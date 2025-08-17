import * as vscode from 'vscode';
import * as path from 'path';
import { VBEDecoder } from './vbeDecoder';
import { SMBFileSystemProvider, SMBConnection } from './smbFileSystemProvider';
import { Logger } from './logger';

let smbProvider: SMBFileSystemProvider;

export function activate(context: vscode.ExtensionContext) {
    // Initialise SMB file system provider
    smbProvider = new SMBFileSystemProvider();
    context.subscriptions.push(
        vscode.workspace.registerFileSystemProvider('smb', smbProvider, { 
            isCaseSensitive: false,
            isReadonly: false
        })
    );

    // Register commands
    const openSMBCommand = vscode.commands.registerCommand('vbeDecoder.openSMB', openSMBShare);
    const decodeVBECommand = vscode.commands.registerCommand('vbeDecoder.decodeVBE', decodeVBEFile);
    const decodeVBEContentCommand = vscode.commands.registerCommand('vbeDecoder.decodeVBEContent', decodeVBEContent);
    const openLocalFileCommand = vscode.commands.registerCommand('vbeDecoder.openLocalFile', openLocalFile);
    const toggleSplitPaneCommand = vscode.commands.registerCommand('vbeDecoder.toggleSplitPane', toggleSplitPane);

    context.subscriptions.push(openSMBCommand, decodeVBECommand, decodeVBEContentCommand, openLocalFileCommand, toggleSplitPaneCommand);

    // Register file decorations for .vbe files
    const vbeFileDecoration = vscode.window.registerFileDecorationProvider({
        provideFileDecoration(uri: vscode.Uri): vscode.FileDecoration | undefined {
            if (uri.path.endsWith('.vbe')) {
                return {
                    badge: 'VBE',
                    tooltip: 'Visual Basic Encoded File',
                    color: new vscode.ThemeColor('charts.orange')
                };
            }
            return undefined;
        }
    });

    context.subscriptions.push(vbeFileDecoration);

    // Auto-decode VBE files when opened
    const onDidOpenTextDocument = vscode.workspace.onDidOpenTextDocument(async (document) => {
        if (document.fileName.endsWith('.vbe') || VBEDecoder.containsVBE(document.getText())) {
            await handleVBEFileOpened(document);
        }
    });

    context.subscriptions.push(onDidOpenTextDocument);
}

export function deactivate() {
    if (smbProvider) {
        smbProvider.closeAllConnections();
    }
}

/**
 * Open SMB share connection
 */
async function openSMBShare(): Promise<void> {
    try {
        const connectionInfo = await promptForSMBConnection();
        if (!connectionInfo) {
            return;
        }

        // Create the connection
        const connectionId = await smbProvider.createConnection(connectionInfo);
        
        // Open the SMB share in workspace
        const uri = vscode.Uri.parse(`smb://${connectionId}/`);
        
        // Add to workspace folders
        const workspaceEdit = new vscode.WorkspaceEdit();
        const success = vscode.workspace.updateWorkspaceFolders(
            vscode.workspace.workspaceFolders ? vscode.workspace.workspaceFolders.length : 0,
            0,
            { uri, name: connectionInfo.name || `SMB: ${connectionInfo.host}/${connectionInfo.share}` }
        );

        if (success) {
            vscode.window.showInformationMessage(`Successfully connected to SMB share: ${connectionInfo.host}/${connectionInfo.share}`);
            
            // Save connection for future use
            await saveConnectionConfiguration(connectionInfo);
        } else {
            vscode.window.showErrorMessage('Failed to add SMB share to workspace');
        }

    } catch (error) {
        vscode.window.showErrorMessage(`Failed to connect to SMB share: ${error}`);
    }
}

/**
 * Decode VBE file from explorer context menu
 */
async function decodeVBEFile(uri?: vscode.Uri): Promise<void> {
    try {
        let targetUri = uri;
        
        if (!targetUri) {
            // No URI provided, prompt user to select file
            const fileUris = await vscode.window.showOpenDialog({
                canSelectFiles: true,
                canSelectFolders: false,
                canSelectMany: false,
                filters: {
                    'VBE Files': ['vbe'],
                    'VBScript Files': ['vbs'],
                    'All Files': ['*']
                },
                title: 'Select VBE file to decode'
            });

            if (!fileUris || fileUris.length === 0) {
                return;
            }

            targetUri = fileUris[0];
        }

        await processVBEFile(targetUri);

    } catch (error) {
        vscode.window.showErrorMessage(`Failed to decode VBE file: ${error}`);
    }
}

/**
 * Decode VBE content from selected text
 */
async function decodeVBEContent(): Promise<void> {
    const editor = vscode.window.activeTextEditor;
    if (!editor) {
        vscode.window.showWarningMessage('No active editor found');
        return;
    }

    const selection = editor.selection;
    const selectedText = editor.document.getText(selection);

    if (!selectedText) {
        vscode.window.showWarningMessage('No text selected');
        return;
    }

    try {
        if (VBEDecoder.containsVBE(selectedText)) {
            const decoded = VBEDecoder.processFile(selectedText);
            
            // Create new document with decoded content
            const newDocument = await vscode.workspace.openTextDocument({
                content: decoded,
                language: 'vbscript'
            });

            await vscode.window.showTextDocument(newDocument);
            vscode.window.showInformationMessage('VBE content decoded successfully');
        } else {
            vscode.window.showWarningMessage('Selected text does not contain VBE encoded content');
        }

    } catch (error) {
        vscode.window.showErrorMessage(`Failed to decode VBE content: ${error}`);
    }
}

/**
 * Process VBE file and show decoded content
 */
async function processVBEFile(uri: vscode.Uri): Promise<void> {
    try {
        const fileContent = await vscode.workspace.fs.readFile(uri);
        const textContent = Buffer.from(fileContent).toString('utf8');

        if (!VBEDecoder.containsVBE(textContent)) {
            vscode.window.showWarningMessage('File does not contain VBE encoded content');
            return;
        }

        const decoded = VBEDecoder.processFile(textContent);
        
        // Create new document with decoded content
        const fileName = path.basename(uri.path, '.vbe') + '_decoded.vbs';
        const newDocument = await vscode.workspace.openTextDocument({
            content: decoded,
            language: 'vbscript'
        });

        await vscode.window.showTextDocument(newDocument);
        vscode.window.showInformationMessage(`VBE file decoded successfully: ${fileName}`);

        // Offer to save the decoded file
        const saveDecoded = await vscode.window.showInformationMessage(
            'Would you like to save the decoded content?',
            'Save',
            'Don\'t Save'
        );

        if (saveDecoded === 'Save') {
            const saveUri = await vscode.window.showSaveDialog({
                defaultUri: vscode.Uri.file(fileName),
                filters: {
                    'VBScript Files': ['vbs'],
                    'Text Files': ['txt'],
                    'All Files': ['*']
                }
            });

            if (saveUri) {
                await vscode.workspace.fs.writeFile(saveUri, Buffer.from(decoded, 'utf8'));
                vscode.window.showInformationMessage(`Decoded file saved: ${saveUri.fsPath}`);
            }
        }

    } catch (error) {
        vscode.window.showErrorMessage(`Failed to process VBE file: ${error}`);
    }
}

/**
 * Handle when a VBE file is opened
 */
async function handleVBEFileOpened(document: vscode.TextDocument): Promise<void> {
    // Check configuration for auto-decode setting
    const config = vscode.workspace.getConfiguration('vbeDecoder');
    const autoDecodeInPane = config.get<boolean>('autoDecodeInPane', true);
    const autoPromptDecode = config.get<boolean>('autoPromptDecode', true);

    if (!autoPromptDecode && !autoDecodeInPane) {
        return;
    }

    // If auto-decode in pane is enabled, decode immediately in split view
    if (autoDecodeInPane) {
        await showDecodedInSplitPane(document);
        return;
    }

    // Otherwise show prompt
    const option = await vscode.window.showInformationMessage(
        'This file contains VBE encoded content. Would you like to decode it?',
        'Decode in Split',
        'Decode',
        'View Raw',
        'Don\'t Ask Again'
    );

    if (option === 'Decode in Split') {
        await showDecodedInSplitPane(document);
    } else if (option === 'Decode') {
        const decoded = VBEDecoder.processFile(document.getText());
        
        const newDocument = await vscode.workspace.openTextDocument({
            content: decoded,
            language: 'vbscript'
        });

        await vscode.window.showTextDocument(newDocument);
    } else if (option === 'Don\'t Ask Again') {
        // Save preference
        await config.update('autoPromptDecode', false, vscode.ConfigurationTarget.Global);
    }
}

/**
 * Show decoded content in a split pane
 */
async function showDecodedInSplitPane(document: vscode.TextDocument): Promise<void> {
    try {
        const decoded = VBEDecoder.processFile(document.getText());
        
        // Get configuration
        const config = vscode.workspace.getConfiguration('vbeDecoder');
        const splitPosition = config.get<string>('splitPanePosition', 'beside');
        const syncScrolling = config.get<boolean>('syncScrolling', false);
        
        // Create a virtual document with decoded content
        const fileName = path.basename(document.fileName, '.vbe') + '_decoded.vbs';
        const decodedDocument = await vscode.workspace.openTextDocument({
            content: decoded,
            language: 'vbscript'
        });

        // Get the current active editor
        const currentEditor = vscode.window.activeTextEditor;
        
        // Determine view column based on configuration
        const viewColumn = splitPosition === 'below' 
            ? vscode.ViewColumn.Active 
            : vscode.ViewColumn.Beside;
        
        // Show decoded content in split pane
        await vscode.window.showTextDocument(
            decodedDocument,
            {
                viewColumn: viewColumn,
                preserveFocus: false,
                preview: false
            }
        );

        // Add a status bar message
        vscode.window.setStatusBarMessage('✓ VBE file decoded in split pane', 3000);

        // If we have the original editor and sync is enabled, setup sync
        if (currentEditor && syncScrolling) {
            setupScrollSync(currentEditor, vscode.window.activeTextEditor!);
        }

    } catch (error) {
        vscode.window.showErrorMessage(`Failed to decode VBE file: ${error}`);
    }
}

/**
 * Setup synchronized scrolling between original and decoded views
 */
function setupScrollSync(originalEditor: vscode.TextEditor, decodedEditor: vscode.TextEditor): void {
    // Optional: Implement scroll synchronization
    const disposable = vscode.window.onDidChangeTextEditorVisibleRanges((event) => {
        if (event.textEditor === originalEditor) {
            // Sync scroll position to decoded editor
            const range = originalEditor.visibleRanges[0];
            if (range) {
                decodedEditor.revealRange(
                    new vscode.Range(range.start.line, 0, range.start.line, 0),
                    vscode.TextEditorRevealType.AtTop
                );
            }
        }
    });

    // Store disposable to clean up later if needed
    // This is simplified - in production you'd want to manage these disposables properly
}

/**
 * Toggle split pane view for current VBE file
 */
async function toggleSplitPane(): Promise<void> {
    const editor = vscode.window.activeTextEditor;
    if (!editor) {
        vscode.window.showWarningMessage('No active editor found');
        return;
    }

    const document = editor.document;
    
    // Check if current file contains VBE content
    if (!document.fileName.endsWith('.vbe') && !VBEDecoder.containsVBE(document.getText())) {
        vscode.window.showWarningMessage('Current file does not contain VBE encoded content');
        return;
    }

    // Check if we already have a split pane with decoded content
    const allEditors = vscode.window.visibleTextEditors;
    const decodedEditor = allEditors.find(e => 
        e.document.languageId === 'vbscript' && 
        e.document.getText().includes('') && // Check for decoded content
        e !== editor
    );

    if (decodedEditor) {
        // Close the decoded pane
        await vscode.commands.executeCommand('workbench.action.closeActiveEditor');
        vscode.window.setStatusBarMessage('Decoded pane closed', 3000);
    } else {
        // Open decoded content in split pane
        await showDecodedInSplitPane(document);
    }
}

/**
 * Prompt user for SMB connection details
 */
async function promptForSMBConnection(): Promise<SMBConnection | undefined> {
    // Check for saved connections first
    const savedConnections = await getSavedConnections();
    
    if (savedConnections.length > 0) {
        const options = [
            ...savedConnections.map(conn => ({ label: conn.name, connection: conn })),
            { label: '$(plus) Add New Connection', connection: null }
        ];

        const selected = await vscode.window.showQuickPick(options, {
            placeHolder: 'Select SMB connection or create new'
        });

        if (!selected) {
            return undefined;
        }

        if (selected.connection) {
            // Get password for saved connection
            const password = await vscode.window.showInputBox({
                prompt: `Enter password for ${selected.connection.username}@${selected.connection.host}`,
                password: true
            });

            if (password === undefined) {
                return undefined;
            }

            return { ...selected.connection, password };
        }
    }

    // Create new connection
    const name = await vscode.window.showInputBox({
        prompt: 'Enter connection name (optional)',
        placeHolder: 'My SMB Share'
    });

    const host = await vscode.window.showInputBox({
        prompt: 'Enter SMB server hostname or IP address',
        placeHolder: '192.168.1.100 or server.domain.com',
        validateInput: (value) => {
            if (!value || value.trim() === '') {
                return 'Host is required';
            }
            return null;
        }
    });

    if (!host) {
        return undefined;
    }

    const share = await vscode.window.showInputBox({
        prompt: 'Enter share name',
        placeHolder: 'shared_folder',
        validateInput: (value) => {
            if (!value || value.trim() === '') {
                return 'Share name is required';
            }
            return null;
        }
    });

    if (!share) {
        return undefined;
    }

    const username = await vscode.window.showInputBox({
        prompt: 'Enter username',
        placeHolder: 'domain\\username or username'
    });

    if (!username) {
        return undefined;
    }

    const password = await vscode.window.showInputBox({
        prompt: 'Enter password',
        password: true
    });

    if (password === undefined) {
        return undefined;
    }

    const domain = await vscode.window.showInputBox({
        prompt: 'Enter domain (optional)',
        placeHolder: 'WORKGROUP'
    });

    return {
        name: name || `${host}/${share}`,
        host,
        share,
        username,
        password,
        domain
    };
}

/**
 * Get saved SMB connections from configuration
 */
async function getSavedConnections(): Promise<SMBConnection[]> {
    const config = vscode.workspace.getConfiguration('vbeDecoder');
    return config.get<SMBConnection[]>('smbConnections', []);
}

/**
 * Save SMB connection configuration (without password)
 */
async function saveConnectionConfiguration(connection: SMBConnection): Promise<void> {
    const shouldSave = await vscode.window.showInformationMessage(
        'Would you like to save this connection for future use? (Password will not be saved)',
        'Save',
        'Don\'t Save'
    );

    if (shouldSave === 'Save') {
        const config = vscode.workspace.getConfiguration('vbeDecoder');
        const savedConnections = config.get<SMBConnection[]>('smbConnections', []);
        
        // Remove password before saving
        const connectionToSave = { ...connection };
        delete connectionToSave.password;
        
        // Check if connection already exists
        const existingIndex = savedConnections.findIndex(
            conn => conn.host === connection.host && conn.share === connection.share && conn.username === connection.username
        );

        if (existingIndex >= 0) {
            savedConnections[existingIndex] = connectionToSave;
        } else {
            savedConnections.push(connectionToSave);
        }

        await config.update('smbConnections', savedConnections, vscode.ConfigurationTarget.Global);
        vscode.window.showInformationMessage('Connection saved successfully');
    }
}

/**
 * Open local file for VBE decoding
 */
async function openLocalFile(): Promise<void> {
    try {
        const fileUris = await vscode.window.showOpenDialog({
            canSelectFiles: true,
            canSelectFolders: false,
            canSelectMany: true,
            filters: {
                'VBE Files': ['vbe'],
                'VBScript Files': ['vbs'],
                'Text Files': ['txt'],
                'All Files': ['*']
            },
            title: 'Select files to open'
        });

        if (!fileUris || fileUris.length === 0) {
            return;
        }

        // Open each selected file
        for (const uri of fileUris) {
            try {
                const document = await vscode.workspace.openTextDocument(uri);
                await vscode.window.showTextDocument(document);

                // Check if it's a VBE file and offer to decode
                if (uri.path.endsWith('.vbe') || VBEDecoder.containsVBE(document.getText())) {
                    const shouldDecode = await vscode.window.showInformationMessage(
                        `File "${path.basename(uri.path)}" contains VBE encoded content. Decode it?`,
                        'Decode',
                        'Skip'
                    );

                    if (shouldDecode === 'Decode') {
                        await processVBEFile(uri);
                    }
                }
            } catch (error) {
                vscode.window.showErrorMessage(`Failed to open file ${uri.path}: ${error}`);
            }
        }

    } catch (error) {
        vscode.window.showErrorMessage(`Failed to open files: ${error}`);
    }
}