import * as vscode from 'vscode';
import * as path from 'path';
const SMB2 = require('@marsaud/smb2');

export interface SMBConnection {
    name: string;
    host: string;
    share: string;
    username: string;
    password?: string;
    domain?: string;
}

export class SMBFileSystemProvider implements vscode.FileSystemProvider {
    private connections = new Map<string, any>();
    private _emitter = new vscode.EventEmitter<vscode.FileChangeEvent[]>();
    
    readonly onDidChangeFile: vscode.Event<vscode.FileChangeEvent[]> = this._emitter.event;

    constructor() {}

    /**
     * Create a new SMB connection
     * @param connectionInfo SMB connection details
     * @returns Connection identifier
     */
    public async createConnection(connectionInfo: SMBConnection): Promise<string> {
        const connectionId = `${connectionInfo.host}/${connectionInfo.share}`;
        
        const smbOptions = {
            share: `\\\\${connectionInfo.host}\\${connectionInfo.share}`,
            domain: connectionInfo.domain || '',
            username: connectionInfo.username,
            password: connectionInfo.password || '',
            autoCloseTimeout: 0
        };

        try {
            const smbClient = new SMB2(smbOptions);
            this.connections.set(connectionId, smbClient);
            return connectionId;
        } catch (error) {
            throw new Error(`Failed to create SMB connection: ${error}`);
        }
    }

    /**
     * Get SMB client for a given URI
     * @param uri File URI
     * @returns SMB client instance
     */
    private getSMBClient(uri: vscode.Uri): any {
        const connectionId = `${uri.authority}`;
        const client = this.connections.get(connectionId);
        
        if (!client) {
            throw new Error(`No SMB connection found for ${connectionId}`);
        }
        
        return client;
    }

    /**
     * Convert VSCode URI to SMB path
     * @param uri VSCode URI
     * @returns SMB-compatible path
     */
    private uriToSMBPath(uri: vscode.Uri): string {
        return uri.path.replace(/\//g, '\\');
    }

    watch(uri: vscode.Uri, options: { recursive: boolean; excludes: string[]; }): vscode.Disposable {
        // SMB file watching is not implemented in this basic version
        return new vscode.Disposable(() => {});
    }

    async stat(uri: vscode.Uri): Promise<vscode.FileStat> {
        const client = this.getSMBClient(uri);
        const smbPath = this.uriToSMBPath(uri);

        return new Promise((resolve, reject) => {
            client.exists(smbPath, (exists: boolean) => {
                if (!exists) {
                    reject(new vscode.FileSystemError('File not found'));
                    return;
                }

                // Get file stats
                client.readdir(path.dirname(smbPath), (err: any, files: any[]) => {
                    if (err) {
                        // Try to read as file
                        client.readFile(smbPath, (readErr: any, data: any) => {
                            if (readErr) {
                                reject(new vscode.FileSystemError('Cannot access file'));
                                return;
                            }
                            
                            resolve({
                                type: vscode.FileType.File,
                                ctime: 0,
                                mtime: 0,
                                size: data ? data.length : 0
                            });
                        });
                        return;
                    }

                    const fileName = path.basename(smbPath);
                    const fileInfo = files.find((file: any) => file.Filename === fileName);
                    
                    if (fileInfo) {
                        const isDirectory = fileInfo.FileAttributes & 0x10; // FILE_ATTRIBUTE_DIRECTORY
                        
                        resolve({
                            type: isDirectory ? vscode.FileType.Directory : vscode.FileType.File,
                            ctime: new Date(fileInfo.CreationTime).getTime(),
                            mtime: new Date(fileInfo.LastWriteTime).getTime(),
                            size: fileInfo.EndOfFile || 0
                        });
                    } else {
                        reject(new vscode.FileSystemError('File not found'));
                    }
                });
            });
        });
    }

    async readDirectory(uri: vscode.Uri): Promise<[string, vscode.FileType][]> {
        const client = this.getSMBClient(uri);
        const smbPath = this.uriToSMBPath(uri);

        return new Promise((resolve, reject) => {
            client.readdir(smbPath, (err: any, files: any[]) => {
                if (err) {
                    reject(new vscode.FileSystemError('Cannot read directory'));
                    return;
                }

                const entries: [string, vscode.FileType][] = files
                    .filter((file: any) => file.Filename !== '.' && file.Filename !== '..')
                    .map((file: any) => {
                        const isDirectory = file.FileAttributes & 0x10; // FILE_ATTRIBUTE_DIRECTORY
                        return [
                            file.Filename,
                            isDirectory ? vscode.FileType.Directory : vscode.FileType.File
                        ];
                    });

                resolve(entries);
            });
        });
    }

    async createDirectory(uri: vscode.Uri): Promise<void> {
        const client = this.getSMBClient(uri);
        const smbPath = this.uriToSMBPath(uri);

        return new Promise((resolve, reject) => {
            client.mkdir(smbPath, (err: any) => {
                if (err) {
                    reject(new vscode.FileSystemError('Cannot create directory'));
                    return;
                }
                resolve();
            });
        });
    }

    async readFile(uri: vscode.Uri): Promise<Uint8Array> {
        const client = this.getSMBClient(uri);
        const smbPath = this.uriToSMBPath(uri);

        return new Promise((resolve, reject) => {
            client.readFile(smbPath, (err: any, data: any) => {
                if (err) {
                    reject(new vscode.FileSystemError('Cannot read file'));
                    return;
                }
                
                resolve(new Uint8Array(data));
            });
        });
    }

    async writeFile(uri: vscode.Uri, content: Uint8Array, options: { create: boolean; overwrite: boolean; }): Promise<void> {
        const client = this.getSMBClient(uri);
        const smbPath = this.uriToSMBPath(uri);

        return new Promise((resolve, reject) => {
            const buffer = Buffer.from(content);
            
            client.writeFile(smbPath, buffer, (err: any) => {
                if (err) {
                    reject(new vscode.FileSystemError('Cannot write file'));
                    return;
                }
                resolve();
            });
        });
    }

    async delete(uri: vscode.Uri, options: { recursive: boolean; }): Promise<void> {
        const client = this.getSMBClient(uri);
        const smbPath = this.uriToSMBPath(uri);

        return new Promise((resolve, reject) => {
            // Check if it's a directory or file first
            this.stat(uri).then(stat => {
                if (stat.type === vscode.FileType.Directory) {
                    client.rmdir(smbPath, (err: any) => {
                        if (err) {
                            reject(new vscode.FileSystemError('Cannot delete directory'));
                            return;
                        }
                        resolve();
                    });
                } else {
                    client.unlink(smbPath, (err: any) => {
                        if (err) {
                            reject(new vscode.FileSystemError('Cannot delete file'));
                            return;
                        }
                        resolve();
                    });
                }
            }).catch(reject);
        });
    }

    async rename(oldUri: vscode.Uri, newUri: vscode.Uri, options: { overwrite: boolean; }): Promise<void> {
        const client = this.getSMBClient(oldUri);
        const oldPath = this.uriToSMBPath(oldUri);
        const newPath = this.uriToSMBPath(newUri);

        return new Promise((resolve, reject) => {
            client.rename(oldPath, newPath, (err: any) => {
                if (err) {
                    reject(new vscode.FileSystemError('Cannot rename file'));
                    return;
                }
                resolve();
            });
        });
    }

    /**
     * Close a specific SMB connection
     * @param connectionId Connection identifier
     */
    public closeConnection(connectionId: string): void {
        const client = this.connections.get(connectionId);
        if (client) {
            client.disconnect();
            this.connections.delete(connectionId);
        }
    }

    /**
     * Close all SMB connections
     */
    public closeAllConnections(): void {
        for (const [id, client] of this.connections) {
            client.disconnect();
        }
        this.connections.clear();
    }
}