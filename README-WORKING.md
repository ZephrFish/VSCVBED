# VBE Decoder Extension - Working Status

## ✅ Fixed Issues
1. **Package Dependencies**: Updated from deprecated `node-smb2` to `@marsaud/smb2`
2. **TypeScript Compilation**: Fixed all type errors and compilation issues
3. **VSCode API**: Fixed deprecated `createFileDecorationProvider` to `registerFileDecorationProvider`

## 🔧 Current Status
The extension now compiles successfully and can be packaged as a VSCode extension. The core functionality includes:

### Features
- **VBE File Decoding**: Decodes Visual Basic Encoded (.vbe) files
- **SMB Share Support**: Connect to and browse SMB/CIFS network shares
- **Local File Support**: Open and decode local VBE files
- **Split Pane View**: View encoded and decoded content side-by-side
- **Auto-Detection**: Automatically detects VBE content in opened files

### Commands Available
- `VBE Decoder: Open SMB Share` - Connect to an SMB share
- `VBE Decoder: Decode VBE File` - Decode a VBE file
- `VBE Decoder: Decode VBE Content` - Decode selected VBE content
- `VBE Decoder: Open Local File` - Open local files for decoding
- `VBE Decoder: Toggle Decoded Split Pane` - Toggle split view

## 📦 Installation Instructions

1. **Build the Extension**:
   ```bash
   npm install
   npm run compile
   ```

2. **Package the Extension** (requires vsce):
   ```bash
   npm install -g vsce
   vsce package
   ```
   This will create a `.vsix` file

3. **Install in VSCode**:
   - Open VSCode
   - Go to Extensions view (Ctrl+Shift+X)
   - Click the "..." menu at the top
   - Select "Install from VSIX..."
   - Choose the generated `.vsix` file

## 🚀 Usage

### For Local VBE Files:
1. Open a `.vbe` file in VSCode
2. The extension will auto-detect and offer to decode
3. Or right-click and select "Decode VBE File"

### For SMB Shares:
1. Run command: `VBE Decoder: Open SMB Share`
2. Enter connection details (host, share, credentials)
3. Browse and open VBE files directly from the share

### For Inline Decoding:
1. Select VBE-encoded text in any file
2. Right-click and select "Decode VBE Content"

## ⚠️ Known Limitations

### VBE Decoding Algorithm
The current VBE decoding implementation may not handle all VBE encoding variations correctly. The algorithm is based on the standard VBE encoding scheme but may need adjustments for:
- Different encoding versions
- Multi-byte character encodings
- Special character handling

### For Security Researchers
This tool is designed for legitimate security analysis and malware research. When analyzing potentially malicious VBE files:
- Always work in an isolated environment
- Use proper sandboxing/virtualization
- Follow your organization's security policies
- Never execute decoded scripts without proper analysis

## 🔒 Security Considerations
- SMB credentials are handled in memory only (not persisted with passwords)
- Decoded content is displayed in VSCode editor (not automatically executed)
- File system access follows VSCode's security model

## 📝 Configuration
The extension can be configured via VSCode settings:
- `vbeDecoder.autoDecodeInPane`: Auto-decode in split pane
- `vbeDecoder.autoPromptDecode`: Show decode prompt
- `vbeDecoder.splitPanePosition`: Split pane position (beside/below)
- `vbeDecoder.syncScrolling`: Synchronize scrolling between panes