# VBE Decoder - VSCode Extension

A powerful VSCode extension that provides seamless access to SMB shares, local files, and decoding of VBE (Visual Basic Encoded) files directly within the editor.

![VBE Decoder](https://img.shields.io/badge/VSCode-Extension-blue)
![TypeScript](https://img.shields.io/badge/TypeScript-4.9+-blue)
![Node.js](https://img.shields.io/badge/Node.js-16+-green)

## 🚀 Features

### 🔓 VBE Decoding
- **Automatic Split Pane**: Automatically shows decoded content in a split pane when opening VBE files
- **Automatic Detection**: Automatically detects VBE encoded content in files
- **Context Menu Integration**: Right-click on .vbe or .vbs files to decode
- **Selected Text Decoding**: Decode VBE content from selected text in any file
- **Real-time Processing**: Fast decoding using optimized TypeScript algorithms
- **Save Decoded Files**: Option to save decoded content as .vbs files
- **Toggle Split View**: Quickly toggle between split and single pane views
- **Synchronized Scrolling**: Optional synchronized scrolling between original and decoded views

### 🌐 SMB Share Support
- **Direct SMB Access**: Browse and edit files on SMB shares as if they were local
- **Connection Management**: Save and reuse SMB connection configurations
- **Secure Authentication**: Support for domain authentication
- **File System Integration**: Full VSCode file system provider implementation
- **Workspace Integration**: Add SMB shares directly to your workspace

### 📁 Local File Access
- **Multi-file Selection**: Open multiple files simultaneously
- **Format Support**: Support for .vbe, .vbs, .txt, and all file types
- **Auto-decode Prompt**: Automatically offers to decode VBE files when opened

### 🎨 Visual Enhancements
- **File Decorations**: Visual indicators for VBE files in the explorer
- **Syntax Highlighting**: VBScript language support with proper highlighting
- **Command Palette**: Easy access to all features via command palette

## 📦 Installation

### From VSIX Package
1. Download the `.vsix` file from the releases
2. Open VSCode
3. Go to Extensions view (Ctrl+Shift+X)
4. Click the "..." menu and select "Install from VSIX..."
5. Select the downloaded `.vsix` file

### From Source
1. Clone this repository
2. Install dependencies:
   ```bash
   npm install
   ```
3. Compile the extension:
   ```bash
   npm run compile
   ```
4. Press F5 to run the extension in development mode

## 🛠️ Usage

### Decoding VBE Files

#### Method 1: Context Menu
1. Right-click on a `.vbe` or `.vbs` file in the explorer
2. Select "Decode VBE File"
3. The decoded content will open in a new tab

#### Method 2: Command Palette
1. Open Command Palette (Ctrl+Shift+P)
2. Type "VBE Decoder: Decode VBE File"
3. Select the file to decode

#### Method 3: Selected Text
1. Select VBE encoded text in any editor
2. Right-click and select "Decode VBE Content"
3. The decoded content will open in a new tab

### Connecting to SMB Shares

#### Method 1: Command Palette
1. Open Command Palette (Ctrl+Shift+P)
2. Type "VBE Decoder: Open SMB Share"
3. Enter connection details:
   - **Host**: SMB server hostname or IP address
   - **Share**: Share name
   - **Username**: Authentication username
   - **Password**: Authentication password
   - **Domain**: Domain (optional)

#### Method 2: Saved Connections
1. Use a previously saved connection
2. Only password needs to be entered

### Opening Local Files
1. Open Command Palette (Ctrl+Shift+P)
2. Type "VBE Decoder: Open Local File"
3. Select files to open
4. VBE files will automatically prompt for decoding

## ⚙️ Configuration

### Extension Settings

```json
{
  "vbeDecoder.autoDecodeInPane": true,        // Auto-decode VBE files in split pane
  "vbeDecoder.autoPromptDecode": true,        // Show decode prompt when opening VBE files
  "vbeDecoder.splitPanePosition": "beside",   // Position of decoded pane: "beside" or "below"
  "vbeDecoder.syncScrolling": false,          // Synchronize scrolling between panes
  "vbeDecoder.smbConnections": [
    {
      "name": "My Server",
      "host": "192.168.1.100",
      "share": "shared_folder",
      "username": "user",
      "domain": "WORKGROUP"
    }
  ]
}
```

### Workspace Configuration
Add SMB shares to your workspace settings:

```json
{
  "folders": [
    {
      "name": "SMB Share",
      "uri": "smb://192.168.1.100/shared_folder/"
    }
  ]
}
```

## 📋 Commands

| Command | Description | Keyboard Shortcut |
|---------|-------------|-------------------|
| `vbeDecoder.openSMB` | Open SMB Share | - |
| `vbeDecoder.decodeVBE` | Decode VBE File | - |
| `vbeDecoder.decodeVBEContent` | Decode Selected VBE Content | - |
| `vbeDecoder.openLocalFile` | Open Local File | - |
| `vbeDecoder.toggleSplitPane` | Toggle Decoded Split Pane | - |

## 🔧 VBE Decoding Algorithm

This extension implements the complete VBE decoding algorithm, including:

- **Magic Number Processing**: Handles the decoding offset (9)
- **Character Mapping**: Uses the complete 108-element decoding table
- **Combination Switching**: Implements the 64-element combination array
- **String Replacements**: Processes special character sequences (@&, @#, @*, @!, @)
- **Byte Filtering**: Excludes bad bytes (60, 62, 64)
- **Pattern Detection**: Regex pattern matching for VBE content (`#@~^......==...==^#~@`)

### Technical Details

The decoder processes VBE files by:
1. Extracting encoded sections using regex patterns
2. Applying character replacements for known sequences
3. Using lookup tables for character transformation
4. Implementing combination switching for proper decoding
5. Reconstructing the original VBScript code

## 🔒 Security Considerations

- **Password Storage**: SMB passwords are never saved in configuration
- **Memory Management**: Sensitive data is cleared from memory after use
- **Connection Security**: Uses secure SMB2 protocol
- **Input Validation**: All user inputs are validated before processing

## 🤝 Contributing

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🐛 Troubleshooting

### Common Issues

#### SMB Connection Fails
- Verify network connectivity to the SMB server
- Check firewall settings on both client and server
- Ensure SMB2 protocol is enabled on the server
- Verify username, password, and domain are correct

#### VBE Decoding Fails
- Ensure the file contains valid VBE encoded content
- Check the VBE pattern format: `#@~^......==...==^#~@`
- Verify the file is not corrupted

#### Extension Not Loading
- Check VSCode version compatibility (1.74.0+)
- Verify all dependencies are installed
- Check the Output panel for error messages

### Debug Mode
Enable debug logging in the Output panel:
1. Open Output panel (View → Output)
2. Select "VBE Decoder" from the dropdown
3. Monitor log messages for troubleshooting

## 📚 References

- [VSCode Extension API](https://code.visualstudio.com/api)
- [SMB2 Protocol](https://docs.microsoft.com/en-us/openspecs/windows_protocols/ms-smb2/)
- [VBScript Encoding Format](https://github.com/unixfreaxjp/vbe-decoder)

## 🙏 Acknowledgments

- Original VBE decoding algorithm inspiration from various security research
- VSCode extension development community
- SMB2 protocol implementation contributors

---

**Made with ❤️ for the VSCode community**