# VBE Decoder for VSCode

## TL;DR

A VSCode extension that decodes VBE (Visual Basic Encoded) files and provides seamless SMB share access. Open encoded VBScript files, view the decoded content side-by-side, and work with network shares as if they were local folders.

## Overview

VBE Decoder brings comprehensive Visual Basic Script support to VSCode, handling both encoded and standard VBScript files whilst providing native SMB network share integration. The extension automatically detects and decodes VBE content, making it invaluable for security analysis, legacy code recovery, and systems administration.

## Key Features

### VBE Decoding
- Automatic detection and decoding of VBE encoded content
- Split-pane view showing original and decoded content side-by-side
- Context menu integration for quick decoding
- Selected text decoding from any file type
- Batch processing of multiple files

### SMB Share Integration
- Direct access to SMB/CIFS network shares within VSCode
- Saved connection profiles for quick access
- Full file system provider implementation
- Secure authentication with domain support

### File Management
- Multi-file selection and processing
- Automatic format detection
- Configurable decoding behaviour
- Export decoded content to standard VBS files

## Installation

### Prerequisites
- VSCode version 1.74.0 or higher
- Node.js 16.x or higher (for development)

### Installing the Extension

**From VSIX Package:**
1. Download the latest `.vsix` file from releases
2. Open VSCode and navigate to the Extensions view (Ctrl+Shift+X)
3. Click the menu icon and select "Install from VSIX..."
4. Select the downloaded file

**From Source:**
```bash
git clone https://github.com/ZephrFish/VSC-VBEDecoder
cd VSC-VBEDecoder
npm install
npm run compile
```
Press F5 in VSCode to launch the extension in development mode.

## Usage Guide

### Decoding VBE Files

**Using the Context Menu:**
Right-click any `.vbe` file in the explorer and select "Decode VBE File". The decoded content opens automatically in a new editor tab.

**Using the Command Palette:**
Press Ctrl+Shift+P, type "VBE Decoder: Decode VBE File", and select your target file.

**Decoding Selected Text:**
Highlight VBE encoded text in any editor, right-click, and choose "Decode VBE Content".

### Working with SMB Shares

**Connecting to a Share:**
1. Open the Command Palette (Ctrl+Shift+P)
2. Run "VBE Decoder: Open SMB Share"
3. Enter your connection details:
   - Server hostname or IP address
   - Share name
   - Username and password
   - Domain (if applicable)

The share appears in your workspace and behaves like a local folder.

**Managing Connections:**
Frequently used connections can be saved (without passwords) for quick access. Select from saved connections when opening shares, entering only your password.

## Configuration

Configure the extension through VSCode settings:

```json
{
  "vbeDecoder.autoDecodeInPane": true,
  "vbeDecoder.autoPromptDecode": true,
  "vbeDecoder.splitPanePosition": "beside",
  "vbeDecoder.syncScrolling": false,
  "vbeDecoder.smbConnections": []
}
```

**Settings Explained:**
- `autoDecodeInPane`: Automatically decode VBE files in split view when opened
- `autoPromptDecode`: Show decode prompt for VBE files
- `splitPanePosition`: Position decoded content "beside" or "below" original
- `syncScrolling`: Synchronise scrolling between original and decoded panes
- `smbConnections`: Saved SMB connection profiles (passwords excluded)

## Technical Information

### VBE Decoding Algorithm

The extension implements the complete VBE decoding specification:
- Magic number processing with proper offset handling
- 108-element character transformation table
- 64-element combination switching array
- Special character sequence processing
- Invalid byte filtering

### Security Considerations

- Passwords are never stored in configuration files
- Sensitive data is cleared from memory after use
- All network communication uses SMB2 protocol
- Input validation prevents injection attacks

## Commands Reference

| Command | Description |
|---------|-------------|
| `vbeDecoder.openSMB` | Connect to SMB share |
| `vbeDecoder.decodeVBE` | Decode VBE file |
| `vbeDecoder.decodeVBEContent` | Decode selected text |
| `vbeDecoder.openLocalFile` | Open local file with decode option |
| `vbeDecoder.toggleSplitPane` | Toggle split view for decoded content |

## Troubleshooting

### SMB Connection Issues
Check network connectivity and firewall settings. Ensure SMB2 is enabled on the target server. Verify credentials and domain information are correct.

### Decoding Problems
Confirm the file contains valid VBE encoded content. Check for the characteristic pattern markers. Ensure the file hasn't been corrupted during transfer.

### Extension Not Loading
Verify VSCode version compatibility. Check the Output panel for error messages. Ensure all dependencies are properly installed.

## Development

### Building from Source
```bash
npm install        # Install dependencies
npm run compile    # Compile TypeScript
npm test          # Run tests
npm run watch     # Watch mode for development
```

### Project Structure
- `src/` - TypeScript source files
- `out/` - Compiled JavaScript
- `test/` - Test suite
- `package.json` - Extension manifest

## Contributing

Contributions are welcome. Please fork the repository, create a feature branch, and submit a pull request with your changes. Ensure all tests pass and follow the existing code style.

## Licence

MIT Licence - see the [LICENSE](LICENSE) file for details.

## Acknowledgements

Built on the foundation of security research into VBE encoding formats and the excellent VSCode extension API. Special thanks to the SMB2 protocol implementation contributors.

---

For bug reports and feature requests, please use the [GitHub Issues](https://github.com/ZephrFish/VSC-VBEDecoder/issues) page.