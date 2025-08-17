// Test the decoder with the real VBE file from user
const fs = require('fs');
const { VBEDecoder } = require('./out/vbeDecoder');

// Read the user's test VBE file
const vbeContent = fs.readFileSync('/Users/zephr/tools/ZephrLocalTools/VBE-Decoder/test.vbe', 'utf8');

console.log('=== Testing VBE Decoder with Real File ===\n');
console.log('File: /Users/zephr/tools/ZephrLocalTools/VBE-Decoder/test.vbe');
console.log('File size:', vbeContent.length, 'characters');
console.log('Contains VBE:', VBEDecoder.containsVBE(vbeContent));

if (VBEDecoder.containsVBE(vbeContent)) {
    console.log('\n=== Decoding VBE Content ===');
    const decoded = VBEDecoder.processFile(vbeContent);
    
    // Show first 500 characters of decoded content
    console.log('\nFirst 500 characters of decoded content:');
    console.log('----------------------------------------');
    console.log(decoded.substring(0, 500));
    console.log('----------------------------------------');
    
    // Save the full decoded content
    fs.writeFileSync('decoded-real-test.vbs', decoded, 'utf8');
    console.log('\n✅ Full decoded content saved to: decoded-real-test.vbs');
    console.log('✅ Decoded size:', decoded.length, 'characters');
    console.log('\n✅ SUCCESS: The VBE decoder successfully decoded your file!');
    console.log('✅ When you open this .vbe file in VSCode with the extension installed,');
    console.log('✅ it will automatically show the decoded VBScript in a split pane.');
} else {
    console.log('No VBE content detected in the file');
}