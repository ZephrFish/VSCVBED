// Test the VBE decoder directly
const fs = require('fs');
const path = require('path');

// Import the compiled decoder
const { VBEDecoder } = require('./out/vbeDecoder');

// Read the test VBE file
const vbeContent = fs.readFileSync('msgbox-test.vbe', 'utf8');

console.log('=== Testing VBE Decoder ===\n');
console.log('Input file: test-vbe.vbe');
console.log('Contains VBE:', VBEDecoder.containsVBE(vbeContent));

if (VBEDecoder.containsVBE(vbeContent)) {
    console.log('\n=== Decoded Content ===\n');
    const decoded = VBEDecoder.processFile(vbeContent);
    console.log(decoded);
    
    // Save decoded output
    fs.writeFileSync('test-decoded.vbs', decoded, 'utf8');
    console.log('\n=== Decoded content saved to test-decoded.vbs ===');
} else {
    console.log('No VBE content found in file');
}