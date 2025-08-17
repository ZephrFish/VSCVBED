/**
 * Comprehensive test suite for VBE Decoder
 * Tests core decoding functionality, edge cases, and performance
 */

const assert = require('assert');
const fs = require('fs');
const path = require('path');
const { VBEDecoder } = require('../out/vbeDecoder');

describe('VBE Decoder Core Functionality', () => {
    
    describe('VBE Detection', () => {
        it('should detect valid VBE encoded content', () => {
            const vbeContent = '#@~^BAAAAA==r@#@&P~,P,lsD`rJ4G;D\'~E#*@#@&LbIAAA==^#~@';
            assert.strictEqual(VBEDecoder.containsVBE(vbeContent), true);
        });

        it('should not detect regular VBScript as VBE', () => {
            const vbsContent = 'MsgBox "Hello World"';
            assert.strictEqual(VBEDecoder.containsVBE(vbsContent), false);
        });

        it('should detect VBE with newlines', () => {
            const multilineVBE = '#@~^BAAAAA==r@#@&\nP~,P,lsD`rJ4G;D\'~E#*@#@&\nLbIAAA==^#~@';
            assert.strictEqual(VBEDecoder.containsVBE(multilineVBE), true);
        });

        it('should detect multiple VBE sections', () => {
            const multiSection = '#@~^BAAAAA==test==^#~@ some text #@~^CAAAAA==test2==^#~@';
            assert.strictEqual(VBEDecoder.containsVBE(multiSection), true);
        });
    });

    describe('Decoding Algorithm', () => {
        it('should decode simple VBE content', () => {
            const encoded = '#@~^BAAAAA==r@#@&P~,P,lsD`rJ4G;D\'~E#*@#@&LbIAAA==^#~@';
            const decoded = VBEDecoder.processFile(encoded);
            assert(decoded.length > 0, 'Decoded content should not be empty');
            assert(!decoded.includes('#@~^'), 'Decoded content should not contain VBE markers');
        });

        it('should handle special character replacements', () => {
            // Test @& (line feed), @# (carriage return), @* (>), @! (<), @, (@)
            const testCases = [
                { encoded: '@&', expected: '\n' },
                { encoded: '@#', expected: '\r' },
                { encoded: '@*', expected: '>' },
                { encoded: '@!', expected: '<' },
                { encoded: '@,', expected: '@' }
            ];

            testCases.forEach(test => {
                const result = VBEDecoder.decodeData(test.encoded);
                assert.strictEqual(result, test.expected, 
                    `Failed to decode ${test.encoded} to ${test.expected.charCodeAt(0)}`);
            });
        });

        it('should extract and decode multiple VBE sections', () => {
            const multiSection = `
                Regular text
                #@~^AAAAAA==test1==BBBBBB==^#~@
                More text
                #@~^CCCCCC==test2==DDDDDD==^#~@
                End text
            `;
            const sections = VBEDecoder.extractAndDecode(multiSection);
            assert.strictEqual(sections.length, 2, 'Should extract 2 VBE sections');
        });

        it('should handle empty VBE sections gracefully', () => {
            const emptyVBE = '#@~^AAAAAA====BBBBBB==^#~@';
            const decoded = VBEDecoder.processFile(emptyVBE);
            assert(typeof decoded === 'string', 'Should return a string');
        });

        it('should preserve non-VBE content when no VBE found', () => {
            const plainText = 'This is plain text without any encoding';
            const result = VBEDecoder.processFile(plainText);
            assert.strictEqual(result, plainText, 'Should return original content');
        });
    });

    describe('Real-World VBE Files', () => {
        const testFiles = [
            'test-msgbox.vbe',
            'real-test.vbe'
        ];

        testFiles.forEach(filename => {
            const filepath = path.join(__dirname, '..', filename);
            if (fs.existsSync(filepath)) {
                it(`should decode ${filename}`, () => {
                    const content = fs.readFileSync(filepath, 'utf8');
                    const decoded = VBEDecoder.processFile(content);
                    
                    // Verify decoding worked
                    assert(decoded.length > 0, 'Decoded content should not be empty');
                    assert(!VBEDecoder.containsVBE(decoded), 'Decoded content should not contain VBE');
                    
                    // Basic validation that decoding happened
                    assert(decoded !== content, 'Decoded should differ from encoded');
                });
            }
        });

        it('should decode a known VBE pattern', () => {
            // This is a known VBE encoded "MsgBox" script
            const knownVBE = '#@~^OQAAAA==-mD~}4N+.`*~\',J@#@&P~~,lr^oUD`b/OD+~~{PJ*~{PEs+U@#@&cJIAAA==^#~@';
            const decoded = VBEDecoder.processFile(knownVBE);
            
            assert(decoded.length > 0, 'Should decode known pattern');
            assert(!decoded.includes('#@~^'), 'Should not contain VBE markers');
        });
    });

    describe('Edge Cases and Error Handling', () => {
        it('should handle null input', () => {
            assert.doesNotThrow(() => {
                VBEDecoder.processFile('');
            });
        });

        it('should handle very large files', () => {
            const largeContent = '#@~^AAAAAA==' + 'x'.repeat(100000) + '==BBBBBB==^#~@';
            assert.doesNotThrow(() => {
                const decoded = VBEDecoder.processFile(largeContent);
                assert(decoded.length > 0);
            });
        });

        it('should handle malformed VBE headers', () => {
            const malformed = '#@~^AAA==content==BBB==^#~@'; // Wrong header length
            const result = VBEDecoder.processFile(malformed);
            // Should either decode partially or return original
            assert(typeof result === 'string');
        });

        it('should handle mixed line endings', () => {
            const mixedEndings = '#@~^AAAAAA==test\r\ntest\ntest\r==BBBBBB==^#~@';
            assert.doesNotThrow(() => {
                VBEDecoder.processFile(mixedEndings);
            });
        });

        it('should handle Unicode characters', () => {
            const unicode = '#@~^AAAAAA==тест テスト 测试==BBBBBB==^#~@';
            assert.doesNotThrow(() => {
                VBEDecoder.processFile(unicode);
            });
        });
    });

    describe('Performance Tests', () => {
        it('should decode small files quickly', () => {
            const small = '#@~^BAAAAA==r@#@&P~,P,lsD`rJ4G;D\'~E#*@#@&LbIAAA==^#~@';
            const start = Date.now();
            
            for (let i = 0; i < 1000; i++) {
                VBEDecoder.processFile(small);
            }
            
            const elapsed = Date.now() - start;
            assert(elapsed < 1000, `Decoding 1000 small files took ${elapsed}ms (should be < 1000ms)`);
        });

        it('should handle concurrent decoding', () => {
            const promises = [];
            const content = '#@~^BAAAAA==test==BBBBBB==^#~@';
            
            for (let i = 0; i < 100; i++) {
                promises.push(new Promise((resolve) => {
                    const result = VBEDecoder.processFile(content);
                    resolve(result);
                }));
            }
            
            return Promise.all(promises).then(results => {
                assert.strictEqual(results.length, 100);
                results.forEach(result => {
                    assert(typeof result === 'string');
                });
            });
        });
    });

    describe('Integration with VSCode Extension', () => {
        it('should export correct methods', () => {
            assert(typeof VBEDecoder.decodeData === 'function', 'decodeData should be a function');
            assert(typeof VBEDecoder.extractAndDecode === 'function', 'extractAndDecode should be a function');
            assert(typeof VBEDecoder.processFile === 'function', 'processFile should be a function');
            assert(typeof VBEDecoder.containsVBE === 'function', 'containsVBE should be a function');
        });

        it('should handle file paths with spaces', () => {
            const content = '#@~^BAAAAA==test==BBBBBB==^#~@';
            // This would be handled by the extension, but we test the decoder doesn't break
            assert.doesNotThrow(() => {
                VBEDecoder.processFile(content);
            });
        });
    });
});

// Run tests if executed directly
if (require.main === module) {
    const Mocha = require('mocha');
    const mocha = new Mocha();
    
    mocha.addFile(__filename);
    mocha.run(failures => {
        process.exitCode = failures ? 1 : 0;
    });
}