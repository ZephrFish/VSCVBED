/**
 * Integration tests for VSCode Extension functionality
 * Tests command registration, activation, and file handling
 */

const assert = require('assert');
const path = require('path');

describe('VSCode Extension Integration', () => {
    
    describe('Extension Module', () => {
        let extension;
        
        before(() => {
            extension = require('../out/extension');
        });

        it('should export activate function', () => {
            assert(typeof extension.activate === 'function', 
                'Extension should export activate function');
        });

        it('should export deactivate function', () => {
            assert(typeof extension.deactivate === 'function', 
                'Extension should export deactivate function');
        });

        it('should handle activation without context', () => {
            // Mock minimal context
            const mockContext = {
                subscriptions: [],
                extensionUri: { fsPath: __dirname },
                globalState: {
                    get: () => undefined,
                    update: () => Promise.resolve()
                },
                workspaceState: {
                    get: () => undefined,
                    update: () => Promise.resolve()
                }
            };

            assert.doesNotThrow(() => {
                // Note: This may not fully activate without VSCode APIs
                // but should not throw errors
                extension.activate(mockContext);
            }, 'Activation should not throw errors');
        });

        it('should handle deactivation gracefully', () => {
            assert.doesNotThrow(() => {
                extension.deactivate();
            }, 'Deactivation should not throw errors');
        });
    });

    describe('SMB File System Provider', () => {
        let SMBFileSystemProvider;
        
        before(() => {
            const module = require('../out/smbFileSystemProvider');
            SMBFileSystemProvider = module.SMBFileSystemProvider;
        });

        it('should export SMBFileSystemProvider class', () => {
            assert(typeof SMBFileSystemProvider === 'function', 
                'Should export SMBFileSystemProvider constructor');
        });

        it('should create provider instance', () => {
            const provider = new SMBFileSystemProvider();
            assert(provider !== null, 'Should create provider instance');
            
            // Check required methods exist
            assert(typeof provider.connect === 'function', 'Should have connect method');
            assert(typeof provider.disconnect === 'function', 'Should have disconnect method');
            assert(typeof provider.readFile === 'function', 'Should have readFile method');
            assert(typeof provider.writeFile === 'function', 'Should have writeFile method');
        });

        it('should handle connection options', () => {
            const provider = new SMBFileSystemProvider();
            const options = {
                share: '\\\\localhost\\share',
                username: 'user',
                password: 'pass',
                domain: 'WORKGROUP'
            };

            // Should not throw when connecting (may fail without actual SMB server)
            assert.doesNotThrow(() => {
                provider.connect(options).catch(() => {
                    // Expected to fail without actual SMB server
                });
            });
        });

        it('should handle disconnection', () => {
            const provider = new SMBFileSystemProvider();
            assert.doesNotThrow(() => {
                provider.disconnect();
            }, 'Disconnect should not throw');
        });
    });

    describe('Logger Module', () => {
        let Logger;
        
        before(() => {
            const module = require('../out/logger');
            Logger = module.Logger;
        });

        it('should export Logger class', () => {
            assert(typeof Logger === 'function', 'Should export Logger constructor');
        });

        it('should implement singleton pattern', () => {
            const instance1 = Logger.getInstance();
            const instance2 = Logger.getInstance();
            assert.strictEqual(instance1, instance2, 'Should return same instance');
        });

        it('should have logging methods', () => {
            const logger = Logger.getInstance();
            assert(typeof logger.info === 'function', 'Should have info method');
            assert(typeof logger.error === 'function', 'Should have error method');
            assert(typeof logger.warn === 'function', 'Should have warn method');
            assert(typeof logger.debug === 'function', 'Should have debug method');
        });

        it('should handle log messages without errors', () => {
            const logger = Logger.getInstance();
            assert.doesNotThrow(() => {
                logger.info('Test info message');
                logger.error('Test error message');
                logger.warn('Test warning message');
                logger.debug('Test debug message');
            }, 'Logging should not throw errors');
        });
    });

    describe('Command Registration', () => {
        it('should define expected commands in package.json', () => {
            const packageJson = require('../package.json');
            const commands = packageJson.contributes.commands;
            
            const expectedCommands = [
                'vbeDecoder.openSMB',
                'vbeDecoder.decodeVBE',
                'vbeDecoder.decodeVBEContent',
                'vbeDecoder.openLocalFile',
                'vbeDecoder.toggleSplitPane'
            ];
            
            expectedCommands.forEach(cmdId => {
                const found = commands.some(cmd => cmd.command === cmdId);
                assert(found, `Command ${cmdId} should be registered in package.json`);
            });
        });

        it('should have proper activation events', () => {
            const packageJson = require('../package.json');
            const activationEvents = packageJson.activationEvents;
            
            assert(activationEvents.includes('onFileSystem:smb'), 
                'Should activate on SMB file system');
            assert(activationEvents.includes('onLanguage:vbscript'), 
                'Should activate on VBScript language');
            assert(activationEvents.includes('onCommand:vbeDecoder.decodeVBE'), 
                'Should activate on decode command');
        });

        it('should have context menus configured', () => {
            const packageJson = require('../package.json');
            const menus = packageJson.contributes.menus;
            
            assert(menus['explorer/context'], 'Should have explorer context menu');
            assert(menus['editor/context'], 'Should have editor context menu');
            assert(menus['editor/title'], 'Should have editor title menu');
        });
    });

    describe('Configuration Settings', () => {
        it('should define configuration properties', () => {
            const packageJson = require('../package.json');
            const config = packageJson.contributes.configuration;
            
            assert(config.title === 'VBE Decoder', 'Should have correct title');
            
            const properties = config.properties;
            assert(properties['vbeDecoder.autoDecodeInPane'], 
                'Should have autoDecodeInPane setting');
            assert(properties['vbeDecoder.autoPromptDecode'], 
                'Should have autoPromptDecode setting');
            assert(properties['vbeDecoder.splitPanePosition'], 
                'Should have splitPanePosition setting');
            assert(properties['vbeDecoder.syncScrolling'], 
                'Should have syncScrolling setting');
            assert(properties['vbeDecoder.smbConnections'], 
                'Should have smbConnections setting');
        });

        it('should have valid configuration schemas', () => {
            const packageJson = require('../package.json');
            const properties = packageJson.contributes.configuration.properties;
            
            // Check boolean settings
            assert.strictEqual(properties['vbeDecoder.autoDecodeInPane'].type, 'boolean');
            assert.strictEqual(properties['vbeDecoder.autoPromptDecode'].type, 'boolean');
            assert.strictEqual(properties['vbeDecoder.syncScrolling'].type, 'boolean');
            
            // Check enum setting
            assert.strictEqual(properties['vbeDecoder.splitPanePosition'].type, 'string');
            assert(Array.isArray(properties['vbeDecoder.splitPanePosition'].enum));
            
            // Check array setting
            assert.strictEqual(properties['vbeDecoder.smbConnections'].type, 'array');
        });
    });

    describe('Language Support', () => {
        it('should register VBScript language', () => {
            const packageJson = require('../package.json');
            const languages = packageJson.contributes.languages;
            
            assert(Array.isArray(languages), 'Should have languages array');
            
            const vbscript = languages.find(lang => lang.id === 'vbscript');
            assert(vbscript, 'Should register vbscript language');
            assert(vbscript.extensions.includes('.vbs'), 'Should support .vbs extension');
            assert(vbscript.extensions.includes('.vbe'), 'Should support .vbe extension');
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