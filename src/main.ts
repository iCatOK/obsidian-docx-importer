import {
  Plugin,
  Notice,
  TFile,
  Menu,
  TFolder,
  WorkspaceLeaf,
} from 'obsidian';
import {
  DocxImporterSettings,
  DEFAULT_SETTINGS,
  ConversionResult,
} from './types';
import { DocxParser } from './parsers/docxParser';
import { MainConverter } from './converters/mainConverter';
import { DocxImporterSettingsTab } from './ui/settingsTab';
import { ImportModal } from './ui/importModal';
import { sanitizeFileName, normalizeWhitespace, trimLines } from './utils/helpers';

export default class DocxImporterPlugin extends Plugin {
  settings: DocxImporterSettings = DEFAULT_SETTINGS;
  private parser: DocxParser;
  private converter: MainConverter | null = null;

  async onload() {
    console.log('Loading DOCX Importer plugin');

    await this.loadSettings();

    this.parser = new DocxParser();
    this.converter = new MainConverter(this.settings, this.app.vault);

    // Add ribbon icon
    this.addRibbonIcon('file-input', 'Import DOCX', () => {
      new ImportModal(this.app, this).open();
    });

    // Add command
    this.addCommand({
      id: 'import-docx',
      name: 'Import DOCX file',
      callback: () => {
        new ImportModal(this.app, this).open();
      },
    });

    // Add command for drag and drop import
    this.addCommand({
      id: 'import-docx-from-clipboard',
      name: 'Import DOCX from file path in clipboard',
      callback: async () => {
        try {
          const clipboardText = await navigator.clipboard.readText();
          if (clipboardText.endsWith('.docx')) {
            new Notice('Please use the Import DOCX command and select a file');
          }
        } catch (error) {
          new Notice('Could not read clipboard');
        }
      },
    });

    // Register file menu event for .docx files dropped in vault
    this.registerEvent(
      this.app.workspace.on('file-menu', (menu: Menu, file: TFile | TFolder) => {
        if (file instanceof TFile && file.extension === 'docx') {
          menu.addItem((item) => {
            item
              .setTitle('Convert to Markdown')
              .setIcon('file-text')
              .onClick(async () => {
                await this.convertDocxInVault(file);
              });
          });
        }
      })
    );

    // Register drop handler for external files
    this.registerDomEvent(document, 'drop', async (evt: DragEvent) => {
      const files = evt.dataTransfer?.files;
      if (files && files.length > 0) {
        for (let i = 0; i < files.length; i++) {
          const file = files[i];
          if (file.name.endsWith('.docx')) {
            evt.preventDefault();
            const arrayBuffer = await file.arrayBuffer();
            const outputName = file.name.replace('.docx', '');
            await this.importDocx(arrayBuffer, outputName);
          }
        }
      }
    });

    // Add settings tab
    this.addSettingTab(new DocxImporterSettingsTab(this.app, this));

    console.log('DOCX Importer plugin loaded');
  }

  onunload() {
    console.log('Unloading DOCX Importer plugin');
  }

  async loadSettings() {
    this.settings = Object.assign({}, DEFAULT_SETTINGS, await this.loadData());
  }

  async saveSettings() {
    await this.saveData(this.settings);
    if (this.converter) {
      this.converter.updateSettings(this.settings);
    }
  }

  async importDocx(
    arrayBuffer: ArrayBuffer,
    outputFileName: string
  ): Promise<ConversionResult> {
    if (!this.converter) {
      this.converter = new MainConverter(this.settings, this.app.vault);
    }

    // Parse the DOCX file
    const parsed = await this.parser.parse(arrayBuffer);

    // Get current folder path
    const activeFile = this.app.workspace.getActiveFile();
    const currentPath = activeFile?.parent?.path || '';

    // Convert to Markdown
    const result = await this.converter.convert(parsed, currentPath);

    // Clean up the markdown
    let markdown = result.markdown;
    markdown = normalizeWhitespace(markdown);
    markdown = trimLines(markdown);

    // Create the output file
    const sanitizedName = sanitizeFileName(outputFileName);
    let filePath = `${sanitizedName}.md`;

    // If we're in a folder, create the file there
    if (currentPath) {
      filePath = `${currentPath}/${filePath}`;
    }

    // Check if file already exists
    let finalPath = filePath;
    let counter = 1;
    while (this.app.vault.getAbstractFileByPath(finalPath)) {
      finalPath = currentPath
        ? `${currentPath}/${sanitizedName}_${counter}.md`
        : `${sanitizedName}_${counter}.md`;
      counter++;
    }

    // Create the file
    const newFile = await this.app.vault.create(finalPath, markdown);

    // Open the new file
    const leaf = this.app.workspace.getLeaf(false);
    await leaf.openFile(newFile);

    return result;
  }

  async convertDocxInVault(file: TFile): Promise<void> {
    try {
      new Notice('Converting DOCX file...');

      const arrayBuffer = await this.app.vault.readBinary(file);
      const outputName = file.basename;

      const result = await this.importDocx(arrayBuffer, outputName);

      // Optionally delete the original DOCX file
      // await this.app.vault.delete(file);

      if (result.warnings.length > 0) {
        new Notice(
          `Conversion completed with ${result.warnings.length} warning(s)`
        );
      } else {
        new Notice('DOCX file converted successfully!');
      }
    } catch (error) {
      console.error('Conversion error:', error);
      new Notice(`Conversion failed: ${error instanceof Error ? error.message : 'Unknown error'}`);
    }
  }
}