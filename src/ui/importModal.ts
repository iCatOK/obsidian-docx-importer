import { App, Modal, Notice, Setting, TFile } from 'obsidian';
import DocxImporterPlugin from '../main';

export class ImportModal extends Modal {
  plugin: DocxImporterPlugin;
  file: File | null = null;
  outputFileName: string = '';

  constructor(app: App, plugin: DocxImporterPlugin) {
    super(app);
    this.plugin = plugin;
  }

  onOpen() {
    const { contentEl } = this;
    contentEl.empty();

    contentEl.createEl('h2', { text: 'Import DOCX File' });

    // File input
    new Setting(contentEl)
      .setName('Select DOCX file')
      .setDesc('Choose a .docx file to import')
      .addButton((button) =>
        button.setButtonText('Choose File').onClick(() => {
          const input = document.createElement('input');
          input.type = 'file';
          input.accept = '.docx';
          input.onchange = (e) => {
            const target = e.target as HTMLInputElement;
            if (target.files && target.files.length > 0) {
              this.file = target.files[0];
              this.outputFileName = this.file.name.replace('.docx', '');
              fileNameDisplay.setText(`Selected: ${this.file.name}`);
              outputNameInput.setValue(this.outputFileName);
            }
          };
          input.click();
        })
      );

    const fileNameDisplay = contentEl.createEl('p', {
      text: 'No file selected',
      cls: 'docx-importer-file-display',
    });

    // Output file name
    let outputNameInput: any;
    new Setting(contentEl)
      .setName('Output file name')
      .setDesc('Name for the created markdown file (without extension)')
      .addText((text) => {
        outputNameInput = text;
        text
          .setPlaceholder('output')
          .setValue(this.outputFileName)
          .onChange((value) => {
            this.outputFileName = value;
          });
      });

    // Import button
    new Setting(contentEl).addButton((button) =>
      button
        .setButtonText('Import')
        .setCta()
        .onClick(async () => {
          if (!this.file) {
            new Notice('Please select a DOCX file first');
            return;
          }

          if (!this.outputFileName) {
            new Notice('Please enter an output file name');
            return;
          }

          await this.importFile();
        })
    );

    // Cancel button
    new Setting(contentEl).addButton((button) =>
      button.setButtonText('Cancel').onClick(() => {
        this.close();
      })
    );
  }

  async importFile() {
    if (!this.file) return;

    try {
      new Notice('Importing DOCX file...');

      const arrayBuffer = await this.file.arrayBuffer();
      const result = await this.plugin.importDocx(
        arrayBuffer,
        this.outputFileName
      );

      if (result.warnings.length > 0) {
        console.warn('Import warnings:', result.warnings);
        new Notice(
          `Import completed with ${result.warnings.length} warning(s). Check console for details.`
        );
      } else {
        new Notice('DOCX file imported successfully!');
      }

      this.close();
    } catch (error) {
      console.error('Import error:', error);
      new Notice(`Import failed: ${error instanceof Error ? error.message : 'Unknown error'}`);
    }
  }

  onClose() {
    const { contentEl } = this;
    contentEl.empty();
  }
}