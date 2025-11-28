import { App, PluginSettingTab, Setting } from 'obsidian';
import DocxImporterPlugin from '../main';
import { DocxImporterSettings } from '../types';

export class DocxImporterSettingsTab extends PluginSettingTab {
  plugin: DocxImporterPlugin;

  constructor(app: App, plugin: DocxImporterPlugin) {
    super(app, plugin);
    this.plugin = plugin;
  }

  display(): void {
    const { containerEl } = this;
    containerEl.empty();

    containerEl.createEl('h2', { text: 'DOCX Importer Settings' });

    // Image Settings Section
    containerEl.createEl('h3', { text: 'Image Settings' });

    new Setting(containerEl)
      .setName('Image folder')
      .setDesc('Folder where imported images will be saved')
      .addText((text) =>
        text
          .setPlaceholder('attachments')
          .setValue(this.plugin.settings.imageFolder)
          .onChange(async (value) => {
            this.plugin.settings.imageFolder = value || 'attachments';
            await this.plugin.saveSettings();
          })
      );

    new Setting(containerEl)
      .setName('Create image folder')
      .setDesc('Automatically create the image folder if it does not exist')
      .addToggle((toggle) =>
        toggle
          .setValue(this.plugin.settings.createImageFolder)
          .onChange(async (value) => {
            this.plugin.settings.createImageFolder = value;
            await this.plugin.saveSettings();
          })
      );

    new Setting(containerEl)
      .setName('Use relative image paths')
      .setDesc('Use relative paths for image references (recommended for portability)')
      .addToggle((toggle) =>
        toggle
          .setValue(this.plugin.settings.useRelativeImagePaths)
          .onChange(async (value) => {
            this.plugin.settings.useRelativeImagePaths = value;
            await this.plugin.saveSettings();
          })
      );

    // Formula Settings Section
    containerEl.createEl('h3', { text: 'Formula Settings' });

    new Setting(containerEl)
      .setName('Formula format')
      .setDesc('Output format for mathematical formulas')
      .addDropdown((dropdown) =>
        dropdown
          .addOption('latex', 'LaTeX')
          .addOption('mathml', 'MathML')
          .setValue(this.plugin.settings.formulaFormat)
          .onChange(async (value: 'latex' | 'mathml') => {
            this.plugin.settings.formulaFormat = value;
            await this.plugin.saveSettings();
          })
      );

    // Table Settings Section
    containerEl.createEl('h3', { text: 'Table Settings' });

    new Setting(containerEl)
      .setName('Default table alignment')
      .setDesc('Default alignment for table columns')
      .addDropdown((dropdown) =>
        dropdown
          .addOption('left', 'Left')
          .addOption('center', 'Center')
          .addOption('right', 'Right')
          .setValue(this.plugin.settings.tableAlignment)
          .onChange(async (value: 'left' | 'center' | 'right') => {
            this.plugin.settings.tableAlignment = value;
            await this.plugin.saveSettings();
          })
      );

    // Formatting Settings Section
    containerEl.createEl('h3', { text: 'Formatting Settings' });

    new Setting(containerEl)
      .setName('Preserve line breaks')
      .setDesc('Keep line breaks from the original document')
      .addToggle((toggle) =>
        toggle
          .setValue(this.plugin.settings.preserveLineBreaks)
          .onChange(async (value) => {
            this.plugin.settings.preserveLineBreaks = value;
            await this.plugin.saveSettings();
          })
      );

    new Setting(containerEl)
      .setName('Handle list numbering')
      .setDesc('Properly handle numbered and bulleted lists')
      .addToggle((toggle) =>
        toggle
          .setValue(this.plugin.settings.handleNumbering)
          .onChange(async (value) => {
            this.plugin.settings.handleNumbering = value;
            await this.plugin.saveSettings();
          })
      );
  }
}