import { ImageData, ConversionContext } from '../types';
import { TFile, Vault } from 'obsidian';

export class ImageConverter {
  private vault: Vault;

  constructor(vault: Vault) {
    this.vault = vault;
  }

  async saveImage(
    imageId: string,
    context: ConversionContext
  ): Promise<string | null> {
    // Get the relationship for this image
    const relationship = context.relationships.get(imageId);
    if (!relationship) {
      console.warn(`No relationship found for image: ${imageId}`);
      return null;
    }

    // Get image data
    const imagePath = relationship.target.replace('../', '');
    const imageData = context.images.get(imagePath);

    if (!imageData) {
      console.warn(`No image data found for: ${imagePath}`);
      return null;
    }

    // Create image folder if needed
    const imageFolder = context.settings.imageFolder;
    if (context.settings.createImageFolder) {
      await this.ensureFolderExists(imageFolder);
    }

    // Generate unique filename
    const fileName = await this.generateUniqueFileName(
      imageFolder,
      imageData.fileName
    );

    // Save the image
    const fullPath = `${imageFolder}/${fileName}`;
    await this.vault.createBinary(fullPath, imageData.data);

    // Return markdown image reference
    if (context.settings.useRelativeImagePaths) {
      return `![[${fileName}]]`;
    } else {
      return `![[${fullPath}]]`;
    }
  }

  private async ensureFolderExists(folderPath: string): Promise<void> {
    const folder = this.vault.getAbstractFileByPath(folderPath);
    if (!folder) {
      await this.vault.createFolder(folderPath);
    }
  }

  private async generateUniqueFileName(
    folder: string,
    originalName: string
  ): Promise<string> {
    const baseName = originalName.replace(/\.[^.]+$/, '');
    const extension = originalName.split('.').pop() || 'png';

    let fileName = `${baseName}.${extension}`;
    let counter = 1;

    while (this.vault.getAbstractFileByPath(`${folder}/${fileName}`)) {
      fileName = `${baseName}_${counter}.${extension}`;
      counter++;
    }

    return fileName;
  }

  getImageMarkdown(imagePath: string, altText?: string): string {
    if (altText) {
      return `![[${imagePath}|${altText}]]`;
    }
    return `![[${imagePath}]]`;
  }
}