import {
  ParsedDocx,
  DocxElement,
  ConversionContext,
  ConversionResult,
  DocxImporterSettings,
} from '../types';
import { FormulaConverter } from './formulaConverter';
import { ImageConverter } from './imageConverter';
import { TableConverter } from './tableConverter';
import { ListConverter } from './listConverter';
import { Vault } from 'obsidian';

export class MainConverter {
  private formulaConverter: FormulaConverter;
  private imageConverter: ImageConverter;
  private tableConverter: TableConverter;
  private listConverter: ListConverter;
  private settings: DocxImporterSettings;
  private vault: Vault;

  constructor(settings: DocxImporterSettings, vault: Vault) {
    this.settings = settings;
    this.vault = vault;
    this.formulaConverter = new FormulaConverter();
    this.imageConverter = new ImageConverter(vault);
    this.tableConverter = new TableConverter(
      this.formulaConverter,
      this.imageConverter
    );
    this.listConverter = new ListConverter();
  }

  async convert(
    parsed: ParsedDocx,
    currentPath: string
  ): Promise<ConversionResult> {
    const context: ConversionContext = {
      settings: this.settings,
      images: parsed.images,
      styles: parsed.styles,
      numbering: parsed.numbering,
      relationships: parsed.relationships,
      listCounters: new Map(),
      currentPath,
    };

    const warnings: string[] = [];
    const markdownParts: string[] = [];

    // Group list items together
    const groupedElements = this.groupListItems(parsed.document.body);

    for (const element of groupedElements) {
      try {
        const markdown = await this.convertElement(element, context);
        if (markdown.trim()) {
          markdownParts.push(markdown);
        }
      } catch (error) {
        console.error('Conversion error:', error);
        warnings.push(`Failed to convert element: ${Array.isArray(element) ? 'list' : element.type}`);
      }
    }

    return {
      markdown: markdownParts.join('\n\n'),
      images: new Map(),
      warnings,
    };
  }

  private groupListItems(elements: DocxElement[]): (DocxElement | DocxElement[])[] {
    const grouped: (DocxElement | DocxElement[])[] = [];
    let currentList: DocxElement[] = [];
    let currentListId: number | null = null;

    for (const element of elements) {
      if (element.type === 'listItem') {
        const numId = element.properties?.numId;

        if (currentList.length === 0 || numId === currentListId) {
          currentList.push(element);
          currentListId = numId ?? null;
        } else {
          // Different list, save current and start new
          if (currentList.length > 0) {
            grouped.push(currentList);
          }
          currentList = [element];
          currentListId = numId ?? null;
        }
      } else {
        // Not a list item, flush current list
        if (currentList.length > 0) {
          grouped.push(currentList);
          currentList = [];
          currentListId = null;
        }
        grouped.push(element);
      }
    }

    // Don't forget remaining list items
    if (currentList.length > 0) {
      grouped.push(currentList);
    }

    return grouped;
  }

  private async convertElement(
    element: DocxElement | DocxElement[],
    context: ConversionContext
  ): Promise<string> {
    // Handle grouped list items
    if (Array.isArray(element)) {
      return this.listConverter.convert(element, context);
    }

    switch (element.type) {
      case 'heading':
        return this.convertHeading(element, context);
      case 'paragraph':
        return this.convertParagraph(element, context);
      case 'table':
        return this.tableConverter.convert(element, context);
      case 'listItem':
        return this.listConverter.convertItem(element, context);
      default:
        return '';
    }
  }

  private async convertHeading(
    element: DocxElement,
    context: ConversionContext
  ): Promise<string> {
    const level = element.properties?.headingLevel || 1;
    const prefix = '#'.repeat(Math.min(level, 6)) + ' ';

    const content = await this.convertParagraphContent(element, context);
    return prefix + content;
  }

  private async convertParagraph(
    element: DocxElement,
    context: ConversionContext
  ): Promise<string> {
    const content = await this.convertParagraphContent(element, context);

    // Handle alignment
    const alignment = element.properties?.alignment;
    if (alignment === 'center') {
      return `<center>${content}</center>`;
    } else if (alignment === 'right') {
      return `<div style="text-align: right">${content}</div>`;
    }

    return content;
  }

  private async convertParagraphContent(
    element: DocxElement,
    context: ConversionContext
  ): Promise<string> {
    const children = element.children || [];
    const parts: string[] = [];

    for (const child of children) {
      const converted = await this.convertInlineElement(child, context);
      parts.push(converted);
    }

    return parts.join('');
  }

  private async convertInlineElement(
    element: DocxElement,
    context: ConversionContext
  ): Promise<string> {
    switch (element.type) {
      case 'run':
        return this.convertRun(element, context);
      case 'hyperlink':
        return this.convertHyperlink(element, context);
      case 'formula':
        return this.convertFormula(element, context);
      case 'image':
        return this.convertImage(element, context);
      case 'break':
        return context.settings.preserveLineBreaks ? '\n' : ' ';
      case 'tab':
        return '    ';
      default:
        return '';
    }
  }

  private async convertRun(
    element: DocxElement,
    context: ConversionContext
  ): Promise<string> {
    const children = element.children || [];
    const parts: string[] = [];

    for (const child of children) {
      switch (child.type) {
        case 'text':
          parts.push(child.content || '');
          break;
        case 'break':
          parts.push(context.settings.preserveLineBreaks ? '\n' : ' ');
          break;
        case 'tab':
          parts.push('    ');
          break;
        case 'image':
          const imgMd = await this.convertImage(child, context);
          parts.push(imgMd);
          break;
        case 'formula':
          const formulaMd = this.convertFormula(child, context);
          parts.push(formulaMd);
          break;
      }
    }

    let text = parts.join('');

    // Apply formatting
    text = this.applyFormatting(text, element.properties || {});

    return text;
  }

  private applyFormatting(
    text: string,
    props: DocxElement['properties']
  ): string {
    if (!props || !text.trim()) return text;

    // Preserve leading/trailing whitespace
    const leadingSpace = text.match(/^\s*/)?.[0] || '';
    const trailingSpace = text.match(/\s*$/)?.[0] || '';
    let content = text.trim();

    if (!content) return text;

    // Apply formatting in specific order
    if (props.strikethrough) {
      content = `~~${content}~~`;
    }

    if (props.bold && props.italic) {
      content = `***${content}***`;
    } else if (props.bold) {
      content = `**${content}**`;
    } else if (props.italic) {
      content = `*${content}*`;
    }

    if (props.highlight) {
      content = `==${content}==`;
    }

    if (props.superscript) {
      content = `<sup>${content}</sup>`;
    }

    if (props.subscript) {
      content = `<sub>${content}</sub>`;
    }

    if (props.underline) {
      content = `<u>${content}</u>`;
    }

    return leadingSpace + content + trailingSpace;
  }

  private async convertHyperlink(
    element: DocxElement,
    context: ConversionContext
  ): Promise<string> {
    const rId = element.properties?.hyperlinkUrl;
    let url = '#';

    if (rId) {
      const relationship = context.relationships.get(rId);
      if (relationship) {
        url = relationship.target;
      }
    }

    // Convert children to get link text
    const children = element.children || [];
    const textParts: string[] = [];

    for (const child of children) {
      if (child.type === 'run') {
        const runChildren = child.children || [];
        for (const runChild of runChildren) {
          if (runChild.type === 'text') {
            textParts.push(runChild.content || '');
          }
        }
      }
    }

    const linkText = textParts.join('');

    // Handle internal links (bookmarks)
    if (url.startsWith('#')) {
      return `[[${url.substring(1)}|${linkText}]]`;
    }

    return `[${linkText}](${url})`;
  }

  private convertFormula(
    element: DocxElement,
    context: ConversionContext
  ): string {
    const omml = element.properties?.omml;
    if (!omml) return '';

    try {
      const latex = this.formulaConverter.convert(omml);

      // Determine if inline or block formula
      // For now, use inline format
      if (context.settings.formulaFormat === 'latex') {
        return `$${latex}$`;
      } else {
        // MathML format (if needed in future)
        return `$${latex}$`;
      }
    } catch (error) {
      console.error('Formula conversion error:', error);
      return '[Formula]';
    }
  }

  private async convertImage(
    element: DocxElement,
    context: ConversionContext
  ): Promise<string> {
    const imageId = element.properties?.imageId;
    if (!imageId) return '';

    try {
      const markdown = await this.imageConverter.saveImage(imageId, context);
      return markdown || '[Image]';
    } catch (error) {
      console.error('Image conversion error:', error);
      return '[Image]';
    }
  }

  // Public method to convert block formulas (displayed equations)
  convertBlockFormula(omml: string): string {
    try {
      const latex = this.formulaConverter.convert(omml);
      return `$$\n${latex}\n$$`;
    } catch (error) {
      console.error('Block formula conversion error:', error);
      return '$$[Formula]$$';
    }
  }

  // Update settings
  updateSettings(settings: DocxImporterSettings): void {
    this.settings = settings;
  }
}