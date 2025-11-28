import { DocxElement, ConversionContext } from '../types';
import { FormulaConverter } from './formulaConverter';
import { ImageConverter } from './imageConverter';

export class TableConverter {
  private formulaConverter: FormulaConverter;
  private imageConverter: ImageConverter;

  constructor(formulaConverter: FormulaConverter, imageConverter: ImageConverter) {
    this.formulaConverter = formulaConverter;
    this.imageConverter = imageConverter;
  }

  async convert(table: DocxElement, context: ConversionContext): Promise<string> {
    const rows = table.children || [];
    if (rows.length === 0) return '';

    const processedRows: string[][] = [];
    let maxCols = 0;

    // Process each row
    for (const row of rows) {
      const cells = row.children || [];
      const processedCells: string[] = [];

      for (const cell of cells) {
        const cellContent = await this.convertCellContent(cell, context);
        const colspan = cell.properties?.colspan || 1;

        // Add cell content
        processedCells.push(cellContent);

        // Add empty cells for colspan
        for (let i = 1; i < colspan; i++) {
          processedCells.push('');
        }
      }

      processedRows.push(processedCells);
      maxCols = Math.max(maxCols, processedCells.length);
    }

    // Normalize row lengths
    for (const row of processedRows) {
      while (row.length < maxCols) {
        row.push('');
      }
    }

    // Build markdown table
    return this.buildMarkdownTable(processedRows, context);
  }

  private async convertCellContent(
    cell: DocxElement,
    context: ConversionContext
  ): Promise<string> {
    const paragraphs = cell.children || [];
    const contents: string[] = [];

    for (const para of paragraphs) {
      const content = await this.convertParagraphInCell(para, context);
      if (content.trim()) {
        contents.push(content.trim());
      }
    }

    // Join with <br> for multiple paragraphs in a cell
    return contents.join('<br>');
  }

  private async convertParagraphInCell(
    paragraph: DocxElement,
    context: ConversionContext
  ): Promise<string> {
    const runs = paragraph.children || [];
    let result = '';

    for (const run of runs) {
      if (run.type === 'formula') {
        const latex = this.formulaConverter.convert(run.properties?.omml || '');
        result += `$${latex}$`;
      } else if (run.type === 'image') {
        const imageMd = await this.imageConverter.saveImage(
          run.properties?.imageId || '',
          context
        );
        if (imageMd) result += imageMd;
      } else if (run.type === 'run') {
        result += await this.convertRunContent(run, context);
      }
    }

    return result;
  }

  private async convertRunContent(run: DocxElement, context: ConversionContext): Promise<string> {
    const children = run.children || [];
    const segments: string[] = [];

    for (const child of children) {
      if (child.type === 'text') {
        segments.push(child.content || '');
      } else if (child.type === 'break') {
        segments.push('<br>');
      } else if (child.type === 'tab') {
        segments.push('    ');
      } else if (child.type === 'image') {
        // handle images inside a run
        const imageMd = await this.imageConverter.saveImage(
          child.properties?.imageId || '',
          context
        );
        if (imageMd) segments.push(imageMd);
      } else if (child.type === 'formula') {
        const latex = this.formulaConverter.convert(child.properties?.omml || '');
        segments.push(`$${latex}$`);
      }
    }

    let text = segments.join('');

    // Apply formatting
    const props = run.properties || {};

    if (props.bold && props.italic) {
      text = `***${text}***`;
    } else if (props.bold) {
      text = `**${text}**`;
    } else if (props.italic) {
      text = `*${text}*`;
    }

    if (props.strikethrough) {
      text = `~~${text}~~`;
    }

    if (props.highlight) {
      text = `==${text}==`;
    }

    return text;
  }

  private buildMarkdownTable(
    rows: string[][],
    context: ConversionContext
  ): string {
    if (rows.length === 0) return '';

    const lines: string[] = [];
    const colCount = rows[0].length;

    // Header row
    lines.push('| ' + rows[0].map((c) => this.escapeCell(c)).join(' | ') + ' |');

    // Separator row with alignment
    const alignment = context.settings.tableAlignment;
    let separator: string;
    switch (alignment) {
      case 'center':
        separator = ':---:';
        break;
      case 'right':
        separator = '---:';
        break;
      default:
        separator = '---';
    }
    lines.push('| ' + Array(colCount).fill(separator).join(' | ') + ' |');

    // Data rows
    for (let i = 1; i < rows.length; i++) {
      lines.push(
        '| ' + rows[i].map((c) => this.escapeCell(c)).join(' | ') + ' |'
      );
    }

    return lines.join('\n');
  }

  private escapeCell(content: string): string {
    // Escape pipe characters in cell content
    return content.replace(/\|/g, '\\|').replace(/\n/g, '<br>');
  }
}