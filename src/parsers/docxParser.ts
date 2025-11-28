import * as JSZip from 'jszip';
import { XMLParser } from 'fast-xml-parser';
import {
  ParsedDocx,
  DocxDocument,
  DocxElement,
  ElementProperties,
  ImageData,
  StyleDefinition,
  NumberingDefinition,
  AbstractNum,
  NumberingLevel,
  NumInstance,
  Relationship,
} from '../types';

export class DocxParser {
  private zip: JSZip | null = null;
  private xmlParser: XMLParser;

  constructor() {
    this.xmlParser = new XMLParser({
      ignoreAttributes: false,
      attributeNamePrefix: '@_',
      textNodeName: '#text',
      preserveOrder: true,
      trimValues: false,
    });
  }

  async parse(arrayBuffer: ArrayBuffer): Promise<ParsedDocx> {
    this.zip = await JSZip.loadAsync(arrayBuffer);

    const [document, images, styles, numbering, relationships] =
      await Promise.all([
        this.parseDocument(),
        this.parseImages(),
        this.parseStyles(),
        this.parseNumbering(),
        this.parseRelationships(),
      ]);

    return {
      document,
      images,
      styles,
      numbering,
      relationships,
    };
  }

  private async parseDocument(): Promise<DocxDocument> {
    const documentXml = await this.getFileContent('word/document.xml');
    if (!documentXml) {
      throw new Error('document.xml not found in DOCX');
    }

    const parsed = this.xmlParser.parse(documentXml);
    const body = this.extractBody(parsed);
    return { body: this.parseBodyElements(body) };
  }

  private extractBody(parsed: any): any[] {
    // Navigate through the parsed structure to find body
    for (const item of parsed) {
      if (item['w:document']) {
        const doc = item['w:document'];
        for (const docChild of doc) {
          if (docChild['w:body']) {
            return docChild['w:body'];
          }
        }
      }
    }
    return [];
  }

  private parseBodyElements(body: any[]): DocxElement[] {
    if (!body || !Array.isArray(body)) return [];

    const elements: DocxElement[] = [];

    for (const child of body) {
      const element = this.parseElement(child);
      if (element) {
        elements.push(element);
      }
    }

    return elements;
  }

  private parseElement(node: any): DocxElement | null {
    if (node['w:p']) {
      return this.parseParagraph(node['w:p']);
    } else if (node['w:tbl']) {
      return this.parseTable(node['w:tbl']);
    } else if (node['w:sdt']) {
      return this.parseStructuredDocument(node['w:sdt']);
    }
    return null;
  }

  private parseParagraph(paraContent: any[]): DocxElement {
    const properties = this.parseParagraphProperties(paraContent);
    const children = this.parseParagraphChildren(paraContent);

    // Check if it's a heading
    if (properties.headingLevel) {
      return {
        type: 'heading',
        properties,
        children,
      };
    }

    // Check if it's a list item
    if (properties.numId !== undefined) {
      return {
        type: 'listItem',
        properties,
        children,
      };
    }

    return {
      type: 'paragraph',
      properties,
      children,
    };
  }

  private parseParagraphProperties(paraContent: any[]): ElementProperties {
    const props: ElementProperties = {};

    for (const item of paraContent) {
      if (item['w:pPr']) {
        const pPr = item['w:pPr'];

        for (const prItem of pPr) {
          // Check for heading style
          if (prItem['w:pStyle']) {
            const styleAttr = prItem[':@'];
            const styleVal = styleAttr?.['@_w:val'];
            if (styleVal) {
              const headingMatch = styleVal.match(/Heading(\d)/i);
              if (headingMatch) {
                props.headingLevel = parseInt(headingMatch[1], 10);
              }
            }
          }

          // Check for list numbering
          if (prItem['w:numPr']) {
            const numPr = prItem['w:numPr'];
            for (const numPrItem of numPr) {
              if (numPrItem['w:ilvl']) {
                const ilvlAttr = numPrItem[':@'];
                props.ilvl = parseInt(ilvlAttr?.['@_w:val'] || '0', 10);
              }
              if (numPrItem['w:numId']) {
                const numIdAttr = numPrItem[':@'];
                props.numId = parseInt(numIdAttr?.['@_w:val'] || '0', 10);
              }
            }
          }

          // Alignment
          if (prItem['w:jc']) {
            const jcAttr = prItem[':@'];
            const val = jcAttr?.['@_w:val'];
            if (val === 'center' || val === 'right' || val === 'left') {
              props.alignment = val;
            }
          }
        }
      }
    }

    return props;
  }

  private parseParagraphChildren(paraContent: any[]): DocxElement[] {
    const children: DocxElement[] = [];

    for (const item of paraContent) {
      if (item['w:r']) {
        const run = this.parseRun(item['w:r']);
        if (run) children.push(run);
      } else if (item['w:hyperlink']) {
        const hyperlink = this.parseHyperlink(item);
        if (hyperlink) children.push(hyperlink);
      } else if (item['m:oMathPara'] || item['m:oMath']) {
        const formula = this.parseFormula(item);
        if (formula) children.push(formula);
      }
    }

    return children;
  }

  private parseRun(runContent: any[]): DocxElement | null {
    const properties = this.parseRunProperties(runContent);
    const children: DocxElement[] = [];

    for (const item of runContent) {
      if (item['w:t']) {
        const textContent = this.extractText(item['w:t']);
        children.push({
          type: 'text',
          content: textContent,
        });
      } else if (item['w:br']) {
        children.push({ type: 'break' });
      } else if (item['w:tab']) {
        children.push({ type: 'tab' });
      } else if (item['w:drawing']) {
        const image = this.parseDrawing(item['w:drawing']);
        if (image) children.push(image);
      } else if (item['m:oMath']) {
        const formula = this.parseFormula(item);
        if (formula) children.push(formula);
      }
    }

    if (children.length === 0) return null;

    return {
      type: 'run',
      properties,
      children,
    };
  }

  private extractText(textNode: any): string {
    if (Array.isArray(textNode)) {
      return textNode.map((t) => this.extractText(t)).join('');
    }
    if (typeof textNode === 'string') {
      return textNode;
    }
    if (textNode && textNode['#text'] !== undefined) {
      return textNode['#text'];
    }
    return '';
  }

  private parseRunProperties(runContent: any[]): ElementProperties {
    const props: ElementProperties = {};

    for (const item of runContent) {
      if (item['w:rPr']) {
        const rPr = item['w:rPr'];

        for (const prItem of rPr) {
          if (prItem['w:b'] !== undefined) {
            const attr = prItem[':@'];
            props.bold = attr?.['@_w:val'] !== 'false';
            if (prItem['w:b'] === '' || Array.isArray(prItem['w:b'])) {
              props.bold = true;
            }
          }
          if (prItem['w:i'] !== undefined) {
            const attr = prItem[':@'];
            props.italic = attr?.['@_w:val'] !== 'false';
            if (prItem['w:i'] === '' || Array.isArray(prItem['w:i'])) {
              props.italic = true;
            }
          }
          if (prItem['w:u'] !== undefined) {
            props.underline = true;
          }
          if (prItem['w:strike'] !== undefined) {
            props.strikethrough = true;
          }
          if (prItem['w:vertAlign']) {
            const attr = prItem[':@'];
            const vertVal = attr?.['@_w:val'];
            if (vertVal === 'superscript') props.superscript = true;
            if (vertVal === 'subscript') props.subscript = true;
          }
          if (prItem['w:highlight']) {
            const attr = prItem[':@'];
            props.highlight = attr?.['@_w:val'];
          }
          if (prItem['w:color']) {
            const attr = prItem[':@'];
            props.color = attr?.['@_w:val'];
          }
        }
      }
    }

    return props;
  }

  private parseHyperlink(node: any): DocxElement | null {
    const attr = node[':@'];
    const rId = attr?.['@_r:id'];
    const hyperlinkContent = node['w:hyperlink'];
    const children: DocxElement[] = [];

    if (Array.isArray(hyperlinkContent)) {
      for (const item of hyperlinkContent) {
        if (item['w:r']) {
          const run = this.parseRun(item['w:r']);
          if (run) children.push(run);
        }
      }
    }

    return {
      type: 'hyperlink',
      properties: { hyperlinkUrl: rId },
      children,
    };
  }

  private parseDrawing(drawingContent: any[]): DocxElement | null {
    // Navigate through drawing structure to find image reference
    const blipEmbed = this.findBlipEmbed(drawingContent);
    if (!blipEmbed) return null;

    return {
      type: 'image',
      properties: { imageId: blipEmbed },
    };
  }

  private findBlipEmbed(content: any): string | null {
    if (!content) return null;

    if (Array.isArray(content)) {
      for (const item of content) {
        const result = this.findBlipEmbed(item);
        if (result) return result;
      }
    }

    if (typeof content === 'object') {
      // Check for blip
      if (content['a:blip'] !== undefined) {
        const attr = content[':@'];
        return attr?.['@_r:embed'] || null;
      }

      // Recursively search all properties
      for (const key of Object.keys(content)) {
        if (key !== ':@') {
          const result = this.findBlipEmbed(content[key]);
          if (result) return result;
        }
      }
    }

    return null;
  }

  private parseFormula(node: any): DocxElement {
    // Store the raw OMML for later conversion
    return {
      type: 'formula',
      properties: {
        omml: JSON.stringify(node),
      },
    };
  }

  private parseTable(tableContent: any[]): DocxElement {
    const rows: DocxElement[] = [];

    for (const item of tableContent) {
      if (item['w:tr']) {
        rows.push(this.parseTableRow(item['w:tr']));
      }
    }

    return {
      type: 'table',
      children: rows,
    };
  }

  private parseTableRow(rowContent: any[]): DocxElement {
    const cells: DocxElement[] = [];

    for (const item of rowContent) {
      if (item['w:tc']) {
        cells.push(this.parseTableCell(item['w:tc']));
      }
    }

    return {
      type: 'tableRow',
      children: cells,
    };
  }

  private parseTableCell(cellContent: any[]): DocxElement {
    const properties = this.parseTableCellProperties(cellContent);
    const children: DocxElement[] = [];

    for (const item of cellContent) {
      if (item['w:p']) {
        const paragraph = this.parseParagraph(item['w:p']);
        if (paragraph) children.push(paragraph);
      }
    }

    return {
      type: 'tableCell',
      properties,
      children,
    };
  }

  private parseTableCellProperties(cellContent: any[]): ElementProperties {
    const props: ElementProperties = {};

    for (const item of cellContent) {
      if (item['w:tcPr']) {
        const tcPr = item['w:tcPr'];

        for (const prItem of tcPr) {
          if (prItem['w:gridSpan']) {
            const attr = prItem[':@'];
            props.colspan = parseInt(attr?.['@_w:val'] || '1', 10);
          }
          if (prItem['w:vMerge']) {
            const attr = prItem[':@'];
            const val = attr?.['@_w:val'];
            if (val === 'restart') {
              props.rowspan = 1;
            }
          }
        }
      }
    }

    return props;
  }

  private parseStructuredDocument(sdtContent: any[]): DocxElement | null {
    for (const item of sdtContent) {
      if (item['w:sdtContent']) {
        const content = item['w:sdtContent'];
        for (const contentItem of content) {
          const element = this.parseElement(contentItem);
          if (element) return element;
        }
      }
    }
    return null;
  }

  private async parseImages(): Promise<Map<string, ImageData>> {
    const images = new Map<string, ImageData>();

    if (!this.zip) return images;

    const files = Object.keys(this.zip.files).filter((f) =>
      f.startsWith('word/media/')
    );

    for (const filePath of files) {
      const file = this.zip.file(filePath);
      if (!file) continue;

      const data = await file.async('arraybuffer');
      const fileName = filePath.split('/').pop() || '';
      const extension = fileName.split('.').pop()?.toLowerCase() || '';

      const contentType = this.getImageContentType(extension);
      const relPath = filePath.replace('word/', '');

      images.set(relPath, {
        data,
        contentType,
        fileName,
        extension,
      });
    }

    return images;
  }

  private getImageContentType(extension: string): string {
    const types: Record<string, string> = {
      png: 'image/png',
      jpg: 'image/jpeg',
      jpeg: 'image/jpeg',
      gif: 'image/gif',
      bmp: 'image/bmp',
      svg: 'image/svg+xml',
      webp: 'image/webp',
      emf: 'image/emf',
      wmf: 'image/wmf',
    };
    return types[extension] || 'application/octet-stream';
  }

  private async parseStyles(): Promise<Map<string, StyleDefinition>> {
    const styles = new Map<string, StyleDefinition>();

    const stylesXml = await this.getFileContent('word/styles.xml');
    if (!stylesXml) return styles;

    const parsed = this.xmlParser.parse(stylesXml);

    // Navigate to styles
    for (const item of parsed) {
      if (item['w:styles']) {
        const stylesContent = item['w:styles'];
        for (const styleItem of stylesContent) {
          if (styleItem['w:style']) {
            const attr = styleItem[':@'];
            const styleId = attr?.['@_w:styleId'];
            if (styleId) {
              styles.set(styleId, this.parseStyleDefinition(styleItem, attr));
            }
          }
        }
      }
    }

    return styles;
  }

  private parseStyleDefinition(node: any, attr: any): StyleDefinition {
    const styleId = attr?.['@_w:styleId'] || '';
    const type = attr?.['@_w:type'] || '';

    let name = '';
    let basedOn: string | undefined;

    const styleContent = node['w:style'];
    if (Array.isArray(styleContent)) {
      for (const item of styleContent) {
        if (item['w:name']) {
          const nameAttr = item[':@'];
          name = nameAttr?.['@_w:val'] || '';
        }
        if (item['w:basedOn']) {
          const basedOnAttr = item[':@'];
          basedOn = basedOnAttr?.['@_w:val'];
        }
      }
    }

    return {
      styleId,
      name,
      type,
      basedOn,
      properties: {},
    };
  }

  private async parseNumbering(): Promise<NumberingDefinition | null> {
    const numberingXml = await this.getFileContent('word/numbering.xml');
    if (!numberingXml) return null;

    const parsed = this.xmlParser.parse(numberingXml);

    const abstractNums = new Map<number, AbstractNum>();
    const nums = new Map<number, NumInstance>();

    for (const item of parsed) {
      if (item['w:numbering']) {
        const numberingContent = item['w:numbering'];

        for (const numItem of numberingContent) {
          if (numItem['w:abstractNum']) {
            const attr = numItem[':@'];
            const abstractNum = this.parseAbstractNum(numItem['w:abstractNum'], attr);
            abstractNums.set(abstractNum.abstractNumId, abstractNum);
          }
          if (numItem['w:num']) {
            const attr = numItem[':@'];
            const num = this.parseNumInstance(numItem['w:num'], attr);
            nums.set(num.numId, num);
          }
        }
      }
    }

    return { abstractNums, nums };
  }

  private parseAbstractNum(content: any[], attr: any): AbstractNum {
    const abstractNumId = parseInt(attr?.['@_w:abstractNumId'] || '0', 10);
    const levels = new Map<number, NumberingLevel>();

    for (const item of content) {
      if (item['w:lvl']) {
        const lvlAttr = item[':@'];
        const level = this.parseNumberingLevel(item['w:lvl'], lvlAttr);
        levels.set(level.ilvl, level);
      }
    }

    return { abstractNumId, levels };
  }

  private parseNumberingLevel(content: any[], attr: any): NumberingLevel {
    const ilvl = parseInt(attr?.['@_w:ilvl'] || '0', 10);
    let numFmt = 'decimal';
    let lvlText = '%1.';
    let start = 1;

    for (const item of content) {
      if (item['w:numFmt']) {
        const fmtAttr = item[':@'];
        numFmt = fmtAttr?.['@_w:val'] || 'decimal';
      }
      if (item['w:lvlText']) {
        const textAttr = item[':@'];
        lvlText = textAttr?.['@_w:val'] || '';
      }
      if (item['w:start']) {
        const startAttr = item[':@'];
        start = parseInt(startAttr?.['@_w:val'] || '1', 10);
      }
    }

    return { ilvl, numFmt, lvlText, start };
  }

  private parseNumInstance(content: any[], attr: any): NumInstance {
    const numId = parseInt(attr?.['@_w:numId'] || '0', 10);
    let abstractNumId = 0;

    for (const item of content) {
      if (item['w:abstractNumId']) {
        const absAttr = item[':@'];
        abstractNumId = parseInt(absAttr?.['@_w:val'] || '0', 10);
      }
    }

    return { numId, abstractNumId };
  }

  private async parseRelationships(): Promise<Map<string, Relationship>> {
    const relationships = new Map<string, Relationship>();

    const relsXml = await this.getFileContent('word/_rels/document.xml.rels');
    if (!relsXml) return relationships;

    const parsed = this.xmlParser.parse(relsXml);

    for (const item of parsed) {
      if (item['Relationships']) {
        const relsContent = item['Relationships'];

        for (const relItem of relsContent) {
          if (relItem['Relationship']) {
            const attr = relItem[':@'];
            const id = attr?.['@_Id'];
            const type = attr?.['@_Type'];
            const target = attr?.['@_Target'];

            if (id) {
              relationships.set(id, { id, type, target });
            }
          }
        }
      }
    }

    return relationships;
  }

  private async getFileContent(path: string): Promise<string | null> {
    if (!this.zip) return null;

    const file = this.zip.file(path);
    if (!file) return null;

    return await file.async('string');
  }
}