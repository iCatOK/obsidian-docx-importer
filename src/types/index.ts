export interface DocxImporterSettings {
  imageFolder: string;
  useRelativeImagePaths: boolean;
  formulaFormat: 'latex' | 'mathml';
  preserveLineBreaks: boolean;
  tableAlignment: 'left' | 'center' | 'right';
  defaultImageWidth: number;
  createImageFolder: boolean;
  handleNumbering: boolean;
}

export const DEFAULT_SETTINGS: DocxImporterSettings = {
  imageFolder: 'attachments',
  useRelativeImagePaths: true,
  formulaFormat: 'latex',
  preserveLineBreaks: true,
  tableAlignment: 'left',
  defaultImageWidth: 600,
  createImageFolder: true,
  handleNumbering: true,
};

export interface ParsedDocx {
  document: DocxDocument;
  images: Map<string, ImageData>;
  styles: Map<string, StyleDefinition>;
  numbering: NumberingDefinition | null;
  relationships: Map<string, Relationship>;
}

export interface DocxDocument {
  body: DocxElement[];
}

export interface DocxElement {
  type: ElementType;
  content?: string;
  children?: DocxElement[];
  properties?: ElementProperties;
}

export type ElementType =
  | 'paragraph'
  | 'run'
  | 'text'
  | 'heading'
  | 'table'
  | 'tableRow'
  | 'tableCell'
  | 'image'
  | 'formula'
  | 'list'
  | 'listItem'
  | 'hyperlink'
  | 'bookmark'
  | 'break'
  | 'tab';

export interface ElementProperties {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strikethrough?: boolean;
  superscript?: boolean;
  subscript?: boolean;
  highlight?: string;
  color?: string;
  fontSize?: number;
  fontFamily?: string;
  alignment?: 'left' | 'center' | 'right' | 'justify';
  headingLevel?: number;
  listLevel?: number;
  listType?: 'bullet' | 'number';
  numId?: number;
  ilvl?: number;
  hyperlinkUrl?: string;
  imageId?: string;
  colspan?: number;
  rowspan?: number;
  cellWidth?: number;
  verticalAlign?: string;
  omml?: string;
}

export interface ImageData {
  data: ArrayBuffer;
  contentType: string;
  fileName: string;
  extension: string;
}

export interface StyleDefinition {
  styleId: string;
  name: string;
  type: string;
  basedOn?: string;
  properties: ElementProperties;
}

export interface NumberingDefinition {
  abstractNums: Map<number, AbstractNum>;
  nums: Map<number, NumInstance>;
}

export interface AbstractNum {
  abstractNumId: number;
  levels: Map<number, NumberingLevel>;
}

export interface NumberingLevel {
  ilvl: number;
  numFmt: string;
  lvlText: string;
  start: number;
}

export interface NumInstance {
  numId: number;
  abstractNumId: number;
}

export interface Relationship {
  id: string;
  type: string;
  target: string;
}

export interface ConversionContext {
  settings: DocxImporterSettings;
  images: Map<string, ImageData>;
  styles: Map<string, StyleDefinition>;
  numbering: NumberingDefinition | null;
  relationships: Map<string, Relationship>;
  listCounters: Map<string, number[]>;
  currentPath: string;
}

export interface ConversionResult {
  markdown: string;
  images: Map<string, { data: ArrayBuffer; fileName: string }>;
  warnings: string[];
}