import { XMLParser } from 'fast-xml-parser';

export class OmmlParser {
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

  parse(ommlJson: string): OmmlElement {
    try {
      const node = JSON.parse(ommlJson);
      return this.parseNode(node);
    } catch (error) {
      console.error('OMML parse error:', error);
      return {
        type: 'unknown',
        tagName: '',
        children: [],
        text: null,
        attributes: {},
      };
    }
  }

  private parseNode(node: any): OmmlElement {
    // Handle different node structures from fast-xml-parser
    if (Array.isArray(node)) {
      // If it's an array, create a container element
      const children = node.map((child) => this.parseNode(child));
      return {
        type: 'math',
        tagName: 'm:oMath',
        children,
        text: null,
        attributes: {},
      };
    }

    // Find the tag name (first key that's not :@ or #text)
    let tagName = '';
    let content: any = null;
    let attributes: Record<string, string> = {};

    for (const key of Object.keys(node)) {
      if (key === ':@') {
        // Attributes
        for (const attrKey of Object.keys(node[key])) {
          const cleanKey = attrKey.replace('@_', '');
          attributes[cleanKey] = node[key][attrKey];
        }
      } else if (key === '#text') {
        // Text content
        content = node[key];
      } else {
        // Element tag
        tagName = key;
        content = node[key];
      }
    }

    const children: OmmlElement[] = [];
    let text: string | null = null;

    if (typeof content === 'string') {
      text = content;
    } else if (Array.isArray(content)) {
      for (const child of content) {
        children.push(this.parseNode(child));
      }
    } else if (content && typeof content === 'object') {
      children.push(this.parseNode(content));
    }

    return {
      type: this.getElementType(tagName),
      tagName,
      children,
      text,
      attributes,
    };
  }

  private getElementType(tagName: string): OmmlElementType {
    const typeMap: Record<string, OmmlElementType> = {
      'm:oMath': 'math',
      'm:oMathPara': 'mathPara',
      'm:r': 'run',
      'm:t': 'text',
      'm:f': 'fraction',
      'm:num': 'numerator',
      'm:den': 'denominator',
      'm:rad': 'radical',
      'm:deg': 'degree',
      'm:e': 'element',
      'm:sup': 'superscript',
      'm:sub': 'subscript',
      'm:sSup': 'superscriptContainer',
      'm:sSub': 'subscriptContainer',
      'm:sSubSup': 'subSupContainer',
      'm:nary': 'nary',
      'm:limLow': 'limitLow',
      'm:limUpp': 'limitUpper',
      'm:lim': 'limit',
      'm:m': 'matrix',
      'm:mr': 'matrixRow',
      'm:d': 'delimiter',
      'm:dPr': 'delimiterProps',
      'm:begChr': 'beginChar',
      'm:endChr': 'endChar',
      'm:func': 'function',
      'm:fName': 'functionName',
      'm:eqArr': 'equationArray',
      'm:acc': 'accent',
      'm:accPr': 'accentProps',
      'm:chr': 'character',
      'm:bar': 'bar',
      'm:box': 'box',
      'm:groupChr': 'groupChar',
      'm:borderBox': 'borderBox',
      'm:sPre': 'preSuperSubscript',
      'm:naryPr': 'naryProps',
      'm:ctrlPr': 'controlProps',
      'm:rPr': 'runProps',
      '': 'unknown',
    };

    return typeMap[tagName] || 'unknown';
  }
}

export interface OmmlElement {
  type: OmmlElementType;
  tagName: string;
  children: OmmlElement[];
  text: string | null;
  attributes: Record<string, string>;
}

export type OmmlElementType =
  | 'math'
  | 'mathPara'
  | 'run'
  | 'text'
  | 'fraction'
  | 'numerator'
  | 'denominator'
  | 'radical'
  | 'degree'
  | 'element'
  | 'superscript'
  | 'subscript'
  | 'superscriptContainer'
  | 'subscriptContainer'
  | 'subSupContainer'
  | 'nary'
  | 'naryProps'
  | 'limitLow'
  | 'limitUpper'
  | 'limit'
  | 'matrix'
  | 'matrixRow'
  | 'delimiter'
  | 'delimiterProps'
  | 'beginChar'
  | 'endChar'
  | 'function'
  | 'functionName'
  | 'equationArray'
  | 'accent'
  | 'accentProps'
  | 'character'
  | 'bar'
  | 'box'
  | 'groupChar'
  | 'borderBox'
  | 'preSuperSubscript'
  | 'controlProps'
  | 'runProps'
  | 'unknown';