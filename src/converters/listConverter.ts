import { DocxElement, ConversionContext, NumberingLevel } from '../types';

export class ListConverter {
  convert(items: DocxElement[], context: ConversionContext): string {
    const lines: string[] = [];
    const counterKey = this.getCounterKey(items);

    // Initialize or get counter for this list
    if (!context.listCounters.has(counterKey)) {
      context.listCounters.set(counterKey, []);
    }

    for (const item of items) {
      const line = this.convertItem(item, context);
      if (line) lines.push(line);
    }

    return lines.join('\n');
  }

  convertItem(item: DocxElement, context: ConversionContext): string {
    const props = item.properties || {};
    const numId = props.numId;
    const ilvl = props.ilvl ?? 0;

    // Determine list type
    const listInfo = this.getListInfo(numId, ilvl, context);

    // Get indent
    const indent = '  '.repeat(ilvl);

    // Get prefix
    let prefix: string;
    if (listInfo.isBullet) {
      prefix = '- ';
    } else {
      const counter = this.getListCounter(numId, ilvl, context);
      prefix = `${counter}. `;
    }

    // Convert content
    const content = this.convertItemContent(item, context);

    return `${indent}${prefix}${content}`;
  }

  private getListInfo(
    numId: number | undefined,
    ilvl: number,
    context: ConversionContext
  ): { isBullet: boolean; level: NumberingLevel | null } {
    if (numId === undefined || !context.numbering) {
      return { isBullet: true, level: null };
    }

    const numInstance = context.numbering.nums.get(numId);
    if (!numInstance) {
      return { isBullet: true, level: null };
    }

    const abstractNum = context.numbering.abstractNums.get(
      numInstance.abstractNumId
    );
    if (!abstractNum) {
      return { isBullet: true, level: null };
    }

    const level = abstractNum.levels.get(ilvl);
    if (!level) {
      return { isBullet: true, level: null };
    }

    const isBullet = this.isBulletFormat(level.numFmt);
    return { isBullet, level };
  }

  private isBulletFormat(numFmt: string): boolean {
    const bulletFormats = ['bullet', 'none'];
    return bulletFormats.includes(numFmt.toLowerCase());
  }

  private getListCounter(
    numId: number | undefined,
    ilvl: number,
    context: ConversionContext
  ): number {
    if (numId === undefined) return 1;

    const counterKey = `${numId}`;
    let counters = context.listCounters.get(counterKey);

    if (!counters) {
      counters = [];
      context.listCounters.set(counterKey, counters);
    }

    // Ensure array is long enough
    while (counters.length <= ilvl) {
      counters.push(0);
    }

    // Increment counter at this level
    counters[ilvl]++;

    // Reset deeper levels
    for (let i = ilvl + 1; i < counters.length; i++) {
      counters[i] = 0;
    }

    return counters[ilvl];
  }

  private convertItemContent(
    item: DocxElement,
    context: ConversionContext
  ): string {
    const runs = item.children || [];
    let result = '';

    for (const run of runs) {
      if (run.type === 'run') {
        result += this.convertRun(run);
      } else if (run.type === 'hyperlink') {
        result += this.convertHyperlink(run, context);
      }
    }

    return result.trim();
  }

  private convertRun(run: DocxElement): string {
    const children = run.children || [];
    let text = '';

    for (const child of children) {
      if (child.type === 'text') {
        text += child.content || '';
      }
    }

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

    return text;
  }

  private convertHyperlink(
    hyperlink: DocxElement,
    context: ConversionContext
  ): string {
    const rId = hyperlink.properties?.hyperlinkUrl;
    const relationship = rId ? context.relationships.get(rId) : null;
    const url = relationship?.target || '#';

    let text = '';
    for (const run of hyperlink.children || []) {
      text += this.convertRun(run);
    }

    return `[${text}](${url})`;
  }

  private getCounterKey(items: DocxElement[]): string {
    const firstItem = items[0];
    if (firstItem?.properties?.numId !== undefined) {
      return `list_${firstItem.properties.numId}`;
    }
    return 'default_list';
  }

  resetCounters(context: ConversionContext): void {
    context.listCounters.clear();
  }
}