import { OmmlParser, OmmlElement, OmmlElementType } from '../parsers/ommlParser';

export class FormulaConverter {
  private parser: OmmlParser;

  // Unicode to LaTeX mapping
  private unicodeToLatex: Map<string, string> = new Map([
    // Greek letters
    ['α', '\\alpha '],
    ['β', '\\beta '],
    ['γ', '\\gamma '],
    ['δ', '\\delta '],
    ['ε', '\\epsilon '],
    ['ζ', '\\zeta '],
    ['η', '\\eta '],
    ['θ', '\\theta '],
    ['ι', '\\iota'],
    ['κ', '\\kappa '],
    ['λ', '\\lambda '],
    ['μ', '\\mu '],
    ['ν', '\\nu '],
    ['ξ', '\\xi '],
    ['π', '\\pi '],
    ['ρ', '\\rho '],
    ['σ', '\\sigma '],
    ['τ', '\\tau '],
    ['υ', '\\upsilon '],
    ['φ', '\\phi '],
    ['χ', '\\chi '],
    ['ψ', '\\psi '],
    ['ω', '\\omega '],
    ['Γ', '\\Gamma '],
    ['Δ', '\\Delta '],
    ['Θ', '\\Theta '],
    ['Λ', '\\Lambda '],
    ['Ξ', '\\Xi '],
    ['Π', '\\Pi'],
    ['Σ', '\\Sigma'],
    ['Υ', '\\Upsilon'],
    ['Φ', '\\Phi'],
    ['Ψ', '\\Psi'],
    ['Ω', '\\Omega'],
    // Operators
    ['×', '\\times'],
    ['÷', '\\div'],
    ['±', '\\pm'],
    ['∓', '\\mp'],
    ['·', '\\cdot'],
    ['°', '^\\circ'],
    ['∞', '\\infty'],
    ['≈', '\\approx'],
    ['≠', '\\neq'],
    ['≤', '\\leq'],
    ['≥', '\\geq'],
    ['≪', '\\ll'],
    ['≫', '\\gg'],
    ['∝', '\\propto'],
    ['≡', '\\equiv'],
    ['∼', '\\sim'],
    ['≃', '\\simeq'],
    ['≅', '\\cong'],
    // Set theory
    ['∈', '\\in'],
    ['∉', '\\notin'],
    ['⊂', '\\subset'],
    ['⊃', '\\supset'],
    ['⊆', '\\subseteq'],
    ['⊇', '\\supseteq'],
    ['∪', '\\cup'],
    ['∩', '\\cap'],
    ['∅', '\\emptyset'],
    ['∀', '\\forall'],
    ['∃', '\\exists'],
    ['∄', '\\nexists'],
    // Arrows
    ['→', '\\rightarrow'],
    ['←', '\\leftarrow'],
    ['↔', '\\leftrightarrow'],
    ['⇒', '\\Rightarrow'],
    ['⇐', '\\Leftarrow'],
    ['⇔', '\\Leftrightarrow'],
    ['↑', '\\uparrow'],
    ['↓', '\\downarrow'],
    ['↦', '\\mapsto'],
    // Calculus
    ['∂', '\\partial'],
    ['∇', '\\nabla'],
    ['∫', '\\int'],
    ['∬', '\\iint'],
    ['∭', '\\iiint'],
    ['∮', '\\oint'],
    ['∑', '\\sum'],
    ['∏', '\\prod'],
    ['√', '\\sqrt'],
    // Logic
    ['∧', '\\land'],
    ['∨', '\\lor'],
    ['¬', '\\neg'],
    ['⊕', '\\oplus'],
    ['⊗', '\\otimes'],
    ['⊥', '\\perp'],
    ['∥', '\\parallel'],
    // Misc
    ['ℕ', '\\mathbb{N}'],
    ['ℤ', '\\mathbb{Z}'],
    ['ℚ', '\\mathbb{Q}'],
    ['ℝ', '\\mathbb{R}'],
    ['ℂ', '\\mathbb{C}'],
    ['ℏ', '\\hbar'],
    ['ℓ', '\\ell'],
    ['′', "'"],
    ['″', "''"],
  ]);

  private functionNames: Map<string, string> = new Map([
    ['sin', '\\sin'],
    ['cos', '\\cos'],
    ['tan', '\\tan'],
    ['cot', '\\cot'],
    ['sec', '\\sec'],
    ['csc', '\\csc'],
    ['arcsin', '\\arcsin'],
    ['arccos', '\\arccos'],
    ['arctan', '\\arctan'],
    ['sinh', '\\sinh'],
    ['cosh', '\\cosh'],
    ['tanh', '\\tanh'],
    ['log', '\\log'],
    ['ln', '\\ln'],
    ['exp', '\\exp'],
    ['lim', '\\lim'],
    ['max', '\\max'],
    ['min', '\\min'],
    ['sup', '\\sup'],
    ['inf', '\\inf'],
    ['det', '\\det'],
    ['dim', '\\dim'],
    ['ker', '\\ker'],
    ['gcd', '\\gcd'],
    ['mod', '\\mod'],
    ['arg', '\\arg'],
    ['deg', '\\deg'],
  ]);

  private accentMap: Map<string, string> = new Map([
    ['̂', '\\hat'],
    ['̃', '\\tilde'],
    ['̄', '\\bar'],
    ['⃗', '\\vec'],
    ['̇', '\\dot'],
    ['̈', '\\ddot'],
    ['̆', '\\breve'],
    ['̌', '\\check'],
    ['ˆ', '\\hat'],
    ['˜', '\\tilde'],
    ['¯', '\\bar'],
    ['→', '\\vec'],
  ]);

  constructor() {
    this.parser = new OmmlParser();
  }

  convert(ommlJson: string): string {
    try {
      const element = this.parser.parse(ommlJson);
      const latex = this.elementToLatex(element);
      return this.cleanupLatex(latex);
    } catch (error) {
      console.error('Formula conversion error:', error);
      return '[Formula conversion error]';
    }
  }

  private elementToLatex(element: OmmlElement): string {
    switch (element.type) {
      case 'math':
      case 'mathPara':
        return this.convertMath(element);
      case 'run':
        return this.convertRun(element);
      case 'text':
        return this.convertText(element);
      case 'fraction':
        return this.convertFraction(element);
      case 'radical':
        return this.convertRadical(element);
      case 'superscriptContainer':
        return this.convertSuperscript(element);
      case 'subscriptContainer':
        return this.convertSubscript(element);
      case 'subSupContainer':
        return this.convertSubSup(element);
      case 'nary':
        return this.convertNary(element);
      case 'limitLow':
        return this.convertLimitLow(element);
      case 'limitUpper':
        return this.convertLimitUpper(element);
      case 'matrix':
        return this.convertMatrix(element);
      case 'delimiter':
        return this.convertDelimiter(element);
      case 'function':
        return this.convertFunction(element);
      case 'equationArray':
        return this.convertEquationArray(element);
      case 'accent':
        return this.convertAccent(element);
      case 'bar':
        return this.convertBar(element);
      case 'box':
        return this.convertBox(element);
      case 'groupChar':
        return this.convertGroupChar(element);
      case 'borderBox':
        return this.convertBorderBox(element);
      case 'preSuperSubscript':
        return this.convertPreSuperSubscript(element);
      case 'element':
      case 'numerator':
      case 'denominator':
      case 'degree':
      case 'superscript':
      case 'subscript':
      case 'limit':
      case 'matrixRow':
      case 'functionName':
        return this.convertChildren(element);
      case 'naryProps':
      case 'delimiterProps':
      case 'accentProps':
      case 'controlProps':
      case 'runProps':
        return ''; // Skip property elements
      default:
        return this.convertChildren(element);
    }
  }

  private convertMath(element: OmmlElement): string {
    return this.convertChildren(element);
  }

  private convertRun(element: OmmlElement): string {
    return this.convertChildren(element);
  }

  private convertText(element: OmmlElement): string {
    let text = element.text || '';
    
    // Also check children for text
    for (const child of element.children) {
      if (child.text) {
        text += child.text;
      } else if (child.type === 'text') {
        text += this.convertText(child);
      }
    }

    // Convert each character
    let result = '';
    let prevWasOperator = false;

    for (const char of text) {
      if (this.unicodeToLatex.has(char)) {
        result += this.unicodeToLatex.get(char);
        prevWasOperator = ['+', '-', '=', '×', '÷'].includes(char);
      } else if (char === ' ') {
        // Добавляйте пробелы только где нужно
        if (!prevWasOperator && result && !result.endsWith(' ')) {
          result += ' ';
        }
      } else if (['+', '-', '='].includes(char)) {
        result += ` ${char} `;
        prevWasOperator = true;
      } else {
        result += char;
        prevWasOperator = false;
      }
    }

    return result.trim();
  }

  private convertFraction(element: OmmlElement): string {
    let numerator = '';
    let denominator = '';

    for (const child of element.children) {
      if (child.type === 'numerator' || child.tagName === 'm:num') {
        numerator = this.convertChildren(child);
      } else if (child.type === 'denominator' || child.tagName === 'm:den') {
        denominator = this.convertChildren(child);
      }
    }

    if (numerator.includes('+') || numerator.includes('-') || numerator.includes('*')) {
      numerator = `{${numerator}}`;
    }

    if (denominator.includes('+') || denominator.includes('-') || denominator.includes('*')) {
      denominator = `{${denominator}}`;
    }

    return `\\frac{${numerator}}{${denominator}}`;
  }

  private convertRadical(element: OmmlElement): string {
    let degree = '';
    let base = '';

    for (const child of element.children) {
      if (child.type === 'degree' || child.tagName === 'm:deg') {
        degree = this.convertChildren(child);
      } else if (child.type === 'element' || child.tagName === 'm:e') {
        base = this.convertChildren(child);
      }
    }

    if (degree && degree.trim() !== '' && degree.trim() !== '2') {
      return `\\sqrt[${degree}]{${base}}`;
    }
    return `\\sqrt{${base}}`;
  }

  private convertSuperscript(element: OmmlElement): string {
    let base = '';
    let sup = '';

    for (const child of element.children) {
      if (child.type === 'element' || child.tagName === 'm:e') {
        base = this.convertChildren(child);
      } else if (child.type === 'superscript' || child.tagName === 'm:sup') {
        sup = this.convertChildren(child);
      }
    }

    return `${base}^{${sup}}`;
  }

  private convertSubscript(element: OmmlElement): string {
    let base = '';
    let sub = '';

    for (const child of element.children) {
      if (child.type === 'element' || child.tagName === 'm:e') {
        base = this.convertChildren(child);
      } else if (child.type === 'subscript' || child.tagName === 'm:sub') {
        sub = this.convertChildren(child);
      }
    }

    return `${base}_{${sub}}`;
  }

  private convertSubSup(element: OmmlElement): string {
    let base = '';
    let sub = '';
    let sup = '';

    for (const child of element.children) {
      if (child.type === 'element' || child.tagName === 'm:e') {
        base = this.convertChildren(child);
      } else if (child.type === 'subscript' || child.tagName === 'm:sub') {
        sub = this.convertChildren(child);
      } else if (child.type === 'superscript' || child.tagName === 'm:sup') {
        sup = this.convertChildren(child);
      }
    }

    return `${base}_{${sub}}^{${sup}}`;
  }

  private convertNary(element: OmmlElement): string {
    let operator = '\\int';
    let sub = '';
    let sup = '';
    let base = '';

    for (const child of element.children) {
      if (child.type === 'naryProps' || child.tagName === 'm:naryPr') {
        for (const prop of child.children) {
          if (prop.tagName === 'm:chr' || prop.type === 'character') {
            const chr = prop.attributes['m:val'] || prop.attributes['val'];
            if (chr) {
              operator = this.convertNaryChar(chr);
            }
          }
        }
      } else if (child.type === 'subscript' || child.tagName === 'm:sub') {
        sub = this.convertChildren(child);
      } else if (child.type === 'superscript' || child.tagName === 'm:sup') {
        sup = this.convertChildren(child);
      } else if (child.type === 'element' || child.tagName === 'm:e') {
        base = this.convertChildren(child);
      }
    }

    let result = operator;
    if (sub) result += `_{${sub}}`;
    if (sup) result += `^{${sup}}`;
    result += ` ${base}`;

    return result;
  }

  private convertNaryChar(char: string): string {
    const naryMap: Record<string, string> = {
      '∫': '\\int',
      '∬': '\\iint',
      '∭': '\\iiint',
      '∮': '\\oint',
      '∑': '\\sum',
      '∏': '\\prod',
      '⋃': '\\bigcup',
      '⋂': '\\bigcap',
      '⋁': '\\bigvee',
      '⋀': '\\bigwedge',
    };

    return naryMap[char] || '\\int';
  }

  private convertLimitLow(element: OmmlElement): string {
    let base = '';
    let lim = '';

    for (const child of element.children) {
      if (child.type === 'element' || child.tagName === 'm:e') {
        base = this.convertChildren(child);
      } else if (child.type === 'limit' || child.tagName === 'm:lim') {
        lim = this.convertChildren(child);
      }
    }

    return `${base}_{${lim}}`;
  }

  private convertLimitUpper(element: OmmlElement): string {
    let base = '';
    let lim = '';

    for (const child of element.children) {
      if (child.type === 'element' || child.tagName === 'm:e') {
        base = this.convertChildren(child);
      } else if (child.type === 'limit' || child.tagName === 'm:lim') {
        lim = this.convertChildren(child);
      }
    }

    return `${base}^{${lim}}`;
  }

  private convertMatrix(element: OmmlElement): string {
    const rows: string[] = [];

    for (const child of element.children) {
      if (child.type === 'matrixRow' || child.tagName === 'm:mr') {
        const cells: string[] = [];
        for (const cell of child.children) {
          if (cell.type === 'element' || cell.tagName === 'm:e') {
            cells.push(this.convertChildren(cell));
          }
        }
        rows.push(cells.join(' & '));
      }
    }

    return `\\begin{pmatrix} ${rows.join(' \\\\ ')} \\end{pmatrix}`;
  }

  private convertDelimiter(element: OmmlElement): string {
    let beginChar = '(';
    let endChar = ')';
    const contents: string[] = [];

    for (const child of element.children) {
      if (child.type === 'delimiterProps' || child.tagName === 'm:dPr') {
        for (const prop of child.children) {
          if (prop.tagName === 'm:begChr' || prop.type === 'beginChar') {
            beginChar = prop.attributes['m:val'] || prop.attributes['val'] || '(';
          }
          if (prop.tagName === 'm:endChr' || prop.type === 'endChar') {
            endChar = prop.attributes['m:val'] || prop.attributes['val'] || ')';
          }
        }
      } else if (child.type === 'element' || child.tagName === 'm:e') {
        contents.push(this.convertChildren(child));
      }
    }

    const leftDelim = this.convertDelimiterChar(beginChar);
    const rightDelim = this.convertDelimiterChar(endChar);

    return `\\left${leftDelim} ${contents.join(', ')} \\right${rightDelim}`;
  }

  private convertDelimiterChar(char: string): string {
    const delimMap: Record<string, string> = {
      '(': '(',
      ')': ')',
      '[': '[',
      ']': ']',
      '{': '\\{',
      '}': '\\}',
      '|': '|',
      '‖': '\\|',
      '⌈': '\\lceil',
      '⌉': '\\rceil',
      '⌊': '\\lfloor',
      '⌋': '\\rfloor',
      '⟨': '\\langle',
      '⟩': '\\rangle',
      '': '.',
    };

    return delimMap[char] || char;
  }

  private convertFunction(element: OmmlElement): string {
    let funcName = '';
    let arg = '';

    for (const child of element.children) {
      if (child.type === 'functionName' || child.tagName === 'm:fName') {
        funcName = this.convertChildren(child).trim();
      } else if (child.type === 'element' || child.tagName === 'm:e') {
        arg = this.convertChildren(child);
      }
    }

    if (this.functionNames.has(funcName)) {
      return `${this.functionNames.get(funcName)} ${arg}`;
    }

    return `\\operatorname{${funcName}} ${arg}`;
  }

  private convertEquationArray(element: OmmlElement): string {
    const equations: string[] = [];

    for (const child of element.children) {
      if (child.type === 'element' || child.tagName === 'm:e') {
        equations.push(this.convertChildren(child));
      }
    }

    if (equations.length === 1) {
      return equations[0];
    }

    return `\\begin{aligned} ${equations.join(' \\\\ ')} \\end{aligned}`;
  }

  private convertAccent(element: OmmlElement): string {
    let accentChar = '^';
    let base = '';

    for (const child of element.children) {
      if (child.type === 'accentProps' || child.tagName === 'm:accPr') {
        for (const prop of child.children) {
          if (prop.tagName === 'm:chr' || prop.type === 'character') {
            accentChar = prop.attributes['m:val'] || prop.attributes['val'] || '^';
          }
        }
      } else if (child.type === 'element' || child.tagName === 'm:e') {
        base = this.convertChildren(child);
      }
    }

    const accentCmd = this.accentMap.get(accentChar) || '\\hat';
    return `${accentCmd}{${base}}`;
  }

  private convertBar(element: OmmlElement): string {
    let base = '';

    for (const child of element.children) {
      if (child.type === 'element' || child.tagName === 'm:e') {
        base = this.convertChildren(child);
      }
    }

    return `\\overline{${base}}`;
  }

  private convertBox(element: OmmlElement): string {
    return this.convertChildren(element);
  }

  private convertGroupChar(element: OmmlElement): string {
    let base = '';
    let chr = '';

    for (const child of element.children) {
      if (child.tagName === 'm:groupChrPr') {
        for (const prop of child.children) {
          if (prop.tagName === 'm:chr') {
            chr = prop.attributes['m:val'] || prop.attributes['val'] || '';
          }
        }
      } else if (child.type === 'element' || child.tagName === 'm:e') {
        base = this.convertChildren(child);
      }
    }

    if (chr === '⏟' || chr === '︸') {
      return `\\underbrace{${base}}`;
    } else if (chr === '⏞' || chr === '︷') {
      return `\\overbrace{${base}}`;
    }

    return base;
  }

  private convertBorderBox(element: OmmlElement): string {
    let content = '';

    for (const child of element.children) {
      if (child.type === 'element' || child.tagName === 'm:e') {
        content = this.convertChildren(child);
      }
    }

    return `\\boxed{${content}}`;
  }

  private convertPreSuperSubscript(element: OmmlElement): string {
    let base = '';
    let sub = '';
    let sup = '';

    for (const child of element.children) {
      if (child.type === 'element' || child.tagName === 'm:e') {
        base = this.convertChildren(child);
      } else if (child.type === 'subscript' || child.tagName === 'm:sub') {
        sub = this.convertChildren(child);
      } else if (child.type === 'superscript' || child.tagName === 'm:sup') {
        sup = this.convertChildren(child);
      }
    }

    return `{}_{${sub}}^{${sup}}${base}`;
  }

  private convertChildren(element: OmmlElement): string {
    if (element.text) {
      return this.convertTextContent(element.text);
    }

    return element.children
      .map((child) => this.elementToLatex(child))
      .join('');
  }

  private convertTextContent(text: string): string {
    let result = '';
    for (const char of text) {
      if (this.unicodeToLatex.has(char)) {
        result += this.unicodeToLatex.get(char) + ' ';
      } else if (char === ' ') {
        result += '\\ ';
      } else {
        result += char;
      }
    }
    return result.trim();
  }

  private cleanupLatex(latex: string): string {
    return latex
      .replace(/\s+([+\-*/=])/g, ' $1')          // Пробел перед операциями
      .replace(/([+\-*/=])\s+/g, '$1 ')          // Пробел после операций
      .replace(/\s+/g, ' ')                       // Множественные пробелы в одиночные
      .replace(/\{\s+/g, '{')                     // Пробелы после открытой скобки
      .replace(/\s+\}/g, '}')                     // Пробелы перед закрытой скобкой
      .replace(/\s+_/g, '_')                      // Пробелы перед нижним индексом
      .replace(/\s+\^/g, '^')                     // Пробелы перед верхним индексом
      .replace(/_\s+\{/g, '_{')
      .replace(/\^\s+\{/g, '^{')
      .trim();
  }
}