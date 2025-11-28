import {
  OmmlParser,
  OmmlElement,
  OmmlElementType,
} from "../parsers/ommlParser";

export class FormulaConverter {
  private parser: OmmlParser;

  // Unicode to LaTeX mapping
  private unicodeToLatex: Map<string, string> = new Map([
    // Greek letters
    ["α", "\\alpha"],
    ["β", "\\beta"],
    ["γ", "\\gamma"],
    ["δ", "\\delta"],
    ["ε", "\\epsilon"],
    ["ζ", "\\zeta"],
    ["η", "\\eta"],
    ["θ", "\\theta"],
    ["ι", "\\iota"],
    ["κ", "\\kappa"],
    ["λ", "\\lambda"],
    ["μ", "\\mu"],
    ["ν", "\\nu"],
    ["ξ", "\\xi"],
    ["π", "\\pi"],
    ["ρ", "\\rho"],
    ["σ", "\\sigma"],
    ["τ", "\\tau"],
    ["υ", "\\upsilon"],
    ["φ", "\\phi"],
    ["χ", "\\chi"],
    ["ψ", "\\psi"],
    ["ω", "\\omega"],
    ["Γ", "\\Gamma"],
    ["Δ", "\\Delta"],
    ["Θ", "\\Theta"],
    ["Λ", "\\Lambda"],
    ["Ξ", "\\Xi"],
    ["Π", "\\Pi"],
    ["Σ", "\\Sigma"],
    ["Υ", "\\Upsilon"],
    ["Φ", "\\Phi"],
    ["Ψ", "\\Psi"],
    ["Ω", "\\Omega"],
    // Operators
    ["×", "\\times"],
    ["÷", "\\div"],
    ["±", "\\pm"],
    ["∓", "\\mp"],
    ["·", "\\cdot"],
    ["°", "^\\circ"],
    ["∞", "\\infty"],
    ["≈", "\\approx"],
    ["≠", "\\neq"],
    ["≤", "\\leq"],
    ["≥", "\\geq"],
    ["≪", "\\ll"],
    ["≫", "\\gg"],
    ["∝", "\\propto"],
    ["≡", "\\equiv"],
    ["∼", "\\sim"],
    ["≃", "\\simeq"],
    ["≅", "\\cong"],
    // Set theory
    ["∈", "\\in"],
    ["∉", "\\notin"],
    ["⊂", "\\subset"],
    ["⊃", "\\supset"],
    ["⊆", "\\subseteq"],
    ["⊇", "\\supseteq"],
    ["∪", "\\cup"],
    ["∩", "\\cap"],
    ["∅", "\\emptyset"],
    ["∀", "\\forall"],
    ["∃", "\\exists"],
    ["∄", "\\nexists"],
    // Arrows
    ["→", "\\rightarrow"],
    ["←", "\\leftarrow"],
    ["↔", "\\leftrightarrow"],
    ["⇒", "\\Rightarrow"],
    ["⇐", "\\Leftarrow"],
    ["⇔", "\\Leftrightarrow"],
    ["↑", "\\uparrow"],
    ["↓", "\\downarrow"],
    ["↦", "\\mapsto"],
    // Calculus
    ["∂", "\\partial"],
    ["∇", "\\nabla"],
    ["∫", "\\int"],
    ["∬", "\\iint"],
    ["∭", "\\iiint"],
    ["∮", "\\oint"],
    ["∑", "\\sum"],
    ["∏", "\\prod"],
    ["√", "\\sqrt"],
    // Logic
    ["∧", "\\land"],
    ["∨", "\\lor"],
    ["¬", "\\neg"],
    ["⊕", "\\oplus"],
    ["⊗", "\\otimes"],
    ["⊥", "\\perp"],
    ["∥", "\\parallel"],
    // Misc
    ["ℕ", "\\mathbb{N}"],
    ["ℤ", "\\mathbb{Z}"],
    ["ℚ", "\\mathbb{Q}"],
    ["ℝ", "\\mathbb{R}"],
    ["ℂ", "\\mathbb{C}"],
    ["ℏ", "\\hbar"],
    ["ℓ", "\\ell"],
    ["′", "'"],
    ["″", "''"],
  ]);

  private functionNames: Map<string, string> = new Map([
    ["sin", "\\sin"],
    ["cos", "\\cos"],
    ["tan", "\\tan"],
    ["cot", "\\cot"],
    ["sec", "\\sec"],
    ["csc", "\\csc"],
    ["arcsin", "\\arcsin"],
    ["arccos", "\\arccos"],
    ["arctan", "\\arctan"],
    ["sinh", "\\sinh"],
    ["cosh", "\\cosh"],
    ["tanh", "\\tanh"],
    ["log", "\\log"],
    ["ln", "\\ln"],
    ["exp", "\\exp"],
    ["lim", "\\lim"],
    ["max", "\\max"],
    ["min", "\\min"],
    ["sup", "\\sup"],
    ["inf", "\\inf"],
    ["det", "\\det"],
    ["dim", "\\dim"],
    ["ker", "\\ker"],
    ["gcd", "\\gcd"],
    ["mod", "\\mod"],
    ["arg", "\\arg"],
    ["deg", "\\deg"],
  ]);

  private accentMap: Map<string, string> = new Map([
    ["̂", "\\hat"],
    ["̃", "\\tilde"],
    ["̄", "\\bar"],
    ["⃗", "\\vec"],
    ["̇", "\\dot"],
    ["̈", "\\ddot"],
    ["̆", "\\breve"],
    ["̌", "\\check"],
    ["ˆ", "\\hat"],
    ["˜", "\\tilde"],
    ["¯", "\\bar"],
    ["→", "\\vec"],
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
      console.error("Formula conversion error:", error);
      return "[Formula conversion error]";
    }
  }

  private elementToLatex(element: OmmlElement): string {
    switch (element.type) {
      case "math":
      case "mathPara":
        return this.convertMath(element);
      case "run":
        return this.convertRun(element);
      case "text":
        return this.convertText(element);
      case "fraction":
        return this.convertFraction(element);
      case "radical":
        return this.convertRadical(element);
      case "superscriptContainer":
        return this.convertSuperscript(element);
      case "subscriptContainer":
        return this.convertSubscript(element);
      case "subSupContainer":
        return this.convertSubSup(element);
      case "nary":
        return this.convertNary(element);
      case "limitLow":
        return this.convertLimitLow(element);
      case "limitUpper":
        return this.convertLimitUpper(element);
      case "matrix":
        return this.convertMatrix(element);
      case "delimiter":
        return this.convertDelimiter(element);
      case "function":
        return this.convertFunction(element);
      case "equationArray":
        return this.convertEquationArray(element);
      case "accent":
        return this.convertAccent(element);
      case "bar":
        return this.convertBar(element);
      case "box":
        return this.convertBox(element);
      case "groupChar":
        return this.convertGroupChar(element);
      case "borderBox":
        return this.convertBorderBox(element);
      case "preSuperSubscript":
        return this.convertPreSuperSubscript(element);
      case "element":
      case "numerator":
      case "denominator":
      case "degree":
      case "superscript":
      case "subscript":
      case "limit":
      case "matrixRow":
      case "functionName":
        return this.convertChildren(element);
      case "naryProps":
      case "delimiterProps":
      case "accentProps":
      case "controlProps":
      case "runProps":
        return ""; // Skip property elements
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

  /**
   * Проверяет, является ли символ буквой (латинской или другой)
   */
  private isLetter(char: string): boolean {
    return /^[a-zA-Z]$/.test(char);
  }

  /**
   * Проверяет, является ли символ цифрой
   */
  private isDigit(char: string): boolean {
    return /^[0-9]$/.test(char);
  }

  /**
   * Проверяет, является ли символ алфавитно-цифровым
   */
  private isAlphanumeric(char: string): boolean {
    return this.isLetter(char) || this.isDigit(char);
  }

  /**
   * Проверяет, заканчивается ли строка на LaTeX-команду (например, \Delta)
   * которая требует пробела перед следующей буквой
   */
  private endsWithLatexCommand(str: string): boolean {
    // Ищем LaTeX-команду в конце строки: \commandname
    const match = str.match(/\\[a-zA-Z]+$/);
    return match !== null;
  }

  /**
   * Проверяет, нужен ли пробел между текущим результатом и следующим символом
   */
  private needsSpaceBefore(result: string, nextChar: string): boolean {
    if (!result) return false;

    // Если следующий символ - буква или цифра, и результат заканчивается на LaTeX-команду
    if (this.isAlphanumeric(nextChar) && this.endsWithLatexCommand(result)) {
      return true;
    }

    return false;
  }

  private convertText(element: OmmlElement): string {
    let text = element.text || "";

    // Also check children for text
    for (const child of element.children) {
      if (child.text) {
        text += child.text;
      } else if (child.type === "text") {
        text += this.convertText(child);
      }
    }

    return this.convertTextContent(text);
  }

  private convertFraction(element: OmmlElement): string {
    let numerator = "";
    let denominator = "";

    for (const child of element.children) {
      if (child.type === "numerator" || child.tagName === "m:num") {
        numerator = this.convertChildren(child);
      } else if (child.type === "denominator" || child.tagName === "m:den") {
        denominator = this.convertChildren(child);
      }
    }

    if (
      numerator.includes("+") ||
      numerator.includes("-") ||
      numerator.includes("*")
    ) {
      numerator = `{${numerator}}`;
    }

    if (
      denominator.includes("+") ||
      denominator.includes("-") ||
      denominator.includes("*")
    ) {
      denominator = `{${denominator}}`;
    }

    return `\\frac{${numerator}}{${denominator}}`;
  }

  private convertRadical(element: OmmlElement): string {
    let degree = "";
    let base = "";

    for (const child of element.children) {
      if (child.type === "degree" || child.tagName === "m:deg") {
        degree = this.convertChildren(child);
      } else if (child.type === "element" || child.tagName === "m:e") {
        base = this.convertChildren(child);
      }
    }

    if (degree && degree.trim() !== "" && degree.trim() !== "2") {
      return `\\sqrt[${degree}]{${base}}`;
    }
    return `\\sqrt{${base}}`;
  }

  private convertSuperscript(element: OmmlElement): string {
    let base = "";
    let sup = "";

    for (const child of element.children) {
      if (child.type === "element" || child.tagName === "m:e") {
        base = this.convertChildren(child);
      } else if (child.type === "superscript" || child.tagName === "m:sup") {
        sup = this.convertChildren(child);
      }
    }

    return `${base}^{${sup}}`;
  }

  private convertSubscript(element: OmmlElement): string {
    let base = "";
    let sub = "";

    for (const child of element.children) {
      if (child.type === "element" || child.tagName === "m:e") {
        base = this.convertChildren(child);
      } else if (child.type === "subscript" || child.tagName === "m:sub") {
        sub = this.convertChildren(child);
      }
    }

    return `${base}_{${sub}}`;
  }

  private convertSubSup(element: OmmlElement): string {
    let base = "";
    let sub = "";
    let sup = "";

    for (const child of element.children) {
      if (child.type === "element" || child.tagName === "m:e") {
        base = this.convertChildren(child);
      } else if (child.type === "subscript" || child.tagName === "m:sub") {
        sub = this.convertChildren(child);
      } else if (child.type === "superscript" || child.tagName === "m:sup") {
        sup = this.convertChildren(child);
      }
    }

    return `${base}_{${sub}}^{${sup}}`;
  }

  private convertNary(element: OmmlElement): string {
    let operator = "\\int";
    let sub = "";
    let sup = "";
    let base = "";

    for (const child of element.children) {
      if (child.type === "naryProps" || child.tagName === "m:naryPr") {
        for (const prop of child.children) {
          if (prop.tagName === "m:chr" || prop.type === "character") {
            const chr = prop.attributes["m:val"] || prop.attributes["val"];
            if (chr) {
              operator = this.convertNaryChar(chr);
            }
          }
        }
      } else if (child.type === "subscript" || child.tagName === "m:sub") {
        sub = this.convertChildren(child);
      } else if (child.type === "superscript" || child.tagName === "m:sup") {
        sup = this.convertChildren(child);
      } else if (child.type === "element" || child.tagName === "m:e") {
        base = this.convertChildren(child);
      }
    }

    let result = operator;
    if (sub) result += `_{${sub}}`;
    if (sup) result += `^{${sup}}`;
    if (base) result += ` ${base}`;

    return result;
  }

  private convertNaryChar(char: string): string {
    const naryMap: Record<string, string> = {
      "∫": "\\int",
      "∬": "\\iint",
      "∭": "\\iiint",
      "∮": "\\oint",
      "∑": "\\sum",
      "∏": "\\prod",
      "⋃": "\\bigcup",
      "⋂": "\\bigcap",
      "⋁": "\\bigvee",
      "⋀": "\\bigwedge",
    };

    return naryMap[char] || "\\int";
  }

  private convertLimitLow(element: OmmlElement): string {
    let base = "";
    let lim = "";

    for (const child of element.children) {
      if (child.type === "element" || child.tagName === "m:e") {
        base = this.convertChildren(child);
      } else if (child.type === "limit" || child.tagName === "m:lim") {
        lim = this.convertChildren(child);
      }
    }

    return `${base}_{${lim}}`;
  }

  private convertLimitUpper(element: OmmlElement): string {
    let base = "";
    let lim = "";

    for (const child of element.children) {
      if (child.type === "element" || child.tagName === "m:e") {
        base = this.convertChildren(child);
      } else if (child.type === "limit" || child.tagName === "m:lim") {
        lim = this.convertChildren(child);
      }
    }

    return `${base}^{${lim}}`;
  }

  private convertMatrix(element: OmmlElement): string {
    const rows: string[] = [];

    for (const child of element.children) {
      if (child.type === "matrixRow" || child.tagName === "m:mr") {
        const cells: string[] = [];
        for (const cell of child.children) {
          if (cell.type === "element" || cell.tagName === "m:e") {
            cells.push(this.convertChildren(cell));
          }
        }
        rows.push(cells.join(" & "));
      }
    }

    return `\\begin{pmatrix} ${rows.join(" \\\\ ")} \\end{pmatrix}`;
  }

  private convertDelimiter(element: OmmlElement): string {
    let beginChar = "(";
    let endChar = ")";
    const contents: string[] = [];

    for (const child of element.children) {
      if (child.type === "delimiterProps" || child.tagName === "m:dPr") {
        for (const prop of child.children) {
          if (prop.tagName === "m:begChr" || prop.type === "beginChar") {
            beginChar =
              prop.attributes["m:val"] || prop.attributes["val"] || "(";
          }
          if (prop.tagName === "m:endChr" || prop.type === "endChar") {
            endChar = prop.attributes["m:val"] || prop.attributes["val"] || ")";
          }
        }
      } else if (child.type === "element" || child.tagName === "m:e") {
        contents.push(this.convertChildren(child));
      }
    }

    const leftDelim = this.convertDelimiterChar(beginChar);
    const rightDelim = this.convertDelimiterChar(endChar);

    return `\\left${leftDelim} ${contents.join(", ")} \\right${rightDelim}`;
  }

  private convertDelimiterChar(char: string): string {
    const delimMap: Record<string, string> = {
      "(": "(",
      ")": ")",
      "[": "[",
      "]": "]",
      "{": "\\{",
      "}": "\\}",
      "|": "|",
      "‖": "\\|",
      "⌈": "\\lceil",
      "⌉": "\\rceil",
      "⌊": "\\lfloor",
      "⌋": "\\rfloor",
      "⟨": "\\langle",
      "⟩": "\\rangle",
      "": ".",
    };

    return delimMap[char] || char;
  }

  private convertFunction(element: OmmlElement): string {
    let funcName = "";
    let arg = "";

    for (const child of element.children) {
      if (child.type === "functionName" || child.tagName === "m:fName") {
        funcName = this.convertChildren(child).trim();
      } else if (child.type === "element" || child.tagName === "m:e") {
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
      if (child.type === "element" || child.tagName === "m:e") {
        equations.push(this.convertChildren(child));
      }
    }

    if (equations.length === 1) {
      return equations[0];
    }

    return `\\begin{aligned} ${equations.join(" \\\\ ")} \\end{aligned}`;
  }

  private convertAccent(element: OmmlElement): string {
    let accentChar = "^";
    let base = "";

    for (const child of element.children) {
      if (child.type === "accentProps" || child.tagName === "m:accPr") {
        for (const prop of child.children) {
          if (prop.tagName === "m:chr" || prop.type === "character") {
            accentChar =
              prop.attributes["m:val"] || prop.attributes["val"] || "^";
          }
        }
      } else if (child.type === "element" || child.tagName === "m:e") {
        base = this.convertChildren(child);
      }
    }

    const accentCmd = this.accentMap.get(accentChar) || "\\hat";
    return `${accentCmd}{${base}}`;
  }

  private convertBar(element: OmmlElement): string {
    let base = "";

    for (const child of element.children) {
      if (child.type === "element" || child.tagName === "m:e") {
        base = this.convertChildren(child);
      }
    }

    return `\\overline{${base}}`;
  }

  private convertBox(element: OmmlElement): string {
    return this.convertChildren(element);
  }

  private convertGroupChar(element: OmmlElement): string {
    let base = "";
    let chr = "";

    for (const child of element.children) {
      if (child.tagName === "m:groupChrPr") {
        for (const prop of child.children) {
          if (prop.tagName === "m:chr") {
            chr = prop.attributes["m:val"] || prop.attributes["val"] || "";
          }
        }
      } else if (child.type === "element" || child.tagName === "m:e") {
        base = this.convertChildren(child);
      }
    }

    if (chr === "⏟" || chr === "︸") {
      return `\\underbrace{${base}}`;
    } else if (chr === "⏞" || chr === "︷") {
      return `\\overbrace{${base}}`;
    }

    return base;
  }

  private convertBorderBox(element: OmmlElement): string {
    let content = "";

    for (const child of element.children) {
      if (child.type === "element" || child.tagName === "m:e") {
        content = this.convertChildren(child);
      }
    }

    return `\\boxed{${content}}`;
  }

  private convertPreSuperSubscript(element: OmmlElement): string {
    let base = "";
    let sub = "";
    let sup = "";

    for (const child of element.children) {
      if (child.type === "element" || child.tagName === "m:e") {
        base = this.convertChildren(child);
      } else if (child.type === "subscript" || child.tagName === "m:sub") {
        sub = this.convertChildren(child);
      } else if (child.type === "superscript" || child.tagName === "m:sup") {
        sup = this.convertChildren(child);
      }
    }

    return `{}_{${sub}}^{${sup}}${base}`;
  }

  private convertChildren(element: OmmlElement): string {
    if (element.text) {
      return this.convertTextContent(element.text);
    }

    const parts: string[] = [];
    for (const child of element.children) {
      const converted = this.elementToLatex(child);
      if (converted) {
        parts.push(converted);
      }
    }

    // Объединяем части с учётом необходимости пробелов
    let result = "";
    for (let i = 0; i < parts.length; i++) {
      const part = parts[i];
      if (i === 0) {
        result = part;
      } else {
        // Проверяем, нужен ли пробел между предыдущей частью и текущей
        const firstCharOfPart = part.charAt(0);
        if (this.needsSpaceBefore(result, firstCharOfPart)) {
          result += " " + part;
        } else {
          result += part;
        }
      }
    }

    return result;
  }

  private convertTextContent(text: string): string {
    // Normalize non-breaking spaces to plain spaces from the input (Word often uses U+00A0)
    text = text.replace(/\u00A0/g, " ");
    let result = "";
    const chars = Array.from(text); // Корректная работа с Unicode

    for (let i = 0; i < chars.length; i++) {
      const char = chars[i];
      const nextChar = i + 1 < chars.length ? chars[i + 1] : null;

      if (this.unicodeToLatex.has(char)) {
        const latexCmd = this.unicodeToLatex.get(char)!;

        // Добавляем пробел перед командой, если результат заканчивается на букву/цифру
        if (result && this.isAlphanumeric(result.charAt(result.length - 1))) {
          result += " ";
        }

        result += latexCmd;

        // Добавляем пробел после команды, если следующий символ - буква или цифра
        // и команда не заканчивается на }
        if (
          nextChar &&
          this.isAlphanumeric(nextChar) &&
          !latexCmd.endsWith("}")
        ) {
          result += " ";
        }
      } else if (char === " ") {
        // Пробел добавляем только если он нужен
        if (result && !result.endsWith(" ") && !result.endsWith("{")) {
          result += " ";
        }
      } else if (["+", "-", "="].includes(char)) {
        // Операторы с пробелами вокруг, но избегаем двойных пробелов
        if (result && !result.endsWith(" ")) {
          result += " ";
        }
        result += char;
        // Пробел после добавим только если следующий символ не пробел
        if (nextChar && nextChar !== " ") {
          result += " ";
        }
      } else {
        // Обычный символ
        // Проверяем, нужен ли пробел перед ним (после LaTeX-команды)
        if (this.isAlphanumeric(char) && this.endsWithLatexCommand(result)) {
          result += " ";
        }
        result += char;
      }
    }

    return result;
  }

  private cleanupLatex(latex: string): string {
    return (
      latex
        // Normalize non-breaking spaces as HTML may render U+00A0 as &nbsp;
        .replace(/\u00A0/g, " ")
        // Нормализуем пробелы вокруг операторов, но сохраняем один пробел
        .replace(/\s*([+\-=])\s*/g, " $1 ")
        // Убираем множественные пробелы
        .replace(/\s{2,}/g, " ")
        // Убираем пробелы внутри фигурных скобок (в начале и конце)
        .replace(/\{\s+/g, "{")
        .replace(/\s+\}/g, "}")
        // Убираем пробел между закрывающей скобкой и _ или ^
        .replace(/\}\s+_/g, "}_")
        .replace(/\}\s+\^/g, "}^")
        // Убираем пробел перед _ и ^ только если перед ним не LaTeX-команда
        // Это делается через negative lookbehind
        .replace(/([^\\a-zA-Z])\s+_/g, "$1_")
        .replace(/([^\\a-zA-Z])\s+\^/g, "$1^")
        // Но при этом сохраняем пробел после LaTeX-команд перед индексами
        // НЕ убираем пробел в случаях типа \Delta _
        // Нормализуем _{  и ^{
        .replace(/_\s+\{/g, "_{")
        .replace(/\^\s+\{/g, "^{")
        .trim()
    );
  }
}
