'use strict';

var beautifyHtml = require('js-beautify').html;
import beautifyVbscript from './beautifyVbscript';
import beautifyJscript from './beautifyJscript';

// Tokenize to safely extract ASP blocks even when '%>' appears inside strings
function extractAspBlocks(text: string, options: any) {
  const preservedBlocks: Array<{ placeholder: string; content: string; type: string }> = [];
  let placeholderIndex = 0;
  const out: string[] = [];
  const len = text.length;
  let i = 0;

  const pushPlaceholder = (content: string, type: string) => {
    if (!shouldPreserveBlock(type, options)) {
      out.push(content);
      return;
    }
    if ((type === 'server' || type === 'output') && (options.formatVbscriptInAspBlocks || options.formatJscriptInAspBlocks)) {
      try { content = formatAspBlock(content, options); } catch {}
    }
    const placeholder = `__ASP_BLOCK_${placeholderIndex++}__`;
    preservedBlocks.push({ placeholder, content, type });
    out.push(placeholder);
  };

  while (i < len) {
    const ch = text[i];
    // SSI include <!--#include ... --> (allow whitespace after <!--)
    if (ch === '<' && text.substr(i, 4) === '<!--' && /^<!--\s*#include/i.test(text.substr(i, 12))) {
      const end = text.indexOf('-->', i + 4);
      const endPos = end >= 0 ? end + 3 : len;
      pushPlaceholder(text.substring(i, endPos), 'include');
      i = endPos;
      continue;
    }
    // ASP comment <%-- ... --%>
    if (ch === '<' && text.substr(i, 5) === '<%--') {
      const end = text.indexOf('--%>', i + 4);
      const endPos = end >= 0 ? end + 4 : len;
      pushPlaceholder(text.substring(i, endPos), 'comment');
      i = endPos;
      continue;
    }
    // ASP server-side blocks
    if (ch === '<' && text.substr(i, 2) === '<%') {
      let type: 'directive' | 'output' | 'declaration' | 'server' = 'server';
      if (text.substr(i, 3) === '<%@') type = 'directive';
      else if (text.substr(i, 3) === '<%=') type = 'output';
      else if (text.substr(i, 3) === '<%!') type = 'declaration';

      // scan until %> while respecting quoted strings
      let j = i + 2;
      let inString = false;
      let quote = '';
      while (j < len) {
        const c = text[j];
        if (!inString && (c === '"' || c === "'")) { inString = true; quote = c; j++; continue; }
        if (inString && c === quote) { inString = false; quote = ''; j++; continue; }
        if (!inString && c === '%' && text[j + 1] === '>') { j += 2; break; }
        j++;
      }
      const endPos = j;
      pushPlaceholder(text.substring(i, endPos), type);
      i = endPos;
      continue;
    }
    // default: copy
    out.push(ch);
    i++;
  }

  return { processedText: out.join(''), preservedBlocks };
}

function shouldPreserveBlock(type, options) {
  switch (type) {
    case 'directive':
    case 'server':
    case 'output':
    case 'declaration':
      return true; // Always preserve core ASP blocks
    case 'include':
      return options.preserveIncludeDirectives !== false;
    case 'comment':
      return options.preserveAspComments !== false;
    default:
      return false;
  }
}

function detectAspLanguage(code: string): 'vbscript' | 'jscript' | 'unknown' {
  if (!code.trim()) return 'unknown';

  const lowerCode = code.toLowerCase();

  // Check for explicit language directives first
  if (lowerCode.includes('@language="vbscript"') || lowerCode.includes("@language='vbscript'")) {
    return 'vbscript';
  }
  if (lowerCode.includes('@language="jscript"') || lowerCode.includes("@language='jscript'") ||
      lowerCode.includes('@language="javascript"') || lowerCode.includes("@language='javascript'")) {
    return 'jscript';
  }

  // Enhanced VBScript indicators with weights
  const vbscriptIndicators = [
    { pattern: /\bdim\s+\w/i, weight: 3 },
    { pattern: /\bset\s+\w+\s*=/i, weight: 3 },
    { pattern: /\bfunction\s+\w+\s*\(/i, weight: 3 },
    { pattern: /\bsub\s+\w+\s*\(/i, weight: 3 },
    { pattern: /\bif\s+.*\bthen\b/i, weight: 3 },
    { pattern: /\bselect\s+case\b/i, weight: 3 },
    { pattern: /\bfor\s+.*\bto\b/i, weight: 3 },
    { pattern: /\bwhile\b/i, weight: 2 },
    { pattern: /\bwend\b/i, weight: 2 },
    { pattern: /\bnext\b/i, weight: 2 },
    { pattern: /\bcreateobject\b/i, weight: 3 },
    { pattern: /\bserver\./i, weight: 3 },
    { pattern: /\bresponse\./i, weight: 3 },
    { pattern: /\brequest\./i, weight: 3 },
    { pattern: /\bsession\./i, weight: 3 },
    { pattern: /\bapplication\./i, weight: 3 },
    { pattern: /\bado/i, weight: 2 },
    { pattern: /\bfilesystemobject\b/i, weight: 2 },
    { pattern: /\bend\s+(if|function|sub|select|class|with)\b/i, weight: 2 },
    { pattern: /\boption\s+explicit\b/i, weight: 2 },
    { pattern: /\bon\s+error\b/i, weight: 2 },
    { pattern: /\berror\b/i, weight: 1 },
    { pattern: /\bcall\s+\w/i, weight: 2 },
    { pattern: /\bexit\s+(function|sub|do|for)\b/i, weight: 2 }
  ];

  // Enhanced JScript indicators with weights
  const jscriptIndicators = [
    { pattern: /\bvar\s+\w+\s*=/i, weight: 3 },
    { pattern: /\blet\s+\w+\s*=/i, weight: 3 },
    { pattern: /\bconst\s+\w+\s*=/i, weight: 3 },
    { pattern: /\bfunction\s+\w+\s*\(/i, weight: 3 },
    { pattern: /\w+\s*\([^)]*\)\s*\{/i, weight: 3 },
    { pattern: /\bif\s*\(/i, weight: 3 },
    { pattern: /\bfor\s*\(/i, weight: 3 },
    { pattern: /\bwhile\s*\(/i, weight: 3 },
    { pattern: /\btry\s*\{/i, weight: 3 },
    { pattern: /\bcatch\s*\(/i, weight: 3 },
    { pattern: /\w+\.\w+\s*\(/i, weight: 2 },
    { pattern: /\bnew\s+\w+/i, weight: 3 },
    { pattern: /\bthis\./i, weight: 2 },
    { pattern: /\bprototype\b/i, weight: 2 },
    { pattern: /\btypeof\b/i, weight: 2 },
    { pattern: /\binstanceof\b/i, weight: 2 },
    { pattern: /\belse\s*\{/i, weight: 2 },
    { pattern: /\}\s*else\b/i, weight: 2 },
    { pattern: /\breturn\b/i, weight: 2 },
    { pattern: /\bbreak\b/i, weight: 2 },
    { pattern: /\bcontinue\b/i, weight: 2 },
    { pattern: /\w+\+\+|\+\+\w+/i, weight: 2 },
    { pattern: /\w+--|--\w+/i, weight: 2 },
    { pattern: /==|!=|===|!==/i, weight: 2 },
    { pattern: /&&|\|\|/i, weight: 2 }
  ];

  let vbscriptScore = 0;
  let jscriptScore = 0;

  // Count VBScript indicators with weights
  vbscriptIndicators.forEach(indicator => {
    if (indicator.pattern.test(lowerCode)) {
      vbscriptScore += indicator.weight;
    }
  });

  // Count JScript indicators with weights
  jscriptIndicators.forEach(indicator => {
    if (indicator.pattern.test(lowerCode)) {
      jscriptScore += indicator.weight;
    }
  });

  // Additional heuristics
  // Check for semicolon usage (JScript)
  const semicolonCount = (code.match(/;/g) || []).length;
  const linesCount = code.split('\n').length;
  if (semicolonCount > linesCount * 0.5) {
    jscriptScore += 2;
  }

  // Check for comment styles
  if (code.includes('//')) {
    jscriptScore += 2;
  }
  if (code.includes("'")) {
    vbscriptScore += 2;
  }

  // Check for ASP-specific patterns
  if (lowerCode.includes('response.write') || lowerCode.includes('request.')) {
    // Could be either, but slightly favor VBScript for Classic ASP
    vbscriptScore += 1;
  }

  // Determine language based on scores with a threshold
  const scoreDifference = Math.abs(vbscriptScore - jscriptScore);
  const minScoreThreshold = 3;

  if (vbscriptScore > jscriptScore && vbscriptScore >= minScoreThreshold) {
    return 'vbscript';
  } else if (jscriptScore > vbscriptScore && jscriptScore >= minScoreThreshold) {
    return 'jscript';
  } else if (scoreDifference < 2 && (vbscriptScore > 0 || jscriptScore > 0)) {
    // Close scores, default to VBScript for Classic ASP compatibility
    return 'vbscript';
  } else if (vbscriptScore > 0 || jscriptScore > 0) {
    // Some indicators found but not conclusive
    return vbscriptScore >= jscriptScore ? 'vbscript' : 'jscript';
  }

  return 'unknown';
}

function formatAspBlock(content: string, options: any): string {
  // Check if formatting is enabled for ASP blocks
  if (!options.formatVbscriptInAspBlocks && !options.formatJscriptInAspBlocks) return content;

  // Handle different ASP block types
  let code = '';
  let prefix = '<%';
  let suffix = '%>';

  // Extract code from ASP delimiters based on block type
  if (content.startsWith('<%=')) {
    const match = content.match(/^<%=(.*)%>$/);
    if (match) {
      code = match[1];
      prefix = '<%=';
    }
  } else if (content.startsWith('<%')) {
    const match = content.match(/^<%(.*)%>$/);
    if (match) {
      code = match[1];
    }
  } else {
    // Not a standard ASP block, return as-is
    return content;
  }

  if (!code.trim()) return content;

  // Detect language if auto-detection is enabled
  let language: 'vbscript' | 'jscript' | 'unknown' = 'unknown';
  if (options.detectAspLanguage !== false) {
    language = detectAspLanguage(code);
  } else {
    // Fallback to manual settings
    if (options.formatVbscriptInAspBlocks) {
      language = 'vbscript';
    } else if (options.formatJscriptInAspBlocks) {
      language = 'jscript';
    }
  }

  let formattedCode = code;

  // Apply appropriate formatter based on detected language
  if (language === 'vbscript' && options.formatVbscriptInAspBlocks) {
    try {
      formattedCode = beautifyVbscript(code, {
        indentSize: options.vbscriptIndentSize || 4,
        indentChar: options.indentChar || ' ',
        alignAssignments: options.vbscriptAlignAssignments || false,
        preserveAspComments: options.preserveAspComments !== false,
        maxLineLength: options.maxLineLength || 120
      });
    } catch (error) {
      // If formatting fails, keep original code
      console.warn('VBScript formatting failed:', error.message);
      formattedCode = code;
    }
  } else if (language === 'jscript' && options.formatJscriptInAspBlocks) {
    try {
      formattedCode = beautifyJscript(code, {
        indentSize: options.jscriptIndentSize || 2,
        indentChar: options.indentChar || ' ',
        semicolons: options.jscriptSemicolons !== false,
        alignAssignments: options.jscriptAlignAssignments || false,
        preserveAspComments: options.preserveAspComments !== false,
        maxLineLength: options.maxLineLength || 100
      });
    } catch (error) {
      // If formatting fails, keep original code
      console.warn('JScript formatting failed:', error.message);
      formattedCode = code;
    }
  }

  return `${prefix}${formattedCode}${suffix}`;
}

function reinsertAspBlocks(text: string, preservedBlocks: Array<{ placeholder: string, content: string, type: string }>, options: any): string {
  if (!preservedBlocks || preservedBlocks.length === 0) return text;

  // Create quick lookup map
  const map = new Map<string, { content: string, type: string }>();
  preservedBlocks.forEach(b => map.set(b.placeholder, { content: b.content, type: b.type }));

  // Replace placeholders with aligned content
  return text.replace(/__ASP_BLOCK_\d+__/g, (ph, ...args) => {
    const matchIndex = args[args.length - 2] as number; // offset
    const full = args[args.length - 1] as string; // full string
    const entry = map.get(ph);
    if (!entry) return ph;

    const content = entry.content;

    if (options && options.alignServerBlocks === false) {
      return content; // no alignment
    }

    // Determine indentation at placeholder position
    const lineStart = full.lastIndexOf('\n', matchIndex - 1) + 1;
    const before = full.slice(lineStart, matchIndex);
    const isAtLineStart = /^\s*$/.test(before);
    const indent = before; // existing indentation spaces/tabs before placeholder

    if (!isAtLineStart) {
      // Inline placeholder: return content as-is
      return content;
    }

    // Align multi-line ASP content to surrounding HTML indent
    const lines = content.split('\n');
    if (lines.length === 1) {
      return content; // single-line; outer indent already present
    }

    // Strip common leading whitespace from subsequent lines, then re-indent
    const subsequent = lines.slice(1);
    const nonEmpty = subsequent.filter(l => l.trim().length > 0);
    const commonPrefixLen = nonEmpty.length
      ? Math.min(...nonEmpty.map(l => (l.match(/^\s*/)?.[0].length || 0)))
      : 0;

    const adjusted = [lines[0]].concat(
      subsequent.map(l => {
        const stripped = l.replace(new RegExp('^\\s{0,' + commonPrefixLen + '}'), '');
        return indent + stripped;
      })
    );

    return adjusted.join('\n');
  });
}

export default function (originalText, htmlOptions) {
  try {
    // Preserve BOM and EOLs
    const hasBOM = originalText.charCodeAt(0) === 0xFEFF;
    const eol = /\r\n/.test(originalText) ? '\r\n' : '\n';
    const textNoBOM = hasBOM ? originalText.slice(1) : originalText;

    // First, extract and preserve all ASP blocks
    const { processedText, preservedBlocks } = extractAspBlocks(textNoBOM, htmlOptions);

    // Check if we have any ASP blocks to format
    if (preservedBlocks.length === 0) {
      // No ASP blocks found, just format as regular HTML
      let result = beautifyHtml(textNoBOM, htmlOptions);
      if (htmlOptions && htmlOptions.trimTrailingWhitespace !== false) {
        result = result.replace(/[ \t]+$/gm, '');
      }
      if (eol === '\r\n') {
        result = result.replace(/\r?\n/g, '\r\n');
      }
      return hasBOM ? '\ufeff' + result : result;
    }

    // Format the HTML without ASP blocks
    let formattedHtml = beautifyHtml(processedText, htmlOptions);

    // Run post-processing while placeholders are still present
    formattedHtml = postProcessAspContent(formattedHtml, htmlOptions);

    // Re-insert the preserved ASP blocks
    let finalResult = reinsertAspBlocks(formattedHtml, preservedBlocks, htmlOptions);

    // Cleanup: strip trailing whitespace per line
    if (htmlOptions && htmlOptions.trimTrailingWhitespace !== false) {
      finalResult = finalResult.replace(/[ \t]+$/gm, '');
    }

    if (htmlOptions && htmlOptions.trimTrailingWhitespace !== false) {
      finalResult = finalResult.replace(/[ \t]+$/gm, '');
    }
    if (eol === '\r\n') {
      finalResult = finalResult.replace(/\r?\n/g, '\r\n');
    }
    return hasBOM ? '\ufeff' + finalResult : finalResult;

  } catch (error) {
    // If extraction or formatting fails, fall back to original behavior
    console.warn('ASP block processing failed, falling back to standard formatting:', error.message);
    let result = beautifyHtml(originalText, htmlOptions);
    if (htmlOptions && htmlOptions.trimTrailingWhitespace !== false) {
      result = result.replace(/[ \t]+$/gm, '');
    }
    return result;
  }
}

function postProcessAspContent(content: string, options: any): string {
  // Post-processing to handle edge cases in complex ASP files

  // Improve readability when ASP placeholders appear between tags
  // Keep this conservative to avoid altering attribute values.
  content = content.replace(/(>)\s*(__ASP_BLOCK_\d+__)\s*(<)/g, '$1\n$2\n$3');

  return content;
}
