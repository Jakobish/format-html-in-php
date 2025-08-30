'use strict';

export interface JscriptOptions {
  indentSize: number;
  indentChar: string;
  semicolons?: boolean;
  maxLineLength?: number;
  alignAssignments?: boolean;
  preserveAspComments?: boolean;
}

interface ParseContext {
  indentLevel: number;
  inFunction: boolean;
  inObjectLiteral: boolean;
  inArrayLiteral: boolean;
  inStatement: boolean;
  braceStack: string[];
}

export default function beautifyJscript(code: string, options: JscriptOptions): string {
  if (!code.trim()) return code;

  const indentSize = options.indentSize || 2;
  const indentChar = options.indentChar || ' ';
  const indent = indentChar.repeat(indentSize);
  const addSemicolons = options.semicolons !== false;
  const maxLineLength = options.maxLineLength || 100;
  const alignAssignments = options.alignAssignments || false;
  const preserveAspComments = options.preserveAspComments !== false;

  // Split into lines and preserve original line endings
  const lines = code.split(/\r?\n/);
  const formattedLines: string[] = [];
  const context: ParseContext = {
    indentLevel: 0,
    inFunction: false,
    inObjectLiteral: false,
    inArrayLiteral: false,
    inStatement: false,
    braceStack: []
  };

  // Collect assignments for potential alignment
  const assignments: Array<{ lineIndex: number; assignmentPos: number; line: string }> = [];

  for (let i = 0; i < lines.length; i++) {
    let line = lines[i].trim();

    // Skip empty lines
    if (!line) {
      formattedLines.push('');
      continue;
    }

    // Handle comments with ASP context
    if (line.startsWith('//') || (preserveAspComments && line.startsWith('/*'))) {
      const commentIndent = context.inObjectLiteral || context.inArrayLiteral ?
        indent.repeat(context.indentLevel + 1) : indent.repeat(context.indentLevel);
      formattedLines.push(commentIndent + line);
      continue;
    }

    // Handle multi-line comments
    if (line.startsWith('/*') && !line.endsWith('*/')) {
      formattedLines.push(indent.repeat(context.indentLevel) + line);
      continue;
    }
    if (line.endsWith('*/') && !line.startsWith('/*')) {
      formattedLines.push(indent.repeat(context.indentLevel) + line);
      continue;
    }

    // Apply current indentation
    let formattedLine = indent.repeat(context.indentLevel) + line;

    // Track braces for indentation
    const openBraces = (line.match(/\{/g) || []).length;
    const closeBraces = (line.match(/\}/g) || []).length;

    // Handle function declarations and expressions
    if (/\bfunction\s+\w+\s*\(/.test(line) || /\w+\s*\([^)]*\)\s*\{/.test(line)) {
      context.inFunction = true;
    }

    // Handle object literals
    if (line.includes('{') && !line.includes('function') && !line.includes('if') && !line.includes('for')) {
      context.inObjectLiteral = true;
    }

    // Handle array literals
    if (line.includes('[') && !line.includes('function')) {
      context.inArrayLiteral = true;
    }

    // Format ASP-specific constructs
    formattedLine = formatAspJscriptConstructs(formattedLine, indent, context.indentLevel);

    // Format object properties and method calls
    formattedLine = formatObjectAndArrays(formattedLine, indent, context);

    // Handle line breaking for long lines
    if (maxLineLength > 0 && formattedLine.length > maxLineLength) {
      formattedLine = breakLongJscriptLine(formattedLine, maxLineLength, indent, context.indentLevel);
    }

    // Collect assignments for alignment
    if (alignAssignments && /=/.test(line) && !line.includes('if ') && !line.includes('for ') && !line.includes('function')) {
      const assignmentPos = line.indexOf('=');
      if (assignmentPos > 0) {
        assignments.push({ lineIndex: formattedLines.length, assignmentPos, line: formattedLine });
      }
    }

    // Add semicolons if enabled
    if (addSemicolons && shouldAddSemicolon(line) && !formattedLine.endsWith(';') && !formattedLine.endsWith(',')) {
      formattedLine += ';';
    }

    formattedLines.push(formattedLine);

    // Update indent level based on braces
    context.indentLevel += openBraces;
    context.indentLevel -= closeBraces;

    // Reset context flags
    if (closeBraces > 0 && context.inFunction) {
      context.inFunction = false;
    }
    if (closeBraces > 0 && context.inObjectLiteral) {
      context.inObjectLiteral = false;
    }
    if (line.includes(']') && context.inArrayLiteral) {
      context.inArrayLiteral = false;
    }
  }

  // Apply assignment alignment if enabled
  if (alignAssignments && assignments.length > 1) {
    return applyJscriptAssignmentAlignment(formattedLines, assignments, indentSize);
  }

  return formattedLines.join('\n');
}

function formatAspJscriptConstructs(line: string, indent: string, indentLevel: number): string {
  // Format Server.CreateObject calls
  line = line.replace(/Server\.CreateObject\s*\(\s*(["'][^"']+["'])\s*\)/gi,
    `Server.CreateObject($1)`);

  // Format Response.Write calls
  line = line.replace(/Response\.Write\s*\(/gi, 'Response.Write(');

  // Format Request object access
  line = line.replace(/Request\.\w+/gi, match => match);

  // Format Session and Application objects
  line = line.replace(/(Session|Application)\.\w+/gi, match => match);

  return line;
}

function formatObjectAndArrays(line: string, indent: string, context: ParseContext): string {
  // Format object literals with proper indentation
  if (context.inObjectLiteral) {
    line = line.replace(/(\w+):\s*/g, (match, prop) => {
      return `\n${indent.repeat(context.indentLevel + 1)}${prop}: `;
    });
  }

  // Format array elements
  if (context.inArrayLiteral && line.includes(',')) {
    const elements = line.split(',');
    if (elements.length > 1) {
      line = elements.map(el => el.trim()).join(',\n' + indent.repeat(context.indentLevel + 1));
    }
  }

  // Format chained method calls
  line = line.replace(/(\.\w+\([^)]*\))/g, (match) => {
    return `\n${indent.repeat(context.indentLevel + 1)}${match.trim()}`;
  });

  return line;
}

function breakLongJscriptLine(line: string, maxLength: number, indent: string, indentLevel: number): string {
  if (line.length <= maxLength) return line;

  // Break at operators, commas, or method chains
  const breakPoints = [' && ', ' || ', ' + ', ' - ', ' * ', ' / ', ', ', ' ? ', ' : ', '.'];

  for (const breakPoint of breakPoints) {
    if (line.includes(breakPoint)) {
      const parts = line.split(breakPoint);
      if (parts.length > 1) {
        const result = parts[0] + breakPoint.trim() + '\n' +
          indent.repeat(indentLevel + 1) + parts.slice(1).join(breakPoint);
        if (result.length <= maxLength) {
          return result;
        }
      }
    }
  }

  return line;
}

function shouldAddSemicolon(line: string): boolean {
  // Don't add semicolons to comments, blocks, or already terminated lines
  if (line.startsWith('//') || line.startsWith('/*') || line.endsWith('{') ||
      line.endsWith('}') || line.endsWith(';') || line.endsWith(',')) {
    return false;
  }

  // Add semicolons to variable declarations, assignments, and function calls
  return /\b(var|let|const)\s+\w+\s*=/.test(line) ||
         /\w+\s*=/.test(line) ||
         /\w+\([^)]*\)$/.test(line);
}

function applyJscriptAssignmentAlignment(lines: string[], assignments: Array<{ lineIndex: number; assignmentPos: number; line: string }>, indentSize: number): string {
  if (assignments.length < 2) return lines.join('\n');

  // Find the maximum position of '=' for alignment
  const maxPos = Math.max(...assignments.map(a => a.assignmentPos));

  assignments.forEach(assignment => {
    const line = lines[assignment.lineIndex];
    const equalsIndex = line.indexOf('=');
    if (equalsIndex > 0) {
      const spacesNeeded = maxPos - equalsIndex;
      if (spacesNeeded > 0) {
        const beforeEquals = line.substring(0, equalsIndex);
        const afterEquals = line.substring(equalsIndex);
        lines[assignment.lineIndex] = beforeEquals + ' '.repeat(spacesNeeded) + afterEquals;
      }
    }
  });

  return lines.join('\n');
}