// Simple test script for VBScript beautifier
const beautifyVbscript = function(code, options) {
  if (!code.trim()) return code;

  const indentSize = options.indentSize || 4;
  const indentChar = options.indentChar || ' ';
  const indent = indentChar.repeat(indentSize);
  const alignAssignments = options.alignAssignments || false;
  const preserveAspComments = options.preserveAspComments !== false;
  const maxLineLength = options.maxLineLength || 120;

  // Split into lines and preserve original line endings
  const lines = code.split(/\r?\n/);
  const formattedLines = [];
  const context = {
    indentLevel: 0,
    inMultiLineStatement: false,
    inSelectCase: false,
    inWithBlock: false,
    pendingIndentIncrease: false,
    currentBlockType: ''
  };

  // Keywords that increase indentation
  const indentKeywords = /^(if|for|while|do|function|sub|class|with|select|property|try)\b/i;

  // Keywords that decrease indentation
  const outdentKeywords = /^(else|end|next|loop|wend|catch|finally)\b/i;

  // Keywords that handle same line structures
  const sameLineKeywords = /^(then|else)\b/i;

  // Multi-line statement indicators
  const multiLineIndicators = /^(\s*(if|for|while|do|with|select)\b.*[^_]\s*$)/i;

  // Continuation character
  const continuationChar = /\s+_\s*$/;

  // Collect all assignments for alignment if enabled
  const assignments = [];

  for (let i = 0; i < lines.length; i++) {
    let line = lines[i].trim();

    // Skip empty lines
    if (!line) {
      formattedLines.push('');
      continue;
    }

    // Handle comments with better preservation
    if (line.startsWith("'") || (preserveAspComments && line.startsWith("<!--"))) {
      const commentIndent = context.inSelectCase && line.toLowerCase().includes('case') ?
        indent.repeat(context.indentLevel + 1) : indent.repeat(context.indentLevel);
      formattedLines.push(commentIndent + line);
      continue;
    }

    // Handle outdent keywords first
    if (outdentKeywords.test(line)) {
      context.indentLevel = Math.max(0, context.indentLevel - 1);
      context.inMultiLineStatement = false;

      // Handle specific block endings
      if (line.toLowerCase().startsWith('end select')) {
        context.inSelectCase = false;
      } else if (line.toLowerCase().startsWith('end with')) {
        context.inWithBlock = false;
      }
    }

    // Apply current indentation
    let formattedLine = indent.repeat(context.indentLevel) + line;

    // Handle Select Case blocks
    if (line.toLowerCase().startsWith('select case')) {
      context.inSelectCase = true;
      context.pendingIndentIncrease = true;
    } else if (context.inSelectCase && line.toLowerCase().startsWith('case')) {
      formattedLine = indent.repeat(context.indentLevel + 1) + line;
    } else if (context.inSelectCase && line.toLowerCase().startsWith('case else')) {
      formattedLine = indent.repeat(context.indentLevel + 1) + line;
    }

    // Handle With blocks
    if (line.toLowerCase().startsWith('with ')) {
      context.inWithBlock = true;
      context.pendingIndentIncrease = true;
    }

    // Handle array and object formatting
    formattedLine = formatArraysAndObjects(formattedLine, indent, context.indentLevel);

    // Handle ASP-specific constructs
    formattedLine = formatAspConstructs(formattedLine, indent, context.indentLevel);

    // Handle complex expressions and line breaking
    if (maxLineLength > 0 && formattedLine.length > maxLineLength) {
      formattedLine = breakLongLine(formattedLine, maxLineLength, indent, context.indentLevel);
    }

    // Collect assignments for alignment
    if (alignAssignments && /=/.test(line) && !line.toLowerCase().includes('if ') && !line.toLowerCase().includes('then')) {
      const assignmentPos = line.indexOf('=');
      if (assignmentPos > 0) {
        assignments.push({ lineIndex: formattedLines.length, assignmentPos, line: formattedLine });
      }
    }

    // Handle same line keywords (like If ... Then)
    if (sameLineKeywords.test(line)) {
      formattedLine = indent.repeat(context.indentLevel) + line;
    }

    // Handle multi-line structures
    if (multiLineIndicators.test(line) && !continuationChar.test(line)) {
      context.inMultiLineStatement = true;
      context.pendingIndentIncrease = true;
    }

    // Handle line continuation
    if (continuationChar.test(line)) {
      formattedLine += ' _';
      context.inMultiLineStatement = true;
    }

    // Increase indent for keywords
    if (indentKeywords.test(line) && !line.toLowerCase().includes('end')) {
      if (context.pendingIndentIncrease) {
        context.indentLevel++;
        context.pendingIndentIncrease = false;
      }
    }

    formattedLines.push(formattedLine);

    // Handle pending indent increases
    if (context.pendingIndentIncrease) {
      context.indentLevel++;
      context.pendingIndentIncrease = false;
    }
  }

  // Apply assignment alignment if enabled
  if (alignAssignments && assignments.length > 1) {
    return applyAssignmentAlignment(formattedLines, assignments, indentSize);
  }

  return formattedLines.join('\n');
};

function formatArraysAndObjects(line, indent, indentLevel) {
  // Format array literals
  line = line.replace(/Array\(([^)]+)\)/gi, (match, content) => {
    const items = content.split(',').map(item => item.trim());
    if (items.length <= 3) {
      return `Array(${items.join(', ')})`;
    } else {
      const indentedItems = items.map(item => `\n${indent.repeat(indentLevel + 1)}${item}`);
      return `Array(${indentedItems.join(',')}\n${indent.repeat(indentLevel)})`;
    }
  });

  // Format dictionary/object creation
  line = line.replace(/CreateObject\(["']Scripting\.Dictionary["']\)/gi,
    'CreateObject("Scripting.Dictionary")');

  return line;
}

function formatAspConstructs(line, indent, indentLevel) {
  // Format ADODB connection strings and queries
  if (line.toLowerCase().includes('adodb.') || line.toLowerCase().includes('filesystemobject')) {
    line = line.replace(/(\w+)\.(\w+)\s*=\s*/g, '$1.$2 = ');
  }

  // Format Server.CreateObject calls
  line = line.replace(/Server\.CreateObject\((["'][^"']+["'])\)/gi,
    `Server.CreateObject($1)`);

  // Format Response.Write calls
  line = line.replace(/Response\.Write\s*\(/gi, 'Response.Write(');

  return line;
}

function breakLongLine(line, maxLength, indent, indentLevel) {
  if (line.length <= maxLength) return line;

  // Break long lines at operators or commas
  const breakPoints = [' And ', ' Or ', ' + ', ' - ', ' * ', ' / ', ' & ', ', '];
  let result = line;

  for (const breakPoint of breakPoints) {
    if (result.includes(breakPoint)) {
      const parts = result.split(breakPoint);
      if (parts.length > 1) {
        result = parts[0] + breakPoint.trim() + ' _\n' +
          indent.repeat(indentLevel + 1) + parts.slice(1).join(breakPoint);
        if (result.length <= maxLength) break;
      }
    }
  }

  return result;
}

function applyAssignmentAlignment(lines, assignments, indentSize) {
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

// Read the ASP file
const fs = require('fs');
const code = fs.readFileSync('aspExampleFortesting.asp', 'utf8');

const options = {
  indentSize: 4,
  indentChar: ' ',
  alignAssignments: false,
  preserveAspComments: true,
  maxLineLength: 120
};

console.log('Starting beautification...');
const start = Date.now();
const result = beautifyVbscript(code, options);
const end = Date.now();
console.log(`Beautification took ${end - start} ms`);
console.log('First 1000 characters of result:');
console.log(result.substring(0, 1000));
console.log('\n...');
console.log('Last 1000 characters of result:');
console.log(result.substring(result.length - 1000));