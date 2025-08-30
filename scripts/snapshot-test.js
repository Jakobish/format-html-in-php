/* Snapshot test for formatting complex Classic ASP */
/* eslint-disable no-console */
const fs = require('fs');
const path = require('path');

function loadFormatter() {
  try {
    // Use compiled output
    const mod = require(path.join(__dirname, '..', 'out', 'beautifyHtml.js'));
    return mod && (mod.default || mod);
  } catch (e) {
    console.error('Failed to load formatter. Ensure `npm run compile` succeeded.');
    throw e;
  }
}

function getDefaultOptions() {
  return {
    // HTML formatting (js-beautify) options
    indent_with_tabs: false,
    indent_size: 2,
    indent_char: ' ',
    wrap_line_length: 120,
    wrap_attributes: 'auto',
    indent_inner_html: false,
    preserve_newlines: true,
    max_preserve_newlines: 2,
    end_with_newline: false,
    extra_liners: [],
    content_unformatted: [],
    indent_handlebars: false,
    indent_scripts: 'keep',
    // Extension-specific ASP options
    preserveAspComments: true,
    preserveIncludeDirectives: true,
    preserveComplexExpressions: true,
    detectAspLanguage: true,
    formatVbscriptInAspBlocks: true,
    formatJscriptInAspBlocks: false,
    vbscriptIndentSize: 4,
    vbscriptConvertTabsToSpaces: true,
    jscriptIndentSize: 2,
    vbscriptAlignAssignments: false,
    jscriptSemicolons: true,
    maxLineLength: 120,
    normalizeScriptClose: true
  };
}

function runOne(inputPath, beautifyHtml, repoRoot, snapshotsDir) {
  const baseName = path.basename(inputPath).replace(/\.[^.]+$/, '') + '.formatted.asp';
  const snapshotPath = path.join(snapshotsDir, baseName);
  const actualPath = snapshotPath.replace(/\.asp$/, '.actual.asp');

  if (!fs.existsSync(inputPath)) {
    console.error('Input file not found:', inputPath);
    return 2;
  }

  const options = getDefaultOptions();
  const input = fs.readFileSync(inputPath, 'utf8');
  const formatted = beautifyHtml(input, options);

  if (!fs.existsSync(snapshotsDir)) fs.mkdirSync(snapshotsDir);

  if (!fs.existsSync(snapshotPath)) {
    fs.writeFileSync(snapshotPath, formatted, 'utf8');
    console.log('Snapshot created:', path.relative(repoRoot, snapshotPath));
    return 0;
  }

  const expected = fs.readFileSync(snapshotPath, 'utf8');
  const norm = s => s.replace(/\r\n/g, '\n').replace(/\n$/, '');
  if (norm(expected) === norm(formatted)) {
    console.log('Snapshot OK:', path.relative(repoRoot, snapshotPath));
    return 0;
  }

  fs.writeFileSync(actualPath, formatted, 'utf8');
  console.error('Snapshot mismatch. Wrote actual to:', path.relative(repoRoot, actualPath));
  console.error('To accept new output, replace snapshot with actual:');
  console.error('  mv', path.relative(repoRoot, actualPath), path.relative(repoRoot, snapshotPath));
  return 1;
}

(function main() {
  const beautifyHtml = loadFormatter();
  const repoRoot = path.join(__dirname, '..');
  const snapshotsDir = path.join(repoRoot, 'snapshots');
  const inputs = process.argv.slice(2);
  const files = inputs.length ? inputs : [
    // Keep defaults minimal to reduce churn; pass files on CLI to expand
    path.join(repoRoot, 'samples', 'vbscript-indent.asp'),
    path.join(repoRoot, 'samples', 'script-close.asp')
  ];

  let status = 0;
  for (const f of files) {
    const rc = runOne(f, beautifyHtml, repoRoot, snapshotsDir);
    if (rc !== 0) status = rc;
  }
  process.exit(status);
})();
