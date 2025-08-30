'use strict';

function arrayUnique(dirtyArray) {
  var a = dirtyArray.concat();
  for(var i=0; i<a.length; ++i) {
    for(var j=i+1; j<a.length; ++j) {
      if(a[i] === a[j])
        a.splice(j--, 1);
    }
  }
  return a;
}

export default function(config) {
  let extraLinersVal;
  if (typeof config.html.format.extraLiners === 'string') {
    extraLinersVal = config.html.format.extraLiners
      .split(",")
      .map(s => s.trim());
  } else {
    extraLinersVal = config.html.format.extraLiners;
  }

  let contentUnformattedVal;
  if (typeof config.html.format.contentUnformatted === 'string') {
    contentUnformattedVal = config.html.format.contentUnformatted
      .split(",")
      .map(s => s.trim());
  } else {
    contentUnformattedVal = config.html.format.contentUnformatted;
  }

  const defaultContentUnformatted = [];

  let indentTabs;
  if (config.editor.insertSpaces === true) {
    indentTabs = false;
  } else {
    indentTabs = true;
  }

  const options = {
    indent_with_tabs: indentTabs,
    indent_size: config.editor.tabSize,
    indent_char: " ",
    extra_liners: extraLinersVal,
    content_unformatted: arrayUnique(defaultContentUnformatted.concat(contentUnformattedVal)),
    indent_handlebars: config.html.format.indentHandlebars,
    indent_inner_html: config.html.format.indentInnerHtml,
    max_preserve_newlines: config.html.format.maxPreserveNewLines,
    preserve_newlines: config.html.format.preserveNewLines,
    wrap_line_length: config.html.format.wrapLineLength,
    wrap_attributes: config.html.format.wrapAttributes,
    indent_scripts: "keep",
    end_with_newline: config.html.format.endWithNewline,
    // Alignment and whitespace cleanup
    alignServerBlocks: config.get('formatHtmlInAsp.alignServerBlocks') !== undefined ? config.get('formatHtmlInAsp.alignServerBlocks') : true,
    trimTrailingWhitespace: config.get('formatHtmlInAsp.trimTrailingWhitespace') !== undefined ? config.get('formatHtmlInAsp.trimTrailingWhitespace') : true,
    // New ASP preservation options (extension settings)
    preserveAspComments: config.get('formatHtmlInAsp.preserveAspComments') !== undefined ? config.get('formatHtmlInAsp.preserveAspComments') : true,
    preserveIncludeDirectives: config.get('formatHtmlInAsp.preserveIncludeDirectives') !== undefined ? config.get('formatHtmlInAsp.preserveIncludeDirectives') : true,
    preserveComplexExpressions: true,
    // ASP language/formatting options (VBScript only)
    formatVbscriptInAspBlocks: config.get('formatHtmlInAsp.formatVbscriptInAspBlocks') !== undefined ? config.get('formatHtmlInAsp.formatVbscriptInAspBlocks') : true,
    vbscriptConvertTabsToSpaces: config.get('formatHtmlInAsp.vbscriptConvertTabsToSpaces') !== undefined ? config.get('formatHtmlInAsp.vbscriptConvertTabsToSpaces') : true,
    vbscriptIndentSize: config.get('formatHtmlInAsp.vbscriptIndentSize') !== undefined ? config.get('formatHtmlInAsp.vbscriptIndentSize') : 4,
    // Advanced formatting options
    vbscriptAlignAssignments: config.get('formatHtmlInAsp.vbscriptAlignAssignments') !== undefined ? config.get('formatHtmlInAsp.vbscriptAlignAssignments') : false,
    maxLineLength: config.get('formatHtmlInAsp.maxLineLength') !== undefined ? config.get('formatHtmlInAsp.maxLineLength') : 120,
    // HTML/script normalization
    normalizeScriptClose: config.get('formatHtmlInAsp.normalizeScriptClose') !== undefined ? config.get('formatHtmlInAsp.normalizeScriptClose') : true
  };

  return options;
}
