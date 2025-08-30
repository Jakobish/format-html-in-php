'use strict';

import * as get from 'lodash.get';
import * as has from 'lodash.has';
import * as vscode from 'vscode';
import beautifyHtml from './beautifyHtml';
import optionsFromVSCode from './optionsFromVsCode';

class DocumentWatcher {

  private disposable: vscode.Disposable;

  constructor() {
    const subscriptions: vscode.Disposable[] = [];
    subscriptions.push(vscode.workspace.onWillSaveTextDocument(event => {
      const editor = vscode.window.activeTextEditor;
      if (['asp', 'vbscript'].includes(editor.document.languageId)) {
        const cursor = editor.selection.active;
        const last = editor.document.lineAt(editor.document.lineCount - 1);
        const range = new vscode.Range(
          new vscode.Position(0, 0),
          last.range.end
        );
        event.waitUntil(this.doPreSaveTransformations(
          event.document,
          event.reason
        ).then((content) => {
          editor.edit(edit => {
            if (content !== '') {
              edit.replace(range, content);
            }
          }).then(success => {
            if (success && content !== '') {
              const origSelection = new vscode.Selection(cursor, cursor);
              editor.selection = origSelection;
            }
          });
        }));
      }
    }));
    this.disposable = vscode.Disposable.from.apply(this, subscriptions);
  }

  public dispose() {
    this.disposable.dispose();
  }

  private async doPreSaveTransformations(
    doc: vscode.TextDocument,
    reason: vscode.TextDocumentSaveReason
  ): Promise<string> {
    const config = vscode.workspace.getConfiguration();
    const aspScopedFormat = has(config, '[asp]');
    let aspScopedFormatVal = false;
    if (aspScopedFormat) {
      const aspScopedObj = get(config, '[asp]');
      if (aspScopedObj['editor.formatOnSave']) {
        aspScopedFormatVal = true;
      }
    }
    if (config.get('formatHtmlInAsp.formatOnSave') === true || config.editor.formatOnSave === true || aspScopedFormatVal === true) {
      const html = vscode.window.activeTextEditor.document.getText();
      const options = optionsFromVSCode(config);
      return beautifyHtml(html, options);
    }
    return '';
  }

}

export function activate(formatHtmlInAsp: vscode.ExtensionContext) {
  const docWatch = new DocumentWatcher();
  formatHtmlInAsp.subscriptions.push(docWatch);
  formatHtmlInAsp.subscriptions.push(vscode.commands.registerCommand('formatHtmlInAsp.format', () => {
    const editor = vscode.window.activeTextEditor;
    if (!editor) {
      vscode.window.showErrorMessage('No active editor found');
      return;
    }

    try {
      const config = vscode.workspace.getConfiguration();
      const cursor = editor.selection.active;
      const last = editor.document.lineAt(editor.document.lineCount - 1);
      const range = new vscode.Range(
        new vscode.Position(0, 0),
        last.range.end
      );
      const html = editor.document.getText();
      const options = optionsFromVSCode(config);
      const formattedHtml = beautifyHtml(html, options);

      editor.edit(edit => {
        edit.replace(range, formattedHtml);
      }).then(success => {
        if (success) {
          const origSelection = new vscode.Selection(cursor, cursor);
          editor.selection = origSelection;
          vscode.window.showInformationMessage('HTML formatted successfully');
        } else {
          vscode.window.showErrorMessage('Failed to format HTML');
        }
      });
    } catch (error) {
      vscode.window.showErrorMessage(`Formatting error: ${error.message}`);
    }
  }));

  formatHtmlInAsp.subscriptions.push(vscode.commands.registerCommand('formatHtmlInAsp.formatAllOpen', async () => {
    const aspDocuments = vscode.workspace.textDocuments.filter(doc =>
      ['asp', 'vbscript'].includes(doc.languageId)
    );

    if (aspDocuments.length === 0) {
      vscode.window.showWarningMessage('No ASP files are currently open');
      return;
    }

    const config = vscode.workspace.getConfiguration();
    let successCount = 0;
    let errorCount = 0;

    for (const doc of aspDocuments) {
      try {
        await vscode.window.showTextDocument(doc);
        const editor = vscode.window.activeTextEditor;
        if (editor) {
          const last = doc.lineAt(doc.lineCount - 1);
          const range = new vscode.Range(
            new vscode.Position(0, 0),
            last.range.end
          );
          const html = doc.getText();
          const options = optionsFromVSCode(config);
          const formattedHtml = beautifyHtml(html, options);

          await editor.edit(edit => {
            edit.replace(range, formattedHtml);
          });
          successCount++;
        }
      } catch (error) {
        errorCount++;
        console.error(`Error formatting ${doc.fileName}:`, error);
      }
    }

    if (successCount > 0) {
      vscode.window.showInformationMessage(`Formatted ${successCount} ASP file(s) successfully${errorCount > 0 ? `, ${errorCount} failed` : ''}`);
    } else {
      vscode.window.showErrorMessage('Failed to format any files');
    }
  }));

  formatHtmlInAsp.subscriptions.push(vscode.commands.registerCommand('formatHtmlInAsp.formatSelection', () => {
    const editor = vscode.window.activeTextEditor;
    if (!editor) {
      vscode.window.showErrorMessage('No active editor found');
      return;
    }

    if (editor.selection.isEmpty) {
      vscode.window.showWarningMessage('Please select text to format');
      return;
    }

    try {
      const config = vscode.workspace.getConfiguration();
      const selection = editor.selection;
      const selectedText = editor.document.getText(selection);
      const options = optionsFromVSCode(config);
      const formattedHtml = beautifyHtml(selectedText, options);

      editor.edit(edit => {
        edit.replace(selection, formattedHtml);
      }).then(success => {
        if (success) {
          vscode.window.showInformationMessage('Selection formatted successfully');
        } else {
          vscode.window.showErrorMessage('Failed to format selection');
        }
      });
    } catch (error) {
      vscode.window.showErrorMessage(`Formatting error: ${error.message}`);
    }
  }));

  // Format entire workspace
  formatHtmlInAsp.subscriptions.push(vscode.commands.registerCommand('formatHtmlInAsp.formatWorkspace', async () => {
    const choice = await vscode.window.showInformationMessage(
      'Format Classic ASP files in this workspace?',
      { modal: true },
      'Apply', 'Dry Run', 'Cancel'
    );
    if (!choice || choice === 'Cancel') return;
    const dryRun = (choice === 'Dry Run');

    const config = vscode.workspace.getConfiguration();
    const include = '**/*.{asp,asa,inc}';
    const excludePatterns: string[] = [];
    const filesExclude = config.get<any>('files.exclude') || {};
    const searchExclude = config.get<any>('search.exclude') || {};
    for (const [glob, enabled] of Object.entries(filesExclude)) {
      if (enabled === true) excludePatterns.push(glob as string);
    }
    for (const [glob, enabled] of Object.entries(searchExclude)) {
      if (enabled === true && !excludePatterns.includes(glob as string)) excludePatterns.push(glob as string);
    }
    // Always ignore common folders
    ['**/node_modules/**', '**/.git/**', '**/out/**'].forEach(g => {
      if (!excludePatterns.includes(g)) excludePatterns.push(g);
    });
    const exclude = excludePatterns.length ? `{${excludePatterns.join(',')}}` : undefined;

    const uris = await vscode.workspace.findFiles(include, exclude);
    if (uris.length === 0) {
      vscode.window.showInformationMessage('No Classic ASP files found in workspace.');
      return;
    }

    await vscode.window.withProgress({
      location: vscode.ProgressLocation.Notification,
      title: `Formatting ${uris.length} Classic ASP file(s)` ,
      cancellable: true
    }, async (progress, token) => {
      const total = uris.length;
      let processed = 0;
      let changed = 0;
      let applied = 0;
      let errors = 0;

      const batchEdit = new vscode.WorkspaceEdit();

      for (const uri of uris) {
        if (token.isCancellationRequested) break;
        processed++;
        progress.report({ message: `${processed}/${total}: ${uri.fsPath.split('/').pop()}` , increment: (1 / total) * 100 });
        try {
          const doc = await vscode.workspace.openTextDocument(uri);
          const html = doc.getText();
          const options = optionsFromVSCode(config);
          const formattedHtml = beautifyHtml(html, options);
          if (formattedHtml !== html) {
            changed++;
            if (!dryRun) {
              const last = doc.lineAt(doc.lineCount - 1);
              const range = new vscode.Range(new vscode.Position(0, 0), last.range.end);
              batchEdit.replace(uri, range, formattedHtml);
            }
          }
        } catch (e) {
          errors++;
          console.error(`Format error for ${uri.fsPath}:`, e);
        }
      }

      if (!dryRun) {
        const success = await vscode.workspace.applyEdit(batchEdit);
        if (success) applied = changed;
      }

      if (token.isCancellationRequested) {
        vscode.window.showWarningMessage(`Formatting cancelled. ${processed}/${total} processed, ${changed} changed, ${errors} errors.`);
      } else if (dryRun) {
        vscode.window.showInformationMessage(`Dry run: ${processed} scanned, ${changed} would change, ${errors} errors.`);
      } else {
        vscode.window.showInformationMessage(`Formatted ${applied} file(s). Scanned ${processed}, ${errors} errors.`);
      }
    });
  }));
}
