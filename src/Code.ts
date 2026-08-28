// =============================================================================
// Code.ts — Entry points: menu, sidebar launcher, template helpers
// =============================================================================

/**
 * Runs when a user opens a document that has this add-on installed.
 * Creates the Aprimo menu in the host app's menu bar.
 */
function onOpen(e: any): void {
  const ui = _getUi();
  if (!ui) return;

  ui.createMenu('🔗 Aprimo DAM')
    .addItem('Open Aprimo DAM', 'openSidebar')
    .addSeparator()
    .addItem('Upload Current Document', 'openUploadTab')
    .addItem('Disconnect from Aprimo', 'clearAuth')
    .addToUi();
}

/**
 * Required entry point for editor add-ons installed from the marketplace.
 */
function onInstall(e: any): void {
  onOpen(e);
}

/**
 * One-time setup: creates an installable onOpen trigger for a specific document.
 * Run this ONCE from the Apps Script editor after pasting your test doc ID.
 * Only needed during development — published add-ons use simple triggers automatically.
 *
 * @param docId - The Google Doc ID (from its URL) to bind the trigger to
 */
/** Temporary dev helper — run once from the editor, then ignore. */
function _runSetup(): void {
  // Saves the web app URL so the sidebar can use it as the OAuth redirect_uri
  setWebAppUrl('https://script.google.com/macros/s/AKfycbzydVKcyTnKwnSeAIuh7pR4_ybH4OND49q_xoD3MAQp20dtHTw7lyZ3o01BBA4cL6B5Pw/exec');
  // Installs the onOpen trigger for the test document
  setupTriggerForDoc('132zi-Ta4NnVSH7C3U6VR_8qnIHCWhnbCfDq8hGdizQg');
}

function setupTriggerForDoc(docId: string): void {
  // Remove any existing onOpen triggers for this script to avoid duplicates
  ScriptApp.getProjectTriggers()
    .filter(function(t) { return t.getHandlerFunction() === 'onOpen'; })
    .forEach(function(t) { ScriptApp.deleteTrigger(t); });

  ScriptApp.newTrigger('onOpen')
    .forDocument(docId)
    .onOpen()
    .create();

  Logger.log('✅ onOpen trigger created for document: ' + docId);
}

/**
 * Opens the Aprimo sidebar. Detects which Google app is active and shows
 * the sidebar in the correct context.
 *
 * @param tab - Which tab to show first ('search' | 'upload' | 'recent')
 */
function openSidebar(tab?: string): void {
  const template = HtmlService.createTemplateFromFile('sidebar');
  template.initialTab = tab || 'search';

  const output = template
    .evaluate()
    .setTitle('Aprimo DAM')
    .setWidth(320);

  const ui = _getUi();
  if (ui) {
    ui.showSidebar(output);
  }
}

/**
 * Convenience shortcut — called from "Upload Current Document" menu item.
 */
function openUploadTab(): void {
  openSidebar('upload');
}

/**
 * Include helper for HTML templates. Used via <?!= include('filename') ?> in .html files.
 */
function include(filename: string): string {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Returns the Ui object for whichever Google host app is currently active.
 * Returns null if no supported host is found (e.g. called from a trigger).
 */
function _getUi(): GoogleAppsScript.Base.Ui | null {
  try { return DocumentApp.getUi(); } catch (e) {}
  try { return SpreadsheetApp.getUi(); } catch (e) {}
  try { return SlidesApp.getUi(); } catch (e) {}
  return null;
}

/**
 * Detects which Google app is currently active and returns basic info about
 * the open document. Called from the sidebar on load.
 */
function getActiveDocumentInfo(): object {
  try {
    const doc = DocumentApp.getActiveDocument();
    return {
      id: doc.getId(),
      name: doc.getName(),
      type: 'document',
      exportExtension: 'docx',
      exportMimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    };
  } catch (e) {}

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    return {
      id: ss.getId(),
      name: ss.getName(),
      type: 'spreadsheet',
      exportExtension: 'xlsx',
      exportMimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    };
  } catch (e) {}

  try {
    const pres = SlidesApp.getActivePresentation();
    return {
      id: pres.getId(),
      name: pres.getName(),
      type: 'presentation',
      exportExtension: 'pptx',
      exportMimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
    };
  } catch (e) {}

  return { id: null, name: 'Unknown', type: 'unknown' };
}
