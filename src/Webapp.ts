// =============================================================================
// Webapp.ts — Apps Script Web App entry point
//
// This script is deployed as a web app ("Execute as: User accessing the web
// app", "Who has access: Anyone"). Its URL is used as the OAuth redirect_uri.
//
// Deploy URL format: https://script.google.com/macros/s/{SCRIPT_ID}/exec
// Register this URL as the redirect_uri in your Aprimo OAuth client settings.
// =============================================================================

/**
 * Handles incoming GET requests. Two cases:
 *   ?code=... — OAuth callback after successful Aprimo login
 *   ?error=... — OAuth callback after user denied access
 */
function doGet(e: any): GoogleAppsScript.HTML.HtmlOutput {
  const params = (e && e.parameter) ? e.parameter : {};

  if (params.code) {
    // Use the stored web app URL as redirect_uri — must match exactly what was
    // sent in the authorization request (both come from ScriptProperties).
    const redirectUri = PropertiesService.getScriptProperties()
      .getProperty('aprimo_webapp_url') || ScriptApp.getService().getUrl();
    const result = exchangeToken(params.code, redirectUri);

    const template = HtmlService.createTemplateFromFile('callback');
    template.success = result.success;
    template.errorMessage = result.error || '';

    return template.evaluate()
      .setTitle(result.success ? 'Aprimo — Connected!' : 'Aprimo — Connection Failed')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  if (params.error) {
    const template = HtmlService.createTemplateFromFile('callback');
    template.success = false;
    template.errorMessage = params.error_description
      ? `${params.error}: ${params.error_description}`
      : `Authorization was denied (${params.error}).`;

    return template.evaluate()
      .setTitle('Aprimo — Connection Failed')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  // Direct navigation (not a redirect) — show an informational page
  return HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
      <head>
        <title>Aprimo DAM Add-on</title>
        <style>
          body { font-family: -apple-system, sans-serif; padding: 40px; color: #333; }
          h1 { color: #0057A8; }
        </style>
      </head>
      <body>
        <h1>Aprimo DAM Add-on</h1>
        <p>This URL is the OAuth redirect endpoint for the Aprimo Google Workspace Add-on.</p>
        <p>Open this add-on from within Google Docs, Sheets, or Slides.</p>
      </body>
    </html>
  `).setTitle('Aprimo DAM Add-on');
}
