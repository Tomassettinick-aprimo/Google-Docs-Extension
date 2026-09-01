// =============================================================================
// Upload.ts — Export the active Google document and upload it to Aprimo DAM
// =============================================================================

/**
 * Exports the currently active Google Doc / Sheet / Slides document to its
 * corresponding Office format, then uploads it to Aprimo.
 *
 * @param collectionId    - Target Aprimo collection ID (optional)
 * @param uploadAs        - 'new' = new asset | 'version' = new version of existing
 * @param existingAssetId - Required when uploadAs === 'version'
 * @param classificationId - Required when uploadAs === 'new'
 */
function uploadActiveDocument(
  collectionId: string,
  uploadAs: 'new' | 'version',
  existingAssetId?: string,
  classificationId?: string
): { success: boolean; recordId?: string; viewUrl?: string; error?: string } {
  const token = _getValidToken();
  if (!token) return { success: false, error: 'Not authenticated' };

  // Detect active document
  const docInfoRaw = getActiveDocumentInfo() as any;
  if (!docInfoRaw.id) {
    return { success: false, error: 'No active document found.' };
  }

  try {
    // Export via Drive API (converts Google format → Office format)
    const exportUrl = `https://www.googleapis.com/drive/v3/files/${docInfoRaw.id}/export?mimeType=${encodeURIComponent(docInfoRaw.exportMimeType)}`;

    const exportResp = UrlFetchApp.fetch(exportUrl, {
      headers: { 'Authorization': `Bearer ${ScriptApp.getOAuthToken()}` },
      muteHttpExceptions: true
    });

    if (exportResp.getResponseCode() !== 200) {
      return {
        success: false,
        error: `Export failed (${exportResp.getResponseCode()}). Make sure the document is not empty.`
      };
    }

    const fileName = `${docInfoRaw.name}.${docInfoRaw.exportExtension}`;
    const blob = exportResp.getBlob()
      .setName(fileName)
      .setContentType(docInfoRaw.exportMimeType);

    const result = uploadRecord(
      fileName,
      blob,
      collectionId,
      uploadAs === 'version' ? existingAssetId : undefined,
      classificationId
    );

    if (!result.success) return result;

    addToRecent({
      id: result.recordId!,
      title: docInfoRaw.name,
      fileExtension: docInfoRaw.exportExtension,
      action: uploadAs === 'version' ? 're-uploaded' : 'uploaded'
    });

    return {
      success: true,
      recordId: result.recordId,
      viewUrl: `https://${token.tenant}.aprimo.com/dam/contentitems/${result.recordId}`
    };
  } catch (e: any) {
    Logger.log(`uploadActiveDocument error: ${e.message}`);
    return { success: false, error: e.message };
  }
}
