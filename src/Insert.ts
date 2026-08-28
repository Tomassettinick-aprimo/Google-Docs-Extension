// =============================================================================
// Insert.ts — Insert assets from Aprimo into the active Google document
//
// Supports:
//   - Images → inline image at cursor (Docs) or centered on active slide (Slides)
//   - Office files (DOCX/XLSX/PPTX) → download, convert, open in new tab
//   - Tracks active "open & edit" sessions for the re-upload flow
// =============================================================================

interface EditSession {
  assetId: string;
  driveFileId: string;
  originalTitle: string;
  fileExtension: string;
  startedAt: number;
}

// ─── Insert Image ─────────────────────────────────────────────────────────────

/**
 * Fetches an image from Aprimo (using the server-side token) and inserts it
 * into the currently active Google Doc or Slides presentation.
 *
 * @param assetId    - Aprimo record ID (for recent tracking)
 * @param assetTitle - Display name for the asset
 * @param imageUrl   - Pre-signed download URL from Aprimo
 */
function insertImage(
  assetId: string,
  assetTitle: string,
  imageUrl: string
): { success: boolean; error?: string } {
  const headers = _getApiHeaders();
  if (!headers) return { success: false, error: 'Not authenticated' };

  try {
    // Preview/thumbnail URIs from Aprimo are pre-signed CDN links — no auth header needed.
    const imgResp = UrlFetchApp.fetch(imageUrl, {
      method: 'get',
      muteHttpExceptions: true,
      followRedirects: true
    });

    if (imgResp.getResponseCode() !== 200) {
      return { success: false, error: `Could not download image (${imgResp.getResponseCode()})` };
    }

    const blob = imgResp.getBlob();

    // Try Google Docs first
    const docsResult = _insertIntoDoc(blob);
    if (docsResult !== null) {
      addToRecent({ id: assetId, title: assetTitle, action: 'inserted' });
      return { success: true };
    }

    // Try Google Slides
    const slidesResult = _insertIntoSlide(blob);
    if (slidesResult !== null) {
      addToRecent({ id: assetId, title: assetTitle, action: 'inserted' });
      return { success: true };
    }

    return { success: false, error: 'No active Google Doc or Slides presentation found.' };
  } catch (e: any) {
    Logger.log(`insertImage error: ${e.message}`);
    return { success: false, error: e.message };
  }
}

function _insertIntoDoc(blob: GoogleAppsScript.Base.Blob): boolean | null {
  try {
    const doc = DocumentApp.getActiveDocument();
    const cursor = doc.getCursor();

    let image: GoogleAppsScript.Document.InlineImage;

    if (cursor) {
      image = cursor.insertInlineImage(blob);
    } else {
      // Fall back to end of document body
      image = doc.getBody().appendImage(blob);
    }

    // Cap width at 400 px, preserve aspect ratio
    if (image) {
      const maxW = 400;
      const w = image.getWidth();
      if (w > maxW) {
        image.setWidth(maxW);
        image.setHeight(Math.round(image.getHeight() * (maxW / w)));
      }
    }

    return true;
  } catch (e) {
    return null;
  }
}

function _insertIntoSlide(blob: GoogleAppsScript.Base.Blob): boolean | null {
  try {
    const pres = SlidesApp.getActivePresentation();
    const selection = pres.getSelection();
    const page = selection.getCurrentPage() as GoogleAppsScript.Slides.Slide;
    const slideW = pres.getPageWidth();
    const slideH = pres.getPageHeight();

    const imgElement = page.insertImage(blob);

    // Scale to 60% of slide width, centered
    const targetW = slideW * 0.6;
    const ratio = targetW / imgElement.getWidth();
    imgElement.setWidth(targetW);
    imgElement.setHeight(imgElement.getHeight() * ratio);
    imgElement.setLeft((slideW - imgElement.getWidth()) / 2);
    imgElement.setTop((slideH - imgElement.getHeight()) / 2);

    return true;
  } catch (e) {
    return null;
  }
}

// ─── Open & Edit Flow ─────────────────────────────────────────────────────────

/**
 * Downloads an Aprimo Office asset (DOCX/XLSX/PPTX) to Drive and opens it.
 * Stores an edit session so "Save as New Version" knows what to re-upload.
 *
 * Returns the Drive file ID and a URL the sidebar can open in a new tab.
 */
function openAssetForEditing(
  assetId: string,
  assetTitle: string
): { success: boolean; editUrl?: string; error?: string } {
  const headers = _getApiHeaders();
  if (!headers) return { success: false, error: 'Not authenticated' };

  try {
    // Get the download URL
    const downloadResult = getAssetDownloadUrl(assetId);
    if ('error' in downloadResult) return { success: false, error: downloadResult.error };

    const asset = getAsset(assetId);
    if ('error' in asset) return { success: false, error: asset.error };

    const ext = asset.fileExtension.toLowerCase();

    // Fetch the file content
    const fileResp = UrlFetchApp.fetch(downloadResult.url, {
      method: 'get',
      headers: { 'Authorization': headers['Authorization'] },
      muteHttpExceptions: true,
      followRedirects: true
    });

    if (fileResp.getResponseCode() !== 200) {
      return { success: false, error: `Download failed (${fileResp.getResponseCode()})` };
    }

    const officeMime: { [key: string]: string } = {
      docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      doc:  'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      xls:  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      pptx: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
      ppt:  'application/vnd.openxmlformats-officedocument.presentationml.presentation'
    };

    const googleMime: { [key: string]: string } = {
      docx: MimeType.GOOGLE_DOCS,
      doc:  MimeType.GOOGLE_DOCS,
      xlsx: MimeType.GOOGLE_SHEETS,
      xls:  MimeType.GOOGLE_SHEETS,
      pptx: MimeType.GOOGLE_SLIDES,
      ppt:  MimeType.GOOGLE_SLIDES
    };

    const mime = officeMime[ext] || 'application/octet-stream';
    const blob = fileResp.getBlob()
      .setName(`[Aprimo] ${assetTitle}.${ext}`)
      .setContentType(mime);

    // Save to Drive using DriveApp (no Advanced Service needed).
    // Drive automatically opens Office files in Google Docs/Sheets/Slides
    // compatibility mode when you navigate to their edit URL.
    const driveFile = DriveApp.createFile(blob);
    const driveFileId: string = driveFile.getId();

    // Drive opens Office files in the appropriate Google editor automatically
    const editUrl = `https://drive.google.com/file/d/${driveFileId}/edit`;

    // Persist the edit session
    const session: EditSession = {
      assetId,
      driveFileId,
      originalTitle: assetTitle,
      fileExtension: ext,
      startedAt: Date.now()
    };

    PropertiesService.getUserProperties()
      .setProperty('aprimo_current_edit', JSON.stringify(session));

    addToRecent({ id: assetId, title: assetTitle, action: 'opened' });

    return { success: true, editUrl };
  } catch (e: any) {
    Logger.log(`openAssetForEditing error: ${e.message}`);
    return { success: false, error: e.message };
  }
}

/**
 * Returns the active edit session (if any) so the sidebar can show the
 * "In Progress" banner.
 */
function getCurrentEditSession(): EditSession | null {
  const json = PropertiesService.getUserProperties()
    .getProperty('aprimo_current_edit');
  return json ? JSON.parse(json) : null;
}

/**
 * Exports the Drive file from the edit session back to its Office format and
 * uploads it to Aprimo as a new master file version.
 */
function saveEditAsNewVersion(): { success: boolean; assetId?: string; error?: string } {
  const props = PropertiesService.getUserProperties();
  const json = props.getProperty('aprimo_current_edit');
  if (!json) return { success: false, error: 'No active edit session found.' };

  const session: EditSession = JSON.parse(json);

  try {
    const exportMime: { [key: string]: string } = {
      docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      doc:  'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      xls:  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      pptx: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
      ppt:  'application/vnd.openxmlformats-officedocument.presentationml.presentation'
    };

    const mime = exportMime[session.fileExtension] || 'application/octet-stream';
    const fileName = `${session.originalTitle}.${session.fileExtension}`;

    // File is stored natively in Drive (DOCX/XLSX/PPTX), retrieve it directly.
    // Drive's compatibility-mode edits are saved back to the original Office format.
    const driveFile = DriveApp.getFileById(session.driveFileId);
    const blob = driveFile.getBlob()
      .setName(fileName)
      .setContentType(mime);

    // Upload as a new version
    const result = uploadRecord(fileName, blob, '', '', session.assetId);
    if (!result.success) return result;

    // Clean up: trash the temp Drive file and clear the session
    try { DriveApp.getFileById(session.driveFileId).setTrashed(true); } catch (e) {}
    props.deleteProperty('aprimo_current_edit');

    addToRecent({
      id: session.assetId,
      title: session.originalTitle,
      action: 're-uploaded'
    });

    return { success: true, assetId: session.assetId };
  } catch (e: any) {
    Logger.log(`saveEditAsNewVersion error: ${e.message}`);
    return { success: false, error: e.message };
  }
}

/**
 * Cancels the current edit session without saving. Trashes the temp Drive file.
 */
function cancelEditSession(): void {
  const props = PropertiesService.getUserProperties();
  const json = props.getProperty('aprimo_current_edit');
  if (!json) return;

  const session: EditSession = JSON.parse(json);
  try { DriveApp.getFileById(session.driveFileId).setTrashed(true); } catch (e) {}
  props.deleteProperty('aprimo_current_edit');
}
