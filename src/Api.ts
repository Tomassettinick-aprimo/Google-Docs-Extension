// =============================================================================
// Api.ts — Aprimo DAM REST API client (server-side) — v2
//
// All requests go through UrlFetchApp so auth headers stay server-side.
// The sidebar calls these functions via google.script.run.
//
// API base: https://{tenant}.dam.aprimo.com/api/core
// Headers:  Authorization: Bearer {token} | API-VERSION: 1
// =============================================================================

interface Asset {
  id: string;
  title: string;
  fileExtension: string;
  fileSize: number;
  contentType: string;
  thumbnailUrl: string;   // raw Aprimo URI (for reference)
  thumbnailData: string;  // base64 data URI — safe to use in <img src> without auth
  downloadUrl: string;
  collectionLabel: string;
  modifiedOn: string;
}

interface SearchResult {
  assets: Asset[];
  total: number;
  page: number;
  hasMore: boolean;
  error?: string;
}

interface Collection {
  id: string;
  label: string;
}

// ─── Asset Search ─────────────────────────────────────────────────────────────

/**
 * Full-text search across Aprimo assets. Supports optional file type filtering.
 *
 * @param query     - Free-text search string
 * @param fileTypes - Array of extensions to filter on, e.g. ['png','jpg'] — empty = all
 * @param page      - Zero-indexed page number for pagination
 */
function searchAssets(query: string, fileTypes: string[], page: number): SearchResult {
  const headers = _getApiHeaders();
  const base = _getApiBase();
  if (!headers || !base) {
    return { assets: [], total: 0, page: 0, hasMore: false, error: 'Not authenticated' };
  }

  const pageSize = 24;
  const skip = page * pageSize;

  // Common headers for all Aprimo record requests.
  // select-Record: thumbnail,preview makes Aprimo embed pre-signed CDN URIs
  // directly at record._embedded.thumbnail.uri and record._embedded.preview.uri
  const commonHeaders: { [key: string]: string } = {
    'Authorization': headers['Authorization'],
    'API-VERSION': '1',
    'select-Record': 'thumbnail,preview'
  };

  // Extension → MIME type map for the search expression contentType filter
  const EXT_TO_MIME: { [key: string]: string } = {
    jpg: 'image/jpeg', jpeg: 'image/jpeg', png: 'image/png', gif: 'image/gif',
    webp: 'image/webp', svg: 'image/svg+xml', tiff: 'image/tiff', tif: 'image/tiff',
    heic: 'image/heic', pdf: 'application/pdf',
    docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    doc:  'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    xls:  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    pptx: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    ppt:  'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    mp4: 'video/mp4', mov: 'video/quicktime', mp3: 'audio/mpeg', wav: 'audio/wav'
  };

  const hasQuery    = !!(query && query.trim());
  const hasTypes    = !!(fileTypes && fileTypes.length > 0);

  try {
    let response: GoogleAppsScript.URL_Fetch.HTTPResponse;

    if (hasQuery || hasTypes) {
      // ── Text search or file-type filter ──────────────────────────────────
      // POST /api/core/search/records?page=N&pageSize=N
      // Body: { searchExpression: { expression: "..." } }
      // page is 1-indexed in the Aprimo search API
      let expression: string = hasQuery ? query.trim() : '';

      if (hasTypes) {
        const typeClause = fileTypes
          .map(ext => EXT_TO_MIME[ext.toLowerCase()] || ext)
          .map(mime => `contentType = "${mime}"`)
          .join(' OR ');
        expression = expression ? `(${expression}) AND (${typeClause})` : `(${typeClause})`;
      }

      const aprimoPage = page + 1;
      response = UrlFetchApp.fetch(
        `${base}/search/records?page=${aprimoPage}&pageSize=${pageSize}`,
        {
          method: 'post',
          headers: { ...commonHeaders, 'Content-Type': 'application/json' },
          payload: JSON.stringify({ searchExpression: { expression } }),
          muteHttpExceptions: true
        }
      );
    } else {
      // ── Simple listing (no query, no filter) ─────────────────────────────
      // GET /api/core/records with take/skip sent as HTTP headers (SDK pattern —
      // Aprimo's records endpoint reads pagination from headers, not query string)
      response = UrlFetchApp.fetch(`${base}/records`, {
        method: 'get',
        headers: {
          ...commonHeaders,
          'take':    String(pageSize),
          'skip':    String(skip),
          'orderby': 'modifiedon desc'
        },
        muteHttpExceptions: true
      });
    }

    const code = response.getResponseCode();

    if (code === 401) {
      return { assets: [], total: 0, page: 0, hasMore: false, error: 'Session expired — please reconnect.' };
    }
    if (code !== 200) {
      Logger.log(`searchAssets ${code}: ${response.getContentText()}`);
      return { assets: [], total: 0, page: 0, hasMore: false, error: `API error (${code})` };
    }

    const data = JSON.parse(response.getContentText());
    // HAL+JSON lists live in _embedded.items; fall back to flat items/value
    const items: any[] = data._embedded?.items || data.items || data.value || [];
    const total: number = data.totalCount || items.length;

    return { assets: items.map(_mapAsset), total, page, hasMore: skip + items.length < total };
  } catch (e: any) {
    Logger.log(`searchAssets error: ${e.message}`);
    return { assets: [], total: 0, page: 0, hasMore: false, error: e.message };
  }
}

// ─── Single Asset ─────────────────────────────────────────────────────────────

/**
 * Fetches full details for a single asset by its Aprimo record ID.
 */
function getAsset(id: string): Asset | { error: string } {
  const headers = _getApiHeaders();
  const base = _getApiBase();
  if (!headers || !base) return { error: 'Not authenticated' };

  const expand = encodeURIComponent('masterfilelatestversion($expand=thumbnail,preview)');
  const url = `${base}/records/${id}?$expand=${expand}`;

  try {
    const response = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: {
        'Authorization': headers['Authorization'],
        'API-VERSION': '1',
        'select-Record': 'thumbnail,preview'
      },
      muteHttpExceptions: true
    });

    if (response.getResponseCode() !== 200) {
      return { error: `Failed to fetch asset (${response.getResponseCode()})` };
    }

    return _mapAsset(JSON.parse(response.getContentText()));
  } catch (e: any) {
    return { error: e.message };
  }
}

// ─── Download URL ─────────────────────────────────────────────────────────────

/**
 * Fetches a pre-signed download URL for the latest master file version.
 * Gets the record with masterfilelatestversion expanded and extracts the URI.
 */
function getAssetDownloadUrl(id: string): { url: string } | { error: string } {
  const headers = _getApiHeaders();
  const base = _getApiBase();
  if (!headers || !base) return { error: 'Not authenticated' };

  try {
    // select-Record: thumbnail,preview makes Aprimo return pre-signed CDN URIs
    const expand = encodeURIComponent('masterfilelatestversion($expand=thumbnail,preview)');
    const response = UrlFetchApp.fetch(`${base}/records/${id}?$expand=${expand}`, {
      method: 'get',
      headers: {
        'Authorization': headers['Authorization'],
        'API-VERSION': '1',
        'select-Record': 'thumbnail,preview'
      },
      muteHttpExceptions: true
    });

    const code = response.getResponseCode();
    const body = response.getContentText();

    if (code !== 200) {
      Logger.log(`getAssetDownloadUrl (${code}): ${body}`);
      return { error: `Could not get download URL (${code}): ${body}` };
    }

    const data = JSON.parse(body);
    // HAL+JSON: nested resources in _embedded, with flat fallback
    const mf =
      data._embedded?.masterfilelatestversion ||
      data.masterfilelatestversion ||
      data._embedded?.latestVersion ||
      data.latestVersion ||
      {};

    const preview = mf._embedded?.preview || mf.preview || {};
    const thumb   = mf._embedded?.thumbnail || mf.thumbnail || {};

    // Prefer preview URI (pre-signed, no auth needed); fall back to other known fields
    const downloadUrl: string =
      preview.uri || preview.url ||
      thumb.uri   || thumb.url   ||
      mf.downloaduri || mf.downloadUrl || mf.downloadUri ||
      mf.uri || mf.contentUri || mf.contenturl || '';

    if (!downloadUrl) {
      Logger.log(`getAssetDownloadUrl — no URL in response: ${body}`);
      return { error: 'No download URL available for this asset. The file may still be processing.' };
    }
    return { url: downloadUrl };
  } catch (e: any) {
    return { error: e.message };
  }
}

// ─── Classifications ──────────────────────────────────────────────────────────

/**
 * Returns a flat list of all Aprimo classifications (asset types).
 * Classification is REQUIRED when creating a new record.
 */
function getClassifications(): { classifications: { id: string; label: string }[] } | { error: string } {
  const headers = _getApiHeaders();
  const base = _getApiBase();
  if (!headers || !base) return { error: 'Not authenticated' };

  try {
    const response = UrlFetchApp.fetch(
      `${base}/classifications?$top=200`,
      {
        method: 'get',
        headers: { 'Authorization': headers['Authorization'], 'API-VERSION': '1' },
        muteHttpExceptions: true
      }
    );

    const code = response.getResponseCode();
    const body = response.getContentText();
    if (code !== 200) {
      Logger.log(`getClassifications (${code}): ${body}`);
      return { error: `Could not fetch classifications (${code})` };
    }

    const data = JSON.parse(body);
    const items: any[] = data._embedded?.items || data.items || data.value || [];

    return {
      classifications: items.map((c: any) => ({
        id: c.id,
        label: c.label || c.name || c.id
      }))
    };
  } catch (e: any) {
    return { error: e.message };
  }
}

// ─── Collections ──────────────────────────────────────────────────────────────

/**
 * Returns a flat list of the user's accessible collections, sorted by label.
 * Used to populate the collection picker in the Upload tab.
 */
function getCollections(): { collections: Collection[] } | { error: string } {
  const headers = _getApiHeaders();
  const base = _getApiBase();
  if (!headers || !base) return { error: 'Not authenticated' };

  try {
    const response = UrlFetchApp.fetch(
      `${base}/collections?$top=100&$orderby=label asc`,
      {
        method: 'get',
        headers: { 'Authorization': headers['Authorization'], 'API-VERSION': '1' },
        muteHttpExceptions: true
      }
    );

    if (response.getResponseCode() !== 200) {
      return { error: `Could not fetch collections (${response.getResponseCode()})` };
    }

    const data = JSON.parse(response.getContentText());
    const items: any[] = data._embedded?.items || data.items || data.value || [];

    return {
      collections: items.map((c: any) => ({
        id: c.id,
        label: c.label || c.name || c.id
      }))
    };
  } catch (e: any) {
    return { error: e.message };
  }
}

// ─── Upload ───────────────────────────────────────────────────────────────────

/**
 * Creates a new Aprimo record and uploads a file blob as its master file.
 * If existingRecordId is provided, skips record creation and uploads a new
 * version to the existing record.
 *
 * @param fileName          - Display name + extension (e.g. "Brief.docx")
 * @param blob              - The file content as a Google Apps Script Blob
 * @param collectionId      - Aprimo collection ID to place the asset in (optional)
 * @param tags              - Comma-separated tags string (unused by API, kept for UI)
 * @param existingRecordId  - If set, upload as a new version of this record
 * @param classificationId  - Aprimo classification ID (REQUIRED for new records)
 */
function uploadRecord(
  fileName: string,
  blob: GoogleAppsScript.Base.Blob,
  collectionId: string,
  tags: string,
  existingRecordId?: string,
  classificationId?: string
): { success: boolean; recordId?: string; error?: string } {
  const token = _getValidToken();
  const headers = _getApiHeaders();
  const base = _getApiBase();
  if (!headers || !base || !token) return { success: false, error: 'Not authenticated' };

  try {
    let recordId = existingRecordId;

    // ── New asset: upload file then create record referencing the upload token ──
    if (!recordId) {
      if (!classificationId) {
        return { success: false, error: 'A classification (asset type) is required. Please select one from the dropdown.' };
      }

      // Step 1: POST file to the Aprimo staging area (aprimo.com/uploads, NOT dam API)
      // UrlFetchApp builds multipart/form-data automatically when payload is an object with Blobs.
      const uploadBlob = blob.setName(fileName);
      const stagingResp = UrlFetchApp.fetch(
        `https://${token.tenant}.aprimo.com/uploads`,
        {
          method: 'post',
          headers: { 'Authorization': headers['Authorization'] },
          payload: { file1: uploadBlob },  // UrlFetchApp handles multipart boundary
          muteHttpExceptions: true
        }
      );

      const stagingCode = stagingResp.getResponseCode();
      const stagingBody = stagingResp.getContentText();
      if (stagingCode !== 200 && stagingCode !== 201) {
        Logger.log(`Staging upload failed (${stagingCode}): ${stagingBody}`);
        return { success: false, error: `File upload failed (${stagingCode}): ${stagingBody}` };
      }

      const uploadToken: string = JSON.parse(stagingBody).token;
      if (!uploadToken) {
        return { success: false, error: 'Aprimo staging upload returned no token.' };
      }

      // Step 2: Create the record referencing the upload token
      // title is { value: string }, classification uses addOrUpdate array
      const recordBody: any = {
        title: { value: fileName.replace(/\.[^.]+$/, '') },
        classifications: { addOrUpdate: [{ id: classificationId }] },
        files: {
          master: uploadToken,
          addOrUpdate: [{
            versions: {
              addOrUpdate: [{ id: uploadToken, fileName }]
            }
          }]
        }
      };

      if (collectionId) {
        recordBody.collections = { addOrUpdate: [{ id: collectionId }] };
      }

      const createResp = UrlFetchApp.fetch(`${base}/records`, {
        method: 'post',
        headers,
        payload: JSON.stringify(recordBody),
        muteHttpExceptions: true
      });

      const createCode = createResp.getResponseCode();
      const createBody = createResp.getContentText();
      if (createCode !== 200 && createCode !== 201) {
        Logger.log(`Create record failed (${createCode}): ${createBody}`);
        return { success: false, error: `Failed to create record (${createCode}): ${createBody}` };
      }

      recordId = JSON.parse(createBody).id;

    } else {
      // ── New version of existing record ──
      // Stage the file first, then add a new master file version referencing the token
      const uploadBlob = blob.setName(fileName);
      const stagingResp = UrlFetchApp.fetch(
        `https://${token.tenant}.aprimo.com/uploads`,
        {
          method: 'post',
          headers: { 'Authorization': headers['Authorization'] },
          payload: { file1: uploadBlob },
          muteHttpExceptions: true
        }
      );

      const stagingCode = stagingResp.getResponseCode();
      const stagingBody = stagingResp.getContentText();
      if (stagingCode !== 200 && stagingCode !== 201) {
        Logger.log(`Staging upload (new version) failed (${stagingCode}): ${stagingBody}`);
        return { success: false, error: `File upload failed (${stagingCode}): ${stagingBody}` };
      }

      const uploadToken: string = JSON.parse(stagingBody).token;

      const versionResp = UrlFetchApp.fetch(`${base}/records/${recordId}`, {
        method: 'put',
        headers,
        payload: JSON.stringify({
          files: {
            master: uploadToken,
            addOrUpdate: [{
              versions: {
                addOrUpdate: [{ id: uploadToken, fileName }]
              }
            }]
          }
        }),
        muteHttpExceptions: true
      });

      const versionCode = versionResp.getResponseCode();
      if (versionCode !== 200 && versionCode !== 204) {
        const versionBody = versionResp.getContentText();
        Logger.log(`Add version failed (${versionCode}): ${versionBody}`);
        return { success: false, error: `Failed to add new version (${versionCode}): ${versionBody}` };
      }
    }

    return { success: true, recordId };
  } catch (e: any) {
    Logger.log(`uploadRecord error: ${e.message}`);
    return { success: false, error: e.message };
  }
}

// ─── Recent Assets ────────────────────────────────────────────────────────────

/**
 * Appends an asset to the user's recent list (stored in UserProperties).
 * Keeps the 10 most recent, deduped by id.
 */
function addToRecent(asset: Partial<Asset> & { id: string; action?: string }): void {
  const props = PropertiesService.getUserProperties();
  const existing: any[] = JSON.parse(props.getProperty('aprimo_recent_assets') || '[]');

  const deduped = existing.filter((a: any) => a.id !== asset.id);
  deduped.unshift({ ...asset, accessedAt: Date.now() });

  props.setProperty('aprimo_recent_assets', JSON.stringify(deduped.slice(0, 10)));
}

/**
 * Returns the user's recent asset list.
 */
function getRecentAssets(): any[] {
  const props = PropertiesService.getUserProperties();
  return JSON.parse(props.getProperty('aprimo_recent_assets') || '[]');
}

// ─── Internal helpers ─────────────────────────────────────────────────────────

function _mapAsset(item: any): Asset {
  // ── Thumbnail / Preview ────────────────────────────────────────────────────
  // When select-Record: thumbnail,preview is sent, Aprimo embeds pre-signed CDN
  // URIs DIRECTLY on the record at _embedded.thumbnail / _embedded.preview —
  // NOT nested inside _embedded.masterfilelatestversion._embedded.thumbnail.
  const thumb   = item._embedded?.thumbnail || item.thumbnail   || {};
  const preview = item._embedded?.preview   || item.preview     || {};

  const thumbnailUrl: string = thumb.uri   || thumb.url   || '';
  // preview.uri is full-quality rendition (use for insert); fall back to thumbnail
  const previewUrl: string   = preview.uri || preview.url || thumbnailUrl;

  // ── masterfilelatestversion — only present when $expand was used ───────────
  // (e.g. getAsset / getAssetDownloadUrl). Provides filesize, contenttype, etc.
  const mf =
    item._embedded?.masterfilelatestversion ||
    item.masterfilelatestversion ||
    item._embedded?.latestVersion ||
    item.latestVersion ||
    {};

  // ── Title ─────────────────────────────────────────────────────────────────
  // Aprimo localized fields come back as { value: string } objects. Handle both
  // that format and the plain string form, falling back to the record ID (GUID).
  const rawTitle = item.title;
  const title: string = typeof rawTitle === 'string'
    ? (rawTitle || item.id)
    : ((rawTitle as any)?.value || item.id);

  // ── File extension ─────────────────────────────────────────────────────────
  // Prefer masterfilelatestversion (when $expand was used). For search/listing
  // results, fall back to the extension embedded in the thumbnail/preview object.
  const thumbExt   = (thumb.extension   || '').toLowerCase();
  const previewExt = (preview.extension || '').toLowerCase();
  const fileExtension: string = (
    mf.fileextension || mf.fileExtension || thumbExt || previewExt
  ).toLowerCase();

  // ── Content type ──────────────────────────────────────────────────────────
  const EXT_TO_MIME: { [key: string]: string } = {
    jpg: 'image/jpeg', jpeg: 'image/jpeg', png: 'image/png', gif: 'image/gif',
    webp: 'image/webp', svg: 'image/svg+xml', tiff: 'image/tiff', tif: 'image/tiff',
    heic: 'image/heic', pdf: 'application/pdf',
    docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    doc:  'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    xls:  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    pptx: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    ppt:  'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    mp4: 'video/mp4', mov: 'video/quicktime', mp3: 'audio/mpeg', wav: 'audio/wav'
  };
  const contentType: string =
    mf.contenttype || mf.contentType || EXT_TO_MIME[fileExtension] || '';

  return {
    id: item.id,
    title,
    fileExtension,
    fileSize: mf.filesize || mf.fileSize || 0,
    contentType,
    thumbnailUrl,
    thumbnailData: '',   // unused — thumbnails are pre-signed CDN URLs
    downloadUrl: previewUrl ||
      mf.downloaduri || mf.downloadUrl || mf.downloadUri ||
      mf.uri || mf.contentUri || mf.contenturl || '',
    collectionLabel:
      item._embedded?.collection?.label ||
      item.collection?.label  ||
      item.collection?.name   ||
      '',
    modifiedOn: item.modifiedon || item.modifiedOn || ''
  };
}

/**
 * Public entry point called per-card from the sidebar for lazy thumbnail loading.
 * Returns a base64 data URI, or '' if unavailable.
 */
function getThumbnailBase64(thumbnailUrl: string): string {
  return _fetchThumbnailBase64(thumbnailUrl);
}

/**
 * Fetches a thumbnail from Aprimo server-side (with auth) and returns a
 * base64 data URI safe for use in <img src> without needing auth headers.
 * Returns '' on any failure so callers can fall back to emoji.
 */
function _fetchThumbnailBase64(thumbnailUrl: string): string {
  if (!thumbnailUrl) return '';
  const headers = _getApiHeaders();
  if (!headers) return '';
  try {
    const resp = UrlFetchApp.fetch(thumbnailUrl, {
      method: 'get',
      headers: { 'Authorization': headers['Authorization'] },
      muteHttpExceptions: true,
      followRedirects: true
    });
    if (resp.getResponseCode() !== 200) return '';
    const blob = resp.getBlob();
    const mime = blob.getContentType() || 'image/jpeg';
    const b64 = Utilities.base64Encode(blob.getBytes());
    return `data:${mime};base64,${b64}`;
  } catch (e) {
    return '';
  }
}
