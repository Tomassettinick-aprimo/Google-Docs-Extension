// =============================================================================
// Api.ts — Aprimo DAM REST API client (server-side) — v2
//
// All requests go through UrlFetchApp so auth headers stay server-side.
// The sidebar calls these functions via google.script.run.
//
// API base: https://{tenant}.dam.aprimo.com/api/core
// Headers:  Authorization: Bearer {token} | API-VERSION: 1
// =============================================================================

interface AssetField {
  label: string;
  value: string;
}

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
  classificationLabel: string;
  modifiedOn: string;
  createdOn: string;
  status: string;
  fields: AssetField[];   // all non-empty custom fields, for the detail panel
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
  // select-Record embeds pre-signed CDN URIs, the master file info (fileextension,
  // contenttype), and all custom fields (needed to read the display title).
  const commonHeaders: { [key: string]: string } = {
    'Authorization': headers['Authorization'],
    'API-VERSION': '1',
    'select-Record': 'thumbnail,preview,masterfilelatestversion,fields'
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
        // Use File.Version.Extension (documented Aprimo search field).
        // Do NOT use contentType — in Aprimo's search grammar that refers to the
        // admin-configured Content Type name ("Image", "Document"), not the MIME type.
        const typeClause = fileTypes
          .map(ext => `File.Version.Extension = '${ext.toLowerCase()}'`)
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

  try {
    // Single-record endpoint is singular: /api/core/record/{id}
    // (The collection endpoint /api/core/records is plural — different routes)
    const response = UrlFetchApp.fetch(`${base}/record/${id}`, {
      method: 'get',
      headers: {
        'Authorization': headers['Authorization'],
        'API-VERSION': '1',
        'select-Record': 'thumbnail,preview,masterfilelatestversion,fields,classification'
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
 * Fetches the binary download URL for an asset's latest master file version.
 *
 * Uses select-Record header (not $expand) so we avoid OData encoding issues.
 * Returns the URL plus whether it's a pre-signed CDN link (no auth needed)
 * or an API endpoint (auth header required).
 */
function getAssetDownloadUrl(id: string): { url: string; preSignedUrl: boolean } | { error: string } {
  const headers = _getApiHeaders();
  const base = _getApiBase();
  if (!headers || !base) return { error: 'Not authenticated' };

  try {
    // Fetch with masterfile to get the file ID for the canonical download endpoint.
    // Aprimo download: GET /api/core/record/{recordId}/file/{fileId}
    const response = UrlFetchApp.fetch(`${base}/record/${id}`, {
      method: 'get',
      headers: {
        'Authorization': headers['Authorization'],
        'API-VERSION': '1',
        'select-Record': 'thumbnail,preview,masterfile,masterfilelatestversion'
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

    // masterfile gives us the file ID needed for the download endpoint
    // Aprimo returns camelCase in plain JSON (masterFile) and lowercase in HAL+JSON
    // (_embedded.masterfile). Try all variants.
    const masterFile =
      data._embedded?.masterfile ||
      data._embedded?.masterFile ||
      data.masterfile ||
      data.masterFile ||
      {};
    const mf =
      data._embedded?.masterfilelatestversion ||
      data._embedded?.masterFileLatestVersion ||
      data.masterfilelatestversion ||
      data.masterFileLatestVersion ||
      {};

    Logger.log(`getAssetDownloadUrl masterFile.id=${masterFile.id} mf.id=${mf.id}`);
    Logger.log(`getAssetDownloadUrl mf.content=${JSON.stringify(mf.content || '').substring(0, 200)}`);
    Logger.log(`getAssetDownloadUrl mf._links=${JSON.stringify(mf._links || {}).substring(0, 400)}`);
    Logger.log(`getAssetDownloadUrl masterFile._links.self=${masterFile._links?.self?.href}`);

    // ── 1. mf.content — the binary download URI on the latest file version ──────
    // This is the most direct path: Aprimo stores the content URI on the version object.
    // It may be a pre-signed CDN URL or an API URL.
    const contentRaw = mf.content;
    const contentUrl: string =
      (typeof contentRaw === 'string' && contentRaw.startsWith('http') ? contentRaw : '') ||
      (contentRaw?.uri   || '') ||
      (contentRaw?.href  || '') ||
      (contentRaw?.url   || '');
    if (contentUrl) {
      Logger.log(`getAssetDownloadUrl: using mf.content URL: ${contentUrl}`);
      return { url: contentUrl, preSignedUrl: _isPreSignedUrl(contentUrl) };
    }

    // ── 2. OriginalRendition URI via additionalfiles ───────────────────────────
    // Per the Aprimo SDK (FileVersion model): FileVersion has NO direct download link.
    // The correct pattern is to fetch the version's additionalfiles with
    // select-additionalfile: Uri, then use the item where type === 'OriginalRendition'.
    // That item's .uri is a time-limited CDN URL (4 hours) for the original binary.
    const additionalFilesUrl: string = mf._links?.additionalfiles?.href || '';
    if (additionalFilesUrl) {
      Logger.log(`getAssetDownloadUrl: fetching additionalfiles from: ${additionalFilesUrl}`);
      try {
        const afResp = UrlFetchApp.fetch(additionalFilesUrl, {
          method: 'get',
          headers: {
            'Authorization': headers['Authorization'],
            'API-VERSION': '1',
            'select-additionalfile': 'Uri'
          },
          muteHttpExceptions: true
        });

        if (afResp.getResponseCode() === 200) {
          const afData = JSON.parse(afResp.getContentText());
          const afItems: any[] = afData.items || afData._embedded?.items || [];
          Logger.log(`getAssetDownloadUrl: additionalfiles count=${afItems.length}, types=[${afItems.map((f: any) => f.type).join(', ')}]`);

          // Prefer OriginalRendition — the original uploaded binary
          const original = afItems.find((f: any) => f.type === 'OriginalRendition' && f.uri);
          if (original) {
            Logger.log(`getAssetDownloadUrl: using OriginalRendition URI: ${original.uri.substring(0, 80)}…`);
            return { url: original.uri, preSignedUrl: _isPreSignedUrl(original.uri) };
          }

          // Fall back to any additional file that has a URI
          const anyWithUri = afItems.find((f: any) => f.uri);
          if (anyWithUri) {
            Logger.log(`getAssetDownloadUrl: using additionalfile URI (type=${anyWithUri.type}): ${anyWithUri.uri.substring(0, 80)}…`);
            return { url: anyWithUri.uri, preSignedUrl: _isPreSignedUrl(anyWithUri.uri) };
          }

          Logger.log(`getAssetDownloadUrl: additionalfiles returned but no items with URI. types=[${afItems.map((f: any) => f.type).join(', ')}]`);
        } else {
          Logger.log(`getAssetDownloadUrl: additionalfiles returned ${afResp.getResponseCode()}: ${afResp.getContentText().substring(0, 200)}`);
        }
      } catch (afErr: any) {
        Logger.log(`getAssetDownloadUrl: additionalfiles fetch error: ${afErr.message}`);
      }
    }

    // ── 3. Explicit HAL download/content link on the version ──────────────────
    const mfDownload: string =
      mf._links?.download?.href ||
      mf._links?.content?.href  ||
      masterFile._links?.download?.href ||
      '';
    if (mfDownload) {
      Logger.log(`getAssetDownloadUrl: using HAL download link: ${mfDownload}`);
      return { url: mfDownload, preSignedUrl: _isPreSignedUrl(mfDownload) };
    }

    // ── 4. Direct URI fields on version object ─────────────────────────────────
    const directUrl: string =
      mf.downloaduri || mf.downloadUrl || mf.downloadUri || mf.downloadURL ||
      mf.contentUri  || mf.contentUrl  || mf.contenturl  || mf.uri ||
      '';
    if (directUrl) {
      Logger.log(`getAssetDownloadUrl: using direct URI field: ${directUrl}`);
      return { url: directUrl, preSignedUrl: _isPreSignedUrl(directUrl) };
    }

    Logger.log(`getAssetDownloadUrl — no URL found. mf._links keys: [${Object.keys(mf._links || {}).join(', ')}]`);
    return { error: 'No download URL available for this asset. The file may still be processing.' };
  } catch (e: any) {
    return { error: e.message };
  }
}

/** Returns true if the URL is a pre-signed CDN link (no Authorization header needed). */
function _isPreSignedUrl(url: string): boolean {
  return url.includes('X-Amz-') ||
         url.includes('x-amz-') ||
         url.includes('sig=') ||
         url.includes('token=') ||
         !url.includes('/api/');
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
    // $top=500 covers all but the largest taxonomies.
    // The Classification model has:
    //   labelPath  — full localized path e.g. "Marketing / Campaigns / Digital"  (best for users)
    //   namePath   — full internal-name path e.g. "Marketing/Campaigns/Digital"  (fallback)
    //   name       — leaf internal name e.g. "Digital"                            (last resort)
    // We deliberately skip c.labels[] (localized array) to keep mapping simple.
    const response = UrlFetchApp.fetch(
      `${base}/classifications?$top=500`,
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

    const mapped = items
      // Exclude classifications that are hidden in the DAM UI
      .filter((c: any) => !c.disabledInDAMUI)
      .map((c: any) => ({
        id: c.id,
        // labelPath gives the full hierarchy (e.g. "Brand / Assets / Logos") —
        // much more useful in a flat dropdown than just the leaf name.
        label: c.labelPath || c.namePath || c.name || c.id
      }));

    // Sort alphabetically so the dropdown is easy to scan
    mapped.sort((a: any, b: any) => a.label.localeCompare(b.label));

    return { classifications: mapped };
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
      `${base}/collections?$top=100`,
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

    // Collection model uses `name` (not `label`) — sort alphabetically client-side
    const mapped = items.map((c: any) => ({ id: c.id, label: c.name || c.id }));
    mapped.sort((a: any, b: any) => a.label.localeCompare(b.label));
    return { collections: mapped };
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
 * @param existingRecordId  - If set, upload as a new version of this record
 * @param classificationId  - Aprimo classification ID (REQUIRED for new records)
 */
function uploadRecord(
  fileName: string,
  blob: GoogleAppsScript.Base.Blob,
  collectionId: string,
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
      //
      // The SDK's UpdateRecordRequest.files.addOrUpdate requires the ID of the
      // EXISTING file container (not the upload token) so Aprimo knows which
      // file to add the new version to.  Without it the API treats the entry as
      // a brand-new file container, re-pointing `master` at it — a replacement
      // rather than a new version.
      //
      // Correct flow:
      //   1. GET /record/{id} with select-record: masterfile  → masterFile.id
      //   2. POST /uploads                                    → uploadToken
      //   3. PUT /record/{id} with files.addOrUpdate[{id: masterFile.id, versions: ...}]

      // Step 1: fetch the master file container ID
      const recordResp = UrlFetchApp.fetch(`${base}/record/${recordId}`, {
        method: 'get',
        headers: {
          'Authorization': headers['Authorization'],
          'API-VERSION': '1',
          'select-record': 'masterfile'
        },
        muteHttpExceptions: true
      });

      const recordFetchCode = recordResp.getResponseCode();
      if (recordFetchCode !== 200) {
        Logger.log(`Fetch record for versioning failed (${recordFetchCode}): ${recordResp.getContentText()}`);
        return { success: false, error: `Could not fetch record to prepare new version (${recordFetchCode})` };
      }

      const recordData = JSON.parse(recordResp.getContentText());
      // Plain JSON (no Accept: application/hal+json) uses camelCase for direct properties.
      // The select-record header embeds sub-resources under _embedded with lowercase keys.
      const masterFileContainer =
        recordData._embedded?.masterfile ||
        recordData._embedded?.masterFile ||
        recordData.masterfile ||
        recordData.masterFile ||
        {};
      const masterFileContainerId: string = masterFileContainer.id || '';

      if (!masterFileContainerId) {
        Logger.log(`uploadRecord: no master file container found on record ${recordId}. Data: ${JSON.stringify(recordData).substring(0, 400)}`);
        return { success: false, error: 'Could not find the master file container on this record. Is it a file-based asset?' };
      }

      // Step 2: stage the file
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

      // Step 3: add the new version to the existing file container
      const versionResp = UrlFetchApp.fetch(`${base}/record/${recordId}`, {
        method: 'put',
        headers,
        payload: JSON.stringify({
          files: {
            // `master` tells Aprimo which upload token becomes the new active master
            master: uploadToken,
            addOrUpdate: [{
              // `id` here is the EXISTING file container — required by the API to
              // add a version rather than replace the entire master file.
              id: masterFileContainerId,
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

  // ── masterfilelatestversion — embedded when select-Record includes it ────────
  // Aprimo returns camelCase in plain JSON, lowercase in HAL+JSON — try all.
  const mf =
    item._embedded?.masterfilelatestversion ||
    item._embedded?.masterFileLatestVersion ||
    item.masterfilelatestversion ||
    item.masterFileLatestVersion ||
    item._embedded?.latestVersion ||
    item.latestVersion ||
    {};

  // ── Title ─────────────────────────────────────────────────────────────────
  // 1. Try item.title (localized { value } object or plain string).
  // 2. If that's empty, scan custom fields (_embedded.fields.items) for a
  //    field whose label is "title", "name", or similar — Aprimo tenants
  //    sometimes store the display name as a custom field rather than title.
  // 3. Fall back to a short truncated GUID only as a last resort.
  const rawTitle = item.title;
  let title: string = typeof rawTitle === 'string'
    ? rawTitle.trim()
    : ((rawTitle as any)?.value || '').trim();

  if (!title) {
    // Scan embedded fields for a display name
    const fieldItems: any[] = (item._embedded?.fields?.items) || [];
    // First pass: exact-match on common title field labels
    for (const f of fieldItems) {
      const label = (f.label || f.fieldLabel || f.dataTypeName || '').toLowerCase().trim();
      if (label === 'title' || label === 'name' || label === 'asset title' || label === 'asset name' || label === 'file name') {
        const val = (f.localizedValues?.[0]?.value || f.value || f.stringValue || '').trim();
        if (val) { title = val; break; }
      }
    }
    // Second pass: first non-empty single-line text field
    if (!title) {
      for (const f of fieldItems) {
        const dataType = (f.dataType || '').toLowerCase();
        if (dataType === 'singlelinetext' || dataType === 'multilinetext') {
          const val = (f.localizedValues?.[0]?.value || f.value || f.stringValue || '').trim();
          if (val) { title = val; break; }
        }
      }
    }
  }

  // Absolute fallback — use a short prefix of the GUID so it's still unique
  // but clearly a system ID, not a meaningful title
  if (!title) title = `Asset ${item.id.substring(0, 8)}…`;

  // ── File extension ─────────────────────────────────────────────────────────
  // Prefer masterfilelatestversion — embedded when select-Record includes it.
  // The thumbnail/preview extension is the PREVIEW IMAGE type (usually jpeg),
  // NOT the source file type, so we must NOT use it as the primary source.
  const thumbExt   = (thumb.extension   || '').toLowerCase();
  const previewExt = (preview.extension || '').toLowerCase();
  const mfExt = (mf.fileextension || mf.fileExtension || '').toLowerCase();
  // Only fall back to thumbnail/preview ext if we have nothing from masterfile
  // and the thumb ext is not a generic image format masking a real doc
  const thumbExtIsImage = /^(jpg|jpeg|png|gif|webp|tiff|tif|svg)$/.test(thumbExt);
  const fileExtension: string = mfExt
    || (thumbExtIsImage ? '' : thumbExt)
    || (thumbExtIsImage ? '' : previewExt)
    || '';

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

  // ── Classification ────────────────────────────────────────────────────────
  const classificationLabel: string =
    item._embedded?.classification?.label  ||
    item.classification?.label             ||
    item._embedded?.classifications?.items?.[0]?.label ||
    '';

  // ── Custom fields for the detail panel ────────────────────────────────────
  // Scan all embedded field items; keep those with a non-empty display value.
  const fieldItems: any[] = (item._embedded?.fields?.items) || [];
  const fields: AssetField[] = [];
  for (const f of fieldItems) {
    const label = (f.label || f.fieldLabel || f.dataTypeName || '').trim();
    if (!label) continue;
    // Try localizedValues first (multi-locale), then scalar value fallbacks
    const localVals: any[] = f.localizedValues || [];
    const rawVal =
      (localVals.length > 0 ? localVals.map((lv: any) => (lv.value || '').trim()).filter(Boolean).join(', ') : '') ||
      (f.value !== undefined && f.value !== null ? String(f.value) : '') ||
      (f.stringValue || '') ||
      (f.booleanValue !== undefined ? String(f.booleanValue) : '');
    const val = rawVal.trim();
    if (val) fields.push({ label, value: val });
    if (fields.length >= 30) break;  // cap at 30 to keep payload reasonable
  }

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
    classificationLabel,
    modifiedOn: item.modifiedon || item.modifiedOn || '',
    createdOn:  item.createdon  || item.createdOn  || '',
    status:     item.status     || item.statusLabel || '',
    fields
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
