# Aprimo DAM — Google Workspace Add-on

Search, insert, upload, and version Aprimo DAM assets directly inside Google Docs, Sheets, and Slides.

## Features

| Use Case | How it works |
|---|---|
| **Search & Insert Images** | Search your DAM, click Insert → image appears at cursor in Docs or centered on slide in Slides |
| **Open & Edit Office Files** | Download a DOCX/XLSX/PPTX from Aprimo, edit it as a Google file, save back as a new version |
| **Upload Current Document** | Export the open Google file to its Office format and upload to Aprimo (new asset or new version) |
| **Recent Assets** | Last 10 assets you touched are one click away in the Recent tab |

## Setup

### 1. Install CLASP

```bash
npm install
npm install -g @google/clasp
clasp login
```

### 2. Create the Apps Script project

```bash
cd src
clasp create --type standalone --title "Aprimo DAM"
```

Copy the `scriptId` from the output into `.clasp.json`.

### 3. Push the code

```bash
clasp push
```

### 4. Deploy as a Web App

This is required for the OAuth redirect URI.

1. Open the script: `clasp open`
2. **Deploy → New deployment**
3. Type: **Web app**
4. Execute as: **User accessing the web app**
5. Who has access: **Anyone** (anonymous is fine — no data is returned, it only exchanges the OAuth code)
6. Click **Deploy** and copy the Web App URL

Format: `https://script.google.com/macros/s/{SCRIPT_ID}/exec`

### 5. Register an OAuth Client in Aprimo

1. In your Aprimo tenant, go to **Settings → Integration → OAuth Clients**
2. Create a new client:
   - Name: `Google Workspace Add-on`
   - Grant type: **Authorization Code with PKCE**
   - Redirect URI: paste the Web App URL from step 4
   - Scopes: `api`
3. Copy the **Client ID**

### 6. Connect in Google Docs / Sheets / Slides

1. Open any Google Doc, Sheet, or Slides
2. Go to **Extensions → Aprimo DAM → Open Aprimo DAM**
3. Enter your tenant subdomain (e.g. `mycompany`) and the Client ID
4. Click **Connect** — a popup will open for Aprimo login
5. After logging in, the popup closes and the sidebar shows your connected state

---

## Development

```bash
# Push changes and watch for file changes
npm run push:watch

# View logs from the last execution
npm run logs

# Open the script editor in the browser
npm run open
```

## File Structure

```
src/
├── appsscript.json   Manifest: OAuth scopes, web app config
├── Code.ts           Entry: onOpen(), openSidebar(), getActiveDocumentInfo()
├── Auth.ts           PKCE state, token exchange, getAuthStatus(), _getApiHeaders()
├── Api.ts            searchAssets(), getAsset(), uploadRecord(), getCollections()
├── Insert.ts         insertImage(), openAssetForEditing(), saveEditAsNewVersion()
├── Upload.ts         uploadActiveDocument() — exports current doc and uploads
├── Webapp.ts         doGet() — OAuth callback redirect handler
├── sidebar.html      Full sidebar UI (Search / Upload / Recent tabs)
└── callback.html     OAuth callback page shown in the auth popup
```

## Auth Flow (PKCE)

```
Sidebar (browser)                    Apps Script server          Aprimo
─────────────────                    ──────────────────          ──────
generatePKCE()                       
  → savePKCEState(verifier)    →     UserProperties.set()
  → open popup with auth URL   →     (popup navigates to Aprimo)
                                                             ←   user logs in
                               ←     doGet(?code=...)       ←   redirect
                                     exchangeToken(code)
                                       UrlFetchApp.fetch(token endpoint)
                                       UserProperties.set(token)
sidebar polls getAuthStatus()
  ← connected: true
showApp()
```

## Aprimo API Notes

- Base URL: `https://{tenant}.dam.aprimo.com/api/core`
- Required header: `API-VERSION: 1`
- **Search (text/filter):** `POST /api/core/search/records?page=N&pageSize=N` with body `{ searchExpression: { expression: "query" } }`
- **Listing (no query):** `GET /api/core/records` — pagination via `take`/`skip`/`orderby` **HTTP headers** (not query params)
- **Thumbnails/previews:** add header `select-Record: thumbnail,preview` — Aprimo returns pre-signed CDN URIs directly at `record._embedded.thumbnail.uri` and `record._embedded.preview.uri`
- **Upload (2-step):** `POST https://{tenant}.aprimo.com/uploads` (multipart) → get `token`; then `POST /api/core/records` with `{ title: { value }, classifications: { addOrUpdate: [{ id }] }, files: { master: token, addOrUpdate: [...] } }`
- Title fields are returned as `{ value: string }` objects (localized), not plain strings

## Deployment to Google Workspace Marketplace

1. Set up a Google Cloud project and OAuth consent screen
2. Link it to your Apps Script project in **Project Settings**
3. Publish via the Apps Script dashboard under **Deploy → Publish to Marketplace**
4. For internal rollout: share the add-on URL or use **Domain-wide install** via Google Workspace Admin

## Troubleshooting

| Issue | Fix |
|---|---|
| Popup blocked | Allow popups in the browser for `script.google.com` |
| "No PKCE state found" | The popup took >10 min; click Connect again |
| Token exchange fails | Verify the redirect URI in Aprimo exactly matches the Web App URL |
| Insert does nothing | Make sure your cursor is placed in the document before clicking Insert |
| Drive API not enabled | Enable **Drive API** in the linked Google Cloud project |
