// =============================================================================
// Auth.ts — PKCE OAuth 2.0 flow for Aprimo
//
// Flow:
//   1. Client (sidebar HTML) generates code_verifier + code_challenge
//   2. Client calls savePKCEState() to persist verifier server-side
//   3. Client opens Aprimo auth URL in a popup window
//   4. Aprimo redirects to doGet() (Webapp.ts) with ?code=...
//   5. doGet() calls exchangeToken() — server exchanges code for token
//   6. Token stored in UserProperties; popup shows success and closes
//   7. Sidebar polls getAuthStatus() until connected
// =============================================================================

interface PKCEState {
  verifier: string;
  tenant: string;
  clientId: string;
  clientSecret: string;
  timestamp: number;
}

interface StoredToken {
  accessToken: string;
  tokenType: string;
  expiresAt: number;
  tenant: string;
  clientId: string;
}

interface AuthStatus {
  connected: boolean;
  email?: string;
  tenant?: string;
  expiresAt?: number;
}

// ─── PKCE Generation + Auth URL (server-side) ────────────────────────────────

/**
 * Generates a PKCE pair server-side, saves the verifier, and returns the
 * full Aprimo authorization URL ready for the sidebar to open in a popup.
 *
 * Doing this server-side avoids crypto.subtle / async-await issues in the
 * Apps Script HTML sandbox.
 */
function buildAuthUrl(tenant: string, clientId: string, clientSecret?: string): { authUrl: string } | { error: string } {
  try {
    const cleanTenant = tenant.trim().toLowerCase();
    const cleanClientId = clientId.trim();

    const webAppUrl = PropertiesService.getScriptProperties()
      .getProperty('aprimo_webapp_url') || '';

    if (!webAppUrl) {
      return { error: 'Web app URL not configured. Please run _runSetup() from the Apps Script editor.' };
    }

    // Generate a 32-byte random code verifier using Apps Script Utilities
    const randomBytes: number[] = [];
    for (let i = 0; i < 32; i++) {
      randomBytes.push(Math.floor(Math.random() * 256));
    }
    const verifier = Utilities.base64EncodeWebSafe(randomBytes).replace(/=+$/, '');

    // SHA-256 hash → base64url encode = code_challenge
    const challengeBytes = Utilities.computeDigest(
      Utilities.DigestAlgorithm.SHA_256,
      verifier,
      Utilities.Charset.UTF_8
    );
    const challenge = Utilities.base64EncodeWebSafe(challengeBytes).replace(/=+$/, '');

    // Persist the verifier for the token exchange in doGet()
    const state: PKCEState = {
      verifier,
      tenant: cleanTenant,
      clientId: cleanClientId,
      clientSecret: (clientSecret || '').trim(),
      timestamp: Date.now()
    };
    PropertiesService.getUserProperties()
      .setProperty('aprimo_pkce_state', JSON.stringify(state));

    // Build the full authorization URL
    const params = [
      'response_type=code',
      `client_id=${encodeURIComponent(cleanClientId)}`,
      `redirect_uri=${encodeURIComponent(webAppUrl)}`,
      'scope=api',
      `code_challenge=${encodeURIComponent(challenge)}`,
      'code_challenge_method=S256',
      `state=${encodeURIComponent(Utilities.getUuid())}`
    ].join('&');

    return { authUrl: `https://${cleanTenant}.aprimo.com/login/connect/authorize?${params}` };
  } catch (e: any) {
    Logger.log(`buildAuthUrl error: ${e.message}`);
    return { error: e.message };
  }
}

// ─── Token Exchange ───────────────────────────────────────────────────────────

/**
 * Exchanges an authorization code for an access token.
 * Called from doGet() after Aprimo redirects with ?code=...
 *
 * @param code        - Authorization code from Aprimo redirect
 * @param redirectUri - Must match exactly what was sent in the auth request
 */
function exchangeToken(code: string, redirectUri: string): { success: boolean; error?: string } {
  const props = PropertiesService.getUserProperties();

  const stateJson = props.getProperty('aprimo_pkce_state');
  if (!stateJson) {
    return { success: false, error: 'No PKCE state found. Please start the connection again.' };
  }

  const state: PKCEState = JSON.parse(stateJson);

  // Reject stale PKCE states (10 minute window)
  if (Date.now() - state.timestamp > 10 * 60 * 1000) {
    props.deleteProperty('aprimo_pkce_state');
    return { success: false, error: 'Authorization session expired. Please reconnect.' };
  }

  const tokenUrl = `https://${state.tenant}.aprimo.com/login/connect/token`;

  const formParamParts = [
    `grant_type=authorization_code`,
    `client_id=${encodeURIComponent(state.clientId)}`,
    `code=${encodeURIComponent(code)}`,
    `code_verifier=${encodeURIComponent(state.verifier)}`,
    `redirect_uri=${encodeURIComponent(redirectUri)}`
  ];
  // Include client_secret if one was registered on the Aprimo OAuth client
  if (state.clientSecret) {
    formParamParts.push(`client_secret=${encodeURIComponent(state.clientSecret)}`);
  }
  const formParams = formParamParts.join('&');

  try {
    const response = UrlFetchApp.fetch(tokenUrl, {
      method: 'post',
      contentType: 'application/x-www-form-urlencoded',
      payload: formParams,
      muteHttpExceptions: true
    });

    if (response.getResponseCode() !== 200) {
      const body = response.getContentText();
      Logger.log(`Token exchange failed (${response.getResponseCode()}): ${body}`);
      return { success: false, error: `Token exchange failed (${response.getResponseCode()}): ${body}` };
    }

    const tokenData = JSON.parse(response.getContentText());

    const token: StoredToken = {
      accessToken: tokenData.access_token,
      tokenType: tokenData.token_type || 'Bearer',
      expiresAt: Date.now() + ((tokenData.expires_in || 3600) * 1000),
      tenant: state.tenant,
      clientId: state.clientId
    };

    props.setProperty('aprimo_token', JSON.stringify(token));
    props.deleteProperty('aprimo_pkce_state');

    return { success: true };
  } catch (e: any) {
    Logger.log(`Token exchange error: ${e.message}`);
    return { success: false, error: e.message };
  }
}

// ─── Auth Status & Token Access ───────────────────────────────────────────────

/**
 * Returns the current auth state. Called by the sidebar to decide which
 * screen to show (connect vs. main app).
 */
function getAuthStatus(): AuthStatus {
  const token = _getValidToken();
  if (!token) return { connected: false };

  let email = '';
  try { email = Session.getActiveUser().getEmail(); } catch (e) {}

  return {
    connected: true,
    email,
    tenant: token.tenant,
    expiresAt: token.expiresAt
  };
}

/**
 * Returns the current tenant subdomain (e.g. "mycompany"), or null if not connected.
 */
function getCurrentTenant(): string | null {
  const token = _getValidToken();
  return token ? token.tenant : null;
}

/**
 * Clears all stored auth data and re-opens the sidebar to the connect screen.
 */
function clearAuth(): void {
  const props = PropertiesService.getUserProperties();
  props.deleteProperty('aprimo_token');
  props.deleteProperty('aprimo_pkce_state');
  props.deleteProperty('aprimo_current_edit');
  props.deleteProperty('aprimo_recent_assets');
  openSidebar();
}

/**
 * Returns a valid (non-expired) stored token, or null if not authenticated.
 * Internal helper used by Api.ts and Insert.ts.
 */
function _getValidToken(): StoredToken | null {
  const props = PropertiesService.getUserProperties();
  const tokenJson = props.getProperty('aprimo_token');
  if (!tokenJson) return null;

  const token: StoredToken = JSON.parse(tokenJson);

  // Reject tokens that expire within the next 5 minutes
  if (Date.now() > token.expiresAt - 5 * 60 * 1000) return null;

  return token;
}

/**
 * Returns Authorization + API-VERSION headers for Aprimo DAM REST API calls.
 * Returns null if not authenticated.
 */
function _getApiHeaders(): { [key: string]: string } | null {
  const token = _getValidToken();
  if (!token) return null;
  return {
    'Authorization': `${token.tokenType} ${token.accessToken}`,
    'API-VERSION': '1',
    'Content-Type': 'application/json'
  };
}

/**
 * Returns the Aprimo DAM API base URL for the current tenant.
 * e.g. https://mycompany.dam.aprimo.com/api/core
 */
function _getApiBase(): string | null {
  const token = _getValidToken();
  if (!token) return null;
  return `https://${token.tenant}.dam.aprimo.com/api/core`;
}

/**
 * Returns the deployed web app URL — used as the OAuth redirect_uri.
 * Stored in ScriptProperties during setup so it's available from any context
 * (ScriptApp.getService().getUrl() only works inside doGet()).
 */
function getWebAppUrl(): string {
  return PropertiesService.getScriptProperties()
    .getProperty('aprimo_webapp_url') || '';
}

/**
 * Stores the web app URL in ScriptProperties. Call this once during setup.
 */
function setWebAppUrl(url: string): void {
  PropertiesService.getScriptProperties().setProperty('aprimo_webapp_url', url);
  Logger.log('✅ Web app URL saved: ' + url);
}
