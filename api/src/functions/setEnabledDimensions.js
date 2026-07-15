const { app } = require('@azure/functions');
const { ConfidentialClientApplication } = require('@azure/msal-node');

const TENANT_ID = process.env.ATLAS_TENANT_ID;
const CLIENT_ID = process.env.ATLAS_CLIENT_ID;
const CLIENT_SECRET = process.env.ATLAS_CLIENT_SECRET;
const SITE_ID = process.env.ATLAS_SITE_ID;
const DATA_FOLDER = 'AtlasData';

// Same authorisation model as the other write endpoints.
const ALLOWED_EDITOR_EMAIL = (process.env.ATLAS_ALLOWED_EDITOR_EMAIL || '').toLowerCase();

// Kept in sync with SCORECARD_DIMENSIONS in scorecard-dimensions.js /
// exportHelpers.js -- an explicit allow-list of dimension keys, same reason
// as VALID_DIMENSIONS in setScorecardGlobal.js.
const VALID_DIMENSIONS = [
  'market_opportunity', 'outsourced_management', 'pricing_impact',
  'alignment_of_investment_thinking', 'distribution_resources_required',
  'regulatory_complexity', 'client_servicing', 'local_presence_required',
  'languages_required', 'investor_decision_making', 'comingled_vehicles',
  'consultant_reliant'
];

let msalClient;
function getMsalClient() {
  if (!msalClient) {
    msalClient = new ConfidentialClientApplication({
      auth: {
        clientId: CLIENT_ID,
        authority: `https://login.microsoftonline.com/${TENANT_ID}`,
        clientSecret: CLIENT_SECRET
      }
    });
  }
  return msalClient;
}

async function getAccessToken() {
  const client = getMsalClient();
  const result = await client.acquireTokenByClientCredential({
    scopes: ['https://graph.microsoft.com/.default']
  });
  return result.accessToken;
}

function getClientPrincipal(request) {
  const header = request.headers.get('x-ms-client-principal');
  if (!header) return null;
  try {
    return JSON.parse(Buffer.from(header, 'base64').toString('utf-8'));
  } catch {
    return null;
  }
}

// Saves the global "which scorecard dimensions count toward Overall" toggle
// (Peter's 2026-07-15 request) onto global.json's top-level
// enabled_dimensions field -- a plain {dimensionKey: boolean} map, missing
// keys defaulting to enabled (see isDimensionEnabled() in
// scorecard-dimensions.js / exportHelpers.js). Global and persisted, not a
// per-session filter: every reader of overview.html/country.html/picker.html
// and every Word/PPT export sees the same enabled/disabled set once saved,
// same as the *_basis reminders setScorecardBasis.js writes. Replaces the
// whole map on each save (not a per-key patch) since the client always
// sends the full 12-entry state from its checkboxes.
app.http('setEnabledDimensions', {
  methods: ['POST'],
  authLevel: 'anonymous',
  route: 'global/set-enabled-dimensions',
  handler: async (request, context) => {
    const principal = getClientPrincipal(request);
    const userEmail = ((principal && principal.userDetails) || '').toLowerCase();

    if (!ALLOWED_EDITOR_EMAIL) {
      return { status: 500, jsonBody: { error: 'Server is not configured with an authorised editor email.' } };
    }
    if (!principal || userEmail !== ALLOWED_EDITOR_EMAIL) {
      return { status: 403, jsonBody: { error: 'Not authorised to make changes.' } };
    }

    let body;
    try {
      body = await request.json();
    } catch {
      return { status: 400, jsonBody: { error: 'Invalid request body' } };
    }

    const { enabled_dimensions } = body || {};
    if (!enabled_dimensions || typeof enabled_dimensions !== 'object' || Array.isArray(enabled_dimensions)) {
      return { status: 400, jsonBody: { error: 'enabled_dimensions object is required' } };
    }
    for (const key of Object.keys(enabled_dimensions)) {
      if (!VALID_DIMENSIONS.includes(key)) {
        return { status: 400, jsonBody: { error: `Unknown scorecard dimension: ${key}` } };
      }
      if (typeof enabled_dimensions[key] !== 'boolean') {
        return { status: 400, jsonBody: { error: `enabled_dimensions.${key} must be a boolean` } };
      }
    }

    const fileName = 'global.json';
    const graphBase = `https://graph.microsoft.com/v1.0/sites/${SITE_ID}/drive/root:/${DATA_FOLDER}/${fileName}`;

    try {
      const token = await getAccessToken();

      const getResp = await fetch(`${graphBase}:/content`, {
        headers: { Authorization: `Bearer ${token}` }
      });
      if (getResp.status === 404) {
        return { status: 404, jsonBody: { error: 'No global.json found on SharePoint yet.' } };
      }
      if (!getResp.ok) {
        return { status: 502, jsonBody: { error: `Could not read global.json from SharePoint`, status: getResp.status } };
      }
      const data = await getResp.json();

      data.enabled_dimensions = enabled_dimensions;
      data.enabled_dimensions_updated_date = new Date().toISOString().slice(0, 10);
      data.enabled_dimensions_updated_by = userEmail;

      const putResp = await fetch(`${graphBase}:/content`, {
        method: 'PUT',
        headers: {
          Authorization: `Bearer ${token}`,
          'Content-Type': 'application/json'
        },
        body: JSON.stringify(data, null, 2)
      });

      if (!putResp.ok) {
        const detail = await putResp.text();
        context.error(`Graph PUT error ${putResp.status}: ${detail}`);
        return { status: 502, jsonBody: { error: 'Could not save changes to SharePoint', status: putResp.status } };
      }

      return {
        status: 200,
        jsonBody: {
          success: true,
          enabled_dimensions: data.enabled_dimensions,
          updated_date: data.enabled_dimensions_updated_date,
          updated_by: data.enabled_dimensions_updated_by
        }
      };
    } catch (err) {
      context.error(err);
      return { status: 500, jsonBody: { error: 'Server error updating global.json', detail: err.message } };
    }
  }
});
