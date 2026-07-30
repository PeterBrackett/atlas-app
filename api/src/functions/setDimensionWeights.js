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
// exportHelpers.js, same as setEnabledDimensions.js's VALID_DIMENSIONS.
const VALID_DIMENSIONS = [
  'market_opportunity', 'outsourced_management', 'pricing_impact',
  'alignment_of_investment_thinking', 'distribution_resources_required',
  'regulatory_complexity', 'client_servicing', 'local_presence_required',
  'languages_required', 'investor_decision_making', 'comingled_vehicles',
  'consultant_reliant'
];

// See setEnabledDimensions.js's comment -- client-defined custom factors use
// a `custom_` prefixed key, accepted here alongside the fixed 12 so a pushed
// weighting for a custom factor isn't rejected.
const CUSTOM_DIMENSION_KEY_RE = /^custom_[a-z0-9_]{1,50}$/;
function isValidDimensionKey(key) {
  return VALID_DIMENSIONS.includes(key) || CUSTOM_DIMENSION_KEY_RE.test(key);
}

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

// Saves a global override of each scorecard dimension's weight (used in the
// Overall = sum(score x weight) formula) onto global.json's top-level
// dimension_weights field -- a plain {dimensionKey: number} map, missing
// keys defaulting to that dimension's own hardcoded weight (see
// applyGlobalDimensionWeights() in scorecard-dimensions.js/exportHelpers.js).
// Added 2026-07-30 from weighting.html's "Push to scorecard" button: Peter
// wanted a client's own priorities (explored live via that page's sliders)
// to be able to become the standard weighting everyone sees from then on --
// same "global and persisted, not a personal filter" model as
// setEnabledDimensions.js, and the same one-shot "replace the whole map"
// write (not a per-key patch), since the client always sends the full
// current 12-entry state from its sliders.
app.http('setDimensionWeights', {
  methods: ['POST'],
  authLevel: 'anonymous',
  route: 'global/set-dimension-weights',
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

    const { dimension_weights } = body || {};
    if (!dimension_weights || typeof dimension_weights !== 'object' || Array.isArray(dimension_weights)) {
      return { status: 400, jsonBody: { error: 'dimension_weights object is required' } };
    }
    for (const key of Object.keys(dimension_weights)) {
      if (!isValidDimensionKey(key)) {
        return { status: 400, jsonBody: { error: `Unknown scorecard dimension: ${key}` } };
      }
      if (typeof dimension_weights[key] !== 'number' || !Number.isFinite(dimension_weights[key])) {
        return { status: 400, jsonBody: { error: `dimension_weights.${key} must be a number` } };
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

      data.dimension_weights = dimension_weights;
      data.dimension_weights_updated_date = new Date().toISOString().slice(0, 10);
      data.dimension_weights_updated_by = userEmail;

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
          dimension_weights: data.dimension_weights,
          updated_date: data.dimension_weights_updated_date,
          updated_by: data.dimension_weights_updated_by
        }
      };
    } catch (err) {
      context.error(err);
      return { status: 500, jsonBody: { error: 'Server error updating global.json', detail: err.message } };
    }
  }
});
