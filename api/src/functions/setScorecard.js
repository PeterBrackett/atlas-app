const { app } = require('@azure/functions');
const { ConfidentialClientApplication } = require('@azure/msal-node');

const TENANT_ID = process.env.ATLAS_TENANT_ID;
const CLIENT_ID = process.env.ATLAS_CLIENT_ID;
const CLIENT_SECRET = process.env.ATLAS_CLIENT_SECRET;
const SITE_ID = process.env.ATLAS_SITE_ID;
const DATA_FOLDER = 'AtlasData';

// Same authorisation model as setActiveSource.js: Azure Static Web Apps'
// built-in Entra ID login lets any Microsoft account sign in, so only a
// request carrying this exact email is allowed to actually save a change.
const ALLOWED_EDITOR_EMAIL = (process.env.ATLAS_ALLOWED_EDITOR_EMAIL || '').toLowerCase();

// Kept in sync with scorecard-dimensions.js on the front end — a request
// naming any dimension outside this list is rejected rather than silently
// written into the country JSON as an unrecognised field.
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

// 2026-07-24: writes to {code}_scores.json (via getScores.js), not
// {code}.json -- scores used to live inline on each segment inside the main
// country file, which meant any data-refresh/bug-fix script rewriting that
// file (even carefully) put every score at risk of being lost to a stale
// read-modify-write race. Splitting scores into their own file means nothing
// that touches {code}.json can ever affect them. See getScores.js for the
// full rationale.
app.http('setScorecard', {
  methods: ['POST'],
  authLevel: 'anonymous',
  route: 'segments/set-scorecard',
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

    const { country, segment, dimension, value } = body || {};
    if (!country || !segment || !dimension) {
      return { status: 400, jsonBody: { error: 'country, segment, and dimension are required' } };
    }
    if (!VALID_DIMENSIONS.includes(dimension)) {
      return { status: 400, jsonBody: { error: `Unknown scorecard dimension: ${dimension}` } };
    }
    // value is either a score of 1-3, or null to clear a dimension back to "not yet scored".
    if (value !== null && ![1, 2, 3].includes(value)) {
      return { status: 400, jsonBody: { error: 'value must be 1, 2, 3, or null' } };
    }

    const countryCode = String(country).toLowerCase();
    if (!/^[a-z0-9-]+$/.test(countryCode)) {
      return { status: 400, jsonBody: { error: 'Invalid country identifier' } };
    }

    const fileName = `${countryCode}_scores.json`;
    const graphBase = `https://graph.microsoft.com/v1.0/sites/${SITE_ID}/drive/root:/${DATA_FOLDER}/${fileName}`;

    try {
      const token = await getAccessToken();

      const getResp = await fetch(`${graphBase}:/content`, {
        headers: { Authorization: `Bearer ${token}` }
      });
      // No scores file yet for this country -- start from an empty one
      // rather than treating that as an error, same "create on first write"
      // behaviour setCommentary.js relies on for a missing commentary key.
      const data = getResp.status === 404
        ? { country_code: countryCode, scores: {} }
        : (getResp.ok ? await getResp.json() : null);

      if (data === null) {
        return { status: 502, jsonBody: { error: `Could not read ${fileName} from SharePoint`, status: getResp.status } };
      }
      if (!data.scores || typeof data.scores !== 'object') data.scores = {};

      if (!data.scores[segment]) data.scores[segment] = {};
      if (value === null) {
        delete data.scores[segment][dimension];
      } else {
        data.scores[segment][dimension] = value;
        data.scores[segment].scored_date = new Date().toISOString().slice(0, 10);
      }

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

      return { status: 200, jsonBody: { success: true, segment, scorecard: data.scores[segment] } };
    } catch (err) {
      context.error(err);
      return { status: 500, jsonBody: { error: 'Server error updating scorecard', detail: err.message } };
    }
  }
});
