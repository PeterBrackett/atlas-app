const { app } = require('@azure/functions');
const { ConfidentialClientApplication } = require('@azure/msal-node');

const TENANT_ID = process.env.ATLAS_TENANT_ID;
const CLIENT_ID = process.env.ATLAS_CLIENT_ID;
const CLIENT_SECRET = process.env.ATLAS_CLIENT_SECRET;
const SITE_ID = process.env.ATLAS_SITE_ID;
const DATA_FOLDER = 'AtlasData';

// Same authorisation model as the other write endpoints.
const ALLOWED_EDITOR_EMAIL = (process.env.ATLAS_ALLOWED_EDITOR_EMAIL || '').toLowerCase();

// Which top-level global.json fields this endpoint is allowed to write --
// an explicit allow-list rather than accepting any key name, so a client
// bug can't stomp on an unrelated part of global.json (e.g. `countries`).
const ALLOWED_BASIS_KEYS = ['market_opportunity_basis', 'regulatory_complexity_basis', 'distribution_resources_basis', 'languages_required_basis', 'local_presence_required_basis', 'client_servicing_basis'];

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

// Generalised version of the original setMarketOpportunityBasis.js --
// records what a bulk auto-score run (setScorecardGlobal.js) was actually
// based on onto global.json itself, keyed by an allow-listed basis field
// name, so the front page can show a "here's what this was based on"
// reminder for more than one dimension without a near-duplicate endpoint
// per dimension. First used for Market opportunity (asset allocation
// thresholds); Regulatory complexity (TMF rank thresholds) is the second.
// Supersedes setMarketOpportunityBasis.js -- delete that file if it was
// already deployed, this one replaces it (same global.json field name for
// market_opportunity_basis, so nothing already saved needs migrating).
app.http('setScorecardBasis', {
  methods: ['POST'],
  authLevel: 'anonymous',
  route: 'global/set-scorecard-basis',
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

    const { basis_key, basis } = body || {};
    if (!basis_key || !ALLOWED_BASIS_KEYS.includes(basis_key)) {
      return { status: 400, jsonBody: { error: `Unknown or missing basis_key. Allowed: ${ALLOWED_BASIS_KEYS.join(', ')}` } };
    }
    if (!basis || typeof basis !== 'object') {
      return { status: 400, jsonBody: { error: 'basis object is required' } };
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

      data[basis_key] = {
        ...basis,
        applied_date: new Date().toISOString().slice(0, 10),
        applied_by: userEmail
      };

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

      return { status: 200, jsonBody: { success: true, basis_key, basis: data[basis_key] } };
    } catch (err) {
      context.error(err);
      return { status: 500, jsonBody: { error: 'Server error updating global.json', detail: err.message } };
    }
  }
});
