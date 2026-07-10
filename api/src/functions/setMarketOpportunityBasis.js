const { app } = require('@azure/functions');
const { ConfidentialClientApplication } = require('@azure/msal-node');

const TENANT_ID = process.env.ATLAS_TENANT_ID;
const CLIENT_ID = process.env.ATLAS_CLIENT_ID;
const CLIENT_SECRET = process.env.ATLAS_CLIENT_SECRET;
const SITE_ID = process.env.ATLAS_SITE_ID;
const DATA_FOLDER = 'AtlasData';

// Same authorisation model as the other write endpoints.
const ALLOWED_EDITOR_EMAIL = (process.env.ATLAS_ALLOWED_EDITOR_EMAIL || '').toLowerCase();

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

// Records what the last "auto-score Market opportunity from an allocation
// threshold" run (setScorecardGlobal.js) was actually based on -- asset
// class/style, threshold, optimal, who ran it and when, and how many
// countries/segments it touched -- onto global.json itself, so the basis
// is visible on the overview page's front page rather than only living in
// browser state that resets on reload. Requested specifically because the
// signed-in write actions live behind auth, and it's easy to apply a set of
// thresholds, sign out, and later forget exactly what basis the scores on
// screen are resting on.
app.http('setMarketOpportunityBasis', {
  methods: ['POST'],
  authLevel: 'anonymous',
  route: 'global/set-market-opportunity-basis',
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

    const { asset_type, asset_style, threshold_bn, optimal_bn, countries_updated, segments_updated } = body || {};
    if (!asset_type || typeof threshold_bn !== 'number' || typeof optimal_bn !== 'number') {
      return { status: 400, jsonBody: { error: 'asset_type, threshold_bn, and optimal_bn are required' } };
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

      data.market_opportunity_basis = {
        asset_type,
        asset_style: asset_style || null,
        threshold_bn,
        optimal_bn,
        countries_updated: countries_updated || null,
        segments_updated: segments_updated || null,
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

      return { status: 200, jsonBody: { success: true, market_opportunity_basis: data.market_opportunity_basis } };
    } catch (err) {
      context.error(err);
      return { status: 500, jsonBody: { error: 'Server error updating global.json', detail: err.message } };
    }
  }
});
