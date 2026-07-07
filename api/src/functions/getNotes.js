const { app } = require('@azure/functions');
const { ConfidentialClientApplication } = require('@azure/msal-node');

// Same credentials/site as getData.js — kept as separate constants here so
// this function has no dependency on the other one.
const TENANT_ID = process.env.ATLAS_TENANT_ID;
const CLIENT_ID = process.env.ATLAS_CLIENT_ID;
const CLIENT_SECRET = process.env.ATLAS_CLIENT_SECRET;
const SITE_ID = process.env.ATLAS_SITE_ID;
const DATA_FOLDER = 'AtlasData';

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

// Research notes are deliberately a separate plain-text/markdown file per
// country (e.g. AtlasData/uk-notes.md), not a field inside the JSON data
// file — writing prose into a JSON string means hand-escaping quotes and
// line breaks every time, which isn't a reasonable ask for ongoing research
// notes. A missing notes file is a normal state (nothing written yet for
// that country), not an error, so it returns 200 with empty text rather
// than a 404.
app.http('getNotes', {
  methods: ['GET'],
  authLevel: 'anonymous',
  route: 'notes/{country}',
  handler: async (request, context) => {
    const country = (request.params.country || '').toLowerCase();

    if (!/^[a-z0-9-]+$/.test(country)) {
      return { status: 400, jsonBody: { error: 'Invalid country identifier' } };
    }

    const fileName = `${country}-notes.md`;
    const graphUrl = `https://graph.microsoft.com/v1.0/sites/${SITE_ID}/drive/root:/${DATA_FOLDER}/${fileName}:/content`;

    try {
      const token = await getAccessToken();

      const graphResponse = await fetch(graphUrl, {
        headers: { Authorization: `Bearer ${token}` }
      });

      if (graphResponse.status === 404) {
        return {
          status: 200,
          headers: { 'Content-Type': 'text/plain; charset=utf-8', 'Cache-Control': 'no-store' },
          body: ''
        };
      }

      if (!graphResponse.ok) {
        const detail = await graphResponse.text();
        context.error(`Graph error ${graphResponse.status}: ${detail}`);
        return { status: 502, jsonBody: { error: 'Could not fetch notes from SharePoint', status: graphResponse.status } };
      }

      const text = await graphResponse.text();

      return {
        status: 200,
        headers: { 'Content-Type': 'text/plain; charset=utf-8', 'Cache-Control': 'no-store' },
        body: text
      };
    } catch (err) {
      context.error(err);
      return { status: 500, jsonBody: { error: 'Server error fetching notes', detail: err.message } };
    }
  }
});
