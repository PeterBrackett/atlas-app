const { app } = require('@azure/functions');
const { ConfidentialClientApplication } = require('@azure/msal-node');

const TENANT_ID = process.env.ATLAS_TENANT_ID;
const CLIENT_ID = process.env.ATLAS_CLIENT_ID;
const CLIENT_SECRET = process.env.ATLAS_CLIENT_SECRET;
const SITE_ID = process.env.ATLAS_SITE_ID;
const DATA_FOLDER = 'AtlasData';
const FILE_NAME = 'sources.md';

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

// Serves a single, fixed file (AtlasData/sources.md) — the in-app home for
// methodology and sourcing notes, so explanations like "why is this figure
// a top-down override" live inside the app rather than only in an external
// design document. Same plain-markdown-file pattern as per-country notes,
// just not parameterised by country since this is one shared page.
app.http('getSources', {
  methods: ['GET'],
  authLevel: 'anonymous',
  route: 'sources',
  handler: async (request, context) => {
    const graphUrl = `https://graph.microsoft.com/v1.0/sites/${SITE_ID}/drive/root:/${DATA_FOLDER}/${FILE_NAME}:/content`;

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
        return { status: 502, jsonBody: { error: 'Could not fetch sources from SharePoint', status: graphResponse.status } };
      }

      const text = await graphResponse.text();

      return {
        status: 200,
        headers: { 'Content-Type': 'text/plain; charset=utf-8', 'Cache-Control': 'no-store' },
        body: text
      };
    } catch (err) {
      context.error(err);
      return { status: 500, jsonBody: { error: 'Server error fetching sources', detail: err.message } };
    }
  }
});
