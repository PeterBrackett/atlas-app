const { app } = require('@azure/functions');
const { ConfidentialClientApplication } = require('@azure/msal-node');

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

// "Recent developments" -- the bottom-right box on each country-page
// commentary section (see country.html's renderCommentarySection()) --
// lives in its own {code}_developments.json file, keyed by the same 8
// commentary section keys as {code}.json's `commentary` object:
// { developments: { wealth: [...], pensions: [...], ... } }. Split into its
// own file for the same reason scores were (see getScores.js): this content
// is meant to be refreshed periodically (a scheduled research pass, or a
// manual add), and keeping it out of {code}.json means a data-refresh script
// re-ingesting segment data can never wipe it out.
// Returns { country_code, developments: {} } (200, not 404) when no file
// exists yet, same "not built yet is fine" convention as getScores.js.
app.http('getDevelopments', {
  methods: ['GET'],
  authLevel: 'anonymous',
  route: 'developments/{country}',
  handler: async (request, context) => {
    const country = (request.params.country || '').toLowerCase();

    if (!/^[a-z0-9-]+$/.test(country)) {
      return { status: 400, jsonBody: { error: 'Invalid country identifier' } };
    }

    const fileName = `${country}_developments.json`;
    const graphUrl = `https://graph.microsoft.com/v1.0/sites/${SITE_ID}/drive/root:/${DATA_FOLDER}/${fileName}:/content`;

    try {
      const token = await getAccessToken();

      const graphResponse = await fetch(graphUrl, {
        headers: { Authorization: `Bearer ${token}` }
      });

      if (graphResponse.status === 404) {
        return {
          jsonBody: { country_code: country, developments: {} },
          headers: { 'Cache-Control': 'no-store' }
        };
      }

      if (!graphResponse.ok) {
        const detail = await graphResponse.text();
        context.error(`Graph error ${graphResponse.status}: ${detail}`);
        return { status: 502, jsonBody: { error: 'Could not fetch developments from SharePoint', status: graphResponse.status } };
      }

      const data = await graphResponse.json();

      return {
        jsonBody: data,
        headers: { 'Cache-Control': 'no-store' }
      };
    } catch (err) {
      context.error(err);
      return { status: 500, jsonBody: { error: 'Server error fetching developments', detail: err.message } };
    }
  }
});
