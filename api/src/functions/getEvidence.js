const { app } = require('@azure/functions');
const { ConfidentialClientApplication } = require('@azure/msal-node');

const TENANT_ID = process.env.ATLAS_TENANT_ID;
const CLIENT_ID = process.env.ATLAS_CLIENT_ID;
const CLIENT_SECRET = process.env.ATLAS_CLIENT_SECRET;
const SITE_ID = process.env.ATLAS_SITE_ID;
const DATA_FOLDER = 'AtlasData';
const FILE_NAME = 'evidence_library.json';

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

// Serves a single, fixed file (AtlasData/evidence_library.json) -- the
// cross-cutting "supporting evidence" library (fee benchmarks, market-
// structure surveys, and similar reference material that applies across
// countries/regions rather than to one segment of one country). Same
// single-shared-file pattern as getSources.js, just JSON instead of
// markdown since evidence.html and picker.html's export picker both need
// structured fields (theme, scope, figures) to filter/render by, not just
// free text. Added 2026-07-27 per Peter's request for a separate area for
// this kind of material that doesn't touch the per-country {code}.json
// schema at all.
app.http('getEvidence', {
  methods: ['GET'],
  authLevel: 'anonymous',
  route: 'evidence',
  handler: async (request, context) => {
    const graphUrl = `https://graph.microsoft.com/v1.0/sites/${SITE_ID}/drive/root:/${DATA_FOLDER}/${FILE_NAME}:/content`;

    try {
      const token = await getAccessToken();

      const graphResponse = await fetch(graphUrl, {
        headers: { Authorization: `Bearer ${token}` }
      });

      if (graphResponse.status === 404) {
        // No SharePoint copy uploaded yet -- respond with an empty library
        // rather than an error, same reasoning as getSources.js's 404
        // handling, so evidence.html/picker.html can fall through to the
        // static data/evidence_library.json copy in the repo.
        return {
          status: 200,
          jsonBody: { entries: [] },
          headers: { 'Cache-Control': 'no-store' }
        };
      }

      if (!graphResponse.ok) {
        const detail = await graphResponse.text();
        context.error(`Graph error ${graphResponse.status}: ${detail}`);
        return { status: 502, jsonBody: { error: 'Could not fetch evidence library from SharePoint', status: graphResponse.status } };
      }

      const data = await graphResponse.json();

      return {
        jsonBody: data,
        headers: { 'Cache-Control': 'no-store' }
      };
    } catch (err) {
      context.error(err);
      return { status: 500, jsonBody: { error: 'Server error fetching evidence library', detail: err.message } };
    }
  }
});
