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

// Scorecard values (the opportunity-matrix 1/2/3 ratings, keyed by segment
// name) used to live inline on each segment inside {code}.json, alongside
// AUM/allocation/sources/commentary. That meant every score was only as safe
// as the *next* write to that file -- and {code}.json gets rewritten wholesale
// by a range of things: data-refresh scripts re-ingesting a fresh S&P/OECD
// export, one-off bug-fix scripts, retone/commentary batch edits, and so on.
// Any of those doing a read-modify-write against a stale copy of the file (or
// simply not carrying the `scorecard` key forward) would silently wipe out
// whatever had been scored on the live site in between -- which is exactly
// what Peter reported happening. 2026-07-24: scores moved out into their own
// {code}_scores.json file, keyed by segment name, written to only by
// setScorecard.js / setScorecardBulk.js / setScorecardGlobal.js. Nothing that
// touches {code}.json touches this file, so no amount of data-refresh churn on
// the main file can ever affect scores again.
// This endpoint returns { scores: {} } (200, not 404) when no scores file
// exists yet for a country, so the front end can treat "not scored yet" and
// "file doesn't exist yet" the same way rather than special-casing a 404.
app.http('getScores', {
  methods: ['GET'],
  authLevel: 'anonymous',
  route: 'scores/{country}',
  handler: async (request, context) => {
    const country = (request.params.country || '').toLowerCase();

    if (!/^[a-z0-9-]+$/.test(country)) {
      return { status: 400, jsonBody: { error: 'Invalid country identifier' } };
    }

    const fileName = `${country}_scores.json`;
    const graphUrl = `https://graph.microsoft.com/v1.0/sites/${SITE_ID}/drive/root:/${DATA_FOLDER}/${fileName}:/content`;

    try {
      const token = await getAccessToken();

      const graphResponse = await fetch(graphUrl, {
        headers: { Authorization: `Bearer ${token}` }
      });

      if (graphResponse.status === 404) {
        return {
          jsonBody: { country_code: country, scores: {} },
          headers: { 'Cache-Control': 'no-store' }
        };
      }

      if (!graphResponse.ok) {
        const detail = await graphResponse.text();
        context.error(`Graph error ${graphResponse.status}: ${detail}`);
        return { status: 502, jsonBody: { error: 'Could not fetch scores from SharePoint', status: graphResponse.status } };
      }

      const data = await graphResponse.json();

      return {
        jsonBody: data,
        headers: { 'Cache-Control': 'no-store' }
      };
    } catch (err) {
      context.error(err);
      return { status: 500, jsonBody: { error: 'Server error fetching scores', detail: err.message } };
    }
  }
});
