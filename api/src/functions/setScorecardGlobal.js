const { app } = require('@azure/functions');
const { ConfidentialClientApplication } = require('@azure/msal-node');

const TENANT_ID = process.env.ATLAS_TENANT_ID;
const CLIENT_ID = process.env.ATLAS_CLIENT_ID;
const CLIENT_SECRET = process.env.ATLAS_CLIENT_SECRET;
const SITE_ID = process.env.ATLAS_SITE_ID;
const DATA_FOLDER = 'AtlasData';

// Same authorisation model as setScorecard.js / setScorecardBulk.js.
const ALLOWED_EDITOR_EMAIL = (process.env.ATLAS_ALLOWED_EDITOR_EMAIL || '').toLowerCase();

// Kept in sync with scorecard-dimensions.js on the front end.
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

// Sets one scorecard dimension across many (country, segment) pairs at once,
// each with its own value -- unlike setScorecardBulk.js (one value applied
// to every segment in a single country), this is for the global-overview
// "classify every segment by an asset-allocation threshold" feature, where
// each segment's score is computed client-side from its own data and can
// differ segment to segment and country to country. Batches by country (one
// SharePoint read + one write per country, not per segment) to keep this to
// roughly one Graph round-trip per country regardless of how many segments
// in it changed.
// 2026-07-24: writes to {code}_scores.json, not {code}.json -- see
// getScores.js for why scores were split into their own file.
app.http('setScorecardGlobal', {
  methods: ['POST'],
  authLevel: 'anonymous',
  route: 'segments/set-scorecard-global',
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

    const { dimension, assignments } = body || {};
    if (!dimension || !VALID_DIMENSIONS.includes(dimension)) {
      return { status: 400, jsonBody: { error: `Unknown scorecard dimension: ${dimension}` } };
    }
    if (!Array.isArray(assignments) || !assignments.length) {
      return { status: 400, jsonBody: { error: 'assignments must be a non-empty array of {country, segment, value}' } };
    }
    for (const a of assignments) {
      if (!a || !a.country || !a.segment || ![1, 2, 3].includes(a.value)) {
        return { status: 400, jsonBody: { error: 'Each assignment needs country, segment, and value in [1, 2, 3]' } };
      }
    }

    // Group by country so each SharePoint file is read and written once,
    // regardless of how many of its segments are being updated.
    const byCountry = {};
    for (const a of assignments) {
      const countryCode = String(a.country).toLowerCase();
      if (!/^[a-z0-9-]+$/.test(countryCode)) {
        return { status: 400, jsonBody: { error: `Invalid country identifier: ${a.country}` } };
      }
      if (!byCountry[countryCode]) byCountry[countryCode] = [];
      byCountry[countryCode].push(a);
    }

    const token = await getAccessToken();
    const today = new Date().toISOString().slice(0, 10);
    const results = {};

    for (const [countryCode, countryAssignments] of Object.entries(byCountry)) {
      const scoresFileName = `${countryCode}_scores.json`;
      const scoresGraphBase = `https://graph.microsoft.com/v1.0/sites/${SITE_ID}/drive/root:/${DATA_FOLDER}/${scoresFileName}`;

      try {
        // No need to touch {code}.json at all here -- the caller already
        // knows which segments exist (it built `assignments` from data it
        // already loaded), so this only ever reads/writes the scores file.
        const scoresResp = await fetch(`${scoresGraphBase}:/content`, {
          headers: { Authorization: `Bearer ${token}` }
        });
        const scoresData = scoresResp.status === 404
          ? { country_code: countryCode, scores: {} }
          : (scoresResp.ok ? await scoresResp.json() : null);

        if (scoresData === null) {
          results[countryCode] = { success: false, error: `Could not read scores from SharePoint (status ${scoresResp.status})` };
          continue;
        }
        if (!scoresData.scores || typeof scoresData.scores !== 'object') scoresData.scores = {};

        let updated = 0;
        countryAssignments.forEach(a => {
          if (!scoresData.scores[a.segment]) scoresData.scores[a.segment] = {};
          scoresData.scores[a.segment][dimension] = a.value;
          scoresData.scores[a.segment].scored_date = today;
          updated += 1;
        });

        const putResp = await fetch(`${scoresGraphBase}:/content`, {
          method: 'PUT',
          headers: {
            Authorization: `Bearer ${token}`,
            'Content-Type': 'application/json'
          },
          body: JSON.stringify(scoresData, null, 2)
        });

        if (!putResp.ok) {
          const detail = await putResp.text();
          context.error(`Graph PUT error ${putResp.status} for ${countryCode}: ${detail}`);
          results[countryCode] = { success: false, error: `Could not save changes (status ${putResp.status})` };
          continue;
        }

        results[countryCode] = { success: true, updated };
      } catch (err) {
        context.error(`Error updating ${countryCode}: ${err.message}`);
        results[countryCode] = { success: false, error: err.message };
      }
    }

    const anySuccess = Object.values(results).some(r => r.success);
    return {
      status: anySuccess ? 200 : 502,
      jsonBody: { success: anySuccess, results }
    };
  }
});
