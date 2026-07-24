const { app } = require('@azure/functions');
const { ConfidentialClientApplication } = require('@azure/msal-node');

const TENANT_ID = process.env.ATLAS_TENANT_ID;
const CLIENT_ID = process.env.ATLAS_CLIENT_ID;
const CLIENT_SECRET = process.env.ATLAS_CLIENT_SECRET;
const SITE_ID = process.env.ATLAS_SITE_ID;
const DATA_FOLDER = 'AtlasData';

// Same authorisation model as setScorecard.js / setActiveSource.js.
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

// Sets one scorecard dimension to the same value across every segment in a
// country, in a single read-modify-write -- the "this dimension genuinely
// doesn't vary by segment" shortcut (e.g. Languages required, Local presence
// required), requested so a country-wide value doesn't need entering once
// per segment by hand. Deliberately overwrites every segment's existing
// value for that dimension, including ones already scored differently --
// that's the point of a "set all" action, not a fill-blanks-only one; a user
// who wants a specific segment to differ can still edit that cell
// individually afterwards, same as any other scorecard edit.
// 2026-07-24: writes to {code}_scores.json, not {code}.json -- see
// getScores.js for why scores were split into their own file.
app.http('setScorecardBulk', {
  methods: ['POST'],
  authLevel: 'anonymous',
  route: 'segments/set-scorecard-bulk',
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

    const { country, dimension, value } = body || {};
    if (!country || !dimension) {
      return { status: 400, jsonBody: { error: 'country and dimension are required' } };
    }
    if (!VALID_DIMENSIONS.includes(dimension)) {
      return { status: 400, jsonBody: { error: `Unknown scorecard dimension: ${dimension}` } };
    }
    // value is either a score of 1-3, or null to clear every segment's dimension back to "not yet scored".
    if (value !== null && ![1, 2, 3].includes(value)) {
      return { status: 400, jsonBody: { error: 'value must be 1, 2, 3, or null' } };
    }

    const countryCode = String(country).toLowerCase();
    if (!/^[a-z0-9-]+$/.test(countryCode)) {
      return { status: 400, jsonBody: { error: 'Invalid country identifier' } };
    }

    const dataFileName = `${countryCode}.json`;
    const scoresFileName = `${countryCode}_scores.json`;
    const dataGraphBase = `https://graph.microsoft.com/v1.0/sites/${SITE_ID}/drive/root:/${DATA_FOLDER}/${dataFileName}`;
    const scoresGraphBase = `https://graph.microsoft.com/v1.0/sites/${SITE_ID}/drive/root:/${DATA_FOLDER}/${scoresFileName}`;

    try {
      const token = await getAccessToken();

      // {code}.json is read-only here -- it's only consulted to get the
      // list of segment names this country actually has, never written
      // back to. Scores themselves live in and are written to
      // {code}_scores.json exclusively (see getScores.js).
      const dataResp = await fetch(`${dataGraphBase}:/content`, {
        headers: { Authorization: `Bearer ${token}` }
      });
      if (dataResp.status === 404) {
        return { status: 404, jsonBody: { error: `No SharePoint data file found for '${countryCode}' yet.` } };
      }
      if (!dataResp.ok) {
        return { status: 502, jsonBody: { error: `Could not read ${dataFileName} from SharePoint`, status: dataResp.status } };
      }
      const countryData = await dataResp.json();

      const segments = countryData.segments || [];
      if (!segments.length) {
        return { status: 404, jsonBody: { error: 'No segments found for this country' } };
      }

      const scoresResp = await fetch(`${scoresGraphBase}:/content`, {
        headers: { Authorization: `Bearer ${token}` }
      });
      const scoresData = scoresResp.status === 404
        ? { country_code: countryCode, scores: {} }
        : (scoresResp.ok ? await scoresResp.json() : null);

      if (scoresData === null) {
        return { status: 502, jsonBody: { error: `Could not read ${scoresFileName} from SharePoint`, status: scoresResp.status } };
      }
      if (!scoresData.scores || typeof scoresData.scores !== 'object') scoresData.scores = {};

      const today = new Date().toISOString().slice(0, 10);
      segments.forEach(seg => {
        const name = seg.segment;
        if (!scoresData.scores[name]) scoresData.scores[name] = {};
        if (value === null) {
          delete scoresData.scores[name][dimension];
        } else {
          scoresData.scores[name][dimension] = value;
          scoresData.scores[name].scored_date = today;
        }
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
        context.error(`Graph PUT error ${putResp.status}: ${detail}`);
        return { status: 502, jsonBody: { error: 'Could not save changes to SharePoint', status: putResp.status } };
      }

      return { status: 200, jsonBody: { success: true, scores: scoresData.scores } };
    } catch (err) {
      context.error(err);
      return { status: 500, jsonBody: { error: 'Server error updating scorecard in bulk', detail: err.message } };
    }
  }
});
