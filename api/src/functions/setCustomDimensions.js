const { app } = require('@azure/functions');
const { ConfidentialClientApplication } = require('@azure/msal-node');

const TENANT_ID = process.env.ATLAS_TENANT_ID;
const CLIENT_ID = process.env.ATLAS_CLIENT_ID;
const CLIENT_SECRET = process.env.ATLAS_CLIENT_SECRET;
const SITE_ID = process.env.ATLAS_SITE_ID;
const DATA_FOLDER = 'AtlasData';

// Same authorisation model as the other write endpoints.
const ALLOWED_EDITOR_EMAIL = (process.env.ATLAS_ALLOWED_EDITOR_EMAIL || '').toLowerCase();

// A custom factor's key is always client-generated as `custom_` plus a
// slugified label (see slugifyDimensionKey() in scorecard-dimensions.js) --
// this lets setScorecard.js/setScorecardBulk.js/setScorecardGlobal.js/
// setEnabledDimensions.js/setDimensionWeights.js each accept a custom key by
// pattern match alone, with no need to fetch this endpoint's own data to
// know which custom keys currently exist (see each file's own comment).
const CUSTOM_DIMENSION_KEY_RE = /^custom_[a-z0-9_]{1,50}$/;

// Layout ceiling, not a data-model limit -- the PowerPoint export lays out
// one fixed-height row per dimension on a single slide (see exportPptx.js's
// SCORECARD_ROW_H), so an unbounded number of custom factors would eventually
// overflow the slide. 8 leaves comfortable headroom alongside the existing
// 12 fixed dimensions, AUM/allocation/Scored/Overall rows.
const MAX_CUSTOM_DIMENSIONS = 8;
const MAX_LABEL_LENGTH = 80;

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

// Saves the client-defined additional scorecard factors (Peter's 2026-07-30
// request: "clients might want to add in their own factors") onto
// global.json's top-level custom_dimensions field -- an array of
// {key, label, weight} entries, same "global and persisted, not a per-session
// filter" model as enabled_dimensions/dimension_weights: every reader of
// country.html/picker.html/overview.html/map.html and every Word/PPT export
// sees the same custom factor set once saved. Replaces the whole array on
// each save (not a per-key patch), same as setEnabledDimensions.js/
// setDimensionWeights.js, since the client always sends its full current
// state from the Factor selector panel (add/remove/rename all happen
// client-side first, then the resulting full list is pushed here).
app.http('setCustomDimensions', {
  methods: ['POST'],
  authLevel: 'anonymous',
  route: 'global/set-custom-dimensions',
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

    const { custom_dimensions } = body || {};
    if (!Array.isArray(custom_dimensions)) {
      return { status: 400, jsonBody: { error: 'custom_dimensions array is required (an empty array clears all custom factors)' } };
    }
    if (custom_dimensions.length > MAX_CUSTOM_DIMENSIONS) {
      return { status: 400, jsonBody: { error: `No more than ${MAX_CUSTOM_DIMENSIONS} custom factors are supported` } };
    }

    const seenKeys = new Set();
    for (const dim of custom_dimensions) {
      if (!dim || typeof dim !== 'object') {
        return { status: 400, jsonBody: { error: 'Each custom_dimensions entry must be an object' } };
      }
      if (typeof dim.key !== 'string' || !CUSTOM_DIMENSION_KEY_RE.test(dim.key)) {
        return { status: 400, jsonBody: { error: `Invalid custom dimension key: ${dim.key}` } };
      }
      if (seenKeys.has(dim.key)) {
        return { status: 400, jsonBody: { error: `Duplicate custom dimension key: ${dim.key}` } };
      }
      seenKeys.add(dim.key);
      if (typeof dim.label !== 'string' || !dim.label.trim() || dim.label.length > MAX_LABEL_LENGTH) {
        return { status: 400, jsonBody: { error: `Each custom dimension needs a label up to ${MAX_LABEL_LENGTH} characters` } };
      }
      if (typeof dim.weight !== 'number' || !Number.isFinite(dim.weight) || dim.weight < 0 || dim.weight > 5) {
        return { status: 400, jsonBody: { error: 'Each custom dimension needs a weight between 0 and 5' } };
      }
    }

    // Normalise to just {key, label, weight} -- strip anything else a client
    // might have sent along.
    const cleaned = custom_dimensions.map((d) => ({ key: d.key, label: d.label.trim(), weight: d.weight }));

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

      data.custom_dimensions = cleaned;
      data.custom_dimensions_updated_date = new Date().toISOString().slice(0, 10);
      data.custom_dimensions_updated_by = userEmail;

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

      return {
        status: 200,
        jsonBody: {
          success: true,
          custom_dimensions: data.custom_dimensions,
          updated_date: data.custom_dimensions_updated_date,
          updated_by: data.custom_dimensions_updated_by
        }
      };
    } catch (err) {
      context.error(err);
      return { status: 500, jsonBody: { error: 'Server error updating global.json', detail: err.message } };
    }
  }
});
