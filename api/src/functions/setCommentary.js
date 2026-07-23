const { app } = require('@azure/functions');
const { ConfidentialClientApplication } = require('@azure/msal-node');

const TENANT_ID = process.env.ATLAS_TENANT_ID;
const CLIENT_ID = process.env.ATLAS_CLIENT_ID;
const CLIENT_SECRET = process.env.ATLAS_CLIENT_SECRET;
const SITE_ID = process.env.ATLAS_SITE_ID;
const DATA_FOLDER = 'AtlasData';

// Same authorisation model as setScorecardBulk.js / setActiveSource.js.
const ALLOWED_EDITOR_EMAIL = (process.env.ATLAS_ALLOWED_EDITOR_EMAIL || '').toLowerCase();

// The two fixed commentary sections -- see country.html's COMMENTARY_SECTIONS
// comment for the reasoning (Wealth & key pools of capital, Pensions
// structure). Kept as a small fixed set rather than free-form section keys
// so every country ends up with the same framework.
const VALID_SECTIONS = ['wealth', 'pensions'];

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

// sources is an array of {label, url} -- url is optional (a source can be a
// citation with no link, e.g. "S&P Money Manager Database export"), label is
// required so every listed source reads as something, not a blank line.
function sanitizeSources(rawSources) {
  if (!Array.isArray(rawSources)) return [];
  return rawSources
    .map((s) => ({
      label: typeof (s && s.label) === 'string' ? s.label.trim() : '',
      url: typeof (s && s.url) === 'string' ? s.url.trim() : ''
    }))
    .filter((s) => s.label || s.url);
}

// One country's commentary is {wealth: {text, sources[]}, pensions: {text,
// sources[]}} on the country JSON's top-level `commentary` key, alongside
// `segments` -- this endpoint only ever touches the one section it's asked
// to save, leaving the other section (and everything else in the file)
// untouched, same read-modify-write-the-whole-file approach as
// setScorecardBulk.js.
app.http('setCommentary', {
  methods: ['POST'],
  authLevel: 'anonymous',
  route: 'commentary/set',
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

    const { country, section } = body || {};
    if (!country || !section) {
      return { status: 400, jsonBody: { error: 'country and section are required' } };
    }
    if (!VALID_SECTIONS.includes(section)) {
      return { status: 400, jsonBody: { error: `Unknown commentary section: ${section}` } };
    }
    const text = typeof body.text === 'string' ? body.text : '';
    const sources = sanitizeSources(body.sources);

    const countryCode = String(country).toLowerCase();
    if (!/^[a-z0-9-]+$/.test(countryCode)) {
      return { status: 400, jsonBody: { error: 'Invalid country identifier' } };
    }

    const fileName = `${countryCode}.json`;
    const graphBase = `https://graph.microsoft.com/v1.0/sites/${SITE_ID}/drive/root:/${DATA_FOLDER}/${fileName}`;

    try {
      const token = await getAccessToken();

      const getResp = await fetch(`${graphBase}:/content`, {
        headers: { Authorization: `Bearer ${token}` }
      });
      if (getResp.status === 404) {
        return { status: 404, jsonBody: { error: `No SharePoint data file found for '${countryCode}' yet.` } };
      }
      if (!getResp.ok) {
        return { status: 502, jsonBody: { error: `Could not read ${fileName} from SharePoint`, status: getResp.status } };
      }
      const data = await getResp.json();

      if (!data.commentary || typeof data.commentary !== 'object') data.commentary = {};
      const today = new Date().toISOString().slice(0, 10);
      data.commentary[section] = { text, sources, last_updated: today };

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

      return { status: 200, jsonBody: { success: true, commentary: data.commentary } };
    } catch (err) {
      context.error(err);
      return { status: 500, jsonBody: { error: 'Server error updating commentary', detail: err.message } };
    }
  }
});
