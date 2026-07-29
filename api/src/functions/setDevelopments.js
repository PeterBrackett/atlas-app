const { app } = require('@azure/functions');
const { ConfidentialClientApplication } = require('@azure/msal-node');

const TENANT_ID = process.env.ATLAS_TENANT_ID;
const CLIENT_ID = process.env.ATLAS_CLIENT_ID;
const CLIENT_SECRET = process.env.ATLAS_CLIENT_SECRET;
const SITE_ID = process.env.ATLAS_SITE_ID;
const DATA_FOLDER = 'AtlasData';

// Same authorisation model as setCommentary.js / setEvidence.js.
const ALLOWED_EDITOR_EMAIL = (process.env.ATLAS_ALLOWED_EDITOR_EMAIL || '').toLowerCase();

// Same fixed section set as setCommentary.js's VALID_SECTIONS -- developments
// are tagged to the same 8 commentary sections so each section's bottom-right
// box only shows items relevant to it.
const VALID_SECTIONS = ['wealth', 'pensions', 'insurance', 'charities', 'foundations', 'family_offices', 'sovereign_wealth_funds', 'ocio'];

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

// Turns a headline into a stable-ish slug id, unique within that section's
// own array (not globally) -- same approach as setEvidence.js's slugify(),
// just scoped smaller since developments are per-section, per-country.
function slugify(headline, existingIds) {
  const base = String(headline || 'item')
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '-')
    .replace(/^-+|-+$/g, '') || 'item';
  if (!existingIds.has(base)) return base;
  let i = 2;
  while (existingIds.has(`${base}-${i}`)) i += 1;
  return `${base}-${i}`;
}

// One endpoint handles add, edit and delete for a country's recent-
// developments entries -- same read-modify-write-the-whole-file shape as
// setEvidence.js, just against {code}_developments.json (see
// getDevelopments.js) instead of the shared evidence file, and with entries
// nested one level deeper under a section key. `id` present + found => update
// in place; `id` present + not found => error; `id` absent => create a new
// entry with a slug generated from the headline. `delete: true` + `id`
// removes that entry.
app.http('setDevelopments', {
  methods: ['POST'],
  authLevel: 'anonymous',
  route: 'developments/set',
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
      return { status: 400, jsonBody: { error: `Unknown developments section: ${section}` } };
    }

    const countryCode = String(country).toLowerCase();
    if (!/^[a-z0-9-]+$/.test(countryCode)) {
      return { status: 400, jsonBody: { error: 'Invalid country identifier' } };
    }

    const fileName = `${countryCode}_developments.json`;
    const graphBase = `https://graph.microsoft.com/v1.0/sites/${SITE_ID}/drive/root:/${DATA_FOLDER}/${fileName}`;

    try {
      const token = await getAccessToken();

      const getResp = await fetch(`${graphBase}:/content`, {
        headers: { Authorization: `Bearer ${token}` }
      });

      let data;
      if (getResp.status === 404) {
        data = { country_code: countryCode, developments: {} };
      } else if (!getResp.ok) {
        return { status: 502, jsonBody: { error: `Could not read ${fileName} from SharePoint`, status: getResp.status } };
      } else {
        data = await getResp.json();
      }
      if (!data.developments || typeof data.developments !== 'object') data.developments = {};
      if (!Array.isArray(data.developments[section])) data.developments[section] = [];

      const list = data.developments[section];
      const today = new Date().toISOString().slice(0, 10);

      if (body.delete) {
        const { id } = body;
        if (!id) return { status: 400, jsonBody: { error: 'id is required to delete an entry' } };
        const before = list.length;
        data.developments[section] = list.filter((e) => e.id !== id);
        if (data.developments[section].length === before) {
          return { status: 404, jsonBody: { error: `No development entry found with id '${id}'` } };
        }
      } else {
        const headline = typeof body.headline === 'string' ? body.headline.trim() : '';
        if (!headline) return { status: 400, jsonBody: { error: 'headline is required' } };

        const entryFields = {
          headline,
          date: typeof body.date === 'string' ? body.date.trim() : '',
          summary: typeof body.summary === 'string' ? body.summary.trim() : '',
          source: typeof body.source === 'string' ? body.source.trim() : '',
          url: typeof body.url === 'string' ? body.url.trim() : ''
        };

        if (body.id) {
          const existing = list.find((e) => e.id === body.id);
          if (!existing) {
            return { status: 404, jsonBody: { error: `No development entry found with id '${body.id}'` } };
          }
          Object.assign(existing, entryFields, { last_updated: today });
        } else {
          const existingIds = new Set(list.map((e) => e.id));
          const id = slugify(headline, existingIds);
          list.push(Object.assign({ id }, entryFields, { added: { date: today, by: userEmail } }));
        }
      }

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

      return { status: 200, jsonBody: { success: true, developments: data.developments } };
    } catch (err) {
      context.error(err);
      return { status: 500, jsonBody: { error: 'Server error updating developments', detail: err.message } };
    }
  }
});
