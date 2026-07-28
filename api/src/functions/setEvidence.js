const { app } = require('@azure/functions');
const { ConfidentialClientApplication } = require('@azure/msal-node');

const TENANT_ID = process.env.ATLAS_TENANT_ID;
const CLIENT_ID = process.env.ATLAS_CLIENT_ID;
const CLIENT_SECRET = process.env.ATLAS_CLIENT_SECRET;
const SITE_ID = process.env.ATLAS_SITE_ID;
const DATA_FOLDER = 'AtlasData';
const FILE_NAME = 'evidence_library.json';

// Same authorisation model as setCommentary.js / addSource.js -- anyone can
// sign in via Static Web Apps' built-in Entra login, but only a request
// carrying this exact email is allowed to write anything.
const ALLOWED_EDITOR_EMAIL = (process.env.ATLAS_ALLOWED_EDITOR_EMAIL || '').toLowerCase();

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

// Turns a title into a stable-ish slug id, with a numeric suffix if that
// slug's already taken -- e.g. two entries both titled "Mercer Fee Survey"
// in different years become mercer-fee-survey and mercer-fee-survey-2.
function slugify(title, existingIds) {
  const base = String(title || 'entry')
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '-')
    .replace(/^-+|-+$/g, '') || 'entry';
  if (!existingIds.has(base)) return base;
  let i = 2;
  while (existingIds.has(`${base}-${i}`)) i += 1;
  return `${base}-${i}`;
}

// figures[] rows are free-form enough (region/metric/value/unit/as_of) that
// there's no strict schema to enforce beyond "value should be a number if
// present at all" -- same light-touch sanitizing as setCommentary.js's
// sanitizeSources(), just enough to avoid storing stray blank rows from the
// on-screen "+ Add figure" button.
function sanitizeFigures(rawFigures) {
  if (!Array.isArray(rawFigures)) return [];
  return rawFigures
    .map((f) => ({
      region: typeof (f && f.region) === 'string' ? f.region.trim() : '',
      metric: typeof (f && f.metric) === 'string' ? f.metric.trim() : '',
      value: typeof (f && f.value) === 'number' && !Number.isNaN(f.value) ? f.value : (typeof (f && f.value) === 'string' && f.value.trim() && !Number.isNaN(Number(f.value)) ? Number(f.value) : undefined),
      unit: typeof (f && f.unit) === 'string' ? f.unit.trim() : '',
      as_of: typeof (f && f.as_of) === 'string' ? f.as_of.trim() : ''
    }))
    .filter((f) => f.region || f.metric || typeof f.value === 'number');
}

function sanitizeStringArray(raw) {
  if (Array.isArray(raw)) return raw.map((s) => String(s).trim()).filter(Boolean);
  if (typeof raw === 'string') return raw.split(',').map((s) => s.trim()).filter(Boolean);
  return [];
}

// One endpoint handles add, edit and delete for the supporting-evidence
// library (see evidence.html / evidence_library.json) -- same read-modify-
// write-the-whole-file approach as setCommentary.js, just against the one
// shared evidence file instead of a per-country file. `id` present + found
// => update that entry in place; `id` present + not found => error rather
// than silently creating a duplicate; `id` absent => create a new entry
// with a slug generated from the title. `delete: true` + a valid `id`
// removes that entry instead of touching its fields.
app.http('setEvidence', {
  methods: ['POST'],
  authLevel: 'anonymous',
  route: 'evidence/set',
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

    const graphBase = `https://graph.microsoft.com/v1.0/sites/${SITE_ID}/drive/root:/${DATA_FOLDER}/${FILE_NAME}`;

    try {
      const token = await getAccessToken();

      const getResp = await fetch(`${graphBase}:/content`, {
        headers: { Authorization: `Bearer ${token}` }
      });

      let data;
      if (getResp.status === 404) {
        // No file uploaded yet -- start a fresh one rather than erroring,
        // so the very first entry (added through the UI) can create it.
        data = { schema_note: 'Cross-cutting supporting evidence -- see evidence.html.', entries: [] };
      } else if (!getResp.ok) {
        return { status: 502, jsonBody: { error: `Could not read ${FILE_NAME} from SharePoint`, status: getResp.status } };
      } else {
        data = await getResp.json();
      }
      if (!Array.isArray(data.entries)) data.entries = [];

      const today = new Date().toISOString().slice(0, 10);

      if (body.delete) {
        const { id } = body;
        if (!id) return { status: 400, jsonBody: { error: 'id is required to delete an entry' } };
        const before = data.entries.length;
        data.entries = data.entries.filter((e) => e.id !== id);
        if (data.entries.length === before) {
          return { status: 404, jsonBody: { error: `No entry found with id '${id}'` } };
        }
      } else {
        const title = typeof body.title === 'string' ? body.title.trim() : '';
        if (!title) return { status: 400, jsonBody: { error: 'title is required' } };

        const entryFields = {
          title,
          theme: typeof body.theme === 'string' ? body.theme.trim() : '',
          scope: sanitizeStringArray(body.scope),
          related_dimensions: sanitizeStringArray(body.related_dimensions),
          source: typeof body.source === 'string' ? body.source.trim() : '',
          as_of: typeof body.as_of === 'string' ? body.as_of.trim() : '',
          access: typeof body.access === 'string' ? body.access.trim() : '',
          summary: typeof body.summary === 'string' ? body.summary.trim() : '',
          figures: sanitizeFigures(body.figures),
          note: typeof body.note === 'string' ? body.note.trim() : ''
        };

        if (body.id) {
          const existing = data.entries.find((e) => e.id === body.id);
          if (!existing) {
            return { status: 404, jsonBody: { error: `No entry found with id '${body.id}'` } };
          }
          Object.assign(existing, entryFields, { last_updated: today });
        } else {
          const existingIds = new Set(data.entries.map((e) => e.id));
          const id = slugify(title, existingIds);
          data.entries.push(Object.assign({ id }, entryFields, { added: { date: today, by: userEmail } }));
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

      return { status: 200, jsonBody: { success: true, entries: data.entries } };
    } catch (err) {
      context.error(err);
      return { status: 500, jsonBody: { error: 'Server error updating evidence library', detail: err.message } };
    }
  }
});
