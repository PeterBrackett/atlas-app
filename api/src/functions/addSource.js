const { app } = require('@azure/functions');
const { ConfidentialClientApplication } = require('@azure/msal-node');

const TENANT_ID = process.env.ATLAS_TENANT_ID;
const CLIENT_ID = process.env.ATLAS_CLIENT_ID;
const CLIENT_SECRET = process.env.ATLAS_CLIENT_SECRET;
const SITE_ID = process.env.ATLAS_SITE_ID;
const DATA_FOLDER = 'AtlasData';

// Same single-trusted-editor gate as setActiveSource.js and setScorecard.js —
// anyone can sign in via Static Web Apps' built-in Entra login, but only a
// request carrying this exact email is allowed to write anything.
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

// Appends a brand-new candidate figure to a segment's sources[] list --
// this is the "I found an official number, let me record it" action,
// distinct from setActiveSource.js which only switches which *existing*
// source is prevailing. Only targets SharePoint's AtlasData copy (the same
// store setActiveSource.js and getData.js use); countries without a
// SharePoint file yet (e.g. Ireland, still served from the GitHub static
// fallback) aren't reachable through this endpoint until one exists there.
app.http('addSource', {
  methods: ['POST'],
  authLevel: 'anonymous',
  route: 'segments/add-source',
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

    const { country, segment, basis, aum_bn, as_of, source, note, setActive } = body || {};
    if (!country || !segment || !basis || typeof aum_bn !== 'number' || Number.isNaN(aum_bn)) {
      return { status: 400, jsonBody: { error: 'country, segment, basis, and a numeric aum_bn are required' } };
    }

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
        return { status: 404, jsonBody: { error: `No SharePoint data file found for '${countryCode}' yet -- this endpoint only edits SharePoint's copy, not the GitHub fallback.` } };
      }
      if (!getResp.ok) {
        return { status: 502, jsonBody: { error: `Could not read ${fileName} from SharePoint`, status: getResp.status } };
      }
      const data = await getResp.json();

      const seg = (data.segments || []).find(s => s.segment === segment);
      if (!seg) {
        return { status: 404, jsonBody: { error: 'Segment not found' } };
      }

      if (!Array.isArray(seg.sources)) seg.sources = [];

      const newSource = {
        basis: String(basis),
        aum_bn,
        as_of: as_of ? String(as_of) : 'unknown',
        source: source ? String(source) : '',
        note: note ? String(note) : '',
        active: false
      };

      if (setActive) {
        seg.sources.forEach((src) => { src.active = false; });
        newSource.active = true;
        seg.aum_bn = newSource.aum_bn;
        seg.basis = newSource.basis;
      }

      seg.sources.push(newSource);

      if (!Array.isArray(seg.source_history)) seg.source_history = [];
      seg.source_history.push({
        date: new Date().toISOString().slice(0, 10),
        event: setActive ? 'added+switch' : 'added',
        by: userEmail,
        detail: `New source added: "${newSource.basis}" (${newSource.aum_bn.toLocaleString(undefined, { maximumFractionDigits: 2 })}bn, as of ${newSource.as_of}).${setActive ? ' Set as the active source immediately.' : ''}`
      });

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

      return { status: 200, jsonBody: { success: true, segment: seg } };
    } catch (err) {
      context.error(err);
      return { status: 500, jsonBody: { error: 'Server error adding source', detail: err.message } };
    }
  }
});
