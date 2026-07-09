const { app } = require('@azure/functions');
const { ConfidentialClientApplication } = require('@azure/msal-node');

const TENANT_ID = process.env.ATLAS_TENANT_ID;
const CLIENT_ID = process.env.ATLAS_CLIENT_ID;
const CLIENT_SECRET = process.env.ATLAS_CLIENT_SECRET;
const SITE_ID = process.env.ATLAS_SITE_ID;
const DATA_FOLDER = 'AtlasData';

// Only this email is allowed to actually change anything, regardless of who
// manages to sign in. Azure Static Web Apps' built-in Microsoft Entra ID
// login (free on every plan) lets any Microsoft account authenticate — it's
// not tenant-restricted unless you pay for the Standard plan's custom auth.
// Rather than pay for that, authorisation is enforced here instead: anyone
// can log in, but only a request carrying this exact email can write.
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

// Azure Static Web Apps injects this header on every request once a user is
// signed in, containing base64-encoded JSON with their identity details.
function getClientPrincipal(request) {
  const header = request.headers.get('x-ms-client-principal');
  if (!header) return null;
  try {
    return JSON.parse(Buffer.from(header, 'base64').toString('utf-8'));
  } catch {
    return null;
  }
}

app.http('setActiveSource', {
  methods: ['POST'],
  authLevel: 'anonymous',
  route: 'segments/set-active-source',
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

    const { country, segment, sourceIndex } = body || {};
    if (!country || !segment || typeof sourceIndex !== 'number') {
      return { status: 400, jsonBody: { error: 'country, segment, and sourceIndex are required' } };
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
      if (!getResp.ok) {
        return { status: 502, jsonBody: { error: `Could not read ${fileName} from SharePoint`, status: getResp.status } };
      }
      const data = await getResp.json();

      const seg = (data.segments || []).find(s => s.segment === segment);
      if (!seg || !Array.isArray(seg.sources) || !seg.sources[sourceIndex]) {
        return { status: 404, jsonBody: { error: 'Segment or source index not found' } };
      }

      const previous = seg.sources.find(src => src.active);
      const chosen = seg.sources[sourceIndex];

      // No-op if the requested source is already active — don't log a
      // switch event that didn't actually change anything.
      if (previous !== chosen) {
        seg.sources.forEach((src, i) => { src.active = (i === sourceIndex); });
        seg.aum_bn = chosen.aum_bn;
        seg.basis = chosen.basis;

        if (!Array.isArray(seg.source_history)) seg.source_history = [];
        seg.source_history.push({
          date: new Date().toISOString().slice(0, 10),
          event: 'switch',
          by: userEmail,
          detail: `Active source switched from "${previous ? previous.basis : 'unknown'}" to "${chosen.basis}" (${typeof chosen.aum_bn === 'number' ? chosen.aum_bn.toLocaleString(undefined, { maximumFractionDigits: 2 }) + 'bn' : 'n/a'}).`
        });
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

      return { status: 200, jsonBody: { success: true, segment: seg } };
    } catch (err) {
      context.error(err);
      return { status: 500, jsonBody: { error: 'Server error updating source', detail: err.message } };
    }
  }
});
