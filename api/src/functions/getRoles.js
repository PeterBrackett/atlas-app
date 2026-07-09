const { app } = require('@azure/functions');

// Static Web Apps calls this endpoint automatically after every successful
// sign-in, once "rolesSource": "/api/GetRoles" is set in
// staticwebapp.config.json's auth section. This replaces the portal's
// invitation-based role system entirely -- per Microsoft's own docs, once a
// rolesSource function is configured, invitation-assigned roles are ignored.
// SWA secures this route itself once rolesSource is set; it isn't reachable
// via a normal external HTTP call.
//
// Simple allow-list by email for now. Add colleagues here as needed rather
// than through the Role management portal blade, which invitations kept
// 404ing against.
const ATLAS_READERS = [
  'peter.brackett@institutionaladviser.co.uk'
];

app.http('GetRoles', {
  methods: ['POST'],
  authLevel: 'anonymous',
  route: 'GetRoles',
  handler: async (request, context) => {
    let body;
    try {
      body = await request.json();
    } catch {
      body = {};
    }

    const email = (body.userDetails || '').trim().toLowerCase();
    const isReader = ATLAS_READERS.some((e) => e.toLowerCase() === email);

    return {
      status: 200,
      jsonBody: { roles: isReader ? ['atlasreader'] : [] }
    };
  }
});
