const https = require('https');
const querystring = require('querystring');
const CLIENT_ID = process.env.CLIENT_ID;
const TENANT_ID = process.env.TENANT_ID;
const REFRESH_TOKEN = process.env.REFRESH_TOKEN;

function request(options, body) {
  return new Promise((resolve, reject) => {
    const req = https.request(options, res => {
      let d = '';
      res.on('data', c => d += c);
      res.on('end', () => { try { resolve({ s: res.statusCode, b: JSON.parse(d) }); } catch(e) { resolve({ s: res.statusCode, b: d }); } });
    });
    req.on('error', reject);
    if (body) req.write(body);
    req.end();
  });
}

async function graph(token, path) {
  return request({ hostname: 'graph.microsoft.com', path: `/v1.0${path}`, method: 'GET', headers: { 'Authorization': `Bearer ${token}`, 'Content-Type': 'application/json' } });
}

module.exports = async (req, res) => {
  try {
    const b = querystring.stringify({ grant_type: 'refresh_token', client_id: CLIENT_ID, refresh_token: REFRESH_TOKEN, scope: 'Files.ReadWrite offline_access' });
    const tok = (await request({ hostname: 'login.microsoftonline.com', path: `/${TENANT_ID}/oauth2/v2.0/token`, method: 'POST', headers: { 'Content-Type': 'application/x-www-form-urlencoded', 'Content-Length': Buffer.byteLength(b) } }, b)).b;
    if (!tok.access_token) return res.status(500).json({ error: 'Token failed', detail: tok });
    const token = tok.access_token;

    const search = await graph(token, `/me/drive/search(q='Jupiter_IT_PowerAutomate_Backend')`);
    const shared = await graph(token, `/me/drive/sharedWithMe`);
    const drives = await graph(token, `/me/drives`);

    const files = (search.b.value || []).filter(f => f.name && f.name.includes('Jupiter'));
    const sharedFiles = (shared.b.value || []).filter(f => f.name && f.name.includes('Jupiter'));

    return res.status(200).json({
      driveSearch: files.map(f => ({ name: f.name, FILE_ID: f.id, DRIVE_ID: f.parentReference.driveId, path: f.parentReference.path })),
      sharedWithMe: sharedFiles.map(f => ({ name: f.name, FILE_ID: f.remoteItem ? f.remoteItem.id : f.id, DRIVE_ID: f.remoteItem ? f.remoteItem.parentReference.driveId : f.parentReference.driveId })),
      allDrives: (drives.b.value || []).map(d => ({ name: d.name, id: d.id, type: d.driveType }))
    });
  } catch (err) {
    return res.status(500).json({ error: err.message });
  }
};
