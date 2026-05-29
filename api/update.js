const https = require('https');
const querystring = require('querystring');
const CLIENT_ID = process.env.CLIENT_ID;
const TENANT_ID = process.env.TENANT_ID;
const REFRESH_TOKEN = process.env.REFRESH_TOKEN;
const TABLE_NAME = 'ITR_Data';

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

async function getToken() {
  const b = querystring.stringify({ grant_type: 'refresh_token', client_id: CLIENT_ID, refresh_token: REFRESH_TOKEN, scope: 'Files.ReadWrite offline_access' });
  const r = await request({ hostname: 'login.microsoftonline.com', path: `/${TENANT_ID}/oauth2/v2.0/token`, method: 'POST', headers: { 'Content-Type': 'application/x-www-form-urlencoded', 'Content-Length': Buffer.byteLength(b) } }, b);
  if (!r.b.access_token) throw new Error('Token error: ' + JSON.stringify(r.b));
  return r.b.access_token;
}

async function graph(token, method, path, data) {
  const body = data ? JSON.stringify(data) : null;
  const headers = { 'Authorization': `Bearer ${token}`, 'Content-Type': 'application/json' };
  if (body) headers['Content-Length'] = Buffer.byteLength(body);
  return request({ hostname: 'graph.microsoft.com', path: `/v1.0${path}`, method, headers }, body);
}

module.exports = async (req, res) => {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, GET, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type, Authorization');
  if (req.method === 'OPTIONS') return res.status(200).end();
  if (req.method === 'GET') return res.status(200).json({ status: 'Jupiter ITR Backend running OK', time: new Date().toISOString() });
  if (req.method !== 'POST') return res.status(405).json({ error: 'Method not allowed' });

  try {
    const { srNo, action, expectedDate, filerName, filerEmail, dateOfFiling } = req.body;
    if (!srNo || !action) return res.status(400).json({ error: 'Missing srNo or action' });

    const token = await getToken();

    let driveId = process.env.DRIVE_ID;
    let fileId = process.env.FILE_ID;

    if (!driveId || !fileId) {
      const search = await graph(token, 'GET', `/me/drive/search(q='Jupiter_IT_PowerAutomate_Backend')`);
      const file = (search.b.value || []).find(f => f.name && f.name.includes('Jupiter_IT_PowerAutomate_Backend'));
      if (!file) throw new Error('Excel file not found. Please set DRIVE_ID and FILE_ID env vars in Vercel.');
      driveId = file.parentReference.driveId;
      fileId = file.id;
    }

    const rows = await graph(token, 'GET', `/drives/${driveId}/items/${fileId}/workbook/tables/${TABLE_NAME}/rows`);
    if (rows.s !== 200) throw new Error('Cannot read table: ' + JSON.stringify(rows.b));

    let idx = -1;
    for (let i = 0; i < rows.b.value.length; i++) {
      if (String(rows.b.value[i].values[0][0]) === String(srNo)) { idx = i; break; }
    }
    if (idx === -1) return res.status(404).json({ error: 'Sr.No. ' + srNo + ' not found' });

    const vals = [...rows.b.value[idx].values[0]];
    if (action === 'under_process') { vals[14] = 'Under Process'; vals[15] = expectedDate || ''; }
    else if (action === 'filed') { vals[14] = 'Pending with Approver'; }
    else if (action === 'assign_filer') { vals[12] = filerName || ''; vals[13] = filerEmail || ''; vals[14] = 'Assigned to Filer'; }
    else if (action === 'date_filed') { vals[16] = dateOfFiling || ''; vals[14] = 'Filed'; }

    const update = await graph(token, 'PATCH', `/drives/${driveId}/items/${fileId}/workbook/tables/${TABLE_NAME}/rows/itemAt(index=${idx})`, { values: [vals] });
    if (update.s !== 200) throw new Error('Update failed: ' + JSON.stringify(update.b));

    return res.status(200).json({ success: true, message: 'Sr.No. ' + srNo + ' updated: ' + action });
  } catch (err) {
    console.error(err);
    return res.status(500).json({ error: err.message });
  }
};
