/**
 * fetch-data.js — Weekly Figures 2026 dashboard
 * Mirrors the GIM dashboard pattern exactly:
 *   - Same Azure AD app (TENANT_ID, CLIENT_ID, CLIENT_SECRET)
 *   - Same Graph API flow: token → site → drive → file → download → parse
 *   - Writes data.json to repo root for index.html to fetch
 *
 * File on SharePoint: Weekly_Figures_2026.xlsx  (2 sheets: IA + GIM)
 */

const fetch = (...args) => import('node-fetch').then(({ default: f }) => f(...args));
const XLSX  = require('xlsx');
const fs    = require('fs');

const TENANT_ID     = process.env.TENANT_ID;
const CLIENT_ID     = process.env.CLIENT_ID;
const CLIENT_SECRET = process.env.CLIENT_SECRET;
const FILE_NAME     = process.env.FILE_NAME || 'Weekly_Figures_2026.xlsx';

const MONTHS = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];

async function main() {

  // ── 1. Auth (identical to GIM dashboard) ─────────────────────────────────
  console.log('Fetching access token...');
  const tokenResp = await fetch(
    `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`,
    {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams({
        grant_type:    'client_credentials',
        client_id:     CLIENT_ID,
        client_secret: CLIENT_SECRET,
        scope:         'https://graph.microsoft.com/.default',
      }).toString(),
    }
  );
  const tokenData = await tokenResp.json();
  if (!tokenData.access_token) {
    console.error('Token error:', JSON.stringify(tokenData));
    process.exit(1);
  }
  const token = tokenData.access_token;
  console.log('✓ Token obtained');

  // ── 2. Site (identical to GIM dashboard) ─────────────────────────────────
  console.log('Getting site...');
  const siteResp = await fetch(
    'https://graph.microsoft.com/v1.0/sites/seveninsurancebrokers.sharepoint.com:/sites/SIBIntranet',
    { headers: { Authorization: `Bearer ${token}` } }
  );
  const siteData = await siteResp.json();
  if (!siteData.id) {
    console.error('Site error:', JSON.stringify(siteData));
    process.exit(1);
  }
  const siteId = siteData.id;
  console.log('✓ Site ID:', siteId);

  // ── 3 & 4. Search every drive for the file ──────────────────────────────
  console.log('Listing drives...');
  const drivesResp = await fetch(
    `https://graph.microsoft.com/v1.0/sites/${siteId}/drives`,
    { headers: { Authorization: `Bearer ${token}` } }
  );
  const drivesData = await drivesResp.json();
  console.log('Available drives:', drivesData.value?.map(d => d.name));

  // Try every drive until the file is found
  let downloadUrl = null;
  for (const drive of (drivesData.value || [])) {
    console.log(`Searching drive: ${drive.name}...`);
    const fileResp = await fetch(
      `https://graph.microsoft.com/v1.0/drives/${drive.id}/root:/${FILE_NAME}`,
      { headers: { Authorization: `Bearer ${token}` } }
    );
    const fileData = await fileResp.json();
    if (fileData['@microsoft.graph.downloadUrl']) {
      console.log(`Found in drive: ${drive.name}`);
      downloadUrl = fileData['@microsoft.graph.downloadUrl'];
      break;
    }
    console.log(`  Not in ${drive.name}`);
  }
  if (!downloadUrl) {
    console.error(`File "${FILE_NAME}" not found in any drive. Check the exact filename on SharePoint.`);
    process.exit(1);
  }
  console.log('Download URL obtained');

  // ── 5. Download ───────────────────────────────────────────────────────────
  console.log('Downloading...');
  const dlResp = await fetch(downloadUrl);
  const buf    = await dlResp.arrayBuffer();
  console.log('✓ File size:', (buf.byteLength / 1024).toFixed(1), 'KB');

  // ── 6. Parse both sheets ──────────────────────────────────────────────────
  console.log('Parsing Excel...');
  const wb = XLSX.read(Buffer.from(buf), { type: 'buffer', cellDates: true });
  console.log('Sheet names:', wb.SheetNames);

  const gimSheetName = wb.SheetNames.find(n => n.includes('GIM'));
  const iaSheetName  = wb.SheetNames.find(n => n.includes('IA'));
  if (!gimSheetName || !iaSheetName) {
    console.error('Could not find GIM or IA sheets. Found:', wb.SheetNames);
    process.exit(1);
  }
  console.log(`✓ GIM sheet: "${gimSheetName}", IA sheet: "${iaSheetName}"`);

  const gimRaw = XLSX.utils.sheet_to_json(wb.Sheets[gimSheetName], { header: 1, defval: null, raw: false });
  const iaRaw  = XLSX.utils.sheet_to_json(wb.Sheets[iaSheetName],  { header: 1, defval: null, raw: false });

  const gimData = parseGIM(gimRaw);
  const iaData  = parseIA(iaRaw);
  console.log(`✓ Parsed: ${gimData.length} GIM brokers, ${iaData.length} IA brokers`);

  // ── 7. Find latest date for dashboard default landing ─────────────────────
  const allDates = [
    ...gimData.flatMap(b => b.weeks.map(w => w.date)),
    ...iaData.flatMap(b  => b.weeks.map(w => w.date)),
  ].sort();
  const latestDate  = allDates[allDates.length - 1] || '';
  const latestMonth = latestDate ? MONTHS[parseInt(latestDate.split('-')[1]) - 1] : '';

  // ── 8. Write data.json ────────────────────────────────────────────────────
  const output = {
    updated:     new Date().toISOString(),
    latestDate,
    latestMonth,
    gim: gimData,
    ia:  iaData,
  };
  fs.writeFileSync('data.json', JSON.stringify(output));
  const kb = (fs.statSync('data.json').size / 1024).toFixed(1);
  console.log(`✓ Saved data.json (${kb} KB) — latest: ${latestDate} (${latestMonth})`);
  console.log('=== Done ===');
}

// ── Helpers ──────────────────────────────────────────────────────────────────

function safeNum(v) {
  if (v == null || v === '') return 0;
  const n = parseFloat(v);
  return isNaN(n) ? 0 : Math.round(n * 100) / 100;
}

function getWeekCols(rows) {
  // Row index 3 has week start dates
  const cols = [];
  if (!rows[3]) return cols;
  rows[3].forEach((v, i) => {
    if (!v || i === 0) return;
    const d = v instanceof Date ? v : new Date(v);
    if (!isNaN(d.getTime())) {
      cols.push({ col: i, date: d.toISOString().slice(0, 10) });
    }
  });
  return cols;
}

function dateToMonth(dateStr) {
  return MONTHS[parseInt(dateStr.split('-')[1]) - 1];
}

function parseGIM(rows) {
  /**
   * GIM sheet: Row 3 = dates, Row 5 = metrics (step 4 per week)
   * Per week: +0=Sales, +1=Written Gross (Invoicing), +2=Issued Gross (Revenue)
   */
  const weekCols = getWeekCols(rows);
  const skipNames = new Set([
    'General & Medical Total A', 'TOPSURANCE GIM Total B',
    'General & Medical Total (A+B+C)', 'TOPSURANCE', 'General & Medical', 'new', '',
  ]);

  const brokers = [];
  for (let ri = 6; ri < rows.length; ri++) {
    if (!rows[ri]) continue;
    const name = rows[ri][0] ? String(rows[ri][0]).trim() : '';
    if (!name || skipNames.has(name)) continue;

    const weeks = [];
    for (const { col, date } of weekCols) {
      const sales = safeNum(rows[ri][col]);
      const inv   = safeNum(rows[ri][col + 1]); // Written Gross  = Invoicing
      const rev   = safeNum(rows[ri][col + 2]); // Issued Gross   = Revenue
      if (sales || inv || rev) {
        weeks.push({ date, month: dateToMonth(date), sales, inv, rev });
      }
    }
    if (weeks.length) brokers.push({ name, weeks });
  }
  return brokers;
}

function parseIA(rows) {
  /**
   * IA sheet: Row 3 = dates, Row 5 = metrics (step 5 per week)
   * Per week: +0=Issued Cases (Invoicing), +1=Trails (Revenue), +2=Written Indem, +3=Written Non-Indem
   */
  const weekCols = getWeekCols(rows);

  const teamRanges = [
    [9,  14, 'Alpha'],
    [18, 25, 'APW'],
    [30, 33, 'Prinden'],
    [39, 41, 'PB Team'],
    [47, 70, 'SVN Capital'],
    [75, 76, 'Joe Barnaby'],
    [82, 83, 'Abacus'],
    [88, 100, 'Individual'],
    [105, 107, 'Akhil'],
  ];

  const skipRows = new Set([15,27,36,44,72,79,85,102,109,114,116,131,133,143,145,147,149,157,158]);
  const skipNames = new Set([
    'Alpha Team','APW ','Prinden','PB Team','SVN CAPITAL','Joe Barnaby Team',
    'Abacus Team','INDIVIDUAL TEAM','AKHIL TEAM','Sonny Ridgewell  Others',
    'General & Medical','TOPSURANCE','Alpha Team Total','APW Team Total',
    'Prinden Total A','PB Team Total','SVN CAPITAL Total','Joe Barnaby Team  Total ',
    'Abacus Team  Total ','INDIVIDUAL TEAM TOTAL','Akhil  Team Total',
    'Sonny Ridgewell - Others Total','General & Medical Total A',
    'TOPSURANCE GIM Total B','General & Medical Total (A+B+C)',
    'GROUP TOTAL','Life Total','ia','gim','',
  ]);

  function getTeam(ri) {
    for (const [s, e, t] of teamRanges) if (ri >= s && ri <= e) return t;
    return 'Other';
  }

  const brokers = [];
  for (let ri = 8; ri < 150; ri++) {
    if (skipRows.has(ri) || !rows[ri]) continue;
    const name = rows[ri][0] ? String(rows[ri][0]).trim() : '';
    if (!name || skipNames.has(name)) continue;

    const team  = getTeam(ri);
    const weeks = [];
    for (const { col, date } of weekCols) {
      const inv    = safeNum(rows[ri][col]);       // Issued Cases Gross = Invoicing
      const trails = safeNum(rows[ri][col + 1]);   // Trails & Fees      = Revenue
      const wIndem = safeNum(rows[ri][col + 2]);   // Written Indemnified
      const wNon   = safeNum(rows[ri][col + 3]);   // Written Non-Indemnified
      const sales  = wIndem + wNon;
      if (inv || trails || sales) {
        weeks.push({ date, month: dateToMonth(date), inv, trails, sales });
      }
    }
    if (weeks.length) brokers.push({ name, team, weeks });
  }

  // Remove GIM brokers (belong in GIM dashboard) and catch-all Others rows
  const iaOthersSkip = new Set([
    'Alpha Others','PB-Others','TH-Others','SVN-Others','Akhil Others',
    'Others','Prinden Others','Sonny Ridgewell  Others',
  ]);
  return brokers.filter(b => b.team !== 'GIM (via IA)' && b.team !== 'Other' && !iaOthersSkip.has(b.name.trim()));
}

main().catch(err => {
  console.error('ERROR:', err.message);
  process.exit(1);
});
