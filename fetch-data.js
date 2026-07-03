/**
 * fetch-data.js — Weekly Figures 2026 dashboard
 *   - Same Azure AD app (TENANT_ID, CLIENT_ID, CLIENT_SECRET)
 *   - Same Graph API flow: token → site → drive → file → download → parse
 *   - Writes data.json to repo root for index.html to fetch
 *
 * File on SharePoint: Weekly_Figures_2026.xlsx  (2 sheets: IA + GIM)
 *
 * FIX (2026-07): IA parser rewritten to be marker-driven instead of using
 * hardcoded row ranges. The old teamRanges cut off the LAST member of several
 * teams (Sabina, Brett Staveley-Abacus, Akhil Others, New Broker, Prinden
 * Others) because the range end was one row short — those rows fell through to
 * team 'Other' and were then filtered out. It also silently dropped every
 * zero-data member (whole APW team, SVN sub-brokers, Kartik Soni). The block
 * boundaries are now found from the team HEADER and TOTAL labels in column A,
 * so the parser self-corrects if rows are inserted/removed on SharePoint, and
 * every roster member is emitted (zero-data members included, at 0).
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

  // ── 1. Auth ───────────────────────────────────────────────────────────────
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

  // ── 2. Site ───────────────────────────────────────────────────────────────
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

  // ── 3 & 4. Search every drive for the file ───────────────────────────────
  console.log('Listing drives...');
  const drivesResp = await fetch(
    `https://graph.microsoft.com/v1.0/sites/${siteId}/drives`,
    { headers: { Authorization: `Bearer ${token}` } }
  );
  const drivesData = await drivesResp.json();
  console.log('Available drives:', drivesData.value?.map(d => d.name));

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

  // ── 5. Download ─────────────────────────────────────────────────────────────
  console.log('Downloading...');
  const dlResp = await fetch(downloadUrl);
  const buf    = await dlResp.arrayBuffer();
  console.log('✓ File size:', (buf.byteLength / 1024).toFixed(1), 'KB');

  // ── 6. Parse both sheets ────────────────────────────────────────────────────
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

  const gimRaw = XLSX.utils.sheet_to_json(wb.Sheets[gimSheetName], { header: 1, defval: null, raw: true });
  const iaRaw  = XLSX.utils.sheet_to_json(wb.Sheets[iaSheetName],  { header: 1, defval: null, raw: true });

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

  // ── 8. Write data.json ──────────────────────────────────────────────────────
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

function norm(s) {
  return String(s == null ? '' : s).trim().replace(/\s+/g, ' ');
}

function getWeekCols(rows) {
  // Row index 3 has week start dates. raw:true returns Excel serial numbers.
  const cols = [];
  if (!rows[3]) return cols;
  rows[3].forEach((v, i) => {
    if (!v || i === 0) return;
    let d;
    if (v instanceof Date) {
      d = v;
    } else if (typeof v === 'number') {
      d = new Date(Math.round((v - 25569) * 86400 * 1000));
    } else {
      d = new Date(v);
    }
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
  ].map(norm));

  const brokers = [];
  for (let ri = 6; ri < rows.length; ri++) {
    if (!rows[ri]) continue;
    const name = norm(rows[ri][0]);
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
    const group = ['Miro','Kelly','Ruben','Topsurance Others','Waqqaz','Aletha Mudyahoto','Leila Kedenge','Retention'].some(n=>name.startsWith(n)) ? 'Topsurance' : 'General & Medical';
    if (weeks.length) brokers.push({ name: rows[ri][0].toString().trim(), group, weeks });
  }
  return brokers;
}

function parseIA(rows) {
  /**
   * IA sheet: Row 3 = dates, weekly metric blocks step 5 per week.
   * Per week: +0=Issued Cases (Invoicing), +1=Trails (Revenue),
   *           +2=Written Indemnified, +3=Written Non-Indemnified
   * YTD summary columns: 276=Issued, 277=Trails, 280=W.Indem, 281=W.Non
   *
   * Team blocks are located by their HEADER label and TOTAL label in column A
   * (not hardcoded row numbers). Every member row between header and total is
   * emitted — including zero-data rows — so no team member ever disappears.
   */
  const weekCols = getWeekCols(rows);

  const blocks = [
    { team: 'Alpha',       header: 'Alpha Team',       total: 'Alpha Team Total' },
    { team: 'APW',         header: 'APW',              total: 'APW Team Total' },
    { team: 'Prinden',     header: 'Prinden',          total: 'Prinden Total A' },
    { team: 'PB Team',     header: 'PB Team',          total: 'PB Team Total' },
    { team: 'SVN Capital', header: 'SVN CAPITAL',      total: 'SVN CAPITAL Total' },
    { team: 'Joe Barnaby', header: 'Joe Barnaby Team', total: 'Joe Barnaby Team Total' },
    { team: 'Abacus',      header: 'Abacus Team',      total: 'Abacus Team Total' },
    { team: 'Individual',  header: 'INDIVIDUAL TEAM',  total: 'INDIVIDUAL TEAM TOTAL' },
    { team: 'Akhil',       header: 'AKHIL TEAM',       total: 'Akhil Team Total' },
  ];

  // Section markers / sub-headers that must never be treated as a member.
  const skipNames = new Set([
    'Sonny Ridgewell Others', 'General & Medical', 'TOPSURANCE',
    'GROUP TOTAL', 'Life Total', 'ia', 'gim', '',
  ].map(norm));

  const findRow = (label, from = 0) => {
    const L = norm(label);
    for (let r = from; r < rows.length; r++) {
      if (rows[r] && norm(rows[r][0]) === L) return r;
    }
    return -1;
  };

  const brokers = [];
  const seenWithData = new Set(); // name -> already emitted a version carrying figures
  let cursor = 0;

  for (const b of blocks) {
    const hRow = findRow(b.header, cursor);
    if (hRow < 0) { console.warn('parseIA: header not found:', b.header); continue; }
    let tRow = findRow(b.total, hRow + 1);
    if (tRow < 0) { console.warn('parseIA: total not found:', b.total); continue; }
    cursor = tRow;

    for (let ri = hRow + 1; ri < tRow; ri++) {
      if (!rows[ri]) continue;
      const name = norm(rows[ri][0]);
      if (!name || skipNames.has(name)) continue;
      if (name === norm(b.header) || /total/i.test(name)) continue;

      const weeks = [];
      for (const { col, date } of weekCols) {
        const inv    = safeNum(rows[ri][col]);
        const trails = safeNum(rows[ri][col + 1]);
        const wIndem = safeNum(rows[ri][col + 2]);
        const wNon   = safeNum(rows[ri][col + 3]);
        const sales  = Math.round((wIndem + wNon) * 100) / 100;
        if (inv || trails || sales) {
          weeks.push({ date, month: dateToMonth(date), inv, trails, sales, w_indem: wIndem, w_non: wNon });
        }
      }
      const inv_ytd     = safeNum(rows[ri][276]);
      const trails_ytd  = safeNum(rows[ri][277]);
      const w_indem_ytd = safeNum(rows[ri][280]);
      const w_non_ytd   = safeNum(rows[ri][281]);

      const hasData = weeks.length > 0 || inv_ytd || trails_ytd || w_indem_ytd || w_non_ytd;

      // De-dupe cross-listed placeholders (e.g. "Brett Staveley - Abacus" appears
      // as an empty row under SVN Capital and as the real row under Abacus).
      // If we've already emitted a version WITH data, skip an empty duplicate;
      // if this one has data and a prior empty was emitted, replace it.
      if (seenWithData.has(name) && !hasData) continue;
      if (hasData) {
        const prevEmptyIdx = brokers.findIndex(x => norm(x.name) === name);
        if (prevEmptyIdx >= 0) {
          const prev = brokers[prevEmptyIdx];
          const prevHas = prev.weeks.length > 0 || prev.inv_ytd || prev.trails_ytd || prev.w_indem_ytd || prev.w_non_ytd;
          if (!prevHas) brokers.splice(prevEmptyIdx, 1);
        }
        seenWithData.add(name);
      }

      brokers.push({
        name: rows[ri][0].toString().trim().replace(/\s+/g, ' '),
        team: b.team, inv_ytd, trails_ytd, w_indem_ytd, w_non_ytd, weeks,
      });
    }
  }
  return brokers;
}

main().catch(err => {
  console.error('ERROR:', err.message);
  process.exit(1);
});
