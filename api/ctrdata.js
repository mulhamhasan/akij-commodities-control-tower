const XLSX = require('xlsx');

const REGISTER_ID = "1me7vjHVAMV6Wazw4GnIzoPvvb3B_Rhz1CyDUY9FGLh4";

const COMMODITY_ORDER = [
  'Wheat', 'Lentil', 'SBM / Soybean Meal', 'Corn', 'Maize', 'Soyabean',
  'Mustard', 'Yellow Peas', 'Chickpeas', 'Canola', 'Coal', 'General / Macro',
  'Palm Oil', 'Jute', 'Lysine/Threonine', 'Rice Bran', 'Groundnut', 'Turmeric'
];

const RULES = [
  ['lysine', 'Lysine/Threonine'], ['threonine', 'Lysine/Threonine'],
  ['methionine', 'Lysine/Threonine'], ['choline', 'Lysine/Threonine'],
  ['monocalcium', 'Lysine/Threonine'], ['limestone', 'Lysine/Threonine'],
  ['sodium bicarbonate', 'Lysine/Threonine'], ['dcp', 'Lysine/Threonine'],
  ['rice br', 'Rice Bran'], ['dorb', 'Rice Bran'],
  ['turmeric', 'Turmeric'], ['groundnut', 'Groundnut'], ['jute', 'Jute'],
  ['palm', 'Palm Oil'], ['coal', 'Coal'],
  ['canola', 'Canola'], ['rapeseed', 'Canola'],
  ['chickpea', 'Chickpeas'],
  ['yellowpea', 'Yellow Peas'], ['yellow pea', 'Yellow Peas'],
  ['peas', 'Yellow Peas'], ['pulse', 'Yellow Peas'],
  ['mustard', 'Mustard'], ['lentil', 'Lentil'],
  ['soya bean meal', 'SBM / Soybean Meal'], ['soya meal', 'SBM / Soybean Meal'],
  ['soybean meal', 'SBM / Soybean Meal'], ['soy meal', 'SBM / Soybean Meal'],
  ['sbm', 'SBM / Soybean Meal'],
  ['soybean extraction', 'Soyabean'], ['soyabean extraction', 'Soyabean'],
  ['soybean', 'Soyabean'], ['soyabean', 'Soyabean'],
  ['corn', 'Corn'], ['wheat', 'Wheat'], ['cwrs', 'Wheat'], ['maize', 'Maize']
];

function classify(item) {
  const s = ' ' + item.toLowerCase() + ' ';
  for (const [k, com] of RULES) {
    if (s.includes(k)) return com;
  }
  return 'General / Macro';
}

function cellAt(ws, row, col) {
  return ws[XLSX.utils.encode_cell({ r: row - 1, c: col - 1 })];
}

function cellValue(ws, row, col) {
  const cell = cellAt(ws, row, col);
  return cell ? cell.v : null;
}

function cellHref(ws, row, col) {
  const cell = cellAt(ws, row, col);
  if (cell && cell.l) {
    if (cell.l.Rel && cell.l.Rel.Target) return cell.l.Rel.Target;
    if (cell.l.Target) return cell.l.Target;
  }
  return null;
}

function extractSheet(ws, category) {
  const range = XLSX.utils.decode_range(ws['!ref']);
  const maxCol = range.e.c + 1;
  const maxRow = range.e.r + 1;

  let hdr = -1;
  for (let r = 1; r <= Math.min(maxRow, 10); r++) {
    const a = String(cellValue(ws, r, 1) || '');
    const c = String(cellValue(ws, r, 3) || '');
    if (a.includes('Report Type') || c.includes('Sub')) { hdr = r; break; }
  }
  if (hdr < 0) return [];

  const mcols = {};
  for (let c = 1; c <= maxCol; c++) {
    const v = cellValue(ws, hdr, c);
    if (v && /meeting/i.test(String(v))) {
      const m = String(v).match(/(\d+)/);
      if (m) mcols[parseInt(m[1], 10)] = c;
    }
  }
  const sortedNums = Object.keys(mcols).map(Number).sort((a, b) => a - b);

  const rows = [];
  let curGroup = null;
  for (let r = hdr + 1; r <= maxRow; r++) {
    const g = cellValue(ws, r, 1);
    if (g && String(g).trim()) curGroup = String(g).trim();
    const item = cellValue(ws, r, 3);
    if (!item || !String(item).trim()) continue;
    const timeline = cellValue(ws, r, 4);

    let link = null;
    for (let c = 5; c <= maxCol; c++) {
      const h = cellHref(ws, r, c);
      if (h) { link = h; break; }
    }
    const meetings = {};
    let cur = link;
    for (const n of sortedNums) {
      const h = cellHref(ws, r, mcols[n]);
      if (h) cur = h;
      meetings[String(n)] = cur;
    }

    rows.push({
      category,
      group: curGroup,
      item: String(item).trim(),
      timeline: timeline ? String(timeline).trim() : '',
      link,
      meetings
    });
  }
  return rows;
}

async function build() {
  const tokenResp = await fetch('https://oauth2.googleapis.com/token', {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({
      client_id: process.env.GOOGLE_CLIENT_ID,
      client_secret: process.env.GOOGLE_CLIENT_SECRET,
      refresh_token: process.env.GOOGLE_REFRESH_TOKEN,
      grant_type: 'refresh_token'
    })
  });
  const tokenJson = await tokenResp.json();
  if (!tokenJson.access_token) throw new Error('token refresh failed: ' + JSON.stringify(tokenJson));
  const token = tokenJson.access_token;

  const xlsxResp = await fetch(
    `https://www.googleapis.com/drive/v3/files/${REGISTER_ID}/export?mimeType=application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`,
    { headers: { Authorization: `Bearer ${token}` } }
  );
  if (!xlsxResp.ok) throw new Error('Drive export failed: ' + xlsxResp.status);
  const buf = await xlsxResp.arrayBuffer();
  const wb = XLSX.read(buf, { cellHyperlinks: true });

  const allRows = [];
  for (const name of wb.SheetNames) {
    const t = name.toLowerCase();
    let cat = null;
    if (t.includes('supply')) cat = 'supply';
    else if (t.includes('demand')) cat = 'demand';
    else if (t.includes('finance')) cat = 'finance';
    if (cat) allRows.push(...extractSheet(wb.Sheets[name], cat));
  }

  const data = {};
  for (const row of allRows) {
    const com = classify(row.item);
    if (!data[com]) data[com] = { supply: [], demand: [], finance: [] };
    data[com][row.category].push({
      item: row.item,
      timeline: row.timeline,
      link: row.link,
      group: row.group,
      meetings: row.meetings
    });
  }

  const ordered = {};
  for (const k of COMMODITY_ORDER) if (data[k]) ordered[k] = data[k];
  for (const k in data) if (!ordered[k]) ordered[k] = data[k];

  const nums = new Set();
  for (const name of wb.SheetNames) {
    const t = name.toLowerCase();
    if (!(t.includes('supply') || t.includes('demand') || t.includes('finance'))) continue;
    const ws = wb.Sheets[name];
    const range = XLSX.utils.decode_range(ws['!ref']);
    for (let r = 1; r <= Math.min(range.e.r + 1, 10); r++) {
      for (let c = 1; c <= range.e.c + 1; c++) {
        const v = cellValue(ws, r, c);
        if (v && /meeting/i.test(String(v))) {
          const m = String(v).match(/(\d+)/);
          if (m) nums.add(parseInt(m[1], 10));
        }
      }
    }
  }
  const sortedMeetingNums = [...nums].sort((a, b) => a - b);
  const meetings = sortedMeetingNums.map((n, i) => {
    const ord = n === 1 ? 'st' : n === 2 ? 'nd' : n === 3 ? 'rd' : 'th';
    return { n, label: `${n}${ord} Meeting`, status: i === sortedMeetingNums.length - 1 ? 'latest' : 'past' };
  });

  return '// Auto-generated from "Source File" register.\n' +
    'window.MEETINGS = ' + JSON.stringify(meetings) + ';\n' +
    'window.CTR_DATA = ' + JSON.stringify(ordered) + ';\n';
}

module.exports = async function handler(req, res) {
  try {
    const js = await build();
    res.setHeader('Content-Type', 'application/javascript; charset=utf-8');
    res.setHeader('Cache-Control', 'no-store');
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.status(200).send(js);
  } catch (e) {
    res.setHeader('Content-Type', 'text/plain; charset=utf-8');
    res.status(500).send('Error: ' + e.message);
  }
};
