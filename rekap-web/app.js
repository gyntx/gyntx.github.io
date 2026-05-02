// ─── CONFIG ───────────────────────────────────────────────────────────────────
const ROOM_TYPES = { 'MG-V': 'MANGGAR', 'QN-V': 'QUEEN', 'MS-V': 'MAHLIGAI' };
const SEGMENT_MAP = {
  'FIT': 'FIT', 'CA': 'COMPANY', 'COMPANY': 'COMPANY',
  'GOV': 'GOVERNMENT', 'GOVERNMENT': 'GOVERNMENT',
  'WALK IN': 'WALK IN', 'OTA': 'OTA',
  'TA': 'TA', 'TRAVEL AGENT': 'TA',
  'COMPL': 'COMPLIMENT', 'COMPLIMENT': 'COMPLIMENT', 'COMPLIMENTARY': 'COMPLIMENT',
};
const SEGMENTS = ['FIT', 'COMPANY', 'GOVERNMENT', 'WALK IN', 'OTA', 'TA', 'COMPLIMENT'];
const DAYS_ORDER = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];

// ─── STATE ────────────────────────────────────────────────────────────────────
let uploadedFile = null;
let outputBlob = null;

// ─── DOM REFS ─────────────────────────────────────────────────────────────────
const dropZone     = document.getElementById('drop-zone');
const fileInput    = document.getElementById('file-input');
const configSec    = document.getElementById('config-section');
const progressSec  = document.getElementById('progress-section');
const resultSec    = document.getElementById('result-section');
const uploadSec    = document.getElementById('upload-section');
const progressFill = document.getElementById('progress-fill');
const progressText = document.getElementById('progress-text');
const logEl        = document.getElementById('log');
const summaryGrid  = document.getElementById('summary-grid');

// ─── UPLOAD HANDLERS ──────────────────────────────────────────────────────────
dropZone.addEventListener('dragover', e => { e.preventDefault(); dropZone.classList.add('dragover'); });
dropZone.addEventListener('dragleave', () => dropZone.classList.remove('dragover'));
dropZone.addEventListener('drop', e => {
  e.preventDefault();
  dropZone.classList.remove('dragover');
  handleFile(e.dataTransfer.files[0]);
});
// Klik drop zone → buka file picker (kecuali klik dari label agar tidak double trigger)
dropZone.addEventListener('click', e => {
  if (e.target.id === 'pick-btn') return; // ditangani sendiri
  fileInput.value = '';
  fileInput.click();
});
document.getElementById('pick-btn').addEventListener('click', e => {
  e.stopPropagation();
  fileInput.value = '';
  fileInput.click();
});
fileInput.addEventListener('change', () => {
  if (fileInput.files[0]) handleFile(fileInput.files[0]);
});

function handleFile(file) {
  if (!file || !file.name.endsWith('.xlsx')) {
    alert('Harap upload file .xlsx');
    return;
  }
  uploadedFile = file;

  // Auto-detect bulan & tahun dari nama file
  const monthNames = ['januari','februari','maret','april','mei','juni','juli','agustus','september','oktober','november','desember'];
  const match = file.name.match(/(\w+)[\s'`]*(\d{2,4})/i);
  if (match) {
    const year  = match[2].length === 2 ? '20' + match[2] : match[2];
    const mIdx  = monthNames.indexOf(match[1].toLowerCase());
    if (mIdx >= 0) {
      const mm = String(mIdx + 1).padStart(2, '0');
      document.getElementById('month-picker').value = `${year}-${mm}`;
    }
  }

  // Auto-detect total kamar dari isi file
  autoDetectTotalRooms(file).then(totalRooms => {
    document.getElementById('total-rooms').value = totalRooms;
    document.getElementById('file-name').textContent = file.name;
    document.getElementById('file-meta').textContent = `${(file.size / 1024).toFixed(1)} KB`;
    configSec.classList.remove('hidden');
    configSec.scrollIntoView({ behavior: 'smooth' });
  });
}

async function autoDetectTotalRooms(file) {
  try {
    const buffer = await file.arrayBuffer();
    const wb = XLSX.read(buffer, { type: 'array', sheetRows: 40 });
    const ws = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, raw: false });
    for (const r of rows) {
      const label = String(r[1] || '').trim();
      if (label.includes('T O T A L  ROOM AVAILABLE')) {
        const val = parseInt(String(r[3] || '').replace(/,/g, ''));
        if (val > 0) return val;
      }
    }
  } catch {}
  return 14;
}

// ─── PROCESS ──────────────────────────────────────────────────────────────────
document.getElementById('process-btn').addEventListener('click', async () => {
  const monthPicker = document.getElementById('month-picker').value; // format: "2026-03"
  const totalRooms  = parseInt(document.getElementById('total-rooms').value);

  if (!monthPicker) { alert('Pilih bulan & tahun terlebih dahulu'); return; }
  if (!totalRooms)  { alert('Isi total kamar tersedia'); return; }

  const [year, month] = monthPicker.split('-').map(Number);
  const monthNames = ['Januari','Februari','Maret','April','Mei','Juni','Juli','Agustus','September','Oktober','November','Desember'];
  const monthLabel = `${monthNames[month - 1]}-${year}`;

  configSec.classList.add('hidden');
  progressSec.classList.remove('hidden');
  logEl.innerHTML = '';

  try {
    const dailyData = await parseXlsx(uploadedFile, year, month);
    setProgress(60, 'Membuat file rekap...');

    const blob = await generateExcel(dailyData, monthLabel, totalRooms);
    outputBlob = blob;

    setProgress(100, 'Selesai!');
    showResult(dailyData, monthLabel);
  } catch (err) {
    addLog('❌ Error: ' + err.message, 'error');
    console.error(err);
  }
});

// ─── PARSE XLSX ───────────────────────────────────────────────────────────────
async function parseXlsx(file, year, month) {
  addLog('📂 Membaca file xlsx...');
  setProgress(10, 'Membaca file...');

  const buffer = await file.arrayBuffer();
  const wb = XLSX.read(buffer, { type: 'array', cellDates: true });

  addLog(`📋 Ditemukan ${wb.SheetNames.length} sheet`);
  const dailyData = [];

  wb.SheetNames.forEach((sheetName, idx) => {
    setProgress(10 + Math.round((idx / wb.SheetNames.length) * 45), `Parsing sheet: ${sheetName}`);

    // Ambil tanggal dari nama sheet
    const dayMatch = sheetName.match(/(\d+)/);
    if (!dayMatch) { addLog(`⚠️ Skip: ${sheetName} (tidak ada angka)`); return; }
    const day = parseInt(dayMatch[1]);
    let sheetDate;
    try { sheetDate = new Date(year, month - 1, day); } catch { return; }
    const dayOfWeek = DAYS_ORDER[sheetDate.getDay()];

    const ws = wb.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, raw: false });

    // Cari header row tamu
    let headerIdx = -1;
    for (let i = 0; i < rows.length; i++) {
      if (rows[i][0] === 'No' && rows[i][2] === 'Type Kamar') { headerIdx = i; break; }
    }

    const guests = [];
    if (headerIdx >= 0) {
      for (let i = headerIdx + 1; i < rows.length; i++) {
        const r = rows[i];
        if (!r[0] || !/^\d+$/.test(String(r[0]).trim())) continue;
        const roomType = String(r[2] || '').trim();
        const price    = parseFloat(String(r[7] || '0').replace(/,/g, '')) || 0;
        const segRaw   = String(r[13] || '').trim().toUpperCase();
        const seg      = SEGMENT_MAP[segRaw] || 'FIT';
        guests.push({
          roomNo:      String(r[1] || '').trim(),
          roomType,
          name:        String(r[3] || '').trim(),
          pax:         String(r[5] || '').trim(),
          nationality: String(r[6] || '').trim(),
          price,
          company:     String(r[9] || '').trim(),
          checkin:     String(r[10] || '').trim(),
          checkout:    String(r[11] || '').trim(),
          nights:      String(r[12] || '').trim(),
          segment: seg,
          keterangan:  String(r[14] || '').trim(),
          dayOfWeek,
        });
      }
    }

    // Parse summary
    const summary = { totalOccupied: 0, grossRoomRevenue: 0, rentalRevenue: 0, otherRoomRevenue: 0, fbRevenue: 0 };
    for (const r of rows) {
      const label = String(r[1] || '').trim();
      const val   = parseFloat(String(r[3] || '').replace(/,/g, '')) || 0;
      if (label.includes('T O T A L  OCCUPIED'))                                                                    summary.totalOccupied    = val;
      else if (label.includes('GROSS ROOM REVENUE') && !label.includes('AVERAGE') && !label.includes('NETT') && !label.includes('TOTAL GROSS')) summary.grossRoomRevenue = val;
      else if (label === 'RENTAL ITEM REVENUE')                                                                      summary.rentalRevenue    = val;
      else if (label === 'OTHER ROOM REVENUE')                                                                       summary.otherRoomRevenue = val;
      else if (label === 'FB REVENUE')                                                                               summary.fbRevenue        = val;
    }
    if (!summary.totalOccupied)    summary.totalOccupied    = guests.length;
    if (!summary.grossRoomRevenue) summary.grossRoomRevenue = guests.reduce((s, g) => s + g.price, 0);

    addLog(`✅ ${sheetName}: ${guests.length} tamu, Rev: ${fmt(summary.grossRoomRevenue)}`);
    dailyData.push({ date: sheetDate, sheetName, guests, summary });
  });

  dailyData.sort((a, b) => a.date - b.date);
  addLog(`📊 Total: ${dailyData.length} hari diproses`);
  return dailyData;
}

// ─── GENERATE EXCEL ───────────────────────────────────────────────────────────
async function generateExcel(dailyData, monthLabel, totalRooms) {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'Rekap Bulanan Hotel';

  buildOccRecap(wb, dailyData, monthLabel, totalRooms);
  buildRevenuePerType(wb, dailyData, monthLabel);
  buildSegmentasi(wb, dailyData, monthLabel);
  buildTamuRecap(wb, dailyData, monthLabel);

  const buffer = await wb.xlsx.writeBuffer();
  return new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
}

// ─── HELPERS STYLE ────────────────────────────────────────────────────────────
const thinBorder = { style: 'thin', color: { argb: 'FFB0B0B0' } };
const allBorders = { top: thinBorder, left: thinBorder, bottom: thinBorder, right: thinBorder };

function hdrCell(ws, row, col, value, bgArgb, fgArgb = 'FFFFFFFF', wrapText = true) {
  const cell = ws.getCell(row, col);
  cell.value = value;
  cell.fill  = { type: 'pattern', pattern: 'solid', fgColor: { argb: bgArgb } };
  cell.font  = { bold: true, color: { argb: fgArgb } };
  cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText };
  cell.border = allBorders;
}

function dataCell(ws, row, col, value, numFmt = null, align = 'center') {
  const cell = ws.getCell(row, col);
  cell.value  = value;
  cell.border = allBorders;
  cell.alignment = { horizontal: align, vertical: 'middle' };
  if (numFmt) cell.numFmt = numFmt;
}

function totalCell(ws, row, col, value, numFmt = null) {
  const cell = ws.getCell(row, col);
  cell.value = value;
  cell.fill  = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFC000' } };
  cell.font  = { bold: true };
  cell.border = allBorders;
  cell.alignment = { horizontal: 'center', vertical: 'middle' };
  if (numFmt) cell.numFmt = numFmt;
}

// ─── SHEET 1: OCC RECAP ───────────────────────────────────────────────────────
function buildOccRecap(wb, dailyData, monthLabel, totalRooms) {
  const ws = wb.addWorksheet('MBR OCC RECAPITULATION');

  // Title
  ws.mergeCells('A1:D2'); ws.getCell('A1').value = 'REKAPITULASI OKUPANSI';
  ws.getCell('A1').font = { bold: true, size: 14, color: { argb: 'FF1F4E79' } };
  ws.getCell('A1').alignment = { horizontal: 'left', vertical: 'middle' };

  ws.mergeCells('F1:J2'); ws.getCell('F1').value = monthLabel;
  ws.getCell('F1').font = { bold: true, size: 12 };
  ws.getCell('F1').alignment = { horizontal: 'center', vertical: 'middle' };

  // Header row 4
  const headers = ['No', 'Date', 'Room Sold', 'Total Room\nAvailable', 'Occ', 'Arr',
                   'Room Revenue', 'Other Room\nRevenue', 'Other Revenue', 'Total Revenue'];
  headers.forEach((h, i) => hdrCell(ws, 4, i + 1, h, 'FF1F4E79'));
  ws.getRow(4).height = 32;

  let [tSold, tAvail, tRev, tOtherRoom, tOther, tArrNum] = [0, 0, 0, 0, 0, 0];

  dailyData.forEach((day, idx) => {
    const r   = idx + 5;
    const s   = day.summary;
    const sold      = s.totalOccupied;
    const avail     = totalRooms;
    const roomRev   = s.grossRoomRevenue;
    const otherRoom = s.otherRoomRevenue;
    const other     = s.rentalRevenue + s.fbRevenue;
    const total     = roomRev + otherRoom + other;
    const occ       = avail ? sold / avail : 0;
    const arr       = sold  ? roomRev / sold : 0;

    const dateStr = day.date.toLocaleDateString('en-US', { weekday: 'long', day: '2-digit', month: 'long', year: 'numeric' });
    const rowBg   = idx % 2 === 1 ? 'FFD9E1F2' : 'FFFFFFFF';

    const vals = [idx + 1, dateStr, sold, avail, occ, arr, roomRev, otherRoom, other, total];
    vals.forEach((v, ci) => {
      const cell = ws.getCell(r, ci + 1);
      cell.value  = v;
      cell.border = allBorders;
      cell.fill   = { type: 'pattern', pattern: 'solid', fgColor: { argb: rowBg } };
      cell.alignment = { horizontal: ci === 1 ? 'left' : 'center', vertical: 'middle' };
      if (ci === 4) cell.numFmt = '0.00%';
      else if (ci >= 5) cell.numFmt = '#,##0';
    });

    tSold += sold; tAvail += avail; tRev += roomRev;
    tOtherRoom += otherRoom; tOther += other; tArrNum += roomRev;
  });

  // Total row
  const tr = dailyData.length + 5;
  const tOcc = tAvail ? tSold / tAvail : 0;
  const tArr = tSold  ? tArrNum / tSold : 0;
  ['', 'TOTAL', tSold, tAvail, tOcc, tArr, tRev, tOtherRoom, tOther, tRev + tOtherRoom + tOther]
    .forEach((v, ci) => {
      totalCell(ws, tr, ci + 1, v, ci === 4 ? '0.00%' : ci >= 5 ? '#,##0' : null);
    });

  // Col widths
  [5, 30, 10, 12, 8, 14, 16, 16, 14, 16].forEach((w, i) => { ws.getColumn(i + 1).width = w; });
}

// ─── SHEET 2: REVENUE PER TIPE KAMAR ─────────────────────────────────────────
function buildRevenuePerType(wb, dailyData, monthLabel) {
  const ws = wb.addWorksheet('REVENUE PER TIPE KAMAR');

  const headers = ['TANGGAL', 'MANGGAR QTY', 'MANGGAR REVENUE', 'QUEEN QTY', 'QUEEN REVENUE',
                   'MAHLIGAI QTY', 'MAHLIGAI REVENUE', 'TOTAL ROOM REVENUE\n(By Room Type)'];
  headers.forEach((h, i) => hdrCell(ws, 1, i + 1, h, 'FF1F4E79'));
  ws.getRow(1).height = 36;

  const totals = { 'MG-V': [0, 0], 'QN-V': [0, 0], 'MS-V': [0, 0] };

  dailyData.forEach((day, idx) => {
    const r   = idx + 2;
    const qty = { 'MG-V': 0, 'QN-V': 0, 'MS-V': 0 };
    const rev = { 'MG-V': 0, 'QN-V': 0, 'MS-V': 0 };

    day.guests.forEach(g => {
      if (qty[g.roomType] !== undefined) { qty[g.roomType]++; rev[g.roomType] += g.price; }
    });

    const totalRev = Object.values(rev).reduce((s, v) => s + v, 0);
    const rowBg    = idx % 2 === 1 ? 'FFDDEEFF' : 'FFFFFFFF';

    const vals = [day.date, qty['MG-V'], rev['MG-V'], qty['QN-V'], rev['QN-V'], qty['MS-V'], rev['MS-V'], totalRev];
    vals.forEach((v, ci) => {
      const cell = ws.getCell(r, ci + 1);
      cell.value  = v;
      cell.border = allBorders;
      cell.fill   = { type: 'pattern', pattern: 'solid', fgColor: { argb: rowBg } };
      cell.alignment = { horizontal: ci === 0 ? 'left' : 'center', vertical: 'middle' };
      if (ci === 0) cell.numFmt = 'D-MMM';
      else if ([2, 4, 6, 7].includes(ci)) cell.numFmt = '#,##0';
    });

    Object.keys(totals).forEach(rt => { totals[rt][0] += qty[rt]; totals[rt][1] += rev[rt]; });
  });

  // Total row
  const tr = dailyData.length + 2;
  const totalRevAll = Object.values(totals).reduce((s, v) => s + v[1], 0);
  ['TOTAL', totals['MG-V'][0], totals['MG-V'][1], totals['QN-V'][0], totals['QN-V'][1],
   totals['MS-V'][0], totals['MS-V'][1], totalRevAll]
    .forEach((v, ci) => {
      const cell = ws.getCell(tr, ci + 1);
      cell.value = v;
      cell.fill  = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F4E79' } };
      cell.font  = { bold: true, color: { argb: 'FFFFFFFF' } };
      cell.border = allBorders;
      cell.alignment = { horizontal: 'center', vertical: 'middle' };
      if ([2, 4, 6, 7].includes(ci)) cell.numFmt = '#,##0';
    });

  [12, 12, 18, 10, 16, 12, 18, 22].forEach((w, i) => { ws.getColumn(i + 1).width = w; });
}

// ─── SHEET 3: SEGMENTASI ──────────────────────────────────────────────────────
function buildSegmentasi(wb, dailyData, monthLabel) {
  const ws = wb.addWorksheet('SEGMENTASI');

  // Aggregate
  const qtyByDay = {};
  const revByDay = {};
  SEGMENTS.forEach(s => {
    qtyByDay[s] = {}; revByDay[s] = {};
    DAYS_ORDER.forEach(d => { qtyByDay[s][d] = 0; revByDay[s][d] = 0; });
  });

  dailyData.forEach(day => {
    day.guests.forEach(g => {
      const seg = qtyByDay[g.segment] ? g.segment : 'FIT';
      if (DAYS_ORDER.includes(g.dayOfWeek)) {
        qtyByDay[seg][g.dayOfWeek]++;
        revByDay[seg][g.dayOfWeek] += g.price;
      }
    });
  });

  // Title
  ws.mergeCells('C3:J3');
  ws.getCell('C3').value = 'Market Segmentasi';
  ws.getCell('C3').font  = { bold: true, size: 12 };
  ws.getCell('K3').value = monthLabel;
  ws.getCell('K3').font  = { bold: true };
  ws.getCell('K3').alignment = { horizontal: 'right' };

  function writeTable(startRow, bgArgb, isArr, srcDict) {
    hdrCell(ws, startRow, 3, 'SEGMENTASI', bgArgb, 'FF000000');
    DAYS_ORDER.forEach((d, i) => hdrCell(ws, startRow, 4 + i, d, bgArgb, 'FF000000'));
    hdrCell(ws, startRow, 11, 'Total', bgArgb, 'FF000000');

    SEGMENTS.forEach((seg, si) => {
      const r = startRow + 1 + si;
      const cell = ws.getCell(r, 3);
      cell.value  = seg;
      cell.font   = { color: { argb: ['FIT', 'OTA'].includes(seg) ? 'FFC00000' : 'FF000000' } };
      cell.border = allBorders;

      let rowTotal = 0;
      DAYS_ORDER.forEach((day, di) => {
        let v;
        if (isArr) {
          const q = qtyByDay[seg][day], rv = revByDay[seg][day];
          v = q ? rv / q : '-';
        } else {
          v = srcDict[seg][day] || null;
        }
        dataCell(ws, r, 4 + di, v, typeof v === 'number' ? '#,##0' : null);
        if (typeof v === 'number' && !isArr) rowTotal += v;
      });

      let kVal;
      if (isArr) {
        const tq = DAYS_ORDER.reduce((s, d) => s + qtyByDay[seg][d], 0);
        const tr = DAYS_ORDER.reduce((s, d) => s + revByDay[seg][d], 0);
        kVal = tq ? tr / tq : '-';
      } else {
        kVal = rowTotal || null;
      }
      dataCell(ws, r, 11, kVal, typeof kVal === 'number' ? '#,##0' : null);
    });

    const totalR = startRow + SEGMENTS.length + 1;
    const tcell  = ws.getCell(totalR, 3);
    tcell.value  = 'Total';
    tcell.font   = { color: { argb: 'FF0070C0' }, bold: true };
    tcell.border = allBorders;

    let grandTotalQ = 0, grandTotalRv = 0, grandTotal = 0;
    DAYS_ORDER.forEach((day, di) => {
      let v;
      if (isArr) {
        const tq = SEGMENTS.reduce((s, sg) => s + qtyByDay[sg][day], 0);
        const tr = SEGMENTS.reduce((s, sg) => s + revByDay[sg][day], 0);
        v = tq ? tr / tq : '-';
        grandTotalQ += tq; grandTotalRv += tr;
      } else {
        v = SEGMENTS.reduce((s, sg) => s + srcDict[sg][day], 0) || null;
        if (typeof v === 'number') grandTotal += v;
      }
      dataCell(ws, totalR, 4 + di, v, typeof v === 'number' ? '#,##0' : null);
    });

    const gt = isArr ? (grandTotalQ ? grandTotalRv / grandTotalQ : null) : (grandTotal || null);
    dataCell(ws, totalR, 11, gt, gt ? '#,##0' : null);
  }

  writeTable(5,  'FFFFC000', false, qtyByDay);  // QTY
  writeTable(15, 'FF92D050', false, revByDay);  // Revenue
  writeTable(25, 'FFBDD7EE', true,  revByDay);  // ARR

  [3, 3, 14, 10, 10, 10, 12, 12, 10, 10, 14].forEach((w, i) => { ws.getColumn(i + 1).width = w; });
}

// ─── RESULT & DOWNLOAD ────────────────────────────────────────────────────────
// ─── SHEET 4: REKAPITULASI TAMU ─────────────────────────────────────────────
function buildTamuRecap(wb, dailyData, monthLabel) {
  const ws = wb.addWorksheet('REKAPITULASI TAMU');

  ws.mergeCells('A1:F1'); ws.getCell('A1').value = 'REKAPITULASI TAMU MENGINAP';
  ws.getCell('A1').font = { bold: true, size: 13 };
  ws.getCell('A1').alignment = { horizontal: 'left', vertical: 'middle' };
  ws.mergeCells('H1:M1'); ws.getCell('H1').value = monthLabel.toUpperCase();
  ws.getCell('H1').font = { bold: true, size: 12 };
  ws.getCell('H1').alignment = { horizontal: 'center', vertical: 'middle' };
  ws.getRow(1).height = 20;

  const headers = ['No', 'Nomor\nKamar', 'Type\nKamar', 'Nama Tamu', 'Dewasa\n/Anak',
                   'Kebangsaan', 'Harga Kamar', 'Perusahaan/Travel',
                   'CheckIn', 'CheckOut', 'Lama\nMenginap', 'Segmentasi', 'Keterangan'];
  headers.forEach((h, i) => hdrCell(ws, 3, i + 1, h, 'FF1F4E79'));
  ws.getRow(3).height = 30;

  const fmtDate = s => {
    try {
      const d = new Date(s);
      if (isNaN(d)) return s;
      return `${String(d.getDate()).padStart(2,'0')}/${String(d.getMonth()+1).padStart(2,'0')}/${String(d.getFullYear()).slice(-2)}`;
    } catch { return s; }
  };

  let no = 0, totalHarga = 0;
  dailyData.forEach(day => {
    day.guests.forEach(g => {
      no++;
      const row = 3 + no;
      const bg  = no % 2 === 0 ? 'FFF7FAFC' : 'FFFFFFFF';
      const vals = [
        no, g.roomNo, g.roomType, g.name, g.pax,
        g.nationality, g.price, g.company,
        fmtDate(g.checkin), fmtDate(g.checkout),
        g.nights, g.segment, g.keterangan
      ];
      vals.forEach((v, ci) => {
        const cell = ws.getCell(row, ci + 1);
        cell.value  = v;
        cell.border = allBorders;
        cell.fill   = { type: 'pattern', pattern: 'solid', fgColor: { argb: bg } };
        cell.alignment = { horizontal: [3, 7, 12].includes(ci) ? 'left' : 'center', vertical: 'middle' };
        if (ci === 6) cell.numFmt = '#,##0';
      });
      totalHarga += g.price;
    });
  });

  const tr = 3 + no + 1;
  ws.mergeCells(`A${tr}:F${tr}`);
  ws.getCell(tr, 1).value = 'TOTAL';
  ws.getCell(tr, 1).font  = { bold: true };
  ws.getCell(tr, 1).alignment = { horizontal: 'center' };
  ws.getCell(tr, 7).value  = totalHarga;
  ws.getCell(tr, 7).numFmt = '#,##0';
  ws.getCell(tr, 7).font   = { bold: true };
  for (let c = 1; c <= 13; c++) {
    ws.getCell(tr, c).border = allBorders;
    ws.getCell(tr, c).fill   = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFC000' } };
  }

  [5, 8, 8, 28, 8, 10, 14, 20, 10, 10, 8, 12, 22].forEach((w, i) => { ws.getColumn(i + 1).width = w; });
}

function showResult(dailyData, monthLabel) {
  progressSec.classList.add('hidden');
  resultSec.classList.remove('hidden');

  const totalGuests  = dailyData.reduce((s, d) => s + d.guests.length, 0);
  const totalRevenue = dailyData.reduce((s, d) => s + d.summary.grossRoomRevenue, 0);
  const avgOcc       = dailyData.reduce((s, d) => s + d.summary.totalOccupied, 0) /
                       (dailyData.length * parseInt(document.getElementById('total-rooms').value));

  summaryGrid.innerHTML = `
    <div class="summary-card"><div class="value">${dailyData.length}</div><div class="label">Hari Operasional</div></div>
    <div class="summary-card"><div class="value">${totalGuests}</div><div class="label">Total Tamu</div></div>
    <div class="summary-card"><div class="value">${(avgOcc * 100).toFixed(1)}%</div><div class="label">Rata-rata OCC</div></div>
    <div class="summary-card"><div class="value">Rp ${fmt(totalRevenue)}</div><div class="label">Total Room Revenue</div></div>
    <div class="summary-card"><div class="value">4</div><div class="label">Sheet Dibuat</div></div>
    <div class="summary-card"><div class="value">1</div><div class="label">File Output</div></div>
  `;

  document.getElementById('download-btn').onclick = () => {
    const url  = URL.createObjectURL(outputBlob);
    const a    = document.createElement('a');
    a.href     = url;
    a.download = `Rekap Bulanan ${monthLabel}.xlsx`;
    a.click();
    URL.revokeObjectURL(url);
  };
}

document.getElementById('reset-btn').addEventListener('click', () => {
  uploadedFile = null; outputBlob = null;
  fileInput.value = '';
  resultSec.classList.add('hidden');
  progressSec.classList.add('hidden');
  configSec.classList.add('hidden');
  uploadSec.scrollIntoView({ behavior: 'smooth' });
});

// ─── UTILS ────────────────────────────────────────────────────────────────────
function setProgress(pct, text) {
  progressFill.style.width = pct + '%';
  progressText.textContent = text;
}

function addLog(msg, type = 'info') {
  const line = document.createElement('div');
  line.textContent = msg;
  if (type === 'error') line.style.color = '#fc8181';
  logEl.appendChild(line);
  logEl.scrollTop = logEl.scrollHeight;
}

function fmt(n) {
  return Math.round(n).toLocaleString('id-ID');
}
