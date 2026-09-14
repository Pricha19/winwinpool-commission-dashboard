/* ==========================================================================
   WinWinPool Commission Dashboard — client logic
   All Excel parsing + business-rule computation happens in the browser.
   Only pre-aggregated results are sent to /api/upload.
   ========================================================================== */

/* ---------- constants ---------- */
const BRANCH_CODES = ['ไม่ระบุ', 'สำนักงานใหญ่ (กรุงเทพฯ)', 'สาขาหัวหิน', 'สาขาพัทยา']; // index 0-3
const CAT_CODES = ['ในประเทศ', 'นำเข้าเอง', 'บริการ']; // index 0-2, fixed display order
const CAT_KEYS = ['domestic', 'import', 'service'];
const STATUS_LIST = [
  'ตรงกัน (ยืนยันแล้ว)',
  'ไม่ตรง (ยังไม่ยืนยันการชำระเงิน)',
  'ไม่นับ (รอตรวจสอบ - ฐานภาษีไม่ตรงกับยอดหลังหักส่วนลด)',
  'ไม่นับ (ใบสำคัญเป็น Draft ยังไม่อนุมัติ)',
  'ไม่นับ (ใบสำคัญถูกยกเลิก)',
];
const CONFIRMED_STATUS = 'ตรงกัน (ยืนยันแล้ว)';
const IMPORT_TOKENS = ['-WW-', '-WP-', '-LS-', '-PT-', '-ADS-'];

/* ---------- date helpers (Excel serial + พ.ศ./ค.ศ. fix) ---------- */
function excelSerialToDate(serial) {
  return new Date(Math.round((serial - 25569) * 86400 * 1000)); // UTC
}
function normalizeYear(y) {
  return y > 2100 ? y - 543 : y; // > 2100 => พ.ศ. leaked into a ค.ศ. date cell
}
function yearMonthFromCell(rawDate) {
  let d = null;
  if (typeof rawDate === 'number' && isFinite(rawDate)) {
    d = excelSerialToDate(rawDate);
  } else if (rawDate instanceof Date) {
    d = rawDate;
  } else if (typeof rawDate === 'string' && rawDate.trim()) {
    // fallback: try dd/mm/yyyy style strings
    const m = rawDate.trim().match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})/);
    if (m) {
      let yy = parseInt(m[3], 10);
      if (yy < 100) yy += 2500; // 2-digit พ.ศ. shorthand, rare
      const yFixed = normalizeYear(yy);
      const mm = String(parseInt(m[2], 10)).padStart(2, '0');
      return `${yFixed}-${mm}`;
    }
    return 'unknown';
  }
  if (!d || isNaN(d.getTime())) return 'unknown';
  const y = normalizeYear(d.getUTCFullYear());
  const mm = String(d.getUTCMonth() + 1).padStart(2, '0');
  return `${y}-${mm}`;
}

/* ---------- product categorisation ---------- */
function categorizeProduct(code) {
  const c = '-' + String(code || '').toUpperCase() + '-';
  if (c.includes('-SV-') || String(code || '').toUpperCase().startsWith('SV-')) return 2; // service, checked first
  for (const tok of IMPORT_TOKENS) {
    if (c.includes(tok)) return 1; // self-imported
  }
  return 0; // domestic resale (default)
}

/* ---------- branch derivation from salesCode ---------- */
function branchFromSalesCode(code) {
  const first = String(code || '').trim().charAt(0).toUpperCase();
  if (first === 'B') return 1;
  if (first === 'H') return 2;
  if (first === 'P') return 3;
  return 0;
}

/* ---------- generic header-row scan (label-based, not fixed position) ---------- */
function findHeaderRowIndex(rows, colCLabel) {
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    if (row && row[0] === 'No' && row[2] === colCLabel) return i;
  }
  return -1;
}
function buildHeaderMap(headerRow) {
  const map = {};
  headerRow.forEach((label, idx) => {
    if (label !== null && label !== undefined && label !== '') {
      const key = String(label).trim();
      if (!(key in map)) map[key] = idx;
    }
  });
  return map;
}

/* ---------- parse an INV or DEP sheet into a list of "documents" ----------
   Each document = { invoiceNo, ym, apr, description, staffName, docStatus,
                      totalBeforeDiscount, totalDiscount, source, items:[...] }
   Item = { code, qty, unitPrice, lineSubtotal, salesCode }
   Item-row columns are FIXED (same in both INV and DEP):
     E(4)=code, G(6)=qty, I(8)=unitPrice, J(9)=lineSubtotal, P(15)=salesCode
--------------------------------------------------------------------------- */
function parseInvoiceLikeSheet(rows, source) {
  const headerIdx = findHeaderRowIndex(rows, 'เลขใบสำคัญ');
  if (headerIdx === -1) return null;
  const hmap = buildHeaderMap(rows[headerIdx]);
  const need = ['Apr', 'เลขใบสำคัญ', 'วันที่', 'คำอธิบาย', 'พนักงาน', 'สถานะ', 'ราคาสินค้า', 'ส่วนลด', 'เลขใบกำกับภาษี'];
  for (const label of need) {
    if (!(label in hmap)) return null; // not a recognisable sheet of this kind
  }
  const docs = [];
  let current = null;
  for (let r = headerIdx + 1; r < rows.length; r++) {
    const row = rows[r];
    if (!row) continue;
    const colA = row[0];
    if (colA !== null && colA !== undefined && colA !== '') {
      if (current) docs.push(current);
      const totalBeforeDiscount = Number(row[hmap['ราคาสินค้า']]) || 0;
      const totalDiscount = Number(row[hmap['ส่วนลด']]) || 0;
      current = {
        source,
        invoiceNo: row[hmap['เลขใบกำกับภาษี']] != null ? String(row[hmap['เลขใบกำกับภาษี']]).trim() : '',
        apr: row[hmap['Apr']] != null ? String(row[hmap['Apr']]).trim() : '',
        description: row[hmap['คำอธิบาย']] != null ? String(row[hmap['คำอธิบาย']]) : '',
        staffName: row[hmap['พนักงาน']] != null ? String(row[hmap['พนักงาน']]).trim() : '',
        docStatus: row[hmap['สถานะ']] != null ? String(row[hmap['สถานะ']]) : '',
        totalBeforeDiscount,
        totalDiscount,
        ym: yearMonthFromCell(row[hmap['วันที่']]),
        items: [],
      };
    } else if (current && row[4] !== null && row[4] !== undefined && row[4] !== '') {
      current.items.push({
        code: String(row[4]).trim(),
        qty: Number(row[6]) || 0,
        unitPrice: Number(row[8]) || 0,
        lineSubtotal: Number(row[9]) || 0,
        salesCode: row[15] != null ? String(row[15]).trim() : '',
      });
    }
  }
  if (current) docs.push(current);
  return docs;
}

/* ---------- parse the TAX sheet -> Map(invoiceNo -> summed ฐานภาษี) ---------- */
function parseTaxSheet(rows) {
  const headerIdx = findHeaderRowIndex(rows, 'เลขใบกำกับภาษี');
  if (headerIdx === -1) return null;
  const hmap = buildHeaderMap(rows[headerIdx]);
  if (!('เลขใบกำกับภาษี' in hmap) || !('ฐานภาษี' in hmap)) return null;
  const map = new Map();
  for (let r = headerIdx + 1; r < rows.length; r++) {
    const row = rows[r];
    if (!row) continue;
    const inv = row[hmap['เลขใบกำกับภาษี']];
    if (inv === null || inv === undefined || inv === '') continue;
    const invoiceNo = String(inv).trim();
    const base = Number(row[hmap['ฐานภาษี']]) || 0;
    map.set(invoiceNo, (map.get(invoiceNo) || 0) + base);
  }
  return map;
}

/* ---------- try every sheet in a workbook until one matches a known shape ---------- */
function extractRowsFromWorkbook(workbook, tryFns) {
  for (const sheetName of workbook.SheetNames) {
    const ws = workbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
    for (const fn of tryFns) {
      const result = fn(rows);
      if (result) return result;
    }
  }
  return null;
}

/* ==========================================================================
   Status determination + net-sales computation (per business spec 3.3 / 3.4)
   Produces a flat array of "flat rows":
     { ym, source, status, branch(0-3), cat(0-2), salesCode, staffName, netSales }
   netSales is 0 for any row whose status !== CONFIRMED_STATUS.
   ========================================================================== */

function determineBaseStatus(doc, taxMap) {
  const descAndStatus = (doc.description || '') + ' ' + (doc.docStatus || '');
  if (descAndStatus.includes('ยกเลิก')) return 'ไม่นับ (ใบสำคัญถูกยกเลิก)';
  if (doc.apr === 'Draft') return 'ไม่นับ (ใบสำคัญเป็น Draft ยังไม่อนุมัติ)';
  const taxBase = doc.invoiceNo ? taxMap.get(doc.invoiceNo) : undefined;
  if (taxBase === undefined || taxBase === null || taxBase === 0) {
    return 'ไม่ตรง (ยังไม่ยืนยันการชำระเงิน)';
  }
  return null; // needs further per-source checks
}

/* ---- INV: flatten one document into flat item rows ---- */
function flattenInvDoc(doc, taxMap) {
  const base = determineBaseStatus(doc, taxMap);
  const out = [];
  if (base) {
    for (const item of doc.items) {
      out.push(makeFlatRow(doc, item, base, 0));
    }
    return out;
  }
  const taxBase = taxMap.get(doc.invoiceNo) || 0;
  const diff = Math.round((doc.totalBeforeDiscount - doc.totalDiscount - taxBase) * 100) / 100;
  let status;
  if (diff !== 0) {
    status = 'ไม่นับ (รอตรวจสอบ - ฐานภาษีไม่ตรงกับยอดหลังหักส่วนลด)';
  } else {
    status = CONFIRMED_STATUS;
  }
  for (const item of doc.items) {
    let netSales = 0;
    if (status === CONFIRMED_STATUS) {
      netSales = doc.totalBeforeDiscount === 0
        ? item.lineSubtotal
        : item.lineSubtotal * (1 - doc.totalDiscount / doc.totalBeforeDiscount);
    }
    out.push(makeFlatRow(doc, item, status, netSales));
  }
  return out;
}

/* ---- DEP: flatten one document (deposit allocation across the whole order) ---- */
function flattenDepDoc(doc, taxMap) {
  const base = determineBaseStatus(doc, taxMap);
  const out = [];
  if (base) {
    for (const item of doc.items) {
      out.push(makeFlatRow(doc, item, base, 0));
    }
    return out;
  }
  const taxBase = taxMap.get(doc.invoiceNo) || 0;
  const status = CONFIRMED_STATUS;
  const isExact = Math.round((doc.totalBeforeDiscount - doc.totalDiscount - taxBase) * 100) / 100 === 0;

  if (isExact) {
    for (const item of doc.items) {
      const netSales = doc.totalBeforeDiscount === 0
        ? item.lineSubtotal
        : item.lineSubtotal * (1 - doc.totalDiscount / doc.totalBeforeDiscount);
      out.push(makeFlatRow(doc, item, status, netSales));
    }
    return out;
  }

  // partial deposit: SV- rows paid first, remainder split evenly among the rest
  const svItems = doc.items.filter((it) => categorizeProduct(it.code) === 2);
  const otherItems = doc.items.filter((it) => categorizeProduct(it.code) !== 2);
  const svTotalRaw = svItems.reduce((s, it) => s + it.lineSubtotal, 0);
  const svRecognized = Math.min(svTotalRaw, taxBase);
  const remaining = taxBase - svRecognized;

  for (const it of svItems) {
    const netSales = svTotalRaw > 0 ? (it.lineSubtotal / svTotalRaw) * svRecognized : 0;
    out.push(makeFlatRow(doc, it, status, netSales));
  }
  if (otherItems.length === 1) {
    out.push(makeFlatRow(doc, otherItems[0], status, remaining));
  } else if (otherItems.length > 1) {
    const share = remaining / otherItems.length;
    for (const it of otherItems) {
      out.push(makeFlatRow(doc, it, status, share));
    }
  }
  return out;
}

function makeFlatRow(doc, item, status, netSales) {
  return {
    ym: doc.ym,
    source: doc.source,
    status,
    branch: branchFromSalesCode(item.salesCode),
    cat: categorizeProduct(item.code),
    salesCode: item.salesCode || '',
    staffName: doc.staffName || '',
    netSales: status === CONFIRMED_STATUS ? netSales : 0,
  };
}

/* ---- run the full pipeline over parsed INV + DEP docs ---- */
function computeFlatRows(invDocs, depDocs, taxMap) {
  const rows = [];
  for (const doc of invDocs) rows.push(...flattenInvDoc(doc, taxMap));
  for (const doc of depDocs) rows.push(...flattenDepDoc(doc, taxMap));
  return rows;
}

/* ==========================================================================
   Month bucketing, aggregation per month, and sales-name directory
   ========================================================================== */

const THAI_MONTHS = ['มกราคม','กุมภาพันธ์','มีนาคม','เมษายน','พฤษภาคม','มิถุนายน',
  'กรกฎาคม','สิงหาคม','กันยายน','ตุลาคม','พฤศจิกายน','ธันวาคม'];

function thaiMonthLabel(ym) {
  if (ym === 'unknown') return 'ไม่ทราบวันที่';
  const [y, m] = ym.split('-').map(Number);
  const buddhistYear = y + 543;
  return `${THAI_MONTHS[m - 1] || m} ${buddhistYear}`;
}

function emptyStatusCounts() {
  const o = {};
  for (const s of STATUS_LIST) o[s] = 0;
  return o;
}

function aggregateFlatRowsForMonth(rows, ym) {
  const totals = { import: 0, domestic: 0, service: 0, grand: 0 };
  const statusCounts = emptyStatusCounts();
  const sourceConfirmed = { INV: 0, DEP: 0 };
  const sourceNet = { INV: 0, DEP: 0 };
  const groupMap = new Map();
  let confirmedRows = 0;

  for (const r of rows) {
    statusCounts[r.status] = (statusCounts[r.status] || 0) + 1;
    if (r.status === CONFIRMED_STATUS) {
      confirmedRows++;
      sourceConfirmed[r.source] = (sourceConfirmed[r.source] || 0) + 1;
      sourceNet[r.source] = (sourceNet[r.source] || 0) + r.netSales;
      const catKey = CAT_KEYS[r.cat];
      totals[catKey] += r.netSales;
      totals.grand += r.netSales;
      const key = r.branch + '\u0001' + r.cat + '\u0001' + r.salesCode;
      groupMap.set(key, (groupMap.get(key) || 0) + r.netSales);
    }
  }
  const rowsArr = [];
  for (const [key, sum] of groupMap.entries()) {
    const [b, c, code] = key.split('\u0001');
    rowsArr.push([Number(b), Number(c), code, Math.round(sum * 100) / 100]);
  }
  return {
    ym,
    label: thaiMonthLabel(ym),
    totals: {
      import: Math.round(totals.import * 100) / 100,
      domestic: Math.round(totals.domestic * 100) / 100,
      service: Math.round(totals.service * 100) / 100,
      grand: Math.round(totals.grand * 100) / 100,
    },
    rows: rowsArr,
    statusCounts,
    totalRows: rows.length,
    confirmedRows,
    sourceConfirmed,
    sourceNet,
  };
}

function buildMonthDocs(allFlatRows) {
  const byYm = new Map();
  for (const r of allFlatRows) {
    if (!byYm.has(r.ym)) byYm.set(r.ym, []);
    byYm.get(r.ym).push(r);
  }
  const monthDocs = {};
  for (const [ym, rows] of byYm.entries()) {
    monthDocs[ym] = aggregateFlatRowsForMonth(rows, ym);
  }
  return monthDocs;
}

function buildSalesNamesDirectory(allFlatRows) {
  const tally = new Map(); // salesCode -> Map(name -> count)
  for (const r of allFlatRows) {
    if (r.status !== CONFIRMED_STATUS || !r.salesCode) continue;
    if (!tally.has(r.salesCode)) tally.set(r.salesCode, new Map());
    const m = tally.get(r.salesCode);
    const name = r.staffName || '';
    if (!name) continue;
    m.set(name, (m.get(name) || 0) + 1);
  }
  const result = {};
  for (const [code, m] of tally.entries()) {
    let best = null, bestCount = -1;
    for (const [name, cnt] of m.entries()) {
      if (cnt > bestCount) { best = name; bestCount = cnt; }
    }
    if (best) result[code] = best;
  }
  return result;
}

/* ==========================================================================
   Upload pipeline: read the 3 Excel files client-side, run the full pipeline,
   POST only the aggregated result to /api/upload.
   ========================================================================== */

const uploadState = { INV: null, TAX: null, DEP: null };

function readWorkbook(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const wb = XLSX.read(data, { type: 'array', cellDates: false });
        resolve(wb);
      } catch (err) { reject(err); }
    };
    reader.onerror = () => reject(new Error('อ่านไฟล์ไม่สำเร็จ'));
    reader.readAsArrayBuffer(file);
  });
}

async function processAndUpload() {
  const msgBox = document.getElementById('uploadMsg');
  const processBtn = document.getElementById('processBtn');
  const processingLabel = document.getElementById('processingLabel');
  msgBox.innerHTML = '';
  if (!uploadState.INV || !uploadState.TAX) {
    msgBox.innerHTML = '<div class="msg error">ต้องมีไฟล์ INV และ TAX อย่างน้อยครับ</div>';
    return;
  }
  processBtn.disabled = true;
  processingLabel.style.display = 'inline';
  try {
    const invWb = await readWorkbook(uploadState.INV);
    const taxWb = await readWorkbook(uploadState.TAX);
    const depWb = uploadState.DEP ? await readWorkbook(uploadState.DEP) : null;

    const invDocs = extractRowsFromWorkbook(invWb, [(rows) => parseInvoiceLikeSheet(rows, 'INV')]);
    const taxMap = extractRowsFromWorkbook(taxWb, [parseTaxSheet]);
    const depDocs = depWb ? extractRowsFromWorkbook(depWb, [(rows) => parseInvoiceLikeSheet(rows, 'DEP')]) : [];

    if (!invDocs) throw new Error('ไม่พบโครงสร้างที่รู้จักในไฟล์ INV (หาแถวหัวตาราง "No"/"เลขใบสำคัญ" ไม่เจอ)');
    if (!taxMap) throw new Error('ไม่พบโครงสร้างที่รู้จักในไฟล์ TAX (หาแถวหัวตาราง "No"/"เลขใบกำกับภาษี" ไม่เจอ)');
    if (depWb && !depDocs) throw new Error('ไม่พบโครงสร้างที่รู้จักในไฟล์ DEP (หาแถวหัวตาราง "No"/"เลขใบสำคัญ" ไม่เจอ)');

    const invRowCount = invDocs.reduce((s, d) => s + d.items.length, 0);
    const depRowCount = (depDocs || []).reduce((s, d) => s + d.items.length, 0);
    let taxRowCount = 0;
    taxMap.forEach(() => { taxRowCount++; });

    const flatRows = computeFlatRows(invDocs, depDocs || [], taxMap);
    const unknownDateRows = flatRows.filter((r) => r.ym === 'unknown').length;

    const monthDocs = buildMonthDocs(flatRows);
    const salesNames = buildSalesNamesDirectory(flatRows);

    const body = {
      monthDocs,
      salesNames,
      uploaderName: document.getElementById('uploaderName').value.trim(),
      invFileName: uploadState.INV.name,
      taxFileName: uploadState.TAX.name,
      depFileName: uploadState.DEP ? uploadState.DEP.name : '',
      invRowCount,
      taxRowCount,
      depRowCount,
      unknownDateRows,
    };

    const resp = await fetch('/api/upload', {
      method: 'POST',
      headers: { 'content-type': 'application/json' },
      body: JSON.stringify(body),
    });
    const result = await resp.json();
    if (!resp.ok) throw new Error(result.error || 'บันทึกข้อมูลไม่สำเร็จ');

    let successMsg = `บันทึกสำเร็จ: อัปเดต ${result.monthsUpdated.length} เดือน`;
    if (unknownDateRows > 0) {
      successMsg += ` — คำเตือน: มี ${unknownDateRows} รายการที่อ่านวันที่ไม่ได้ (เก็บไว้ใน "ไม่ทราบวันที่")`;
    }
    msgBox.innerHTML = `<div class="msg success">${successMsg}</div>`;

    await loadIndexAndRender(true);
  } catch (err) {
    console.error(err);
    msgBox.innerHTML = `<div class="msg error">${(err && err.message) || 'เกิดข้อผิดพลาดไม่ทราบสาเหตุ'}</div>`;
  } finally {
    processBtn.disabled = !(uploadState.INV && uploadState.TAX);
    processingLabel.style.display = 'none';
  }
}

/* ==========================================================================
   Dashboard: fetch, merge, filter, and render
   ========================================================================== */

const dash = {
  indexData: { months: [], salesNames: {}, latestUpload: null },
  merged: null,        // current merged aggregate for the selected date range
  allTime: false,
  expandedBranches: new Set([1, 2, 3, 0]), // default: branches open
  expandedCats: new Set(),                 // default: categories closed
};

function fmtBaht(n) {
  const v = Math.round((n || 0) * 100) / 100;
  return '฿' + v.toLocaleString('th-TH', { minimumFractionDigits: 0, maximumFractionDigits: 0 });
}

function currentRealMonth() {
  const d = new Date();
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`;
}

function shiftMonth(ym, delta) {
  if (ym === 'unknown' || !ym) return ym;
  const [y, m] = ym.split('-').map(Number);
  const d = new Date(Date.UTC(y, m - 1 + delta, 1));
  return `${d.getUTCFullYear()}-${String(d.getUTCMonth() + 1).padStart(2, '0')}`;
}

function emptyAggregate() {
  return {
    totals: { import: 0, domestic: 0, service: 0, grand: 0 },
    rows: [],
    statusCounts: emptyStatusCounts(),
    totalRows: 0,
    confirmedRows: 0,
    sourceConfirmed: { INV: 0, DEP: 0 },
    sourceNet: { INV: 0, DEP: 0 },
  };
}

function mergeMonthDocs(docsObj) {
  const merged = emptyAggregate();
  const rowMap = new Map();
  for (const ym in docsObj) {
    const d = docsObj[ym];
    if (!d) continue;
    merged.totals.import += d.totals?.import || 0;
    merged.totals.domestic += d.totals?.domestic || 0;
    merged.totals.service += d.totals?.service || 0;
    merged.totals.grand += d.totals?.grand || 0;
    merged.totalRows += d.totalRows || 0;
    merged.confirmedRows += d.confirmedRows || 0;
    for (const s of STATUS_LIST) merged.statusCounts[s] += (d.statusCounts?.[s] || 0);
    merged.sourceConfirmed.INV += d.sourceConfirmed?.INV || 0;
    merged.sourceConfirmed.DEP += d.sourceConfirmed?.DEP || 0;
    merged.sourceNet.INV += d.sourceNet?.INV || 0;
    merged.sourceNet.DEP += d.sourceNet?.DEP || 0;
    for (const r of (d.rows || [])) {
      const key = r[0] + '\u0001' + r[1] + '\u0001' + r[2];
      rowMap.set(key, (rowMap.get(key) || 0) + r[3]);
    }
  }
  merged.rows = Array.from(rowMap.entries()).map(([key, sum]) => {
    const [b, c, code] = key.split('\u0001');
    return [Number(b), Number(c), code, sum];
  });
  return merged;
}

function filterRowsByEmployee(rows, salesCode) {
  if (!salesCode) return rows;
  return rows.filter((r) => r[2] === salesCode);
}

function totalsFromRows(rows) {
  const t = { import: 0, domestic: 0, service: 0, grand: 0 };
  for (const [, c, , sum] of rows) {
    t[CAT_KEYS[c]] += sum;
    t.grand += sum;
  }
  return t;
}

function resolveName(code) {
  return dash.indexData.salesNames[code] || code;
}

/* ---------- fetching ---------- */
async function fetchIndex() {
  const resp = await fetch('/api/index');
  if (!resp.ok) throw new Error('เชื่อมต่อเซิร์ฟเวอร์ไม่สำเร็จ');
  return resp.json();
}
async function fetchMonths(params) {
  const resp = await fetch('/api/months?' + params);
  if (!resp.ok) throw new Error('เชื่อมต่อเซิร์ฟเวอร์ไม่สำเร็จ');
  return resp.json();
}

async function loadIndexAndRender(forceReloadRange) {
  const metaPill = document.getElementById('metaPill');
  try {
    dash.indexData = await fetchIndex();
    if (dash.indexData.latestUpload && dash.indexData.latestUpload.uploadedAt) {
      const dt = new Date(dash.indexData.latestUpload.uploadedAt);
      const who = dash.indexData.latestUpload.uploaderName ? ` โดย ${dash.indexData.latestUpload.uploaderName}` : '';
      metaPill.textContent = `อัปเดตล่าสุด ${dt.toLocaleString('th-TH')}${who}`;
      metaPill.classList.remove('err');
    } else {
      metaPill.textContent = 'ยังไม่เคยมีการอัปโหลดข้อมูล';
      metaPill.classList.remove('err');
    }
  } catch (err) {
    metaPill.textContent = 'เชื่อมต่อเซิร์ฟเวอร์ไม่สำเร็จ ลองรีเฟรชอีกครั้ง';
    metaPill.classList.add('err');
  }
  await loadRangeAndRender();
}

async function loadRangeAndRender() {
  const fromEl = document.getElementById('fromMonth');
  const toEl = document.getElementById('toMonth');
  let docsObj = {};
  try {
    if (dash.allTime) {
      docsObj = await fetchMonths('all=1');
    } else {
      docsObj = await fetchMonths(`from=${fromEl.value}&to=${toEl.value}`);
    }
    document.getElementById('metaPill').classList.remove('err');
  } catch (err) {
    document.getElementById('metaPill').textContent = 'เชื่อมต่อเซิร์ฟเวอร์ไม่สำเร็จ ลองรีเฟรชอีกครั้ง';
    document.getElementById('metaPill').classList.add('err');
  }
  dash.merged = mergeMonthDocs(docsObj);
  renderAll();
}

/* ---------- rendering ---------- */
function renderAll() {
  const hasAnyUpload = dash.indexData.months && dash.indexData.months.length > 0;
  const hasDataInRange = dash.merged && dash.merged.totalRows > 0;
  const empSelect = document.getElementById('employeeSelect');
  const selectedEmp = empSelect.value;

  updateEmployeeOptions();

  const summaryGrid = document.getElementById('summaryGrid');
  const emptyCard = document.getElementById('emptyStateCard');
  const qualityCard = document.getElementById('qualityCard');
  const drillCard = document.getElementById('drillCard');

  if (!hasAnyUpload) {
    summaryGrid.style.display = 'none';
    qualityCard.style.display = 'none';
    drillCard.style.display = 'none';
    emptyCard.style.display = 'block';
    document.getElementById('emptyIcon').textContent = '📭';
    document.getElementById('emptyText').textContent = 'ยังไม่เคยมีการอัปโหลดข้อมูล — เริ่มต้นด้วยการอัปโหลดไฟล์ INV/TAX/DEP ด้านบน';
    return;
  }
  if (!hasDataInRange) {
    summaryGrid.style.display = 'none';
    qualityCard.style.display = 'none';
    drillCard.style.display = 'none';
    emptyCard.style.display = 'block';
    document.getElementById('emptyIcon').textContent = '🗓️';
    document.getElementById('emptyText').textContent = 'ไม่มีข้อมูลในช่วงเวลาที่เลือก ลองเปลี่ยนช่วงเวลาดูครับ';
    return;
  }

  const rowsForEmployee = filterRowsByEmployee(dash.merged.rows, selectedEmp);
  if (selectedEmp && rowsForEmployee.length === 0) {
    empSelect.value = '';
    renderAll();
    return;
  }
  emptyCard.style.display = 'none';
  summaryGrid.style.display = 'grid';

  const displayTotals = totalsFromRows(rowsForEmployee);
  document.getElementById('valImport').textContent = fmtBaht(displayTotals.import);
  document.getElementById('valDomestic').textContent = fmtBaht(displayTotals.domestic);
  document.getElementById('valService').textContent = fmtBaht(displayTotals.service);
  document.getElementById('valGrand').textContent = fmtBaht(displayTotals.grand);
  const grand = displayTotals.grand || 1;
  document.getElementById('barImport').style.width = (100 * displayTotals.import / grand) + '%';
  document.getElementById('barDomestic').style.width = (100 * displayTotals.domestic / grand) + '%';
  document.getElementById('barService').style.width = (100 * displayTotals.service / grand) + '%';

  if (!selectedEmp) {
    qualityCard.style.display = 'block';
    renderQualityChips();
  } else {
    qualityCard.style.display = 'none';
  }

  drillCard.style.display = 'block';
  renderTree(rowsForEmployee);

  renderFilterSummary();
}

function renderQualityChips() {
  const chipRow = document.getElementById('chipRow');
  chipRow.innerHTML = '';
  const classFor = {
    'ตรงกัน (ยืนยันแล้ว)': 'ok',
    'ไม่ตรง (ยังไม่ยืนยันการชำระเงิน)': 'warn',
    'ไม่นับ (รอตรวจสอบ - ฐานภาษีไม่ตรงกับยอดหลังหักส่วนลด)': 'warn',
    'ไม่นับ (ใบสำคัญเป็น Draft ยังไม่อนุมัติ)': 'neutral',
    'ไม่นับ (ใบสำคัญถูกยกเลิก)': 'bad',
  };
  for (const status of STATUS_LIST) {
    const count = dash.merged.statusCounts[status] || 0;
    const chip = document.createElement('span');
    chip.className = 'chip ' + (classFor[status] || 'neutral');
    chip.innerHTML = `<b>${count.toLocaleString('th-TH')}</b> ${status}`;
    chipRow.appendChild(chip);
  }
  const depCount = dash.merged.sourceConfirmed.DEP || 0;
  const depSum = dash.merged.sourceNet.DEP || 0;
  document.getElementById('depLine').textContent =
    `รวมยอดจากใบสั่งขาย/มัดจำ (DEP) ที่ยืนยันชำระแล้ว ${depCount.toLocaleString('th-TH')} รายการ เป็นเงิน ${fmtBaht(depSum)}`;
}

function updateEmployeeOptions() {
  const select = document.getElementById('employeeSelect');
  const current = select.value;
  const rows = (dash.merged && dash.merged.rows) || [];
  const codesWithSales = new Set(rows.filter((r) => r[3] > 0).map((r) => r[2]));
  const codes = Array.from(codesWithSales).sort((a, b) => resolveName(a).localeCompare(resolveName(b), 'th'));
  select.innerHTML = '<option value="">ทุกคน</option>';
  for (const code of codes) {
    const opt = document.createElement('option');
    opt.value = code;
    opt.textContent = `${resolveName(code)} — ${code}`;
    select.appendChild(opt);
  }
  if (codesWithSales.has(current)) select.value = current;
  else select.value = '';
}

function renderFilterSummary() {
  const el = document.getElementById('filterSummary');
  const emp = document.getElementById('employeeSelect').value;
  let rangeText;
  if (dash.allTime) rangeText = 'ทั้งหมด';
  else {
    const from = document.getElementById('fromMonth').value;
    const to = document.getElementById('toMonth').value;
    rangeText = from === to ? thaiMonthLabel(from) : `${thaiMonthLabel(from)} – ${thaiMonthLabel(to)}`;
  }
  const empText = emp ? ` · พนักงาน: ${resolveName(emp)}` : '';
  el.textContent = rangeText + empText;
}

/* ---------- drill-down tree ---------- */
function renderTree(rows) {
  const root = document.getElementById('treeRoot');
  const search = document.getElementById('searchBox').value.trim().toLowerCase();
  const showZero = document.getElementById('showZeroChk').checked;

  const tree = new Map();
  for (const [b, c, code, sum] of rows) {
    if (!tree.has(b)) tree.set(b, { sum: 0, cats: new Map() });
    const bt = tree.get(b);
    if (!bt.cats.has(c)) bt.cats.set(c, { sum: 0, people: [] });
    const ct = bt.cats.get(c);
    ct.people.push({ code, sum });
  }
  // recompute sums bottom-up so hidden/zero filtering doesn't skew percentages
  for (const [, bt] of tree) {
    for (const [, ct] of bt.cats) {
      ct.sum = ct.people.reduce((s, p) => s + p.sum, 0);
    }
    bt.sum = Array.from(bt.cats.values()).reduce((s, ct) => s + ct.sum, 0);
  }
  const grandTotal = Array.from(tree.values()).reduce((s, bt) => s + bt.sum, 0) || 1;

  root.innerHTML = '';
  const branchOrder = [1, 2, 3, 0];
  let renderedAny = false;

  for (const b of branchOrder) {
    const bt = tree.get(b);
    if (!bt) continue;
    if (bt.sum === 0 && !showZero) continue;

    const catEntries = [0, 1, 2]
      .filter((c) => bt.cats.has(c))
      .map((c) => ({ c, ...bt.cats.get(c) }))
      .map((ct) => {
        let people = ct.people.slice().sort((a, b2) => b2.sum - a.sum);
        if (search) {
          people = people.filter((p) => p.code.toLowerCase().includes(search) || resolveName(p.code).toLowerCase().includes(search));
        }
        if (!showZero) people = people.filter((p) => p.sum !== 0);
        return { ...ct, people };
      })
      .filter((ct) => ct.people.length > 0 || (showZero && !search));

    if (catEntries.length === 0) continue;
    renderedAny = true;

    const branchRow = document.createElement('div');
    branchRow.className = 'row branch' + (dash.expandedBranches.has(b) ? ' open' : '');
    branchRow.innerHTML = `<span class="caret">▸</span><span class="name">${BRANCH_CODES[b]}</span><span class="amt">${fmtBaht(bt.sum)}</span><span class="pct">${(100 * bt.sum / grandTotal).toFixed(1)}%</span>`;
    const body = document.createElement('div');
    body.className = 'group-body' + (dash.expandedBranches.has(b) ? ' show' : '');
    branchRow.addEventListener('click', () => {
      if (dash.expandedBranches.has(b)) dash.expandedBranches.delete(b);
      else dash.expandedBranches.add(b);
      renderTree(rows);
    });
    root.appendChild(branchRow);

    for (const ct of catEntries) {
      const catKeyId = b + '_' + ct.c;
      const catRow = document.createElement('div');
      catRow.className = 'row category' + (dash.expandedCats.has(catKeyId) ? ' open' : '');
      const dotClass = ct.c === 0 ? 'domestic' : ct.c === 1 ? 'import' : 'service';
      catRow.innerHTML = `<span class="caret">▸</span><span class="cat-dot ${dotClass}"></span><span class="name">${CAT_CODES[ct.c]}</span><span class="amt">${fmtBaht(ct.sum)}</span><span class="pct">${(100 * ct.sum / (bt.sum || 1)).toFixed(1)}%</span>`;
      const catBody = document.createElement('div');
      catBody.className = 'group-body' + (dash.expandedCats.has(catKeyId) ? ' show' : '');
      catRow.addEventListener('click', () => {
        if (dash.expandedCats.has(catKeyId)) dash.expandedCats.delete(catKeyId);
        else dash.expandedCats.add(catKeyId);
        renderTree(rows);
      });
      body.appendChild(catRow);
      body.appendChild(catBody);

      for (const p of ct.people) {
        const personRow = document.createElement('div');
        personRow.className = 'row person';
        personRow.innerHTML = `<span class="name">${resolveName(p.code)} — ${p.code}</span><span class="amt">${fmtBaht(p.sum)}</span><span class="pct">${(100 * p.sum / (ct.sum || 1)).toFixed(1)}%</span>`;
        catBody.appendChild(personRow);
      }
    }
    root.appendChild(body);
  }

  if (!renderedAny) {
    root.innerHTML = '<div class="empty-state">ไม่พบรายการที่ตรงกับเงื่อนไข</div>';
  }
}

/* ==========================================================================
   Wiring: dropzones, filter bar, polling, init
   ========================================================================== */

function setupDropzone(id, fileInputId, fileLabelId, kind) {
  const zone = document.getElementById(id);
  const input = document.getElementById(fileInputId);
  const label = document.getElementById(fileLabelId);
  const defaultText = label.textContent;

  function setFile(file) {
    if (!file) return;
    uploadState[kind] = file;
    zone.classList.add('filled');
    label.textContent = file.name;
    const processBtn = document.getElementById('processBtn');
    processBtn.disabled = !(uploadState.INV && uploadState.TAX);
  }

  zone.addEventListener('click', () => input.click());
  input.addEventListener('change', () => setFile(input.files[0]));
  zone.addEventListener('dragover', (e) => { e.preventDefault(); zone.classList.add('dragover'); });
  zone.addEventListener('dragleave', () => zone.classList.remove('dragover'));
  zone.addEventListener('drop', (e) => {
    e.preventDefault();
    zone.classList.remove('dragover');
    if (e.dataTransfer.files && e.dataTransfer.files[0]) setFile(e.dataTransfer.files[0]);
  });
}

function setupFilterBar() {
  const fromEl = document.getElementById('fromMonth');
  const toEl = document.getElementById('toMonth');
  const nowYm = currentRealMonth();
  fromEl.value = nowYm;
  toEl.value = nowYm;

  function exitAllTime() { dash.allTime = false; }

  fromEl.addEventListener('change', () => {
    exitAllTime();
    if (fromEl.value > toEl.value) toEl.value = fromEl.value;
    loadRangeAndRender();
  });
  toEl.addEventListener('change', () => {
    exitAllTime();
    if (toEl.value < fromEl.value) fromEl.value = toEl.value;
    loadRangeAndRender();
  });
  document.getElementById('prevBtn').addEventListener('click', () => {
    exitAllTime();
    fromEl.value = shiftMonth(fromEl.value, -1);
    toEl.value = shiftMonth(toEl.value, -1);
    loadRangeAndRender();
  });
  document.getElementById('nextBtn').addEventListener('click', () => {
    exitAllTime();
    fromEl.value = shiftMonth(fromEl.value, 1);
    toEl.value = shiftMonth(toEl.value, 1);
    loadRangeAndRender();
  });
  document.getElementById('thisMonthBtn').addEventListener('click', () => {
    exitAllTime();
    const m = currentRealMonth();
    fromEl.value = m;
    toEl.value = m;
    loadRangeAndRender();
  });
  document.getElementById('allTimeBtn').addEventListener('click', () => {
    dash.allTime = true;
    loadRangeAndRender();
  });
  document.getElementById('employeeSelect').addEventListener('change', renderAll);
  document.getElementById('searchBox').addEventListener('input', () => {
    if (dash.merged) renderTree(filterRowsByEmployee(dash.merged.rows, document.getElementById('employeeSelect').value));
  });
  document.getElementById('showZeroChk').addEventListener('change', () => {
    if (dash.merged) renderTree(filterRowsByEmployee(dash.merged.rows, document.getElementById('employeeSelect').value));
  });
}

function init() {
  setupDropzone('dzInv', 'fileInv', 'dzInvFile', 'INV');
  setupDropzone('dzTax', 'fileTax', 'dzTaxFile', 'TAX');
  setupDropzone('dzDep', 'fileDep', 'dzDepFile', 'DEP');
  document.getElementById('processBtn').addEventListener('click', processAndUpload);
  setupFilterBar();

  loadIndexAndRender().then(() => {
    // open the upload box automatically if nothing has ever been uploaded
    if (!dash.indexData.months || dash.indexData.months.length === 0) {
      document.getElementById('uploadBox').open = true;
    }
  });

  setInterval(() => { loadRangeAndRender(); }, 60000);
}

document.addEventListener('DOMContentLoaded', init);
