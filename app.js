/** ========= Cody Wealth WebApp: server (app.gs) =========
 *  動態表頭整合版（A1 為表頭）
 *  - 全部讀寫一律用 getColMap_() 取欄位位置（不再寫死欄位號）
 *  - NOW：Current + posted(-金額)，planned 不算進 NOW
 *  - submitCashflow：status=auto → 依當列日期產生 R1C1 公式
 */

const CF_SHEET   = 'Cashflow'; // 資料分頁
const HEADER_ROW = 1;          // 表頭列（A1）
const COLS = {                 // 後備：找不到表頭時的預設欄位（A..F）
  date:1, item:2, amount:3, account:4, status:5, note:6
};

/** Web App 入口 */
function doGet() {
  const tpl = HtmlService.createTemplateFromFile('index'); // ← 指向 index
  return tpl.evaluate()
    .setTitle('Cody Wealth')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// 讓 <?!= include('xxx'); ?> 可用
function include(name) {
  return HtmlService.createHtmlOutputFromFile(name).getContent();
}

/** 取選單資料（動態表頭版）：帳戶清單、狀態清單 */
// app.gs
/** 取選單資料（動態表頭版）：帳戶清單、狀態清單 */
function getMeta(){
  const sh = SpreadsheetApp.getActive().getSheetByName(CF_SHEET);
  if (!sh) return { accounts: [], statuses: ['auto','posted','planned'] };

  const lastRow = sh.getLastRow();
  // 限制讀取範圍到已使用的列數，而非整個工作表
  // 假設帳戶和狀態欄位分別在第4和第5欄
  const accountsCol = sh.getRange(1, 4, lastRow).getValues().flat();
  const statusesCol = sh.getRange(1, 5, lastRow).getValues().flat();

  // 使用 Set 去重
  const uniqueAccounts = new Set(accountsCol.filter(v => v !== ''));
  const uniqueStatuses = new Set(statusesCol.filter(v => v !== ''));

  return {
    accounts: Array.from(uniqueAccounts),
    statuses: ['auto', 'posted', 'planned', ...Array.from(uniqueStatuses)],
  };
}

/** Dashboard：依帳戶彙總 NOW（Current + posted(-金額)） */
function getSummaryNow(){
  const map = calcNow_(); // {acc: now}
  return Object.keys(map).sort().map(a=>({ account:a, now: map[a] }));
}

/** Dashboard：取 planned 清單（日期、帳戶、金額、明細） */
function getPlannedList(){
  const sh = SpreadsheetApp.getActive().getSheetByName(CF_SHEET);
  if (!sh) return [];
  const lastRow = sh.getLastRow();
  if (lastRow <= HEADER_ROW) return [];

  const n = lastRow - HEADER_ROW;
  const vals = sh.getRange(HEADER_ROW+1, 1, n, sh.getLastColumn()).getValues();
  const map = getColMap_(sh);
  const iD = map.date-1, iI = map.item-1, iA = map.amount-1, iACC = map.account-1, iS = map.status-1;

  const out = [];
  const tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone() || 'Asia/Taipei';

  for (const r of vals){
    if (!r[iS]) continue;
    const sta = String(r[iS]).trim().toLowerCase();
    if (sta !== 'planned') continue;

    const d = r[iD] instanceof Date ? r[iD] : (r[iD] ? new Date(r[iD]) : null);
    out.push({
      date: d ? Utilities.formatDate(d, tz, 'yyyy-MM-dd') : '',
      account: String(r[iACC]||'').trim(),
      amount: Number(r[iA]||0) * -1, // planned 顯示為 -金額
      item:   String(r[iI]||'').trim()
    });
  }
  out.sort((a,b)=> String(a.date).localeCompare(String(b.date)));
  return out;
}

/** 依種類取清單：kind = 'planned' | 'now' | 'final'
 *  planned: 只取 status=planned
 *  now:     只取 status=posted（排除 item='Current'）
 *  final:   posted + planned（都排除 Current）
 *  輸出：[{date, account, amount, item}]
 */
function getTxByKind(kind) {
  kind = String(kind || '').toLowerCase();
  if (!['planned','now','final'].includes(kind)) kind = 'planned';

  const sh = SpreadsheetApp.getActive().getSheetByName(CF_SHEET);
  const lastRow = sh.getLastRow();
  if (lastRow <= HEADER_ROW) return [];

  const n = lastRow - HEADER_ROW;
  const vals = sh.getRange(HEADER_ROW+1, 1, n, sh.getLastColumn()).getValues();
  const M = getColMap_(sh);
  const iD=M.date-1, iI=M.item-1, iA=M.amount-1, iACC=M.account-1, iS=M.status-1;

  const out = [];
  for (const r of vals) {
    const item = String(r[iI]||'').trim();
    const sta  = String(r[iS]||'').trim().toLowerCase();
    const acc  = String(r[iACC]||'').trim();
    const rawD = r[iD];
    const dObj = rawD instanceof Date ? rawD : (rawD ? new Date(rawD) : null);

    // 篩選邏輯
    if (kind === 'planned') {
      if (sta !== 'planned') continue;
    } else if (kind === 'now') {
      if (sta !== 'posted') continue;
      if (item === 'Current') continue; // 不列 snapshot
    } else if (kind === 'final') {
      if (!(sta === 'posted' || sta === 'planned')) continue;
      if (item === 'Current') continue;
    }

    // 金額規則：posted / planned 都是 -amt（支出正 → 餘額減）
    const amt = Number(r[iA]||0) * -1;

    out.push({
      date: dObj ? Utilities.formatDate(dObj, SpreadsheetApp.getActive().getSpreadsheetTimeZone()||'Asia/Taipei','yyyy-MM-dd') : '',
      account: acc,
      amount: amt,
      item: item
    });
  }
  // 時間由近到遠
  out.sort((a,b)=> String(b.date||'').localeCompare(String(a.date||'')));
  return out;
}

/** 明細清單（可過濾：account / dateFrom / dateTo；空值=不過濾） */
function getTransactions(filter){
  filter = filter || {};
  const accPick = String(filter.account||'').trim();
  const fromStr = String(filter.dateFrom||'').trim();
  const toStr   = String(filter.dateTo||'').trim();
  const from = fromStr ? new Date(fromStr) : null;
  const to   = toStr   ? new Date(toStr)   : null;

  const sh = SpreadsheetApp.getActive().getSheetByName(CF_SHEET);
  if (!sh) return [];
  const lastRow = sh.getLastRow();
  if (lastRow <= HEADER_ROW) return [];

  const n = lastRow - HEADER_ROW;
  const vals = sh.getRange(HEADER_ROW+1, 1, n, sh.getLastColumn()).getValues();
  const M = getColMap_(sh);
  const iD=M.date-1, iI=M.item-1, iA=M.amount-1, iACC=M.account-1, iS=M.status-1, iN=M.note-1;
  const tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone() || 'Asia/Taipei';

  const out = [];
  for (let i=0;i<vals.length;i++){
    const r = vals[i];
    const _row = HEADER_ROW + 1 + i; // ← 真實列號（非常重要）

    const dRaw = r[iD];
    const dObj = dRaw instanceof Date ? dRaw : (dRaw ? new Date(dRaw) : null);
    if (from && dObj && dObj < from) continue;
    if (to   && dObj && dObj > to)   continue;

    const acc = String(r[iACC]||'').trim();
    if (accPick && acc !== accPick) continue;

    out.push({
      _row,
      date: dObj ? Utilities.formatDate(dObj, tz, 'yyyy-MM-dd') : '',
      item: String(r[iI]||'').trim(),
      amount: Number(r[iA]||0),
      account: acc,
      status: String(r[iS]||'').trim().toLowerCase(),
      note: String(r[iN]||'').trim()
    });
  }
  // 近到遠
  out.sort((a,b)=> String(b.date).localeCompare(String(a.date)));
  return out;
}

/** Cashflow 分頁 URL（給「開啟分頁」用） */
function getCashflowUrl(){
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(CF_SHEET);
  if (!sh) return ss.getUrl();
  return ss.getUrl() + '#gid=' + sh.getSheetId();
}

/* ====== 共同工具 ====== */

/** 讀表頭 → 動態欄位定位；若找不到就回退到 COLS */
function getColMap_(sh){
  const row = sh.getRange(HEADER_ROW, 1, 1, sh.getLastColumn()).getValues()[0] || [];
  const norm = s => String(s||'').trim().toLowerCase();
  const alias = {
    date:['date','日期'],
    item:['item','明細','項目'],
    amount:['amount','金額','金額(+-)'],
    account:['account','帳戶','帳號'],
    status:['status','狀態'],
    note:['note','備註','註記']
  };
  const idx = {};
  for (const k of Object.keys(alias)){
    const names = alias[k];
    let found = 0;
    for (let c=0; c<row.length; c++){
      if (names.includes(norm(row[c]))) { found = c+1; break; }
    }
    idx[k] = found || COLS[k];
  }
  return idx; // {date, item, amount, account, status, note}
}

/** 現有餘額（NOW）：Current + posted(-金額)；planned 不計 */
function calcNow_(){
  const sh = SpreadsheetApp.getActive().getSheetByName(CF_SHEET);
  const out = {};
  if (!sh) return out;

  const lastRow = sh.getLastRow();
  if (lastRow <= HEADER_ROW) return out;

  const vals = sh.getRange(HEADER_ROW+1, 1, lastRow-HEADER_ROW, sh.getLastColumn()).getValues();
  const M = getColMap_(sh);
  const iItem=M.item-1, iAmt=M.amount-1, iAcc=M.account-1, iSta=M.status-1;

  for (const r of vals){
    const acc = String(r[iAcc]||'').trim(); 
    if (!acc) continue;
    if (!(acc in out)) out[acc]=0;

    const item = String(r[iItem]||'').trim();
    const sta  = String(r[iSta]||'').trim().toLowerCase();
    const amt  = Number(r[iAmt]||0);

    if (item === 'Current') out[acc]+=amt;
    else if (sta === 'posted') out[acc]+=-amt;
  }
  return out;
}

/** 回傳指定帳戶的 Balance(NOW)；name 空字串→回全部加總 */
function getBalanceForAccount_(name) {
  const sh = SpreadsheetApp.getActive().getSheetByName(CF_SHEET);
  if (!sh) return 0;
  const lastRow = sh.getLastRow();
  if (lastRow <= HEADER_ROW) return 0;

  const vals = sh.getRange(HEADER_ROW+1, 1, lastRow-HEADER_ROW, sh.getLastColumn()).getValues();
  const M = getColMap_(sh);
  const iItem = M.item - 1;
  const iAmt  = M.amount - 1;
  const iAcc  = M.account - 1;
  const iSta  = M.status - 1;

  const out = {};
  for (const r of vals) {
    const acc = String(r[iAcc]||'').trim();
    if (!acc) continue;
    if (!(acc in out)) out[acc]=0;

    const item = String(r[iItem]||'').trim();
    const sta  = String(r[iSta]||'').trim().toLowerCase();
    const amt  = Number(r[iAmt]||0);

    if (item === 'Current') out[acc] += amt;
    else if (sta === 'posted') out[acc] += -amt;
  }

  if (!name) return Object.values(out).reduce((a,b)=>a+Number(b||0),0);
  return out[name] ?? 0;
}

/** 寫入一筆（status=auto → 依欄位距離置入 R1C1 公式），回傳該帳戶 NOW */
function submitCashflow(rec) {
  const sh  = SpreadsheetApp.getActive().getSheetByName(CF_SHEET);
  if (!sh) throw new Error('找不到資料分頁：' + CF_SHEET);

  const row = Math.max(sh.getLastRow() + 1, HEADER_ROW + 1);
  const M = getColMap_(sh);

  const d = rec.date ? new Date(rec.date) : '';
  const item = String(rec.item||'').trim();
  const amt  = rec.amount === '' || rec.amount == null ? '' : Number(rec.amount);
  const acc  = String(rec.account||'').trim();
  let   st   = String(rec.status||'').trim(); // 'auto' | 'posted' | 'planned'
  const note = String(rec.note||'').trim();

  if (M.date)    sh.getRange(row, M.date).setValue(d);
  if (M.item)    sh.getRange(row, M.item).setValue(item);
  if (M.amount)  sh.getRange(row, M.amount).setValue(amt);
  if (M.account) sh.getRange(row, M.account).setValue(acc);
  if (M.note)    sh.getRange(row, M.note).setValue(note);

  if (M.status) {
    const stCell = sh.getRange(row, M.status);
    if (!st || /^auto$/i.test(st)) {
      const delta = (M.date || 0) - (M.status || 0);
      stCell.setFormulaR1C1(`=IF(RC[${delta}]<=TODAY(),"posted","planned")`);
    } else if (/^posted$/i.test(st)) {
      stCell.clearFormat().setValue('posted');
    } else if (/^planned$/i.test(st)) {
      stCell.clearFormat().setValue('planned');
    } else {
      const delta = (M.date || 0) - (M.status || 0);
      stCell.setFormulaR1C1(`=IF(RC[${delta}]<=TODAY(),"posted","planned")`);
    }
  }

  applyFormatAndValidation_?.(row);
  SpreadsheetApp.flush();       // ★ 同理
  const balanceNow = getBalanceForAccount_(acc);
  return { row, balanceNow };
}

/** 一鍵修復：把缺少驗證/格式的列補齊 */
function fixMissingValidationAll_(){
  const sh = SpreadsheetApp.getActive().getSheetByName(CF.sheetName);
  if (!sh) return;
  const lastRow = sh.getLastRow();
  if (lastRow <= CF.headerRow) return;

  const width = CF.lastCol - CF.firstCol + 1;
  const rng = sh.getRange(CF.headerRow + 1, CF.firstCol, lastRow - CF.headerRow, width);
  const dv  = rng.getDataValidations(); // 2D 陣列
  const startRow = CF.headerRow + 1;

  for (let r = 0; r < dv.length; r++){
    if (dv[r].some(cell => cell == null)) {
      applyFormatAndValidation_(startRow + r);
    }
  }
  SpreadsheetApp.getUi().alert('已檢查並補齊缺失的格式與下拉驗證。');
}

/* ===== 你前端會叫用的：CTBC investment（取 ✅ 分頁 H,K,L,N,T,U,V,AF,AH） ===== */
function getCTBCInvestments(){
  const SHEET = '✅';
  // A=1, ..., AF=32, AG=33, AH=34
  const INDEX = {H:8, K:11, L:12, N:14, R:18, T:20, U:21, V:22, AF:32, AH:34};
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET);
  if (!sh) return [];

  const lastRow = sh.getLastRow();
  if (lastRow <= 1) return [];

  const n = lastRow - 1;

  // 只讀取到實際需要的欄位，避免超取
  const needCols = Math.max.apply(null, Object.values(INDEX));
  const readCols = Math.min(sh.getLastColumn(), needCols);

  const values = sh.getRange(2, 1, n, readCols).getValues();
  const tz = Session.getScriptTimeZone();

  const toStr = (v) => v instanceof Date
    ? Utilities.formatDate(v, tz, 'yyyy-MM-dd')
    : (v == null ? '' : String(v));

  // 幣別正規化（不改動原 V，另提供 CUR）
  function normCur_(v){
    const s = String(v||'').toUpperCase();
    if (s.includes('USD')) return 'USD';
    if (s.includes('TWD') || s.includes('NTD') || /(NT|台幣|新台幣)/i.test(s)) return 'TWD';
    return '';
  }

  // 輸出：每列是一個物件，包含列號 ROW（2 起算）
  const out = values.map((r, i) => ({
    ROW: i + 2, // 真實試算表列號
    H:  toStr(r[INDEX.H  - 1]),
    K:  toStr(r[INDEX.K  - 1]),
    L:  Number(r[INDEX.L  - 1]) || 0,
    N:  Number(r[INDEX.N  - 1]) || 0,
    R:   toStr(r[INDEX.R - 1]),
    QUOTE: Number(r[INDEX.R - 1]) || 0, 
    T:  Number(r[INDEX.T  - 1]) || 0,
    U:  Number(r[INDEX.U  - 1]) || 0,
    V:  toStr(r[INDEX.V  - 1]),
    CUR: normCur_(r[INDEX.V - 1]),
    AF: Number(r[INDEX.AF - 1]) || 0,
    AH: toStr(r[INDEX.AH - 1]),
  }));

  return out;
}

/**
 * CTBC 匯總（供 ctbc_js 直接呼叫）
 * 來源：getCTBCInvestments()（不可重複實作）
 * 回傳鍵名對齊前端 IDs：
 *   ctbcAllN, ctbcAllT, ctbcAllNoInt, ctbcAllAF,
 *   ctbcTWD_N, ctbcTWD_T, ctbcTWD_AF,
 *   ctbcUSD_N, ctbcUSD_T, ctbcUSD_AF,
 *   ctbcAllROI, ctbcTWD_ROI, ctbcUSD_ROI
 */
function getCTBCAggregates(kind){
  const rowsAll = getCTBCInvestments() || [];

  // 允許 'fund' 或 '基金' 觸發基金篩選（比對 H/K）；允許 'stock' 或 '股票' 觸發股票篩選
  const k = String(kind||'').toLowerCase();
  const isFundReq  = k.includes('fund')  || k.includes('基金');
  const isStockReq = k.includes('stock') || k.includes('股票');
  const isFund = (r)=> /基金|FUND/i.test(String(r.H||'')) || /基金|FUND/i.test(String(r.K||''));
  const rows = isFundReq ? rowsAll.filter(isFund)
              : isStockReq ? rowsAll.filter(r=> !isFund(r))
              : rowsAll;

  // 幣別挑選：優先 CUR，否則以 V 判斷
  const pick = (code)=> rows.filter(r=>{
    const cur = (r.CUR || '').toUpperCase();
    if (cur) return cur === code;
    const v = String(r.V||'').toUpperCase();
    if (code === 'USD') return v.includes('USD');
    if (code === 'TWD') return v.includes('TWD') || v.includes('NTD') || /(NT|台幣|新台幣)/i.test(v);
    return false;
  });

  // 與前端 ctbc_js 一致的 normalizeU：|U|≤1 視為小數，乘 100；否則視為百分比
  function normalizeU(u){
    var n = Number(u||0);
    if (!isFinite(n)) n = 0;
    return (Math.abs(n) <= 1) ? (n * 100) : n;
  }

  // 匯總：N/T/AF/L，同時計算「以 N 加權」的 U 加權平均（roiU）
  function agg(list){
    var N=0, T=0, AF=0, Lsum=0, w=0, wU=0;
    for (var i=0;i<list.length;i++){
      var r = list[i];
      var n = Number(r.N||0);
      var t = Number(r.T||0);
      var af= Number(r.AF||0);
      var l = Number(r.L||0);
      var u = normalizeU(r.U);
      if (!isFinite(n)) n = 0;
      if (!isFinite(t)) t = 0;
      if (!isFinite(af)) af = 0;
      if (!isFinite(l)) l = 0;
      N += n; T += t; AF += af; Lsum += l;
      if (n){ w += n; wU += (u * n); }
    }
    var roiU = w ? (wU / w) : null; // 與前端一致：N 加權的 U 平均（百分比）
    return { N:N, T:T, AF:AF, Lsum:Lsum, roiU:roiU };
  }

  var all = agg(rows);
  var twd = agg(pick('TWD'));
  var usd = agg(pick('USD'));

  // AllNoInt = 現值合計（ΣN），依照當前篩選後的集合計算
  var allNoInt = all.N;

  return {
    ctbcAllN:     all.Lsum,
    ctbcAllT:     all.T,
    ctbcAllNoInt: allNoInt,
    ctbcAllAF:    all.AF,

    ctbcTWD_N:    twd.N,
    ctbcTWD_T:    twd.T,
    ctbcTWD_AF:   twd.AF,

    ctbcUSD_N:    usd.N,
    ctbcUSD_T:    usd.T,
    ctbcUSD_AF:   usd.AF,

    // ROI 與前端相同：以 U 的 N 加權平均（百分比值）
    ctbcAllROI:   (all.roiU==null?0:all.roiU),
    ctbcTWD_ROI:  (twd.roiU==null?0:twd.roiU),
    ctbcUSD_ROI:  (usd.roiU==null?0:usd.roiU)
  };
}

/**
 * 一次回傳「基金 / 股票」的現值與 ROI（與前端一致：U 以 N 加權；U 在 -1..1 視為小數×100）
 * 回傳：{ fund:{ value, roi }, stock:{ value, roi } }
 */
function getCTBCFundStock(){
  const rowsAll = getCTBCInvestments() || [];
  const isFund = (r)=> /基金|FUND/i.test(String(r.H||'')) || /基金|FUND/i.test(String(r.K||''));
  const fundRows  = rowsAll.filter(isFund);
  const stockRows = rowsAll.filter(r=> !isFund(r));

  function normalizeU(u){
    var n = Number(u||0); if (!isFinite(n)) n = 0;
    return (Math.abs(n) <= 1) ? (n * 100) : n;
  }
  function agg(list){
    var N=0, w=0, wU=0;
    for (var i=0;i<list.length;i++){
      var n = Number(list[i].N||0); if (!isFinite(n)) n=0;
      var u = normalizeU(list[i].U);
      N += n; if (n){ w += n; wU += (u*n); }
    }
    return { value: N, roi: w ? (wU/w) : 0 };
  }

  return { fund: agg(fundRows), stock: agg(stockRows) };
}

/** 讀取某列 AF 儲存格的公式與顯示值（優先回公式） */
function getCTBCInterestCell(rowNumber){
  const SHEET = '✅';
  const COL_AF = 32;
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET);
  if (!sh) throw new Error('找不到工作表：' + SHEET);
  if (rowNumber < 2) throw new Error('列號不合法');

  const rng = sh.getRange(rowNumber, COL_AF);
  return {
    formula: rng.getFormula() || '',
    value: rng.getDisplayValue() || ''
  };
}

/** 前端寫回「已領利息(USD)」（AF=32），並讓觸發器自動回填 AH(34) */
function setCTBCInterestUSD(rowNumber, input){
  const SHEET = '✅';
  const COL_AF = 32; // 已領利息
  const COL_AH = 34; // 更新日
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET);
  if (!sh) throw new Error('找不到工作表：' + SHEET);
  if (rowNumber < 2) throw new Error('列號不合法');

  const rngAF = sh.getRange(rowNumber, COL_AF);
  const s = (typeof input === 'string') ? input.trim() : '';

  if (s && s[0] === '=') {
    rngAF.setFormula(s);         // ← 保留公式
  } else {
    rngAF.setValue(Number(input)||0);
  }

  // 注意：用程式寫值不會觸發你的 onEdit，所以這裡直接寫 AH（見第 4 點）
  const tz = Session.getScriptTimeZone() || 'Asia/Taipei';
  const today = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  sh.getRange(rowNumber, COL_AH).setValue(today);

  SpreadsheetApp.flush();
  return { row: rowNumber };
}

/** 讀取某列 QUOTE(R) 的公式/值 */
function getCTBCQuoteCell(rowNumber){
  const SHEET = '✅';
  const COL_R = 18; // 目前報價
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET);
  if (!sh) throw new Error('找不到工作表：' + SHEET);
  if (rowNumber < 2) throw new Error('列號不合法');
  const rng = sh.getRange(rowNumber, COL_R);
  return { formula: rng.getFormula() || '', value: rng.getDisplayValue() || '' };
}

/** 寫回「目前報價 (USD)」（R=18），並更新 AH(34) */
function setCTBCQuoteUSD(rowNumber, input){
  const SHEET = '✅';
  const COL_R = 18; // 目前報價
  const COL_AH = 34; // 更新日
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET);
  if (!sh) throw new Error('找不到工作表：' + SHEET);
  if (rowNumber < 2) throw new Error('列號不合法');

  const rngR = sh.getRange(rowNumber, COL_R);
  const s = (typeof input === 'string') ? input.trim() : '';
  if (s && s[0] === '=') rngR.setFormula(s); else rngR.setValue(Number(input)||0);

  const tz = Session.getScriptTimeZone() || 'Asia/Taipei';
  const today = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  sh.getRange(rowNumber, COL_AH).setValue(today);

  SpreadsheetApp.flush();
  return { row: rowNumber };
}

/* ===== Cashflow 彙總表（給前端 Dashboard 用） ===== */

/** 表內可用：=CF_SUMMARY()  → 含表頭與最後一列 TOTAL */
function CF_SUMMARY() {
  const sh = SpreadsheetApp.getActive().getSheetByName(CF_SHEET);
  const lastRow = sh.getLastRow();
  if (lastRow <= HEADER_ROW) return [["Account","NOW","Planned","Final"]];

  const M = getColMap_(sh);
  const vals = sh.getRange(HEADER_ROW+1, 1, lastRow-HEADER_ROW, sh.getLastColumn()).getValues();
  const iI=M.item-1, iA=M.amount-1, iACC=M.account-1, iS=M.status-1;

  const agg = {}; // acc -> {now, planned}
  for (const r of vals){
    const acc = String(r[iACC]||'').trim(); if (!acc) continue;
    const item= String(r[iI]||'').trim();
    const sta = String(r[iS]||'').trim().toLowerCase();
    const amt = Number(r[iA]||0);
    if (!agg[acc]) agg[acc] = {now:0, planned:0};
    if (item === 'Current') agg[acc].now += amt;
    else if (sta === 'posted')  agg[acc].now += -amt;
    else if (sta === 'planned') agg[acc].planned += -amt;
  }

  const rows = Object.keys(agg).sort().map(acc=>{
    const now = agg[acc].now||0, pln = agg[acc].planned||0;
    return [acc, now, pln, now+pln];
  });

  const tNow = rows.reduce((s,r)=>s+Number(r[1]||0),0);
  const tPln = rows.reduce((s,r)=>s+Number(r[2]||0),0);
  const tFin = rows.reduce((s,r)=>s+Number(r[3]||0),0);

  return [
    ["Account","NOW","Planned","Final"],
    ...rows,
    ["TOTAL", tNow, tPln, tFin]
  ];
}

/** 給前端 Dashboard 用（物件陣列） */
function getSummaryCF(){
  const t = CF_SUMMARY();                     // 含表頭 + TOTAL
  return (t||[]).slice(1).map(r=>({           // 含 TOTAL 一起回傳，前端自行決定是否顯示
    account:String(r[0]??''),
    now:Number(r[1]??0),
    planned:Number(r[2]??0),
    final:Number(r[3]??0)
  }));
}


/** 設定 **/
const RECORD_SHEET_NAME = 'record'; // 你的 sheet 名稱

/**
 * 超快速讀取 record（支援投影欄 cols、分頁 offset/limit）
 * @param {Object} opt
 * @param {number} [opt.offset=0]
 * @param {number} [opt.limit=200]
 * @param {number[]} [opt.cols]  // 1-based 欄索引，若省略→全部欄
 * @return {{headers:string[], rows:string[][], from:number, to:number, total:number, allHeaders:string[]}}
 */
/** 給 Database 虛擬表：以第2列為表頭，第3列開始是資料；支援 order: 'desc' | 'asc'（預設 desc） */
function getDatabaseRows(opt) {
  opt = opt || {};
  var offset = Math.max(0, Number(opt.offset || 0));
  var limit  = Math.max(0, Number(opt.limit  || 200));
  var order  = String(opt.order || 'desc').toLowerCase(); // 預設最新→最舊

  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(RECORD_SHEET_NAME);
  if (!sh) return { headers: [], rows: [], from: 0, to: 0, total: 0, allHeaders: [] };

  var lastRow = sh.getLastRow();
  var lastCol = sh.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return { headers: [], rows: [], from: 0, to: 0, total: 0, allHeaders: [] };

  // 表頭在第 2 列
  var headers = sh.getRange(2, 1, 1, lastCol).getDisplayValues()[0] || [];
  var dataStart = 3;

  if (lastRow < dataStart) {
    return { headers: headers, rows: [], from: 0, to: 0, total: 0, allHeaders: headers };
  }

  // 一次性把第3列到最後列都抓出來（顯示值），後面用陣列處理，避免逐列 getRange 造成「無此範圍」
  var block = sh.getRange(dataStart, 1, lastRow - dataStart + 1, lastCol).getDisplayValues();

  // 修剪尾端「工作列/空白列」→ 規則改為：整列顯示值皆空白才刪
  function isBlankRow(arr){
    return arr.every(function(v){ return String(v||'').trim()===''; });
  }

  var endIdx = block.length - 1;
  while (endIdx >= 0) {
    var rowArr = block[endIdx];
    if (!isBlankRow(rowArr)) break;  // 只要有任何顯示值就保留（即便全是公式）
    endIdx--;
  }
  if (endIdx < 0) {
    return { headers: headers, rows: [], from: 0, to: 0, total: 0, allHeaders: headers };
  }

  // 有效資料塊
  var data = block.slice(0, endIdx + 1);
  var total = data.length;

  // 順序：desc = 由新到舊（視為資料底部是最新 → 反轉）；asc = 由舊到新（保持）
  var ordered = (order === 'desc') ? data.slice().reverse() : data;

  // 切頁
  var from = Math.min(offset, total);
  var to   = Math.min(offset + limit, total);
  var rows = ordered.slice(from, to);

  return { headers: headers, rows: rows, from: from, to: to, total: total, allHeaders: headers };
}

/** 左側最新資料卡：用欄位字母定位 + 同時取值(getValues)與顯示(getDisplayValues) */
function getRecordLatest(){
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(RECORD_SHEET_NAME);
  if (!sh) return null;

  var lastRow = sh.getLastRow();
  var lastCol = sh.getLastColumn();
  if (lastRow < 3) return null;

  // 欄字母→1-based 欄號（支援 AA、BB、CV 等）
  function colLetter(s){
    s = String(s||'').trim().toUpperCase();
    var n = 0;
    for (var i=0;i<s.length;i++){
      var code = s.charCodeAt(i);
      if (code >= 65 && code <= 90) n = n*26 + (code - 64);
    }
    return n;
  }

  // 直接指定欄位：A/C/D/E/F/G/BA/CU   ← AQ 移除後向左位移一欄
    var cDate = colLetter('A');
    var cWhole= colLetter('C');
    var cWithDad=colLetter('D');
    var cAvail= colLetter('E');
    var cPnL  = colLetter('F');  // ★ 當日損益
    var cCash = colLetter('G');
    var cDebt = colLetter('BA'); // 原 BB 因移除 AQ 改為 BA
    var cUSD  = colLetter('CU'); // 原 CV 因移除 AQ 改為 CU

  // 判斷是否為空列：關鍵欄位全都空 → 才算空
  function isEmptyRow(r){
    var disp = sh.getRange(r, 1, 1, lastCol).getDisplayValues()[0];
    var keys = [cDate,cWhole,cWithDad,cAvail,cPnL,cCash,cDebt,cUSD];
    return keys.every(function(c){
      return String(disp[c-1] || '').trim() === '';
    });
  }

  var rowIdx = lastRow;
  while (rowIdx >= 3 && isEmptyRow(rowIdx)) rowIdx--;
  if (rowIdx < 3) return null;

  // 同時抓「原始值」與「顯示字串」
  var rVal  = sh.getRange(rowIdx, 1, 1, lastCol).getValues()[0];
  var rDisp = sh.getRange(rowIdx, 1, 1, lastCol).getDisplayValues()[0];

  // 將顯示字串轉數字（支援 1,234 / 1 234 / 1，234 / (1,234) / 12.3%）
  function toNum(vVal, vDisp){
    // 原始值若為 number，直接用
    if (typeof vVal === 'number' && isFinite(vVal)) return vVal;

    var s = String(vDisp || '').trim();
    if (!s) return 0;

    // 括號負數
    var neg = false;
    if (/^\(.*\)$/.test(s)){ neg = true; s = s.slice(1,-1); }

    // 移除各種空白（含 NBSP）與各地區千分位逗號/全形逗號
    s = s.replace(/[\u00A0\s,，]/g,'');

    // 百分比
    var isPct = /%$/.test(s);
    if (isPct) s = s.replace(/%$/,'');

    var n = Number(s);
    if (!isFinite(n)) return 0;
    if (isPct) n = n / 100;
    if (neg) n = -n;
    return n;
  }

  // 有沒有任何關鍵值
  function hasVal(v){ return String(v||'').trim() !== ''; }
  var keyHasValue =
    hasVal(rDisp[cWhole-1]) || hasVal(rDisp[cWithDad-1]) || hasVal(rDisp[cAvail-1]) ||
    hasVal(rDisp[cCash-1])  || hasVal(rDisp[cDebt-1])    || hasVal(rDisp[cUSD-1]) ||
    hasVal(rDisp[cDate-1]);
  if (!keyHasValue) return null;

  // 日期輸出：優先 Date，其次顯示字串
  var dVal  = rVal[cDate-1];
  var dDisp = rDisp[cDate-1];
  var tz    = SpreadsheetApp.getActive().getSpreadsheetTimeZone() || 'Asia/Taipei';
  var outDate = (dVal instanceof Date)
    ? Utilities.formatDate(dVal, tz, 'yyyy-MM-dd')
    : String(dDisp || '');

  return {
    date: outDate,
    C:   toNum(rVal[cWhole-1],   rDisp[cWhole-1]),
    D:   toNum(rVal[cWithDad-1], rDisp[cWithDad-1]),
    E:   toNum(rVal[cAvail-1],   rDisp[cAvail-1]),
    F:   toNum(rVal[cPnL-1],     rDisp[cPnL-1]),
    G:   toNum(rVal[cCash-1],    rDisp[cCash-1]),
    BA:  toNum(rVal[cDebt-1],    rDisp[cDebt-1]),   // 債務（內部計算用）
    CU:  toNum(rVal[cUSD-1],     rDisp[cUSD-1])    // 持有 USD
  };
}

/** 一次回傳 Dashboard 基礎資料（Database 最新 + Cashflow TOTAL + CTBC 總覽） */
function getDashboardData(){
  // --- Database 最新（已有 getRecordLatest）
  const latest = getRecordLatest() || { date:'', C:0, D:0, E:0, G:0, BA:0, CU:0 };
  const kpi_total = Number(latest.C||0);
  const kpi_debt  = Number(latest.BA||0); // 內部用於淨資產計算
  const kpi_cash  = Number(latest.G||0);
  const kpi_avail = Number(latest.E||0);
  const kpi_usd   = Number(latest.CU||0);
  const kpi_net   = Number(latest.D||0); // 以 D 欄 Whole Assets 作為「淨資產」顯示

  // --- Cashflow TOTAL（已有 getSummaryCF：含 TOTAL）
  const cfRows = getSummaryCF() || [];
  const totalRow = cfRows.find(r => String(r.account||'').toUpperCase()==='TOTAL') || { now:0, planned:0, final:0 };
  const cfTotal = { now:Number(totalRow.now||0), planned:Number(totalRow.planned||0), final:Number(totalRow.final||0) };

  // 取近期 5 筆 planned（給 Dashboard 卡片列）
  const plannedList = (getTxByKind('planned')||[]).slice(0,5);

  // --- CTBC：以 getCTBCInvestments 累加 USD 值
  const ct = getCTBCInvestments() || [];
  const all = ct.reduce((a,r)=>{
    const v = Number(r.N||0);    // 現值 (USD)
    const pnl = Number(r.T||0);  // 含息損益 (USD)
    const intv= Number(r.AF||0); // 已領利息 (USD)
    a.value+=v; a.pnl+=pnl; a.interest+=intv; return a;
  }, {value:0, pnl:0, interest:0});
  const ret = all.value!==0 ? (all.pnl/all.value)*100 : 0;

  return {
    kpis: {
      date: latest.date || '',
      total: kpi_total,
      net:   kpi_net,
      cash:  kpi_cash,
      avail: kpi_avail,
      usd:   kpi_usd
    },
    cashflow: {
      total: cfTotal,
      plannedTop: plannedList
    },
    ctbc: {
      allUSD: { value: all.value, pnl: all.pnl, ret, interest: all.interest }
    },
    database: {
      latest // {date,C,D,E,G,BA,CU}
    }
  };
}

/** ====== 基本設定：Dash 分頁、欄位位置 ====== */
const DASH_SHEET_NAME = 'Dash';
const START_ROW = 3;      // A3 起算
const NAME_COL  = 1;      // A: 帳戶名
const ORIG_COL  = 2;      // B: 原幣（這裡是我們要寫回的目標欄）
const TWD_COL   = 3;      // C: 換算新台幣（通常由公式帶出，不直接寫）

/** 工具：找到最後一筆帳戶列（A 欄連續非空白，遇到空白停） */
function _dash_lastRow_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(DASH_SHEET_NAME);
  const maxRows = sh.getMaxRows();
  for (let r = START_ROW; r <= maxRows; r++) {
    const v = sh.getRange(r, NAME_COL).getDisplayValue();
    if (!v) return r - 1; // 空白就回前一列
  }
  return maxRows;
}

/** 工具：從名稱推測幣別（供前端徽章顯示提示用） */
function _inferCurrency_(name) {
  const s = String(name || '').toUpperCase();
  if (/HKD/.test(s)) return 'HKD';
  if (/USD/.test(s)) return 'USD';
  if (/VND/.test(s)) return 'VND';
  if (/JPY/.test(s)) return 'JPY';
  if (/KRW/.test(s)) return 'KRW';
  if (/(NT|NTD|TWD)/.test(s)) return 'TWD';
  return '';
}

/** 讀取帳戶清單（對應 Dash!A3:C…），無 ID → 以「工作表列號」當 ID */
function getAccountsLite() {
  const sh = SpreadsheetApp.getActive().getSheetByName(DASH_SHEET_NAME);
  const last = _dash_lastRow_();
  if (last < START_ROW) return [];

  const rng = sh.getRange(START_ROW, NAME_COL, last - START_ROW + 1, 3);
  const values = rng.getValues();              // [[name, orig, twd], ...]
  const display = rng.getDisplayValues();      // 保留 $/逗號格式（若你需要）
  const res = [];

  for (let i = 0; i < values.length; i++) {
    const rowNo = START_ROW + i;
    const name = values[i][0];
    if (!name) continue;
    // 跳過小計/總計列（例如「總持有美金」）
    if (/^總/.test(String(name))) continue;

    const origNum = Number(values[i][1]) || 0; // 原幣數字
    const cur = _inferCurrency_(name);

    res.push({
      ID: String(rowNo),              // 以「列號」充當 ID
      AccountName: name,
      Currency: cur,                  // 只做提示用
      Balance: origNum                // 前端就是編輯這個欄位（B 欄）
    });
  }
  return res;
}

/** 寫入餘額：將數值寫回 Dash 的「原幣（B 欄）」；ID=列號 */
function setBalance(payload) {
  const { id, value } = payload; // value: number（前端已處理試算）
  const sh = SpreadsheetApp.getActive().getSheetByName(DASH_SHEET_NAME);
  const rowNo = Number(id);
  if (!rowNo || rowNo < START_ROW) throw new Error('Bad id/row: ' + id);

  // 寫回原幣（B 欄）
  sh.getRange(rowNo, ORIG_COL).setValue(Number(value) || 0);

  // 讓公式連動（例如 C 欄、Dash!B2/E3 等）
  SpreadsheetApp.flush();
  const dash = getDashNumbers();
  return { ok: true, saved: Number(value) || 0, dash };
}

/* ===== 連結清單（給前端管理用） ===== */
function saveLinkNote(row, note){
  const sh = getLinksSheet_();
  if (row < 2 || row > sh.getLastRow()) throw new Error('列號超出範圍');
  sh.getRange(row, 3).setValue(note || '');
  return true;
}

/** （可選）=公式交給 Sheets 引擎算，例如 =SUM(100, 200) → 回計算後的數字 */
function evalBySheetsEngine(expr) {
  if (!/^=/.test(expr)) throw new Error('Not a formula');
  const lock = LockService.getScriptLock(); lock.tryLock(5000);
  try {
    const ss = SpreadsheetApp.getActive();
    const sh = ss.getSheetByName('_EvalTmp') || ss.insertSheet('_EvalTmp');
    const cell = sh.getRange('A1');
    cell.setFormula(expr);
    SpreadsheetApp.flush();
    const val = cell.getValue();
    cell.clearContent();
    return { ok: true, value: Number(val) };
  } finally { lock.releaseLock(); }
}

/** Dash 指標：B1=總金額、E3=今日增減 */
function getDashNumbers() {
  const sh = SpreadsheetApp.getActive().getSheetByName('Dash');
  return {
    total: Number(sh.getRange('B1').getValue()) || 0,
    today: Number(sh.getRange('E3').getValue()) || 0
  };
}

/**
 * 覆寫第 row 列的交易（Cashflow!A-F = date/item/Amount/Account/Status/remark）
 * @param {number} row  試算表實際列號（含表頭，通常 >= 2）
 * @param {Object} rec  {date,item,amount,account,status,note}
 */

function updateTransaction(row, rec){
  const sh = SpreadsheetApp.getActive().getSheetByName(CF_SHEET);
  if (!sh) throw new Error('找不到工作表：' + CF_SHEET);

  if (!row || row < HEADER_ROW + 1) throw new Error('row 不合法（需為資料列，通常 ≥ ' + (HEADER_ROW+1) + '）');

  const M = getColMap_(sh);
  const map = {
    date:    M.date,
    item:    M.item,
    amount:  M.amount,
    account: M.account,
    status:  M.status,
    note:    M.note
  };

  // 轉型
  const d = rec.date ? new Date(rec.date) : '';
  const item = String(rec.item||'').trim();
  const amt  = rec.amount === '' || rec.amount == null ? '' : Number(rec.amount);
  const acc  = String(rec.account||'').trim();
  let   st   = String(rec.status||'').trim();
  const note = String(rec.note||'').trim();

  // 寫入
  if (map.date)    sh.getRange(row, map.date).setValue(d);
  if (map.item)    sh.getRange(row, map.item).setValue(item);
  if (map.amount)  sh.getRange(row, map.amount).setValue(amt);
  if (map.account) sh.getRange(row, map.account).setValue(acc);
  if (map.note)    sh.getRange(row, map.note).setValue(note);

  if (map.status) {
    const stCell = sh.getRange(row, map.status);
    if (!st || /^auto$/i.test(st)) {
      const delta = (map.date || 0) - (map.status || 0);
      stCell.setFormulaR1C1(`=IF(RC[${delta}]<=TODAY(),"posted","planned")`);
    } else if (/^posted$/i.test(st)) {
      stCell.clearFormat().setValue('posted');
    } else if (/^planned$/i.test(st)) {
      stCell.clearFormat().setValue('planned');
    } else {
      const delta = (map.date || 0) - (map.status || 0);
      stCell.setFormulaR1C1(`=IF(RC[${delta}]<=TODAY(),"posted","planned")`);
    }
  }
  SpreadsheetApp.flush();       // ★ 讓前端下一次讀到最新值
  return { ok:true, row };
}

/** 刪除指定列（Cashflow!A:F 的整列）
 * @param {number} row 真實試算表列號（含表頭，通常 >= 2）
 * @return {{ok:boolean,row:number}}
 */
function deleteTransaction(row){
  const sh = SpreadsheetApp.getActive().getSheetByName(CF_SHEET);
  if (!sh) throw new Error('找不到工作表：' + CF_SHEET);
  const last = sh.getLastRow();
  const nRow = Number(row);
  if (!nRow || nRow < HEADER_ROW + 1 || nRow > last) {
    throw new Error('row 不合法（需為資料列，通常 ≥ ' + (HEADER_ROW+1) + ' 且 ≤ ' + last + '）：' + row);
  }
  sh.deleteRow(nRow);
  SpreadsheetApp.flush();
  return { ok:true, row:nRow };
}


/** 保單成長：讀 A, I, J, L, N, S 欄（第 11 列為表頭） */
function getPolicyGrowth(){
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('保單成長');
  if (!sh) throw new Error('找不到「保單成長」分頁');

  const last = sh.getLastRow();
  if (last < 12) return { headers:['年/月','年齡','紅利模型','累計提領','Total P/L','累計投入'], rows:[], years:[] };

  // 指定欄位（從第 12 列開始是資料）
  const rA = sh.getRange(12, 1, last-11, 1).getValues();   // 年份（A）
  const rI = sh.getRange(12, 9, last-11, 1).getValues();   // 累計投入（I）
  const rJ = sh.getRange(12,10, last-11, 1).getValues();   // 累計提領（J）
  const rL = sh.getRange(12,12, last-11, 1).getValues();   // Total P/L（L）
  const rN = sh.getRange(12,14, last-11, 1).getValues();   // 紅利模型（N）
  const rS = sh.getRange(12,19, last-11, 1).getValues();   // 年齡（S）

  // 轉成 rows
  const rows = [];
  const years = [];
  for (let i=0; i<rA.length; i++){
    const y = rA[i][0];
    const yyMM = formatYYMM(y);         // 需求 4：顯示 25/08 格式
    years.push(yyMM);
    rows.push([
      yyMM,                              // 0 年/月（顯示用）
      Number(rS[i][0]) || null,          // 1 年齡
      String(rN[i][0] ?? ''),            // 2 紅利模型
      Number(rJ[i][0]) || 0,             // 3 累計提領
      Number(rL[i][0]) || 0,             // 4 Total P/L
      Number(rI[i][0]) || 0              // 5 累計投入（新增）
    ]);
  }

  return {
    headers: ['年/月','年齡','紅利模型','累計提領','Total P/L','累計投入'],
    rows,
    years
  };
}

/** A 欄可能是日期或數字，統一回傳 yy/MM */
function formatYYMM(v){
  let d;
  if (v instanceof Date) d = v;
  else if (typeof v === 'number') {
    // 既可能是 Excel 序號也可能是 202508 這類整數
    // 先試序號（Google 表單序號起點 1899-12-30）
    const asDate = new Date(Math.round((v - 25569) * 86400 * 1000));
    if (!isNaN(asDate.getTime()) && asDate.getFullYear() > 1970) d = asDate;
    else {
      // 再試 202508 / 2025-08 這類
      const s = String(v);
      if (/^\d{6}$/.test(s)) { // YYYYMM
        d = new Date(Number(s.slice(0,4)), Number(s.slice(4,6))-1, 1);
      }
    }
  }
  if (!d) {
    const s = String(v||'').trim().replaceAll('-', '/').replaceAll('.', '/');
    const tryD = new Date(s);
    if (!isNaN(tryD.getTime())) d = tryD;
  }
  if (!d) return String(v||'--');
  const yy = String(d.getFullYear()).slice(-2);
  const mm = String(d.getMonth()+1).padStart(2,'0');
  return `${yy}/${mm}`;
}

/** 讀 A1:B10 控制項（直接回傳數值陣列，前端顯示，不再跳轉） */
function getPolicyControlsRange(){
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('保單成長');
  if (!sh) throw new Error('找不到「保單成長」分頁');
  return sh.getRange('A1:B10').getDisplayValues(); // 保留你表格上的格式
}



/**====================record=========================*/

/** 前端按鈕用：呼叫「凍結最後一列 + 新增骨架」 */
function runRecordAddRow(){
  var name = (typeof CONFIG !== 'undefined' && CONFIG.recordSheet) ? CONFIG.recordSheet : 'record';
  addRowAndSnapshot(name);  // 你已經實作好的主程式
  return { ok:true, msg:'已新增一列（H~末欄轉值、A 寫時間），並在下一列鋪回公式骨架。' };
}

/**===================links========================== */

// —— Links 後端 API ——
// 資料表：sheet 名稱 "links"，A: title, B: url
function getLinksSheet_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('links') || ss.insertSheet('links');
  const firstRow = sh.getRange(1,1,1,3).getValues()[0];
  const isHeaderMissing = !firstRow[0] && !firstRow[1] && !firstRow[2];
  if (isHeaderMissing) sh.getRange(1,1,1,3).setValues([['title','url','note']]);
  return sh;
}

function getLinks() {
  const sh = getLinksSheet_();
  const last = sh.getLastRow();
  if (last <= 1) return [];
  const values = sh.getRange(2, 1, last - 1, 3).getValues(); // A:C
  const out = [];
  for (let i = 0; i < values.length; i++) {
    const [title, url, note] = values[i];
    if (!url) continue;
    const row = i + 2;
    const domain = tryGetDomain_(url);
    out.push({
      row, title: title || '', url, note: note || '',
      domain,
      favicon: domain ? ('https://www.google.com/s2/favicons?sz=64&domain=' + encodeURIComponent(domain)) : '',
      thumb: ''
    });
  }
  return out;
}

function addLink(payload) {
  const url = (payload && payload.url || '').trim();
  let title = (payload && payload.title || '').trim();
  if (!url) throw new Error('缺少網址');

  const meta = safeProbe_(url);
  if (!title) title = meta.title || url;

  const sh = getLinksSheet_();
  const row = sh.getLastRow() + 1;
  sh.getRange(row, 1, 1, 3).setValues([[title, url, ""]]); // A,B,C

  const domain = tryGetDomain_(url);
  return {
    row, title, url, note: '',
    domain,
    favicon: domain ? ('https://www.google.com/s2/favicons?sz=64&domain=' + encodeURIComponent(domain)) : '',
    thumb: meta.image || ''
  };
}

function removeLink(row) {
  const sh = getLinksSheet_();
  const last = sh.getLastRow();
  if (row < 2 || row > last) throw new Error('列號超出範圍');
  sh.deleteRow(row);
  return true;
}

// 用於新增彈窗貼上網址自動抓標題
function probeUrlMeta(url) {
  const meta = safeProbe_(url);
  return { title: meta.title || '', image: meta.image || '' };
}

// —— Helpers —— 
function tryGetDomain_(u) {
  try { return (new URL(u)).hostname; } catch (e) { return ''; }
}

function safeProbe_(u) {
  try {
    const res = UrlFetchApp.fetch(u, { muteHttpExceptions: true, followRedirects: true, timeout: 10000 });
    const html = res.getContentText() || '';
    return {
      title: parseTitle_(html) || parseOG_(html, 'og:title') || parseMetaName_(html, 'title') || '',
      image: parseOG_(html, 'og:image') || ''
    };
  } catch (e) {
    // 失敗就回空
    return { title: '', image: '' };
  }
}

function parseTitle_(html) {
  const m = html.match(/<title[^>]*>([\s\S]*?)<\/title>/i);
  if (m && m[1]) return sanitize_(m[1]);
  return '';
}
function parseOG_(html, property) {
  const re = new RegExp(`<meta[^>]+property=["']${property}["'][^>]+content=["']([^"']+)["']`, 'i');
  const m = html.match(re);
  return m && m[1] ? m[1] : '';
}
function parseMetaName_(html, name) {
  const re = new RegExp(`<meta[^>]+name=["']${name}["'][^>]+content=["']([^"']+)["']`, 'i');
  const m = html.match(re);
  return m && m[1] ? m[1] : '';
}
function sanitize_(s) {
  return s.replace(/\s+/g,' ').trim();
}



/** === STK 讀取：A..V === */
function getSTKRows(){
  const SHEET = 'STK';
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET);
  if (!sh) return [];
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow <= 1) return [];

  const n = lastRow - 1;
  const values = sh.getRange(2, 1, n, Math.min(22, lastCol)).getValues(); // V=22
  const tz = Session.getScriptTimeZone() || 'Asia/Taipei';
  const toStr = v => v instanceof Date ? Utilities.formatDate(v, tz, 'yyyy-MM-dd') : (v==null?'':String(v));

  return values.map((r,i)=>({
    ROW: i+2,
    A:r[0],  B:r[1],  C:num(r[2]),  D:num(r[3]),  E:num(r[4]),
    F:num(r[5]), G:num(r[6]), H:r[7], I:r[8],
    J:num(r[9]), K:num(r[10]),
    L:toStr(r[11]), M:toStr(r[12]), N:num(r[13]),
    O:num(r[14]), P:num(r[15]), Q:num(r[16]),
    R:num(r[17]), S:toStr(r[18]), T:toStr(r[19]), U:toStr(r[20]), V:num(r[21])
  }));

  function num(x){ x = Number(x); return isFinite(x)? x : 0; }
}

/** STK 偵錯：確認工作表/列數/欄數＋前 5 列樣本（顯示值） */
function getSTKDebug(){
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('STK');
  if (!sh) return { sheet:false, lastRow:0, lastCol:0, hasData:false, sample:[] };
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  const hasData = lastRow > 1;
  const rowsToRead = Math.min(5, Math.max(0, lastRow - 1));
  const colsToRead = Math.min(22, lastCol);
  const sample = rowsToRead>0 && colsToRead>0
    ? sh.getRange(2, 1, rowsToRead, colsToRead).getDisplayValues()
    : [];
  return { sheet:true, lastRow, lastCol, hasData, sample };
}

/** === STK 單格寫入通用：可選擇順便寫「更新日」(S 欄) ===
 * @param {number} row           - 實際列號（含表頭從第2列開始，所以通常 >=2）
 * @param {string} colLetter     - 欄位字母（例如 'E' 或 'R'）
 * @param {number|string} value  - 要寫入的值（數字或字串都可）
 * @param {string|null} updateDateColLetter - 若提供（例如 'S'），會一併把該欄寫成今天 yyyy-MM-dd
 */
function setSTKValue(row, colLetter, value, updateDateColLetter){
  const SHEET = 'STK';
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET);
  if (!sh) throw new Error('找不到工作表：' + SHEET);
  if (!row || row < 2) throw new Error('列號不合法');

  const col = colIndex(colLetter);
  sh.getRange(row, col).setValue(value);

  if (updateDateColLetter){
    const tz = Session.getScriptTimeZone() || 'Asia/Taipei';
    const today = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
    const uCol = colIndex(updateDateColLetter);
    sh.getRange(row, uCol).setValue(today);
  }
  return true;

  function colIndex(letter){
    let s = String(letter||'').trim().toUpperCase(), n=0;
    for (let i=0;i<s.length;i++){ const code = s.charCodeAt(i); if (code>=65 && code<=90) n = n*26 + (code-64); }
    return n;
  }
}

/** === 便利包：專寫現價(E)、股息(R) === */
function setSTKPrice(row, price){
  return setSTKValue(row, 'E', Number(price) || 0, 'S');
}
function setSTKDividend(row, divd){
  return setSTKValue(row, 'R', Number(divd) || 0, 'S');
}

/** 新增一列到 STK，未列欄位沿用上一列公式（若無上一列，僅寫入值） */
function addSTKItem(item){
  const SHEET = 'STK';
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET);
  if (!sh) throw new Error('找不到工作表：' + SHEET);

  var needCols = 22; // A..V
  if (sh.getMaxColumns() < needCols){
    sh.insertColumnsAfter(sh.getMaxColumns(), needCols - sh.getMaxColumns());
  }

  var lastRow = sh.getLastRow();
  var dataStart = 2;

  // 插入新列在最後一列之後
  var insertAfter = Math.max(1, lastRow);
  sh.insertRowAfter(insertAfter);
  var targetRow = insertAfter + 1;

  // 若有模板列（上一列或第 2 列），先複製其格式/公式
  var sourceRow = (lastRow >= dataStart) ? lastRow : dataStart;
  if (sourceRow >= dataStart && sourceRow <= sh.getMaxRows()){
    sh.getRange(sourceRow, 1, 1, needCols).copyTo(
      sh.getRange(targetRow, 1, 1, needCols),
      {contentsOnly:false}
    );
  }

  // A 狀態清空（未賣出）、M 賣日清空
  sh.getRange(targetRow, col('A')).setValue(false);
  sh.getRange(targetRow, col('M')).clearContent();

  // 寫入主要欄位
  if (item.B) sh.getRange(targetRow, col('B')).setValue(item.B);
  if (item.C!=null) sh.getRange(targetRow, col('C')).setValue(Number(item.C)||0);
  if (item.D!=null) sh.getRange(targetRow, col('D')).setValue(Number(item.D)||0);
  if (item.T) sh.getRange(targetRow, col('T')).setValue(String(item.T));
  if (item.R!=null) sh.getRange(targetRow, col('R')).setValue(Number(item.R)||0);
  if (item.L) {
    var tz = Session.getScriptTimeZone() || 'Asia/Taipei';
    var d = new Date(item.L);
    sh.getRange(targetRow, col('L')).setValue(Utilities.formatDate(d, tz, 'yyyy-MM-dd'));
  }
  if (item.U) sh.getRange(targetRow, col('U')).setValue(String(item.U));

  // E 現價：強制改為指定公式（依列號帶入）
  var r = targetRow;
  var formula = '=IF(T'+r+'="US",IF(A'+r+'=TRUE,"sold price",googlefinance(B'+r+',"PRICE")),IF(A'+r+'=TRUE,"sold price",vlookup(B'+r+",'🔄️'!B:E,3,0)))";
  sh.getRange(targetRow, col('E')).setFormula(formula);

  return { row: targetRow };

  function col(letter){
    var s = String(letter||'').trim().toUpperCase(), n=0;
    for (var i=0;i<s.length;i++){ var code = s.charCodeAt(i); if (code>=65 && code<=90) n = n*26 + (code-64); }
    return n;
  }
}

/** 依據賣出數量處理賣出；若部分賣出，會拆分一筆新列保留剩餘股數 */
function splitSell(row, qty, price){
  const SHEET = 'STK';
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET);
  if (!sh) throw new Error('找不到工作表：' + SHEET);
  var r = Number(row);
  if (!r || r<2) throw new Error('row 不合法');

  var needCols = 22;
  if (sh.getMaxColumns() < needCols){
    sh.insertColumnsAfter(sh.getMaxColumns(), needCols - sh.getMaxColumns());
  }

  var currentQty = Number(sh.getRange(r, col('C')).getValue()) || 0;
  if (qty > currentQty) throw new Error('股數不可大於現有股數');

  // 全部賣出：標記 A=已賣出、E=賣價、M=今天
  var tz = Session.getScriptTimeZone() || 'Asia/Taipei';
  var today = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');

  if (qty === currentQty){
    sh.getRange(r, col('A')).setValue(true);
    sh.getRange(r, col('E')).setValue(Number(price)); // 固定賣價
    sh.getRange(r, col('M')).setValue(today);
    return { row:r, type:'full' };
  }

  // 部分賣出：原列改為賣出 qty；新列保留剩餘
  var remain = currentQty - qty;

  // 原列寫入賣出資訊
  sh.getRange(r, col('A')).setValue(true);
  sh.getRange(r, col('C')).setValue(Number(qty));
  sh.getRange(r, col('E')).setValue(Number(price));
  sh.getRange(r, col('M')).setValue(today);

  // 插入新列，複製原列格式/公式，再寫入剩餘股數，並清空賣出標記/賣日，現價改回公式
  sh.insertRowAfter(r);
  var nr = r+1;
  sh.getRange(r, 1, 1, needCols).copyTo(sh.getRange(nr, 1, 1, needCols), {contentsOnly:false});

  // 新列：未賣出、股數=remain、賣日清空
  sh.getRange(nr, col('A')).setValue(false);
  sh.getRange(nr, col('C')).setValue(Number(remain));
  sh.getRange(nr, col('M')).clearContent();

  // 新列：現價公式
  var formula = '=IF(T'+nr+'="US",IF(A'+nr+'=TRUE,"sold price",googlefinance(B'+nr+',"PRICE")),IF(A'+nr+'=TRUE,"sold price",vlookup(B'+nr+",'🔄️'!B:E,3,0)))";
  sh.getRange(nr, col('E')).setFormula(formula);

  return { row:r, remainRow:nr, type:'partial' };

  function col(letter){
    var s = String(letter||'').trim().toUpperCase(), n=0;
    for (var i=0;i<s.length;i++){ var code = s.charCodeAt(i); if (code>=65 && code<=90) n = n*26 + (code-64); }
    return n;
  }
}
/** 刪除 STK 指定列（含表頭起算），row>=2 */
function deleteSTKRow(row){
  const SHEET = 'STK';
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET);
  if (!sh) throw new Error('找不到工作表：' + SHEET);
  const last = sh.getLastRow();
  const r = Number(row);
  if (!r || r < 2 || r > last) throw new Error('列號超出範圍');
  sh.deleteRow(r);
  return { ok:true, row:r };
}