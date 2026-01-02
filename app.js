// ---- helpers for parsing and indexing ----
function parseMaybeString(v) {
  try {
    return typeof v === 'string' ? JSON.parse(v) : v;
  } catch (e) {
    console.warn('[parseMaybeString] parse failed, returning original', e);
    return v;
  }
}
function clamp(n, min, max) {
  if (Number.isNaN(n)) return min;
  return Math.max(min, Math.min(n, max));
}
// -----------------------------------------
/** ========= Cody Wealth WebApp: server (app.gs) =========
 *  動態表頭整合版（A1 為表頭）
 *  - 全部讀寫一律用 getColMap_() 取欄位位置（不再寫死欄位號）
 *  - NOW：Current + posted(-金額)，planned 不算進 NOW
 *  - submitCashflow：status=auto → 依當列日期產生 R1C1 公式
 */

const CF_SHEET = 'Cashflow'; // 資料分頁
const HEADER_ROW = 1;          // 表頭列（A1）
const COLS = {                 // 後備：找不到表頭時的預設欄位（A..F）
  date: 1, item: 2, amount: 3, account: 4, status: 5, note: 6
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
function getMeta() {
  const sh = SpreadsheetApp.getActive().getSheetByName(CF_SHEET);
  if (!sh) return { accounts: [], statuses: ['auto', 'posted', 'planned'] };

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
    statuses: ['auto', 'posted', 'planned', 'using', ...Array.from(uniqueStatuses)],
  };
}

/** Dashboard：依帳戶彙總 NOW（Current + posted(-金額)） */
function getSummaryNow() {
  const map = calcNow_(); // {acc: now}
  return Object.keys(map).sort().map(a => ({ account: a, now: map[a] }));
}

/** Dashboard：取 planned 清單（日期、帳戶、金額、明細） */
function getPlannedList() {
  const sh = SpreadsheetApp.getActive().getSheetByName(CF_SHEET);
  if (!sh) return [];
  const lastRow = sh.getLastRow();
  if (lastRow <= HEADER_ROW) return [];

  const n = lastRow - HEADER_ROW;
  const vals = sh.getRange(HEADER_ROW + 1, 1, n, sh.getLastColumn()).getValues();
  const map = getColMap_(sh);
  const iD = map.date - 1, iI = map.item - 1, iA = map.amount - 1, iACC = map.account - 1, iS = map.status - 1;

  const out = [];
  const tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone() || 'Asia/Taipei';

  for (const r of vals) {
    if (!r[iS]) continue;
    const sta = String(r[iS]).trim().toLowerCase();
    if (sta !== 'planned') continue;

    const d = r[iD] instanceof Date ? r[iD] : (r[iD] ? new Date(r[iD]) : null);
    out.push({
      date: d ? Utilities.formatDate(d, tz, 'yyyy-MM-dd') : '',
      account: String(r[iACC] || '').trim(),
      amount: Number(r[iA] || 0) * -1, // planned 顯示為 -金額
      item: String(r[iI] || '').trim()
    });
  }
  out.sort((a, b) => String(a.date).localeCompare(String(b.date)));
  return out;
}

/** 
 * History Consolidation (User Triggered)
 * 1. Find account's existing 'Current' snapshot (or create one).
 * 2. Sum all 'posted' transactions up to CUTOFF DATE (inclusive).
 * 3. Update 'Current' amount = OldSnapshot + Sum(-Posted).
 * 4. Update 'Current' date = Cutoff Date.
 * 5. Delete consolidated rows.
 */
function consolidateAccount(accName, cutoffDateStr) {
  if (!accName) throw new Error('Account name is required');
  if (!cutoffDateStr) throw new Error('Cutoff date is required');

  const sh = SpreadsheetApp.getActive().getSheetByName(CF_SHEET);
  if (!sh) throw new Error('Sheet not found');

  const tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone() || 'Asia/Taipei';

  // Logic: consolidate items where date <= cutoffDateStr.
  // Safest: compare timestamps.
  const cutDate = new Date(cutoffDateStr || '');
  // Set to end of day to ensuring full coverage (23:59:59.999)
  if (isNaN(cutDate.getTime())) throw new Error('Invalid cutoff date format');
  cutDate.setHours(23, 59, 59, 999);
  const cutTime = cutDate.getTime();

  const lastRow = sh.getLastRow();
  if (lastRow <= HEADER_ROW) return { ok: true, msg: 'No data to consolidate' };

  // Read all data
  const range = sh.getRange(HEADER_ROW + 1, 1, lastRow - HEADER_ROW, sh.getLastColumn());
  const vals = range.getValues();
  const M = getColMap_(sh);
  const iItem = M.item - 1, iAmt = M.amount - 1, iAcc = M.account - 1, iSta = M.status - 1, iDate = M.date - 1;

  // Find targets
  let snapRowIdx = -1;  // 0-based relative to vals
  let snapAmt = 0;
  const toDelete = [];
  let sumPosted = 0;
  let count = 0;

  for (let i = 0; i < vals.length; i++) {
    const r = vals[i];
    const rAcc = String(r[iAcc] || '').trim();
    if (rAcc !== accName) continue;

    const item = String(r[iItem] || '').trim();
    const sta = String(r[iSta] || '').trim().toLowerCase();

    // Check Date
    const rDate = r[iDate] instanceof Date ? r[iDate] : (r[iDate] ? new Date(r[iDate]) : null);

    // Snapshot?
    if (item === 'Current') {
      snapRowIdx = i;
      snapAmt = Number(r[iAmt] || 0);
      continue;
    }

    // Target for consolidation: Posted AND Date <= Cutoff (Timestamp compare)
    // Ignore if date is empty or invalid
    if (sta === 'posted' && rDate) {
      // Comparison
      const rTime = rDate.getTime();
      if (!isNaN(rTime) && rTime <= cutTime) {
        sumPosted += Number(r[iAmt] || 0);
        toDelete.push(i);
        count++;
      }
    }
  }

  if (count === 0) {
    return { ok: true, msg: 'No transactions found to consolidate (up to ' + yesterStr + ')' };
  }

  // Execute update
  let finalSnapAmt = snapAmt - sumPosted; // Expense(pos) reduces snapshot

  // 1. Update or Create Snapshot
  // Note: If no snapshot existed, start from 0 - sumPosted
  if (snapRowIdx >= 0) {
    const absRow = HEADER_ROW + 1 + snapRowIdx;
    sh.getRange(absRow, iAmt + 1).setValue(finalSnapAmt);
    // Update date to cutoff date (just store the raw string or parsed date)
    sh.getRange(absRow, iDate + 1).setValue(cutoffDateStr);
  } else {
    // New row
    const newRow = [];
    // initialize empty
    for (let k = 0; k < sh.getLastColumn(); k++) newRow.push('');
    newRow[iDate] = cutoffDateStr;
    newRow[iItem] = 'Current';
    newRow[iAmt] = finalSnapAmt;
    newRow[iAcc] = accName;
    newRow[iSta] = 'posted';
    sh.appendRow(newRow);
  }

  // 2. Delete rows (from bottom to top to preserve indices)
  // Convert relative index -> absolute row number
  const absRowsToDelete = toDelete.map(idx => HEADER_ROW + 1 + idx).sort((a, b) => b - a);

  // Optimization: delete one by one is slow. Direct Sheets API is better but here we use simple loop.
  // Or check if contiguous? Let's assume standard usage.
  for (const rNum of absRowsToDelete) {
    sh.deleteRow(rNum);
  }

  return { ok: true, msg: `Consolidated ${count} rows. New Start Balance: ${finalSnapAmt}` };
}

/** 依種類取清單：kind = 'planned' | 'now' | 'final'
 *  planned: 只取 status=planned
 *  now:     只取 status=posted（排除 item='Current'）
 *  final:   posted + planned（都排除 Current）
 *  輸出：[{date, account, amount, item}]
 */
function getTxByKind(kind) {
  kind = String(kind || '').toLowerCase();
  if (!['planned', 'now', 'final'].includes(kind)) kind = 'planned';

  const sh = SpreadsheetApp.getActive().getSheetByName(CF_SHEET);
  const lastRow = sh.getLastRow();
  if (lastRow <= HEADER_ROW) return [];

  const n = lastRow - HEADER_ROW;
  const vals = sh.getRange(HEADER_ROW + 1, 1, n, sh.getLastColumn()).getValues();
  const M = getColMap_(sh);
  const iD = M.date - 1, iI = M.item - 1, iA = M.amount - 1, iACC = M.account - 1, iS = M.status - 1;

  const out = [];
  for (const r of vals) {
    const item = String(r[iI] || '').trim();
    const sta = String(r[iS] || '').trim().toLowerCase();
    const acc = String(r[iACC] || '').trim();
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
    const amt = Number(r[iA] || 0) * -1;

    out.push({
      date: dObj ? Utilities.formatDate(dObj, SpreadsheetApp.getActive().getSpreadsheetTimeZone() || 'Asia/Taipei', 'yyyy-MM-dd') : '',
      account: acc,
      amount: amt,
      item: item
    });
  }
  // 時間由近到遠
  out.sort((a, b) => String(b.date || '').localeCompare(String(a.date || '')));
  return out;
}

/** 明細清單（可過濾：account / dateFrom / dateTo；空值=不過濾） */
/** 明細清單（可過濾：account / dateFrom / dateTo；空值=不過濾） */
function getTransactions(filter) {
  filter = filter || {};
  const accPick = String(filter.account || '').trim();
  const fromStr = String(filter.dateFrom || '').trim();
  const toStr = String(filter.dateTo || '').trim();
  const from = fromStr ? new Date(fromStr) : null;
  const to = toStr ? new Date(toStr) : null;
  const limit = Number(filter.limit) || 0; // 0 = no limit

  const sh = SpreadsheetApp.getActive().getSheetByName(CF_SHEET);
  if (!sh) return [];
  const lastRow = sh.getLastRow();
  if (lastRow <= HEADER_ROW) return [];

  const n = lastRow - HEADER_ROW;
  const vals = sh.getRange(HEADER_ROW + 1, 1, n, sh.getLastColumn()).getValues();
  const M = getColMap_(sh);
  const iD = M.date - 1, iI = M.item - 1, iA = M.amount - 1, iACC = M.account - 1, iS = M.status - 1, iN = M.note - 1;
  const tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone() || 'Asia/Taipei';

  const out = [];
  for (let i = 0; i < vals.length; i++) {
    const r = vals[i];
    const _row = HEADER_ROW + 1 + i; // ← 真實列號（非常重要）

    const dRaw = r[iD];
    const dObj = dRaw instanceof Date ? dRaw : (dRaw ? new Date(dRaw) : null);
    if (from && dObj && dObj < from) continue;
    if (to && dObj && dObj > to) continue;

    const acc = String(r[iACC] || '').trim();
    if (accPick && acc !== accPick) continue;

    out.push({
      _row,
      date: dObj ? Utilities.formatDate(dObj, tz, 'yyyy-MM-dd') : '',
      item: String(r[iI] || '').trim(),
      amount: Number(r[iA] || 0),
      account: acc,
      status: String(r[iS] || '').trim().toLowerCase(),
      note: String(r[iN] || '').trim()
    });
  }
  // 近到遠
  out.sort((a, b) => String(b.date).localeCompare(String(a.date)));

  if (limit > 0 && out.length > limit) {
    return out.slice(0, limit);
  }
  return out;
}

/** Cashflow 分頁 URL（給「開啟分頁」用） */
function getCashflowUrl() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(CF_SHEET);
  if (!sh) return ss.getUrl();
  return ss.getUrl() + '#gid=' + sh.getSheetId();
}

/* ====== 共同工具 ====== */

/** 讀表頭 → 動態欄位定位；若找不到就回退到 COLS */
function getColMap_(sh) {
  const row = sh.getRange(HEADER_ROW, 1, 1, sh.getLastColumn()).getValues()[0] || [];
  const norm = s => String(s || '').trim().toLowerCase();
  const alias = {
    date: ['date', '日期', 'due date', 'time', 'day'],
    item: ['item', '明細', '項目', 'desc', 'description'],
    amount: ['amount', '金額', '金額(+-)'],
    account: ['account', '帳戶', '帳號', 'acct', 'bank'],
    status: ['status', '狀態', 'state', 'st'],
    note: ['note', '備註', '註記', 'remark', 'memo']
  };
  const idx = {};
  for (const k of Object.keys(alias)) {
    const names = alias[k];
    let found = 0;
    for (let c = 0; c < row.length; c++) {
      if (names.includes(norm(row[c]))) { found = c + 1; break; }
    }
    idx[k] = found || COLS[k];
  }
  return idx; // {date, item, amount, account, status, note}
}

/** 現有餘額（NOW）：Current + posted(-金額)；planned 不計
 *  修正邏輯：
 *  1. 找到該帳戶「最新」的一筆 item='Current' (Snapshot)
 *  2. 餘額 = Snapshot.amount + Σ(-posted.amount)
 *     其中 posted 必須是「日期 > Snapshot.date」的交易
 *     (若無 Snapshot，則 Snapshot.date = 0，即加總所有 posted)
 *  3. [NEW] 若 status='planned' 且 date <= Today，也視為 posted (Effective Posted)
 */
function calcNow_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(CF_SHEET);
  const out = {};
  if (!sh) return out;

  const lastRow = sh.getLastRow();
  if (lastRow <= HEADER_ROW) return out;

  const vals = sh.getRange(HEADER_ROW + 1, 1, lastRow - HEADER_ROW, sh.getLastColumn()).getValues();
  const M = getColMap_(sh);
  const iItem = M.item - 1, iAmt = M.amount - 1, iAcc = M.account - 1, iSta = M.status - 1, iDate = M.date - 1;

  // 1. Group by Account
  const byAcc = {};
  for (const r of vals) {
    const acc = String(r[iAcc] || '').trim();
    if (!acc) continue;
    if (!byAcc[acc]) byAcc[acc] = [];
    byAcc[acc].push(r);
  }

  // Helper date
  const tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone() || 'Asia/Taipei';
  const todayStr = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');

  // 2. Calculate per account
  for (const acc of Object.keys(byAcc)) {
    const rows = byAcc[acc];

    // 為了準確，我們先把 rows 轉成物件並 parse date
    const parsed = rows.map(r => {
      const dRaw = r[iDate];
      const dObj = dRaw instanceof Date ? dRaw : (dRaw ? new Date(dRaw) : null);
      const ts = dObj ? dObj.getTime() : 0;
      const dStr = dObj ? Utilities.formatDate(dObj, tz, 'yyyy-MM-dd') : '';
      return {
        r,
        ts,
        dStr,
        item: String(r[iItem] || '').trim(),
        sta: String(r[iSta] || '').trim().toLowerCase(),
        amt: Number(r[iAmt] || 0)
      };
    });

    // 找最新 Snapshot (Current)
    let snapDate = -1;
    let snapBal = 0;

    for (const p of parsed) {
      if (p.item === 'Current') {
        if (p.ts >= snapDate) {
          snapDate = p.ts;
          snapBal = p.amt;
        }
      }
    }

    // 計算餘額：Base + Sum(effective_posted where date > snapDate)
    let bal = snapDate === -1 ? 0 : snapBal;

    for (const p of parsed) {
      if (p.item === 'Current') continue;

      let isEffectivePosted = false;
      if (p.sta === 'posted') isEffectivePosted = true;
      else if (p.sta === 'planned') {
        // 若日期 <= 今天，視為已入帳
        if (p.dStr <= todayStr) isEffectivePosted = true;
      }

      if (isEffectivePosted) {
        // 只有日期「晚於」Snapshot 才納入
        // 這裡如果是「當日 Planned turned Posted」，只要它的日期 (00:00) > snapDate 就算
        // 若 snapDate 是今天的 Snapshot (time > 0)，則 planned (time=0) 可能被濾掉 -> 導致「當日 snapshot 後的新交易」漏算
        // 但通常 snapshot 是歷史定錨。若 snapshot 就是今天，那今天的新交易理應要 > snapshot time 才會算
        // 可是 planned 通常沒有時間 (00:00)，若 snapshot 是今天下午做的，那今天的 planned 就會變成 <= snapDate 而被忽略 (視為已包含在 snapshot)
        // 這符合邏輯：若 snapshot 比較新，那它應該已經包含了今天的 planned (如果當時已入帳)
        // 但如果 snapshot 是舊的 (e.g. 昨天)，那今天的 planned > snapDate，就會被納入 (扣除) -> 正確
        if (p.ts > snapDate) {
          bal += -p.amt;
        }
      }
    }
    out[acc] = bal;
  }
  return out;
}

/** [NEW] 回傳所有帳戶的餘額概況 { acc: { posted, planned, sum } } */
function getAccountBalances() {
  const sh = SpreadsheetApp.getActive().getSheetByName(CF_SHEET);
  if (!sh) return {};

  // 1. 取得 Posted/Now 餘額 (已有新邏輯)
  const postedMap = calcNow_();

  // 2. 取得 Planned (Future) 餘額
  const lastRow = sh.getLastRow();
  const plannedMap = {};

  if (lastRow > HEADER_ROW) {
    const vals = sh.getRange(HEADER_ROW + 1, 1, lastRow - HEADER_ROW, sh.getLastColumn()).getValues();
    const M = getColMap_(sh);
    const iItem = M.item - 1, iAmt = M.amount - 1, iAcc = M.account - 1, iSta = M.status - 1, iDate = M.date - 1;
    const tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone() || 'Asia/Taipei';
    const todayStr = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');

    for (const r of vals) {
      const acc = String(r[iAcc] || '').trim();
      if (!acc) continue;

      const item = String(r[iItem] || '').trim();
      if (item === 'Current') continue;

      const sta = String(r[iSta] || '').trim().toLowerCase();
      if (sta !== 'planned') continue;

      const dRaw = r[iDate];
      const dObj = dRaw instanceof Date ? dRaw : (dRaw ? new Date(dRaw) : null);
      const dStr = dObj ? Utilities.formatDate(dObj, tz, 'yyyy-MM-dd') : '';

      // 只算未來的 planned (日期 > 今天)
      if (dStr > todayStr) {
        const amt = Number(r[iAmt] || 0);
        // planned 顯示為 -金額 (支出為正，故用負號)
        if (!plannedMap[acc]) plannedMap[acc] = 0;
        plannedMap[acc] += -amt;
      }
    }
  }

  // 3. 取得 Using 總和 (維持原邏輯)
  const usingMap = {};
  if (lastRow > HEADER_ROW) {
    const vals = sh.getRange(HEADER_ROW + 1, 1, lastRow - HEADER_ROW, sh.getLastColumn()).getValues();
    const M = getColMap_(sh);
    const iAmt = M.amount - 1, iAcc = M.account - 1, iSta = M.status - 1, iItem = M.item - 1;

    for (const r of vals) {
      const acc = String(r[iAcc] || '').trim();
      if (!acc) continue;
      if (String(r[iItem] || '').toLowerCase() === 'current') continue;
      const sta = String(r[iSta] || '').trim().toLowerCase();
      if (sta === 'using') {
        const amt = Number(r[iAmt] || 0);
        if (!usingMap[acc]) usingMap[acc] = 0;
        usingMap[acc] += -amt; // using 視同支出/投資，用負號
      }
    }
  }

  // 4. 合併結果
  const out = [];
  // 收集所有帳戶
  const allAcc = new Set([...Object.keys(postedMap), ...Object.keys(plannedMap), ...Object.keys(usingMap)]);

  allAcc.forEach(acc => {
    const p = postedMap[acc] || 0;
    const pl = plannedMap[acc] || 0;
    const u = usingMap[acc] || 0;
    out.push({
      account: acc,
      posted: p,
      planned: pl,
      using: u,
      sum: p + pl // 一般定義 sum = posted + planned (using 只有貸款帳戶特殊處理，前端會再算)
    });
  });

  return out.sort((a, b) => a.account.localeCompare(b.account, 'zh-Hant'));
}

/** 回傳指定帳戶的 Balance(NOW)；name 空字串→回全部加總 */
function getBalanceForAccount_(name) {
  const all = calcNow_();
  if (!name) return Object.values(all).reduce((a, b) => a + Number(b || 0), 0);
  return all[name] ?? 0;
}

/** 寫入一筆（status=auto → 依欄位距離置入 R1C1 公式），回傳該帳戶 NOW */
function submitCashflow(rec) {
  const sh = SpreadsheetApp.getActive().getSheetByName(CF_SHEET);
  if (!sh) throw new Error('找不到資料分頁：' + CF_SHEET);

  const row = Math.max(sh.getLastRow() + 1, HEADER_ROW + 1);
  const M = getColMap_(sh);

  const d = rec.date ? new Date(rec.date) : '';
  const item = String(rec.item || '').trim();
  const amt = rec.amount === '' || rec.amount == null ? '' : Number(rec.amount);
  const acc = String(rec.account || '').trim();
  let st = String(rec.status || '').trim(); // 'auto' | 'posted' | 'planned'
  const note = String(rec.note || '').trim();

  if (M.date) sh.getRange(row, M.date).setValue(d);
  if (M.item) sh.getRange(row, M.item).setValue(item);
  if (M.amount) sh.getRange(row, M.amount).setValue(amt);
  if (M.account) sh.getRange(row, M.account).setValue(acc);
  if (M.note) sh.getRange(row, M.note).setValue(note);

  if (M.status) {
    const stCell = sh.getRange(row, M.status);
    if (/^using$/i.test(st)) {
      // 保留 using 狀態，不套自動公式
      stCell.clearFormat().setValue('using');
    } else {
      const delta = (M.date || 0) - (M.status || 0);
      stCell.setFormulaR1C1(`=IF(INT(RC[${delta}])<=TODAY(),"posted","planned")`);
    }
  }

  applyFormatAndValidation_?.(row);
  SpreadsheetApp.flush();       // ★ 同理
  const balanceNow = getBalanceForAccount_(acc);
  return { row, balanceNow };
}

/** 一鍵修復：把缺少驗證/格式的列補齊 */
function fixMissingValidationAll_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(CF.sheetName);
  if (!sh) return;
  const lastRow = sh.getLastRow();
  if (lastRow <= CF.headerRow) return;

  const width = CF.lastCol - CF.firstCol + 1;
  const rng = sh.getRange(CF.headerRow + 1, CF.firstCol, lastRow - CF.headerRow, width);
  const dv = rng.getDataValidations(); // 2D 陣列
  const startRow = CF.headerRow + 1;

  for (let r = 0; r < dv.length; r++) {
    if (dv[r].some(cell => cell == null)) {
      applyFormatAndValidation_(startRow + r);
    }
  }
  SpreadsheetApp.getUi().alert('已檢查並補齊缺失的格式與下拉驗證。');
}

/* ===== 你前端會叫用的：CTBC investment（取 ✅ 分頁 H,K,L,N,T,U,V,AF,AH） ===== */
function getCTBCInvestments() {
  const SHEET = '✅';
  // A=1, ..., AF=32, AG=33, AH=34
  const INDEX = { G: 7, H: 8, K: 11, L: 12, M: 13, N: 14, P: 16, Q: 17, R: 18, T: 20, U: 21, V: 22, AF: 32, AH: 34, AI: 35 };
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
  function normCur_(v) {
    const s = String(v || '').toUpperCase();
    if (s.includes('USD')) return 'USD';
    if (s.includes('TWD') || s.includes('NTD') || /(NT|台幣|新台幣)/i.test(s)) return 'TWD';
    return '';
  }

  // 輸出：每列是一個物件，包含列號 ROW（2 起算）
  const out = values.map((r, i) => ({
    ROW: i + 2, // 真實試算表列號
    G: num_(r[INDEX.G - 1]),
    H: toStr(r[INDEX.H - 1]),
    K: toStr(r[INDEX.K - 1]),
    L: num_(r[INDEX.L - 1]),
    M: num_(r[INDEX.M - 1]),
    N: num_(r[INDEX.N - 1]),
    R: toStr(r[INDEX.R - 1]),
    QUOTE: num_(r[INDEX.R - 1]),
    T: num_(r[INDEX.T - 1]),
    U: num_(r[INDEX.U - 1]),
    V: toStr(r[INDEX.V - 1]),
    CUR: normCur_(r[INDEX.V - 1]),
    AF: num_(r[INDEX.AF - 1]),
    AH: toStr(r[INDEX.AH - 1]),
    AI: r[INDEX.AI - 1],
    P: toStr(r[INDEX.P - 1]),
    Q: toStr(r[INDEX.Q - 1])
  }));

  function num_(x) { x = Number(x); return isFinite(x) ? x : 0; }

  return out;
}

/** 更新 ✅ 分頁的值 */
function setInvestValue(rowId, colKey, val) {
  const SHEET = '✅';
  const INDEX = { G: 7, H: 8, K: 11, L: 12, M: 13, N: 14, P: 16, Q: 17, R: 18, T: 20, U: 21, V: 22, AF: 32, AH: 34, AI: 35 };
  const colIndex = INDEX[colKey];
  if (!colIndex) throw new Error('Invalid column key: ' + colKey);

  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET);
  if (!sh) throw new Error('Sheet not found: ' + SHEET);

  sh.getRange(rowId, colIndex).setValue(val);
  return true;
}

/** 在 ✅ 分頁新增一筆投資項目 */
function addInvestEntry(category, name) {
  const SHEET = '✅';
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET);
  if (!sh) throw new Error('Sheet not found');

  // 假設架構：G(7)=類別, K(11)=項目名稱, AI(35)=狀態(TRUE)
  // 我們先建立一個 36 欄的陣列
  const row = new Array(37).fill(''); // 到 AK (37)
  row[6] = category; // G
  row[10] = name;    // K
  row[34] = false;   // AI (狀態預設為持有中 -> FALSE)

  sh.appendRow(row);
  return { success: true };
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
function getCTBCAggregates(kind) {
  const rowsAll = getCTBCInvestments() || [];

  // 允許 'fund' 或 '基金' 觸發基金篩選（比對 H/K）；允許 'stock' 或 '股票' 觸發股票篩選
  const k = String(kind || '').toLowerCase();
  const isFundReq = k.includes('fund') || k.includes('基金');
  const isStockReq = k.includes('stock') || k.includes('股票');
  const isFund = (r) => /基金|FUND/i.test(String(r.H || '')) || /基金|FUND/i.test(String(r.K || ''));
  const rows = isFundReq ? rowsAll.filter(isFund)
    : isStockReq ? rowsAll.filter(r => !isFund(r))
      : rowsAll;

  // 幣別挑選：優先 CUR，否則以 V 判斷
  const pick = (code) => rows.filter(r => {
    const cur = (r.CUR || '').toUpperCase();
    if (cur) return cur === code;
    const v = String(r.V || '').toUpperCase();
    if (code === 'USD') return v.includes('USD');
    if (code === 'TWD') return v.includes('TWD') || v.includes('NTD') || /(NT|台幣|新台幣)/i.test(v);
    return false;
  });

  // 與前端 ctbc_js 一致的 normalizeU：|U|<10 視為小數，乘 100；否則視為百分比
  function normalizeU(u) {
    var n = Number(u || 0);
    if (!isFinite(n)) n = 0;
    // 修正：閾值改為 10 (1000%)，避免 >100% (1.x) 被誤判為小數
    return (Math.abs(n) < 10) ? (n * 100) : n;
  }

  // 匯總：N/T/AF/L，同時計算「以 N 加權」的 U 加權平均（roiU）
  function agg(list) {
    var N = 0, T = 0, AF = 0, Lsum = 0, w = 0, wU = 0;
    for (var i = 0; i < list.length; i++) {
      var r = list[i];
      var n = Number(r.N || 0);
      var t = Number(r.T || 0);
      var af = Number(r.AF || 0);
      var l = Number(r.L || 0);
      var u = normalizeU(r.U);
      if (!isFinite(n)) n = 0;
      if (!isFinite(t)) t = 0;
      if (!isFinite(af)) af = 0;
      if (!isFinite(l)) l = 0;
      N += n; T += t; AF += af; Lsum += l;
      if (n) { w += n; wU += (u * n); }
    }
    var roiU = w ? (wU / w) : null; // 與前端一致：N 加權的 U 平均（百分比）
    return { N: N, T: T, AF: AF, Lsum: Lsum, roiU: roiU };
  }

  var all = agg(rows);
  var twd = agg(pick('TWD'));
  var usd = agg(pick('USD'));

  // AllNoInt = 現值合計（ΣN），依照當前篩選後的集合計算
  var allNoInt = all.N;

  return {
    ctbcAllN: all.Lsum,
    ctbcAllT: all.T,
    ctbcAllNoInt: allNoInt,
    ctbcAllAF: all.AF,

    ctbcTWD_N: twd.N,
    ctbcTWD_T: twd.T,
    ctbcTWD_AF: twd.AF,

    ctbcUSD_N: usd.N,
    ctbcUSD_T: usd.T,
    ctbcUSD_AF: usd.AF,

    // ROI 與前端相同：以 U 的 N 加權平均（百分比值）
    ctbcAllROI: (all.roiU == null ? 0 : all.roiU),
    ctbcTWD_ROI: (twd.roiU == null ? 0 : twd.roiU),
    ctbcUSD_ROI: (usd.roiU == null ? 0 : usd.roiU)
  };
}

/**
 * 一次回傳「基金 / 股票」的現值與 ROI（與前端一致：U 以 N 加權；U 在 -1..1 視為小數×100）
 * 回傳：{ fund:{ value, roi }, stock:{ value, roi } }
 */
function getCTBCFundStock() {
  const rowsAll = getCTBCInvestments() || [];
  const isFund = (r) => /基金|FUND/i.test(String(r.H || '')) || /基金|FUND/i.test(String(r.K || ''));
  const fundRows = rowsAll.filter(isFund);
  const stockRows = rowsAll.filter(r => !isFund(r));

  function normalizeU(u) {
    var n = Number(u || 0); if (!isFinite(n)) n = 0;
    // 修正：閾值改為 10 (1000%)
    return (Math.abs(n) < 10) ? (n * 100) : n;
  }
  function agg(list) {
    var N = 0, w = 0, wU = 0;
    for (var i = 0; i < list.length; i++) {
      var n = Number(list[i].N || 0); if (!isFinite(n)) n = 0;
      var u = normalizeU(list[i].U);
      N += n; if (n) { w += n; wU += (u * n); }
    }
    return { value: N, roi: w ? (wU / w) : 0 };
  }

  return { fund: agg(fundRows), stock: agg(stockRows) };
}

/** 讀取某列 AF 儲存格的公式與顯示值（優先回公式） */
function getCTBCInterestCell(rowNumber) {
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
function setCTBCInterestUSD(rowNumber, input) {
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
    rngAF.setValue(Number(input) || 0);
  }

  // 注意：用程式寫值不會觸發你的 onEdit，所以這裡直接寫 AH（見第 4 點）
  const tz = Session.getScriptTimeZone() || 'Asia/Taipei';
  const today = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  sh.getRange(rowNumber, COL_AH).setValue(today);

  SpreadsheetApp.flush();
  return { row: rowNumber };
}

/** 讀取某列 QUOTE(R) 的公式/值 */
function getCTBCQuoteCell(rowNumber) {
  const SHEET = '✅';
  const COL_R = 18; // 目前報價
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET);
  if (!sh) throw new Error('找不到工作表：' + SHEET);
  if (rowNumber < 2) throw new Error('列號不合法');
  const rng = sh.getRange(rowNumber, COL_R);
  return { formula: rng.getFormula() || '', value: rng.getDisplayValue() || '' };
}

/** 寫回「目前報價 (USD)」（R=18），並更新 AH(34) */
function setCTBCQuoteUSD(rowNumber, input) {
  const SHEET = '✅';
  const COL_R = 18; // 目前報價
  const COL_AH = 34; // 更新日
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET);
  if (!sh) throw new Error('找不到工作表：' + SHEET);
  if (rowNumber < 2) throw new Error('列號不合法');

  const rngR = sh.getRange(rowNumber, COL_R);
  const s = (typeof input === 'string') ? input.trim() : '';
  if (s && s[0] === '=') rngR.setFormula(s); else rngR.setValue(Number(input) || 0);

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
  if (lastRow <= HEADER_ROW) return [["Account", "NOW", "Planned", "Final"]];

  const M = getColMap_(sh);
  const vals = sh.getRange(HEADER_ROW + 1, 1, lastRow - HEADER_ROW, sh.getLastColumn()).getValues();
  const iI = M.item - 1, iA = M.amount - 1, iACC = M.account - 1, iS = M.status - 1;

  const agg = {}; // acc -> {now, planned}
  for (const r of vals) {
    const acc = String(r[iACC] || '').trim(); if (!acc) continue;
    const item = String(r[iI] || '').trim();
    const sta = String(r[iS] || '').trim().toLowerCase();
    const amt = Number(r[iA] || 0);
    if (!agg[acc]) agg[acc] = { now: 0, planned: 0 };
    if (item === 'Current') agg[acc].now += amt;
    else if (sta === 'posted') agg[acc].now += -amt;
    else if (sta === 'planned') agg[acc].planned += -amt;
  }

  const rows = Object.keys(agg).sort().map(acc => {
    const now = agg[acc].now || 0, pln = agg[acc].planned || 0;
    return [acc, now, pln, now + pln];
  });

  const tNow = rows.reduce((s, r) => s + Number(r[1] || 0), 0);
  const tPln = rows.reduce((s, r) => s + Number(r[2] || 0), 0);
  const tFin = rows.reduce((s, r) => s + Number(r[3] || 0), 0);

  return [
    ["Account", "NOW", "Planned", "Final"],
    ...rows,
    ["TOTAL", tNow, tPln, tFin]
  ];
}

/** 給前端 Dashboard 用（物件陣列） */
function getSummaryCF() {
  const t = CF_SUMMARY();                     // 含表頭 + TOTAL
  return (t || []).slice(1).map(r => ({           // 含 TOTAL 一起回傳，前端自行決定是否顯示
    account: String(r[0] ?? ''),
    now: Number(r[1] ?? 0),
    planned: Number(r[2] ?? 0),
    final: Number(r[3] ?? 0)
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
  var limit = Math.max(0, Number(opt.limit || 200));
  var order = String(opt.order || 'desc').toLowerCase(); // 預設最新→最舊

  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(RECORD_SHEET_NAME);
  if (!sh) return { headers: [], rows: [], from: 0, to: 0, total: 0, allHeaders: [], lastUsedRow: 2 };

  var lastRow = sh.getLastRow();
  var lastCol = sh.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return { headers: [], rows: [], from: 0, to: 0, total: 0, allHeaders: [], lastUsedRow: 2 };

  // 表頭在第 2 列
  var headers = sh.getRange(2, 1, 1, lastCol).getDisplayValues()[0] || [];
  var dataStart = 3;

  if (lastRow < dataStart) {
    return { headers, rows: [], from: 0, to: 0, total: 0, allHeaders: headers, lastUsedRow: dataStart - 1 };
  }

  // 一次性把第3列到最後列都抓出來（顯示值），後面用陣列處理，避免逐列 getRange 造成「無此範圍」
  var block = sh.getRange(dataStart, 1, lastRow - dataStart + 1, lastCol).getDisplayValues();

  // 修剪尾端「工作列/空白列」→ 規則改為：整列顯示值皆空白才刪
  function isBlankRow(arr) {
    return arr.every(function (v) { return String(v || '').trim() === ''; });
  }

  var endIdx = block.length - 1;
  while (endIdx >= 0) {
    var rowArr = block[endIdx];
    if (!isBlankRow(rowArr)) break;  // 只要有任何顯示值就保留（即便全是公式）
    endIdx--;
  }
  if (endIdx < 0) {
    return { headers, rows: [], from: 0, to: 0, total: 0, allHeaders: headers, lastUsedRow: dataStart - 1 };
  }

  var dataStart = 3; // 資料從第 3 列開始（如果你原本就有這個常數，保持一致即可）

  var lastUsedRow = dataStart + endIdx; // 真實試算表最後一列（有顯示值）

  // 有效資料塊
  var data = block.slice(0, endIdx + 1);
  var total = data.length;

  // 順序：desc = 由新到舊（視為資料底部是最新 → 反轉）；asc = 由舊到新（保持）
  var ordered = (order === 'desc') ? data.slice().reverse() : data;

  // 切頁
  var from = Math.min(offset, total);
  var to = Math.min(offset + limit, total);
  var rows = ordered.slice(from, to);

  return { headers, rows, from, to, total, allHeaders: headers, lastUsedRow };
}

/**
 * 依虛擬列索引（最新在上，0-based）計算所在分頁並回傳該分頁資料
 * 用途：解法「row847 查無資料」—先用此 API 算 offset，再抓對的 page
 * @param {Object} opt
 * @param {number} opt.row    目標虛擬列索引（0 表「最新一列」）
 * @param {number} [opt.limit=200] 每頁筆數
 * @param {string} [opt.order='desc'] 'desc'=最新在上、'asc'=最舊在上（需與前端一致）
 * @return {{headers:string[], rows:string[][], from:number, to:number, total:number, allHeaders:string[], lastUsedRow:number, selectedIndex:number, selectedRow:string[]|null}}
 */
function getDatabasePageForRow(opt) {
  opt = opt || {};
  var want = Math.max(0, Number(opt.row || 0));
  var limit = Math.max(1, Number(opt.limit || 200));
  var order = String(opt.order || 'desc').toLowerCase();

  // 與 getDatabaseRows 同步的讀取邏輯（避免「尾端骨架列」干擾）
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(RECORD_SHEET_NAME);
  if (!sh) return { headers: [], rows: [], from: 0, to: 0, total: 0, allHeaders: [], lastUsedRow: 2, selectedIndex: -1, selectedRow: null };

  var lastRow = sh.getLastRow();
  var lastCol = sh.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return { headers: [], rows: [], from: 0, to: 0, total: 0, allHeaders: [], lastUsedRow: 2, selectedIndex: -1, selectedRow: null };

  var headers = sh.getRange(2, 1, 1, lastCol).getDisplayValues()[0] || [];
  var dataStart = 3;
  if (lastRow < dataStart) return { headers, rows: [], from: 0, to: 0, total: 0, allHeaders: headers, lastUsedRow: dataStart - 1, selectedIndex: -1, selectedRow: null };

  var block = sh.getRange(dataStart, 1, lastRow - dataStart + 1, lastCol).getDisplayValues();
  function isBlankRow(arr) { return arr.every(function (v) { return String(v || '').trim() === ''; }); }
  var endIdx = block.length - 1;
  while (endIdx >= 0) { if (!isBlankRow(block[endIdx])) break; endIdx--; }
  if (endIdx < 0) return { headers, rows: [], from: 0, to: 0, total: 0, allHeaders: headers, lastUsedRow: dataStart - 1, selectedIndex: -1, selectedRow: null };

  var lastUsedRow = dataStart + endIdx;  // 真實最後一列（有顯示值）
  var data = block.slice(0, endIdx + 1);
  var total = data.length;               // 虛擬清單總筆數

  // 索引解讀：desc = 最新在上（反轉），asc = 最舊在上（保持）
  var ordered = (order === 'desc') ? data.slice().reverse() : data;

  // 重要：row 需夾在 [0, total)
  var clamped = Math.max(0, Math.min(want, Math.max(0, total - 1)));

  // 計算 page 邊界（offset）與該頁 rows
  var offset = Math.floor(clamped / limit) * limit; // page 的起點（含）
  var from = Math.min(offset, total);
  var to = Math.min(offset + limit, total);
  var rows = ordered.slice(from, to);

  // 該頁中的相對 index 與那一列的資料
  var selectedIndex = (clamped >= from && clamped < to) ? (clamped - from) : -1;
  var selectedRow = selectedIndex >= 0 ? rows[selectedIndex] : null;

  return { headers, rows, from, to, total, allHeaders: headers, lastUsedRow, selectedIndex, selectedRow };
}

/**
 * 依虛擬列索引直接取該列原始值（不分頁）
 * @param {number} rowIndexFromTop 0-based（與前端虛擬清單一致）
 * @param {string} [order='desc']  'desc'=最新在上
 * @return {{headers:string[], row:any[]|null, total:number, lastUsedRow:number}}
 */
function getDatabaseRowByIndex(rowIndexFromTop, order) {
  var page = getDatabasePageForRow({ row: Number(rowIndexFromTop || 0), limit: 1, order: order || 'desc' });
  return { headers: page.headers || [], row: page.selectedRow || null, total: page.total || 0, lastUsedRow: page.lastUsedRow || 0 };
}

/** 左側最新資料卡：用欄位字母定位 + 同時取值(getValues)與顯示(getDisplayValues) */
function getRecordLatest() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(RECORD_SHEET_NAME);
  if (!sh) return null;

  var lastRow = sh.getLastRow();
  var lastCol = sh.getLastColumn();
  if (lastRow < 3) return null;

  // 欄字母→1-based 欄號（支援 AA、BB、CV 等）
  function colLetter(s) {
    s = String(s || '').trim().toUpperCase();
    var n = 0;
    for (var i = 0; i < s.length; i++) {
      var code = s.charCodeAt(i);
      if (code >= 65 && code <= 90) n = n * 26 + (code - 64);
    }
    return n;
  }

  // 直接指定欄位：A/C/D/E/F/G/BA/CU   ← AQ 移除後向左位移一欄
  var cDate = colLetter('A');
  var cWhole = colLetter('C');
  var cWithDad = colLetter('D');
  var cAvail = colLetter('E');
  var cPnL = colLetter('F');  // ★ 當日損益
  var cCash = colLetter('G');
  var cDebt = colLetter('BA'); // 原 BB 因移除 AQ 改為 BA
  var cUSD = colLetter('CU'); // 原 CV 因移除 AQ 改為 CU

  // 判斷是否為空列：關鍵欄位全都空 → 才算空
  function isEmptyRow(r) {
    var disp = sh.getRange(r, 1, 1, lastCol).getDisplayValues()[0];
    var keys = [cDate, cWhole, cWithDad, cAvail, cPnL, cCash, cDebt, cUSD];
    return keys.every(function (c) {
      return String(disp[c - 1] || '').trim() === '';
    });
  }

  var rowIdx = lastRow;
  while (rowIdx >= 3 && isEmptyRow(rowIdx)) rowIdx--;
  if (rowIdx < 3) return null;

  // 同時抓「原始值」與「顯示字串」
  var rVal = sh.getRange(rowIdx, 1, 1, lastCol).getValues()[0];
  var rDisp = sh.getRange(rowIdx, 1, 1, lastCol).getDisplayValues()[0];

  // 將顯示字串轉數字（支援 1,234 / 1 234 / 1，234 / (1,234) / 12.3%）
  function toNum(vVal, vDisp) {
    // 原始值若為 number，直接用
    if (typeof vVal === 'number' && isFinite(vVal)) return vVal;

    var s = String(vDisp || '').trim();
    if (!s) return 0;

    // 括號負數
    var neg = false;
    if (/^\(.*\)$/.test(s)) { neg = true; s = s.slice(1, -1); }

    // 移除各種空白（含 NBSP）與各地區千分位逗號/全形逗號
    s = s.replace(/[\u00A0\s,，]/g, '');

    // 百分比
    var isPct = /%$/.test(s);
    if (isPct) s = s.replace(/%$/, '');

    var n = Number(s);
    if (!isFinite(n)) return 0;
    if (isPct) n = n / 100;
    if (neg) n = -n;
    return n;
  }

  // 有沒有任何關鍵值
  function hasVal(v) { return String(v || '').trim() !== ''; }
  var keyHasValue =
    hasVal(rDisp[cWhole - 1]) || hasVal(rDisp[cWithDad - 1]) || hasVal(rDisp[cAvail - 1]) ||
    hasVal(rDisp[cCash - 1]) || hasVal(rDisp[cDebt - 1]) || hasVal(rDisp[cUSD - 1]) ||
    hasVal(rDisp[cDate - 1]);
  if (!keyHasValue) return null;

  // 日期輸出：優先 Date，其次顯示字串
  var dVal = rVal[cDate - 1];
  var dDisp = rDisp[cDate - 1];
  var tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone() || 'Asia/Taipei';
  var outDate = (dVal instanceof Date)
    ? Utilities.formatDate(dVal, tz, 'yyyy-MM-dd')
    : String(dDisp || '');

  return {
    date: outDate,
    C: toNum(rVal[cWhole - 1], rDisp[cWhole - 1]),
    D: toNum(rVal[cWithDad - 1], rDisp[cWithDad - 1]),
    E: toNum(rVal[cAvail - 1], rDisp[cAvail - 1]),
    F: toNum(rVal[cPnL - 1], rDisp[cPnL - 1]),
    G: toNum(rVal[cCash - 1], rDisp[cCash - 1]),
    BA: toNum(rVal[cDebt - 1], rDisp[cDebt - 1]),   // 債務（內部計算用）
    CU: toNum(rVal[cUSD - 1], rDisp[cUSD - 1])    // 持有 USD
  };
}

/** 一次回傳 Dashboard 基礎資料（Database 最新 + Cashflow TOTAL + CTBC 總覽） */
function getDashboardData() {
  // --- Database 最新（已有 getRecordLatest）
  const latest = getRecordLatest() || { date: '', C: 0, D: 0, E: 0, G: 0, BA: 0, CU: 0 };
  const kpi_total = Number(latest.C || 0);
  const kpi_debt = Number(latest.BA || 0); // 內部用於淨資產計算
  const kpi_cash = Number(latest.G || 0);
  const kpi_avail = Number(latest.E || 0);
  const kpi_usd = Number(latest.CU || 0);
  const kpi_net = Number(latest.D || 0); // 以 D 欄 Whole Assets 作為「淨資產」顯示

  // --- Cashflow TOTAL（已有 getSummaryCF：含 TOTAL）
  const cfRows = getSummaryCF() || [];
  const totalRow = cfRows.find(r => String(r.account || '').toUpperCase() === 'TOTAL') || { now: 0, planned: 0, final: 0 };
  const cfTotal = { now: Number(totalRow.now || 0), planned: Number(totalRow.planned || 0), final: Number(totalRow.final || 0) };

  // 取近期 5 筆 planned（給 Dashboard 卡片列）
  const plannedList = (getTxByKind('planned') || []).slice(0, 5);

  // --- CTBC：改為基金-only 指標 + 市場（STK）未賣出總市值 ---
  const ctAll = getCTBCInvestments() || [];
  const isFund = (r) => /基金|FUND/i.test(String(r.H || '')) || /基金|FUND/i.test(String(r.K || ''));

  // 基金-only 匯總
  const fundAgg = ctAll.reduce((a, r) => {
    if (!isFund(r)) return a;
    const v = Number(r.N || 0);    // 現值 (USD)
    const pnl = Number(r.T || 0);  // 含息損益 (USD)
    const intv = Number(r.AF || 0); // 已領利息 (USD)
    a.value += v; a.pnl += pnl; a.interest += intv;
    // U 正規化（與前端一致：|U|≤1 視為小數，×100）
    let u = Number(r.U || 0); if (!isFinite(u)) u = 0; if (Math.abs(u) <= 1) u *= 100;
    if (v) { a.w += v; a.wU += (u * v); }
    return a;
  }, { value: 0, pnl: 0, interest: 0, w: 0, wU: 0 });
  const fundROI = fundAgg.w ? (fundAgg.wU / fundAgg.w) : 0;

  // 市場（STK）未賣出總市值：讀 STK，A=false 的 N 欄加總
  let marketsMV = 0;
  try {
    const stkRows = getSTKRows() || [];
    marketsMV = stkRows.reduce((s, r) => s + ((r && r.A === true) ? 0 : Number(r.N || 0)), 0);
  } catch (e) { marketsMV = 0; }

  const headerTotalMV = (Number(fundAgg.value || 0) + Number(marketsMV || 0));

  return {
    kpis: {
      date: latest.date || '',
      total: kpi_total,
      net: kpi_net,
      cash: kpi_cash,
      avail: kpi_avail,
      usd: kpi_usd
    },
    cashflow: {
      total: cfTotal,
      plannedTop: plannedList
    },
    // 儀表板用：ctbc 區塊回「基金-only」與「市場總市值」
    ctbc: {
      fund: { value: fundAgg.value, pnl: fundAgg.pnl, interest: fundAgg.interest, roi: fundROI },
      markets: { totalMV: marketsMV },
      // 兼容舊鍵：allUSD 的 pnl/interest 改回傳基金-only，value 仍回基金現值（現值總合顯示交給 header pill）
      allUSD: { value: fundAgg.value, pnl: fundAgg.pnl, ret: fundROI, interest: fundAgg.interest },
      headerTotalMV: headerTotalMV
    },
    database: {
      latest // {date,C,D,E,G,BA,CU}
    }
  };
}

/** ====== 基本設定：Dash 分頁、欄位位置 ====== */
const DASH_SHEET_NAME = 'Dash';
const START_ROW = 3;      // A3 起算
const NAME_COL = 1;      // A: 帳戶名
const ORIG_COL = 2;      // B: 原幣（這裡是我們要寫回的目標欄）
const TWD_COL = 3;      // C: 換算新台幣（通常由公式帶出，不直接寫）

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
function saveLinkNote(row, note) {
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

/**
 * 取得 USD→TWD 匯率（優先 GOOGLEFINANCE；10 分鐘快取；失敗回 32）
 * 前端可直接呼叫 google.script.run.getUSDTWD()
 */
function getUSDTWD() {
  try {
    var cache = CacheService.getDocumentCache();
    var hit = cache && cache.get('fx_USD_TWD');
    if (hit != null && hit !== '') return Number(hit);
  } catch (e) { /* ignore cache errors */ }

  var v = NaN;
  try {
    var r = evalBySheetsEngine('=GOOGLEFINANCE("CURRENCY:USDTWD")');
    if (r && r.ok && isFinite(r.value) && Number(r.value) > 0) v = Number(r.value);
  } catch (e) { /* fallthrough */ }

  // 一些情況 GOOGLEFINANCE 會回表格；INDEX 取最新價格
  if (!isFinite(v) || v <= 0) {
    try {
      var r2 = evalBySheetsEngine('=INDEX(GOOGLEFINANCE("CURRENCY:USDTWD"),2,2)');
      if (r2 && r2.ok && isFinite(r2.value) && Number(r2.value) > 0) v = Number(r2.value);
    } catch (e) { /* fallthrough */ }
  }

  // 兜底（離線或服務異常時）
  if (!isFinite(v) || v <= 0) {
    v = 32; // safe default
    console.warn('[getUSDTWD] Fallback to 32 used. Google Finance failed.');
  }

  try { if (cache) cache.put('fx_USD_TWD', String(v), 600); } catch (e) { }
  return v;
}

/**
 * 通用：取得 base→quote 匯率，例如 getFxRate('USD','TWD')
 * 失敗回 NaN；呼叫者自行處理預設值
 */
function getFxRate(base, quote) {
  base = String(base || '').trim().toUpperCase();
  quote = String(quote || '').trim().toUpperCase();
  if (!base || !quote) return NaN;

  var key = 'fx_' + base + '_' + quote;
  try {
    var cache = CacheService.getDocumentCache();
    var hit = cache && cache.get(key);
    if (hit != null && hit !== '') return Number(hit);
  } catch (e) { /* ignore cache errors */ }

  var v = NaN;
  var expr = '=GOOGLEFINANCE("CURRENCY:' + base + quote + '")';
  try {
    var r = evalBySheetsEngine(expr);
    if (r && r.ok && isFinite(r.value) && Number(r.value) > 0) v = Number(r.value);
  } catch (e) { /* fallthrough */ }
  if (!isFinite(v) || v <= 0) {
    try {
      var r2 = evalBySheetsEngine('=INDEX(GOOGLEFINANCE("CURRENCY:' + base + quote + '"),2,2)');
      if (r2 && r2.ok && isFinite(r2.value) && Number(r2.value) > 0) v = Number(r2.value);
    } catch (e) { /* fallthrough */ }
  }
  if (!isFinite(v) || v <= 0) return NaN;

  try {
    var cache2 = CacheService.getDocumentCache();
    if (cache2) cache2.put(key, String(v), 600);
  } catch (e) { }
  return v;
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

function updateTransaction(row, rec) {
  const sh = SpreadsheetApp.getActive().getSheetByName(CF_SHEET);
  if (!sh) throw new Error('找不到工作表：' + CF_SHEET);

  if (!row || row < HEADER_ROW + 1) throw new Error('row 不合法（需為資料列，通常 ≥ ' + (HEADER_ROW + 1) + '）');

  const M = getColMap_(sh);
  const map = {
    date: M.date,
    item: M.item,
    amount: M.amount,
    account: M.account,
    status: M.status,
    note: M.note
  };

  // 轉型
  const d = rec.date ? new Date(rec.date) : '';
  const item = String(rec.item || '').trim();
  const amt = rec.amount === '' || rec.amount == null ? '' : Number(rec.amount);
  const acc = String(rec.account || '').trim();
  let st = String(rec.status || '').trim();
  const note = String(rec.note || '').trim();

  // 寫入
  if (map.date) sh.getRange(row, map.date).setValue(d);
  if (map.item) sh.getRange(row, map.item).setValue(item);
  if (map.amount) sh.getRange(row, map.amount).setValue(amt);
  if (map.account) sh.getRange(row, map.account).setValue(acc);
  if (map.note) sh.getRange(row, map.note).setValue(note);

  if (map.status) {
    const stCell = sh.getRange(row, map.status);
    if (/^using$/i.test(st)) {
      // 保留 using 狀態，不套自動公式
      stCell.clearFormat().setValue('using');
    } else {
      const delta = (map.date || 0) - (map.status || 0);
      stCell.setFormulaR1C1(`=IF(INT(RC[${delta}])<=TODAY(),"posted","planned")`);
    }
  }
  SpreadsheetApp.flush();       // ★ 讓前端下一次讀到最新值
  return { ok: true, row };
}

/** 刪除指定列（Cashflow!A:F 的整列）
 * @param {number} row 真實試算表列號（含表頭，通常 >= 2）
 * @return {{ok:boolean,row:number}}
 */
function deleteTransaction(row) {
  const sh = SpreadsheetApp.getActive().getSheetByName(CF_SHEET);
  if (!sh) throw new Error('找不到工作表：' + CF_SHEET);
  const last = sh.getLastRow();
  const nRow = Number(row);
  if (!nRow || nRow < HEADER_ROW + 1 || nRow > last) {
    throw new Error('row 不合法（需為資料列，通常 ≥ ' + (HEADER_ROW + 1) + ' 且 ≤ ' + last + '）：' + row);
  }
  sh.deleteRow(nRow);
  SpreadsheetApp.flush();
  return { ok: true, row: nRow };
}


/** 保單成長：讀 A, I, J, L, N, S 欄（第 11 列為表頭） */
function getPolicyGrowth() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('保單成長');
  if (!sh) throw new Error('找不到「保單成長」分頁');

  const last = sh.getLastRow();
  if (last < 12) return { headers: ['年/月', '年齡', '紅利模型', '累計提領', 'Total P/L', '累計投入'], rows: [], years: [] };

  // 指定欄位（從第 12 列開始是資料）
  const rA = sh.getRange(12, 1, last - 11, 1).getValues();   // 年份（A）
  const rI = sh.getRange(12, 9, last - 11, 1).getValues();   // 累計投入（I）
  const rJ = sh.getRange(12, 10, last - 11, 1).getValues();   // 累計提領（J）
  const rL = sh.getRange(12, 12, last - 11, 1).getValues();   // Total P/L（L）
  const rN = sh.getRange(12, 14, last - 11, 1).getValues();   // 紅利模型（N）
  const rS = sh.getRange(12, 19, last - 11, 1).getValues();   // 年齡（S）

  // 轉成 rows
  const rows = [];
  const years = [];
  for (let i = 0; i < rA.length; i++) {
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
    headers: ['年/月', '年齡', '紅利模型', '累計提領', 'Total P/L', '累計投入'],
    rows,
    years
  };
}

/** A 欄可能是日期或數字，統一回傳 yy/MM */
function formatYYMM(v) {
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
        d = new Date(Number(s.slice(0, 4)), Number(s.slice(4, 6)) - 1, 1);
      }
    }
  }
  if (!d) {
    const s = String(v || '').trim().replaceAll('-', '/').replaceAll('.', '/');
    const tryD = new Date(s);
    if (!isNaN(tryD.getTime())) d = tryD;
  }
  if (!d) return String(v || '--');
  const yy = String(d.getFullYear()).slice(-2);
  const mm = String(d.getMonth() + 1).padStart(2, '0');
  return `${yy}/${mm}`;
}

/** 讀 A1:B10 控制項（直接回傳數值陣列，前端顯示，不再跳轉） */
function getPolicyControlsRange() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('保單成長');
  if (!sh) throw new Error('找不到「保單成長」分頁');
  return sh.getRange('A1:B10').getDisplayValues(); // 保留你表格上的格式
}



/**====================record=========================*/




/**
 * 依日期回傳 record 的「當日各欄位明細」（供前端子清單使用）
 * @param {string} dateText  前端點擊主列帶過來的日期字串（預期 yyyy-MM-dd）
 * @return {Array<{label:string,value:any,col_index:number}>}
 */


function getRecordDetailsByDate(dateText) {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(RECORD_SHEET_NAME || 'record');
  if (!sh) return [];

  var lastRow = sh.getLastRow();
  var lastCol = sh.getLastColumn();
  if (lastRow < 3 || lastCol < 1) return [];

  // 表頭：第 1 列（帳戶/銀行）、第 2 列（幣別/子分類）
  var hdr1 = sh.getRange(1, 1, 1, lastCol).getDisplayValues()[0] || [];
  var hdr2 = sh.getRange(2, 1, 1, lastCol).getDisplayValues()[0] || [];

  // 想要比對的日期（允許 "yyyy-MM-dd" 或與 A 欄顯示字串一致；另外用 yyyymmdd 兜底）
  var wantDate = null;
  try {
    if (dateText) {
      var tryD = new Date(dateText);
      if (!isNaN(tryD.getTime())) wantDate = tryD;
    }
  } catch (e) { }
  var tz = ss.getSpreadsheetTimeZone() || 'Asia/Taipei';

  function sameYMD(d1, d2) {
    return d1 && d2 && d1.getFullYear() === d2.getFullYear() && d1.getMonth() === d2.getMonth() && d1.getDate() === d2.getDate();
  }
  function ymdDigits(s) {
    s = String(s || '');
    var only = s.replace(/\D+/g, ''); // 只留數字
    if (only.length >= 8) return only.slice(0, 8); // yyyymmdd
    return only;
  }
  var wantYMD = ymdDigits(dateText);

  // 新增：dateToYMD 及 wantDate 正規化
  function dateToYMD(d) {
    if (!(d instanceof Date) || isNaN(d.getTime())) return '';
    var y = d.getFullYear();
    var m = String(d.getMonth() + 1).padStart(2, '0');
    var da = String(d.getDate()).padStart(2, '0');
    return '' + y + m + da;
  }
  // 若 wantDate 尚未解析成功，但拿得到 8 碼數字，手動建 Date 物件
  if (!wantDate && wantYMD && wantYMD.length === 8) {
    var yy = Number(wantYMD.slice(0, 4));
    var mm = Number(wantYMD.slice(4, 6));
    var dd = Number(wantYMD.slice(6, 8));
    var tmp = new Date(yy, mm - 1, dd);
    if (!isNaN(tmp.getTime())) wantDate = tmp;
  }

  function combineLabel(a, b) {
    a = String(a || '').trim();
    b = String(b || '').trim();
    if (a && b && a !== b) return a + ' (' + b + ')';
    return a || b || '';
  }

  // 一次抓 A 欄：原始值 + 顯示字串（第3列起）
  var nData = lastRow - 2;
  var aVals = sh.getRange(3, 1, nData, 1).getValues();         // Date 或原值
  var aDisp = sh.getRange(3, 1, nData, 1).getDisplayValues();   // 顯示字串
  var bDisp = sh.getRange(3, 2, nData, 1).getDisplayValues();   // 月份（可能是 2025-10 / 25-10 / 2025/10）

  // 尋找目標列 index（0-based 相對於第3列）
  var idx = -1;
  for (var i = 0; i < nData; i++) {
    var v = aVals[i][0];
    var s = aDisp[i][0];

    // 1) 以 Date 物件直接比對年月日
    if (wantDate && v instanceof Date && sameYMD(v, wantDate)) { idx = i; break; }

    // 2) 以顯示字串全等比對
    if (String(s || '').trim() === String(dateText || '').trim()) { idx = i; break; }

    // 3) 以純數字 yyyymmdd 比對：同時檢查 顯示字串 與 原始 Date 轉 yyyymmdd
    if (wantYMD) {
      var dispY = ymdDigits(s);
      if (dispY && dispY === wantYMD) { idx = i; break; }
      if (v instanceof Date) {
        var valY = dateToYMD(v);
        if (valY && valY === wantYMD) { idx = i; break; }
      }
    }
    // 4) Fallback：A 欄只有 M/D、B 欄提供年/月 → 組合成 yyyymmdd 再比對
    if (wantYMD && bDisp && bDisp[i] && (s || '')) {
      var bStr = String(bDisp[i][0] || '');     // 例如 '2025-10'、'25/10'
      var y4 = (bStr.match(/(\d{4})[\/-]?(\d{1,2})/) || [])[1];
      var y2 = (!y4 && (bStr.match(/\b(\d{2})[\/-]?(\d{1,2})\b/) || [])[1]) || '';
      var YY = y4 ? y4 : (y2 ? ((Number(y2) < 70 ? '20' : '19') + y2) : '');

      // 從 A 欄顯示字串取出月/日
      var mdDigits = String(s || '').replace(/\D+/g, ''); // '10/16' -> '1016'  ,  '9/2' -> '92'
      var MM = '', DD = '';
      if (mdDigits.length >= 3) {
        if (mdDigits.length === 3) { MM = mdDigits.slice(0, 1); DD = mdDigits.slice(1); }
        else { MM = mdDigits.slice(0, 2); DD = mdDigits.slice(2); }
      }
      if (MM) MM = String(MM).padStart(2, '0');
      if (DD) DD = String(DD).padStart(2, '0');

      if (YY && MM && DD) {
        var ymdJoin = '' + YY + MM + DD;
        if (ymdJoin === wantYMD) { idx = i; break; }
      }
    }
  }
  if (idx < 0) return [];

  var targetRow = 3 + idx; // 真實列號

  // 取該列的原始值（保留數字/日期型別給前端格式化）
  var rowVals = sh.getRange(targetRow, 1, 1, lastCol).getValues()[0] || [];

  // 輸出為 {label, value, col_index}
  var out = new Array(lastCol);
  for (var c = 1; c <= lastCol; c++) {
    var label = combineLabel(hdr1[c - 1], hdr2[c - 1]);
    if (!label && c === 1) label = '日期'; // A 欄無表頭時給預設
    out[c - 1] = { label: label, value: rowVals[c - 1], col_index: c - 1 };
  }
  return out;
}

/**
 * 依「試算表實際列號」回傳該列所有欄位明細（不涉入排序/索引推算）
 * @param {number} rowNumber  真實列號（≥3）
 */
function getRecordDetailsByRowNumber(rowNumber) {
  var sh = SpreadsheetApp.getActive().getSheetByName(RECORD_SHEET_NAME || 'record');
  if (!sh) return [];
  var lastRow = sh.getLastRow();
  var lastCol = sh.getLastColumn();
  var r = Number(rowNumber || 0);
  if (!r || r < 3 || r > lastRow || lastCol < 1) return [];

  var hdr1 = sh.getRange(1, 1, 1, lastCol).getDisplayValues()[0] || [];
  var hdr2 = sh.getRange(2, 1, 1, lastCol).getDisplayValues()[0] || [];
  var rowVals = sh.getRange(r, 1, 1, lastCol).getValues()[0] || [];

  function combineLabel(a, b) {
    a = String(a || '').trim();
    b = String(b || '').trim();
    if (a && b && a !== b) return a + ' (' + b + ')';
    return a || b || '';
  }
  var out = new Array(lastCol);
  for (var c = 1; c <= lastCol; c++) {
    var label = combineLabel(hdr1[c - 1], hdr2[c - 1]);
    if (!label && c === 1) label = '日期';
    out[c - 1] = { label: label, value: rowVals[c - 1], col_index: c - 1 };
  }
  return out;
}


/**
 * 依前端虛擬清單的列索引回傳該列「所有欄位明細」
 * @param {number} rowIndexFromTop 0-based，0 表最新（對應實際表第 3 列）
 * @return {Array<{label:string,value:any,col_index:number}>}
 */
/**
 * 依前端虛擬清單的列索引回傳該列「所有欄位明細」
 */
function getRecordDetailsByRowIndex(rowIndexFromTop) {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(typeof RECORD_SHEET_NAME === 'string' ? RECORD_SHEET_NAME : 'record');
  if (!sh) return [];
  var lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
  if (lastRow < 3 || lastCol < 1) return [];

  function isBlankDisplayRow(r) {
    var disp = sh.getRange(r, 1, 1, lastCol).getDisplayValues()[0];
    for (var i = 0; i < disp.length; i++) { if (String(disp[i] || '').trim() !== '') return false; }
    return true;
  }
  var lastUsed = lastRow;
  while (lastUsed >= 3 && isBlankDisplayRow(lastUsed)) lastUsed--;
  if (lastUsed < 3) return [];

  var hdr1 = sh.getRange(1, 1, 1, lastCol).getDisplayValues()[0] || [];
  var hdr2 = sh.getRange(2, 1, 1, lastCol).getDisplayValues()[0] || [];
  var i = Math.max(0, Number(rowIndexFromTop || 0));
  var nData = lastUsed - 2;
  if (i >= nData) return [];

  var targetA = lastUsed - i;
  var targetB = 3 + i;
  if (targetA < 3 || targetA > lastUsed) targetA = -1;
  if (targetB < 3 || targetB > lastUsed) targetB = -1;

  function rowHasAnyValue(r) {
    if (r < 3) return false;
    var disp = sh.getRange(r, 1, 1, lastCol).getDisplayValues()[0];
    for (var k = 0; k < disp.length; k++) { if (String(disp[k] || '').trim() !== '') return true; }
    return false;
  }

  var targetRow = -1;
  if (targetA !== -1 && rowHasAnyValue(targetA)) targetRow = targetA;
  else if (targetB !== -1 && rowHasAnyValue(targetB)) targetRow = targetB;
  else targetRow = (targetA !== -1 ? targetA : targetB);
  if (targetRow === -1) return [];

  var rowVals = sh.getRange(targetRow, 1, 1, lastCol).getValues()[0] || [];
  function combineLabel(a, b) {
    a = String(a || '').trim(); b = String(b || '').trim();
    if (a && b && a !== b) return a + ' (' + b + ')';
    return a || b || '';
  }
  var out = new Array(lastCol);
  for (var c = 1; c <= lastCol; c++) {
    var label = combineLabel(hdr1[c - 1], hdr2[c - 1]);
    if (!label && c === 1) label = '日期';
    out[c - 1] = { label: label, value: rowVals[c - 1], col_index: c - 1 };
  }
  return out;
}

/**
 * 比較指定列與其「前一筆」（舊一筆，即下一列）的差異
 */
function getRecordDelta(rowIndexFromTop) {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(RECORD_SHEET_NAME || 'record');
  if (!sh) return [];
  var lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
  if (lastRow < 4) return [];

  function isBlankDisplayRow(r) {
    var disp = sh.getRange(r, 1, 1, lastCol).getDisplayValues()[0];
    for (var i = 0; i < disp.length; i++) { if (String(disp[i] || '').trim() !== '') return false; }
    return true;
  }
  var lastUsed = lastRow;
  while (lastUsed >= 3 && isBlankDisplayRow(lastUsed)) lastUsed--;

  var i = Math.max(0, Number(rowIndexFromTop || 0));
  var r1 = lastUsed - i;
  var r2 = r1 - 1;
  if (r1 < 3 || r2 < 3) return [];

  var hdr1 = sh.getRange(1, 1, 1, lastCol).getDisplayValues()[0] || [];
  var hdr2 = sh.getRange(2, 1, 1, lastCol).getDisplayValues()[0] || [];
  var v1 = sh.getRange(r1, 1, 1, lastCol).getValues()[0] || [];
  var v2 = sh.getRange(r2, 1, 1, lastCol).getValues()[0] || [];

  function combineLabel(a, b) {
    a = String(a || '').trim(); b = String(b || '').trim();
    if (a && b && a !== b) return a + ' (' + b + ')';
    return a || b || '';
  }
  var out = [];
  for (var c = 7; c < lastCol; c++) {
    var cur = Number(v1[c]) || 0, prev = Number(v2[c]) || 0, diff = cur - prev;
    if (Math.abs(diff) > 0.001) {
      out.push({ label: combineLabel(hdr1[c], hdr2[c]) || ('Col ' + (c + 1)), current: cur, previous: prev, delta: diff });
    }
  }
  return out;
}
/** 前端按鈕用：呼叫「凍結最後一列 + 新增骨架」 */
function runRecordAddRow() {
  var name = (typeof CONFIG !== 'undefined' && CONFIG.recordSheet) ? CONFIG.recordSheet : 'record';
  addRowAndSnapshot(name);  // 你已經實作好的主程式
  return { ok: true, msg: '已新增一列（H~末欄轉值、A 寫時間），並在下一列鋪回公式骨架。' };
}

/**===================links========================== */

// —— Links 後端 API ——
// 資料表：sheet 名稱 "links"，A: title, B: url
function getLinksSheet_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('links') || ss.insertSheet('links');
  const firstRow = sh.getRange(1, 1, 1, 3).getValues()[0];
  const isHeaderMissing = !firstRow[0] && !firstRow[1] && !firstRow[2];
  if (isHeaderMissing) sh.getRange(1, 1, 1, 3).setValues([['title', 'url', 'note']]);
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
  return s.replace(/\s+/g, ' ').trim();
}



/** === STK 讀取：A..V === */
function getSTKRows() {
  const SHEET = 'STK';
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET);
  if (!sh) return [];
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow <= 1) return [];

  const n = lastRow - 1;
  const values = sh.getRange(2, 1, n, Math.min(22, lastCol)).getValues(); // V=22
  const tz = Session.getScriptTimeZone() || 'Asia/Taipei';
  const toStr = v => v instanceof Date ? Utilities.formatDate(v, tz, 'yyyy-MM-dd') : (v == null ? '' : String(v));

  return values.map((r, i) => ({
    ROW: i + 2,
    A: r[0], B: r[1], C: num(r[2]), D: num(r[3]), E: num(r[4]),
    F: num(r[5]), G: num(r[6]), H: num(r[7]), I: r[8],
    J: num(r[9]), K: num(r[10]),
    L: toStr(r[11]), M: toStr(r[12]), N: num(r[13]),
    O: num(r[14]), P: num(r[15]), Q: num(r[16]),
    R: num(r[17]), S: toStr(r[18]), T: toStr(r[19]), U: toStr(r[20]), V: num(r[21])
  }));

  function num(x) { x = Number(x); return isFinite(x) ? x : 0; }
}

/** STK 偵錯：確認工作表/列數/欄數＋前 5 列樣本（顯示值） */
function getSTKDebug() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('STK');
  if (!sh) return { sheet: false, lastRow: 0, lastCol: 0, hasData: false, sample: [] };
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  const hasData = lastRow > 1;
  const rowsToRead = Math.min(5, Math.max(0, lastRow - 1));
  const colsToRead = Math.min(22, lastCol);
  const sample = rowsToRead > 0 && colsToRead > 0
    ? sh.getRange(2, 1, rowsToRead, colsToRead).getDisplayValues()
    : [];
  return { sheet: true, lastRow, lastCol, hasData, sample };
}

/** === STK 單格寫入通用：可選擇順便寫「更新日」(S 欄) ===
 * @param {number} row           - 實際列號（含表頭從第2列開始，所以通常 >=2）
 * @param {string} colLetter     - 欄位字母（例如 'E' 或 'R'）
 * @param {number|string} value  - 要寫入的值（數字或字串都可）
 * @param {string|null} updateDateColLetter - 若提供（例如 'S'），會一併把該欄寫成今天 yyyy-MM-dd
 */
function setSTKValue(row, colLetter, value, updateDateColLetter) {
  const SHEET = 'STK';
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET);
  if (!sh) throw new Error('找不到工作表：' + SHEET);
  if (!row || row < 2) throw new Error('列號不合法');

  const col = colIndex(colLetter);
  sh.getRange(row, col).setValue(value);

  if (updateDateColLetter) {
    const tz = Session.getScriptTimeZone() || 'Asia/Taipei';
    const today = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
    const uCol = colIndex(updateDateColLetter);
    sh.getRange(row, uCol).setValue(today);
  }
  return true;

  function colIndex(letter) {
    let s = String(letter || '').trim().toUpperCase(), n = 0;
    for (let i = 0; i < s.length; i++) { const code = s.charCodeAt(i); if (code >= 65 && code <= 90) n = n * 26 + (code - 64); }
    return n;
  }
}

/** === 便利包：專寫現價(E)、股息(R) === */
function setSTKPrice(row, price) {
  return setSTKValue(row, 'E', Number(price) || 0, 'S');
}
function setSTKDividend(row, divd) {
  return setSTKValue(row, 'R', Number(divd) || 0, 'S');
}

/** 新增一列到 STK，未列欄位沿用上一列公式（若無上一列，僅寫入值） */
function addSTKItem(item) {
  const SHEET = 'STK';
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET);
  if (!sh) throw new Error('找不到工作表：' + SHEET);

  var needCols = 22; // A..V
  if (sh.getMaxColumns() < needCols) {
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
  if (sourceRow >= dataStart && sourceRow <= sh.getMaxRows()) {
    sh.getRange(sourceRow, 1, 1, needCols).copyTo(
      sh.getRange(targetRow, 1, 1, needCols),
      { contentsOnly: false }
    );
  }

  // A 狀態清空（未賣出）、M 賣日清空
  sh.getRange(targetRow, col('A')).setValue(false);
  sh.getRange(targetRow, col('M')).clearContent();

  // 寫入主要欄位
  if (item.B) sh.getRange(targetRow, col('B')).setValue(item.B);
  if (item.C != null) sh.getRange(targetRow, col('C')).setValue(Number(item.C) || 0);
  if (item.D != null) sh.getRange(targetRow, col('D')).setValue(Number(item.D) || 0);
  if (item.T) sh.getRange(targetRow, col('T')).setValue(String(item.T));
  if (item.R != null) sh.getRange(targetRow, col('R')).setValue(Number(item.R) || 0);
  if (item.L) {
    var tz = Session.getScriptTimeZone() || 'Asia/Taipei';
    var d = new Date(item.L);
    sh.getRange(targetRow, col('L')).setValue(Utilities.formatDate(d, tz, 'yyyy-MM-dd'));
  }
  if (item.U) sh.getRange(targetRow, col('U')).setValue(String(item.U));

  // E 現價：強制改為指定公式（依列號帶入）
  var r = targetRow;
  var formula = '=IF(T' + r + '="US",IF(A' + r + '=TRUE,"sold price",googlefinance(B' + r + ',"PRICE")),IF(A' + r + '=TRUE,"sold price",vlookup(B' + r + ",'🔄️'!B:E,3,0)))";
  sh.getRange(targetRow, col('E')).setFormula(formula);

  return { row: targetRow };

  function col(letter) {
    var s = String(letter || '').trim().toUpperCase(), n = 0;
    for (var i = 0; i < s.length; i++) { var code = s.charCodeAt(i); if (code >= 65 && code <= 90) n = n * 26 + (code - 64); }
    return n;
  }
}

/** 依據賣出數量處理賣出；若部分賣出，會拆分一筆新列保留剩餘股數 */
function splitSell(row, qty, price) {
  const SHEET = 'STK';
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET);
  if (!sh) throw new Error('找不到工作表：' + SHEET);
  var r = Number(row);
  if (!r || r < 2) throw new Error('row 不合法');

  var needCols = 22;
  if (sh.getMaxColumns() < needCols) {
    sh.insertColumnsAfter(sh.getMaxColumns(), needCols - sh.getMaxColumns());
  }

  var currentQty = Number(sh.getRange(r, col('C')).getValue()) || 0;
  if (qty > currentQty) throw new Error('股數不可大於現有股數');

  // 全部賣出：標記 A=已賣出、E=賣價、M=今天
  var tz = Session.getScriptTimeZone() || 'Asia/Taipei';
  var today = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');

  if (qty === currentQty) {
    sh.getRange(r, col('A')).setValue(true);
    sh.getRange(r, col('E')).setValue(Number(price)); // 固定賣價
    sh.getRange(r, col('M')).setValue(today);
    return { row: r, type: 'full' };
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
  var nr = r + 1;
  sh.getRange(r, 1, 1, needCols).copyTo(sh.getRange(nr, 1, 1, needCols), { contentsOnly: false });

  // 新列：未賣出、股數=remain、賣日清空
  sh.getRange(nr, col('A')).setValue(false);
  sh.getRange(nr, col('C')).setValue(Number(remain));
  sh.getRange(nr, col('M')).clearContent();

  // 新列：現價公式
  var formula = '=IF(T' + nr + '="US",IF(A' + nr + '=TRUE,"sold price",googlefinance(B' + nr + ',"PRICE")),IF(A' + nr + '=TRUE,"sold price",vlookup(B' + nr + ",'🔄️'!B:E,3,0)))";
  sh.getRange(nr, col('E')).setFormula(formula);

  return { row: r, remainRow: nr, type: 'partial' };

  function col(letter) {
    var s = String(letter || '').trim().toUpperCase(), n = 0;
    for (var i = 0; i < s.length; i++) { var code = s.charCodeAt(i); if (code >= 65 && code <= 90) n = n * 26 + (code - 64); }
    return n;
  }
}
/** 刪除 STK 指定列（含表頭起算），row>=2 */
function deleteSTKRow(row) {
  const SHEET = 'STK';
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET);
  if (!sh) throw new Error('找不到工作表：' + SHEET);
  const last = sh.getLastRow();
  const r = Number(row);
  if (!r || r < 2 || r > last) throw new Error('列號超出範圍');
  sh.deleteRow(r);
  return { ok: true, row: r };
}


/** ==================== record：V2 快照/分頁 API ==================== */
/** 設定 */
const RECORD_V2_CFG = {
  sheetName: typeof RECORD_SHEET_NAME === 'string' && RECORD_SHEET_NAME ? RECORD_SHEET_NAME : 'record',
  headerRows: 2,       // 第1列+第2列為表頭
  dataStartRow: 3,     // 第3列開始為資料
  cacheKey: 'record_v2_snapshot',
  cacheTtlSec: 60,     // 快照 60 秒有效
};

/** 小工具：取得「有效最後一列」（自尾端往上掃，整列顯示值皆空才略過） */
function recordV2_lastUsedRow_(sh) {
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < RECORD_V2_CFG.dataStartRow || lastCol < 1) return RECORD_V2_CFG.dataStartRow - 1;

  let r = lastRow;
  while (r >= RECORD_V2_CFG.dataStartRow) {
    const disp = sh.getRange(r, 1, 1, lastCol).getDisplayValues()[0];
    const allBlank = disp.every(v => String(v || '').trim() === '');
    if (!allBlank) break;
    r--;
  }
  return r; // < dataStartRow 表示無有效資料
}

/** 取得 V2 meta（不存 cache）：{sheetId,lastCol,lastUsed,total,version} */
function recordV2_getMeta_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(RECORD_V2_CFG.sheetName);
  if (!sh) return { sheet: false, total: 0, lastCol: 0, lastUsed: 0, version: '' };
  const lastCol = sh.getLastColumn();
  const lastUsed = recordV2_lastUsedRow_(sh);
  const total = Math.max(0, lastUsed - (RECORD_V2_CFG.headerRows));
  const version = new Date().toISOString();
  return { sheet: true, sheetId: sh.getSheetId(), lastCol, lastUsed, total, version };
}

/** 建立/覆寫快照（僅存 meta，不存整包資料；索引→實際列的換算公式即可） */
function recordMakeSnapshotV2() {
  const meta = recordV2_getMeta_();
  const cache = CacheService.getDocumentCache();
  cache.put(RECORD_V2_CFG.cacheKey, JSON.stringify(meta), RECORD_V2_CFG.cacheTtlSec);
  return meta;
}

/** 讀取快照（若失效就重建） */
function recordV2_readMeta_() {
  const cache = CacheService.getDocumentCache();
  const s = cache.get(RECORD_V2_CFG.cacheKey);
  if (s) {
    try { return JSON.parse(s); } catch (e) { /* fallthrough */ }
  }
  return recordMakeSnapshotV2();
}

/** Ping：給前端顯示 ready 狀態 */
function recordPingV2() {
  const meta = recordV2_readMeta_();
  return { total: meta.total || 0, version: meta.version || '' };
}

/**
 * 分頁清單：回 6 欄
 *  A(日期顯示字串)｜C(總金額)｜D(總金額-含dad)｜E(可動用)｜F(損益)｜G(可用現金)
 *  並附上 idx（0=最新在上）與實際 sheet 列號 row
 * @param {{offset?:number, limit?:number}} opt
 */
function recordListV2(opt) {
  opt = opt || {};
  const offset = Math.max(0, Number(opt.offset || 0));
  const limit = Math.max(1, Number(opt.limit || 200));

  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(RECORD_V2_CFG.sheetName);
  if (!sh) return { meta: { total: 0 }, rows: [] };

  const meta = recordV2_readMeta_();
  const total = meta.total || 0;
  if (total <= 0) return { meta, rows: [] };

  const from = Math.min(offset, total);
  const to = Math.min(offset + limit, total);
  const count = Math.max(0, to - from);
  if (count <= 0) return { meta, rows: [] };

  // 將「虛擬 idx（最新在上）」映射到「實際列號」
  // latestIdx=0 → realRow = meta.lastUsed - 0
  // latestIdx=k → realRow = meta.lastUsed - k
  const rows = [];
  const tz = ss.getSpreadsheetTimeZone() || 'Asia/Taipei';
  for (let i = 0; i < count; i++) {
    const idx = from + i;                       // 0-based（最新在上）
    const realRow = meta.lastUsed - idx;        // 實際列號
    if (realRow < RECORD_V2_CFG.dataStartRow) break;

    // 一次取 A..G（7欄），保留數字型別
    const vals = sh.getRange(realRow, 1, 1, 7).getValues()[0];
    const dispA = sh.getRange(realRow, 1, 1, 1).getDisplayValues()[0][0]; // A 欄顯示字串當日期

    rows.push({
      idx,
      row: realRow,
      A: String(dispA || ''),     // 日期（顯示）
      C: Number(vals[2] || 0),
      D: Number(vals[3] || 0),
      E: Number(vals[4] || 0),
      F: Number(vals[5] || 0),
      G: Number(vals[6] || 0),
    });
  }
  return { meta: { total, version: meta.version, lastCol: meta.lastCol }, rows };
}

/**
 * 子清單：依 idx 反查實際列號 → 回傳整列的 {label,value,col_index}[]
 * @param {{idx:number}} payload
 */
function recordDetailV2(payload) {
  const idx = Math.max(0, Number(payload && payload.idx || 0));
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(RECORD_V2_CFG.sheetName);
  if (!sh) return [];

  const meta = recordV2_readMeta_();
  const total = meta.total || 0;
  if (total <= 0 || idx >= total) return [];

  const realRow = meta.lastUsed - idx; // 由 idx 換算實際列號
  if (realRow < RECORD_V2_CFG.dataStartRow) return [];

  // 直接沿用你已存在的函式
  return getRecordDetailsByRowNumber(realRow) || [];
}

/** === V2 Self-Test / Debug === */
function recordV2_selfTest() {
  // 快速驗證後端：回 meta + 前 3 筆列表（若有）
  var meta = recordMakeSnapshotV2();
  var page = recordListV2({ offset: 0, limit: 3 });
  return { ok: true, meta: meta, rows: (page && page.rows) || [] };
}
function recordV2_debugRow(idx) {
  // 檢查某 idx → realRow 映射是否正常，並回傳前 6 欄顯示值
  idx = Math.max(0, Number(idx || 0));
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(RECORD_V2_CFG.sheetName);
  if (!sh) return { ok: false, msg: 'no sheet' };
  var meta = recordV2_readMeta_();
  if (!meta || !meta.total) return { ok: false, msg: 'no meta/total' };
  if (idx >= meta.total) return { ok: false, msg: 'idx>=total', meta: meta };
  var realRow = meta.lastUsed - idx;
  var disp = sh.getRange(realRow, 1, 1, Math.min(6, sh.getLastColumn())).getDisplayValues()[0] || [];
  return { ok: true, meta: meta, idx: idx, realRow: realRow, preview: disp };
}


/***** routines.gs *****/

// 你 routines 分頁的名稱
const ROUTINE_SHEET = 'routines';

// 欄位對照（1-based column index in sheet）
const ROUTINE_COLS = {
  active: 1,  // A
  name: 2,  // B
  cat: 3,  // C
  amount: 4,  // D
  currency: 5,  // E
  freq: 6,  // F
  nextDueDate: 7,  // G
  autoPost: 8,  // H
  acc: 9,  // I
  owner: 10, // J
  dir: 11, // K
  rollupOnly: 12, // L
  lastPostedMonth: 13 // M (不顯示在 UI，但會一起回傳，之後可能用得到)
};


// 後端：讀 routines 全部列
function getRoutinesData() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(ROUTINE_SHEET);
  if (!sh) return { rows: [], meta: {} };

  const lastRow = sh.getLastRow();
  if (lastRow < 2) {
    return { rows: [], meta: {} };
  }

  // 把第2列到最後一列抓出來
  const rng = sh.getRange(2, 1, lastRow - 1, sh.getLastColumn());
  const values = rng.getValues();

  const tz = ss.getSpreadsheetTimeZone() || 'Asia/Taipei';

  const rows = values.map((r, idx) => {
    const rowNumber = 2 + idx; // 真實列號（之後更新會用）
    const d = r[ROUTINE_COLS.nextDueDate - 1];
    // 格式化日期給前端
    const nextDue =
      d instanceof Date
        ? Utilities.formatDate(d, tz, 'yyyy-MM-dd')
        : (d ? String(d) : '');

    return {
      _row: rowNumber,
      active: r[ROUTINE_COLS.active - 1] === true,
      name: String(r[ROUTINE_COLS.name - 1] || ''),
      cat: String(r[ROUTINE_COLS.cat - 1] || ''),
      amount: Number(r[ROUTINE_COLS.amount - 1] || 0),
      currency: String(r[ROUTINE_COLS.currency - 1] || ''),
      freq: String(r[ROUTINE_COLS.freq - 1] || ''),
      nextDueDate: nextDue,
      autoPost: r[ROUTINE_COLS.autoPost - 1] === true,
      acc: String(r[ROUTINE_COLS.acc - 1] || ''),
      owner: String(r[ROUTINE_COLS.owner - 1] || ''),
      dir: String(r[ROUTINE_COLS.dir - 1] || ''),
      rollupOnly: r[ROUTINE_COLS.rollupOnly - 1] === true,
      lastPostedMonth: String(r[ROUTINE_COLS.lastPostedMonth - 1] || '')
    };
  });

  // 這裡把下拉清單的候選值也傳給前端（Acc, Owner, Freq 等）
  // 目前你說 Acc 固定：中信, 台新, 星展, Cash, Richart, 貸款, 聯邦, ECO
  // Owner：漢, 媽, 薰, 昕
  // Freq：monthly / yearly / bi-monthly
  const meta = {
    accList: ['中信', '台新', '星展', 'Cash', 'Richart', '貸款', '聯邦', 'ECO'],
    ownerList: ['漢', '媽', '薰', '昕'],
    freqList: ['monthly', 'bi-monthly', 'yearly'],
    dirList: ['out', 'in', 'xfer'], // 這個我先假設，可自行改
    currencyList: ['TWD', 'USD', 'JPY', 'KRW'] // 也可以依照你實際用的幣別
  };

  return { rows, meta };
}

// 後端：更新單一欄位（例如切 active toggle、改 Acc、改金額…）
// fieldName 必須是我們支援的 key（active/name/amount/...）
// value 是前端送來的新值
function updateRoutineField(rowNumber, fieldName, value) {
  const sh = SpreadsheetApp.getActive().getSheetByName(ROUTINE_SHEET);
  if (!sh) throw new Error('找不到 sheet: ' + ROUTINE_SHEET);

  const colIdx = ROUTINE_COLS[fieldName];
  if (!colIdx) throw new Error('不支援欄位: ' + fieldName);

  // 特別處理 boolean toggle 欄位
  if (fieldName === 'active' || fieldName === 'autoPost' || fieldName === 'rollupOnly') {
    sh.getRange(rowNumber, colIdx).setValue(!!value);
    return { ok: true };
  }

  // 特別處理日期
  if (fieldName === 'nextDueDate') {
    // 預期 value 是 "yyyy-MM-dd" 字串
    if (value) {
      sh.getRange(rowNumber, colIdx).setValue(new Date(value));
    } else {
      sh.getRange(rowNumber, colIdx).setValue('');
    }
    return { ok: true };
  }

  // 其他純值 (文字/數字)
  sh.getRange(rowNumber, colIdx).setValue(value);
  return { ok: true };
}

// （可選）後端：新增一列空白 routine（按"新增"→確認）
// 會在表頭下一列插入，並回傳新的 rowNumber，前端可直接編輯
function addRoutineBlank() {
  const sh = SpreadsheetApp.getActive().getSheetByName(ROUTINE_SHEET);
  if (!sh) throw new Error('找不到 sheet: ' + ROUTINE_SHEET);

  // 我們要插到表頭下一列，也就是第2列
  sh.insertRowAfter(1);
  const newRow = 2;

  // 預設值
  sh.getRange(newRow, ROUTINE_COLS.active).setValue(true);
  sh.getRange(newRow, ROUTINE_COLS.name).setValue('');
  sh.getRange(newRow, ROUTINE_COLS.amount).setValue(0);
  sh.getRange(newRow, ROUTINE_COLS.freq).setValue('monthly');
  sh.getRange(newRow, ROUTINE_COLS.nextDueDate).setValue(new Date());
  sh.getRange(newRow, ROUTINE_COLS.autoPost).setValue(false);
  sh.getRange(newRow, ROUTINE_COLS.rollupOnly).setValue(false);

  return { row: newRow };
}

function deleteRoutineRow(rowNumber) {
  const sh = SpreadsheetApp.getActive().getSheetByName(ROUTINE_SHEET);
  if (!sh) throw new Error('找不到 sheet: ' + ROUTINE_SHEET);

  sh.deleteRow(rowNumber);
  return { ok: true };
}

/**
 * 例行支出自動寫入 Cashflow
 * - 對應舊觸發器名稱：job_routinesAutoPost（可直接用時間觸發器叫這支）
 * - 規則：
 *   - 只處理 active=true 且 autoPost=true，且 rollupOnly != true 的列
 *   - 依 nextDueDate 為基準，一路往後推 freq（monthly / bi-monthly / yearly）
 *   - 每一個 nextDueDate 只會寫入一次（同一個 ym 不重複），寫入後：
 *       - 呼叫 submitCashflow() 產生 Cashflow 列（status='auto' → 用公式判斷 posted/planned）
 *       - 更新 routines.nextDueDate 為下一期
 *       - 更新 routines.lastPostedMonth 為最後一次寫入的 yyyymm
 *   - 補到的上限為「下個月的 10 號」：避免一次衝太遠，只看目前到下個月 10 號之前的項目
 */
function job_routinesAutoPost() {
  const ss = SpreadsheetApp.getActive();
  const shR = ss.getSheetByName(ROUTINE_SHEET);
  const shCF = ss.getSheetByName(CF_SHEET);
  if (!shR || !shCF) return;

  const tz = ss.getSpreadsheetTimeZone() || 'Asia/Taipei';
  const today = new Date();

  // 上限：下個月 10 號
  const curY = today.getFullYear();
  const curM0 = today.getMonth(); // 0-based
  let nextY = curY;
  let nextM0 = curM0 + 1;
  if (nextM0 >= 12) {
    nextM0 -= 12;
    nextY += 1;
  }
  const limitDate = new Date(nextY, nextM0, 10); // JS 月份 0-based
  const limitYM = nextY * 100 + (nextM0 + 1);

  const lastRow = shR.getLastRow();
  if (lastRow < 2) return;

  const lastCol = shR.getLastColumn();
  const vals = shR.getRange(2, 1, lastRow - 1, lastCol).getValues();

  for (let i = 0; i < vals.length; i++) {
    const rowVals = vals[i];
    const rowNo = 2 + i;

    const active = rowVals[ROUTINE_COLS.active - 1] === true;
    const autoPost = rowVals[ROUTINE_COLS.autoPost - 1] === true;
    const rollupOnly = rowVals[ROUTINE_COLS.rollupOnly - 1] === true;

    if (!active || !autoPost || rollupOnly) continue;

    let due = rowVals[ROUTINE_COLS.nextDueDate - 1];
    if (!(due instanceof Date)) continue;

    let lastPostedYM = Number(rowVals[ROUTINE_COLS.lastPostedMonth - 1] || 0);
    const freq = String(rowVals[ROUTINE_COLS.freq - 1] || '').toLowerCase();

    const name = String(rowVals[ROUTINE_COLS.name - 1] || '');
    const rawAmt = Number(rowVals[ROUTINE_COLS.amount - 1] || 0);
    const acc = String(rowVals[ROUTINE_COLS.acc - 1] || '');
    const dir = String(rowVals[ROUTINE_COLS.dir - 1] || 'out').toLowerCase();

    if (!rawAmt || !acc || !name) {
      // 若基本欄位不完整，就不自動寫入
      continue;
    }

    let changed = false;

    // 只要 nextDueDate 還在上限之前，就一路補到 limitDate
    while (due && due <= limitDate) {
      const ym = due.getFullYear() * 100 + (due.getMonth() + 1);
      if (ym > limitYM) break;

      // 已經處理過這個月份就跳過，但仍要往下一期推，避免死循環
      if (ym <= lastPostedYM) {
        due = _routine_nextDate_(due, freq);
        changed = true;
        continue;
      }

      // 決定金額方向：
      // dir='out' 或 'ME_OWE' → 視為支出（負數）
      // dir='in' 或 'THEY_OWE' → 視為收入（正數）
      // 其他照原值
      let amt = rawAmt;
      if (dir === 'out' || dir === 'me_owe') {
        amt = -Math.abs(rawAmt);
      } else if (dir === 'in' || dir === 'they_owe') {
        amt = Math.abs(rawAmt);
      }

      try {
        submitCashflow({
          date: due,
          item: name,
          amount: amt,
          account: acc,
          status: 'auto',           // 讓 status 欄用公式自動判斷 posted / planned
          note: '[routine] ' + name // 你可依喜好調整字樣
        });
      } catch (e) {
        console.warn('[job_routinesAutoPost] submitCashflow failed on row', rowNo, e);
        break; // 這筆出錯就先不要再往下推，以免亂寫
      }

      lastPostedYM = ym;
      due = _routine_nextDate_(due, freq);
      changed = true;
    }

    // 有寫入才回存 nextDueDate / lastPostedMonth
    if (changed) {
      shR.getRange(rowNo, ROUTINE_COLS.nextDueDate).setValue(due);
      shR.getRange(rowNo, ROUTINE_COLS.lastPostedMonth).setValue(lastPostedYM || '');
    }
  }
  SpreadsheetApp.flush();
}

/**
 * 依 freq 推算下一期日期
 * @param {Date} d
 * @param {string} freq 'monthly' | 'bi-monthly' | 'yearly' | ...
 * @return {Date}
 */
function _routine_nextDate_(d, freq) {
  const base = new Date(d.getTime());
  const f = String(freq || '').toLowerCase();
  if (f === 'monthly') {
    base.setMonth(base.getMonth() + 1);
  } else if (f === 'bi-monthly' || f === 'bimonthly' || f === 'bi_monthly') {
    base.setMonth(base.getMonth() + 2);
  } else if (f === 'yearly' || f === 'annual' || f === 'year') {
    base.setFullYear(base.getFullYear() + 1);
  } else {
    // 預設當作 monthly
    base.setMonth(base.getMonth() + 1);
  }
  return base;
}

// ===== Dash 用：沿用 page_input 的計算規則（後端版） =====

/** 後端版：依帳戶計算 posted Balance（含 Current 快照邏輯） */
function _dash_computePostedBalancesByAccount_(rows) {
  rows = Array.isArray(rows) ? rows : [];
  // 只取 posted
  var posted = rows.filter(function (r) {
    return String(r.status || '').toLowerCase() === 'posted';
  });

  var byAcct = {};
  posted.forEach(function (r) {
    var acct = String(r.account || '').trim();
    if (!acct) return;
    if (!byAcct[acct]) byAcct[acct] = [];
    byAcct[acct].push(r);
  });

  var out = [];
  Object.keys(byAcct).forEach(function (acct) {
    var list = byAcct[acct];

    // Current 列
    var currentRows = list.filter(function (r) {
      return /\bcurrent\b/i.test(String(r.item || ''));
    });

    var base = 0;
    var snapTs = -1;

    if (currentRows.length) {
      // 找最新一筆 Current（依日期）
      var latest = currentRows
        .map(function (r) {
          return {
            amount: Number(r.amount) || 0,
            _ts: Date.parse(r.date || '') || 0
          };
        })
        .sort(function (a, b) { return b._ts - a._ts; })[0];
      base = latest.amount;
      snapTs = latest._ts;
    }

    // 其他 posted 列（非 Current）
    var others = list.filter(function (r) {
      return !/\bcurrent\b/i.test(String(r.item || ''));
    });

    // 修正：只扣除「日期 > Snapshot」的交易
    var sumOthers = others.reduce(function (t, r) {
      var rTs = Date.parse(r.date || '') || 0;
      if (rTs > snapTs) {
        return t + (Number(r.amount) || 0);
      }
      return t;
    }, 0);

    // 規則：balance = base - sumOthers
    out.push({ account: acct, balance: base - sumOthers });
  });

  out.sort(function (a, b) {
    return String(a.account).localeCompare(String(b.account), 'zh-Hant-TW');
  });
  return out;
}

/** 後端版：依帳戶計算 planned 總額（顯示為 -金額） */
function _dash_computePlannedTotalsByAccount_(rows) {
  rows = Array.isArray(rows) ? rows : [];
  var byAcct = {};
  rows.forEach(function (r) {
    var sta = String(r.status || '').toLowerCase();
    if (sta !== 'planned') return;

    // 排除 Current
    if (/\bcurrent\b/i.test(String(r.item || ''))) return;

    var acct = String(r.account || '').trim();
    if (!acct) return;

    var v = -(Number(r.amount) || 0); // planned 顯示為 -金額
    byAcct[acct] = (byAcct[acct] || 0) + v;
  });

  var out = [];
  Object.keys(byAcct).forEach(function (acct) {
    out.push({ account: acct, balance: byAcct[acct] });
  });
  out.sort(function (a, b) {
    return String(a.account).localeCompare(String(b.account), 'zh-Hant-TW');
  });
  return out;
}

/** 後端版：依帳戶計算 using 總額（顯示為 -金額，和 input.js 相同） */
function _dash_computeUsingTotalsByAccount_(rows) {
  rows = Array.isArray(rows) ? rows : [];
  var byAcct = {};
  rows.forEach(function (r) {
    var sta = String(r.status || '').toLowerCase();
    if (sta !== 'using') return;

    // 排除 Current
    if (/\bcurrent\b/i.test(String(r.item || ''))) return;

    var acct = String(r.account || '').trim();
    if (!acct) return;

    var v = -(Number(r.amount) || 0); // using 顯示為 -金額（支出為正）
    byAcct[acct] = (byAcct[acct] || 0) + v;
  });

  var out = [];
  Object.keys(byAcct).forEach(function (acct) {
    out.push({ account: acct, balance: byAcct[acct] });
  });
  out.sort(function (a, b) {
    return String(a.account).localeCompare(String(b.account), 'zh-Hant-TW');
  });
  return out;
}

function getDashBalanceSummary() {
  var rows = getTransactions({}) || [];

  var posted = _dash_computePostedBalancesByAccount_(rows);
  var planned = _dash_computePlannedTotalsByAccount_(rows);
  var using = _dash_computeUsingTotalsByAccount_(rows);

  var postedMap = {};
  var plannedMap = {};
  var usingMap = {};

  posted.forEach(function (r) {
    postedMap[String(r.account)] = Number(r.balance) || 0;
  });
  planned.forEach(function (r) {
    plannedMap[String(r.account)] = Number(r.balance) || 0;
  });
  using.forEach(function (r) {
    usingMap[String(r.account)] = Number(r.balance) || 0;
  });

  function mark(acct, set) {
    if (!acct) return;
    set[acct] = true;
  }

  var acctSet = {};
  Object.keys(postedMap).forEach(function (a) { mark(a, acctSet); });
  Object.keys(plannedMap).forEach(function (a) { mark(a, acctSet); });
  Object.keys(usingMap).forEach(function (a) { mark(a, acctSet); });

  var acctList = Object.keys(acctSet).sort(function (a, b) {
    return String(a).localeCompare(String(b), 'zh-Hant-TW');
  });

  function isLoanAccount(name) {
    var s = String(name || '');
    return s.indexOf('貸款') !== -1;
  }
  function isPinnedAccount(name) {
    var s = String(name || '');
    var sLow = s.toLowerCase();
    if (isLoanAccount(name)) return false;
    return s.indexOf('中信') !== -1 || sLow === 'cash';
  }

  var loan = [];
  var normal = [];

  acctList.forEach(function (acct) {
    var pln = Number(plannedMap[acct] || 0);
    var pst = Number(postedMap[acct] || 0);
    var useAgg = Number(usingMap[acct] || 0); // = -(原始 using 總和)

    if (!pln && !pst && !useAgg) return;

    if (isLoanAccount(acct)) {
      // 貸款帳戶：posted / using / planned / sum
      var usingRaw = -useAgg;              // 原始方向
      var postedDisplay = pst - usingRaw;       // posted 要扣掉 using
      var sumLoan = postedDisplay + usingRaw + pln; // = posted + using + planned

      loan.push({
        account: acct,
        posted: postedDisplay,
        using: usingRaw,
        planned: pln,
        sum: sumLoan
      });
    } else {
      // 一般帳戶：排除固定置頂（中信 / Cash）
      if (isPinnedAccount(acct)) return;
      var sum = pst + pln;                      // 一般帳戶 sum 定義 = posted + planned
      normal.push({
        account: acct,
        posted: pst,
        planned: pln,
        sum: sum
      });
    }
  });

  loan.sort(function (a, b) {
    return String(a.account).localeCompare(String(b.account), 'zh-Hant-TW');
  });
  normal.sort(function (a, b) {
    return String(a.account).localeCompare(String(b.account), 'zh-Hant-TW');
  });

  // ★ 這裡是關鍵：只算「一般帳戶」的 sum
  var normalTotalSum = normal.reduce(function (t, r) {
    return t + (Number(r.sum) || 0);
  }, 0);

  // 如果你之後想看貸款合計，也幫你算好
  var loanTotalSum = loan.reduce(function (t, r) {
    return t + (Number(r.sum) || 0);
  }, 0);

  return {
    loan: loan,
    normal: normal,
    normalTotalSum: normalTotalSum,  // Dash 卡片總計請用這個
    loanTotalSum: loanTotalSum       // 目前不用，單純備用
  };
}