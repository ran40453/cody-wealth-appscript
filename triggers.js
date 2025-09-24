/** ========= triggers.gs (hardened) =========
 * 唯一 onEdit：集中判斷 → 呼叫對應 handler
 */
function onEdit(e) {
  try {
    if (!e) return;

    const sh   = e.range.getSheet();
    const name = sh.getName();
    const r    = e.range.getRow();
    const c    = e.range.getColumn();
    const vStr = (typeof e.value === 'string' ? e.value : String(e.value || '')).toUpperCase();

    // ------- record：B1 勾選 → 凍結最後一列 + 鋪新列骨架（R1C1 版）
    if (name === CONFIG.recordSheet && r === 1 && c === COL.B) {
      const checked = vStr === 'TRUE' || (typeof e.range.isChecked === 'function' && e.range.isChecked());
      if (checked) {
        SpreadsheetApp.getActive().toast('Record：新增一列並凍結上一列…', 'onEdit', 2);
        addRowAndSnapshot(name);       // ← 確保用的是 R1C1 寫回的那一版（見第 2 段）
        e.range.setValue(false);       // 反勾回去
        return;
      }
    }

    // ------- M-Rpt：C1 勾選 → 新增列（沿用你的 addRowBasic）
    if (name === CONFIG.mrpSheet && r === 1 && c === 3 && vStr === 'TRUE') {
      if (typeof addRowBasic === 'function') addRowBasic(name);
      e.range.setValue(false);
      return;
    }

    // ------- ✅ 表（timestamp）：Q(17) 或 AF(32) 單格改動 → AH(34) 寫時間戳
    if (name === CONFIG.timestampSheet &&
        r > 2 && (c === 17 || c === 32) &&
        e.range.getNumRows() === 1 && e.range.getNumColumns() === 1 &&
        e.value !== '' && e.value != null) {
      sh.getRange(r, 34).setValue(new Date());
      return;
    }

    // ------- GRW：H1(8)/X1(24) 任一變動 → Y1(25) 寫刷新時間
    if (name === CONFIG.grwSheet && (c === 8 || c === 24)) {
      const x = sh.getRange(1, 24).getValue();
      const h = sh.getRange(1, 8).getValue();
      if (x != e.value || h != e.value) sh.getRange(1, 25).setValue(new Date());
      return;
    }

    // ------- Cashflow：最後一列新增時，自動補格式/驗證
    if (name === CF.sheetName) {
      const last = sh.getLastRow();
      for (let rr = r; rr < r + e.range.getNumRows(); rr++) {
        if (rr === last && rr > CF.headerRow && typeof applyFormatAndValidation_ === 'function') {
          applyFormatAndValidation_(rr);
        }
      }
      return;
    }
  } catch (err) {
    SpreadsheetApp.getActive().toast('onEdit 錯誤：' + (err && err.message || err), 'Error', 5);
    throw err;
  }
}

function onOpen(){
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Cody Wealth')
    .addItem('重裝觸發器', 'installTriggers')
    .addSeparator()
    .addItem('修復 Cashflow 格式/下拉', 'fixMissingValidationAll_') // 若該函式存在
    .addItem('夜間快照（立即）', 'job_nightlySnapshot')           // 在 jobs.gs
    .addItem('更新 Dashboard 快取（立即）', 'job_refreshDashboardCache')
    .addToUi();
}