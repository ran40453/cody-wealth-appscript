/** ========= cashflow_format.gs =========
 * Cashflow：套用第 2 列的「外觀 + 下拉驗證」
 */

function applyFormatAndValidation_(row){
  const sh = SpreadsheetApp.getActive().getSheetByName(CONFIG.cashflowSheet);
  if (!sh) return;
  const width = CF.lastCol - CF.firstCol + 1;
  const src = sh.getRange(CF.templateRow, CF.firstCol, 1, width);
  const dst = sh.getRange(row,            CF.firstCol, 1, width);

  // 外觀
  src.copyTo(dst, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  // 下拉驗證
  src.copyTo(dst, SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false);
}

/** 手動一鍵修復（選單按鈕） */
function fixMissingValidationAll_(){
  const sh = SpreadsheetApp.getActive().getSheetByName(CONFIG.cashflowSheet);
  if (!sh) return;
  const lastRow = sh.getLastRow();
  if (lastRow <= CF.headerRow) return;

  const width = CF.lastCol - CF.firstCol + 1;
  const rng = sh.getRange(CF.headerRow + 1, CF.firstCol, lastRow - CF.headerRow, width);
  const dv  = rng.getDataValidations(); // 2D

  const startRow = CF.headerRow + 1;
  for (let r = 0; r < dv.length; r++){
    if (dv[r].some(cell => cell == null)) {
      applyFormatAndValidation_(startRow + r);
    }
  }
  SpreadsheetApp.getUi().alert('已檢查並補齊缺失的格式與下拉驗證。');
}