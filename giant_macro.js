/** ========= jobs_record.gs =========
 * `record` 工作表專用：
 * 凍結最後一列 (H~end 轉值，A 寫入靜態時間)，並在下一列鋪回公式骨架（用 copyTo 讓相對參照自動位移）。
 */
function addRowAndSnapshot(sheetName){
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(sheetName);
  if (!sh) throw new Error('找不到工作表：' + sheetName);

  SpreadsheetApp.flush(); // 先讓現有公式算完

  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 2) {            // 只有標題或空表時
    sh.insertRowAfter(lastRow); // 直接加一列
    return;
  }

  // 1) 在「末列」後面新增一列
  sh.insertRowAfter(lastRow);
  const newRow = lastRow + 1;

  // 2) 用 copyTo 複製整列（格式/公式/驗證），相對參照會自動往下位移 1 列
  const source = sh.getRange(lastRow, 1, 1, lastCol);
  const target = sh.getRange(newRow, 1, 1, lastCol);
  source.copyTo(target, { contentsOnly: false });

  // 3) 凍結舊末列：H ~ 最末欄 轉成值、A 欄寫靜態時間
  const hStart = COL.H;                       // 依你專案的常數（例：COL.A=1, COL.H=8）
  const hWidth = lastCol - hStart + 1;
  const hEndRange = sh.getRange(lastRow, hStart, 1, hWidth);
  hEndRange.setValues(hEndRange.getValues()); // 公式 → 值

  sh.getRange(lastRow, COL.A).setValue(new Date()); // A 欄寫入時間戳

  SpreadsheetApp.flush();
}

/** 工具：2D 公式陣列是否至少含一個公式字串（保留可用在其他情境） */
function arrayHasAnyFormula_(formulas2d){
  if (!formulas2d || !formulas2d.length) return false;
  for (let r = 0; r < formulas2d.length; r++){
    const row = formulas2d[r] || [];
    for (let c = 0; c < row.length; c++){
      if (row[c]) return true;
    }
  }
  return false;
}

/** 舊程式可能會呼叫的 stub，保留避免錯誤 */
function clearCheckBox(){ /* no-op */ }