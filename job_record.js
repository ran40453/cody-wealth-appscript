/** ========= jobs_record.gs (R1C1 版) =========
 * 凍結最後一列 (H~end 轉值，A 寫入靜態時間)，並在下一列鋪回 R1C1 公式骨架。
 */
function addRowAndSnapshot(sheetName){
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(sheetName);
  if (!sh) throw new Error('找不到工作表：' + sheetName);

  SpreadsheetApp.flush(); // 先讓現有公式算完

  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 2) { sh.insertRowAfter(lastRow); return; }

  // A~G（不含 H），抓 R1C1 公式骨架
  const widthAG = COL.H - COL.A; // =7
  const agRng   = sh.getRange(lastRow, COL.A, 1, widthAG);
  const agR1C1  = agRng.getFormulasR1C1();

  // H ~ end，抓 R1C1 公式骨架
  const hWidth  = lastCol - COL.H + 1;
  const hRng    = sh.getRange(lastRow, COL.H, 1, hWidth);
  const hR1C1   = hRng.getFormulasR1C1();

  // 凍結最後一列：H~end 轉值、A 寫入日期時間
  hRng.setValues(hRng.getValues());
  sh.getRange(lastRow, COL.A).setValue(new Date());
  SpreadsheetApp.flush();

  // 新增一列 → 鋪回 R1C1 骨架（自動位移）
  sh.insertRowAfter(lastRow);
  const newRow = lastRow + 1;

  if (hasAnyFormula_(agR1C1)) {
    sh.getRange(newRow, COL.A, 1, widthAG).setFormulasR1C1(agR1C1);
  } else {
    sh.getRange(newRow, COL.A).setFormulaR1C1('=IF(LEN(R[-1]C), NOW(), "")');
  }
  if (hasAnyFormula_(hR1C1)) {
    sh.getRange(newRow, COL.H, 1, hWidth).setFormulasR1C1(hR1C1);
  }
  SpreadsheetApp.flush();
}

function hasAnyFormula_(r1c1_2d){
  if (!r1c1_2d || !r1c1_2d.length) return false;
  for (let r = 0; r < r1c1_2d.length; r++){
    for (let c = 0; c < r1c1_2d[r].length; c++){
      if (r1c1_2d[r][c]) return true;
    }
  }
  return false;
}