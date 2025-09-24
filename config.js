/** ========= config.gs ========= */

// 工作表名稱
const CONFIG = {
  recordSheet:     'record',
  mrpSheet:        'M-Rpt',
  timestampSheet:  '✅',
  grwSheet:        'GRW',
  timezone:        'Asia/Taipei',
  cashflowSheet:   'Cashflow' // 新增這一行
};

// 通用欄位（1-based）
const COL = {
  A:1, B:2, C:3, D:4, E:5, F:6, G:7, H:8,
  BB:53, CV:99
};

// Cashflow 欄位（給格式/驗證用）
const CF = {
  sheetName: 'Cashflow',
  headerRow: 1,          // 表頭列
  templateRow: 2,        // 以第 2 列當範本
  firstCol: 1,           // A
  lastCol:  6            // F（你的表單實際範圍：日期/明細/金額/帳戶/狀態/備註）
};