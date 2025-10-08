/** ========= config.gs (Apps Script) ========= */

// 工作表名稱 / 時區等全域設定
const CONFIG = {
  recordSheet:    'record',
  mrpSheet:       'M-Rpt',
  timestampSheet: '✅',
  grwSheet:       'GRW',
  cashflowSheet:  'Cashflow',
  timezone:       'Asia/Taipei'
};

/**
 * 以欄位字母回傳 1-based 欄號。
 * 例：A→1, Z→26, AA→27, AZ→52, BA→53, BB→54, CU→100 (C=3, V=22; 3*26+22=100)
 */
function COL_INDEX(letter){
  var s = String(letter||'').trim().toUpperCase();
  var n = 0;
  for (var i=0;i<s.length;i++){
    var code = s.charCodeAt(i);
    if (code >= 65 && code <= 90) n = n*26 + (code - 64);
  }
  return n;
}

// 常用欄位快速存取（會自動以正確索引計算，避免手動數錯/欄位位移）
const COL = {
  // 基礎 A~H
  A: COL_INDEX('A'), B: COL_INDEX('B'), C: COL_INDEX('C'), D: COL_INDEX('D'),
  E: COL_INDEX('E'), F: COL_INDEX('F'), G: COL_INDEX('G'), H: COL_INDEX('H'),
  // 你常用的遠端欄位（注意：原先 AQ 被移除後，BB→BA、CV→CU 等位移請直接用字母）
  BA: COL_INDEX('BA'),
  BB: COL_INDEX('BB'),
  CU: COL_INDEX('CU'),
  CV: COL_INDEX('CV')
};

// Cashflow 欄位範圍（給格式/驗證用，表單實際範圍：日期/明細/金額/帳戶/狀態/備註）
const CF = {
  sheetName: 'Cashflow',
  headerRow: 1,          // 表頭列
  templateRow: 2,        // 以第 2 列當範本
  firstCol: COL.A,       // A
  lastCol:  COL.F        // F
};

if (typeof window !== 'undefined') window.APP_CONFIG = {
  // Dashboard KPI 卡位的 DOM id（改這裡就好）
  KPI_IDS: {
    total: 'kpiTotal',
    net:   'kpiNet',    // 以 D 欄 Whole Assets 顯示
    cash:  'kpiCash',
    avail: 'kpiAvail',
    usd:   'kpiUSD'     // 取代過去的 debt 顯示
  },

  // 從後端 latest 物件取值時使用的鍵（後端固定回傳英文字母欄位）
  // AQ 拔除後：BB→BA、CV→CU
  LATEST_KEYS: {
    date:  'date',
    total: 'C',   // 總資產
    whole: 'D',   // Whole Assets（你要的「淨資產」顯示來源）
    avail: 'E',   // 可動用
    pnl:   'F',   // 當日損益
    cash:  'G',   // 現金
    debt:  'BA',  // 債務（仍可在 record 頁顯示/備查）
    usd:   'CU'   // 持有 USD
  },

  // record 頁左側摘要卡要顯示哪些欄位與標籤（可自由增減/改名）
  RECORD_LATEST_FIELDS: [
    { label: '日期',       key: 'date',  type: 'text' },
    { label: '總資產',     key: 'C',     type: 'number' },
    { label: '加上爸爸',   key: 'D',     type: 'number' },
    { label: '可動用',     key: 'E',     type: 'number' },
    { label: '當日損益',   key: 'F',     type: 'pnl' },
    { label: '現金',       key: 'G',     type: 'number' },
    { label: '債務',       key: 'BA',    type: 'number' }, // 可選：dashboard 已移除
    { label: '持有 USD',   key: 'CU',    type: 'number' }
  ]
}; // end APP_CONFIG
