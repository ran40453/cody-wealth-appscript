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

/** ================= FX graph from sheet pairs =================
// 若你想固定讀哪一張工作表，填上名稱（例如 'Rates'）。留空則讀目前作用中的工作表。 */
var FX_SHEET_NAME = '💱';


// 簡單快取，避免每儲存格重複讀表
var __FX_CACHE__ = {
  ts: 0,
  graph: null
};

/** 重新載入匯率圖的快取 */
function FX_RELOAD(){
  __FX_CACHE__.graph = null;
  __FX_CACHE__.ts = 0;
  return 'OK';
}

/** 列出目前解析到的邊（除錯用）：=FX_DUMP() */
function FX_DUMP(){
  var g = getFxGraphFromSheet_();
  var out = [['FROM','TO','RATE']];
  for (var a in g){
    for (var b in g[a]){
      out.push([a,b,g[a][b]]);
    }
  }
  return out;
}

/** 自訂函數：=FX(base, quote, [amount])
 *  回傳 base->quote 的匯率；若提供 amount，回傳換算後金額。
 *  例：=FX("USD","KRW") 或 =FX("USD","KRW", 25)
 */
function FX(base, quote, amount){
  base = String(base || '').trim().toUpperCase();
  quote = String(quote || '').trim().toUpperCase();
  if (!base || !quote) return 'ERR: empty base/quote';

  var graph = getFxGraphFromSheet_();
  if (!graph[base]) return 'ERR: unknown base ' + base;
  if (!graph[quote]) return 'ERR: unknown quote ' + quote;

  var rate = fxRateBetween_(graph, base, quote);
  if (rate == null) return 'ERR: no path ' + base + '->' + quote;
  if (amount == null || amount === '') return rate;
  var num = Number(amount);
  if (!isFinite(num)) return 'ERR: amount not number';
  return num * rate;
}

/** 讀取 J5:K13，建立無向圖（含倒數邊） */
function getFxGraphFromSheet_(){
  var now = Date.now();
  // 5 分鐘更新一次
  if (__FX_CACHE__.graph && (now - __FX_CACHE__.ts) < 5*60*1000) return __FX_CACHE__.graph;

  var ss = SpreadsheetApp.getActive();
  var sh = FX_SHEET_NAME ? ss.getSheetByName(FX_SHEET_NAME) : ss.getActiveSheet();
  if (!sh) throw new Error('找不到 FX_SHEET_NAME 指定的工作表：' + FX_SHEET_NAME);
  var rng = sh.getRange('J5:K13');
  var values = rng.getValues(); // [[label, rate], ...]

  var graph = Object.create(null);

  function ensure(n){ if(!graph[n]) graph[n] = Object.create(null); }

  for (var i=0;i<values.length;i++){
    var label = (values[i][0] + '').trim();
    var val = values[i][1];
    if (!label) continue;
    // 跳過空值與非數字
    var num = Number(val);
    if (!isFinite(num)) continue;

    // 正規化：
    // 支援 "AAA / BBB"、"AAA/BBB"、"AAA : BBB"、"1K AAA : BBB"
    var m1 = label.match(/^([A-Za-z0-9]+)\s*\/\s*([A-Za-z0-9]+)$/);
    var m2 = label.match(/^([A-Za-z0-9]+)\s*:\s*([A-Za-z0-9]+)$/);
    var m3 = label.match(/^1[Kk]\s+([A-Za-z]{3})\s*:\s*([A-Za-z]{3})$/);

    var from=null, to=null, scale=1;
    if (m3){
      from = m3[1].toUpperCase();
      to   = m3[2].toUpperCase();
      // 1000 from -> num to  =>  1 from -> num/1000 to
      num = num / 1000;
    } else if (m1){
      from = m1[1].toUpperCase();
      to   = m1[2].toUpperCase();
    } else if (m2){
      from = m2[1].toUpperCase();
      to   = m2[2].toUpperCase();
    } else {
      // 也允許像 "USD KRW" 或 "USD->KRW"
      var t = label.replace(/\s+/g,'').replace('->','/');
      var m = t.match(/^([A-Za-z]{3})\/([A-Za-z]{3})$/);
      if (m){ from=m[1].toUpperCase(); to=m[2].toUpperCase(); }
    }

    if (!from || !to) continue;

    ensure(from); ensure(to);
    // 建邊
    graph[from][to] = num;           // from -> to
    if (num !== 0) graph[to][from] = 1/num; // reciprocal
  }

  __FX_CACHE__.graph = graph;
  __FX_CACHE__.ts = now;
  return graph;
}

/** 在乘法圖上找 base->quote 匯率。使用 -log(w) 作為權重，跑 Dijkstra。 */
function fxRateBetween_(graph, base, quote){
  if (base === quote) return 1;
  var nodes = Object.keys(graph);
  var dist = Object.create(null);
  var prev = Object.create(null);
  var visited = Object.create(null);

  nodes.forEach(function(n){ dist[n] = Infinity; });
  dist[base] = 0;

  function popMin(){
    var best=null, bestD=Infinity;
    for (var n in dist){
      if (visited[n]) continue;
      if (dist[n] < bestD){ bestD = dist[n]; best = n; }
    }
    return best;
  }

  while(true){
    var u = popMin();
    if (!u) break;
    if (u === quote) break;
    visited[u] = true;

    var edges = graph[u] || {};
    for (var v in edges){
      var w = edges[v];
      if (!isFinite(w) || w <= 0) continue;
      var cost = -Math.log(w);
      var alt = dist[u] + cost;
      if (alt < dist[v]){
        dist[v] = alt;
        prev[v] = u;
      }
    }
  }

  if (!isFinite(dist[quote])) return null;
  // dist 是 -log(乘積)。轉回匯率：exp(-dist)
  return Math.exp(-dist[quote]);
}

/** ================= Yahoo Finance auto rates =================
 *  用 Yahoo Finance 抓最新匯率，不需要在 J5:K13 手填數字。
 *  提供：
 *    - =YF_RATE(base, quote)         // 回傳 1 base -> ? quote
 *    - =YF_RATES(base, range_codes)  // 回傳一整欄，range_codes 例如 J6:J13
 *  實作策略：優先嘗試直接 pair，如 USDKRW=X；若無，走 USD 交叉：rate(base->quote)= (USD->quote)/(USD->base)
 */

// 以函數呼叫期內的快取避免重複請求
var __YF_MEMO__ = {};

function YF_RATE(base, quote){
  base = String(base||'').trim().toUpperCase();
  quote = String(quote||'').trim().toUpperCase();
  if (!base || !quote) return 'ERR: empty base/quote';
  if (base === quote) return 1;
  try{
    // 先嘗試直接 pair
    var direct = yfFetchDirectPair_(base, quote);
    if (direct != null) return direct;

    // 再走 USD 交叉
    var usdToQuote = (quote === 'USD') ? 1 : yfUsdTo_(quote);
    var usdToBase  = (base  === 'USD') ? 1 : yfUsdTo_(base);
    if (!isFinite(usdToQuote) || !isFinite(usdToBase) || usdToBase === 0)
      return 'ERR: cross not available';
    return usdToQuote / usdToBase; // (quote/USD) / (base/USD) = quote/base
  }catch(e){
    return 'ERR: ' + e.message;
  }
}

/**
 * 自動填滿一整欄：=YF_RATES("TWD", J6:J13)
 * 會輸出與 range_codes 同高的一欄數字，可直接放在 K6（Google 會向下展開）。
 */
function YF_RATES(base, range_codes){
  base = String(base||'').trim().toUpperCase();
  var values = Array.isArray(range_codes) ? range_codes : [[range_codes]];
  var out = [];
  for (var r=0; r<values.length; r++){
    var code = (values[r] && values[r][0] != null) ? (''+values[r][0]).trim().toUpperCase() : '';
    if (!code) { out.push(['']); continue; }
    out.push([ YF_RATE(base, code) ]);
  }
  return out;
}

// ===== helpers =====
function yfFetchDirectPair_(base, quote){
  var sym1 = base + quote + '=X';
  var p1 = yfFetchPrice_(sym1);
  if (p1 != null) return p1;
  var sym2 = quote + base + '=X';
  var p2 = yfFetchPrice_(sym2);
  if (p2 != null && isFinite(p2) && p2 !== 0) return 1/p2; // 反向
  return null;
}

function yfUsdTo_(code){
  // 1 USD -> ? code
  if (__YF_MEMO__['USD_'+code] != null) return __YF_MEMO__['USD_'+code];
  var price = yfFetchPrice_('USD' + code + '=X');
  if (price == null){
    // 嘗試反向（codeUSD）
    var rev = yfFetchPrice_(code + 'USD=X');
    if (rev != null && isFinite(rev) && rev !== 0) price = 1/rev;
  }
  if (price == null) throw new Error('USD cross missing for ' + code);
  __YF_MEMO__['USD_'+code] = price;
  return price;
}

function yfFetchPrice_(symbol){
  if (__YF_MEMO__['sym_'+symbol] != null) return __YF_MEMO__['sym_'+symbol];
  var url = 'https://query1.finance.yahoo.com/v8/finance/chart/' + encodeURIComponent(symbol) + '?range=1d&interval=1d';
  var res = UrlFetchApp.fetch(url, {muteHttpExceptions: true});
  var text = res.getContentText();
  try{
    var json = JSON.parse(text);
    if (!json || !json.chart || !json.chart.result || !json.chart.result[0]) return null;
    var meta = json.chart.result[0].meta;
    var px = meta && meta.regularMarketPrice;
    if (px == null || !isFinite(px)){
      // 後備：抓 quotes
      var quotes = json.chart.result[0].indicators && json.chart.result[0].indicators.quote;
      if (quotes && quotes[0] && Array.isArray(quotes[0].close)){
        var arr = quotes[0].close.filter(function(n){ return isFinite(n); });
        px = arr.length ? arr[arr.length-1] : null;
      }
    }
    if (px == null) return null;
    __YF_MEMO__['sym_'+symbol] = px;
    return px;
  }catch(e){
    return null;
  }
}