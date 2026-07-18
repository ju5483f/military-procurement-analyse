// 獨立重算演示檔的逐案風險評分（與 HTML 各自實作，用於交叉驗證）
// 六個定案錨點：金額 1萬~100億、待平衡 1萬~2億、結匯率上限300%、w=0.5
const fs = require('fs');
const path = require('path');
const XLSX = require('./xlsx.full.min.js');

const XLSX_FILE = process.argv[2] ||
  'C:/Users/gyzhu/Desktop/電腦小組/1150615研析試作區/執行中軍購案個案管制表(演示用).xlsx';

// ---- 讀檔 ----
const wb = XLSX.read(fs.readFileSync(XLSX_FILE), { type: 'buffer' });
const ws = wb.Sheets[wb.SheetNames[0]];
const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
const header = rows[0];
const dataRows = rows.slice(1).filter(r => r.some(c => c !== '' && c !== null && c !== undefined));

function num(v) {
  if (v === null || v === undefined || v === '') return null;
  const n = (typeof v === 'number') ? v : parseFloat(String(v).replace(/,/g, ''));
  return isNaN(n) ? null : n;
}
function calc(row) {
  const H = num(row[7]), I = num(row[8]), J = num(row[9]), K = num(row[10]),
        L = num(row[11]), M = num(row[12]), O = num(row[14]), P = num(row[15]);
  return {
    H, I, J, K, L, M, O, Pv: P,
    diff: (H !== null && H > 0 && K !== null && J !== null) ? Math.abs(K / H - J) : null,
    drawn: (I !== null && O !== null) ? I - O : null,
    OI: (I !== null && I > 0 && O !== null) ? O / I : null,
    kr: (I !== null && O !== null && (I - O) > 0 && K !== null) ? K / (I - O) : null,
    gap: (I !== null && O !== null && K !== null) ? Math.abs(K - (I - O)) : null,
    jl: (J !== null && L !== null) ? J - L : null
  };
}
const isZW = row => String(row[5] == null ? '' : row[5]).trim() === '作維';

// ---- 20 控制點（預設門檻，與 v1150715 相同）----
const CP = {
  '1-1': (v) => v.K !== null && v.K === 0,
  '1-2': (v) => v.L !== null && v.L > 0 && v.L < 0.5,
  '1-3': (v) => v.diff !== null && v.diff >= 0.4,
  '1-4': (v) => v.diff !== null && v.diff < 0.4 && v.M !== null && v.M >= 3e7,
  '1-5': (v) => v.drawn !== null && v.drawn > 0 && v.K !== null && (v.K / v.drawn) < 0.1 && Math.abs(v.K - v.drawn) > 1e8,
  '1-6': (v) => v.J !== null && v.J > 1.0,
  '1-7': (v) => v.OI !== null && v.OI >= 0.4,
  '1-8': (v) => v.OI !== null && v.OI < 0.4 && v.O !== null && v.O >= 3e7,
  '1-9': (v) => v.Pv !== null && v.Pv > 0,
  '1-10': (v) => v.J !== null && v.J >= 0.9 && v.L !== null && v.L < 0.5,
  '2-1': (v) => v.K !== null && v.K === 0 && v.I !== null && v.I >= 3e6,
  '2-2': (v) => v.L !== null && v.L > 0 && v.L < 0.5 && v.M !== null && v.M >= 3e7,
  '2-3': (v) => v.diff !== null && v.diff >= 0.5 && v.M !== null && v.M >= 3e6,
  '2-4': (v) => v.diff !== null && v.diff < 0.5 && v.M !== null && v.M >= 3e7,
  '2-5': (v) => v.drawn !== null && v.drawn > 0 && v.K !== null && (v.K / v.drawn) < 0.1 && Math.abs(v.K - v.drawn) > 1e7,
  '2-6': (v) => v.J !== null && v.J > 1.0,
  '2-7': (v) => v.OI !== null && v.OI >= 0.5 && v.O !== null && v.O >= 3e6,
  '2-8': (v) => v.OI !== null && v.OI < 0.5 && v.O !== null && v.O >= 1e7,
  '2-9': (v) => v.Pv !== null && v.Pv > 0,
  '2-10': (v) => v.J !== null && v.J >= 0.9 && v.L !== null && v.L < 0.5
};
const JT_IDS = ['1-1','1-2','1-3','1-4','1-5','1-6','1-7','1-8','1-9','1-10'];
const ZW_IDS = ['2-1','2-2','2-3','2-4','2-5','2-6','2-7','2-8','2-9','2-10'];

// ---- 錨點與嚴重度 ----
const A = { moneyHi: 1e10, moneyLo: 1e4, pLo: 1e4, pHi: 2e8, jHi: 3.0, w: 0.5 };
const clip = x => Math.max(0, Math.min(1, x));
function ratioSev(x, lo, hi) { if (x == null) return 0; return clip((x - lo) / (hi - lo)); }
function moneySev(x, lo, hi) { if (x == null) return 0; return clip(Math.log(Math.max(x, lo) / lo) / Math.log(hi / lo)); }
// 各維度評分（回傳 0~10）；lo/hi 已代入定案錨點
const DIM = {
  '1-1': { risk: '執行', score: () => 10 },
  '1-2': { risk: '執行', score: v => 1 + 9 * ratioSev(v.L, 0.5, 0) },
  '1-3/1-4': { risk: '執行', score: v => 1 + 9 * (0.5 * ratioSev(v.diff, 0, 1) + 0.5 * moneySev(v.M, A.moneyLo, A.moneyHi)) },
  '1-5': { risk: '執行', score: v => 1 + 9 * (0.5 * ratioSev(v.kr, 0.1, 0) + 0.5 * moneySev(v.gap, A.moneyLo, A.moneyHi)) },
  '1-6': { risk: '財務', score: v => 1 + 9 * ratioSev(v.J, 1.0, A.jHi) },
  '1-7/1-8': { risk: '財務', score: v => 1 + 9 * (0.5 * ratioSev(v.OI, 0, 1) + 0.5 * moneySev(v.O, A.moneyLo, A.moneyHi)) },
  '1-9': { risk: '財務', score: v => 1 + 9 * moneySev(v.Pv, A.pLo, A.pHi) },
  '2-1': { risk: '執行', score: () => 10 },
  '2-2': { risk: '執行', score: v => 1 + 9 * (0.5 * ratioSev(v.L, 0.5, 0) + 0.5 * moneySev(v.M, A.moneyLo, A.moneyHi)) },
  '2-3/2-4': { risk: '執行', score: v => 1 + 9 * (0.5 * ratioSev(v.diff, 0, 1) + 0.5 * moneySev(v.M, A.moneyLo, A.moneyHi)) },
  '2-5': { risk: '執行', score: v => 1 + 9 * (0.5 * ratioSev(v.kr, 0.1, 0) + 0.5 * moneySev(v.gap, A.moneyLo, A.moneyHi)) },
  '2-6': { risk: '財務', score: v => 1 + 9 * ratioSev(v.J, 1.0, A.jHi) },
  '2-7/2-8': { risk: '財務', score: v => 1 + 9 * (0.5 * ratioSev(v.OI, 0, 1) + 0.5 * moneySev(v.O, A.moneyLo, A.moneyHi)) },
  '2-9': { risk: '財務', score: v => 1 + 9 * moneySev(v.Pv, A.pLo, A.pHi) },
  '2-10': { risk: '執行', score: v => 1 + 9 * ratioSev(v.jl, 0.4, 1.0) }
};
function dimIdForCP(id) {
  if (id === '1-3' || id === '1-4') return '1-3/1-4';
  if (id === '1-7' || id === '1-8') return '1-7/1-8';
  if (id === '2-3' || id === '2-4') return '2-3/2-4';
  if (id === '2-7' || id === '2-8') return '2-7/2-8';
  if (id === '1-10') return null;
  return id;
}

// ---- 逐案評分 ----
const cases = dataRows.map((row, idx) => {
  const v = calc(row);
  const zw = isZW(row);
  const ids = (zw ? ZW_IDS : JT_IDS).filter(id => CP[id](v));
  const perDim = {};
  let exec = 0, fin = 0;
  for (const id of ids) {
    const did = dimIdForCP(id);
    if (did === null || perDim[did] !== undefined) continue;
    const sc = DIM[did].score(v);
    perDim[did] = sc;
    if (DIM[did].risk === '執行') exec += sc; else fin += sc;
  }
  return { idx, caseNo: row[1], name: row[4], svc: row[2], zw,
    exec: +exec.toFixed(2), fin: +fin.toFixed(2), total: +(exec + fin).toFixed(2),
    H: num(row[7]), perDim };
});

const scored = cases.filter(c => c.total > 0);
const unscored = cases.length - scored.length;
console.log('總案數:', cases.length, '｜有分數:', scored.length, '｜未命中(0分):', unscored);
console.log('全部總分加總:', cases.reduce((a, c) => a + c.total, 0).toFixed(2));

function top(zw, n) {
  return cases.filter(c => c.zw === zw && c.total > 0).sort((a, b) => b.total - a.total).slice(0, n);
}
console.log('\n=== 軍投 Top 8 ===');
for (const c of top(false, 8)) console.log(`${c.caseNo}｜執${c.exec} 財${c.fin} 總${c.total}｜發價值${c.H}｜${Object.entries(c.perDim).map(([k,s])=>k+':'+s.toFixed(2)).join(' ')}`);
console.log('\n=== 作維 Top 8 ===');
for (const c of top(true, 8)) console.log(`${c.caseNo}｜執${c.exec} 財${c.fin} 總${c.total}｜發價值${c.H}｜${Object.entries(c.perDim).map(([k,s])=>k+':'+s.toFixed(2)).join(' ')}`);

// 輸出逐案預期（與瀏覽器 caseScores 逐 index 比對）
const expected = cases.map(c => ({ i: c.idx, e: c.exec, f: c.fin, t: c.total }));
fs.writeFileSync(path.join(__dirname, 'expected_scores.json'), JSON.stringify(expected));
console.log('\n已寫出 expected_scores.json（', expected.length, '案）');
