// 獨立重算演示檔的風險分布圖幾何（與 HTML 各自實作，用於交叉驗證）
// 重用 expected_score.js 的評分核心，再加繪圖座標／半徑／分區
const fs = require('fs');
const path = require('path');
const XLSX = require('./xlsx.full.min.js');

const XLSX_FILE = process.argv[2] ||
  'C:/Users/gyzhu/Desktop/電腦小組/1150615研析試作區/執行中軍購案個案管制表(演示用).xlsx';

const wb = XLSX.read(fs.readFileSync(XLSX_FILE), { type: 'buffer' });
const rows0 = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, defval: '' });
const dataRows = rows0.slice(1).filter(r => r.some(c => c !== '' && c !== null && c !== undefined));

function num(v) {
  if (v === null || v === undefined || v === '') return null;
  const n = (typeof v === 'number') ? v : parseFloat(String(v).replace(/,/g, ''));
  return isNaN(n) ? null : n;
}
function calc(row) {
  const H = num(row[7]), I = num(row[8]), J = num(row[9]), K = num(row[10]),
        L = num(row[11]), M = num(row[12]), O = num(row[14]), Pp = num(row[15]);
  return { H, I, J, K, L, M, O, Pv: Pp,
    diff: (H !== null && H > 0 && K !== null && J !== null) ? Math.abs(K / H - J) : null,
    drawn: (I !== null && O !== null) ? I - O : null,
    OI: (I !== null && I > 0 && O !== null) ? O / I : null,
    kr: (I !== null && O !== null && (I - O) > 0 && K !== null) ? K / (I - O) : null,
    gap: (I !== null && O !== null && K !== null) ? Math.abs(K - (I - O)) : null,
    jl: (J !== null && L !== null) ? J - L : null };
}
const isZW = row => String(row[5] == null ? '' : row[5]).trim() === '作維';

const CP = {
  '1-1': v => v.K !== null && v.K === 0,
  '1-2': v => v.L !== null && v.L > 0 && v.L < 0.5,
  '1-3': v => v.diff !== null && v.diff >= 0.4,
  '1-4': v => v.diff !== null && v.diff < 0.4 && v.M !== null && v.M >= 3e7,
  '1-5': v => v.drawn !== null && v.drawn > 0 && v.K !== null && (v.K / v.drawn) < 0.1 && Math.abs(v.K - v.drawn) > 1e8,
  '1-6': v => v.J !== null && v.J > 1.0,
  '1-7': v => v.OI !== null && v.OI >= 0.4,
  '1-8': v => v.OI !== null && v.OI < 0.4 && v.O !== null && v.O >= 3e7,
  '1-9': v => v.Pv !== null && v.Pv > 0,
  '1-10': v => v.J !== null && v.J >= 0.9 && v.L !== null && v.L < 0.5,
  '2-1': v => v.K !== null && v.K === 0 && v.I !== null && v.I >= 3e6,
  '2-2': v => v.L !== null && v.L > 0 && v.L < 0.5 && v.M !== null && v.M >= 3e7,
  '2-3': v => v.diff !== null && v.diff >= 0.5 && v.M !== null && v.M >= 3e6,
  '2-4': v => v.diff !== null && v.diff < 0.5 && v.M !== null && v.M >= 3e7,
  '2-5': v => v.drawn !== null && v.drawn > 0 && v.K !== null && (v.K / v.drawn) < 0.1 && Math.abs(v.K - v.drawn) > 1e7,
  '2-6': v => v.J !== null && v.J > 1.0,
  '2-7': v => v.OI !== null && v.OI >= 0.5 && v.O !== null && v.O >= 3e6,
  '2-8': v => v.OI !== null && v.OI < 0.5 && v.O !== null && v.O >= 1e7,
  '2-9': v => v.Pv !== null && v.Pv > 0,
  '2-10': v => v.J !== null && v.J >= 0.9 && v.L !== null && v.L < 0.5
};
const JT_IDS = ['1-1','1-2','1-3','1-4','1-5','1-6','1-7','1-8','1-9','1-10'];
const ZW_IDS = ['2-1','2-2','2-3','2-4','2-5','2-6','2-7','2-8','2-9','2-10'];
const A = { moneyHi: 1e10, moneyLo: 1e4, pLo: 1e4, pHi: 2e8, jHi: 3.0, w: 0.5 };
const clip = x => Math.max(0, Math.min(1, x));
const rS = (x, lo, hi) => x == null ? 0 : clip((x - lo) / (hi - lo));
const mS = (x, lo, hi) => x == null ? 0 : clip(Math.log(Math.max(x, lo) / lo) / Math.log(hi / lo));
const DIM = {
  '1-1': { risk: '執行', score: () => 10 },
  '1-2': { risk: '執行', score: v => 1 + 9 * rS(v.L, 0.5, 0) },
  '1-3/1-4': { risk: '執行', score: v => 1 + 9 * (0.5 * rS(v.diff, 0, 1) + 0.5 * mS(v.M, A.moneyLo, A.moneyHi)) },
  '1-5': { risk: '執行', score: v => 1 + 9 * (0.5 * rS(v.kr, 0.1, 0) + 0.5 * mS(v.gap, A.moneyLo, A.moneyHi)) },
  '1-6': { risk: '財務', score: v => 1 + 9 * rS(v.J, 1.0, A.jHi) },
  '1-7/1-8': { risk: '財務', score: v => 1 + 9 * (0.5 * rS(v.OI, 0, 1) + 0.5 * mS(v.O, A.moneyLo, A.moneyHi)) },
  '1-9': { risk: '財務', score: v => 1 + 9 * mS(v.Pv, A.pLo, A.pHi) },
  '2-1': { risk: '執行', score: () => 10 },
  '2-2': { risk: '執行', score: v => 1 + 9 * (0.5 * rS(v.L, 0.5, 0) + 0.5 * mS(v.M, A.moneyLo, A.moneyHi)) },
  '2-3/2-4': { risk: '執行', score: v => 1 + 9 * (0.5 * rS(v.diff, 0, 1) + 0.5 * mS(v.M, A.moneyLo, A.moneyHi)) },
  '2-5': { risk: '執行', score: v => 1 + 9 * (0.5 * rS(v.kr, 0.1, 0) + 0.5 * mS(v.gap, A.moneyLo, A.moneyHi)) },
  '2-6': { risk: '財務', score: v => 1 + 9 * rS(v.J, 1.0, A.jHi) },
  '2-7/2-8': { risk: '財務', score: v => 1 + 9 * (0.5 * rS(v.OI, 0, 1) + 0.5 * mS(v.O, A.moneyLo, A.moneyHi)) },
  '2-9': { risk: '財務', score: v => 1 + 9 * mS(v.Pv, A.pLo, A.pHi) },
  '2-10': { risk: '執行', score: v => 1 + 9 * rS(v.jl, 0.4, 1.0) }
};
const dimIdForCP = id => (id === '1-3' || id === '1-4') ? '1-3/1-4'
  : (id === '1-7' || id === '1-8') ? '1-7/1-8'
  : (id === '2-3' || id === '2-4') ? '2-3/2-4'
  : (id === '2-7' || id === '2-8') ? '2-7/2-8'
  : id === '1-10' ? null : id;

const cases = dataRows.map((row, idx) => {
  const v = calc(row), zw = isZW(row);
  const ids = (zw ? ZW_IDS : JT_IDS).filter(id => CP[id](v));
  const perDim = {}; let exec = 0, fin = 0;
  for (const id of ids) {
    const did = dimIdForCP(id);
    if (did === null || perDim[did] !== undefined) continue;
    const sc = DIM[did].score(v); perDim[did] = sc;
    if (DIM[did].risk === '執行') exec += sc; else fin += sc;
  }
  // execRaw/finRaw：未四捨五入，供繪圖座標比對（HTML 泡泡以原始累加分數定位，非 2dp）
  return { idx, caseNo: row[1], zw, exec: +exec.toFixed(2), fin: +fin.toFixed(2),
    execRaw: exec, finRaw: fin, total: +(exec + fin).toFixed(2), H: num(row[7]) };
});

// ---- 繪圖幾何（須與 HTML 的 CH/chartX/chartY/bubbleR/zoneOf 一致）----
const CH = { PL: 58, PR: 438, PT: 50, PB: 430 };
const chartX = e => CH.PL + (e / 40) * (CH.PR - CH.PL);
const chartY = f => CH.PB - (f / 40) * (CH.PB - CH.PT);
const bubbleR = H => (H === null || !(H > 0)) ? 4 : 4 + 22 * clip((Math.log(H) / Math.LN10 - 4) / 6);
const zoneOf = t => t >= 30 ? 3 : t >= 20 ? 2 : t >= 10 ? 1 : 0;

function report(zw, label) {
  const inCat = cases.filter(c => c.zw === zw);
  const scored = inCat.filter(c => c.total > 0).sort((a, b) => b.total - a.total);
  const zero = inCat.length - scored.length;
  const zoneCnt = [0, 0, 0, 0];
  scored.forEach(c => zoneCnt[zoneOf(c.total)]++);
  console.log(`\n=== ${label} ===`);
  console.log(`泡泡數(有分數) ${scored.length}｜0 分註記 ${zero}｜分區[綠/黃/橘/紅] ${zoneCnt.join('/')}`);
  console.log('前 10 名標籤:', scored.slice(0, 10).map(c => c.caseNo).join(', '));
  return { scored, zero };
}
const jt = report(false, '軍投');
const zw = report(true, '作維');

function spot(no) {
  const c = cases.find(x => x.caseNo === no);
  console.log(`${no}: (exec ${c.exec}, fin ${c.fin}) → cx ${chartX(c.execRaw).toFixed(1)}, cy ${chartY(c.finRaw).toFixed(1)}, r ${bubbleR(c.H).toFixed(1)}, zone ${zoneOf(c.total)}`);
}
console.log('\n=== 座標/半徑抽驗 ===');
spot('TWMYUY'); spot('TWBKQK'); spot('TWPJLD');

// 輸出全部繪入泡泡（供瀏覽器逐案比對）
const out = {};
[[false, 'jt'], [true, 'zw']].forEach(([z, k]) => {
  out[k] = cases.filter(c => c.zw === z && c.total > 0).map(c => ({
    i: c.idx, cx: +chartX(c.execRaw).toFixed(1), cy: +chartY(c.finRaw).toFixed(1), r: +bubbleR(c.H).toFixed(1)
  }));
});
fs.writeFileSync(path.join(__dirname, 'expected_chart.json'), JSON.stringify(out));
console.log('\n已寫出 expected_chart.json（軍投', out.jt.length, '／作維', out.zw.length, '泡泡）');
