/* 錨點校準工具 v1150717 的預期答案（Node 獨立實作，不得與 HTML 共用程式碼）
   以 執行中軍購案個案管制表(演示用).xlsx（451 筆）算出 A / B / C1 / C2 */
'use strict';
const fs = require('fs'), path = require('path'), os = require('os');
const XLSX = require('./xlsx.full.min.js');

const PROJ = 'C:/Users/gyzhu/Desktop/電腦小組/1150615研析試作區';
const SRC = path.join(PROJ, '執行中軍購案個案管制表(演示用).xlsx');

/* ---------- 讀檔 ---------- */
const tmp = path.join(os.tmpdir(), 'calib' + Date.now() + '.xlsx');
fs.copyFileSync(SRC, tmp);
const wb = XLSX.read(fs.readFileSync(tmp), { type: 'buffer' });
const raw = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, defval: '' });
const rows = raw.slice(1).filter(r => r.some(c => c !== '' && c !== null && c !== undefined));

function num(v) {
  if (v === null || v === undefined || v === '') return null;
  const n = (typeof v === 'number') ? v : parseFloat(String(v).replace(/,/g, ''));
  return isNaN(n) ? null : n;
}
function calcVals(row) {
  const H = num(row[7]), I = num(row[8]), J = num(row[9]), K = num(row[10]),
        L = num(row[11]), M = num(row[12]), O = num(row[14]), P = num(row[15]);
  return {
    H, I, J, K, L, M, O, P,
    diff:  (H !== null && H > 0 && K !== null && J !== null) ? Math.abs(K / H - J) : null,
    drawn: (I !== null && O !== null) ? I - O : null,
    OI:    (I !== null && I > 0 && O !== null) ? O / I : null,
    kr:    (I !== null && O !== null && (I - O) > 0 && K !== null) ? K / (I - O) : null,
    gap:   (I !== null && O !== null && K !== null) ? Math.abs(K - (I - O)) : null,
    jl:    (J !== null && L !== null) ? J - L : null
  };
}
const CASES = rows.map(r => ({
  no: String(r[1]).trim(),
  jt: String(r[5]).trim() !== '作維',      // 軍投＋特別預算 → 軍投
  v: calcVals(r)
}));

/* ---------- 預設門檻（同 v1150715 的 PARAM_DEFS） ---------- */
const T = {
  jt2_r: 0.5, jt34_d: 0.4, jt4_a: 3e7, jt5_r: 0.1, jt5_a: 1e8, jt6_r: 1.0,
  jt78_f: 0.4, jt8_a: 3e7, jt10_j: 0.9, jt10_l: 0.5,
  zw1_a: 3e6, zw2_r: 0.5, zw2_a: 3e7, zw34_d: 0.5, zw3_a: 3e6, zw4_a: 3e7,
  zw5_r: 0.1, zw5_a: 1e7, zw6_r: 1.0, zw78_f: 0.5, zw7_a: 3e6, zw8_a: 1e7,
  zw10_j: 0.9, zw10_l: 0.5
};

/* ---------- 16 個評分維度 ---------- */
const R = (get, lo, hi) => ({ kind: 'ratio', get, lo: () => lo, hi: () => hi });
const Mo = (get, thr) => ({ kind: 'money', get, thr });
const nn = (...xs) => xs.every(x => x !== null && x !== undefined);

const DIMS = [
  { id: '1-1', jt: 1, fixed: 10, hit: v => nn(v.K) && v.K === 0 },
  { id: '1-2', jt: 1, axes: [R(v => v.L, T.jt2_r, 0)], hit: v => nn(v.L) && v.L > 0 && v.L < T.jt2_r },
  { id: '1-3/1-4', jt: 1, axes: [R(v => v.diff, 0, 1), Mo(v => v.M, T.jt4_a)],
    hit: v => (nn(v.diff) && v.diff >= T.jt34_d) || (nn(v.diff, v.M) && v.diff < T.jt34_d && v.M >= T.jt4_a) },
  { id: '1-5', jt: 1, axes: [R(v => v.kr, T.jt5_r, 0), Mo(v => v.gap, T.jt5_a)],
    hit: v => nn(v.drawn, v.K, v.kr, v.gap) && v.drawn > 0 && v.kr < T.jt5_r && v.gap > T.jt5_a },
  { id: '1-6', jt: 1, axes: [R(v => v.J, 1.0, null)], hit: v => nn(v.J) && v.J > T.jt6_r },
  { id: '1-7/1-8', jt: 1, axes: [R(v => v.OI, 0, 1), Mo(v => v.O, T.jt8_a)],
    hit: v => (nn(v.OI) && v.OI >= T.jt78_f) || (nn(v.OI, v.O) && v.OI < T.jt78_f && v.O >= T.jt8_a) },
  { id: '1-9', jt: 1, axes: [Mo(v => v.P, null)], hit: v => nn(v.P) && v.P > 0 },
  { id: '1-10', jt: 1, axes: [R(v => v.jl, 0.4, 1.0)],
    hit: v => nn(v.J, v.L) && v.J >= T.jt10_j && v.L < T.jt10_l },

  { id: '2-1', jt: 0, fixed: 10, hit: v => nn(v.K, v.I) && v.K === 0 && v.I >= T.zw1_a },
  { id: '2-2', jt: 0, axes: [R(v => v.L, T.zw2_r, 0), Mo(v => v.M, T.zw2_a)],
    hit: v => nn(v.L, v.M) && v.L > 0 && v.L < T.zw2_r && v.M >= T.zw2_a },
  { id: '2-3/2-4', jt: 0, axes: [R(v => v.diff, 0, 1), Mo(v => v.M, T.zw3_a)],
    hit: v => nn(v.diff, v.M) && ((v.diff >= T.zw34_d && v.M >= T.zw3_a) || (v.diff < T.zw34_d && v.M >= T.zw4_a)) },
  { id: '2-5', jt: 0, axes: [R(v => v.kr, T.zw5_r, 0), Mo(v => v.gap, T.zw5_a)],
    hit: v => nn(v.drawn, v.K, v.kr, v.gap) && v.drawn > 0 && v.kr < T.zw5_r && v.gap > T.zw5_a },
  { id: '2-6', jt: 0, axes: [R(v => v.J, 1.0, null)], hit: v => nn(v.J) && v.J > T.zw6_r },
  { id: '2-7/2-8', jt: 0, axes: [R(v => v.OI, 0, 1), Mo(v => v.O, T.zw7_a)],
    hit: v => nn(v.OI, v.O) && ((v.OI >= T.zw78_f && v.O >= T.zw7_a) || (v.OI < T.zw78_f && v.O >= T.zw8_a)) },
  { id: '2-9', jt: 0, axes: [Mo(v => v.P, null)], hit: v => nn(v.P) && v.P > 0 },
  { id: '2-10', jt: 0, axes: [R(v => v.jl, 0.4, 1.0)],
    hit: v => nn(v.J, v.L) && v.J >= T.zw10_j && v.L < T.zw10_l }
];
const isP = (dim, ax) => ax.kind === 'money' && ax.thr === null;   // 待平衡支付值軸

/* ---------- 嚴重度換算 ---------- */
const clip = x => Math.max(0, Math.min(1, x));
function axLo(dim, ax, A) {
  if (ax.kind === 'ratio') return ax.lo();
  if (isP(dim, ax)) return A.pLo;
  return (A.moneyLo === null) ? ax.thr : A.moneyLo;
}
function axHi(dim, ax, A) {
  if (ax.kind === 'ratio') return (ax.hi() === null) ? A.jHi : ax.hi();
  if (isP(dim, ax)) return A.pHi;
  return A.moneyHi;
}
// 回傳 {s, pinLo, pinHi}；x 為 null 時視為 s=0（不貼頂不貼底）
function sev(dim, ax, v, A) {
  const x = ax.get(v);
  if (x === null || x === undefined) return { s: 0, pinLo: false, pinHi: false };
  const lo = axLo(dim, ax, A), hi = axHi(dim, ax, A);
  if (ax.kind === 'ratio') {
    return { s: clip((x - lo) / (hi - lo)), pinLo: (hi > lo ? x <= lo : x >= lo), pinHi: (hi > lo ? x >= hi : x <= hi) };
  }
  return { s: clip(Math.log(Math.max(x, lo) / lo) / Math.log(hi / lo)), pinLo: x <= lo, pinHi: x >= hi };
}
function scoreOf(dim, v, A) {
  if (dim.fixed !== undefined) return dim.fixed;
  const ss = dim.axes.map(ax => sev(dim, ax, v, A));
  let s;
  if (dim.axes.length === 1) s = ss[0].s;
  else {
    const ri = dim.axes.findIndex(a => a.kind === 'ratio');
    const mi = dim.axes.findIndex(a => a.kind === 'money');
    s = A.w * ss[ri].s + (1 - A.w) * ss[mi].s;
  }
  return 1 + 9 * s;
}

/* ---------- 計算 ---------- */
const DEF = { moneyHi: 1e9, moneyLo: null, pLo: 1e5, pHi: 1e7, jHi: 1.2, w: 0.5 };
function run(A) {
  const out = { A: {}, pinHi: {}, pinLo: {}, C1: {}, totals: {} };
  for (const c of CASES) out.totals[c.no] = 0;
  for (const dim of DIMS) {
    const pool = CASES.filter(c => (c.jt ? 1 : 0) === dim.jt && dim.hit(c.v));
    out.A[dim.id] = pool.length;
    if (!pool.length) continue;
    const scores = pool.map(c => scoreOf(dim, c.v, A));
    pool.forEach((c, i) => { out.totals[c.no] += scores[i]; });
    out.C1[dim.id] = {
      n: pool.length,
      min: Math.min(...scores), max: Math.max(...scores),
      avg: scores.reduce((a, b) => a + b, 0) / scores.length
    };
    if (dim.axes) dim.axes.forEach(ax => {
      const key = dim.id + '｜' + (ax.kind === 'ratio' ? '比率' : '金額');
      const r = pool.map(c => sev(dim, ax, c.v, A));
      out.pinHi[key] = r.filter(x => x.pinHi).length;
      out.pinLo[key] = r.filter(x => x.pinLo).length;
    });
  }
  return out;
}
// 跨軸去重的貼頂案件數 —— 只計「受金額上錨管轄」的軸。
// 待平衡支付值軸（thr===null）受 P 上錨管轄、不隨金額上錨變動，計入會造成恆定偏移。
function pinHiCases(A) {
  const set = new Set();
  for (const dim of DIMS) {
    if (!dim.axes) continue;
    for (const c of CASES.filter(c => (c.jt ? 1 : 0) === dim.jt && dim.hit(c.v)))
      for (const ax of dim.axes)
        if (ax.kind === 'money' && ax.thr !== null && sev(dim, ax, c.v, A).pinHi) set.add(c.no);
  }
  return set.size;
}

/* ---------- 輸出 ---------- */
const f2 = x => x.toFixed(2);
const r = run(DEF);
console.log('=== A 各評分維度命中案件數（預設參數）===');
for (const d of DIMS) console.log('  ' + d.id.padEnd(9) + String(r.A[d.id]).padStart(4));
console.log('  ' + '有分數案件數'.padEnd(9) + String(Object.values(r.totals).filter(x => x > 0).length).padStart(4) + '  ／ 共 ' + CASES.length + ' 案');

console.log('\n=== C1 各維度分數統計（預設參數）===');
console.log('  維度        案數    最低    最高    平均');
for (const d of DIMS) { const c = r.C1[d.id]; if (!c) continue;
  console.log('  ' + d.id.padEnd(10) + String(c.n).padStart(4) + '  ' + f2(c.min).padStart(6) + '  ' + f2(c.max).padStart(6) + '  ' + f2(c.avg).padStart(6)); }

console.log('\n=== B 貼底案件數（預設下錨＝控制點門檻）===');
for (const k of Object.keys(r.pinLo)) if (r.pinLo[k]) console.log('  ' + k.padEnd(18) + String(r.pinLo[k]).padStart(4));
console.log('\n=== B 貼頂案件數（金額上錨 10 億）===');
for (const k of Object.keys(r.pinHi)) if (r.pinHi[k]) console.log('  ' + k.padEnd(18) + String(r.pinHi[k]).padStart(4));

console.log('\n=== B 跨軸去重的貼頂案件數 vs 金額上錨 ===');
for (const HI of [5e7, 1e8, 3e8, 1e9, 5e9, 5e10])
  console.log('  上錨 ' + (HI >= 1e8 ? (HI / 1e8) + '億' : (HI / 1e4) + '萬').padEnd(8) + String(pinHiCases(Object.assign({}, DEF, { moneyHi: HI }))).padStart(4) + ' 案');

console.log('\n=== C2 總分分布（預設參數）===');
const tv = Object.values(r.totals);
const bins = [[0, 0]]; for (let i = 1; i <= 70; i += 10) bins.push([i, i + 9]);
for (const [lo, hi] of bins) {
  const n = tv.filter(x => (lo === 0 ? x === 0 : x >= lo && x < hi + 1)).length;
  console.log('  ' + (lo === 0 ? '0 分' : lo + '–' + hi + ' 分').padEnd(10) + String(n).padStart(4) + ' 案');
}
console.log('  總分最高 = ' + f2(Math.max(...tv)));

console.log('\n=== 破綻 2 佐證：1-7/1-8 下錨改 10 萬 ===');
const r2 = run(Object.assign({}, DEF, { moneyLo: 1e5 }));
const a = r.C1['1-7/1-8'], b = r2.C1['1-7/1-8'];
console.log('  下錨＝門檻(3000萬)：分數 ' + f2(a.min) + ' ~ ' + f2(a.max) + '，貼底 ' + r.pinLo['1-7/1-8｜金額'] + ' 案');
console.log('  下錨＝10萬        ：分數 ' + f2(b.min) + ' ~ ' + f2(b.max) + '，貼底 ' + (r2.pinLo['1-7/1-8｜金額'] || 0) + ' 案');
