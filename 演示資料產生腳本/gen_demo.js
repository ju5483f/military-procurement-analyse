// 產生演示用「執行中軍購案個案管制表」451 筆
// 目標：以 v1150715 分析後，結果與「分析結果0716.xlsx」完全相同
const fs = require("fs");
const path = require("path");
const os = require("os");
const XLSX = require("./xlsx.full.min.js");

const PROJ = "C:/Users/gyzhu/Desktop/電腦小組/1150615研析試作區";

/* ===== 目標矩陣（來自 分析結果0716.xlsx） ===== */
const TARGET = {
  "1-1":  { 陸軍: 1, 海軍: 0, 空軍: 3, 其他: 1, s2: 0 },
  "1-2":  { 陸軍: 7, 海軍: 5, 空軍: 10, 其他: 4, s2: 5 },
  "1-3":  { 陸軍: 4, 海軍: 3, 空軍: 2, 其他: 5, s2: 4 },
  "1-4":  { 陸軍: 4, 海軍: 2, 空軍: 10, 其他: 0, s2: 4 },
  "1-5":  { 陸軍: 0, 海軍: 0, 空軍: 1, 其他: 0, s2: 0 },
  "1-6":  { 陸軍: 0, 海軍: 0, 空軍: 0, 其他: 0, s2: 0 },
  "1-7":  { 陸軍: 5, 海軍: 4, 空軍: 9, 其他: 5, s2: 3 },
  "1-8":  { 陸軍: 3, 海軍: 1, 空軍: 5, 其他: 0, s2: 4 },
  "1-9":  { 陸軍: 0, 海軍: 0, 空軍: 0, 其他: 1, s2: 0 },
  "1-10": { 陸軍: 1, 海軍: 1, 空軍: 0, 其他: 1, s2: 2 },
  "2-1":  { 陸軍: 8, 海軍: 0, 空軍: 0, 其他: 0, s2: 5 },
  "2-2":  { 陸軍: 1, 海軍: 1, 空軍: 2, 其他: 0, s2: 1 },
  "2-3":  { 陸軍: 7, 海軍: 2, 空軍: 2, 其他: 2, s2: 5 },
  "2-4":  { 陸軍: 0, 海軍: 0, 空軍: 5, 其他: 1, s2: 0 },
  "2-5":  { 陸軍: 1, 海軍: 1, 空軍: 5, 其他: 0, s2: 1 },
  "2-6":  { 陸軍: 10, 海軍: 6, 空軍: 7, 其他: 4, s2: 1 },
  "2-7":  { 陸軍: 13, 海軍: 1, 空軍: 21, 其他: 1, s2: 5 },
  "2-8":  { 陸軍: 5, 海軍: 4, 空軍: 9, 其他: 0, s2: 3 },
  "2-9":  { 陸軍: 2, 海軍: 2, 空軍: 19, 其他: 1, s2: 2 },
  "2-10": { 陸軍: 9, 海軍: 4, 空軍: 2, 其他: 6, s2: 10 },
};
const CPS = Object.keys(TARGET);
const SVCS = ["陸軍", "海軍", "空軍", "其他"];

/* ===== 案件分布（總計 451 筆） ===== */
const GROUPS = [
  { svc: "陸軍", cat: "軍投", n: 40 },
  { svc: "陸軍", cat: "作維", n: 110 },
  { svc: "海軍", cat: "軍投", n: 30 },
  { svc: "海軍", cat: "作維", n: 80 },
  { svc: "空軍", cat: "軍投", n: 40 },
  { svc: "空軍", cat: "作維", n: 100 },
  { svc: "其他", cat: "軍投", n: 16 },
  { svc: "其他", cat: "作維", n: 35 },
];

/* ===== 可調 dial 值域 ===== */
const D = {
  J: [0.15, 0.25, 0.35, 0.45, 0.55, 0.65, 0.75, 0.85, 0.92, 0.96, 0.99, 1.03, 1.12],
  L: [0, 0.03, 0.08, 0.15, 0.25, 0.35, 0.45, 0.48, 0.55, 0.65, 0.75, 0.88, 0.97],
  H: [1.5e6, 4e6, 8e6, 1.5e7, 3e7, 6e7, 1.2e8, 2.5e8, 5e8, 9e8, 1.6e9],
  OI: [0.01, 0.05, 0.12, 0.22, 0.32, 0.42, 0.48, 0.55, 0.65, 0.78, 0.9],
  P: [0, 0, 0, 250000, 1200000, 8000000],
};
const r2 = (x) => Math.round(x * 100) / 100;

/* 由 dial 算出實際欄位值（與寫入 xlsx 的值一致） */
function materialize(c) {
  const H = r2(c.H);
  const I = r2(c.J * H);
  const K = r2(c.L * I);
  const J = I / H;
  const L = I > 0 ? K / I : 0;
  const M = r2(I - K);
  const O = r2(c.OI * I);
  return { H, I, J, K, L, M, O, P: c.P };
}

/* ===== 控制點邏輯（完全比照 v1150715） ===== */
function calcVals(v) {
  const { H, I, J, K, L, M, O, P } = v;
  return {
    H, I, J, K, L, M, O, P,
    DA: H > 0 ? K / H : null,
    diff: H > 0 ? Math.abs(K / H - J) : null,
    drawn: I - O,
    OI: I > 0 ? O / I : null,
  };
}
const TESTS = {
  "1-1": (v) => v.K === 0,
  "1-2": (v) => v.L > 0 && v.L < 0.5,
  "1-3": (v) => v.diff !== null && v.diff >= 0.4,
  "1-4": (v) => v.diff !== null && v.diff < 0.4 && v.M >= 3e7,
  "1-5": (v) => v.drawn > 0 && v.K / v.drawn < 0.1 && Math.abs(v.K - v.drawn) > 1e8,
  "1-6": (v) => v.J > 1,
  "1-7": (v) => v.OI !== null && v.OI >= 0.4,
  "1-8": (v) => v.OI !== null && v.OI < 0.4 && v.O >= 3e7,
  "1-9": (v) => v.P > 0,
  "1-10": (v) => v.J >= 0.9 && v.L < 0.5,
  "2-1": (v) => v.K === 0 && v.I >= 3e6,
  "2-2": (v) => v.L > 0 && v.L < 0.5 && v.M >= 3e7,
  "2-3": (v) => v.diff !== null && v.diff >= 0.5 && v.M >= 3e6,
  "2-4": (v) => v.diff !== null && v.diff < 0.5 && v.M >= 3e7,
  "2-5": (v) => v.drawn > 0 && v.K / v.drawn < 0.1 && Math.abs(v.K - v.drawn) > 1e7,
  "2-6": (v) => v.J > 1,
  "2-7": (v) => v.OI !== null && v.OI >= 0.5 && v.O >= 3e6,
  "2-8": (v) => v.OI !== null && v.OI < 0.5 && v.O >= 1e7,
  "2-9": (v) => v.P > 0,
  "2-10": (v) => v.J >= 0.9 && v.L < 0.5,
};

/* 回傳該案件命中的控制點 id 陣列 */
function hitsOf(c) {
  const v = calcVals(materialize(c));
  const out = [];
  for (const id of CPS) {
    if (id.startsWith("1-") !== (c.cat === "軍投")) continue; // 作維案只套 2-x，軍投案只套 1-x
    if (TESTS[id](v)) out.push(id);
  }
  return out;
}

/* ===== 初始化 ===== */
const rnd = (a) => a[Math.floor(Math.random() * a.length)];
let cases = [];
for (const g of GROUPS) {
  for (let i = 0; i < g.n; i++) {
    cases.push({
      svc: g.svc, cat: g.cat,
      J: rnd(D.J), L: rnd(D.L), H: rnd(D.H), OI: rnd(D.OI), P: rnd(D.P),
      match: false,
    });
  }
}

/* ===== 計數矩陣（增量維護） ===== */
function blankMatrix() {
  const m = {};
  for (const id of CPS) m[id] = { 陸軍: 0, 海軍: 0, 空軍: 0, 其他: 0, s2: 0 };
  return m;
}
let MAT = blankMatrix();
const hitCache = new Array(cases.length);
function applyCase(i, sign) {
  const c = cases[i];
  const hs = hitCache[i];
  for (const id of hs) {
    MAT[id][c.svc] += sign;
    if (c.match) MAT[id].s2 += sign;
  }
}
for (let i = 0; i < cases.length; i++) { hitCache[i] = hitsOf(cases[i]); applyCase(i, +1); }

function errorOf() {
  let e = 0;
  for (const id of CPS) {
    for (const s of SVCS) e += Math.abs(MAT[id][s] - TARGET[id][s]);
    e += Math.abs(MAT[id].s2 - TARGET[id].s2) * 2; // 第二階段加權，較難滿足
  }
  return e;
}

/* ===== 退火搜尋 ===== */
let err = errorOf();
console.log("初始誤差:", err);
const MAXIT = 4_000_000;
let temp = 3.0;
for (let it = 0; it < MAXIT && err > 0; it++) {
  if (it % 200000 === 0) temp = Math.max(0.02, 3.0 * (1 - it / MAXIT));
  const i = Math.floor(Math.random() * cases.length);
  const c = cases[i];
  const old = { J: c.J, L: c.L, H: c.H, OI: c.OI, P: c.P, match: c.match };

  applyCase(i, -1);
  const dial = Math.floor(Math.random() * 6);
  if (dial === 0) c.J = rnd(D.J);
  else if (dial === 1) c.L = rnd(D.L);
  else if (dial === 2) c.H = rnd(D.H);
  else if (dial === 3) c.OI = rnd(D.OI);
  else if (dial === 4) c.P = rnd(D.P);
  else if (c.svc === "陸軍" || c.svc === "海軍") c.match = !c.match; // 只有陸/海能命中會議檔
  hitCache[i] = hitsOf(c);
  applyCase(i, +1);

  const ne = errorOf();
  const d = ne - err;
  if (d <= 0 || Math.random() < Math.exp(-d / temp)) {
    err = ne;
  } else {
    applyCase(i, -1);
    Object.assign(c, old);
    hitCache[i] = hitsOf(c);
    applyCase(i, +1);
  }
}
console.log("最終誤差:", err);
if (err !== 0) {
  console.log("\n未收斂，差異明細：");
  for (const id of CPS) {
    const a = MAT[id], t = TARGET[id];
    const diffs = SVCS.filter((s) => a[s] !== t[s]).map((s) => `${s} ${a[s]}≠${t[s]}`);
    if (a.s2 !== t.s2) diffs.push(`第二階段 ${a.s2}≠${t.s2}`);
    if (diffs.length) console.log(`  ${id}: ${diffs.join(", ")}`);
  }
  process.exit(1);
}

/* ===== 收斂後：統計命中案件數，供案號指派 ===== */
const needMatch = { 陸軍: 0, 海軍: 0 };
for (const c of cases) if (c.match) needMatch[c.svc]++;
console.log("需命中會議檔的案件數:", JSON.stringify(needMatch));
fs.writeFileSync(path.join(__dirname, "demo_cases.json"), JSON.stringify(cases, null, 1));
console.log("已寫出 demo_cases.json（", cases.length, "筆 ）");
