// 對 demo_cases 加入抖動，讓金額/比率看起來自然；只接受「命中集合完全不變」的抖動
const fs = require("fs");
const path = require("path");
const cases = JSON.parse(fs.readFileSync(path.join(__dirname, "demo_cases.json"), "utf8"));

const r2 = (x) => Math.round(x * 100) / 100;
function materialize(c) {
  const H = r2(c.H), I = r2(c.J * H), K = r2(c.L * I);
  return { H, I, J: I / H, K, L: I > 0 ? K / I : 0, M: r2(I - K), O: r2(c.OI * I), P: c.P };
}
function calcVals(v) {
  return { ...v, diff: v.H > 0 ? Math.abs(v.K / v.H - v.J) : null, drawn: v.I - v.O, OI: v.I > 0 ? v.O / v.I : null };
}
const TESTS = {
  "1-1": (v) => v.K === 0, "1-2": (v) => v.L > 0 && v.L < 0.5,
  "1-3": (v) => v.diff >= 0.4, "1-4": (v) => v.diff < 0.4 && v.M >= 3e7,
  "1-5": (v) => v.drawn > 0 && v.K / v.drawn < 0.1 && Math.abs(v.K - v.drawn) > 1e8,
  "1-6": (v) => v.J > 1, "1-7": (v) => v.OI >= 0.4, "1-8": (v) => v.OI < 0.4 && v.O >= 3e7,
  "1-9": (v) => v.P > 0, "1-10": (v) => v.J >= 0.9 && v.L < 0.5,
  "2-1": (v) => v.K === 0 && v.I >= 3e6, "2-2": (v) => v.L > 0 && v.L < 0.5 && v.M >= 3e7,
  "2-3": (v) => v.diff >= 0.5 && v.M >= 3e6, "2-4": (v) => v.diff < 0.5 && v.M >= 3e7,
  "2-5": (v) => v.drawn > 0 && v.K / v.drawn < 0.1 && Math.abs(v.K - v.drawn) > 1e7,
  "2-6": (v) => v.J > 1, "2-7": (v) => v.OI >= 0.5 && v.O >= 3e6,
  "2-8": (v) => v.OI < 0.5 && v.O >= 1e7, "2-9": (v) => v.P > 0,
  "2-10": (v) => v.J >= 0.9 && v.L < 0.5,
};
const CPS = Object.keys(TESTS);
function hitsOf(c) {
  const v = calcVals(materialize(c));
  return CPS.filter((id) => id.startsWith("1-") === (c.cat === "軍投") && TESTS[id](v)).join(",");
}

const jit = (x, pct) => x * (1 + (Math.random() * 2 - 1) * pct);
let nOk = 0;
for (const c of cases) {
  const base = hitsOf(c);
  for (let t = 0; t < 400; t++) {
    const cand = {
      ...c,
      H: Math.round(jit(c.H, 0.35) / 100) * 100 + Math.round(Math.random() * 99), // 打散並帶零頭
      J: c.J === 0 ? 0 : jit(c.J, 0.04),
      L: c.L === 0 ? 0 : jit(c.L, 0.05), // L=0 必須維持 0（1-1/2-1）
      OI: jit(c.OI, 0.06),
      P: c.P > 0 ? Math.round(jit(c.P, 0.6)) : 0,
    };
    if (cand.J <= 0 || cand.L < 0 || cand.OI < 0 || cand.OI > 0.98) continue;
    if (hitsOf(cand) === base) { Object.assign(c, cand);nOk++; break; }
  }
}
console.log(`成功抖動 ${nOk}/${cases.length} 筆（命中集合均未改變）`);

// 抖動後統計
const hs = cases.map(hitsOf);
const uniqH = new Set(cases.map((c) => r2(c.H))).size;
console.log("不同發價值數量:", uniqH, "（抖動前僅 11 種）");
fs.writeFileSync(path.join(__dirname, "demo_cases.json"), JSON.stringify(cases, null, 1));
console.log("已更新 demo_cases.json");
