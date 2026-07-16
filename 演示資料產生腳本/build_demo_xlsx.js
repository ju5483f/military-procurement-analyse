// 由 demo_cases.json 產生演示用管制表 xlsx（451 筆）
const fs = require("fs");
const path = require("path");
const os = require("os");
const XLSX = require("./xlsx.full.min.js");

const PROJ = "C:/Users/gyzhu/Desktop/電腦小組/1150615研析試作區";
const cases = JSON.parse(fs.readFileSync(path.join(__dirname, "demo_cases.json"), "utf8"));

/* ===== 讀會議檔案號（真實、公開） ===== */
function readMtg(file) {
  const tmp = path.join(os.tmpdir(), "mb_" + Math.random().toString(36).slice(2) + ".xlsx");
  fs.copyFileSync(path.join(PROJ, "會議資料", file), tmp);
  const wb = XLSX.read(fs.readFileSync(tmp), { type: "buffer" });
  const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, defval: "" });
  const map = new Map(); // 案號 → 案名
  for (const r of rows.slice(1)) {
    const no = String(r[1] || "").trim();
    const nm = String(r[2] || "").trim();
    if (no && no !== "無" && !map.has(no)) map.set(no, nm);
  }
  return map;
}
const mtgArmy = readMtg("陸軍114-4會議資料整理.xlsx");
const mtgNavy = readMtg("海軍114-4會議資料整理.xlsx");
const allMtgNos = new Set([...mtgArmy.keys(), ...mtgNavy.keys()]);
console.log(`會議檔案號：陸軍 ${mtgArmy.size} 個、海軍 ${mtgNavy.size} 個`);

/* ===== 控制點邏輯（比照 v1150715） ===== */
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
  return CPS.filter((id) => id.startsWith("1-") === (c.cat === "軍投") && TESTS[id](v));
}

/* ===== 後處理：無命中的案件不需要會議檔案號 ===== */
for (const c of cases) {
  c._hits = hitsOf(c);
  if (!c._hits.length) c.match = false;
}

/* ===== 名稱素材 ===== */
const UNITS = {
  陸軍: ["陸軍主計處", "陸軍後勤指揮部", "陸軍飛彈指揮部", "陸軍保修指揮部", "陸軍司令部"],
  海軍: ["海軍主計處", "海軍後勤指揮部", "海軍保修指揮部", "海軍艦隊指揮部", "海軍司令部"],
  空軍: ["空軍主計處", "空軍後勤指揮部", "空軍保修指揮部", "空軍作戰指揮部", "空軍司令部"],
  其他: ["國防部主計局", "國防部採購室", "資通電軍指揮部", "國防部參謀本部", "國防部軍備局"],
};
const SVC_LABEL = {
  陸軍: ["陸軍"], 海軍: ["海軍"], 空軍: ["空軍"],
  其他: ["中央", "中央、陸軍", "中央、海軍", "中央、空軍", "多軍種統建", "陸軍、空軍"],
};
const PROJ_WORDS = {
  陸軍: ["戰車", "裝甲車", "反裝甲飛彈", "榴彈砲", "通信裝備", "直升機零附件", "彈藥", "訓練器材", "夜視鏡", "無人機"],
  海軍: ["巡防艦", "魚雷", "艦砲", "聲納", "潛艦裝備", "艦艇零附件", "反艦飛彈", "水雷", "救難裝備", "航材"],
  空軍: ["戰機", "空對空飛彈", "雷達", "航電系統", "發動機", "航材零附件", "空用炸彈", "模擬機", "地面裝備", "電戰莢艙"],
  其他: ["聯合指管系統", "通資基礎設施", "資安防護", "衛星通信", "共用零附件", "聯合後勤", "電子戰系統", "情監偵裝備"],
};
const SUFFIX = ["採購案", "延壽案", "獲得案", "後續支援", "零附件開放式軍購案", "二號訂單軍購案", "回修案", "技術服務", "整補案", "籌購案"];
const CODE = ["150101", "150104", "150201", "150203", "150301", "160101", "160105", "170102", "170204", "180101"];

const rnd = (a) => a[Math.floor(Math.random() * a.length)];
const LET = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
const usedNo = new Set(allMtgNos);
function fakeNo() {
  for (;;) {
    const s = "TW" + rnd(["B", "P", "D", "M", "K"]) + LET[Math.floor(Math.random() * 26)] + LET[Math.floor(Math.random() * 26)] + LET[Math.floor(Math.random() * 26)];
    if (!usedNo.has(s)) { usedNo.add(s); return s; }
  }
}

/* ===== 指派案號／案名 ===== */
const poolArmy = [...mtgArmy.entries()];
const poolNavy = [...mtgNavy.entries()];
let ia = 0, inv = 0;
for (const c of cases) {
  if (c.match && c.svc === "陸軍") {
    const [no, nm] = poolArmy[ia++]; c.no = no; c.name = nm || "陸軍軍購案";
  } else if (c.match && c.svc === "海軍") {
    const [no, nm] = poolNavy[inv++]; c.no = no; c.name = nm || "海軍軍購案";
  } else {
    c.no = fakeNo();
    c.name = rnd(PROJ_WORDS[c.svc]) + rnd(SUFFIX);
  }
}
if (ia > poolArmy.length || inv > poolNavy.length) throw new Error("會議檔案號不足");
console.log(`指派會議檔案號：陸軍 ${ia} 筆、海軍 ${inv} 筆`);

/* ===== 類別欄：部分軍投案標為「特別預算」（仍套軍投控制點） ===== */
let spCount = 0;
for (const c of cases) {
  c.catLabel = c.cat;
  if (c.cat === "軍投" && Math.random() < 0.12) { c.catLabel = "特別預算"; spCount++; }
}
console.log(`類別欄標為「特別預算」者：${spCount} 筆（仍套軍投 10 項控制點）`);

/* ===== 組裝資料列 ===== */
const HEADERS = ["項次", "案號", "軍種", "單位", "案名", "類別", "預算科目",
  "發價值(A)\n美元", "累積結匯值(B)\n美元", "累計結匯率\n(C)=(B)/(A)", "累積交運值(D)\n美元",
  "累計交運率\n(E)=(D)/(B)", "軍購備查帳\n(F)=(B)-(D)\n美元", "軍購備查帳\n(F)=(B)-(D)\n新臺幣元",
  "FRB餘額", "待平衡支付值\n美元"];

// 項次：遞增且有跳號（比照真實樣態）
let seq = 0;
const aoa = [HEADERS];
for (const c of cases) {
  seq += 1 + Math.floor(Math.random() * 3);
  const v = materialize(c);
  aoa.push([
    seq, c.no, rnd(SVC_LABEL[c.svc]), rnd(UNITS[c.svc]), c.name, c.catLabel, rnd(CODE),
    v.H, v.I, v.J, v.K, v.L, v.M, r2(v.M * 30.85), v.O, v.P,
  ]);
}

const ws = XLSX.utils.aoa_to_sheet(aoa);
ws["!cols"] = [{ wch: 6 }, { wch: 10 }, { wch: 12 }, { wch: 18 }, { wch: 30 }, { wch: 10 }, { wch: 10 },
  { wch: 16 }, { wch: 16 }, { wch: 12 }, { wch: 16 }, { wch: 12 }, { wch: 16 }, { wch: 18 }, { wch: 14 }, { wch: 14 }];
const wb = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(wb, ws, "執行中軍購案個案管制表");
const OUT = path.join(PROJ, "執行中軍購案個案管制表(演示用).xlsx");
fs.writeFileSync(OUT, XLSX.write(wb, { bookType: "xlsx", type: "buffer" }));
console.log(`\n已寫出：${OUT}`);
console.log(`資料 ${cases.length} 筆、${HEADERS.length} 欄`);
