import React, { useState, useCallback, useEffect } from "react";
import { BUILTIN_MASTER } from "./masterData";
import { BUILTIN_MASTER_PETIT } from "./masterDataPetit";
import * as XLSX from "xlsx-js-style";
import "./App.css";

// ══════════════════════════════════════════════
// ── 공통 유틸 ──
// ══════════════════════════════════════════════
function normalize(s) {
  return s.replace(/[-_\s()[\]]/g, "").replace(/[가-힣]/g, "").toUpperCase();
}
function extractTokens(s) {
  const tokens = [];
  const re1 = /[A-Za-z]{1,4}[\d][\w\-.]*/g;
  let m;
  while ((m = re1.exec(s)) !== null) tokens.push(m[0]);
  const re2 = /[A-Za-z]+-\d+/g;
  while ((m = re2.exec(s)) !== null) tokens.push(m[0]);
  return tokens;
}
function masterVariants(mc) {
  const set = new Set();
  set.add(normalize(mc));
  extractTokens(mc).forEach((tok) => {
    set.add(normalize(tok));
    const base = tok.match(/^[A-Za-z]{1,4}\d+/);
    if (base) set.add(normalize(base[0]));
  });
  return set;
}
function skuVariants(name) {
  const vars = [];
  const packRe = /Pack_([A-Za-z]{1,4}[\d][\w\-.]*)/g;
  let m;
  while ((m = packRe.exec(name)) !== null) {
    vars.push(normalize(m[1]));
    const base = m[1].match(/^[A-Za-z]{1,4}\d+/);
    if (base) vars.push(normalize(base[0]));
  }
  extractTokens(name).forEach((tok) => {
    vars.push(normalize(tok));
    const base = tok.match(/^[A-Za-z]{1,4}\d+/);
    if (base) vars.push(normalize(base[0]));
  });
  return vars;
}
function buildMasterIndex(masterCodes) {
  const index = {};
  masterCodes.forEach((mc, i) => {
    masterVariants(mc).forEach((key) => {
      if (!(key in index)) index[key] = { mc, i };
    });
  });
  return index;
}
function findRankExact(name, index) {
  name = normalizeColorAlias(name);
  for (const key of skuVariants(name)) {
    if (key in index) return { rank: index[key].i, mc: index[key].mc };
  }
  return null;
}
// 색상 별칭 정규화 (매칭용)
function normalizeColorAlias(s) {
  return s
    // 형광 계열 (일반 색상보다 먼저)
    .replace(/형광녹색|형광그린/g,"FG")
    .replace(/형광노랑|형광황색/g,"FL")
    .replace(/형광핑크|형광분홍/g,"FP")
    .replace(/형광주황|형광오렌지/g,"FO")
    // 일반 색상 → 영문 단자 통일
    .replace(/적색|빨강|빨간/g,"R")
    .replace(/청색|파랑|파란/g,"B")
    .replace(/녹색|초록|초록색/g,"G")
    .replace(/황색|노랑|노란/g,"Y")
    .replace(/혼합/g,"A")
    .replace(/흑색|검정|블랙/g,"K")
    .replace(/백색|흰색|화이트/g,"W")
    .replace(/블루/g,"B").replace(/BLUE/gi,"B")
    .replace(/레드/g,"R").replace(/RED/gi,"R")
    .replace(/그린/g,"G").replace(/GREEN/gi,"G")
    .replace(/옐로/g,"Y").replace(/YELLOW/gi,"Y")
    .replace(/X\d+/gi,"")  // X5, X6 등 무시
    .replace(/ver\.?\d+/gi,"").replace(/v\.?\d+(?=\s|$|\))/gi,""); // ver 무시
}
function findRankFuzzy(name, masterCodes) {
  name = normalizeColorAlias(name);
  const nameN = normalize(name);
  let best = null;
  const seen = new Set();
  masterCodes.forEach((mc, i) => {
    if (seen.has(i)) return;
    seen.add(i);
    const mcN = normalize(mc);
    const minLen = Math.min(nameN.length, mcN.length);
    for (let l = minLen; l > 5; l--) {
      let found = false;
      for (let s = 0; s <= nameN.length - l; s++) {
        const sub = nameN.slice(s, s + l);
        if (mcN.includes(sub)) {
          if (!best || l > best.len) best = { rank: i, mc, len: l };
          found = true; break;
        }
      }
      if (found) break;
    }
  });
  if (best) return best;
  const nameRaw = name.replace(/ver\.?\d+/gi,"").replace(/[\s()[\]]/g, "");
  for (let l = Math.min(8, nameRaw.length); l >= 5; l--) {
    for (let s = 0; s <= nameRaw.length - l; s++) {
      const sub = nameRaw.slice(s, s + l);
      if (!/[가-힣]/.test(sub)) continue;
      for (let i = 0; i < masterCodes.length; i++) {
        const mcRaw = masterCodes[i].replace(/[\s()[\]]/g, "");
        if (mcRaw.includes(sub)) return { rank: i, mc: masterCodes[i], len: l };
      }
    }
  }
  return null;
}
function extractPackQty(name) {
  if (!name) return null;
  const n = name;

  // 소음방지보드는 매 단위 무시, 개/개입만
  const isSoundBoard = /소음방지보드/.test(n);

  // 괄호 있는 패턴 우선
  let m = n.match(/\((\d+)개입\)/);
  if (m) return parseInt(m[1], 10);
  m = n.match(/\((\d+)매입\)/);
  if (m) return parseInt(m[1], 10);
  m = n.match(/\((\d+)개\)/);
  if (m) return parseInt(m[1], 10);
  if (!isSoundBoard) {
    m = n.match(/\((\d+)매\)/);
    if (m) return parseInt(m[1], 10);
  }

  // 괄호 없는 패턴
  if (/3개입/.test(n)) return 3;
  // N장 → N매입
  m = n.match(/(\d+)장/);
  if (m) return parseInt(m[1], 10);
  // N매 → N개입 (소음방지보드 제외)
  if (!isSoundBoard) {
    m = n.match(/(\d+)매(?!입)/);
    if (m) return parseInt(m[1], 10);
  }

  return null;
}
function mergeRows(raw) {
  const skip = new Set();
  const result = [];
  for (let i = 0; i < raw.length; i++) {
    if (skip.has(i)) continue;
    const row = [...raw[i]];
    const cVal = row[2] != null ? String(row[2]).trim() : "";
    const eVal = row[4];
    if (cVal && !/^\d+$/.test(cVal) && (eVal == null || eVal === "" || eVal === 0)) {
      for (let j = i + 1; j < Math.min(i + 4, raw.length); j++) {
        const nextC = raw[j][2] != null ? String(raw[j][2]).trim() : "";
        if (/^\d+$/.test(nextC)) {
          row[4] = parseInt(nextC, 10);
          skip.add(j);
          for (let k = i + 1; k < j; k++) skip.add(k);
          break;
        }
      }
    }
    result.push(row);
  }
  return result;
}
function parseMasterXlsx(arrayBuffer) {
  const wb = XLSX.read(arrayBuffer, { type: "array" });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1 });
  return rows.slice(1).map((r) => (r[0] != null ? String(r[0]).trim() : "")).filter(Boolean);
}

// ══════════════════════════════════════════════
// ── 플로엠 정렬 로직 ──
// ══════════════════════════════════════════════
const SKU_ID_MAP = {
  "59366013": "아크릴 홀더 1단",
  "59366014": "아크릴 홀더 1단",
  "64312192": "아크릴 홀더 검정 1단",
  "64312193": "아크릴 홀더 흰색 1단",
};

// SKU ID → 강제 개입수 (브랜드 공통)
const SKU_ID_PACK_QTY = {
  // 1개입
  "2134831":1,"2134840":1,
  "2044918":1,"2096645":1,"2044941":1,"2113960":1,"2113978":1,
  "2114007":1,"2114011":1,"10221254":1,"2044963":1,
  "547994":1,"547812":1,"547870":1,"547872":1,"547918":1,"547920":1,"547922":1,"547931":1,"547987":1,"547808":1,"547809":1,"547817":1,"547830":1,"547839":1,"547844":1,"547845":1,"547855":1,"547868":1,"547883":1,"547890":1,"547897":1,"547898":1,"547909":1,"547910":1,"547916":1,"547921":1,"547924":1,"547929":1,"547934":1,"547943":1,"547952":1,"547954":1,"547970":1,"547975":1,"547999":1,"548000":1,"548002":1,"548009":1,"548010":1,"548015":1,"547944":1,"547838":1,"547903":1,"547951":1,"547816":1,"548001":1,
  // 2개입
  "14715868":2,"3679747":2,"3679744":2,"3679743":2,"8274318":2,"3679808":2,
  "3679810":2,"3679753":2,"3679812":2,"24301128":2,"24301121":2,
  "17227652":2,"17227653":2,"17227654":2,"17227655":2,"17277626":2,
  // 3개입
  "24301124":3,"24301127":3,
  "6548986":3,"6548987":3,"6561034":3,"6561020":3,"6561006":3,"6561004":3,
  "6561003":3,"6560996":3,"6560992":3,"6560987":3,"3043884":3,
  "6560989":3,"6561114":3,"6561121":3,"6561086":3,"6560975":3,"6560983":3,"6560984":3,"6560985":3,"6560986":3,"6560991":3,"6560994":3,"6561002":3,"6561005":3,"6561022":3,"6561023":3,"6561026":3,"6561033":3,"6561036":3,"6561046":3,"6561052":3,"6561054":3,"6561056":3,"6561076":3,"6561077":3,"6561084":3,"6561085":3,"6561087":3,"6561088":3,"6561091":3,"6561093":3,"6561099":3,"6561100":3,"6561107":3,"6561127":3,"6561128":3,"6561000":3,"6561007":3,"6561009":3,"6561038":3,"6561050":3,"6561053":3,"6561057":3,"6561082":3,"6561089":3,"6561090":3,"6561095":3,"6561096":3,"6561123":3,"6561125":3,"6561126":3,"62368835":3,"6560981":3,"6561035":3,"6561051":3,"6561055":3,"6561101":3,"6561116":3,"6561124":3,
  "3679822":3,"3679837":3,"3679842":3,"3679830":3,"3679811":3,
  "3679783":3,"3679782":3,"3679775":3,"3679773":3,"3679772":3,"3679766":3,
  "3679763":3,"3679762":3,"3679761":3,"3679760":3,"3679759":3,"3679758":3,
  "3679757":3,"3679756":3,"3679750":3,"3679749":3,"3679739":3,"3679732":3,
  "3679729":3,"3679728":3,"3679727":3,"3679726":3,"3267274":3,
  "6561021":3,"6560993":3,"6560999":3,"6561031":3,"6561078":3,"6561115":3,"6561044":3,"6561025":3,"6561122":3,"6560977":3,"6561080":3,"6560974":3,"6560976":3,"6561028":3,"6561024":3,"18314410":3,"6561092":3,"6561081":3,"6561094":3,"6561097":3,"6561098":3,"6561079":3,"6560982":3,"6560988":3,"6560995":3,"6560997":3,"6561029":3,"6561032":3,"6561039":3,"6561040":3,"6561041":3,"6561043":3,"6561045":3,"6561047":3,"6561048":3,"6561049":3,"6561083":3,"6561102":3,"6561104":3,"6561106":3,"6561109":3,"6965227":3,"6561110":3,"6560979":3,"6561117":3,"6561119":3,"6561105":3,"6561037":3,"6561108":3,"6561001":3,"6561103":3,"6561118":3,
  "6561113":3,"3043883":3,"6560990":3,"20283296":3,
  // 4개입
  "8245967":4,"8245986":4,"8245974":4,"8245993":4,"8245977":4,
  "3043880":4,"3043870":4,"3043869":4,"24301123":4,"17906742":4,
  "10155526":4,"11712157":4,
  "18314428":4,"17257844":4,"18501375":4,"17257850":4,"17257863":4,
  // 5개입
  "24301125":5,"18501380":5,"18501374":5,"18501385":5,"20283299":5,"20283318":5,
  "17257847":5,"17257856":5,"17257857":5,"18314418":5,"18314427":5,"18501343":5,"18501369":5,"18501378":5,"18501381":5,"18501383":5,"17257864":5,"18314407":5,"18314417":5,"18314436":5,"18314458":5,
  // 6개입
  "18501377":6,"10467740":6,"10467741":6,"10467755":6,"10467737":6,"10467728":6,"10467757":6,
  "10467734":6,"10467735":6,"10467753":6,"13404787":6,
  "24301126":6,"24301122":6,"10467758":6,"10467727":6,"10467746":6,"10467732":6,
  "18314420":6,"18501341":6,"18501344":6,"18501351":6,
};

function getFloemCode(name) {
  if (!name) return null;
  name = name.replace(/ver\.?\d+/gi,"").replace(/v\.?\d+(?=\s|$|\))/gi,"");
  const packM = name.match(/Pack_([A-Za-z]{1,4}[\d][\w\-.]*)/);
  if (packM) return normalize(packM[1]);
  const hyM = name.match(/\b([A-Za-z]{1,4})-(\d[\w\-]*)/);
  if (hyM) return normalize(hyM[1] + hyM[2]);
  const codes = name.match(/[A-Za-z]{1,4}[\d][\w\-.]*/g);
  return codes ? normalize(codes[0]) : null;
}
function getFloemGroup(code) {
  if (!code) return 99;
  const p = code.match(/^([A-Za-z]+)/);
  if (!p) return 99;
  const prefix = p[1].toUpperCase();
  if (prefix === "FL") return 3;
  if (prefix === "F") return 0;
  if (prefix === "L") return 1;
  if (prefix === "V") return 2;
  if (prefix === "K") return 4;
  return 99;
}
function getFloemSortKey(code) {
  if (!code) return [999999, ""];
  const m = code.match(/^[A-Za-z]+?(\d+)([A-Za-z]*)/);
  return m ? [parseInt(m[1], 10), m[2]] : [999999, ""];
}

const FLOEM_GROUP_LABELS = {
  "-1": "⚠ 수량 미확인", 0: "F 시리즈", 1: "L 시리즈",
  2: "V 시리즈", 3: "FL 시리즈", 4: "K 시리즈", 99: "기타",
};
const FLOEM_GROUP_COLORS = {
  "-1": { bg: "#B71C1C", light: "#FFEBEE", alt: "#FFCDD2" },
  0:   { bg: "#1565C0", light: "#E3F0FF", alt: "#CCE0FF" },
  1:   { bg: "#2E7D32", light: "#E8F5E9", alt: "#D0EBD1" },
  2:   { bg: "#F57F17", light: "#FFF8E1", alt: "#FDEFC3" },
  3:   { bg: "#6A1B9A", light: "#F3E5F5", alt: "#E4C8EE" },
  4:   { bg: "#00695C", light: "#E0F2F1", alt: "#B2DFDB" },
  99:  { bg: "#546E7A", light: "#FAFAFA", alt: "#EFEFEF" },
};
function fgc(group) { return FLOEM_GROUP_COLORS[String(group)] || FLOEM_GROUP_COLORS[99]; }

function processFloem(merged, masterCodes) {
  const index = buildMasterIndex(masterCodes);
  const result = [];
  merged.forEach((row) => {
    const nameStr = row[2] != null ? String(row[2]).trim() : "";
    if (!nameStr || /^\d+$/.test(nameStr)) return;
    const skuId = row[1] != null ? String(row[1]).trim() : "";
    const skuIdMapped = SKU_ID_MAP[skuId] || null;
    let match = findRankExact(nameStr, index);
    let method = "코드";
    if (!match) { const f = findRankFuzzy(nameStr, masterCodes); if (f) { match = f; method = "유사"; } }
    if (skuIdMapped) {
      const mappedIdx = masterCodes.indexOf(skuIdMapped);
      if (mappedIdx >= 0) { match = { rank: mappedIdx, mc: skuIdMapped }; method = "SKUID"; }
    }
    const code = skuIdMapped ? null : getFloemCode(nameStr);
    const group = getFloemGroup(code);
    const [sortNum, sortSuffix] = getFloemSortKey(code);
    const rawPackQty = extractPackQty(nameStr);
    const packQty = SKU_ID_PACK_QTY[skuId] != null ? SKU_ID_PACK_QTY[skuId] : rawPackQty;
    const eVal = row[4];
    const eNum = eVal != null && eVal !== "" ? Number(eVal) : null;
    const total = eNum != null && packQty != null ? eNum * packQty : null;
    const noQty = eNum == null || packQty == null;
    const effectiveSortNum = (sortNum === 999999 && match) ? match.rank : sortNum;
    const displayName = skuIdMapped || null;
    result.push({ group: noQty ? -1 : group, sortNum: effectiveSortNum, sortSuffix, code, displayName, master: match?.mc || null, method, packQty, total, noQty, row: [...row] });
  });
  result.sort((a, b) => {
    // 수량 미확인 맨 위
    if (a.noQty !== b.noQty) return a.noQty ? -1 : 1;
    return a.group - b.group || a.sortNum - b.sortNum || a.sortSuffix.localeCompare(b.sortSuffix);
  });
  return result;
}

// ══════════════════════════════════════════════
// ── 쁘띠팬시 정렬 로직 ──
// ══════════════════════════════════════════════

// 색상 정규화 (영문코드 우선)
const COLOR_ALIAS = {
  "219R": ["219R","219적색","219RED","219R"],
  "B": ["B","청색","블루","BLUE","BL"],
  "A": ["A","혼합","MIX","ASSORT"],
  "W": ["W","흰색","WHITE","화이트"],
  "R": ["R","적색","RED","빨강"],
  "G": ["G","녹색","GREEN","초록"],
  "Y": ["Y","황색","YELLOW","노랑"],
  "K": ["K","흑색","BLACK","검정","블랙"],
};
function normalizeColor(str) {
  const s = str.toUpperCase().replace(/[\s_-]/g,"");
  for (const [canon, aliases] of Object.entries(COLOR_ALIAS)) {
    if (aliases.some(a => s.includes(a.toUpperCase().replace(/[\s_-]/g,"")))) return canon;
  }
  return s;
}

// 쁘띠 카테고리 판별
function getPetitCategory(code, name) {
  if (!code && !name) return "기타";
  const c = (code || "").toUpperCase().replace(/_/g,"");
  const n = (name || "").toUpperCase();
  if (c.startsWith("DA") || c.startsWith("PD") || c.startsWith("TS")) return "스티커";
  if (c.startsWith("20") || c.startsWith("OPM") || c.startsWith("DT") || c.startsWith("HR")) return "견출지";
  if (n.includes("스티커") || n.includes("STICKER")) return "스티커";
  if (n.includes("견출") || n.includes("인덱스")) return "견출지";
  return "기타";
}

// 스티커 그룹
function getPetitStickerGroup(code) {
  if (!code) return 99;
  const c = code.toUpperCase().replace(/_/g,"");
  if (c.startsWith("DA")) return 0;
  if (c.startsWith("PD")) return 1;
  if (c.startsWith("TS")) return 2;
  return 99;
}
// 견출지 그룹
// 견출지 18가지 분류 (C열 SKU이름 기준)
const LABEL_CATS = [
  [0,  "인덱스",    ["인덱스"]],
  [1,  "보호",      ["보호"]],
  [2,  "칼라분류",  ["칼라분류"]],
  [3,  "홀로그램",  ["홀로그램"]],
  [4,  "유광",      ["유광"]],
  [5,  "크라프트",  ["크라프트","크래프트"]],
  [6,  "냉장고",    ["냉장고"]],
  [7,  "봉인",      ["봉인"]],
  [8,  "화살표",    ["화살표"]],
  [9,  "칼라",      ["칼라"]],
  [10, "눈알",      ["눈알"]],
  [11, "모서리",    ["모서리"]],
  [12, "모양",      ["모양"]],
  [13, "장식",      ["장식"]],
  [14, "링라벨",    ["링라벨","링 라벨"]],
  [15, "숫자",      ["숫자"]],
  [16, "마스킹",    ["마스킹"]],
  [17, "스티커명찰",["스티커명찰","스티커 명찰"]],
];
function getPetitLabelGroup(code, name) {
  const n = (name || "").replace(/\s/g,"").toUpperCase();
  for (const [order, label, keywords] of LABEL_CATS) {
    for (const kw of keywords) {
      if (n.includes(kw.replace(/\s/g,"").toUpperCase())) return order;
    }
  }
  // 코드 기반 폴백
  if (!code) return 99;
  const c = code.toUpperCase().replace(/_/g,"");
  if (c.startsWith("20")) return 99;
  if (c.startsWith("OPM")) return 99;
  if (c.startsWith("DT")) return 99;
  if (c.startsWith("HR")) return 99;
  return 99;
}
// 기타 그룹
function getPetitEtcGroup(name) {
  const n = (name||"").replace(/\s/g,"");
  if (/만년스[탬템]프/.test(n)) return 0;
  if (/스[탬템]프/.test(n)) return 1;
  if (/패드/.test(n)) return 2;
  if (/명찰/.test(n)) return 3;
  return 4;
}

function getPetitCode(name) {
  if (!name) return null;
  // Pack_ 제거, 's 제거, ver 제거
  let n = name
    .replace(/Pack_?/gi, "")
    .replace(/_/g, "")
    .replace(/'s/gi, "").replace(/\u2019s/gi, "")
    .replace(/ver\.?\d+/gi, "").replace(/v\.?\d+(?=\s|$|\))/gi, "")
    .trim();

  // 스티커: DA, PD, TS 코드 (견출지보다 먼저 체크)
  const stickerM = n.match(/\b(DA\d+[\w-]*|PD\d+[\w-]*|TS\d+[\w-]*)/i);
  if (stickerM) return stickerM[1].toUpperCase();

  // 견출지: 20- 포함 코드 추출 (색상까지 포함)
  // 패턴: 20-xxx(색상) 또는 20-xxx색상(한글) 또는 20-xxxENG코드
  const labelFullM = n.match(/\b(20-[A-Za-z]?\d+[A-Za-z]*)(\([^)]+\)|[가-힣]+)?/i);
  if (labelFullM) {
    const base = labelFullM[1].toUpperCase();
    const colorPart = labelFullM[2] || "";
    // 뒤에 영문 색상코드가 base 끝에 붙어있는 경우 (20-403G)
    const engColorM = base.match(/^(20-[A-Za-z]?\d+)([A-Z]{1,3})$/i);
    if (engColorM) {
      const inferredColor = inferColorCode(engColorM[2]);
      if (inferredColor) return (engColorM[1] + inferredColor).toUpperCase();
      return base;
    }
    // 괄호 색상 또는 한글 색상 붙어있는 경우
    if (colorPart) {
      const inferredColor = inferColorCode(colorPart);
      if (inferredColor) return (base + inferredColor).toUpperCase();
      return (base + colorPart).toUpperCase();
    }
    return base;
  }

  // OPM-, DT, HR 코드
  const otherLabelM = n.match(/\b(OPM-[A-Za-z0-9]+|DT\d+|HR\d+)/i);
  if (otherLabelM) return otherLabelM[1].toUpperCase();

  // 20- 없는 견출지 추론: 숫자+색상 → 20-번호+색상
  const labelGuessM = n.match(/^(?:[^\w]*)([A-Za-z]{0,2}\d{2,4})(.*)/);
  if (labelGuessM) {
    const numPart = labelGuessM[1];
    const rest = labelGuessM[2] || "";
    // rest 끝의 영문 색상코드 또는 한글 색상
    const engM = rest.match(/([A-Z]{1,3})$/i);
    let colorCode = inferColorCode(rest + n);
    if (!colorCode && engM) colorCode = inferColorCode(engM[1]);
    if (colorCode) return ("20-" + numPart + colorCode).toUpperCase();
    if (numPart.length >= 2) return ("20-" + numPart).toUpperCase();
  }

  // 일반 코드 (X5, X6 제외)
  const codes = n.match(/[A-Za-z]{2,}\d{2,}[\w.-]*/g);
  return codes ? codes[0].toUpperCase() : null;
}

// 영문 색상코드 → 한글 마스터 표기 변환
const ENG_COLOR_MAP = {
  "FG":"(형광녹색)", "FL":"(형광노랑)", "FP":"(형광핑크)",
  "FO":"(형광주황)", "FR":"(형광적색)",
  "R":"(적색)", "BL":"(청색)", "B":"(청색)",
  "G":"(녹색)", "Y":"(노랑)", "A":"(혼합)",
  "BK":"(검정)", "K":"(검정)", "W":"(흰색)",
  "SL":"(은색)",
};
function inferColorCode(s) {
  // 형광 한글
  if (/형광녹색|형광그린/.test(s)) return "(형광녹색)";
  if (/형광노랑|형광황색/.test(s)) return "(형광노랑)";
  if (/형광핑크|형광분홍/.test(s)) return "(형광핑크)";
  if (/형광주황|형광오렌지/.test(s)) return "(형광주황)";
  if (/형광적색|형광빨강/.test(s)) return "(형광적색)";
  // 일반 한글
  if (/빨강|빨간/.test(s)) return "(빨강)";
  if (/적색/.test(s)) return "(적색)";
  if (/파랑|파란/.test(s)) return "(파랑)";
  if (/청색/.test(s)) return "(청색)";
  if (/녹색|초록/.test(s)) return "(녹색)";
  if (/황색|노랑|노란/.test(s)) return "(노랑)";
  if (/혼합/.test(s)) return "(혼합)";
  if (/흑색|검정|블랙/.test(s)) return "(검정)";
  if (/백색|흰색|화이트/.test(s)) return "(흰색)";
  if (/하늘/.test(s)) return "(하늘)";
  if (/분홍|핑크/.test(s)) return "(분홍)";
  if (/보라/.test(s)) return "(보라)";
  if (/투명/.test(s)) return "(투명)";
  if (/은색|실버/.test(s)) return "(은색)";
  if (/금색|골드/.test(s)) return "(금색)";
  // 영문 코드: 문자열 끝에 오는 영문만 (FO, BL, BK, R, B, G 등)
  const engM = s.match(/([A-Z]{1,3})$/i);
  if (engM) {
    const code = engM[1].toUpperCase();
    if (ENG_COLOR_MAP[code]) return ENG_COLOR_MAP[code];
  }
  return null;
}

function getPetitSortNum(code) {
  if (!code) return 999999;
  const m = code.replace(/_/g,"").match(/[A-Za-z]*(\d+)/);
  return m ? parseInt(m[1], 10) : 999999;
}

function processPetit(merged, masterCodes) {
  const index = buildMasterIndex(masterCodes);
  const result = [];
  merged.forEach((row) => {
    const nameStr = row[2] != null ? String(row[2]).trim() : "";
    if (!nameStr || /^\d+$/.test(nameStr)) return;
    const skuId = row[1] != null ? Number(row[1]) : 0;
    const code = getPetitCode(nameStr);
    const category = getPetitCategory(code, nameStr);
    let catGroup, subGroup;
    if (category === "스티커") {
      catGroup = 0;
      subGroup = getPetitStickerGroup(code);
    } else if (category === "견출지") {
      catGroup = 1;
      subGroup = getPetitLabelGroup(code, nameStr);
    } else {
      catGroup = 2;
      subGroup = getPetitEtcGroup(nameStr);
    }
    const sortNum = getPetitSortNum(code);
    let match = findRankExact(nameStr, index);
    if (!match) { const f = findRankFuzzy(nameStr, masterCodes); if (f) match = f; }
    const skuIdStr = String(row[1] || "").trim();
    // 쁘띠견출지 무시 패턴
    const labelIgnorePattern = /\d+매(?!입)|\d+장|300개입|15개입/;
    const nameForQty = (catGroup === 1)
      ? nameStr.replace(labelIgnorePattern, "")
      : nameStr;
    const rawPackQty = extractPackQty(nameForQty);
    // SKU ID 강제 개입수 우선
    let packQty = SKU_ID_PACK_QTY[skuIdStr] != null ? SKU_ID_PACK_QTY[skuIdStr] : rawPackQty;
    // 쁘띠견출지: 개입수 없으면 매입가(K열=index10) ÷ 900으로 계산
    if (packQty == null && catGroup === 1) {
      const price = row[10] != null ? Number(row[10]) : null;
      if (price && price > 0) {
        const ratio = price / 900;
        if (ratio >= 2.7 && ratio <= 3.2) packQty = 3;
      }
    }
    const eVal = row[4];
    const eNum = eVal != null && eVal !== "" ? Number(eVal) : null;
    const total = eNum != null && packQty != null ? eNum * packQty : null;
    const noQty = eNum == null || packQty == null;
    // 특정 SKU ID 강제 분류
    if (skuIdStr === "62368835") { catGroup = 1; subGroup = 2; } // 칼라분류
    result.push({ catGroup, subGroup, sortNum, skuId, code, master: match?.mc || null, packQty, total, noQty, row: [...row] });
  });
  // 정렬: 카테고리 → 서브그룹 → 숫자 → SKU ID 오름차순
  result.sort((a, b) => {
    if (a.noQty !== b.noQty) return a.noQty ? -1 : 1; // 수량없는 건 맨 위
    return a.catGroup - b.catGroup || a.subGroup - b.subGroup || a.sortNum - b.sortNum || a.skuId - b.skuId;
  });
  return result;
}

const PETIT_CAT_LABELS = { 0: "쁘띠스티커", 1: "쁘띠견출지", 2: "쁘띠기타" };
const PETIT_STICKER_LABELS = { 0: "DA", 1: "PD", 2: "TS", 99: "기타" };
const PETIT_LABEL_LABELS = {
  0:"인덱스", 1:"보호", 2:"칼라분류", 3:"홀로그램", 4:"유광",
  5:"크라프트", 6:"냉장고", 7:"봉인", 8:"화살표", 9:"칼라",
  10:"눈알", 11:"모서리", 12:"모양", 13:"장식", 14:"링라벨",
  15:"숫자", 16:"마스킹", 17:"스티커명찰", 99:"기타"
};
const PETIT_ETC_LABELS = { 0: "만년스템프/스탬프", 1: "스탬프/스템프", 2: "패드", 3: "명찰", 4: "기타" };
const PETIT_CAT_COLORS = {
  0: { bg: "#AD1457", light: "#FCE4EC", alt: "#F8BBD9" },
  1: { bg: "#1565C0", light: "#E3F0FF", alt: "#CCE0FF" },
  2: { bg: "#2E7D32", light: "#E8F5E9", alt: "#D0EBD1" },
};
function pgc(catGroup) { return PETIT_CAT_COLORS[catGroup] || PETIT_CAT_COLORS[2]; }
function getPetitGroupLabel(item) {
  if (item.catGroup === 0) return PETIT_STICKER_LABELS[item.subGroup] || "기타";
  if (item.catGroup === 1) return PETIT_LABEL_LABELS[item.subGroup] || "기타";
  return PETIT_ETC_LABELS[item.subGroup] || "기타";
}

// ══════════════════════════════════════════════
// ── 엑셀 생성 ──
// ══════════════════════════════════════════════
function buildFloemExcel(processed) {
  const wb = XLSX.utils.book_new();
  const wsData = [[
    "브랜드","SKU ID","SKU 이름","품목코드","SKU Barcode",
    "발주수량","개입수","합계수량","확정수량","입고수량",
    "매입가","총발주 매입금","발주번호","발주유형","발주현황",
    "물류센터","입고예정일","발주일","매입유형","면세여부",
    "생산연도","제조일자","유통(소비)기한","공급가","부가세","입고금액","Xdock",
  ]];
  let curGroup = -1;
  processed.forEach((item) => {
    if (item.group !== curGroup) {
      curGroup = item.group;
      const cnt = processed.filter((x) => x.group === item.group).length;
      wsData.push([`▶  ${FLOEM_GROUP_LABELS[String(item.group)]} — 숫자 오름차순  (${cnt}개)`]);
    }
    const r = item.row;
    wsData.push([r[0],r[1],r[2],item.code,r[3],r[4],item.packQty,item.total,
      r[5],r[6],r[7],r[8],r[9],r[10],r[11],r[12],r[13],r[14],
      r[15],r[16],r[17],r[18],r[19],r[20],r[21],r[22],r[23]]);
  });
  const ws1 = XLSX.utils.aoa_to_sheet(wsData);
  ws1["!cols"] = [
    {wch:14},{wch:12},{wch:50},{wch:14},{wch:16},{wch:8},{wch:8},{wch:10},
    {wch:8},{wch:8},{wch:10},{wch:14},{wch:12},{wch:8},{wch:12},{wch:8},
    {wch:12},{wch:12},{wch:8},{wch:8},{wch:10},{wch:10},{wch:12},{wch:8},{wch:8},{wch:10},{wch:8},
  ];
  // 그룹 헤더행 노란색
  const yellowFill = { patternType:"solid", fgColor:{ rgb:"FFE500" } };
  const boldFont = { bold:true };
  wsData.forEach((row, wsRow) => {
    const v = row[0] != null ? String(row[0]) : "";
    if (v.includes("숫자 오름차순")) {
      for (let c = 0; c < 27; c++) {
        const addr = XLSX.utils.encode_cell({ r: wsRow, c });
        if (!ws1[addr]) ws1[addr] = { t:"s", v: c===0 ? v : "" };
        ws1[addr].s = { fill:yellowFill, font:boldFont };
      }
    }
  });
  XLSX.utils.book_append_sheet(wb, ws1, "정렬결과");

  // 시트2
  const codeMap = {}, noQtyMap = {};
  processed.forEach((item) => {
    const key = item.displayName || item.code || String(item.row[2]||"").trim();
    if (!key) return;
    if (item.noQty) {
      if (!noQtyMap[key]) noQtyMap[key] = { group:item.group, sortNum:item.sortNum, count:0 };
      noQtyMap[key].count += 1;
    } else {
      if (!codeMap[key]) codeMap[key] = { group:item.group, sortNum:item.sortNum, qty:0, count:0 };
      const q = item.total!=null ? item.total : (item.row[4]!=null ? Number(item.row[4]) : 0);
      codeMap[key].qty += q;
      codeMap[key].count += 1;
    }
  });
  const ws2Data = [["품목코드","발주수량(합계)","행수"]];
  const noQtyEntries = Object.entries(noQtyMap).sort(([,a],[,b]) => a.group-b.group||a.sortNum-b.sortNum);
  if (noQtyEntries.length > 0) {
    ws2Data.push(["⚠ 수량 미확인 — 직접 확인 필요","",""]);
    noQtyEntries.forEach(([key,s]) => ws2Data.push([key,"확인필요",s.count]));
  }
  const summaryRows = Object.entries(codeMap).sort(([,a],[,b]) => a.group-b.group||a.sortNum-b.sortNum);
  let curG2 = -999;
  summaryRows.forEach(([key,s]) => {
    if (s.group!==curG2) { curG2=s.group; ws2Data.push([`▶ ${FLOEM_GROUP_LABELS[String(s.group)]}`,"",""]); }
    ws2Data.push([key,s.qty,s.count]);
  });
  const totalCount = Object.values(codeMap).reduce((s,v)=>s+v.count,0)
                   + Object.values(noQtyMap).reduce((s,v)=>s+v.count,0);
  const totalQty = Object.values(codeMap).reduce((s,v)=>s+v.qty,0);
  ws2Data.push(["▶ 합계",totalQty,totalCount]);
  const ws2 = XLSX.utils.aoa_to_sheet(ws2Data);
  ws2["!cols"] = [{wch:30},{wch:16},{wch:8}];
  XLSX.utils.book_append_sheet(wb, ws2, "품목별 합계");

  // 시트3: 개입수별 합계
  const packMap3 = {};
  processed.forEach((item) => {
    if (item.noQty) return;
    const pq = item.packQty;
    if (!pq) return;
    if (!packMap3[pq]) packMap3[pq] = { totalQty:0, totalCount:0, codeMap:{} };
    const key = item.displayName || item.code || String(item.row[2]||"").trim();
    const skuId = String(item.row[1]||"").trim();
    const skuName = String(item.row[2]||"").trim();
    const pm = packMap3[pq];
    const ckey = key + "_" + skuId;
    if (!pm.codeMap[ckey]) pm.codeMap[ckey] = { code:key, skuName, qty:0, count:0, skuId };
    const q = item.total!=null ? item.total : (item.row[4]!=null ? Number(item.row[4]) : 0);
    pm.codeMap[ckey].qty += q;
    pm.codeMap[ckey].count += 1;
    pm.totalQty += q;
    pm.totalCount += 1;
  });

  const ws3Data = [["개입수","품목코드","발주수량(합계)","행수","SKU 이름"]];
  [1,2,3,4,5,6,7,8,9,10].forEach((pq) => {
    const pm = packMap3[pq];
    if (!pm) return;
    ws3Data.push([`▶ ${pq}개입 (총 ${pm.totalQty}개 / ${pm.totalCount}행)`,"","","",""]);
    Object.values(pm.codeMap)
      .sort((a,b) => a.code.localeCompare(b.code) || Number(a.skuId)-Number(b.skuId))
      .forEach(v => {
        ws3Data.push(["", v.code, v.qty, v.count, v.skuName]);
      });
  });
  const noPackItems3 = processed.filter(x => !x.noQty && !x.packQty);
  if (noPackItems3.length > 0) {
    ws3Data.push(["▶ 개입수 미확인","","","",""]);
    noPackItems3.forEach(item => {
      const key = item.displayName || item.code || String(item.row[2]||"").trim();
      const q = item.row[4]!=null ? Number(item.row[4]) : 0;
      ws3Data.push(["", key, q, 1, String(item.row[2]||"")]);
    });
  }

  const ws3 = XLSX.utils.aoa_to_sheet(ws3Data);
  ws3["!cols"] = [{wch:30},{wch:18},{wch:14},{wch:6},{wch:55}];
  const yFill3 = { patternType:"solid", fgColor:{ rgb:"FFE500" } };
  ws3Data.forEach((row, ri) => {
    const v = row[0] != null ? String(row[0]) : "";
    if (v.startsWith("▶")) {
      for (let c = 0; c < 5; c++) {
        const addr = XLSX.utils.encode_cell({ r:ri, c });
        if (!ws3[addr]) ws3[addr] = { t:"s", v: c===0?v:"" };
        ws3[addr].s = { fill:yFill3, font:{ bold:true } };
      }
    }
  });
  XLSX.utils.book_append_sheet(wb, ws3, "개입수별 합계");
  XLSX.writeFile(wb, "발주정렬결과_플로엠.xlsx");
}

function buildPetitExcel(processed) {
  const wb = XLSX.utils.book_new();
  const wsData = [[
    "브랜드","SKU ID","SKU 이름","품목코드","SKU Barcode",
    "발주수량","개입수","합계수량","확정수량","입고수량",
    "매입가","총발주 매입금","발주번호","발주유형","발주현황",
    "물류센터","입고예정일","발주일","매입유형","면세여부",
    "생산연도","제조일자","유통(소비)기한","공급가","부가세","입고금액","Xdock",
  ]];
  let curCat = -1, curSub = -1;
  processed.forEach((item) => {
    if (item.noQty) return; // 수량없는건 따로
    if (item.catGroup !== curCat || item.subGroup !== curSub) {
      curCat = item.catGroup; curSub = item.subGroup;
      const label = `▶  ${PETIT_CAT_LABELS[item.catGroup]} — ${getPetitGroupLabel(item)}`;
      wsData.push([label]);
    }
    const r = item.row;
    wsData.push([r[0],r[1],r[2],item.code,r[3],r[4],item.packQty,item.total,
      r[5],r[6],r[7],r[8],r[9],r[10],r[11],r[12],r[13],r[14],
      r[15],r[16],r[17],r[18],r[19],r[20],r[21],r[22],r[23]]);
  });
  // 수량없는 것 맨 뒤
  const noQtyItems = processed.filter(x => x.noQty);
  if (noQtyItems.length > 0) {
    wsData.push(["⚠ 수량 미확인"]);
    noQtyItems.forEach(item => {
      const r = item.row;
      wsData.push([r[0],r[1],r[2],item.code,r[3],r[4],item.packQty,item.total,
        r[5],r[6],r[7],r[8],r[9],r[10],r[11],r[12],r[13],r[14],
        r[15],r[16],r[17],r[18],r[19],r[20],r[21],r[22],r[23]]);
    });
  }
  const ws1 = XLSX.utils.aoa_to_sheet(wsData);
  ws1["!cols"] = [
    {wch:14},{wch:12},{wch:50},{wch:14},{wch:16},{wch:8},{wch:8},{wch:10},
    {wch:8},{wch:8},{wch:10},{wch:14},{wch:12},{wch:8},{wch:12},{wch:8},
    {wch:12},{wch:12},{wch:8},{wch:8},{wch:10},{wch:10},{wch:12},{wch:8},{wch:8},{wch:10},{wch:8},
  ];
  const yellowFill = { patternType:"solid", fgColor:{ rgb:"FFE500" } };
  wsData.forEach((row, wsRow) => {
    const v = row[0] != null ? String(row[0]) : "";
    if (v.startsWith("▶  쁘띠")) {
      for (let c = 0; c < 27; c++) {
        const addr = XLSX.utils.encode_cell({ r:wsRow, c });
        if (!ws1[addr]) ws1[addr] = { t:"s", v: c===0?v:"" };
        ws1[addr].s = { fill:yellowFill, font:{ bold:true } };
      }
    }
  });
  XLSX.utils.book_append_sheet(wb, ws1, "정렬결과");

  // 시트2: 품목코드 기준, 모든 SKU ID별 행 표시 + 개입수 포함
  const codeMap = {}, noQtyMap2 = {};
  processed.forEach((item) => {
    const key = item.code || String(item.row[2]||"").trim();
    if (!key) return;
    const skuId = String(item.row[1]||"").trim();
    const skuName = String(item.row[2]||"").trim();
    const packQty = item.packQty;
    if (item.noQty) {
      if (!noQtyMap2[key]) noQtyMap2[key] = { catGroup:item.catGroup, subGroup:item.subGroup, sortNum:item.sortNum, entries:[] };
      noQtyMap2[key].entries.push({ skuId, skuName, packQty, count:1 });
    } else {
      if (!codeMap[key]) codeMap[key] = { catGroup:item.catGroup, subGroup:item.subGroup, sortNum:item.sortNum, skuMap:{} };
      const sm = codeMap[key].skuMap;
      if (!sm[skuId]) sm[skuId] = { qty:0, count:0, skuName, packQty };
      const q = item.total!=null ? item.total : (item.row[4]!=null ? Number(item.row[4]) : 0);
      sm[skuId].qty += q;
      sm[skuId].count += 1;
    }
  });
  // A=품목코드, B=발주수량합계, C=개입수, D=행수, E=SKU이름
  const ws2Data = [["품목코드","발주수량(합계)","개입수","행수","SKU 이름"]];
  // 수량 미확인
  const noQtyE = Object.entries(noQtyMap2).sort(([,a],[,b])=>a.catGroup-b.catGroup||a.subGroup-b.subGroup||a.sortNum-b.sortNum);
  if (noQtyE.length>0) {
    ws2Data.push(["⚠ 수량 미확인","","","",""]);
    noQtyE.forEach(([key,s])=>{
      s.entries.forEach(e=>ws2Data.push([key,"확인필요",e.packQty||"",1,e.skuName]));
    });
  }
  // 품목코드별 정렬
  const sumRows = Object.entries(codeMap).sort(([,a],[,b])=>a.catGroup-b.catGroup||a.subGroup-b.subGroup||a.sortNum-b.sortNum);
  let cc=-1, cs=-1;
  sumRows.forEach(([key,s])=>{
    if(s.catGroup!==cc||s.subGroup!==cs){cc=s.catGroup;cs=s.subGroup;ws2Data.push([`▶ ${PETIT_CAT_LABELS[s.catGroup]}`,"","","",""]);}
    const skuEntries = Object.entries(s.skuMap).sort(([a],[b])=>Number(a)-Number(b));
    skuEntries.forEach(([skuId, sv])=>{
      ws2Data.push([key, sv.qty, sv.packQty||"", sv.count, sv.skuName]);
    });
  });
  const totalCount = Object.values(codeMap).reduce((s,v)=>s+Object.values(v.skuMap).reduce((a,b)=>a+b.count,0),0)
                   + Object.values(noQtyMap2).reduce((s,v)=>s+v.entries.length,0);
  const totalQty = Object.values(codeMap).reduce((s,v)=>s+Object.values(v.skuMap).reduce((a,b)=>a+b.qty,0),0);
  ws2Data.push(["▶ 합계",totalQty,"",totalCount,""]);
  const ws2 = XLSX.utils.aoa_to_sheet(ws2Data);
  ws2["!cols"] = [{wch:18},{wch:14},{wch:8},{wch:6},{wch:55}];
  XLSX.utils.book_append_sheet(wb, ws2, "품목별 합계");

  // 시트3: 개입수별 합계
  const packMap = {}; // packQty → { items: [{code, skuName, qty, count}], totalQty, totalCount }
  processed.forEach((item) => {
    if (item.noQty) return;
    const pq = item.packQty;
    if (!pq) return;
    if (!packMap[pq]) packMap[pq] = { totalQty:0, totalCount:0, codeMap:{} };
    const key = item.code || String(item.row[2]||"").trim();
    const skuId = String(item.row[1]||"").trim();
    const skuName = String(item.row[2]||"").trim();
    const pm = packMap[pq];
    const ckey = key + "_" + skuId;
    if (!pm.codeMap[ckey]) pm.codeMap[ckey] = { code:key, skuName, qty:0, count:0, skuId };
    const q = item.total!=null ? item.total : (item.row[4]!=null ? Number(item.row[4]) : 0);
    pm.codeMap[ckey].qty += q;
    pm.codeMap[ckey].count += 1;
    pm.totalQty += q;
    pm.totalCount += 1;
  });

  const ws3Data = [["개입수","품목코드","발주수량(합계)","행수","SKU 이름"]];
  [1,2,3,4,5,6,7,8,9,10].forEach((pq) => {
    const pm = packMap[pq];
    if (!pm) return;
    // 섹션 헤더
    ws3Data.push([`▶ ${pq}개입 (총 ${pm.totalQty}개 / ${pm.totalCount}행)`,"","","",""]);
    // 품목별 행
    Object.values(pm.codeMap)
      .sort((a,b) => a.code.localeCompare(b.code) || Number(a.skuId)-Number(b.skuId))
      .forEach(v => {
        ws3Data.push(["", v.code, v.qty, v.count, v.skuName]);
      });
  });
  // 개입수 없는 항목
  const noPackItems = processed.filter(x => !x.noQty && !x.packQty);
  if (noPackItems.length > 0) {
    ws3Data.push(["▶ 개입수 미확인","","","",""]);
    noPackItems.forEach(item => {
      const key = item.code || String(item.row[2]||"").trim();
      const q = item.row[4]!=null ? Number(item.row[4]) : 0;
      ws3Data.push(["", key, q, 1, String(item.row[2]||"")]);
    });
  }

  const ws3 = XLSX.utils.aoa_to_sheet(ws3Data);
  ws3["!cols"] = [{wch:30},{wch:18},{wch:14},{wch:6},{wch:55}];

  // 섹션 헤더 노란색
  const yFill = { patternType:"solid", fgColor:{ rgb:"FFE500" } };
  ws3Data.forEach((row, ri) => {
    const v = row[0] != null ? String(row[0]) : "";
    if (v.startsWith("▶")) {
      for (let c = 0; c < 5; c++) {
        const addr = XLSX.utils.encode_cell({ r:ri, c });
        if (!ws3[addr]) ws3[addr] = { t:"s", v: c===0?v:"" };
        ws3[addr].s = { fill:yFill, font:{ bold:true } };
      }
    }
  });

  XLSX.utils.book_append_sheet(wb, ws3, "개입수별 합계");
  XLSX.writeFile(wb, "발주정렬결과_쁘띠팬시.xlsx");
}

// ══════════════════════════════════════════════
// ── 메인 앱 ──
// ══════════════════════════════════════════════
export default function App() {
  const [brand, setBrand] = useState(null); // "floem" | "petit"
  const [masterCodes, setMasterCodes] = useState(null);
  const [masterLoaded, setMasterLoaded] = useState(false);
  const [masterMeta, setMasterMeta] = useState(null);
  const [processed, setProcessed] = useState(null);
  const [stats, setStats] = useState(null);
  const [loading, setLoading] = useState(false);
  const [uploading, setUploading] = useState(false);
  const [uploadMsg, setUploadMsg] = useState(null);
  const [step, setStep] = useState("brand"); // brand | ready | result

  function selectBrand(b) {
    setBrand(b);
    const codes = b === "floem" ? BUILTIN_MASTER : BUILTIN_MASTER_PETIT;
    setMasterCodes(codes);
    setMasterLoaded(true);
    setMasterMeta({ count: codes.length });
    setStep("ready");
  }

  const handleMasterUpload = useCallback(async (e) => {
    const file = e.target.files[0]; if (!file) return;
    setUploading(true); setUploadMsg(null);
    try {
      const buf = await file.arrayBuffer();
      const codes = parseMasterXlsx(buf);
      if (codes.length === 0) throw new Error("품목 코드를 찾을 수 없어요");
      setMasterCodes(codes);
      setMasterMeta({ count: codes.length, updatedAt: new Date().toLocaleString("ko-KR") });
      setUploadMsg({ type:"ok", text:`✓ 임시 적용 ${codes.length.toLocaleString()}개` });
    } catch (err) { setUploadMsg({ type:"err", text:"실패: "+err.message }); }
    setUploading(false); e.target.value = "";
  }, []);

  const handleOrderFile = useCallback((e) => {
    const file = e.target.files[0]; if (!file) return;
    setLoading(true);
    const reader = new FileReader();
    reader.onload = (ev) => {
      try {
        const wb = XLSX.read(ev.target.result, { type:"array" });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const allRows = XLSX.utils.sheet_to_json(ws, { header:1, defval:null });
        const merged = mergeRows(allRows.slice(1));
        const result = brand === "floem"
          ? processFloem(merged, masterCodes)
          : processPetit(merged, masterCodes);
        const matched = result.filter(x => x.master).length;
        setStats({ total:result.length, matched, unmatched:result.length-matched });
        setProcessed(result);
        setStep("result");
      } catch (err) { alert("오류: "+err.message); }
      setLoading(false);
    };
    reader.readAsArrayBuffer(file);
    e.target.value = "";
  }, [masterCodes, brand]);

  const handleReset = () => { setProcessed(null); setStats(null); setStep("ready"); };
  const handleBrandReset = () => { setBrand(null); setProcessed(null); setStats(null); setMasterLoaded(false); setStep("brand"); };

  useEffect(() => {
    if (!uploadMsg) return;
    const t = setTimeout(()=>setUploadMsg(null), 4000);
    return ()=>clearTimeout(t);
  }, [uploadMsg]);

  const gc = brand === "floem" ? fgc : pgc;
  const GROUP_LABELS = brand === "floem" ? FLOEM_GROUP_LABELS : PETIT_CAT_LABELS;

  return (
    <div className="app">
      <header className="header">
        <div className="header-inner">
          <div style={{display:"flex",alignItems:"center",gap:12}}>
            {brand && (
              <button className="btn-back" onClick={handleBrandReset} title="브랜드 선택으로">
                ←
              </button>
            )}
            <div className="logo" onClick={handleBrandReset} style={{cursor:"pointer"}}>
              <span className="logo-mark">{brand === "petit" ? "P" : "F"}</span>
              <span className="logo-text">
                {brand === "petit" ? "쁘띠팬시 자동정렬 합계" : "플로엠제품 자동정렬 합계"}
              </span>
            </div>
          </div>
          <div className="header-right">
            {masterLoaded && masterMeta && (
              <div className="master-badge">✓ 마스터 {masterMeta.count.toLocaleString()}개</div>
            )}
            {brand && (
              <label className={`btn-master-upload ${uploading?"disabled":""}`}>
                {uploading?"업로드 중…":"📋 마스터 업데이트"}
                <input type="file" accept=".xlsx" onChange={handleMasterUpload} hidden disabled={uploading}/>
              </label>
            )}
          </div>
        </div>
      </header>

      {uploadMsg && (
        <div className={`toast ${uploadMsg.type}`} onClick={()=>setUploadMsg(null)}>{uploadMsg.text}</div>
      )}

      <main className="main">
        {/* 브랜드 선택 */}
        {step === "brand" && (
          <div className="card center-card">
            <div className="step-icon">🏪</div>
            <h2>브랜드 선택</h2>
            <p className="desc">정렬할 브랜드를 선택해주세요</p>
            <div style={{display:"flex",gap:16,width:"100%",marginTop:8}}>
              <button className="btn-brand floem" onClick={()=>selectBrand("floem")}>
                <span className="brand-icon">F</span>
                <span>플로엠</span>
              </button>
              <button className="btn-brand petit" onClick={()=>selectBrand("petit")}>
                <span className="brand-icon">P</span>
                <span>쁘띠팬시</span>
              </button>
            </div>
          </div>
        )}

        {/* 발주 파일 업로드 */}
        {step === "ready" && (
          <div className="card center-card">
            <div className="step-icon">📂</div>
            <h2>{brand === "floem" ? "플로엠" : "쁘띠팬시"} 발주파일 업로드</h2>
            <p className="desc">쿠팡에서 받은 {brand === "floem" ? "플로엠" : "쁘띠팬시"} 발주 엑셀을 올려주세요</p>
            <div
              className="upload-zone"
              onDragOver={(e)=>{e.preventDefault();e.currentTarget.classList.add("drag-over");}}
              onDragLeave={(e)=>e.currentTarget.classList.remove("drag-over")}
              onDrop={(e)=>{
                e.preventDefault();e.currentTarget.classList.remove("drag-over");
                const file=e.dataTransfer.files[0];
                if(file) handleOrderFile({target:{files:[file],value:""}});
              }}
              onClick={()=>document.getElementById("order-file-input").click()}
            >
              <input id="order-file-input" type="file" accept=".xlsx,.xls" onChange={handleOrderFile} hidden/>
              <div className="upload-icon">⬆</div>
              <div className="upload-text">{loading?"처리 중…":"파일을 클릭하거나 끌어다 놓으세요"}</div>
              <div className="upload-sub">.xlsx / .xls</div>
            </div>
            <button className="btn-ghost" onClick={handleBrandReset}>← 브랜드 다시 선택</button>
          </div>
        )}

        {/* 결과 */}
        {step === "result" && processed && (
          <div className="result-wrap">
            <div className="stats-row">
              <div className="stat-card"><div className="stat-num">{stats.total}</div><div className="stat-label">전체 품목</div></div>
              <div className="stat-card matched"><div className="stat-num">{stats.matched}</div><div className="stat-label">매칭 완료</div></div>
              <div className="stat-card unmatched"><div className="stat-num">{stats.unmatched}</div><div className="stat-label">미매칭</div></div>
            </div>
            <div className="action-row">
              <button className="btn-download" onClick={()=> brand==="floem" ? buildFloemExcel(processed) : buildPetitExcel(processed)}>
                ⬇ 엑셀 다운로드 (2시트)
              </button>
              <button className="btn-ghost" onClick={handleReset}>새 파일 처리</button>
            </div>

            {/* 미리보기 */}
            <div className="preview">
              <div className="preview-header">정렬결과</div>
              <div className="table-wrap">
                <table>
                  <thead><tr><th>#</th><th>품목코드</th><th>SKU 이름</th><th>발주수량</th><th>개입수</th><th>합계</th><th>매칭</th></tr></thead>
                  <tbody>
                    {(() => {
                      let cg=-1, cs=-1, rn=0;
                      return processed.map((item,idx)=>{
                        const rows=[];
                        const itemGroup = brand==="floem" ? item.group : item.catGroup*100+item.subGroup;
                        const prevGroup = cg;
                        if(itemGroup!==prevGroup){
                          cg=itemGroup;
                          const col = brand==="floem" ? fgc(item.group) : pgc(item.catGroup);
                          const label = brand==="floem"
                            ? FLOEM_GROUP_LABELS[String(item.group)]
                            : `${PETIT_CAT_LABELS[item.catGroup]} — ${getPetitGroupLabel(item)}`;
                          rows.push(<tr key={`g${idx}`} className="group-hdr" style={{background:col.bg}}><td colSpan={7}>▶ {label}</td></tr>);
                        }
                        rn++;
                        const col = brand==="floem" ? fgc(item.group) : pgc(item.catGroup);
                        rows.push(
                          <tr key={idx} style={{background:rn%2===0?col.alt:col.light}}>
                            <td className="td-num">{rn}</td>
                            <td className="td-code">{item.code||"—"}</td>
                            <td className="td-name">{String(item.row[2]||"")}</td>
                            <td className="td-center">{item.row[4]??"—"}</td>
                            <td className="td-center">{item.packQty??"—"}</td>
                            <td className="td-center bold">{item.total??"—"}</td>
                            <td className="td-center"><span className={item.master?"badge-ok":"badge-ng"}>{item.master?"✓":"✗"}</span></td>
                          </tr>
                        );
                        return rows;
                      });
                    })()}
                  </tbody>
                </table>
              </div>
            </div>
          </div>
        )}
      </main>
    </div>
  );
}
