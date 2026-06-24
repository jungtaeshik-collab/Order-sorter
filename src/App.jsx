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
  for (const key of skuVariants(name)) {
    if (key in index) return { rank: index[key].i, mc: index[key].mc };
  }
  return null;
}
function findRankFuzzy(name, masterCodes) {
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
  "547994":1,"547812":1,"547870":1,"547872":1,"547918":1,"547920":1,"547922":1,"547931":1,"547987":1,"547808":1,"547809":1,"547817":1,"547830":1,"547839":1,"547844":1,"547845":1,"547855":1,"547868":1,"547883":1,"547890":1,"547897":1,"547898":1,"547909":1,"547910":1,"547916":1,"547921":1,"547924":1,"547929":1,"547934":1,"547943":1,"547952":1,"547954":1,"547970":1,"547975":1,"547999":1,"548000":1,"548002":1,"548009":1,"548010":1,"548015":1,"547944":1,"547838":1,"547903":1,"547951":1,"547816":1,"548001":1,
  // 2개입
  "3679747":2,"3679744":2,"3679743":2,"8274318":2,"3679808":2,
  "3679810":2,"3679753":2,"3679812":2,"24301128":2,"24301121":2,
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
  // 4개입
  "8245967":4,"8245986":4,"8245974":4,"8245993":4,"8245977":4,
  "3043880":4,"3043870":4,"3043869":4,"24301123":4,"17906742":4,
  "10155526":4,"11712157":4,
  // 5개입
  "24301125":5,"18501380":5,
  // 6개입
  "10467741":6,"10467755":6,"10467737":6,"10467728":6,"10467757":6,
  "10467734":6,"10467735":6,"10467753":6,"13404787":6,
  "24301126":6,"24301122":6,"10467758":6,"10467727":6,"10467746":6,"10467732":6,
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
function getPetitLabelGroup(code) {
  if (!code) return 99;
  const c = code.toUpperCase().replace(/_/g,"");
  if (c.startsWith("20")) return 0;
  if (c.startsWith("OPM")) return 1;
  if (c.startsWith("DT")) return 2;
  if (c.startsWith("HR")) return 3;
  return 4;
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
  // 언더바, 's 무시하고 코드 추출 (멋쟁이팽귄's → 멋쟁이팽귄)
  const n = name.replace(/_/g,"").replace(/'s/gi,"").replace(/’s/gi,"").replace(/ver\.?\d+/gi,"").replace(/v\.?\d+/gi,"");
  const packM = n.match(/Pack_?([A-Za-z]{1,4}[\d][\w\-.]*)/i);
  if (packM) return packM[1].toUpperCase();
  const hyM = n.match(/([A-Za-z]{1,4})-(\d[\w\-]*)/);
  if (hyM) return (hyM[1]+hyM[2]).toUpperCase();
  const codes = n.match(/[A-Za-z]{1,4}[\d][\w\-.]*/g);
  return codes ? codes[0].toUpperCase() : null;
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
      subGroup = getPetitLabelGroup(code);
    } else {
      catGroup = 2;
      subGroup = getPetitEtcGroup(nameStr);
    }
    const sortNum = getPetitSortNum(code);
    let match = findRankExact(nameStr, index);
    if (!match) { const f = findRankFuzzy(nameStr, masterCodes); if (f) match = f; }
    const skuIdStr = String(row[1] || "").trim();
    // 쁘띠견출지 무시 패턴
    const labelIgnorePattern = /\d+매(?!입)|\d+장|300개입/;
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
const PETIT_LABEL_LABELS = { 0: "20-", 1: "OPM", 2: "DT", 3: "HR", 4: "기타" };
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

  // 시트2
  const codeMap = {}, noQtyMap2 = {};
  processed.forEach((item) => {
    const key = item.code || String(item.row[2]||"").trim();
    if (!key) return;
    if (item.noQty) {
      if (!noQtyMap2[key]) noQtyMap2[key] = { catGroup:item.catGroup, subGroup:item.subGroup, sortNum:item.sortNum, count:0 };
      noQtyMap2[key].count += 1;
    } else {
      if (!codeMap[key]) codeMap[key] = { catGroup:item.catGroup, subGroup:item.subGroup, sortNum:item.sortNum, qty:0, count:0 };
      const q = item.total!=null ? item.total : (item.row[4]!=null ? Number(item.row[4]) : 0);
      codeMap[key].qty += q;
      codeMap[key].count += 1;
    }
  });
  const ws2Data = [["품목코드","발주수량(합계)","행수"]];
  const noQtyE = Object.entries(noQtyMap2).sort(([,a],[,b])=>a.catGroup-b.catGroup||a.subGroup-b.subGroup||a.sortNum-b.sortNum);
  if (noQtyE.length>0) {
    ws2Data.push(["⚠ 수량 미확인","",""]);
    noQtyE.forEach(([key,s])=>ws2Data.push([key,"확인필요",s.count]));
  }
  const sumRows = Object.entries(codeMap).sort(([,a],[,b])=>a.catGroup-b.catGroup||a.subGroup-b.subGroup||a.sortNum-b.sortNum);
  let cc=-1, cs=-1;
  sumRows.forEach(([key,s])=>{
    if(s.catGroup!==cc||s.subGroup!==cs){cc=s.catGroup;cs=s.subGroup;ws2Data.push([`▶ ${PETIT_CAT_LABELS[s.catGroup]}`,``,``]);}
    ws2Data.push([key,s.qty,s.count]);
  });
  const totalCount = Object.values(codeMap).reduce((s,v)=>s+v.count,0)+Object.values(noQtyMap2).reduce((s,v)=>s+v.count,0);
  const totalQty = Object.values(codeMap).reduce((s,v)=>s+v.qty,0);
  ws2Data.push(["▶ 합계",totalQty,totalCount]);
  const ws2 = XLSX.utils.aoa_to_sheet(ws2Data);
  ws2["!cols"] = [{wch:30},{wch:16},{wch:8}];
  XLSX.utils.book_append_sheet(wb, ws2, "품목별 합계");
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
