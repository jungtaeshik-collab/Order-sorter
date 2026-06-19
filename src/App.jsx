import React, { useState, useCallback, useEffect } from "react";
import { BUILTIN_MASTER } from "./masterData";
import * as XLSX from "xlsx";
import { initializeApp } from "firebase/app";
import { getDatabase, ref, set, get } from "firebase/database";
import "./App.css";

// ── Firebase 설정 ──
const firebaseConfig = {
  apiKey: "AIzaSyAIJBHt3j73Ustm1BIxA8329yzFS2j1uMM",
  authDomain: "petit-subcon.firebaseapp.com",
  databaseURL: "https://petit-subcon-default-rtdb.firebaseio.com",
  projectId: "petit-subcon",
  storageBucket: "petit-subcon.firebasestorage.app",
  messagingSenderId: "1089448725745",
  appId: "1:1089448725745:web:63769e226cb991a3f4ad50",
};
const firebaseApp = initializeApp(firebaseConfig);
const db = getDatabase(firebaseApp);
const MASTER_PATH = "order-sorter/master";

// ── 유틸 함수 ──
function normalize(s) {
  // 하이픈/공백/괄호/한글 제거 → 영숫자만
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
  // 일반 normalize 매칭
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
  // 한글 포함 품목: 원본 문자열 기준 5~8글자 일치
  const nameRaw = name.replace(/[\s\(\)\[\]]/g, "");
  for (let l = Math.min(8, nameRaw.length); l >= 5; l--) {
    for (let s = 0; s <= nameRaw.length - l; s++) {
      const sub = nameRaw.slice(s, s + l);
      if (!/[가-힣]/.test(sub)) continue; // 한글 포함된 것만
      for (let i = 0; i < masterCodes.length; i++) {
        const mcRaw = masterCodes[i].replace(/[\s\(\)\[\]]/g, "");
        if (mcRaw.includes(sub)) {
          return { rank: i, mc: masterCodes[i], len: l };
        }
      }
    }
  }
  return null;
}
function getCodeFromSku(name) {
  if (!name) return null;
  const packM = name.match(/Pack_([A-Za-z]{1,4}[\d][\w\-.]*)/);
  if (packM) return normalize(packM[1]);
  const hyM = name.match(/\b([A-Za-z]{1,4})-(\d[\w\-]*)/);
  if (hyM) return normalize(hyM[1] + hyM[2]);
  const codes = name.match(/[A-Za-z]{1,4}[\d][\w\-.]*/g);
  return codes ? normalize(codes[0]) : null;
}
function getGroup(code) {
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
function getSortKey(code) {
  // 숫자 + 뒤 알파벳까지 포함 (K1000B < K1000BK < K1000W)
  if (!code) return [999999, ""];
  const m = code.match(/^[A-Za-z]+?(\d+)([A-Za-z]*)/);
  return m ? [parseInt(m[1], 10), m[2]] : [999999, ""];
}
function getSortNum(code) { return getSortKey(code)[0]; }
function getSortSuffix(code) { return getSortKey(code)[1]; }
function extractPackQty(name) {
  if (!name) return null;
  let m = name.match(/\((\d+)개입\)/);
  if (m) return parseInt(m[1], 10);
  m = name.match(/\((\d+)매입\)/);
  if (m) return parseInt(m[1], 10);
  m = name.match(/\((\d+)개\)/);
  if (m) return parseInt(m[1], 10);
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

// ── xlsx → codes 배열 변환 ──
function parseMasterXlsx(arrayBuffer) {
  const wb = XLSX.read(arrayBuffer, { type: "array" });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1 });
  return rows.slice(1)
    .map((r) => (r[0] != null ? String(r[0]).trim() : ""))
    .filter(Boolean);
}

// ── codes 배열 → base64 (DB 저장용) ──
function codesToBase64(codes) {
  const json = JSON.stringify(codes);
  return btoa(unescape(encodeURIComponent(json)));
}
function base64ToCodes(b64) {
  const json = decodeURIComponent(escape(atob(b64)));
  return JSON.parse(json);
}

const GROUP_LABELS = { "-1":"⚠ 수량 미확인", 0:"F 시리즈", 1:"L 시리즈", 2:"V 시리즈", 3:"FL 시리즈", 4:"K 시리즈", 99:"기타" };
const GROUP_COLORS = {
  "-1":{ bg:"#B71C1C", light:"#FFEBEE", alt:"#FFCDD2" },
  0:   { bg:"#1565C0", light:"#E3F0FF", alt:"#CCE0FF" },
  1:   { bg:"#2E7D32", light:"#E8F5E9", alt:"#D0EBD1" },
  2:   { bg:"#F57F17", light:"#FFF8E1", alt:"#FDEFC3" },
  3:   { bg:"#6A1B9A", light:"#F3E5F5", alt:"#E4C8EE" },
  4:   { bg:"#00695C", light:"#E0F2F1", alt:"#B2DFDB" },
  99:  { bg:"#546E7A", light:"#FAFAFA", alt:"#EFEFEF" },
};
function gc(group) { return GROUP_COLORS[String(group)] || GROUP_COLORS[99]; }

function buildExcel(processed) {
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
      wsData.push([`▶  ${GROUP_LABELS[item.group]} — 숫자 오름차순  (${cnt}개)`]);
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
  XLSX.utils.book_append_sheet(wb, ws1, "정렬결과");

  // 시트2: 품목별 합계 (수량없는 항목 맨 위 별도)
  const codeMap = {};
  const noQtyMap = {};
  processed.forEach((item) => {
    if (!item.code) return;
    if (item.noQty) {
      // 수량 미확인 항목
      if (!noQtyMap[item.code]) noQtyMap[item.code] = { group: item.group, sortNum: item.sortNum };
    } else {
      if (!codeMap[item.code]) codeMap[item.code] = { group: item.group, sortNum: item.sortNum, qty: 0 };
      const q = item.total != null ? item.total : (item.row[4] != null ? Number(item.row[4]) : 0);
      codeMap[item.code].qty += q;
    }
  });
  const ws2Data = [["품목코드","발주수량(합계)"]];
  // 수량 미확인 먼저
  const noQtyEntries = Object.entries(noQtyMap).sort(([,a],[,b]) => a.group - b.group || a.sortNum - b.sortNum);
  if (noQtyEntries.length > 0) {
    ws2Data.push(["⚠ 수량 미확인 — 직접 확인 필요"]);
    noQtyEntries.forEach(([code]) => ws2Data.push([code, "확인필요"]));
  }
  // 일반 합계
  const summaryRows = Object.entries(codeMap)
    .sort(([,a],[,b]) => a.group - b.group || a.sortNum - b.sortNum);
  let curG2 = -999;
  summaryRows.forEach(([code, s]) => {
    if (s.group !== curG2) { curG2 = s.group; ws2Data.push([`▶ ${GROUP_LABELS[s.group]}`]); }
    ws2Data.push([code, s.qty]);
  });
  const ws2 = XLSX.utils.aoa_to_sheet(ws2Data);
  ws2["!cols"] = [{wch:18},{wch:16}];
  XLSX.utils.book_append_sheet(wb, ws2, "품목별 합계");
  XLSX.writeFile(wb, "발주정렬결과.xlsx");
}

// ── 메인 앱 ──
export default function App() {
  const [masterCodes, setMasterCodes]   = useState(null);
  const [masterLoaded, setMasterLoaded] = useState(false);
  const [masterMeta, setMasterMeta]     = useState(null); // {count, updatedAt}
  const [masterError, setMasterError]   = useState(null);
  const [processed, setProcessed]       = useState(null);
  const [stats, setStats]               = useState(null);
  const [loading, setLoading]           = useState(false);
  const [uploading, setUploading]       = useState(false);
  const [uploadMsg, setUploadMsg]       = useState(null);
  const [step, setStep]                 = useState("loading");

  // ── 앱 시작 시 DB에서 마스터 자동 로드 ──
  useEffect(() => {
    loadMasterFromDB();
  }, []);

  async function loadMasterFromDB() {
    setLoading(true);
    setMasterError(null);
    try {
      const snap = await get(ref(db, MASTER_PATH));
      if (snap.exists()) {
        const data = snap.val();
        const codes = base64ToCodes(data.codes);
        setMasterCodes(codes);
        setMasterLoaded(true);
        setMasterMeta({ count: codes.length, updatedAt: data.updatedAt });
      } else {
        // Firebase에 없으면 내장 마스터 사용
        setMasterCodes(BUILTIN_MASTER);
        setMasterLoaded(true);
        setMasterMeta({ count: BUILTIN_MASTER.length, updatedAt: "기본내장" });
      }
    } catch (e) {
      // 오류 시에도 내장 마스터로 폴백
      setMasterCodes(BUILTIN_MASTER);
      setMasterLoaded(true);
      setMasterMeta({ count: BUILTIN_MASTER.length, updatedAt: "기본내장" });
    }
    setLoading(false);
    setStep("ready");
  }

  // ── 마스터 xlsx 업로드 → DB 저장 ──
  const handleMasterUpload = useCallback(async (e) => {
    const file = e.target.files[0];
    if (!file) return;
    setUploading(true);
    setUploadMsg(null);
    try {
      const buf = await file.arrayBuffer();
      const codes = parseMasterXlsx(buf);
      if (codes.length === 0) throw new Error("품목 코드를 찾을 수 없어요");
      const now = new Date().toLocaleString("ko-KR");
      await set(ref(db, MASTER_PATH), {
        codes: codesToBase64(codes),
        count: codes.length,
        updatedAt: now,
      });
      setMasterCodes(codes);
      setMasterLoaded(true);
      setMasterMeta({ count: codes.length, updatedAt: now });
      setUploadMsg({ type:"ok", text:`✓ 마스터 업데이트 완료 — ${codes.length.toLocaleString()}개 품목 (${now})` });
    } catch (err) {
      setUploadMsg({ type:"err", text:"업로드 실패: " + err.message });
    }
    setUploading(false);
    e.target.value = "";
  }, []);

  // ── 발주 파일 처리 ──
  const handleOrderFile = useCallback((e) => {
    const file = e.target.files[0];
    if (!file) return;
    setLoading(true);
    const reader = new FileReader();
    reader.onload = (ev) => {
      try {
        const wb = XLSX.read(ev.target.result, { type:"array" });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const allRows = XLSX.utils.sheet_to_json(ws, { header:1, defval:null });
        const merged = mergeRows(allRows.slice(1));
        const index = buildMasterIndex(masterCodes);
        const result = [];
        merged.forEach((row) => {
          const nameStr = row[2] != null ? String(row[2]).trim() : "";
          if (!nameStr || /^\d+$/.test(nameStr)) return;
          let match = findRankExact(nameStr, index);
          let method = "코드";
          if (!match) { const f = findRankFuzzy(nameStr, masterCodes); if (f) { match = f; method = "유사"; } }
          const code = getCodeFromSku(nameStr);
          const group = getGroup(code);
          const sortNum = getSortNum(code);
          const packQty = extractPackQty(nameStr);
          const eVal = row[4];
          const eNum = eVal != null && eVal !== "" ? Number(eVal) : null;
          const total = eNum != null && packQty != null ? eNum * packQty : null;
          // 수량 없는 항목은 그룹 -1 (맨 위)
          // 발주수량 없거나 개입수 없으면 합계 불가 → 맨 위
          const noQty = eNum == null || packQty == null;
          const sortSuffix = getSortSuffix(code);
          // 한글 품목처럼 코드가 없으면 마스터 rank를 sortNum으로 사용
          const effectiveSortNum = (sortNum === 999999 && match) ? match.rank : sortNum;
          result.push({ group: noQty ? -1 : group, sortNum: effectiveSortNum, sortSuffix, code, master:match?.mc||null, method, packQty, total, noQty, row:[...row] });
        });
        result.sort((a, b) => a.group - b.group || a.sortNum - b.sortNum || a.sortSuffix.localeCompare(b.sortSuffix));
        const matched = result.filter((x) => x.master).length;
        setStats({ total:result.length, matched, unmatched:result.length - matched });
        setProcessed(result);
        setStep("result");
      } catch (err) { alert("파일 처리 오류: " + err.message); }
      setLoading(false);
    };
    reader.readAsArrayBuffer(file);
    e.target.value = "";
  }, [masterCodes]);

  const handleReset = () => { setProcessed(null); setStats(null); setStep("ready"); };

  // 토스트 자동 닫기
  useEffect(() => {
    if (!uploadMsg) return;
    const t = setTimeout(() => setUploadMsg(null), 4000);
    return () => clearTimeout(t);
  }, [uploadMsg]);

  return (
    <div className="app">
      <header className="header">
        <div className="header-inner">
          <div className="logo">
            <span className="logo-mark">F</span>
            <span className="logo-text">플로엠 발주 정렬기</span>
          </div>
          <div className="header-right">
            {masterLoaded && masterMeta && (
              <div className="master-badge">
                ✓ 마스터 {masterMeta.count.toLocaleString()}개
                {masterMeta.updatedAt && <span className="badge-date"> · {masterMeta.updatedAt}</span>}
              </div>
            )}
            <label className={`btn-master-upload ${uploading ? "disabled" : ""}`}>
              {uploading ? "업로드 중…" : "📋 마스터 업데이트"}
              <input type="file" accept=".xlsx" onChange={handleMasterUpload} hidden disabled={uploading} />
            </label>
          </div>
        </div>
      </header>

      {uploadMsg && (
        <div className={`toast ${uploadMsg.type}`} onClick={() => setUploadMsg(null)}>
          {uploadMsg.text}
        </div>
      )}

      <main className="main">
        {/* 초기 로딩 */}
        {step === "loading" && (
          <div className="card center-card">
            <div className="spinner" />
            <p className="desc">마스터 품목 불러오는 중…</p>
          </div>
        )}

        {/* 발주 파일 업로드 */}
        {step === "ready" && (
          <div className="card center-card">
            <div className="step-icon">📂</div>
            <h2>발주 파일 업로드</h2>
            <p className="desc">쿠팡에서 받은 발주 엑셀을 올려주세요</p>
            {masterError && !masterLoaded && <div className="error-msg">{masterError}</div>}
            {!masterLoaded ? (
              <div className="no-master">
                마스터가 없어요.<br />
                우측 상단 <strong>📋 마스터 업데이트</strong> 버튼으로<br />
                플로엠리스.xlsx를 먼저 올려주세요.
              </div>
            ) : (
              <div
                className="upload-zone"
                onDragOver={(e) => { e.preventDefault(); e.currentTarget.classList.add("drag-over"); }}
                onDragLeave={(e) => { e.currentTarget.classList.remove("drag-over"); }}
                onDrop={(e) => {
                  e.preventDefault();
                  e.currentTarget.classList.remove("drag-over");
                  const file = e.dataTransfer.files[0];
                  if (file) handleOrderFile({ target: { files: [file], value: "" } });
                }}
                onClick={() => document.getElementById("order-file-input").click()}
              >
                <input id="order-file-input" type="file" accept=".xlsx,.xls" onChange={handleOrderFile} hidden />
                <div className="upload-icon">⬆</div>
                <div className="upload-text">{loading ? "처리 중…" : "파일을 클릭하거나 끌어다 놓으세요"}</div>
                <div className="upload-sub">.xlsx / .xls</div>
              </div>
            )}
          </div>
        )}

        {/* 결과 */}
        {step === "result" && processed && (
          <div className="result-wrap">
            <div className="stats-row">
              <div className="stat-card">
                <div className="stat-num">{stats.total}</div>
                <div className="stat-label">전체 품목</div>
              </div>
              <div className="stat-card matched">
                <div className="stat-num">{stats.matched}</div>
                <div className="stat-label">매칭 완료</div>
              </div>
              <div className="stat-card unmatched">
                <div className="stat-num">{stats.unmatched}</div>
                <div className="stat-label">미매칭</div>
              </div>
            </div>

            <div className="action-row">
              <button className="btn-download" onClick={() => buildExcel(processed)}>
                ⬇ 엑셀 다운로드 (2시트)
              </button>
              <button className="btn-ghost" onClick={handleReset}>새 파일 처리</button>
            </div>

            {/* 정렬결과 미리보기 */}
            <div className="preview">
              <div className="preview-header">정렬결과</div>
              <div className="table-wrap">
                <table>
                  <thead>
                    <tr><th>#</th><th>품목코드</th><th>SKU 이름</th><th>발주수량</th><th>개입수</th><th>합계</th><th>매칭</th></tr>
                  </thead>
                  <tbody>
                    {(() => {
                      let cg = -1, rn = 0;
                      return processed.map((item, idx) => {
                        const rows = [];
                        if (item.group !== cg) {
                          cg = item.group;
                          rows.push(
                            <tr key={`g${idx}`} className="group-hdr" style={{ background: gc(item.group).bg }}>
                              <td colSpan={7}>▶ {GROUP_LABELS[item.group]}</td>
                            </tr>
                          );
                        }
                        rn++;
                        const col = gc(item.group);
                        rows.push(
                          <tr key={idx} style={{ background: rn%2===0 ? col.alt : col.light }}>
                            <td className="td-num">{rn}</td>
                            <td className="td-code">{item.code || "—"}</td>
                            <td className="td-name">{String(item.row[2] || "")}</td>
                            <td className="td-center">{item.row[4] ?? "—"}</td>
                            <td className="td-center">{item.packQty ?? "—"}</td>
                            <td className="td-center bold">{item.total ?? "—"}</td>
                            <td className="td-center">
                              <span className={item.master ? "badge-ok" : "badge-ng"}>
                                {item.master ? "✓" : "✗"}
                              </span>
                            </td>
                          </tr>
                        );
                        return rows;
                      });
                    })()}
                  </tbody>
                </table>
              </div>
            </div>

            {/* 품목별 합계 미리보기 */}
            <div className="preview">
              <div className="preview-header">품목별 합계</div>
              <div className="table-wrap">
                <table>
                  <thead><tr><th>품목코드</th><th>발주수량 합계</th></tr></thead>
                  <tbody>
                    {(() => {
                      const codeMap = {};
                      const noQtyMap = {};
                      processed.forEach((item) => {
                        if (!item.code) return;
                        if (item.noQty) {
                          if (!noQtyMap[item.code]) noQtyMap[item.code] = { group: item.group, sortNum: item.sortNum };
                        } else {
                          if (!codeMap[item.code]) codeMap[item.code] = { group:item.group, sortNum:item.sortNum, qty:0 };
                          const q = item.total!=null ? item.total : (item.row[4]!=null ? Number(item.row[4]) : 0);
                          codeMap[item.code].qty += q;
                        }
                      });
                      const noQtyEntries = Object.entries(noQtyMap).sort(([,a],[,b]) => a.group-b.group||a.sortNum-b.sortNum);
                      const sorted = Object.entries(codeMap).sort(([,a],[,b]) => a.group-b.group||a.sortNum-b.sortNum);
                      const rows = [];
                      // 수량 미확인 먼저
                      if (noQtyEntries.length > 0) {
                        rows.push(<tr key="noqty-hdr" className="group-hdr" style={{background:"#B71C1C"}}><td colSpan={2}>⚠ 수량 미확인 — 직접 확인 필요</td></tr>);
                        noQtyEntries.forEach(([code], i) => rows.push(
                          <tr key={`nq${code}`} style={{background: i%2===0?"#FFEBEE":"#FFCDD2"}}>
                            <td className="td-code">{code}</td>
                            <td className="td-center" style={{color:"#B71C1C",fontWeight:700}}>확인필요</td>
                          </tr>
                        ));
                      }
                      // 일반 합계
                      let cg2=-999, ri=0;
                      sorted.forEach(([code, s]) => {
                        if (s.group !== cg2) {
                          cg2 = s.group;
                          rows.push(<tr key={`sg${code}`} className="group-hdr" style={{background:gc(s.group)?.bg||"#546E7A"}}><td colSpan={2}>▶ {GROUP_LABELS[s.group]}</td></tr>);
                        }
                        ri++;
                        rows.push(
                          <tr key={code} style={{background:ri%2===0?gc(s.group)?.alt:gc(s.group)?.light}}>
                            <td className="td-code">{code}</td>
                            <td className="td-center bold">{s.qty}</td>
                          </tr>
                        );
                      });
                      return rows;
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
