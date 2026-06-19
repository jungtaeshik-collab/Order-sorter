import React, { useState, useCallback } from "react";
import * as XLSX from "xlsx";
import "./App.css";

// ── 마스터 품목 리스트 (플로엠리스.xlsx 내용 — 업데이트 시 이 배열만 교체) ──
// 실제 배포 시: public/master.xlsx 파일을 fetch해서 로드
// 개발 편의상 아래 MASTER_URL 방식 사용
const MASTER_URL = "/master.xlsx";

// ── 유틸 함수 ──
function normalize(s) {
  return s.replace(/[-_\s()[\]]/g, "").toUpperCase();
}

function extractTokens(s) {
  const tokens = [];
  // 영문1-4자 + 숫자 시작 코드
  const re1 = /[A-Za-z]{1,4}[\d][\w\-.]*/g;
  let m;
  while ((m = re1.exec(s)) !== null) tokens.push(m[0]);
  // V-08 스타일
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
  // Pack_코드
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
          found = true;
          break;
        }
      }
      if (found) break;
    }
  });
  return best;
}

function getCodeFromSku(name) {
  if (!name) return null;
  const packM = name.match(/Pack_([A-Za-z]{1,4}[\d][\w\-.]*)/);
  if (packM) return packM[1];
  const codes = name.match(/[A-Za-z]{1,4}[\d][\w\-.]*/g);
  return codes ? codes[0] : null;
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
  return 99;
}

function getSortNum(code) {
  if (!code) return 999999;
  const m = code.match(/^[A-Za-z]+?(\d+)/);
  return m ? parseInt(m[1], 10) : 999999;
}

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

// 쌍 행 병합 (V1907 등 발주수량이 다음 행에 있는 케이스)
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

const GROUP_LABELS = {
  0: "F 시리즈",
  1: "L 시리즈",
  2: "V 시리즈",
  3: "FL 시리즈",
  99: "기타",
};

const GROUP_COLORS = {
  0: { bg: "#1565C0", light: "#E3F0FF", alt: "#CCE0FF" },
  1: { bg: "#2E7D32", light: "#E8F5E9", alt: "#D0EBD1" },
  2: { bg: "#F57F17", light: "#FFF8E1", alt: "#FDEFC3" },
  3: { bg: "#6A1B9A", light: "#F3E5F5", alt: "#E4C8EE" },
  99: { bg: "#546E7A", light: "#FAFAFA", alt: "#EFEFEF" },
};

// ── 엑셀 생성 ──
function buildExcel(processed) {
  const wb = XLSX.utils.book_new();
  const wsData = [];

  // 헤더
  wsData.push([
    "브랜드", "SKU ID", "SKU 이름", "품목코드", "SKU Barcode",
    "발주수량", "개입수", "합계수량", "확정수량", "입고수량",
    "매입가", "총발주 매입금", "발주번호", "발주유형", "발주현황",
    "물류센터", "입고예정일", "발주일", "매입유형", "면세여부",
    "생산연도", "제조일자", "유통(소비)기한", "공급가", "부가세",
    "입고금액", "Xdock",
  ]);

  let curGroup = -1;
  processed.forEach((item) => {
    if (item.group !== curGroup) {
      curGroup = item.group;
      const cnt = processed.filter((x) => x.group === item.group).length;
      wsData.push([`▶  ${GROUP_LABELS[item.group]} — 숫자 오름차순  (${cnt}개)`]);
    }
    const r = item.row;
    wsData.push([
      r[0], r[1], r[2], item.code, r[3], r[4],
      item.packQty, item.total,
      r[5], r[6], r[7], r[8], r[9], r[10], r[11],
      r[12], r[13], r[14], r[15], r[16], r[17], r[18],
      r[19], r[20], r[21], r[22], r[23],
    ]);
  });

  const ws = XLSX.utils.aoa_to_sheet(wsData);
  ws["!cols"] = [
    { wch: 14 }, { wch: 12 }, { wch: 50 }, { wch: 14 }, { wch: 16 },
    { wch: 8 }, { wch: 8 }, { wch: 10 }, { wch: 8 }, { wch: 8 },
    { wch: 10 }, { wch: 14 }, { wch: 12 }, { wch: 8 }, { wch: 12 },
    { wch: 8 }, { wch: 12 }, { wch: 12 }, { wch: 8 }, { wch: 8 },
    { wch: 10 }, { wch: 10 }, { wch: 12 }, { wch: 8 }, { wch: 8 },
    { wch: 10 }, { wch: 8 },
  ];
  XLSX.utils.book_append_sheet(wb, ws, "정렬결과");
  XLSX.writeFile(wb, "발주정렬결과.xlsx");
}

// ── 메인 앱 ──
export default function App() {
  const [masterCodes, setMasterCodes] = useState(null);
  const [masterLoaded, setMasterLoaded] = useState(false);
  const [masterError, setMasterError] = useState(null);
  const [orderFile, setOrderFile] = useState(null);
  const [processed, setProcessed] = useState(null);
  const [stats, setStats] = useState(null);
  const [loading, setLoading] = useState(false);
  const [step, setStep] = useState(1); // 1=마스터로드 2=발주업로드 3=결과

  // 마스터 로드 (public/master.xlsx)
  const loadMaster = useCallback(async () => {
    setLoading(true);
    setMasterError(null);
    try {
      const res = await fetch(MASTER_URL);
      if (!res.ok) throw new Error("master.xlsx 파일을 찾을 수 없어요");
      const buf = await res.arrayBuffer();
      const wb = XLSX.read(buf, { type: "array" });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(ws, { header: 1 });
      const codes = rows.slice(1)
        .map((r) => (r[0] != null ? String(r[0]).trim() : ""))
        .filter(Boolean);
      setMasterCodes(codes);
      setMasterLoaded(true);
      setStep(2);
    } catch (e) {
      setMasterError(e.message);
    }
    setLoading(false);
  }, []);

  // 마스터 파일 직접 업로드 (로컬 테스트용)
  const handleMasterUpload = useCallback((e) => {
    const file = e.target.files[0];
    if (!file) return;
    setLoading(true);
    const reader = new FileReader();
    reader.onload = (ev) => {
      try {
        const wb = XLSX.read(ev.target.result, { type: "array" });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(ws, { header: 1 });
        const codes = rows.slice(1)
          .map((r) => (r[0] != null ? String(r[0]).trim() : ""))
          .filter(Boolean);
        setMasterCodes(codes);
        setMasterLoaded(true);
        setStep(2);
      } catch (err) {
        setMasterError("파일 읽기 실패: " + err.message);
      }
      setLoading(false);
    };
    reader.readAsArrayBuffer(file);
  }, []);

  // 발주 파일 처리
  const handleOrderFile = useCallback((e) => {
    const file = e.target.files[0];
    if (!file) return;
    setOrderFile(file.name);
    setLoading(true);

    const reader = new FileReader();
    reader.onload = (ev) => {
      try {
        const wb = XLSX.read(ev.target.result, { type: "array" });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const allRows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
        const raw = allRows.slice(1);

        const merged = mergeRows(raw);
        const index = buildMasterIndex(masterCodes);

        const result = [];
        merged.forEach((row) => {
          const nameRaw = row[2];
          const nameStr = nameRaw != null ? String(nameRaw).trim() : "";
          if (!nameStr || /^\d+$/.test(nameStr)) return;

          let match = findRankExact(nameStr, index);
          let method = "코드";
          if (!match) {
            const fuzzy = findRankFuzzy(nameStr, masterCodes);
            if (fuzzy) { match = fuzzy; method = "유사"; }
          }

          const code = getCodeFromSku(nameStr);
          const group = getGroup(code);
          const sortNum = getSortNum(code);
          const packQty = extractPackQty(nameStr);
          const eVal = row[4];
          const eNum = eVal != null && eVal !== "" ? Number(eVal) : null;
          const total = eNum != null && packQty != null ? eNum * packQty : null;

          result.push({
            group, sortNum, code,
            master: match?.mc || null,
            method,
            packQty, total,
            row: [...row],
          });
        });

        result.sort((a, b) => a.group - b.group || a.sortNum - b.sortNum);

        const matched = result.filter((x) => x.master).length;
        setStats({ total: result.length, matched, unmatched: result.length - matched });
        setProcessed(result);
        setStep(3);
      } catch (err) {
        alert("파일 처리 오류: " + err.message);
      }
      setLoading(false);
    };
    reader.readAsArrayBuffer(file);
  }, [masterCodes]);

  const handleDownload = () => buildExcel(processed);

  const handleReset = () => {
    setOrderFile(null);
    setProcessed(null);
    setStats(null);
    setStep(2);
  };

  return (
    <div className="app">
      <header className="header">
        <div className="header-inner">
          <div className="logo">
            <span className="logo-mark">F</span>
            <span className="logo-text">플로엠 발주 정렬기</span>
          </div>
          {masterLoaded && (
            <div className="master-badge">
              ✓ 마스터 {masterCodes.length.toLocaleString()}개 로드됨
            </div>
          )}
        </div>
      </header>

      <main className="main">
        {/* STEP 1: 마스터 로드 */}
        {step === 1 && (
          <div className="card center-card">
            <div className="step-icon">📋</div>
            <h2>시작하기</h2>
            <p className="desc">먼저 품목 마스터를 불러와야 해요</p>

            <button className="btn-primary" onClick={loadMaster} disabled={loading}>
              {loading ? "로딩 중…" : "마스터 자동 로드 (플로엠리스.xlsx)"}
            </button>

            <div className="divider"><span>또는</span></div>

            <label className="btn-secondary file-label">
              마스터 파일 직접 선택
              <input type="file" accept=".xlsx" onChange={handleMasterUpload} hidden />
            </label>

            {masterError && <div className="error-msg">⚠ {masterError}</div>}
          </div>
        )}

        {/* STEP 2: 발주 파일 업로드 */}
        {step === 2 && (
          <div className="card center-card">
            <div className="step-icon">📂</div>
            <h2>발주 파일 업로드</h2>
            <p className="desc">쿠팡에서 받은 발주 엑셀 파일을 올려주세요</p>

            <label className="upload-zone">
              <input type="file" accept=".xlsx,.xls" onChange={handleOrderFile} hidden />
              <div className="upload-icon">⬆</div>
              <div className="upload-text">
                {loading ? "처리 중…" : "파일을 클릭하거나 끌어다 놓으세요"}
              </div>
              <div className="upload-sub">.xlsx / .xls</div>
            </label>

            <button className="btn-ghost" onClick={() => { setMasterLoaded(false); setStep(1); }}>
              ← 마스터 다시 로드
            </button>
          </div>
        )}

        {/* STEP 3: 결과 */}
        {step === 3 && processed && (
          <div className="result-wrap">
            {/* 통계 카드 */}
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

            {/* 액션 버튼 */}
            <div className="action-row">
              <button className="btn-download" onClick={handleDownload}>
                ⬇ 엑셀 다운로드
              </button>
              <button className="btn-ghost" onClick={handleReset}>
                새 파일 처리
              </button>
            </div>

            {/* 미리보기 테이블 */}
            <div className="preview">
              <div className="preview-header">미리보기</div>
              <div className="table-wrap">
                <table>
                  <thead>
                    <tr>
                      <th>#</th>
                      <th>품목코드</th>
                      <th>SKU 이름</th>
                      <th>발주수량</th>
                      <th>개입수</th>
                      <th>합계</th>
                      <th>매칭</th>
                    </tr>
                  </thead>
                  <tbody>
                    {(() => {
                      let curGroup = -1;
                      let rowNum = 0;
                      return processed.map((item, idx) => {
                        const rows = [];
                        if (item.group !== curGroup) {
                          curGroup = item.group;
                          const col = GROUP_COLORS[item.group];
                          rows.push(
                            <tr key={`g${idx}`} className="group-hdr" style={{ background: col.bg }}>
                              <td colSpan={7}>▶ {GROUP_LABELS[item.group]}</td>
                            </tr>
                          );
                        }
                        rowNum++;
                        const col = GROUP_COLORS[item.group];
                        const bg = rowNum % 2 === 0 ? col.alt : col.light;
                        rows.push(
                          <tr key={idx} style={{ background: bg }}>
                            <td className="td-num">{rowNum}</td>
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
          </div>
        )}
      </main>
    </div>
  );
}
