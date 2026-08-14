import React, { useState, useEffect, useCallback } from "react";
import {
  Calendar, MapPin, Clock, Users, CheckCircle2, ChevronRight,
  TrendingUp, FileText, Phone, ArrowLeft, RefreshCw, MessageSquare, Sparkles, UserCheck, Lock, Download
} from "lucide-react";
import {
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer, Cell
} from "recharts";
import * as XLSX from "xlsx";

/* ---------- Brand tokens (Mirae Asset CI) ---------- */
const ORANGE = "#F58220";
const ORANGE_DARK = "#CB6015";
const NAVY = "#043B72";
const NAVY_SOFT = "#3A5A7A";
const CREAM = "#FBF9F6";
const LINE = "#E7E2DA";
const INK = "#1E2328";

const TOPIC_OPTIONS = [
  "DB · DC 제도 비교",
  "기금형 퇴직연금 도입",
  "의무가입 대비 전략",
  "세제 · 법률 이슈",
  "IRP · 개인연금 활용",
  "기타",
];

const JOB_OPTIONS = ["인사", "재무 · 회계", "경영지원", "기타"];
const PENSION_ADOPTION_OPTIONS = ["도입", "미도입"];
const INTENT_OPTIONS = ["매우 있음", "있음", "보통", "없음"];

const RATING_FIELDS = [
  { key: "satisfaction", label: "세미나 전반적인 만족도" },
  { key: "usefulness", label: "강의 내용의 실무 유용성" },
  { key: "delivery", label: "강사의 전달력 · 이해도" },
];

// ── Google Sheets backend ──────────────────────────────────────────────
// 1) Google Sheets에 "설문"과 "출석" 탭(시트)을 만드세요.
// 2) 함께 드린 apps-script.gs 코드를 Apps Script에 붙여넣고 "웹 앱"으로 배포하세요.
// 3) 배포 후 발급되는 웹 앱 URL을 아래 GAS_URL에 붙여넣으세요.
const GAS_URL = "https://docs.google.com/spreadsheets/d/1sGpgLTyDIc-ogp4yLWONan-RQrM1k0nK2qVNf2QlWO0/edit?usp=sharing";
const SURVEY_SHEET = "설문";
const ATTENDANCE_SHEET = "출석";
// Simple access code so ordinary attendees can't open the aggregated results.
// Change this before sharing the link, and share the code only with organizers.
const RESULTS_PIN = "2026";
const MIRAE_LOGO_SRC = "data:image/png;base64,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";

function uid() {
  return `${Date.now()}-${Math.random().toString(36).slice(2, 9)}`;
}

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function downloadCSV(filename, headers, rows) {
  const escapeCell = (v) => {
    const s = v === null || v === undefined ? "" : String(v);
    return /[",\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
  };
  const lines = [headers.map(escapeCell).join(",")]
    .concat(rows.map((row) => row.map(escapeCell).join(",")));
  // BOM so Excel/Google Sheets read Korean text as UTF-8 correctly
  const csv = "\uFEFF" + lines.join("\n");
  const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

function downloadXLSX(filename, sheetName, headers, rows) {
  const data = [headers, ...rows.map((row) => row.map((v) => (v === null || v === undefined ? "" : v)))];
  const ws = XLSX.utils.aoa_to_sheet(data);
  // Reasonable column widths so Korean labels aren't truncated on open.
  ws["!cols"] = headers.map((h) => ({ wch: Math.max(String(h).length + 4, 12) }));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, sheetName);
  XLSX.writeFile(wb, filename);
}

// Google Apps Script Web App is occasionally slow to cold-start or hiccups.
// Retry a few times with a short backoff before surfacing an error.
async function submitEntry(sheet, data, attempts = 3) {
  let lastErr;
  for (let i = 0; i < attempts; i++) {
    try {
      const res = await fetch(GAS_URL, {
        method: "POST",
        headers: { "Content-Type": "text/plain;charset=utf-8" }, // avoids CORS preflight on Apps Script
        body: JSON.stringify({ sheet, data }),
      });
      if (!res.ok) throw new Error(`서버 응답 오류 (${res.status})`);
      const json = await res.json();
      if (json && json.ok) return json;
      throw new Error(json?.error || "저장 결과가 올바르지 않습니다.");
    } catch (e) {
      lastErr = e;
    }
    if (i < attempts - 1) await sleep(400 * (i + 1));
  }
  throw lastErr;
}

async function listEntries(sheet) {
  const res = await fetch(`${GAS_URL}?sheet=${encodeURIComponent(sheet)}`);
  if (!res.ok) throw new Error(`서버 응답 오류 (${res.status})`);
  const json = await res.json();
  if (!json || !json.ok) throw new Error(json?.error || "결과를 불러오지 못했습니다.");
  return json.rows || [];
}


/* ---------- Shared bits ---------- */
function HeroBanner() {
  return (
    <div style={{ maxWidth: 640, margin: "0 auto", padding: "18px 20px 0" }}>
      <div style={{
        position: "relative", overflow: "hidden", borderRadius: 24,
        padding: "30px 24px 26px",
        background: "linear-gradient(135deg, #FFCB8E 0%, #F58220 55%, #E8611F 100%)",
        textAlign: "center",
      }}>
        {/* decorative circles */}
        <div style={{
          position: "absolute", top: -46, left: -36, width: 130, height: 130,
          borderRadius: "50%", background: "rgba(255,255,255,0.16)",
        }} />
        <div style={{
          position: "absolute", bottom: -60, right: -30, width: 160, height: 160,
          borderRadius: "50%", background: "rgba(255,255,255,0.14)",
        }} />
        {/* sparkles */}
        <Sparkle style={{ top: 16, left: 22 }} />
        <Sparkle style={{ top: 20, right: 26 }} size={12} />
        <Sparkle style={{ bottom: 22, left: 30 }} size={10} />

        <div style={{ position: "relative", zIndex: 1 }}>
          <div style={{ marginBottom: 16 }}>
            <img
              src={MIRAE_LOGO_SRC}
              alt="MIRAE ASSET 미래에셋증권"
              style={{ height: 72, width: "auto", display: "inline-block" }}
            />
          </div>

          <div style={{
            display: "inline-flex", alignItems: "center", gap: 6, marginTop: 16,
            padding: "6px 14px", borderRadius: 999, background: "rgba(255,255,255,0.22)",
            fontSize: 12.5, fontWeight: 700, color: "#fff",
          }}>
            <Sparkles size={13} color="#fff" />
            2026년 퇴직연금제도 세미나
          </div>

          <h1 style={{
            color: "#fff", fontSize: 22, fontWeight: 800, lineHeight: 1.4,
            margin: "14px 0 0", letterSpacing: -0.2,
          }}>
            퇴직연금제도 변화에 대비하는<br />가장 확실한 시간
          </h1>

          <div style={{
            display: "inline-flex", alignItems: "center", gap: 6, marginTop: 16,
            padding: "7px 14px", borderRadius: 999, background: "rgba(255,255,255,0.9)",
            fontSize: 12.5, fontWeight: 700, color: ORANGE_DARK,
          }}>
            <Calendar size={13} color={ORANGE_DARK} />
            2026년 9월 10일(목)
          </div>
        </div>
      </div>
    </div>
  );
}

function Sparkle({ style, size = 14 }) {
  return (
    <div style={{
      position: "absolute", width: size, height: size,
      background: "rgba(255,255,255,0.55)",
      clipPath: "polygon(50% 0%, 61% 39%, 100% 50%, 61% 61%, 50% 100%, 39% 61%, 0% 50%, 39% 39%)",
      ...style,
    }} />
  );
}

function TopNav({ tab, setTab }) {
  const tabs = [
    { id: "info", label: "행사안내" },
    { id: "attendance", label: "출석" },
    { id: "survey", label: "설문" },
  ];
  return (
    <div style={{ position: "sticky", top: 0, zIndex: 20, background: CREAM }}>
      <div style={{ maxWidth: 640, margin: "0 auto", padding: "0 20px" }}>
        <div style={{ display: "flex", gap: 28, marginTop: 14, borderBottom: `1px solid ${LINE}` }}>
          {tabs.map((t) => (
            <button
              key={t.id}
              onClick={() => setTab(t.id)}
              style={{
                background: "none", border: "none", cursor: "pointer",
                padding: "10px 2px 12px", fontSize: 15.5,
                fontWeight: tab === t.id ? 700 : 500,
                color: tab === t.id ? INK : "#8A8378",
                borderBottom: tab === t.id ? `2px solid ${ORANGE}` : "2px solid transparent",
                fontFamily: "inherit",
              }}
            >
              {t.label}
            </button>
          ))}
        </div>
      </div>
    </div>
  );
}

/* ---------- Info tab ---------- */
function InfoTab() {
  const agenda = [
    { time: "14:00 – 14:10", title: "개회 및 인사말", desc: "연금컨설팅팀 제도컨설팅 파트" },
    { time: "14:10 – 14:50", title: "2026–2027 퇴직연금 제도 변화", desc: "의무가입 로드맵과 기업의 대비 방향" },
    { time: "14:50 – 15:30", title: "DB · DC · 기금형 비교", desc: "제도 선택 기준과 설계 시 유의점" },
    { time: "15:30 – 15:45", title: "휴식" },
    { time: "15:45 – 16:20", title: "세제 · 법률 이슈 Q&A", desc: "성과급 DC 편입, 노란우산공제 등 실무 쟁점" },
    { time: "16:20 – 17:00", title: "개별 상담", desc: "담당 컨설턴트와 1:1 질의응답" },
  ];

  const notes = [
    "사전 등록하신 담당자 성함으로 현장 접수해 주세요.",
    "회사별 제도 현황 자료를 지참하시면 상담이 한층 원활합니다.",
    "행사 중간 참여 · 조기 퇴장 모두 가능합니다.",
    "세미나 종료 후 자료집 PDF를 이메일로 발송해 드립니다.",
  ];

  return (
    <div style={{ maxWidth: 640, margin: "0 auto", padding: "24px 20px 60px" }}>
      <p style={{ fontSize: 14.5, color: "#5B564C", lineHeight: 1.7, margin: "0 0 22px" }}>
        퇴직연금 의무화와 기금형 도입을 앞둔 지금, 우리 회사의 퇴직연금 제도를
        어떻게 준비해야 할지 함께 짚어보는 자리입니다.
      </p>

      {/* Info card */}
      <div style={{
        background: "#fff", border: `1px solid ${LINE}`, borderRadius: 14,
        padding: "18px 20px", display: "flex", flexDirection: "column", gap: 12,
        marginBottom: 28,
      }}>
        <InfoRow icon={<Calendar size={17} color={ORANGE} />} text="2026년 9월 10일(목)" />
        <InfoRow icon={<Clock size={17} color={ORANGE} />} text="오후 2:00 – 5:00 (접수 1:30부터)" />
        <InfoRow icon={<MapPin size={17} color={ORANGE} />} text="미래에셋증권 센터원빌딩 3층 컨퍼런스홀" />
        <InfoRow icon={<Users size={17} color={ORANGE} />} text="기업 인사 · 재무 담당자 (선착순 80명)" />
      </div>

      {/* Agenda */}
      <SectionTitle icon={<FileText size={16} color={NAVY} />} title="프로그램" />
      <div style={{ marginBottom: 30 }}>
        {agenda.map((a, i) => (
          <div key={i} style={{
            display: "flex", gap: 16, padding: "13px 0",
            borderBottom: i < agenda.length - 1 ? `1px solid ${LINE}` : "none",
          }}>
            <div style={{ width: 92, flexShrink: 0, fontSize: 12.5, color: "#9A9385", paddingTop: 1 }}>
              {a.time}
            </div>
            <div>
              <div style={{ fontSize: 14.5, fontWeight: 700, color: INK }}>{a.title}</div>
              {a.desc && (
                <div style={{ fontSize: 12.8, color: "#8A8378", marginTop: 2 }}>{a.desc}</div>
              )}
            </div>
          </div>
        ))}
      </div>

      {/* Notes */}
      <SectionTitle icon={<Sparkles size={16} color={NAVY} />} title="참고 사항" />
      <div style={{ marginBottom: 30 }}>
        {notes.map((n, i) => (
          <div key={i} style={{ display: "flex", gap: 10, alignItems: "flex-start", padding: "7px 0" }}>
            <span style={{
              width: 18, height: 18, borderRadius: "50%", background: "#FCEADA",
              color: ORANGE_DARK, fontSize: 11, fontWeight: 800, flexShrink: 0,
              display: "flex", alignItems: "center", justifyContent: "center", marginTop: 1
            }}>
              {i + 1}
            </span>
            <span style={{ fontSize: 13.6, color: "#4A463D", lineHeight: 1.6 }}>{n}</span>
          </div>
        ))}
      </div>

      {/* Contact */}
      <div style={{
        background: NAVY, borderRadius: 14, padding: "18px 20px",
        display: "flex", alignItems: "center", gap: 12,
      }}>
        <Phone size={18} color="#fff" style={{ flexShrink: 0 }} />
        <div>
          <div style={{ color: "#fff", fontSize: 13.5, fontWeight: 700 }}>
            문의 : 행사스탭 또는 담당 RM에게 문의해주세요.
          </div>
        </div>
      </div>
    </div>
  );
}

function InfoRow({ icon, text }) {
  return (
    <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
      {icon}
      <span style={{ fontSize: 14, color: INK, fontWeight: 500 }}>{text}</span>
    </div>
  );
}

function SectionTitle({ icon, title }) {
  return (
    <div style={{ display: "flex", alignItems: "center", gap: 7, marginBottom: 12 }}>
      {icon}
      <span style={{ fontSize: 15, fontWeight: 800, color: INK }}>{title}</span>
    </div>
  );
}

/* ---------- Rating control ---------- */
function RatingRow({ label, value, onChange }) {
  return (
    <div style={{ marginBottom: 22 }}>
      <div style={{ fontSize: 14, fontWeight: 600, color: INK, marginBottom: 10 }}>{label}</div>
      <div style={{ display: "flex", gap: 8 }}>
        {[1, 2, 3, 4, 5].map((n) => (
          <button
            key={n}
            type="button"
            onClick={() => onChange(n)}
            style={{
              flex: 1, height: 42, borderRadius: 10, cursor: "pointer",
              border: value === n ? `1.5px solid ${ORANGE}` : `1px solid ${LINE}`,
              background: value === n ? "#FEF1E4" : "#fff",
              color: value === n ? ORANGE_DARK : "#9A9385",
              fontWeight: 700, fontSize: 14.5,
            }}
          >
            {n}
          </button>
        ))}
      </div>
      <div style={{ display: "flex", justifyContent: "space-between", marginTop: 5 }}>
        <span style={{ fontSize: 11, color: "#B0AA9C" }}>매우 불만족</span>
        <span style={{ fontSize: 11, color: "#B0AA9C" }}>매우 만족</span>
      </div>
    </div>
  );
}

function ChipGroup({ options, selected, onToggle, multi = true }) {
  return (
    <div style={{ display: "flex", flexWrap: "wrap", gap: 8 }}>
      {options.map((opt) => {
        const active = selected.includes(opt);
        return (
          <button
            key={opt}
            type="button"
            onClick={() => onToggle(opt)}
            style={{
              padding: "9px 14px", borderRadius: 999, cursor: "pointer",
              border: active ? `1.5px solid ${ORANGE}` : `1px solid ${LINE}`,
              background: active ? "#FEF1E4" : "#fff",
              color: active ? ORANGE_DARK : "#5B564C",
              fontSize: 13.3, fontWeight: 600,
            }}
          >
            {opt}
          </button>
        );
      })}
    </div>
  );
}

/* ---------- Attendance tab ---------- */
function AttendanceTab({ onGoResults }) {
  const [name, setName] = useState("");
  const [company, setCompany] = useState("");
  const [department, setDepartment] = useState("");
  const [parking, setParking] = useState("");
  const [submitting, setSubmitting] = useState(false);
  const [checkedInAt, setCheckedInAt] = useState(null);
  const [error, setError] = useState("");

  const canSubmit = name.trim().length > 0 && !submitting;

  const handleCheckIn = async () => {
    if (!canSubmit) return;
    setSubmitting(true);
    setError("");
    const now = new Date();
    const entry = {
      id: uid(),
      checkedInAt: now.toISOString(),
      name: name.trim(),
      company: company.trim(),
      department: department.trim(),
      parking: parking.trim(),
    };
    try {
      await submitEntry(ATTENDANCE_SHEET, entry);
      setCheckedInAt(now);
    } catch (e) {
      console.error("check-in failed", e);
      setError(`출석 체크 중 문제가 발생했습니다 (${e?.message || "알 수 없는 오류"}). 잠시 후 다시 시도해 주세요.`);
    } finally {
      setSubmitting(false);
    }
  };

  if (checkedInAt) {
    return (
      <div style={{
        maxWidth: 640, margin: "0 auto", padding: "70px 24px",
        display: "flex", flexDirection: "column", alignItems: "center", textAlign: "center"
      }}>
        <div style={{
          width: 60, height: 60, borderRadius: "50%", background: "#FEF1E4",
          display: "flex", alignItems: "center", justifyContent: "center", marginBottom: 20
        }}>
          <UserCheck size={28} color={ORANGE} />
        </div>
        <h2 style={{ fontSize: 20, fontWeight: 800, color: INK, margin: "0 0 8px" }}>
          출석이 확인되었습니다
        </h2>
        <p style={{ fontSize: 13.8, color: "#8A8378", lineHeight: 1.6, margin: 0 }}>
          {name.trim()}님, {checkedInAt.toLocaleTimeString("ko-KR", { hour: "2-digit", minute: "2-digit" })}에 체크인되었습니다.
        </p>
        <button
          onClick={onGoResults}
          style={{
            marginTop: 26, background: "none", border: "none", cursor: "pointer",
            color: NAVY_SOFT, fontSize: 12.5, textDecoration: "underline"
          }}
        >
          주최자이신가요? 응답 결과 보기
        </button>
      </div>
    );
  }

  return (
    <div style={{ maxWidth: 640, margin: "0 auto", padding: "28px 20px 70px" }}>
      <h1 style={{ fontSize: 21, fontWeight: 800, color: INK, margin: "0 0 6px" }}>
        출석 체크
      </h1>
      <p style={{ fontSize: 13.5, color: "#8A8378", margin: "0 0 26px" }}>
        성함을 입력하고 출석 체크 버튼을 눌러 주세요.
      </p>

      <Field label="성함">
        <input
          value={name}
          onChange={(e) => setName(e.target.value)}
          placeholder="홍길동"
          style={inputStyle}
        />
      </Field>

      <Field label="소속 회사명">
        <input
          value={company}
          onChange={(e) => setCompany(e.target.value)}
          placeholder="예: OO전자"
          style={inputStyle}
        />
      </Field>

      <Field label="부서">
        <input
          value={department}
          onChange={(e) => setDepartment(e.target.value)}
          placeholder="예: 인사팀"
          style={inputStyle}
        />
      </Field>

      <Field label="주차 등록 (선택)">
        <input
          value={parking}
          onChange={(e) => setParking(e.target.value)}
          placeholder="차량 번호를 입력해 주세요 (예: 12가 3456)"
          style={inputStyle}
        />
      </Field>

      {error && <div style={{ color: "#C0392B", fontSize: 12.5, marginBottom: 12 }}>{error}</div>}

      <button
        onClick={handleCheckIn}
        disabled={!canSubmit}
        style={{
          width: "100%", padding: "15px 0", borderRadius: 12, border: "none",
          background: canSubmit ? ORANGE : "#EBE6DD",
          color: canSubmit ? "#fff" : "#B0AA9C",
          fontSize: 15.5, fontWeight: 700, cursor: canSubmit ? "pointer" : "not-allowed",
          marginTop: 8,
        }}
      >
        {submitting ? "체크 중…" : "출석 체크하기"}
      </button>
    </div>
  );
}

/* ---------- Survey tab ---------- */
function SurveyTab({ onGoResults }) {
  const [job, setJob] = useState("");
  const [pensionAdopted, setPensionAdopted] = useState("");
  const [ratings, setRatings] = useState({});
  const [topics, setTopics] = useState([]);
  const [intent, setIntent] = useState("");
  const [comment, setComment] = useState("");
  const [submitting, setSubmitting] = useState(false);
  const [submitted, setSubmitted] = useState(false);
  const [error, setError] = useState("");

  const toggleTopic = (t) =>
    setTopics((prev) => (prev.includes(t) ? prev.filter((x) => x !== t) : [...prev, t]));

  const allRated = RATING_FIELDS.every((f) => ratings[f.key]);
  const canSubmit = allRated && job && pensionAdopted && intent && !submitting;

  const handleSubmit = async () => {
    if (!canSubmit) return;
    setSubmitting(true);
    setError("");
    const entry = {
      id: uid(),
      submittedAt: new Date().toISOString(),
      job,
      pensionAdopted,
      ratings,
      topics,
      intent,
      comment: comment.trim(),
    };
    try {
      await submitEntry(SURVEY_SHEET, entry);
      setSubmitted(true);
    } catch (e) {
      console.error("submit failed", e);
      setError(`제출 중 문제가 발생했습니다 (${e?.message || "알 수 없는 오류"}). 잠시 후 다시 시도해 주세요.`);
    } finally {
      setSubmitting(false);
    }
  };

  if (submitted) {
    return (
      <div style={{
        maxWidth: 640, margin: "0 auto", padding: "70px 24px",
        display: "flex", flexDirection: "column", alignItems: "center", textAlign: "center"
      }}>
        <div style={{
          width: 60, height: 60, borderRadius: "50%", background: "#FEF1E4",
          display: "flex", alignItems: "center", justifyContent: "center", marginBottom: 20
        }}>
          <CheckCircle2 size={30} color={ORANGE} />
        </div>
        <h2 style={{ fontSize: 20, fontWeight: 800, color: INK, margin: "0 0 8px" }}>
          설문에 참여해 주셔서 감사합니다
        </h2>
        <p style={{ fontSize: 13.8, color: "#8A8378", lineHeight: 1.6, margin: 0 }}>
          보내주신 의견은 다음 세미나 준비에 소중히 반영하겠습니다.
        </p>
        <button
          onClick={onGoResults}
          style={{
            marginTop: 26, background: "none", border: "none", cursor: "pointer",
            color: NAVY_SOFT, fontSize: 12.5, textDecoration: "underline"
          }}
        >
          주최자이신가요? 응답 결과 보기
        </button>
      </div>
    );
  }

  return (
    <div style={{ maxWidth: 640, margin: "0 auto", padding: "28px 20px 70px" }}>
      <h1 style={{ fontSize: 21, fontWeight: 800, color: INK, margin: "0 0 6px" }}>
        세미나 만족도 조사
      </h1>
      <p style={{ fontSize: 13.5, color: "#8A8378", margin: "0 0 26px" }}>
        1분이면 충분합니다. 솔직한 의견을 남겨 주세요.
      </p>

      <Field label="담당 직무분야">
        <ChipGroup options={JOB_OPTIONS} selected={job ? [job] : []} onToggle={(v) => setJob(v)} />
      </Field>

      <Field label="회사의 퇴직연금제도 도입 여부">
        <ChipGroup
          options={PENSION_ADOPTION_OPTIONS}
          selected={pensionAdopted ? [pensionAdopted] : []}
          onToggle={(v) => setPensionAdopted(v)}
        />
      </Field>

      <div style={{ height: 6 }} />

      {RATING_FIELDS.map((f) => (
        <RatingRow
          key={f.key}
          label={f.label}
          value={ratings[f.key]}
          onChange={(n) => setRatings((r) => ({ ...r, [f.key]: n }))}
        />
      ))}

      <Field label="향후 관심 있는 주제 (복수 선택 가능)">
        <ChipGroup options={TOPIC_OPTIONS} selected={topics} onToggle={toggleTopic} />
      </Field>

      <Field label="향후 세미나 · 컨설팅 참여 의향">
        <ChipGroup options={INTENT_OPTIONS} selected={intent ? [intent] : []} onToggle={(v) => setIntent(v)} />
      </Field>

      <Field label="자유 의견 (선택)">
        <textarea
          value={comment}
          onChange={(e) => setComment(e.target.value)}
          placeholder="세미나에 대한 의견이나 다음 세미나에서 다뤘으면 하는 주제를 자유롭게 적어 주세요."
          rows={4}
          style={{ ...inputStyle, resize: "vertical", fontFamily: "inherit" }}
        />
      </Field>

      {error && <div style={{ color: "#C0392B", fontSize: 12.5, marginBottom: 12 }}>{error}</div>}

      <button
        onClick={handleSubmit}
        disabled={!canSubmit}
        style={{
          width: "100%", padding: "15px 0", borderRadius: 12, border: "none",
          background: canSubmit ? ORANGE : "#EBE6DD",
          color: canSubmit ? "#fff" : "#B0AA9C",
          fontSize: 15.5, fontWeight: 700, cursor: canSubmit ? "pointer" : "not-allowed",
          marginTop: 8,
        }}
      >
        {submitting ? "제출 중…" : "제출하기"}
      </button>

      <button
        onClick={onGoResults}
        style={{
          display: "block", margin: "18px auto 0", background: "none", border: "none",
          cursor: "pointer", color: "#B0AA9C", fontSize: 12, textDecoration: "underline"
        }}
      >
        주최자이신가요? 응답 결과 보기
      </button>
    </div>
  );
}

function Field({ label, children }) {
  return (
    <div style={{ marginBottom: 22 }}>
      <div style={{ fontSize: 14, fontWeight: 600, color: INK, marginBottom: 10 }}>{label}</div>
      {children}
    </div>
  );
}

const inputStyle = {
  width: "100%", boxSizing: "border-box", padding: "12px 14px", borderRadius: 10,
  border: `1px solid ${LINE}`, fontSize: 14, color: INK, outline: "none",
};

/* ---------- Results tab ---------- */
function ResultsTab({ onBack }) {
  const [locked, setLocked] = useState(true);
  const [pin, setPin] = useState("");
  const [pinError, setPinError] = useState("");
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState("");
  const [entries, setEntries] = useState([]);
  const [attendees, setAttendees] = useState([]);
  const [expandedIds, setExpandedIds] = useState(() => new Set());

  const toggleExpanded = (id) => {
    setExpandedIds((prev) => {
      const next = new Set(prev);
      if (next.has(id)) next.delete(id); else next.add(id);
      return next;
    });
  };

  const load = useCallback(async () => {
    setLoading(true);
    setError("");
    try {
      const [surveyResults, attendanceResults] = await Promise.all([
        listEntries(SURVEY_SHEET),
        listEntries(ATTENDANCE_SHEET),
      ]);
      // ratings/topics are stored as JSON strings in the sheet; parse them back.
      surveyResults.forEach((e) => {
        if (typeof e.ratings === "string") { try { e.ratings = JSON.parse(e.ratings); } catch (_) { e.ratings = {}; } }
        if (typeof e.topics === "string") { try { e.topics = JSON.parse(e.topics); } catch (_) { e.topics = []; } }
      });
      surveyResults.sort((a, b) => new Date(b.submittedAt) - new Date(a.submittedAt));
      attendanceResults.sort((a, b) => new Date(b.checkedInAt) - new Date(a.checkedInAt));
      setEntries(surveyResults);
      setAttendees(attendanceResults);
    } catch (e) {
      setError("결과를 불러오지 못했습니다.");
    } finally {
      setLoading(false);
    }
  }, []);

  useEffect(() => {
    if (!locked) load();
  }, [locked, load]);

  const handleUnlock = () => {
    if (pin.trim() === RESULTS_PIN) {
      setPinError("");
      setLocked(false);
    } else {
      setPinError("코드가 일치하지 않습니다.");
    }
  };

  if (locked) {
    return (
      <div style={{
        maxWidth: 640, margin: "0 auto", padding: "70px 24px",
        display: "flex", flexDirection: "column", alignItems: "center", textAlign: "center"
      }}>
        <div style={{
          width: 52, height: 52, borderRadius: "50%", background: "#FEF1E4",
          display: "flex", alignItems: "center", justifyContent: "center", marginBottom: 18
        }}>
          <Lock size={22} color={ORANGE} />
        </div>
        <h2 style={{ fontSize: 18, fontWeight: 800, color: INK, margin: "0 0 6px" }}>
          주최자 전용 화면입니다
        </h2>
        <p style={{ fontSize: 13, color: "#8A8378", margin: "0 0 20px", lineHeight: 1.6 }}>
          접근 코드를 입력하면 전체 참여자의 응답 결과를 볼 수 있습니다.
        </p>
        <div style={{ width: "100%", maxWidth: 220 }}>
          <input
            type="password"
            inputMode="numeric"
            value={pin}
            onChange={(e) => { setPin(e.target.value); setPinError(""); }}
            onKeyDown={(e) => { if (e.key === "Enter") handleUnlock(); }}
            placeholder="접근 코드"
            style={{ ...inputStyle, textAlign: "center", letterSpacing: 2 }}
          />
          {pinError && <div style={{ color: "#C0392B", fontSize: 12, marginTop: 8 }}>{pinError}</div>}
          <button
            onClick={handleUnlock}
            style={{
              width: "100%", marginTop: 12, padding: "13px 0", borderRadius: 10, border: "none",
              background: ORANGE, color: "#fff", fontSize: 14.5, fontWeight: 700, cursor: "pointer",
            }}
          >
            확인
          </button>
        </div>
        <button
          onClick={onBack}
          style={{
            marginTop: 22, background: "none", border: "none", cursor: "pointer",
            color: "#B0AA9C", fontSize: 12.5, textDecoration: "underline"
          }}
        >
          돌아가기
        </button>
      </div>
    );
  }


  const avg = (key) => {
    const vals = entries.map((e) => e.ratings?.[key]).filter((v) => typeof v === "number");
    if (!vals.length) return 0;
    return vals.reduce((a, b) => a + b, 0) / vals.length;
  };

  const topicCounts = TOPIC_OPTIONS.map((t) => ({
    name: t,
    count: entries.filter((e) => (e.topics || []).includes(t)).length,
  }));

  const comments = entries.filter((e) => e.comment && e.comment.trim().length > 0);

  return (
    <div style={{ maxWidth: 640, margin: "0 auto", padding: "24px 20px 70px" }}>
      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 22 }}>
        <button
          onClick={onBack}
          style={{ background: "none", border: "none", cursor: "pointer", display: "flex", alignItems: "center", gap: 5, color: NAVY_SOFT, fontSize: 13, fontWeight: 600 }}
        >
          <ArrowLeft size={15} /> 돌아가기
        </button>
        <button
          onClick={load}
          style={{ background: "none", border: "none", cursor: "pointer", display: "flex", alignItems: "center", gap: 5, color: "#8A8378", fontSize: 12.5 }}
        >
          <RefreshCw size={13} /> 새로고침
        </button>
      </div>

      <h1 style={{ fontSize: 20, fontWeight: 800, color: INK, margin: "0 0 4px" }}>설문 결과</h1>
      <p style={{ fontSize: 13, color: "#8A8378", margin: "0 0 16px" }}>
        전체 참여자의 응답을 실시간으로 취합합니다.
      </p>

      {!loading && !error && (attendees.length > 0 || entries.length > 0) && (
        <div style={{ display: "flex", gap: 8, marginBottom: 22 }}>
          <button
            onClick={() => downloadCSV(
              "출석자명단.csv",
              ["이름", "회사명", "부서", "주차등록", "체크인시각"],
              attendees.map((a) => [a.name, a.company, a.department, a.parking, a.checkedInAt])
            )}
            disabled={attendees.length === 0}
            style={{
              flex: 1, display: "flex", alignItems: "center", justifyContent: "center", gap: 6,
              padding: "10px 0", borderRadius: 10, border: `1px solid ${LINE}`, background: "#fff",
              color: attendees.length ? INK : "#C9C3B8", fontSize: 12.5, fontWeight: 700,
              cursor: attendees.length ? "pointer" : "not-allowed",
            }}
          >
            <Download size={13} /> 출석자 CSV
          </button>
          <button
            onClick={() => downloadCSV(
              "설문응답.csv",
              ["직무분야", "퇴직연금도입여부", "만족도", "실무유용성", "전달력", "관심주제", "참여의향", "자유의견", "제출시각"],
              entries.map((e) => [
                e.job, e.pensionAdopted,
                e.ratings?.satisfaction, e.ratings?.usefulness, e.ratings?.delivery,
                (e.topics || []).join(" / "), e.intent, e.comment, e.submittedAt,
              ])
            )}
            disabled={entries.length === 0}
            style={{
              flex: 1, display: "flex", alignItems: "center", justifyContent: "center", gap: 6,
              padding: "10px 0", borderRadius: 10, border: `1px solid ${LINE}`, background: "#fff",
              color: entries.length ? INK : "#C9C3B8", fontSize: 12.5, fontWeight: 700,
              cursor: entries.length ? "pointer" : "not-allowed",
            }}
          >
            <Download size={13} /> 설문 응답 CSV
          </button>
        </div>
      )}

      {!loading && !error && (attendees.length > 0 || entries.length > 0) && (
        <div style={{ display: "flex", gap: 8, marginBottom: 22, marginTop: -14 }}>
          <button
            onClick={() => downloadXLSX(
              "출석자명단.xlsx",
              "출석자",
              ["이름", "회사명", "부서", "주차등록", "체크인시각"],
              attendees.map((a) => [a.name, a.company, a.department, a.parking, a.checkedInAt])
            )}
            disabled={attendees.length === 0}
            style={{
              flex: 1, display: "flex", alignItems: "center", justifyContent: "center", gap: 6,
              padding: "10px 0", borderRadius: 10, border: `1px solid ${NAVY_SOFT}`, background: NAVY,
              color: attendees.length ? "#fff" : "#B9C4CE", fontSize: 12.5, fontWeight: 700,
              cursor: attendees.length ? "pointer" : "not-allowed",
            }}
          >
            <Download size={13} /> 출석자 엑셀(.xlsx)
          </button>
          <button
            onClick={() => downloadXLSX(
              "설문응답.xlsx",
              "설문응답",
              ["직무분야", "퇴직연금도입여부", "만족도", "실무유용성", "전달력", "관심주제", "참여의향", "자유의견", "제출시각"],
              entries.map((e) => [
                e.job, e.pensionAdopted,
                e.ratings?.satisfaction, e.ratings?.usefulness, e.ratings?.delivery,
                (e.topics || []).join(" / "), e.intent, e.comment, e.submittedAt,
              ])
            )}
            disabled={entries.length === 0}
            style={{
              flex: 1, display: "flex", alignItems: "center", justifyContent: "center", gap: 6,
              padding: "10px 0", borderRadius: 10, border: `1px solid ${NAVY_SOFT}`, background: NAVY,
              color: entries.length ? "#fff" : "#B9C4CE", fontSize: 12.5, fontWeight: 700,
              cursor: entries.length ? "pointer" : "not-allowed",
            }}
          >
            <Download size={13} /> 설문 응답 엑셀(.xlsx)
          </button>
        </div>
      )}

      {loading ? (
        <div style={{ padding: "40px 0", textAlign: "center", color: "#B0AA9C", fontSize: 13.5 }}>불러오는 중…</div>
      ) : error ? (
        <div style={{ padding: "40px 0", textAlign: "center", color: "#C0392B", fontSize: 13.5 }}>{error}</div>
      ) : entries.length === 0 && attendees.length === 0 ? (
        <div style={{ padding: "40px 0", textAlign: "center", color: "#B0AA9C", fontSize: 13.5 }}>
          아직 등록된 출석 또는 응답이 없습니다.
        </div>
      ) : (
        <>
          {/* Summary cards */}
          <div style={{ display: "grid", gridTemplateColumns: "repeat(2, 1fr)", gap: 10, marginBottom: 26 }}>
            <StatCard label="총 출석 인원" value={`${attendees.length}명`} icon={<UserCheck size={15} color={ORANGE} />} />
            <StatCard label="총 설문 응답 수" value={`${entries.length}명`} icon={<Users size={15} color={ORANGE} />} />
            <StatCard label="전반적 만족도" value={avg("satisfaction").toFixed(1)} icon={<TrendingUp size={15} color={ORANGE} />} />
            <StatCard label="실무 유용성" value={avg("usefulness").toFixed(1)} icon={<TrendingUp size={15} color={ORANGE} />} />
          </div>

          {/* Attendee list */}
          <SectionTitle icon={<UserCheck size={16} color={NAVY} />} title={`출석자 명단 (${attendees.length})`} />
          <div style={{
            marginBottom: 30, background: "#fff", border: `1px solid ${LINE}`, borderRadius: 14,
            maxHeight: 260, overflowY: "auto"
          }}>
            {attendees.length === 0 ? (
              <div style={{ padding: "16px", fontSize: 13, color: "#B0AA9C" }}>아직 출석 체크한 참여자가 없습니다.</div>
            ) : (
              attendees.map((a, i) => (
                <div key={a.id} style={{
                  display: "flex", justifyContent: "space-between", alignItems: "center",
                  padding: "12px 16px", borderBottom: i < attendees.length - 1 ? `1px solid ${LINE}` : "none",
                }}>
                  <div>
                    <div style={{ fontSize: 13.8, fontWeight: 700, color: INK }}>{a.name}</div>
                    <div style={{ fontSize: 11.5, color: "#8A8378", marginTop: 2 }}>
                      {[a.company, a.department].filter(Boolean).join(" · ") || "-"}
                      {a.parking && ` · 🚗 ${a.parking}`}
                    </div>
                  </div>
                  <div style={{ fontSize: 11.5, color: "#B0AA9C" }}>
                    {new Date(a.checkedInAt).toLocaleTimeString("ko-KR", { hour: "2-digit", minute: "2-digit" })}
                  </div>
                </div>
              ))
            )}
          </div>

          {/* Topic interest chart */}
          <SectionTitle icon={<FileText size={16} color={NAVY} />} title="관심 주제 분포" />
          <div style={{ height: 220, marginBottom: 30, background: "#fff", border: `1px solid ${LINE}`, borderRadius: 14, padding: "14px 10px 4px" }}>
            <ResponsiveContainer width="100%" height="100%">
              <BarChart data={topicCounts} margin={{ top: 4, right: 10, left: -18, bottom: 4 }}>
                <CartesianGrid strokeDasharray="3 3" stroke={LINE} vertical={false} />
                <XAxis dataKey="name" tick={{ fontSize: 10, fill: "#8A8378" }} interval={0} angle={-20} textAnchor="end" height={50} />
                <YAxis tick={{ fontSize: 11, fill: "#8A8378" }} allowDecimals={false} />
                <Tooltip contentStyle={{ fontSize: 12, borderRadius: 8, border: `1px solid ${LINE}` }} />
                <Bar dataKey="count" radius={[6, 6, 0, 0]}>
                  {topicCounts.map((_, i) => <Cell key={i} fill={ORANGE} />)}
                </Bar>
              </BarChart>
            </ResponsiveContainer>
          </div>

          {/* Individual survey responses */}
          <SectionTitle icon={<Users size={16} color={NAVY} />} title={`개별 설문 응답 (${entries.length})`} />
          <div style={{ display: "flex", flexDirection: "column", gap: 8, marginBottom: 30 }}>
            {entries.length === 0 ? (
              <div style={{ fontSize: 13, color: "#B0AA9C" }}>아직 제출된 설문이 없습니다.</div>
            ) : (
              entries.map((e) => (
                <ResponseCard key={e.id} entry={e} expanded={expandedIds.has(e.id)} onToggle={() => toggleExpanded(e.id)} />
              ))
            )}
          </div>

          {/* Comments */}
          <SectionTitle icon={<MessageSquare size={16} color={NAVY} />} title={`자유 의견 (${comments.length})`} />
          <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
            {comments.length === 0 && (
              <div style={{ fontSize: 13, color: "#B0AA9C" }}>등록된 자유 의견이 없습니다.</div>
            )}
            {comments.map((e) => (
              <div key={e.id} style={{ background: "#fff", border: `1px solid ${LINE}`, borderRadius: 12, padding: "13px 15px" }}>
                <div style={{ fontSize: 13.5, color: INK, lineHeight: 1.6 }}>{e.comment}</div>
                <div style={{ fontSize: 11, color: "#B0AA9C", marginTop: 8, display: "flex", gap: 8 }}>
                  <span>{e.job || "-"}</span>
                  <span>·</span>
                  <span>퇴직연금 {e.pensionAdopted || "-"}</span>
                </div>
              </div>
            ))}
          </div>
        </>
      )}
    </div>
  );
}

function ResponseCard({ entry, expanded, onToggle }) {
  const overallAvg =
    RATING_FIELDS.reduce((sum, f) => sum + (entry.ratings?.[f.key] || 0), 0) / RATING_FIELDS.length;

  return (
    <div style={{ background: "#fff", border: `1px solid ${LINE}`, borderRadius: 12, overflow: "hidden" }}>
      <button
        onClick={onToggle}
        style={{
          width: "100%", background: "none", border: "none", cursor: "pointer",
          display: "flex", justifyContent: "space-between", alignItems: "center",
          padding: "12px 15px", textAlign: "left",
        }}
      >
        <div>
          <div style={{ fontSize: 13.5, fontWeight: 700, color: INK }}>
            {entry.job || "직무 미기재"}
            <span style={{
              marginLeft: 8, fontSize: 11, fontWeight: 700, padding: "2px 8px", borderRadius: 999,
              background: entry.pensionAdopted === "도입" ? "#EAF3EC" : "#FBEDEA",
              color: entry.pensionAdopted === "도입" ? "#2E7D4F" : "#C0392B",
            }}>
              퇴직연금 {entry.pensionAdopted || "-"}
            </span>
          </div>
          <div style={{ fontSize: 11.3, color: "#8A8378", marginTop: 4 }}>
            {new Date(entry.submittedAt).toLocaleString("ko-KR", {
              month: "numeric", day: "numeric", hour: "2-digit", minute: "2-digit"
            })}
          </div>
        </div>
        <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
          <span style={{ fontSize: 12.5, fontWeight: 700, color: ORANGE_DARK }}>
            평균 {overallAvg.toFixed(1)}
          </span>
          <ChevronRight
            size={16}
            color="#B0AA9C"
            style={{ transform: expanded ? "rotate(90deg)" : "none", transition: "transform 0.15s" }}
          />
        </div>
      </button>

      {expanded && (
        <div style={{ padding: "0 15px 15px", borderTop: `1px solid ${LINE}` }}>
          <div style={{ display: "flex", flexDirection: "column", gap: 6, marginTop: 12 }}>
            {RATING_FIELDS.map((f) => (
              <div key={f.key} style={{ display: "flex", justifyContent: "space-between", fontSize: 12.5 }}>
                <span style={{ color: "#8A8378" }}>{f.label}</span>
                <span style={{ fontWeight: 700, color: INK }}>{entry.ratings?.[f.key] ?? "-"} / 5</span>
              </div>
            ))}
          </div>

          <div style={{ marginTop: 12 }}>
            <div style={{ fontSize: 11.5, color: "#8A8378", marginBottom: 6 }}>관심 주제</div>
            {entry.topics && entry.topics.length > 0 ? (
              <div style={{ display: "flex", flexWrap: "wrap", gap: 6 }}>
                {entry.topics.map((t) => (
                  <span key={t} style={{
                    fontSize: 11.5, padding: "4px 9px", borderRadius: 999,
                    background: "#FEF1E4", color: ORANGE_DARK, fontWeight: 600,
                  }}>{t}</span>
                ))}
              </div>
            ) : (
              <span style={{ fontSize: 12, color: "#B0AA9C" }}>선택 없음</span>
            )}
          </div>

          <div style={{ display: "flex", justifyContent: "space-between", fontSize: 12.5, marginTop: 12 }}>
            <span style={{ color: "#8A8378" }}>향후 참여 의향</span>
            <span style={{ fontWeight: 700, color: INK }}>{entry.intent || "-"}</span>
          </div>

          {entry.comment && (
            <div style={{ marginTop: 12 }}>
              <div style={{ fontSize: 11.5, color: "#8A8378", marginBottom: 4 }}>자유 의견</div>
              <div style={{ fontSize: 13, color: INK, lineHeight: 1.6 }}>{entry.comment}</div>
            </div>
          )}
        </div>
      )}
    </div>
  );
}

function StatCard({ label, value, icon }) {
  return (
    <div style={{ background: "#fff", border: `1px solid ${LINE}`, borderRadius: 12, padding: "14px 16px" }}>
      <div style={{ display: "flex", alignItems: "center", gap: 6, marginBottom: 8 }}>
        {icon}
        <span style={{ fontSize: 12, color: "#8A8378", fontWeight: 600 }}>{label}</span>
      </div>
      <div style={{ fontSize: 22, fontWeight: 800, color: INK }}>{value}</div>
    </div>
  );
}

/* ---------- App ---------- */
function App() {
  const [tab, setTab] = useState("info");

  return (
    <div style={{
      minHeight: "100vh", background: CREAM,
      fontFamily: "'Apple SD Gothic Neo','Malgun Gothic','Pretendard',-apple-system,sans-serif",
      color: INK,
    }}>
      {tab !== "results" && <HeroBanner />}
      {tab !== "results" && <TopNav tab={tab} setTab={setTab} />}
      {tab === "info" && <InfoTab />}
      {tab === "attendance" && <AttendanceTab onGoResults={() => setTab("results")} />}
      {tab === "survey" && <SurveyTab onGoResults={() => setTab("results")} />}
      {tab === "results" && <ResultsTab onBack={() => setTab("survey")} />}
    </div>
  );
}

/* ---------- Mount ---------- */
import { createRoot } from "react-dom/client";
createRoot(document.getElementById("root")).render(<App />);
