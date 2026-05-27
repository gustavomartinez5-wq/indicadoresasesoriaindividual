import { useState, useMemo, useCallback, useRef, useEffect } from "react";
import * as XLSX from "xlsx";
import {
  BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer,
  PieChart, Pie, Cell, LineChart, Line, CartesianGrid, Legend,
} from "recharts";
import { supabase } from "./supabaseClient";
import {
  ASESORES, ESCUELAS, SERVICIOS, SERVICIO_CLAVE, ESTATUSES,
  SEMESTRES, CAG_OPTS, EXATEC_OPTS, MODALIDADES, INTERESES,
  COMUNIDADES, STATUS_COLORS, CHART_COLORS,
} from "./constants";

/* ═══ TABS ═══ */
const TABS = [
  { id: "home",       icon: "◉", label: "Home" },
  { id: "asesorias",  icon: "◈", label: "Asesorías" },
  { id: "dashboard",  icon: "◐", label: "Dashboard" },
  { id: "asesores",   icon: "◎", label: "Asesores" },
  { id: "pipeline",   icon: "◑", label: "Pipeline" },
  { id: "alumnos",    icon: "◇", label: "Alumnos" },
];

/* ═══ HELPERS ═══ */
function norm(s) {
  return String(s ?? "").toLowerCase().normalize("NFD").replace(/[̀-ͯ]/g, "");
}
function useDebounce(value, delay = 200) {
  const [dv, setDv] = useState(value);
  useEffect(() => {
    const t = setTimeout(() => setDv(value), delay);
    return () => clearTimeout(t);
  }, [value, delay]);
  return dv;
}
function fmtDate(d) {
  if (!d) return "—";
  return new Date(d + "T12:00:00").toLocaleDateString("es-MX", { day: "2-digit", month: "short", year: "numeric" });
}
function todayISO() {
  return new Date().toISOString().slice(0, 10);
}
function countBy(arr, key) {
  const map = {};
  arr.forEach((r) => { const v = r[key] || "N/A"; map[v] = (map[v] || 0) + 1; });
  return Object.entries(map).sort((a, b) => b[1] - a[1]).map(([name, value]) => ({ name, value }));
}
function dlXl(rows, filename) {
  const ws = XLSX.utils.json_to_sheet(rows);
  const maxW = {};
  [Object.keys(rows[0] || {}), ...rows.map((r) => Object.values(r).map(String))].forEach((row) => {
    row.forEach((c, i) => { maxW[i] = Math.max(maxW[i] || 8, String(c).length + 2); });
  });
  ws["!cols"] = Object.values(maxW).map((w) => ({ wch: Math.min(w, 40) }));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Datos");
  XLSX.writeFile(wb, filename);
}
function isCAGS(r) {
  return r.semestre === "8" || norm(r.cag) === norm("Sí");
}
function isDIC25(r) {
  return norm(r.exatec || "").includes("diciembre") || norm(r.exatec || "").includes("dic");
}

/* ═══ PASTE / CLIPBOARD ═══ */
const HEADER_MAP = {
  "día": "dia", "dia": "dia", "fecha": "dia",
  "hora": "hora",
  "matrícula": "matricula", "matricula": "matricula",
  "nombre": "nombre",
  "ap": "ap", "apellido paterno": "ap",
  "am": "am", "apellido materno": "am",
  "servicio": "servicio",
  "clave": "clave",
  "atiende": "atiende", "asesor": "atiende",
  "escuela": "escuela",
  "programa": "programa",
  "estatus": "estatus",
  "interés asesoría": "interes_asesoria",
  "interes asesoria": "interes_asesoria",
  "interés asesoria": "interes_asesoria",
  "interes asesoría": "interes_asesoria",
  "interés aseosría": "interes_asesoria",
  "interes aseosria": "interes_asesoria",
  "interés aseosria": "interes_asesoria",
  "interés": "interes_asesoria",
  "semestre": "semestre",
  "cag": "cag",
  "exatec": "exatec",
  "modalidad": "modalidad",
  "campus": "campus",
  "comunidad": "comunidad",
  "celular": "celular",
  "correo personal": "correo_personal",
  "correo": "correo_personal",
  "email": "correo_personal",
  "linkedin": "linkedin",
  "notas": "notas",
};

const DEFAULT_COL_ORDER = [
  "dia","hora","matricula","nombre","ap","am","servicio","clave","atiende",
  "escuela","programa","estatus","interes_asesoria","semestre","cag","exatec",
  "modalidad","campus","comunidad","celular","correo_personal","linkedin","notas",
];

function parseClipDate(v, fmt = "mdy") {
  if (!v || !v.trim()) return null;
  const s = v.trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  const parts = s.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})$/);
  if (parts) {
    const [, a, b, y] = parts;
    const [m, d] = fmt === "mdy" ? [a, b] : [b, a];
    const mi = parseInt(m), di = parseInt(d);
    if (mi < 1 || mi > 12 || di < 1 || di > 31) return null;
    return `${y}-${m.padStart(2, "0")}-${d.padStart(2, "0")}`;
  }
  return null;
}

function parsePastedTSV(text, dateFmt = "mdy") {
  const lines = text.trim().split(/\r?\n/).filter((l) => l.trim());
  if (!lines.length) return { hasHeaders: false, rows: [] };

  const splitLine = (l) => l.split("\t").map((c) => c.trim());
  const first = splitLine(lines[0]);

  const normalizeH = (h) => norm(h).replace(/\s+/g, " ").trim();
  const mappedHeaders = first.map((h) => HEADER_MAP[normalizeH(h)] || null);
  const hasHeaders = mappedHeaders.filter(Boolean).length >= 2;

  const dataLines = hasHeaders ? lines.slice(1) : lines;
  const colMap = hasHeaders ? mappedHeaders : DEFAULT_COL_ORDER;

  const rows = dataLines.map((line) => {
    const cells = splitLine(line);
    const row = {};
    colMap.forEach((key, i) => {
      if (!key) return;
      const val = cells[i] ?? "";
      if (key === "dia") row[key] = parseClipDate(val, dateFmt) || null;
      else row[key] = val;
    });
    if (row.servicio && !row.clave) row.clave = SERVICIO_CLAVE[row.servicio] || "";
    if (!row.estatus) row.estatus = "Agendado";
    return row;
  }).filter((r) => r.matricula?.trim());

  return { hasHeaders, rows };
}

/* ═══ STYLES ═══ */
const S = {
  card: { background: "rgba(255,255,255,0.04)", border: "1px solid rgba(255,255,255,0.08)", borderRadius: 18, padding: 24, marginBottom: 16 },
  kpi: (color) => ({ background: "rgba(255,255,255,0.04)", border: "1px solid rgba(255,255,255,0.08)", borderRadius: 16, padding: "20px 24px", borderLeft: `4px solid ${color}`, transition: "all .2s" }),
  btn: (color = "#6366f1") => ({ background: `${color}2e`, color, border: "none", borderRadius: 10, padding: "8px 16px", fontSize: 13, fontWeight: 600, cursor: "pointer", fontFamily: "'Plus Jakarta Sans'", transition: "all .2s" }),
  badge: (color = "#6366f1") => ({ display: "inline-block", background: `${color}22`, color, borderRadius: 20, padding: "3px 10px", fontSize: 11, fontWeight: 600 }),
  input: { background: "#0a1525", border: "1px solid rgba(255,255,255,0.1)", borderRadius: 10, padding: "8px 14px", color: "#e8e9ed", fontSize: 13, fontFamily: "'Plus Jakarta Sans'", outline: "none", width: "100%" },
  select: { background: "#0a1525", border: "1px solid rgba(255,255,255,0.1)", borderRadius: 10, padding: "8px 14px", color: "#e8e9ed", fontSize: 13, fontFamily: "'Plus Jakarta Sans'", outline: "none" },
  mono: { fontFamily: "'JetBrains Mono', monospace" },
  dim: { color: "#8e92a6", fontSize: 12 },
  h2: { fontSize: 18, fontWeight: 700, marginBottom: 16 },
  h3: { fontSize: 15, fontWeight: 600, marginBottom: 12, color: "#a5b4fc" },
  grid: (cols) => ({ display: "grid", gridTemplateColumns: `repeat(${cols}, 1fr)`, gap: 16 }),
  overlay: { position: "fixed", inset: 0, background: "rgba(0,0,0,0.6)", backdropFilter: "blur(6px)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 1000 },
  modal: { background: "#0f1628", borderRadius: 22, border: "1px solid rgba(255,255,255,0.1)", padding: 32, maxWidth: 900, width: "90vw", maxHeight: "85vh", overflowY: "auto", position: "relative" },
};

/* ═══ REUSABLE ═══ */
function Cd({ children, style }) { return <div style={{ ...S.card, ...style }}>{children}</div>; }
function Bt({ children, color, onClick, style, disabled }) {
  return (
    <button style={{ ...S.btn(color), ...style, opacity: disabled ? 0.4 : 1, pointerEvents: disabled ? "none" : "auto" }}
      onClick={onClick} disabled={disabled}
      onMouseEnter={(e) => { if (!disabled) e.target.style.opacity = 0.8; }}
      onMouseLeave={(e) => { if (!disabled) e.target.style.opacity = 1; }}>
      {children}
    </button>
  );
}
function KPI({ label, value, sub, color = "#6366f1" }) {
  return (
    <div style={S.kpi(color)}
      onMouseEnter={(e) => { e.currentTarget.style.transform = "translateY(-3px)"; e.currentTarget.style.boxShadow = `0 8px 24px ${color}22`; }}
      onMouseLeave={(e) => { e.currentTarget.style.transform = "translateY(0)"; e.currentTarget.style.boxShadow = "none"; }}>
      <div style={{ color: "#8e92a6", fontSize: 11, fontWeight: 600, textTransform: "uppercase", letterSpacing: 1, marginBottom: 6 }}>{label}</div>
      <div style={{ ...S.mono, fontSize: 28, fontWeight: 700, color }}>{value}</div>
      {sub && <div style={{ color: "#6b6f82", fontSize: 11, marginTop: 4 }}>{sub}</div>}
    </div>
  );
}
function SB({ items, total }) {
  return (
    <div>{items.map(({ name, value, color }, i) => {
      const pct = total ? ((value / total) * 100).toFixed(1) : 0;
      return (
        <div key={i} style={{ marginBottom: 10 }}>
          <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 4 }}>
            <span style={{ fontSize: 12, color: "#e8e9ed" }}>{name}</span>
            <span style={{ ...S.mono, fontSize: 12, color: "#8e92a6" }}>{value} ({pct}%)</span>
          </div>
          <div style={{ height: 5, background: "rgba(255,255,255,0.06)", borderRadius: 3, overflow: "hidden" }}>
            <div style={{ height: "100%", width: `${pct}%`, borderRadius: 3, background: color || CHART_COLORS[i % CHART_COLORS.length], transition: "width .5s" }} />
          </div>
        </div>
      );
    })}</div>
  );
}
const TT = ({ active, payload, label }) => {
  if (!active || !payload?.length) return null;
  return (
    <div style={{ background: "#0f1a2e", border: "1px solid rgba(255,255,255,0.1)", borderRadius: 10, padding: "10px 14px", fontSize: 12 }}>
      <div style={{ fontWeight: 600, marginBottom: 4 }}>{label}</div>
      {payload.map((p, i) => (
        <div key={i} style={{ color: p.color || p.fill, ...S.mono, fontSize: 11 }}>{p.name || p.dataKey}: {p.value}</div>
      ))}
    </div>
  );
};

/* ═══ SPINNER ═══ */
function Spinner() {
  return (
    <div style={{ display: "flex", alignItems: "center", justifyContent: "center", padding: 60 }}>
      <div style={{ width: 32, height: 32, border: "3px solid rgba(99,102,241,0.2)", borderTopColor: "#6366f1", borderRadius: "50%", animation: "spin 0.8s linear infinite" }} />
      <style>{`@keyframes spin { to { transform: rotate(360deg); } }`}</style>
    </div>
  );
}

/* ═══ MODAL ═══ */
function Modal({ onClose, children }) {
  return (
    <div style={S.overlay} onClick={onClose}>
      <div style={S.modal} onClick={(e) => e.stopPropagation()}>
        <button onClick={onClose} style={{ position: "absolute", top: 16, right: 16, background: "none", border: "none", color: "#8e92a6", fontSize: 20, cursor: "pointer" }}>✕</button>
        {children}
      </div>
    </div>
  );
}

/* ═══ STUDENT MODAL ═══ */
function StudentModal({ matricula, records, onClose }) {
  const sorted = [...records].sort((a, b) => (b.dia || "") > (a.dia || "") ? 1 : -1);
  const latest = sorted[0] || {};
  const asist = records.filter((r) => r.estatus === "Asistencia").length;
  const services = [...new Set(records.map((r) => r.servicio).filter(Boolean))];
  const cags = isCAGS(latest);
  const dic25 = isDIC25(latest);
  return (
    <Modal onClose={onClose}>
      <div style={{ marginBottom: 20 }}>
        <div style={{ fontSize: 20, fontWeight: 700 }}>{latest.nombre} {latest.ap} {latest.am}</div>
        <div style={{ ...S.mono, color: "#8e92a6", fontSize: 13, marginTop: 4 }}>{matricula}</div>
        <div style={{ display: "flex", gap: 8, marginTop: 10, flexWrap: "wrap" }}>
          {latest.escuela && <span style={S.badge("#6366f1")}>{latest.escuela}</span>}
          {latest.programa && <span style={S.badge("#10b981")}>{latest.programa}</span>}
          {cags && <span style={{ ...S.badge("#a855f7"), fontWeight: 700 }}>★ CAGS JUN26</span>}
          {dic25 && <span style={{ ...S.badge("#22d3ee"), fontWeight: 700 }}>✓ DIC25</span>}
        </div>
      </div>
      <div style={S.grid(4)}>
        <KPI label="Sesiones" value={records.length} color="#6366f1" />
        <KPI label="Asistencias" value={asist} color="#10b981" />
        <KPI label="Servicios" value={services.length} color="#f59e0b" />
        <KPI label="Tasa asist." value={records.length ? `${((asist / records.length) * 100).toFixed(0)}%` : "—"} color="#8b5cf6" />
      </div>
      <div style={{ marginTop: 20 }}>
        <div style={S.h3}>Servicios utilizados</div>
        <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
          {services.map((s, i) => <span key={i} style={S.badge(CHART_COLORS[i % CHART_COLORS.length])}>{s}</span>)}
        </div>
      </div>
      <div style={{ marginTop: 20 }}>
        <div style={S.h3}>Timeline</div>
        {sorted.map((r, i) => (
          <div key={i} style={{ display: "flex", gap: 12, padding: "10px 0", borderBottom: "1px solid rgba(255,255,255,0.05)", alignItems: "center" }}>
            <div style={{ ...S.mono, fontSize: 11, color: "#6b6f82", minWidth: 90 }}>{fmtDate(r.dia)}</div>
            <div style={{ fontSize: 12, flex: 1 }}>{r.servicio || "—"}</div>
            <div style={{ fontSize: 11, color: "#8e92a6" }}>{r.atiende}</div>
            <span style={S.badge(STATUS_COLORS[r.estatus] || "#6366f1")}>{r.estatus}</span>
          </div>
        ))}
      </div>
    </Modal>
  );
}

/* ═══ TAB HOME ═══ */
function TabHome({ data, onStatusChange }) {
  const today = todayISO();
  const hoy = useMemo(() =>
    data.filter((r) => r.dia === today).sort((a, b) => (a.hora || "") < (b.hora || "") ? -1 : 1),
    [data, today]
  );

  const dateStr = new Date().toLocaleDateString("es-MX", { weekday: "long", day: "numeric", month: "long" });
  const capitalize = (s) => s.charAt(0).toUpperCase() + s.slice(1);

  const stats = useMemo(() => ({
    total: hoy.length,
    asistencia: hoy.filter((r) => r.estatus === "Asistencia").length,
    falta: hoy.filter((r) => r.estatus === "Falta").length,
    cancelacion: hoy.filter((r) => r.estatus === "Cancelación").length,
    express: hoy.filter((r) => r.estatus === "Express").length,
    pendientes: hoy.filter((r) => r.estatus === "Agendado").length,
  }), [hoy]);

  return (
    <div>
      {/* Header del día */}
      <div style={{ marginBottom: 24 }}>
        <h1 style={{ fontSize: 26, fontWeight: 800, margin: 0 }}>
          {capitalize(dateStr)}
        </h1>
        <p style={{ color: "#8e92a6", fontSize: 14, margin: "4px 0 0" }}>
          {hoy.length === 0 ? "Sin asesorías registradas para hoy" : `${hoy.length} asesoría${hoy.length !== 1 ? "s" : ""} agendada${hoy.length !== 1 ? "s" : ""}`}
        </p>
      </div>

      {/* KPIs del día */}
      {hoy.length > 0 && (
        <div style={{ ...S.grid(5), marginBottom: 24 }}>
          <KPI label="Total hoy" value={stats.total} color="#6366f1" />
          <KPI label="Pendientes" value={stats.pendientes} color="#8e92a6" />
          <KPI label="Asistencia" value={stats.asistencia} color="#10b981" />
          <KPI label="Falta" value={stats.falta} color="#ef4444" />
          <KPI label="Cancelación" value={stats.cancelacion} color="#8b5cf6" />
        </div>
      )}

      {/* Lista de asesorías de hoy */}
      {hoy.length === 0 ? (
        <Cd>
          <div style={{ textAlign: "center", padding: 40 }}>
            <div style={{ fontSize: 40, marginBottom: 12, opacity: 0.3 }}>📅</div>
            <div style={{ color: "#8e92a6", fontSize: 14 }}>No hay asesorías registradas para hoy.</div>
            <div style={{ color: "#6b6f82", fontSize: 12, marginTop: 8 }}>Ve a la pestaña Asesorías para agregar nuevas.</div>
          </div>
        </Cd>
      ) : (
        <Cd style={{ padding: 0, overflow: "hidden" }}>
          <div style={{ padding: "16px 20px", borderBottom: "1px solid rgba(255,255,255,0.06)" }}>
            <span style={{ fontWeight: 700, fontSize: 14, color: "#a5b4fc" }}>Asesorías de hoy</span>
          </div>
          {hoy.map((r) => (
            <HomeRow key={r.id} record={r} onStatusChange={onStatusChange} />
          ))}
        </Cd>
      )}

      {/* Últimas asesorías (no de hoy) */}
      <RecentSection data={data} today={today} />
    </div>
  );
}

function HomeRow({ record: r, onStatusChange }) {
  const [saving, setSaving] = useState(false);
  const current = r.estatus;

  const handleStatus = async (newStatus) => {
    if (saving || current === newStatus) return;
    setSaving(true);
    await onStatusChange(r.id, newStatus);
    setSaving(false);
  };

  const nombre = [r.nombre, r.ap, r.am].filter(Boolean).join(" ") || "Sin nombre";
  const cags = isCAGS(r);
  const dic25 = isDIC25(r);

  return (
    <div style={{
      display: "flex", alignItems: "center", gap: 16, padding: "14px 20px",
      borderBottom: "1px solid rgba(255,255,255,0.04)",
      background: current === "Asistencia" ? "rgba(16,185,129,0.04)"
        : current === "Falta" ? "rgba(239,68,68,0.04)"
        : current === "Cancelación" ? "rgba(139,92,246,0.04)"
        : current === "Express" ? "rgba(245,158,11,0.04)"
        : "transparent",
      transition: "background .2s",
    }}
      onMouseEnter={(e) => { if (current === "Agendado") e.currentTarget.style.background = "rgba(99,102,241,0.05)"; }}
      onMouseLeave={(e) => { if (current === "Agendado") e.currentTarget.style.background = "transparent"; }}
    >
      {/* Hora */}
      <div style={{ ...S.mono, fontSize: 18, fontWeight: 700, color: "#6366f1", minWidth: 60 }}>
        {r.hora || "—"}
      </div>

      {/* Info */}
      <div style={{ flex: 1, minWidth: 0 }}>
        <div style={{ fontWeight: 600, fontSize: 15, display: "flex", alignItems: "center", gap: 6, flexWrap: "wrap" }}>
          {nombre}
          {cags  && <span style={{ ...S.badge("#a855f7"), fontSize: 9 }}>CAGS</span>}
          {dic25 && <span style={{ ...S.badge("#22d3ee"), fontSize: 9 }}>DIC25</span>}
        </div>
        <div style={{ color: "#8e92a6", fontSize: 12, marginTop: 2, display: "flex", gap: 12 }}>
          <span>{r.atiende || "Sin asesor"}</span>
          {r.servicio && <span style={{ color: "#6b6f82" }}>· {r.servicio}</span>}
          {r.modalidad && <span style={{ color: "#6b6f82" }}>· {r.modalidad}</span>}
        </div>
      </div>

      {/* Matrícula */}
      <div style={{ ...S.mono, fontSize: 11, color: "#a5b4fc", minWidth: 100, textAlign: "right" }}>
        {r.matricula}
      </div>

      {/* Botones de estatus */}
      <div style={{ display: "flex", gap: 6, flexShrink: 0 }}>
        {["Asistencia","Falta","Cancelación","Express"].map((s) => {
          const col = STATUS_COLORS[s];
          const active = current === s;
          return (
            <button key={s} onClick={() => handleStatus(s)} disabled={saving}
              style={{
                background: active ? `${col}33` : "rgba(255,255,255,0.04)",
                color: active ? col : "#6b6f82",
                border: `1px solid ${active ? col : "rgba(255,255,255,0.08)"}`,
                borderRadius: 8, padding: "5px 10px", fontSize: 11, fontWeight: 600,
                cursor: saving ? "wait" : "pointer", fontFamily: "'Plus Jakarta Sans'",
                transition: "all .15s",
              }}
              onMouseEnter={(e) => { if (!active && !saving) { e.target.style.color = col; e.target.style.borderColor = col; } }}
              onMouseLeave={(e) => { if (!active && !saving) { e.target.style.color = "#6b6f82"; e.target.style.borderColor = "rgba(255,255,255,0.08)"; } }}
            >
              {s === "Asistencia" ? "✓ Asistencia"
                : s === "Falta" ? "✕ Falta"
                : s === "Cancelación" ? "⊘ Canceló"
                : "⚡ Express"}
            </button>
          );
        })}
      </div>
    </div>
  );
}

function RecentSection({ data, today }) {
  const recent = useMemo(() =>
    [...data]
      .filter((r) => r.dia && r.dia !== today)
      .sort((a, b) => (b.dia || "") > (a.dia || "") ? 1 : -1)
      .slice(0, 20),
    [data, today]
  );
  if (!recent.length) return null;
  return (
    <Cd style={{ marginTop: 20 }}>
      <div style={S.h3}>Asesorías recientes</div>
      <div style={{ overflowX: "auto" }}>
        <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
          <thead>
            <tr style={{ borderBottom: "1px solid rgba(255,255,255,0.08)" }}>
              {["Fecha","Hora","Alumno","Matrícula","Asesor","Servicio","Estatus"].map((h) => (
                <th key={h} style={{ textAlign: "left", padding: "8px 10px", fontSize: 10, textTransform: "uppercase", letterSpacing: 1, color: "#6b6f82", fontWeight: 600, whiteSpace: "nowrap" }}>{h}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {recent.map((r) => (
              <tr key={r.id} style={{ borderBottom: "1px solid rgba(255,255,255,0.04)" }}
                onMouseEnter={(e) => e.currentTarget.style.background = "rgba(99,102,241,0.05)"}
                onMouseLeave={(e) => e.currentTarget.style.background = "transparent"}>
                <td style={{ padding: "8px 10px", ...S.mono, fontSize: 11, color: "#6b6f82" }}>{fmtDate(r.dia)}</td>
                <td style={{ padding: "8px 10px", ...S.mono, fontSize: 11 }}>{r.hora || "—"}</td>
                <td style={{ padding: "8px 10px", fontWeight: 500 }}>{[r.nombre, r.ap].filter(Boolean).join(" ") || "—"}</td>
                <td style={{ padding: "8px 10px", ...S.mono, fontSize: 11, color: "#a5b4fc" }}>{r.matricula}</td>
                <td style={{ padding: "8px 10px", fontSize: 11, color: "#8e92a6" }}>{r.atiende}</td>
                <td style={{ padding: "8px 10px", fontSize: 11, maxWidth: 160, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{r.servicio}</td>
                <td style={{ padding: "8px 10px" }}>
                  <span style={S.badge(STATUS_COLORS[r.estatus] || "#6366f1")}>{r.estatus}</span>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </Cd>
  );
}

/* ═══ PASTE MODAL ═══ */
function PasteModal({ onClose, onImport }) {
  const [text, setText] = useState("");
  const [dateFmt, setDateFmt] = useState("mdy");
  const [parsed, setParsed] = useState(null);
  const [importing, setImporting] = useState(false);

  useEffect(() => {
    if (!text.trim()) { setParsed(null); return; }
    setParsed(parsePastedTSV(text, dateFmt));
  }, [text, dateFmt]);

  const handleImport = async () => {
    if (!parsed?.rows?.length) return;
    setImporting(true);
    await onImport(parsed.rows);
    setImporting(false);
    onClose();
  };

  const rows = parsed?.rows || [];

  return (
    <Modal onClose={onClose}>
      <div style={{ fontSize: 18, fontWeight: 700, marginBottom: 8 }}>Pegar desde Excel</div>
      <div style={{ color: "#8e92a6", fontSize: 13, marginBottom: 14 }}>
        Copia tus filas desde Excel <b style={{ color: "#e8e9ed" }}>incluyendo los encabezados</b> y pégalas aquí (Ctrl+V).
        La columna <b style={{ color: "#a5b4fc" }}>Matrícula</b> es obligatoria.
      </div>

      <textarea autoFocus
        style={{ ...S.input, height: 130, resize: "vertical", fontFamily: "'JetBrains Mono', monospace", fontSize: 11, lineHeight: 1.6 }}
        placeholder={"Día\tHora\tMatrícula\tNombre\tAP\tAM\tServicio...\n1/5/2026\t10:00\tA01198340\tAna Paula\tRamos\t..."}
        value={text}
        onChange={(e) => setText(e.target.value)}
      />

      <div style={{ display: "flex", gap: 10, marginTop: 10, alignItems: "center" }}>
        <span style={{ fontSize: 12, color: "#8e92a6" }}>Formato de fecha:</span>
        {[["dmy", "D/M/AAAA"], ["mdy", "M/D/AAAA"]].map(([val, label]) => (
          <button key={val} onClick={() => setDateFmt(val)}
            style={{ ...S.btn(dateFmt === val ? "#6366f1" : "#4a5080"), fontSize: 11, padding: "4px 12px" }}>
            {label}
          </button>
        ))}
        {parsed && (
          <span style={{ marginLeft: "auto", ...S.mono, fontSize: 11, color: "#6b6f82" }}>
            {rows.length} registro{rows.length !== 1 ? "s" : ""} detectado{rows.length !== 1 ? "s" : ""}
          </span>
        )}
      </div>

      {parsed && rows.length > 0 && (
        <div style={{ marginTop: 16 }}>
          <div style={S.h3}>
            Vista previa
            {!parsed.hasHeaders && (
              <span style={{ color: "#f59e0b", fontSize: 11, fontWeight: 400, marginLeft: 8 }}>
                (sin encabezados detectados — usando orden de columnas por defecto)
              </span>
            )}
          </div>
          <div style={{ overflowX: "auto", maxHeight: 220, overflowY: "auto", border: "1px solid rgba(255,255,255,0.08)", borderRadius: 10 }}>
            <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 11 }}>
              <thead>
                <tr style={{ background: "#0a1525", position: "sticky", top: 0 }}>
                  {["Día","Hora","Matrícula","Nombre","AP","Atiende","Servicio","Estatus"].map((h) => (
                    <th key={h} style={{ padding: "7px 10px", textAlign: "left", color: "#6b6f82", fontSize: 10, textTransform: "uppercase", letterSpacing: 0.8, fontWeight: 600, whiteSpace: "nowrap" }}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {rows.slice(0, 15).map((r, i) => (
                  <tr key={i} style={{ borderTop: "1px solid rgba(255,255,255,0.05)" }}>
                    <td style={{ padding: "6px 10px", ...S.mono, fontSize: 10, color: "#6b6f82" }}>{r.dia || "—"}</td>
                    <td style={{ padding: "6px 10px", ...S.mono }}>{r.hora || "—"}</td>
                    <td style={{ padding: "6px 10px", ...S.mono, color: "#a5b4fc" }}>{r.matricula}</td>
                    <td style={{ padding: "6px 10px" }}>{r.nombre || "—"}</td>
                    <td style={{ padding: "6px 10px", color: "#8e92a6" }}>{r.ap || "—"}</td>
                    <td style={{ padding: "6px 10px", color: "#8e92a6" }}>{r.atiende || "—"}</td>
                    <td style={{ padding: "6px 10px", maxWidth: 160, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{r.servicio || "—"}</td>
                    <td style={{ padding: "6px 10px" }}>
                      <span style={S.badge(STATUS_COLORS[r.estatus] || "#6366f1")}>{r.estatus}</span>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
            {rows.length > 15 && (
              <div style={{ padding: "8px 12px", fontSize: 11, color: "#6b6f82", textAlign: "center" }}>
                ... y {rows.length - 15} registro{rows.length - 15 !== 1 ? "s" : ""} más
              </div>
            )}
          </div>
          <div style={{ marginTop: 14, display: "flex", gap: 10 }}>
            <Bt color="#10b981" onClick={handleImport} disabled={importing}>
              {importing ? "Importando..." : `✓ Importar ${rows.length} registro${rows.length !== 1 ? "s" : ""}`}
            </Bt>
            <Bt color="#8e92a6" onClick={onClose}>Cancelar</Bt>
          </div>
        </div>
      )}

      {parsed && rows.length === 0 && (
        <div style={{ marginTop: 12, padding: "12px 16px", borderRadius: 10, fontSize: 13, background: "rgba(239,68,68,0.1)", border: "1px solid rgba(239,68,68,0.3)", color: "#ef4444" }}>
          No se encontraron registros válidos. Asegúrate de incluir la columna <b>Matrícula</b> con datos.
        </div>
      )}
    </Modal>
  );
}

/* ═══ TAB ASESORÍAS (grid de captura) ═══ */
const GRID_COLS = [
  { key: "dia",             label: "Día",         width: 130, type: "date" },
  { key: "hora",            label: "Hora",        width: 80,  type: "text", placeholder: "09:00" },
  { key: "matricula",       label: "Matrícula",   width: 110, type: "text", required: true },
  { key: "nombre",          label: "Nombre",      width: 140, type: "text" },
  { key: "ap",              label: "A. Paterno",  width: 120, type: "text" },
  { key: "am",              label: "A. Materno",  width: 120, type: "text" },
  { key: "atiende",         label: "Atiende",     width: 160, type: "select", options: ASESORES },
  { key: "servicio",        label: "Servicio",    width: 200, type: "select", options: SERVICIOS.map((s) => s.label) },
  { key: "clave",           label: "Clave",       width: 70,  type: "text", readOnly: true },
  { key: "estatus",         label: "Estatus",     width: 120, type: "select", options: ESTATUSES },
  { key: "escuela",         label: "Escuela",     width: 160, type: "select", options: ESCUELAS },
  { key: "programa",        label: "Programa",    width: 160, type: "text" },
  { key: "semestre",        label: "Semestre",    width: 90,  type: "select", options: SEMESTRES },
  { key: "cag",             label: "CAG",         width: 70,  type: "select", options: CAG_OPTS },
  { key: "exatec",          label: "EXATEC",      width: 180, type: "select", options: EXATEC_OPTS },
  { key: "interes_asesoria",label: "Interés",     width: 180, type: "select", options: INTERESES },
  { key: "modalidad",       label: "Modalidad",   width: 110, type: "select", options: MODALIDADES },
  { key: "campus",          label: "Campus",      width: 100, type: "text" },
  { key: "comunidad",       label: "Comunidad",   width: 110, type: "select", options: COMUNIDADES },
  { key: "celular",         label: "Celular",     width: 120, type: "text" },
  { key: "correo_personal", label: "Correo",      width: 180, type: "text" },
  { key: "linkedin",        label: "LinkedIn",    width: 140, type: "text" },
  { key: "notas",           label: "Notas",       width: 200, type: "text" },
];

const FROZEN_COUNT = 4; // dia, hora, matricula, nombre
const FROZEN_LEFT = (() => {
  const map = { __del: 0 };
  let left = 40;
  GRID_COLS.forEach((col, i) => { if (i < FROZEN_COUNT) { map[col.key] = left; left += col.width; } });
  return map;
})();

function emptyRow() {
  return { id: null, dia: todayISO(), hora: "", matricula: "", nombre: "", ap: "", am: "",
    servicio: "", clave: "", atiende: "", escuela: "", programa: "", estatus: "Agendado",
    interes_asesoria: "", semestre: "", cag: "", exatec: "", modalidad: "", campus: "",
    comunidad: "", celular: "", correo_personal: "", linkedin: "", notas: "" };
}

function TabAsesorias({ data, onRefresh }) {
  const [rows, setRows] = useState([]);
  const [editCell, setEditCell] = useState(null); // {rowIdx, key}
  const [saving, setSaving] = useState(new Set());
  const [deleting, setDeleting] = useState(new Set());
  const [search, setSearch] = useState("");
  const [fAsesor, setFAsesor] = useState("");
  const [fEstatus, setFEstatus] = useState("");
  const [importing, setImporting] = useState(false);
  const [importMsg, setImportMsg] = useState(null);
  const [showPaste, setShowPaste] = useState(false);
  const [gridPage, setGridPage] = useState(0);
  const [hoveredRow, setHoveredRow] = useState(null);
  const inputRef = useRef();
  const dSearch = useDebounce(search, 200);

  useEffect(() => {
    setRows([...data, emptyRow()]);
    setGridPage(0);
  }, [data]);

  useEffect(() => { setGridPage(0); }, [dSearch, fAsesor, fEstatus]);

  const filteredWithIdx = useMemo(() => {
    const withIdx = rows.map((row, rowsIdx) => ({ row, rowsIdx }));
    if (!dSearch && !fAsesor && !fEstatus) return withIdx;
    return withIdx.filter(({ row: r }) => {
      if (r.id === null) return true;
      const q = norm(dSearch);
      const matchSearch = !dSearch || norm(r.nombre).includes(q) || norm(r.matricula).includes(q) || norm(r.ap).includes(q);
      const matchAsesor = !fAsesor || r.atiende === fAsesor;
      const matchEstatus = !fEstatus || r.estatus === fEstatus;
      return matchSearch && matchAsesor && matchEstatus;
    });
  }, [rows, dSearch, fAsesor, fEstatus]);

  const GRID_PAGE_SIZE = 150;
  const existingItems = filteredWithIdx.filter(({ row }) => row.id !== null);
  const newItem = filteredWithIdx.find(({ row }) => row.id === null);
  const totalGridPages = Math.ceil(existingItems.length / GRID_PAGE_SIZE);
  const pagedItems = [
    ...existingItems.slice(gridPage * GRID_PAGE_SIZE, (gridPage + 1) * GRID_PAGE_SIZE),
    ...(newItem ? [newItem] : [emptyRow()].map((row) => ({ row, rowsIdx: rows.length - 1 }))),
  ];

  const updateRowLocal = (rowIdx, key, value) => {
    setRows((prev) => {
      const next = [...prev];
      const row = { ...next[rowIdx] };
      row[key] = value;
      if (key === "servicio") row.clave = SERVICIO_CLAVE[value] || "";
      // Add new empty row if editing the last (new) row
      if (row.id === null && key === "matricula" && value) {
        next[rowIdx] = row;
        if (rowIdx === next.length - 1) next.push(emptyRow());
        return next;
      }
      next[rowIdx] = row;
      return next;
    });
  };

  const saveRow = async (rowIdx) => {
    const row = rows[rowIdx];
    if (!row || saving.has(rowIdx)) return;
    if (!row.matricula?.trim()) return;

    const payload = { ...row };
    delete payload.id;

    setSaving((s) => new Set(s).add(rowIdx));
    try {
      if (row.id) {
        await supabase.from("asesorias").update(payload).eq("id", row.id);
      } else {
        const { data: inserted } = await supabase.from("asesorias").insert(payload).select().single();
        if (inserted) {
          setRows((prev) => {
            const next = [...prev];
            next[rowIdx] = inserted;
            return next;
          });
        }
      }
      await onRefresh();
    } finally {
      setSaving((s) => { const n = new Set(s); n.delete(rowIdx); return n; });
    }
  };

  const deleteRow = async (rowIdx) => {
    const row = rows[rowIdx];
    if (!row?.id || deleting.has(rowIdx)) return;
    if (!confirm(`¿Eliminar la asesoría de ${row.nombre || row.matricula}?`)) return;
    setDeleting((s) => new Set(s).add(rowIdx));
    await supabase.from("asesorias").delete().eq("id", row.id);
    await onRefresh();
    setDeleting((s) => { const n = new Set(s); n.delete(rowIdx); return n; });
  };

  const handleCellBlur = (rowIdx) => {
    saveRow(rowIdx);
    setEditCell(null);
  };

  const handleKeyDown = (e, rowIdx, colIdx) => {
    if (e.key === "Enter" || e.key === "Tab") {
      e.preventDefault();
      saveRow(rowIdx);
      // Move to next cell
      const nextCol = colIdx + 1 < GRID_COLS.length ? colIdx + 1 : 0;
      const nextRow = colIdx + 1 < GRID_COLS.length ? rowIdx : rowIdx + 1;
      if (nextRow < rows.length) {
        setEditCell({ rowIdx: nextRow, key: GRID_COLS[nextCol].key });
      } else {
        setEditCell(null);
      }
    }
    if (e.key === "Escape") {
      setEditCell(null);
    }
  };

  /* Paste import */
  const handlePasteImport = async (rows) => {
    const validFields = ["dia","hora","matricula","nombre","ap","am","servicio","clave","atiende",
      "escuela","programa","estatus","interes_asesoria","semestre","cag","exatec","modalidad",
      "campus","comunidad","celular","correo_personal","linkedin","notas"];
    const toInsert = rows.map((r) => {
      const clean = {};
      validFields.forEach((k) => {
        const v = r[k];
        clean[k] = (v === "" || v === undefined) ? null : v;
      });
      return clean;
    });
    const chunkSize = 200;
    let inserted = 0;
    let firstError = null;
    for (let i = 0; i < toInsert.length; i += chunkSize) {
      const { error } = await supabase.from("asesorias").insert(toInsert.slice(i, i + chunkSize));
      if (error) {
        if (!firstError) firstError = error;
        console.error("Supabase insert error:", error, toInsert.slice(i, i + 3));
      } else {
        inserted += Math.min(chunkSize, toInsert.length - i);
      }
    }
    if (firstError) {
      setImportMsg({ type: "error", text: `Error al insertar: ${firstError.message} (${firstError.code}). Se insertaron ${inserted} de ${toInsert.length}.` });
    } else {
      setImportMsg({ type: "ok", text: `${inserted} registros pegados correctamente.` });
    }
    await onRefresh();
  };

  /* Excel import */
  const handleImport = async (file) => {
    if (!file) return;
    setImporting(true);
    setImportMsg(null);
    try {
      const buffer = await file.arrayBuffer();
      const wb = XLSX.read(buffer, { cellDates: false });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const raw = XLSX.utils.sheet_to_json(ws, { defval: "" });
      const normalize = (s) => String(s ?? "").trim();

      const toInsert = raw.map((r) => ({
        dia: parseExcelDate(r["Día"] || r["Dia"] || r["DÍA"]),
        hora: normalize(r["Hora"]),
        matricula: normalize(r["Matrícula"] || r["Matricula"]),
        nombre: normalize(r["Nombre"]),
        ap: normalize(r["AP"]),
        am: normalize(r["AM"]),
        servicio: normalize(r["Servicio"]),
        clave: normalize(r["Clave"]),
        atiende: normalize(r["Atiende"]),
        escuela: normalize(r["Escuela"]),
        programa: normalize(r["Programa"]),
        estatus: normalize(r["Estatus"]) || "Agendado",
        interes_asesoria: normalize(r["Interés Asesoría"] || r["Interes Asesoria"] || r["Interés Asesoria"] || r["Interes Asesoría"]),
        semestre: normalize(r["Semestre"]),
        cag: normalize(r["CAG"]),
        exatec: normalize(r["EXATEC"]),
        modalidad: normalize(r["Modalidad"]),
        campus: normalize(r["Campus"]),
        comunidad: normalize(r["Comunidad"]),
        celular: normalize(r["Celular"]),
        correo_personal: normalize(r["Correo Personal"]),
        linkedin: normalize(r["LinkedIn"]),
        notas: normalize(r["Notas"]),
      })).filter((r) => r.matricula);

      if (!toInsert.length) {
        setImportMsg({ type: "error", text: "No se encontraron registros válidos (¿falta columna Matrícula?)." });
        return;
      }

      const chunkSize = 500;
      let inserted = 0;
      for (let i = 0; i < toInsert.length; i += chunkSize) {
        const { error } = await supabase.from("asesorias").insert(toInsert.slice(i, i + chunkSize));
        if (!error) inserted += Math.min(chunkSize, toInsert.length - i);
      }
      setImportMsg({ type: "ok", text: `${inserted} registros importados correctamente.` });
      await onRefresh();
    } catch (e) {
      setImportMsg({ type: "error", text: `Error: ${e.message}` });
    } finally {
      setImporting(false);
      if (inputRef.current) inputRef.current.value = "";
    }
  };

  return (
    <div>
      {/* Toolbar */}
      <div style={{ display: "flex", gap: 10, marginBottom: 14, flexWrap: "wrap", alignItems: "center" }}>
        <input style={{ ...S.input, maxWidth: 260 }} placeholder="Buscar nombre o matrícula..."
          value={search} onChange={(e) => setSearch(e.target.value)} />
        <select style={S.select} value={fAsesor} onChange={(e) => setFAsesor(e.target.value)}>
          <option value="">Todos los asesores</option>
          {ASESORES.map((a) => <option key={a} value={a}>{a}</option>)}
        </select>
        <select style={S.select} value={fEstatus} onChange={(e) => setFEstatus(e.target.value)}>
          <option value="">Todos los estatus</option>
          {ESTATUSES.map((s) => <option key={s} value={s}>{s}</option>)}
        </select>
        <div style={{ marginLeft: "auto", display: "flex", gap: 8, alignItems: "center" }}>
          <span style={{ ...S.mono, fontSize: 11, color: "#6b6f82" }}>{data.length} registros</span>
          <Bt color="#8b5cf6" onClick={() => setShowPaste(true)}>⌨ Pegar Excel</Bt>
          <label style={{ ...S.btn("#10b981"), cursor: "pointer", display: "inline-block" }}>
            {importing ? "Importando..." : "↑ Importar archivo"}
            <input ref={inputRef} type="file" accept=".xlsx,.xls" style={{ display: "none" }}
              onChange={(e) => handleImport(e.target.files[0])} disabled={importing} />
          </label>
          <Bt color="#6366f1" onClick={() => dlXl(data.map((r) => ({ Día: r.dia, Hora: r.hora, Matrícula: r.matricula, Nombre: r.nombre, AP: r.ap, AM: r.am, Servicio: r.servicio, Clave: r.clave, Atiende: r.atiende, Escuela: r.escuela, Programa: r.programa, Estatus: r.estatus, Interés: r.interes_asesoria, Semestre: r.semestre, CAG: r.cag, EXATEC: r.exatec, Modalidad: r.modalidad, Campus: r.campus, Comunidad: r.comunidad, Celular: r.celular, Correo: r.correo_personal, LinkedIn: r.linkedin, Notas: r.notas })), "asesorias.xlsx")}>
            ↓ Exportar Excel
          </Bt>
        </div>
      </div>

      {showPaste && <PasteModal onClose={() => setShowPaste(false)} onImport={handlePasteImport} />}

      {importMsg && (
        <div style={{ marginBottom: 12, padding: "10px 16px", borderRadius: 10, fontSize: 13,
          background: importMsg.type === "ok" ? "rgba(16,185,129,0.1)" : "rgba(239,68,68,0.1)",
          border: `1px solid ${importMsg.type === "ok" ? "rgba(16,185,129,0.3)" : "rgba(239,68,68,0.3)"}`,
          color: importMsg.type === "ok" ? "#10b981" : "#ef4444" }}>
          {importMsg.text}
          <button onClick={() => setImportMsg(null)} style={{ float: "right", background: "none", border: "none", color: "inherit", cursor: "pointer" }}>✕</button>
        </div>
      )}

      {/* Grid */}
      <div style={{ overflowX: "auto", overflowY: "auto", maxHeight: "calc(100vh - 230px)", borderRadius: 12, border: "1px solid rgba(255,255,255,0.08)" }}>
        <table style={{ borderCollapse: "collapse", fontSize: 12, tableLayout: "fixed", minWidth: GRID_COLS.reduce((a, c) => a + c.width, 0) + 40 }}>
          <thead>
            <tr>
              {/* Delete col — frozen corner */}
              <th style={{ position: "sticky", top: 0, left: 0, zIndex: 21, width: 40, padding: "10px 8px", borderBottom: "1px solid rgba(255,255,255,0.1)", background: "#0a1525" }} />
              {GRID_COLS.map((col, colIdx) => {
                const frozen = colIdx < FROZEN_COUNT;
                const isLastFrozen = colIdx === FROZEN_COUNT - 1;
                return (
                  <th key={col.key} style={{
                    position: "sticky", top: 0,
                    left: frozen ? FROZEN_LEFT[col.key] : undefined,
                    zIndex: frozen ? 20 : 10,
                    width: col.width, padding: "10px 10px", textAlign: "left",
                    fontSize: 10, textTransform: "uppercase", letterSpacing: 0.8,
                    color: frozen ? "#a5b4fc" : "#6b6f82", fontWeight: 600,
                    borderBottom: "1px solid rgba(255,255,255,0.1)",
                    borderRight: isLastFrozen ? "2px solid rgba(99,102,241,0.3)" : undefined,
                    whiteSpace: "nowrap", overflow: "hidden", background: "#0a1525",
                  }}>
                    {col.label}{col.required && <span style={{ color: "#ef4444" }}> *</span>}
                  </th>
                );
              })}
            </tr>
          </thead>
          <tbody>
            {pagedItems.map(({ row, rowsIdx }) => {
              const isNew = row.id === null;
              const isSaving = saving.has(rowsIdx);
              const isHov = hoveredRow === rowsIdx;
              const frozenBg = isNew ? "#0d1424" : isHov ? "#10162e" : "#0b1120";
              const rowBg = isNew ? "rgba(99,102,241,0.04)" : isHov ? "rgba(99,102,241,0.06)" : "transparent";
              return (
                <tr key={row.id || "new"}
                  style={{ background: rowBg, transition: "background .1s" }}
                  onMouseEnter={() => setHoveredRow(rowsIdx)}
                  onMouseLeave={() => setHoveredRow(null)}>
                  {/* Delete / saving — frozen */}
                  <td style={{ position: "sticky", left: 0, zIndex: 2, background: frozenBg, width: 40, padding: "4px 6px", borderBottom: "1px solid rgba(255,255,255,0.04)", textAlign: "center", transition: "background .1s" }}>
                    {isSaving ? (
                      <span style={{ display: "inline-block", width: 12, height: 12, border: "2px solid rgba(99,102,241,0.3)", borderTopColor: "#6366f1", borderRadius: "50%", animation: "spin 0.8s linear infinite" }} />
                    ) : !isNew ? (
                      <button onClick={() => deleteRow(rowsIdx)} title="Eliminar"
                        style={{ background: "none", border: "none", color: "#6b6f82", cursor: "pointer", padding: 2, fontSize: 13, lineHeight: 1 }}
                        onMouseEnter={(e) => e.target.style.color = "#ef4444"}
                        onMouseLeave={(e) => e.target.style.color = "#6b6f82"}>✕</button>
                    ) : (
                      <span style={{ color: "#6b6f82", fontSize: 10 }}>+</span>
                    )}
                  </td>
                  {GRID_COLS.map((col, colIdx) => {
                    const frozen = colIdx < FROZEN_COUNT;
                    const isLastFrozen = colIdx === FROZEN_COUNT - 1;
                    const isEditing = editCell?.rowIdx === rowsIdx && editCell?.key === col.key;
                    const val = row[col.key] ?? "";
                    const isEmpty = !val;
                    return (
                      <td key={col.key}
                        style={{
                          position: frozen ? "sticky" : undefined,
                          left: frozen ? FROZEN_LEFT[col.key] : undefined,
                          zIndex: frozen ? 2 : undefined,
                          background: frozen ? frozenBg : undefined,
                          transition: frozen ? "background .1s" : undefined,
                          padding: "3px 4px",
                          borderBottom: "1px solid rgba(255,255,255,0.04)",
                          borderRight: isLastFrozen ? "2px solid rgba(99,102,241,0.3)" : "1px solid rgba(255,255,255,0.03)",
                          maxWidth: col.width, overflow: "hidden",
                        }}
                        onClick={() => !col.readOnly && setEditCell({ rowIdx: rowsIdx, key: col.key })}>
                        {isEditing && !col.readOnly ? (
                          col.type === "select" ? (
                            <select autoFocus value={val}
                              onChange={(e) => updateRowLocal(rowsIdx, col.key, e.target.value)}
                              onBlur={() => handleCellBlur(rowsIdx)}
                              onKeyDown={(e) => handleKeyDown(e, rowsIdx, colIdx)}
                              style={{ ...S.select, width: "100%", fontSize: 12, padding: "4px 6px", borderRadius: 6, boxSizing: "border-box" }}>
                              <option value="">—</option>
                              {col.options.map((o) => <option key={o} value={o}>{o}</option>)}
                            </select>
                          ) : (
                            <input autoFocus
                              type={col.type === "date" ? "date" : "text"}
                              value={val}
                              onChange={(e) => updateRowLocal(rowsIdx, col.key, e.target.value)}
                              onBlur={() => handleCellBlur(rowsIdx)}
                              onKeyDown={(e) => handleKeyDown(e, rowsIdx, colIdx)}
                              style={{ ...S.input, fontSize: 12, padding: "4px 6px", borderRadius: 6, boxSizing: "border-box", background: "#0d1e38" }}
                            />
                          )
                        ) : (
                          <div style={{
                            padding: "5px 6px", minHeight: 28, fontSize: 12, borderRadius: 6,
                            color: isEmpty && !isNew ? "#3a3f5a" : isEmpty && isNew ? "#4a5080"
                              : col.key === "matricula" ? "#a5b4fc"
                              : col.key === "estatus" ? (STATUS_COLORS[val] || "#e8e9ed")
                              : "#e8e9ed",
                            fontFamily: col.key === "matricula" || col.key === "clave" ? "'JetBrains Mono', monospace" : "inherit",
                            cursor: col.readOnly ? "default" : "text",
                            overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap",
                            border: "1px solid transparent",
                          }}
                            onMouseEnter={(e) => { if (!col.readOnly) e.currentTarget.style.border = "1px solid rgba(99,102,241,0.3)"; }}
                            onMouseLeave={(e) => { e.currentTarget.style.border = "1px solid transparent"; }}>
                            {isEmpty ? (isNew && col.key === "matricula" ? "Nueva asesoría..." : col.placeholder || "") : val}
                          </div>
                        )}
                      </td>
                    );
                  })}
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>

      {/* Paginación */}
      {totalGridPages > 1 && (
        <div style={{ display: "flex", alignItems: "center", gap: 10, padding: "12px 4px" }}>
          <Bt color="#6366f1" onClick={() => setGridPage((p) => Math.max(0, p - 1))}
            style={{ padding: "4px 14px", fontSize: 12, opacity: gridPage === 0 ? 0.3 : 1, pointerEvents: gridPage === 0 ? "none" : "auto" }}>← Ant</Bt>
          <span style={{ ...S.mono, fontSize: 12, color: "#8e92a6" }}>
            Página {gridPage + 1} de {totalGridPages} · {existingItems.length} registros
          </span>
          <Bt color="#6366f1" onClick={() => setGridPage((p) => Math.min(totalGridPages - 1, p + 1))}
            style={{ padding: "4px 14px", fontSize: 12, opacity: gridPage === totalGridPages - 1 ? 0.3 : 1, pointerEvents: gridPage === totalGridPages - 1 ? "none" : "auto" }}>Sig →</Bt>
        </div>
      )}

      <div style={{ marginTop: 8, color: "#6b6f82", fontSize: 11 }}>
        Clic en celda para editar · Tab / Enter para avanzar · se guarda automáticamente al salir
      </div>
    </div>
  );
}

function parseExcelDate(v) {
  if (!v) return null;
  if (typeof v === "number") {
    const d = new Date(Math.round((v - 25569) * 86400000));
    if (isNaN(d)) return null;
    return d.toISOString().slice(0, 10);
  }
  const s = String(v).trim();
  const d = new Date(s);
  if (isNaN(d)) return null;
  return d.toISOString().slice(0, 10);
}

/* ═══ TAB DASHBOARD ═══ */
function TabDashboard({ data }) {
  const total = data.length;
  const asist = data.filter((r) => r.estatus === "Asistencia").length;
  const faltas = data.filter((r) => r.estatus === "Falta").length;
  const express = data.filter((r) => r.estatus === "Express").length;
  const cancel = data.filter((r) => r.estatus === "Cancelación").length;
  const uniq = new Set(data.map((r) => r.matricula)).size;
  const base = total - express;

  const cagsUniq = useMemo(() => new Set(data.filter(isCAGS).map((r) => r.matricula)).size, [data]);
  const dic25Uniq = useMemo(() => new Set(data.filter(isDIC25).map((r) => r.matricula)).size, [data]);

  const weekData = useMemo(() => {
    const wm = {};
    data.forEach((r) => {
      if (!r.dia) return;
      const d = new Date(r.dia + "T12:00:00");
      const jan1 = new Date(d.getFullYear(), 0, 1);
      const w = Math.ceil((((d - jan1) / 86400000) + jan1.getDay() + 1) / 7);
      wm[w] = (wm[w] || 0) + 1;
    });
    return Object.entries(wm).sort((a, b) => +a[0] - +b[0]).map(([w, v]) => ({ name: `S${w}`, value: v }));
  }, [data]);

  const monthData = useMemo(() => {
    const m = {};
    data.forEach((r) => {
      if (!r.dia) return;
      const k = new Date(r.dia + "T12:00:00").toLocaleDateString("es-MX", { month: "short", year: "numeric" });
      m[k] = (m[k] || 0) + 1;
    });
    return Object.entries(m).map(([name, value]) => ({ name, value }));
  }, [data]);

  const statusData = [
    { name: "Asistencia", value: asist, color: "#10b981" },
    { name: "Falta", value: faltas, color: "#ef4444" },
    { name: "Express", value: express, color: "#f59e0b" },
    { name: "Cancelación", value: cancel, color: "#8b5cf6" },
  ];
  const escData = countBy(data, "escuela");
  const intData = countBy(data, "interes_asesoria");
  const srvData = countBy(data, "servicio");
  const modData = countBy(data, "modalidad");

  const ref1 = useRef(), ref2 = useRef();

  return (
    <div>
      <div style={S.grid(5)}>
        <KPI label="Total asesorías" value={total.toLocaleString()} color="#6366f1" sub={uniq ? `${(total / uniq).toFixed(1)} por alumno` : ""} />
        <KPI label="Alumnos únicos" value={uniq.toLocaleString()} color="#3b82f6" />
        <KPI label="Tasa asistencia" value={base ? `${((asist / base) * 100).toFixed(1)}%` : "—"} color="#10b981" sub={`${asist} asistencias`} />
        <KPI label="Faltas" value={faltas} color="#ef4444" sub={base ? `${((faltas / base) * 100).toFixed(1)}%` : ""} />
        <KPI label="Express" value={express} color="#f59e0b" />
      </div>
      <div style={{ ...S.grid(2), marginBottom: 4 }}>
        <KPI label="CAGS — Candidatos JUN 2026" value={cagsUniq} color="#a855f7" sub={uniq ? `${((cagsUniq / uniq) * 100).toFixed(1)}% de alumnos únicos` : ""} />
        <KPI label="DIC 2025 — Generación graduada" value={dic25Uniq} color="#22d3ee" sub={uniq ? `${((dic25Uniq / uniq) * 100).toFixed(1)}% de alumnos únicos` : ""} />
      </div>

      <div style={S.grid(2)}>
        <Cd>
          <div style={S.h3}>Asesorías por semana</div>
          <div ref={ref1}><ResponsiveContainer width="100%" height={240}>
            <BarChart data={weekData} margin={{ top: 5, right: 20, bottom: 5, left: 0 }}>
              <XAxis dataKey="name" tick={{ fill: "#6b6f82", fontSize: 10 }} />
              <YAxis tick={{ fill: "#6b6f82", fontSize: 10 }} />
              <Tooltip content={<TT />} />
              <Bar dataKey="value" fill="#6366f1" radius={[6, 6, 0, 0]} />
            </BarChart>
          </ResponsiveContainer></div>
        </Cd>
        <Cd>
          <div style={S.h3}>Tendencia mensual</div>
          <div ref={ref2}><ResponsiveContainer width="100%" height={240}>
            <LineChart data={monthData} margin={{ top: 5, right: 20, bottom: 5, left: 0 }}>
              <CartesianGrid strokeDasharray="3 3" stroke="rgba(255,255,255,0.06)" />
              <XAxis dataKey="name" tick={{ fill: "#6b6f82", fontSize: 10 }} />
              <YAxis tick={{ fill: "#6b6f82", fontSize: 10 }} />
              <Tooltip content={<TT />} />
              <Line type="monotone" dataKey="value" stroke="#6366f1" strokeWidth={2.5} dot={{ fill: "#6366f1", r: 4 }} />
            </LineChart>
          </ResponsiveContainer></div>
        </Cd>
      </div>

      <div style={S.grid(3)}>
        <Cd>
          <div style={S.h3}>Estatus</div>
          <SB items={statusData} total={total} />
        </Cd>
        <Cd>
          <div style={S.h3}>Escuelas</div>
          <SB items={escData.map((d, i) => ({ ...d, color: CHART_COLORS[i % CHART_COLORS.length] }))} total={total} />
        </Cd>
        <Cd>
          <div style={S.h3}>Interés de asesoría</div>
          <SB items={intData.map((d, i) => ({ ...d, color: CHART_COLORS[i % CHART_COLORS.length] }))} total={total} />
        </Cd>
      </div>

      <div style={S.grid(2)}>
        <Cd>
          <div style={S.h3}>Servicios</div>
          <ResponsiveContainer width="100%" height={Math.max(200, srvData.length * 26)}>
            <BarChart data={srvData} layout="vertical" margin={{ left: 160, right: 20, top: 5, bottom: 5 }}>
              <XAxis type="number" tick={{ fill: "#6b6f82", fontSize: 10 }} />
              <YAxis type="category" dataKey="name" tick={{ fill: "#8e92a6", fontSize: 10 }} width={155} />
              <Tooltip content={<TT />} />
              <Bar dataKey="value" radius={[0, 6, 6, 0]}>{srvData.map((_, i) => <Cell key={i} fill={CHART_COLORS[i % CHART_COLORS.length]} />)}</Bar>
            </BarChart>
          </ResponsiveContainer>
        </Cd>
        <Cd>
          <div style={S.h3}>Modalidad</div>
          <ResponsiveContainer width="100%" height={220}>
            <PieChart>
              <Pie data={modData} dataKey="value" nameKey="name" cx="50%" cy="50%" outerRadius={80} innerRadius={35} paddingAngle={5}>
                {modData.map((_, i) => <Cell key={i} fill={CHART_COLORS[i % CHART_COLORS.length]} />)}
              </Pie>
              <Tooltip content={<TT />} />
              <Legend formatter={(v) => <span style={{ color: "#8e92a6", fontSize: 11 }}>{v}</span>} />
            </PieChart>
          </ResponsiveContainer>
        </Cd>
      </div>
    </div>
  );
}

/* ═══ TAB ASESORES ═══ */
function AsesorModal({ asesor, records, onClose }) {
  const [selectedStudent, setSelectedStudent] = useState(null);
  const [aSearch, setASearch] = useState("");
  const dASearch = useDebounce(aSearch, 200);

  const asist = records.filter((r) => r.estatus === "Asistencia").length;
  const faltas = records.filter((r) => r.estatus === "Falta").length;
  const express = records.filter((r) => r.estatus === "Express").length;
  const base = records.length - express;

  const alumnosList = useMemo(() => {
    const m = {};
    records.forEach((r) => {
      if (!m[r.matricula]) m[r.matricula] = { matricula: r.matricula, nombre: r.nombre, ap: r.ap, sesiones: 0, asistencias: 0, records: [], escuela: r.escuela, programa: r.programa };
      const s = m[r.matricula];
      s.sesiones++;
      s.records.push(r);
      if (r.estatus === "Asistencia") s.asistencias++;
    });
    return Object.values(m).sort((a, b) => b.sesiones - a.sesiones);
  }, [records]);

  const alumnosFiltrados = useMemo(() => {
    if (!dASearch) return alumnosList;
    const q = norm(dASearch);
    return alumnosList.filter((a) => norm(a.nombre).includes(q) || norm(a.matricula).includes(q));
  }, [alumnosList, dASearch]);

  return (
    <>
      {selectedStudent && (
        <StudentModal matricula={selectedStudent.matricula} records={selectedStudent.records} onClose={() => setSelectedStudent(null)} />
      )}
      <Modal onClose={onClose}>
        <div style={{ fontSize: 20, fontWeight: 700, marginBottom: 20 }}>{asesor}</div>
        <div style={S.grid(5)}>
          <KPI label="Total" value={records.length} color="#6366f1" />
          <KPI label="Asistencias" value={asist} color="#10b981" sub={base ? `${((asist / base) * 100).toFixed(1)}%` : "—"} />
          <KPI label="Faltas" value={faltas} color="#ef4444" sub={base ? `${((faltas / base) * 100).toFixed(1)}%` : "—"} />
          <KPI label="Alumnos únicos" value={alumnosList.length} color="#8b5cf6" />
          <KPI label="CAGS atendidos" value={new Set(records.filter(isCAGS).map((r) => r.matricula)).size} color="#a855f7" />
        </div>
        <div style={{ marginTop: 20 }}>
          <div style={{ display: "flex", gap: 8, marginBottom: 10, alignItems: "center" }}>
            <div style={S.h3}>Alumnos atendidos ({alumnosFiltrados.length})</div>
            <input style={{ ...S.input, maxWidth: 200, padding: "6px 10px", marginLeft: "auto" }} placeholder="Buscar alumno..."
              value={aSearch} onChange={(e) => setASearch(e.target.value)} />
          </div>
          <div style={{ overflowX: "auto", maxHeight: 320, overflowY: "auto" }}>
            <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
              <thead style={{ position: "sticky", top: 0, background: "#0f1628", zIndex: 1 }}>
                <tr style={{ borderBottom: "1px solid rgba(255,255,255,0.08)" }}>
                  {["Alumno", "Matrícula", "Sesiones", "Asistencias", "Escuela"].map((h) => (
                    <th key={h} style={{ textAlign: "left", padding: "8px 10px", fontSize: 10, textTransform: "uppercase", letterSpacing: 1, color: "#6b6f82", fontWeight: 600 }}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {alumnosFiltrados.map((a, i) => (
                  <tr key={i} style={{ borderBottom: "1px solid rgba(255,255,255,0.04)", cursor: "pointer" }}
                    onClick={() => setSelectedStudent(a)}
                    onMouseEnter={(e) => e.currentTarget.style.background = "rgba(99,102,241,0.06)"}
                    onMouseLeave={(e) => e.currentTarget.style.background = "transparent"}>
                    <td style={{ padding: "8px 10px", fontWeight: 500 }}>{[a.nombre, a.ap].filter(Boolean).join(" ")}</td>
                    <td style={{ padding: "8px 10px", ...S.mono, fontSize: 11, color: "#a5b4fc" }}>{a.matricula}</td>
                    <td style={{ padding: "8px 10px", textAlign: "center" }}><span style={S.badge("#6366f1")}>{a.sesiones}</span></td>
                    <td style={{ padding: "8px 10px", textAlign: "center" }}><span style={S.badge("#10b981")}>{a.asistencias}</span></td>
                    <td style={{ padding: "8px 10px", fontSize: 11, color: "#8e92a6" }}>{a.escuela}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      </Modal>
    </>
  );
}

function TabAsesores({ data }) {
  const [selected, setSelected] = useState(null);
  const byAsesor = useMemo(() => {
    const m = {};
    data.forEach((r) => { (m[r.atiende] = m[r.atiende] || []).push(r); });
    return Object.entries(m).sort((a, b) => b[1].length - a[1].length);
  }, [data]);

  return (
    <div>
      {selected && <AsesorModal asesor={selected} records={data.filter((r) => r.atiende === selected)} onClose={() => setSelected(null)} />}
      <div style={S.grid(2)}>
        {byAsesor.map(([name, recs], idx) => {
          const a = recs.filter((r) => r.estatus === "Asistencia").length;
          const ex = recs.filter((r) => r.estatus === "Express").length;
          const f = recs.filter((r) => r.estatus === "Falta").length;
          const b = recs.length - ex;
          const color = CHART_COLORS[idx % CHART_COLORS.length];
          const cagsCount = new Set(recs.filter(isCAGS).map((r) => r.matricula)).size;
          const dic25Count = new Set(recs.filter(isDIC25).map((r) => r.matricula)).size;
          return (
            <Cd key={name} style={{ cursor: "pointer", borderLeft: `4px solid ${color}` }} onClick={() => setSelected(name)}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 10 }}>
                <div style={{ fontSize: 16, fontWeight: 700 }}>{name}</div>
                <span style={{ ...S.mono, fontSize: 22, fontWeight: 700, color }}>{recs.length}</span>
              </div>
              <div style={{ display: "flex", gap: 16, marginBottom: 8, fontSize: 12, color: "#8e92a6" }}>
                <span>Asistencia: <b style={{ color: "#10b981" }}>{b ? `${((a / b) * 100).toFixed(0)}%` : "—"}</b></span>
                <span>Alumnos únicos: <b style={{ color: "#a5b4fc" }}>{new Set(recs.map((r) => r.matricula)).size}</b></span>
              </div>
              <div style={{ display: "flex", gap: 6, marginBottom: 8 }}>
                {cagsCount > 0 && <span style={S.badge("#a855f7")}>CAGS: {cagsCount}</span>}
                {dic25Count > 0 && <span style={S.badge("#22d3ee")}>DIC25: {dic25Count}</span>}
              </div>
              <div style={{ display: "flex", height: 6, borderRadius: 3, overflow: "hidden", gap: 2 }}>
                {a > 0 && <div style={{ flex: a, background: "#10b981", borderRadius: 3 }} />}
                {f > 0 && <div style={{ flex: f, background: "#ef4444", borderRadius: 3 }} />}
                {ex > 0 && <div style={{ flex: ex, background: "#f59e0b", borderRadius: 3 }} />}
              </div>
            </Cd>
          );
        })}
      </div>
    </div>
  );
}

/* ═══ TAB PIPELINE ═══ */
function TabPipeline({ data }) {
  const srvData = useMemo(() => countBy(data, "servicio"), [data]);
  const top5 = srvData.slice(0, 5);
  const maxVal = top5[0]?.value || 1;

  return (
    <div>
      <div style={S.grid(5)}>
        {top5.map((s, i) => (
          <KPI key={s.name} label={s.name} value={s.value} color={CHART_COLORS[i]}
            sub={`${data.length ? ((s.value / data.length) * 100).toFixed(1) : 0}% del total`} />
        ))}
      </div>
      <Cd style={{ marginTop: 16 }}>
        <div style={S.h3}>Funnel de servicios</div>
        <div style={{ display: "flex", gap: 8, alignItems: "flex-end", height: 200, padding: "20px 10px" }}>
          {srvData.map((s, i) => (
            <div key={s.name} style={{ flex: 1, display: "flex", flexDirection: "column", alignItems: "center", gap: 6 }}>
              <span style={{ ...S.mono, fontSize: 10, color: "#8e92a6" }}>{s.value}</span>
              <div style={{ width: "100%", maxWidth: 60, height: `${(s.value / maxVal) * 160}px`, background: `linear-gradient(180deg, ${CHART_COLORS[i % CHART_COLORS.length]}, ${CHART_COLORS[i % CHART_COLORS.length]}66)`, borderRadius: "8px 8px 4px 4px", minHeight: 4 }} />
              <span style={{ fontSize: 8, color: "#6b6f82", textAlign: "center", lineHeight: 1.2, maxWidth: 60 }}>{s.name}</span>
            </div>
          ))}
        </div>
      </Cd>
      <Cd>
        <div style={S.h3}>Todos los servicios</div>
        <SB items={srvData.map((d, i) => ({ ...d, color: CHART_COLORS[i % CHART_COLORS.length] }))} total={data.length} />
      </Cd>
    </div>
  );
}

/* ═══ TAB ALUMNOS ═══ */
const PAGE_SIZE = 100;

function TabAlumnos({ data }) {
  const [search, setSearch] = useState("");
  const [fAsesor, setFAsesor] = useState("");
  const [fEscuela, setFEscuela] = useState("");
  const [fEstatus, setFEstatus] = useState("");
  const [fCAGS, setFCAGS] = useState(false);
  const [fDIC25, setFDIC25] = useState(false);
  const [selectedStudent, setSelectedStudent] = useState(null);
  const [page, setPage] = useState(0);
  const dSearch = useDebounce(search, 200);

  const students = useMemo(() => {
    const m = {};
    data.forEach((r) => {
      if (!m[r.matricula]) m[r.matricula] = { matricula: r.matricula, records: [] };
      m[r.matricula].records.push(r);
    });
    return Object.values(m).map((s) => {
      const sorted = [...s.records].sort((a, b) => (b.dia || "") > (a.dia || "") ? 1 : -1);
      const latest = sorted[0] || {};
      return {
        matricula: s.matricula,
        nombre: sorted.find((r) => r.nombre)?.nombre || "—",
        ap: sorted.find((r) => r.ap)?.ap || "",
        sesiones: s.records.length,
        asistencias: s.records.filter((r) => r.estatus === "Asistencia").length,
        ultimoAsesor: latest.atiende || "—",
        ultimoServicio: latest.servicio || "—",
        escuela: latest.escuela || "—",
        programa: latest.programa || "—",
        isCAGS: s.records.some(isCAGS),
        isDIC25: s.records.some(isDIC25),
        records: s.records,
      };
    });
  }, [data]);

  const asesores = useMemo(() => [...new Set(data.map((r) => r.atiende).filter(Boolean))].sort(), [data]);
  const escuelas = useMemo(() => [...new Set(data.map((r) => r.escuela).filter(Boolean))].sort(), [data]);
  const estatuses = useMemo(() => [...new Set(data.map((r) => r.estatus).filter(Boolean))].sort(), [data]);

  const filtered = useMemo(() => {
    let f = students;
    if (dSearch) { const q = norm(dSearch); f = f.filter((s) => norm(s.nombre).includes(q) || norm(s.matricula).includes(q) || norm(s.ap).includes(q)); }
    if (fAsesor) f = f.filter((s) => s.records.some((r) => r.atiende === fAsesor));
    if (fEscuela) f = f.filter((s) => s.escuela === fEscuela);
    if (fEstatus) f = f.filter((s) => s.records.some((r) => r.estatus === fEstatus));
    if (fCAGS) f = f.filter((s) => s.isCAGS);
    if (fDIC25) f = f.filter((s) => s.isDIC25);
    return f.sort((a, b) => b.sesiones - a.sesiones);
  }, [students, dSearch, fAsesor, fEscuela, fEstatus, fCAGS, fDIC25]);

  useEffect(() => setPage(0), [dSearch, fAsesor, fEscuela, fEstatus, fCAGS, fDIC25]);

  const totalPages = Math.ceil(filtered.length / PAGE_SIZE);
  const displayed = dSearch ? filtered : filtered.slice(page * PAGE_SIZE, (page + 1) * PAGE_SIZE);

  return (
    <div>
      {selectedStudent && <StudentModal matricula={selectedStudent.matricula} records={selectedStudent.records} onClose={() => setSelectedStudent(null)} />}
      <div style={{ display: "flex", gap: 8, marginBottom: 14, flexWrap: "wrap", alignItems: "center" }}>
        <input style={{ ...S.input, maxWidth: 280 }} placeholder="Buscar nombre o matrícula..." value={search} onChange={(e) => setSearch(e.target.value)} />
        <select style={S.select} value={fAsesor} onChange={(e) => setFAsesor(e.target.value)}>
          <option value="">Todos los asesores</option>
          {asesores.map((a) => <option key={a} value={a}>{a}</option>)}
        </select>
        <select style={S.select} value={fEscuela} onChange={(e) => setFEscuela(e.target.value)}>
          <option value="">Todas las escuelas</option>
          {escuelas.map((e) => <option key={e} value={e}>{e}</option>)}
        </select>
        <select style={S.select} value={fEstatus} onChange={(e) => setFEstatus(e.target.value)}>
          <option value="">Todos los estatus</option>
          {estatuses.map((s) => <option key={s} value={s}>{s}</option>)}
        </select>
        <button onClick={() => { setFCAGS((v) => !v); setFDIC25(false); }}
          style={{ ...S.btn(fCAGS ? "#a855f7" : "#8e92a6"), fontSize: 11, padding: "5px 12px", opacity: fDIC25 ? 0.4 : 1 }}>
          ★ Solo CAGS
        </button>
        <button onClick={() => { setFDIC25((v) => !v); setFCAGS(false); }}
          style={{ ...S.btn(fDIC25 ? "#22d3ee" : "#8e92a6"), fontSize: 11, padding: "5px 12px", opacity: fCAGS ? 0.4 : 1 }}>
          ✓ Solo DIC25
        </button>
        <Bt color="#8e92a6" onClick={() => { setSearch(""); setFAsesor(""); setFEscuela(""); setFEstatus(""); setFCAGS(false); setFDIC25(false); }} style={{ fontSize: 11 }}>Limpiar</Bt>
        <span style={{ ...S.mono, fontSize: 11, color: "#6b6f82", marginLeft: "auto" }}>{filtered.length} alumnos</span>
      </div>

      <div style={{ overflowX: "auto" }}>
        <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
          <thead>
            <tr style={{ borderBottom: "1px solid rgba(255,255,255,0.08)" }}>
              {["Matrícula","Nombre","Sesiones","Asistencias","Último asesor","Último servicio","Escuela","Programa"].map((h) => (
                <th key={h} style={{ textAlign: "left", padding: "8px 10px", fontSize: 10, textTransform: "uppercase", letterSpacing: 1, color: "#6b6f82", fontWeight: 600, whiteSpace: "nowrap" }}>{h}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {displayed.map((s) => (
              <tr key={s.matricula} style={{ borderBottom: "1px solid rgba(255,255,255,0.04)", cursor: "pointer" }}
                onClick={() => setSelectedStudent(s)}
                onMouseEnter={(e) => e.currentTarget.style.background = "rgba(99,102,241,0.06)"}
                onMouseLeave={(e) => e.currentTarget.style.background = "transparent"}>
                <td style={{ padding: "8px 10px", ...S.mono, fontSize: 11, color: "#a5b4fc" }}>{s.matricula}</td>
                <td style={{ padding: "8px 10px", fontWeight: 500 }}>
                  {[s.nombre, s.ap].filter(Boolean).join(" ")}
                  {s.isCAGS && <span style={{ ...S.badge("#a855f7"), marginLeft: 6, fontSize: 9 }}>CAGS</span>}
                  {s.isDIC25 && <span style={{ ...S.badge("#22d3ee"), marginLeft: 6, fontSize: 9 }}>DIC25</span>}
                </td>
                <td style={{ padding: "8px 10px", textAlign: "center" }}><span style={S.badge("#6366f1")}>{s.sesiones}</span></td>
                <td style={{ padding: "8px 10px", textAlign: "center" }}><span style={S.badge("#10b981")}>{s.asistencias}</span></td>
                <td style={{ padding: "8px 10px", fontSize: 11, color: "#8e92a6" }}>{s.ultimoAsesor}</td>
                <td style={{ padding: "8px 10px", fontSize: 11, maxWidth: 180, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{s.ultimoServicio}</td>
                <td style={{ padding: "8px 10px", fontSize: 11, color: "#8e92a6" }}>{s.escuela}</td>
                <td style={{ padding: "8px 10px" }}><span style={S.badge("#3b82f6")}>{s.programa}</span></td>
              </tr>
            ))}
          </tbody>
        </table>
        {filtered.length === 0 && (
          <div style={{ textAlign: "center", padding: 40, color: "#6b6f82", fontSize: 13 }}>Sin resultados</div>
        )}
        {!dSearch && totalPages > 1 && (
          <div style={{ display: "flex", justifyContent: "center", alignItems: "center", gap: 10, padding: 16 }}>
            <Bt color="#6366f1" onClick={() => setPage((p) => Math.max(0, p - 1))} style={{ padding: "4px 14px", fontSize: 12, opacity: page === 0 ? 0.3 : 1, pointerEvents: page === 0 ? "none" : "auto" }}>← Ant</Bt>
            <span style={{ ...S.mono, fontSize: 12, color: "#8e92a6" }}>Página {page + 1} de {totalPages} · {filtered.length} alumnos</span>
            <Bt color="#6366f1" onClick={() => setPage((p) => Math.min(totalPages - 1, p + 1))} style={{ padding: "4px 14px", fontSize: 12, opacity: page === totalPages - 1 ? 0.3 : 1, pointerEvents: page === totalPages - 1 ? "none" : "auto" }}>Sig →</Bt>
          </div>
        )}
      </div>
    </div>
  );
}

/* ═══ MAIN CRM ═══ */
export default function CRM() {
  const [data, setData] = useState([]);
  const [loading, setLoading] = useState(true);
  const [tab, setTab] = useState("home");

  const fetchData = useCallback(async () => {
    const { data: rows, error } = await supabase
      .from("asesorias")
      .select("*")
      .order("dia", { ascending: false })
      .order("hora", { ascending: true });
    if (!error) setData(rows || []);
    setLoading(false);
  }, []);

  useEffect(() => { fetchData(); }, [fetchData]);

  const handleStatusChange = async (id, newStatus) => {
    await supabase.from("asesorias").update({ estatus: newStatus }).eq("id", id);
    setData((prev) => prev.map((r) => r.id === id ? { ...r, estatus: newStatus } : r));
  };

  if (loading) {
    return (
      <div style={{ minHeight: "100vh", display: "flex", alignItems: "center", justifyContent: "center", background: "#0b1120" }}>
        <div style={{ textAlign: "center" }}>
          <div style={{ width: 48, height: 48, border: "4px solid rgba(99,102,241,0.2)", borderTopColor: "#6366f1", borderRadius: "50%", animation: "spin 0.8s linear infinite", margin: "0 auto 16px" }} />
          <div style={{ color: "#8e92a6", fontSize: 14 }}>Cargando asesorías…</div>
          <style>{`@keyframes spin { to { transform: rotate(360deg); } }`}</style>
        </div>
      </div>
    );
  }

  return (
    <div style={{ minHeight: "100vh" }}>
      <header style={{ position: "sticky", top: 0, zIndex: 100, background: "rgba(11,17,32,0.92)", backdropFilter: "blur(16px)", borderBottom: "1px solid rgba(255,255,255,0.06)", padding: "0 24px" }}>
        <div style={{ display: "flex", alignItems: "center", height: 56, maxWidth: 1400, margin: "0 auto" }}>
          <div style={{ display: "flex", alignItems: "center", gap: 10, marginRight: 32 }}>
            <div style={{ width: 30, height: 30, borderRadius: 8, background: "linear-gradient(135deg, #6366f1, #8b5cf6)", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 14, fontWeight: 800, color: "#fff" }}>C</div>
            <span style={{ fontWeight: 700, fontSize: 14 }}>CVDP <span style={{ color: "#8e92a6", fontWeight: 400 }}>Empleabilidad</span></span>
          </div>
          <nav style={{ display: "flex", gap: 0, flex: 1 }}>
            {TABS.map((t) => (
              <button key={t.id} onClick={() => setTab(t.id)}
                style={{ background: "none", border: "none", borderBottom: tab === t.id ? "2px solid #6366f1" : "2px solid transparent", color: tab === t.id ? "#a5b4fc" : "#7d8296", padding: "16px 16px", fontSize: 13, fontWeight: tab === t.id ? 600 : 400, cursor: "pointer", fontFamily: "'Plus Jakarta Sans'", transition: "all .2s", display: "flex", alignItems: "center", gap: 6 }}>
                <span style={{ fontSize: 14 }}>{t.icon}</span> {t.label}
              </button>
            ))}
          </nav>
          <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
            <span style={{ ...S.mono, fontSize: 11, color: "#6b6f82" }}>{data.length.toLocaleString()} registros</span>
            <Bt color="#6366f1" onClick={fetchData} style={{ fontSize: 11, padding: "5px 12px" }}>↻ Actualizar</Bt>
          </div>
        </div>
      </header>

      <main style={{ maxWidth: 1400, margin: "0 auto", padding: "24px 24px 60px" }}>
        {tab === "home"      && <TabHome data={data} onStatusChange={handleStatusChange} />}
        {tab === "asesorias" && <TabAsesorias data={data} onRefresh={fetchData} />}
        {tab === "dashboard" && <TabDashboard data={data} />}
        {tab === "asesores"  && <TabAsesores data={data} />}
        {tab === "pipeline"  && <TabPipeline data={data} />}
        {tab === "alumnos"   && <TabAlumnos data={data} />}
      </main>
    </div>
  );
}
