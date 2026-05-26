import { useState, useMemo, useCallback, useRef, useEffect } from "react";
import * as XLSX from "xlsx";
import {
  BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer, PieChart, Pie, Cell,
  LineChart, Line, CartesianGrid, Legend
} from "recharts";

/* ═══════════════ CONSTANTS ═══════════════ */
const CHART_COLORS = ["#6366f1","#10b981","#f59e0b","#ef4444","#8b5cf6","#ec4899","#f97316","#14b8a6","#3b82f6","#a855f7","#84cc16","#06b6d4","#e11d48","#22d3ee"];
const STATUS_COLORS = { Asistencia:"#10b981", Falta:"#ef4444", Express:"#f59e0b", "Cancelación":"#8b5cf6" };
const TABS = [
  { id:"dashboard", icon:"◈", label:"Dashboard" },
  { id:"asesores", icon:"◎", label:"Asesores" },
  { id:"pipeline", icon:"◐", label:"Pipeline" },
  { id:"alumnos", icon:"◇", label:"Alumnos" },
  { id:"custom", icon:"◆", label:"Personalizado" }
];
const DIMS = ["servicio","asesor","escuela","programa","estatus","interes","modalidad","comunidad","semestre"];

/* ═══════════════ NORMALIZATION ═══════════════ */
function nSrv(s) {
  if (!s) return "Sin servicio";
  const l = s.toLowerCase();
  if (l.includes("diagnóstico") || l.includes("diagnostico")) {
    if (l.includes("cv")) return "Diagnóstico CV";
    if (l.includes("linkedin")) return "Diagnóstico LinkedIn";
    if (l.includes("bolsa")) return "Diagnóstico Bolsa de Trabajo";
    return "Diagnóstico";
  }
  if (l.includes("cv")) return "Asesoría CV";
  if (l.includes("linkedin")) return "Asesoría LinkedIn";
  if (l.includes("bolsa")) return "Asesoría Bolsa de Trabajo";
  if (l.includes("entrevista") && l.includes("español")) return "Entrevista Español";
  if (l.includes("entrevista") && (l.includes("inglés") || l.includes("ingl"))) return "Entrevista Inglés";
  if (l.includes("individual")) return "Asesoría Individual";
  if (l.includes("carta") || l.includes("oferta")) return "Carta Oferta";
  if (l.includes("plan de vida")) return "Plan de Vida y Carrera";
  if (l.includes("cover letter")) return "Cover Letter";
  if (l.includes("portafolio")) return "Portafolio";
  return s.trim();
}
function nEsc(s) {
  if (!s) return "Sin escuela";
  const l = s.toLowerCase();
  if (l.includes("ingeniería") || l.includes("ingenieria")) return "Ingeniería y Ciencias";
  if (l.includes("negocio")) return "Negocios";
  if (l.includes("humanidades")) return "Humanidades y Educación";
  if (l.includes("sociales")) return "Ciencias Sociales y Gobierno";
  if (l.includes("arquitectura")) return "Arquitectura, Arte y Diseño";
  if (l.includes("medicina")) return "Medicina y Ciencias de la Salud";
  return s.trim();
}
function nInt(s) {
  if (!s) return "Sin interés";
  const l = s.toLowerCase();
  if (l.includes("empleo") && l.includes("internacional")) return "Empleo Internacional";
  if (l.startsWith("empleo")) return "Empleo";
  if (l.includes("prácticas") || l.includes("practicas")) return "Prácticas Profesionales";
  if (l.includes("estancia")) return "Estancia Profesional";
  if (l.includes("on campus")) return "On Campus Intern";
  if (l.includes("posgrado")) return "Posgrado";
  if (l.includes("programa internacional")) return "Programa Internacional";
  if (l.includes("grupos")) return "Grupos Estudiantiles";
  return s.trim();
}
function nMod(s) {
  if (!s) return "Sin modalidad";
  const l = s.trim().toLowerCase();
  if (l.startsWith("v")) return "Virtual";
  if (l.startsWith("p")) return "Presencial";
  return s.trim();
}

/* ═══════════════ SEARCH UTILITIES ═══════════════ */
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

function buildAccentRegex(q) {
  if (!q) return null;
  const map = { a:"[aáàäâã]", e:"[eéèëê]", i:"[iíìïî]", o:"[oóòöôõ]", u:"[uúùüû]", n:"[nñ]" };
  try {
    const escaped = q.toLowerCase().split("").map(c => map[c] || c.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")).join("");
    return new RegExp(escaped, "gi");
  } catch { return null; }
}

/* ═══════════════ EXCEL DATE PARSING ═══════════════ */
function parseDate(v) {
  if (!v) return null;
  if (v instanceof Date) return isNaN(v) ? null : v;
  if (typeof v === "number") {
    const d = new Date(Math.round((v - 25569) * 86400000));
    return isNaN(d) ? null : d;
  }
  const s = String(v).trim();
  const d = new Date(s);
  return isNaN(d) ? null : d;
}
function weekNum(d) {
  if (!d) return 0;
  const jan1 = new Date(d.getFullYear(), 0, 1);
  return Math.ceil((((d - jan1) / 86400000) + jan1.getDay() + 1) / 7);
}
function fmtDate(d) {
  if (!d) return "—";
  return d.toLocaleDateString("es-MX", { day:"2-digit", month:"short", year:"numeric" });
}

/* parseExcel movido a src/parseWorker.js — ver handleFile en CRM() */

/* ═══════════════ UTILITIES ═══════════════ */
function countBy(arr, key) {
  const map = {};
  arr.forEach(r => { const v = r[key] || "N/A"; map[v] = (map[v] || 0) + 1; });
  return Object.entries(map).sort((a, b) => b[1] - a[1]).map(([name, value]) => ({ name, value }));
}

function dlXl(rows, filename) {
  const ws = XLSX.utils.json_to_sheet(rows);
  const maxW = {};
  [Object.keys(rows[0] || {}), ...rows.map(r => Object.values(r).map(String))].forEach(row => {
    row.forEach((c, i) => { maxW[i] = Math.max(maxW[i] || 8, String(c).length + 2); });
  });
  ws["!cols"] = Object.values(maxW).map(w => ({ wch: Math.min(w, 40) }));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Datos");
  XLSX.writeFile(wb, filename);
}

function dlPng(svgEl, filename) {
  if (!svgEl) return;
  const svg = svgEl.cloneNode(true);
  svg.setAttribute("xmlns", "http://www.w3.org/2000/svg");
  const xml = new XMLSerializer().serializeToString(svg);
  const blob = new Blob([xml], { type: "image/svg+xml" });
  const url = URL.createObjectURL(blob);
  const img = new Image();
  img.onload = () => {
    const canvas = document.createElement("canvas");
    canvas.width = img.width * 2;
    canvas.height = img.height * 2;
    const ctx = canvas.getContext("2d");
    ctx.fillStyle = "#0b1120";
    ctx.fillRect(0, 0, canvas.width, canvas.height);
    ctx.scale(2, 2);
    ctx.drawImage(img, 0, 0);
    canvas.toBlob(b => {
      const a = document.createElement("a");
      a.href = URL.createObjectURL(b);
      a.download = filename;
      a.click();
    }, "image/png");
    URL.revokeObjectURL(url);
  };
  img.src = url;
}

/* ═══════════════ STYLES ═══════════════ */
const S = {
  card: { background:"rgba(255,255,255,0.04)", border:"1px solid rgba(255,255,255,0.08)", borderRadius:18, padding:24, marginBottom:16 },
  kpi: (color) => ({ background:"rgba(255,255,255,0.04)", border:"1px solid rgba(255,255,255,0.08)", borderRadius:16, padding:"20px 24px", borderLeft:`4px solid ${color}`, cursor:"default", transition:"all .2s" }),
  btn: (color="#6366f1") => ({ background:`${color}2e`, color, border:"none", borderRadius:10, padding:"8px 16px", fontSize:13, fontWeight:600, cursor:"pointer", fontFamily:"'Plus Jakarta Sans'", transition:"all .2s" }),
  badge: (color="#6366f1") => ({ display:"inline-block", background:`${color}22`, color, borderRadius:20, padding:"3px 10px", fontSize:11, fontWeight:600 }),
  input: { background:"#0a1525", border:"1px solid rgba(255,255,255,0.1)", borderRadius:10, padding:"8px 14px", color:"#e8e9ed", fontSize:13, fontFamily:"'Plus Jakarta Sans'", outline:"none", width:"100%" },
  select: { background:"#0a1525", border:"1px solid rgba(255,255,255,0.1)", borderRadius:10, padding:"8px 14px", color:"#e8e9ed", fontSize:13, fontFamily:"'Plus Jakarta Sans'", outline:"none" },
  mono: { fontFamily:"'JetBrains Mono', monospace" },
  dim: { color:"#8e92a6", fontSize:12 },
  h2: { fontSize:18, fontWeight:700, marginBottom:16 },
  h3: { fontSize:15, fontWeight:600, marginBottom:12, color:"#a5b4fc" },
  grid: (cols) => ({ display:"grid", gridTemplateColumns:`repeat(${cols}, 1fr)`, gap:16 }),
  flex: { display:"flex", alignItems:"center", gap:8 },
  overlay: { position:"fixed", inset:0, background:"rgba(0,0,0,0.6)", backdropFilter:"blur(6px)", display:"flex", alignItems:"center", justifyContent:"center", zIndex:1000 },
  modal: { background:"#0f1628", borderRadius:22, border:"1px solid rgba(255,255,255,0.1)", padding:32, maxWidth:900, width:"90vw", maxHeight:"85vh", overflowY:"auto", position:"relative" },
};

/* ═══════════════ REUSABLE COMPONENTS ═══════════════ */
function Cd({ children, style }) { return <div style={{ ...S.card, ...style }}>{children}</div>; }
function Bt({ children, color, onClick, style }) {
  return <button style={{ ...S.btn(color), ...style }} onClick={onClick}
    onMouseEnter={e => e.target.style.opacity=0.8} onMouseLeave={e => e.target.style.opacity=1}
  >{children}</button>;
}

function KPI({ label, value, sub, color = "#6366f1" }) {
  return (
    <div style={S.kpi(color)} onMouseEnter={e => { e.currentTarget.style.transform="translateY(-3px)"; e.currentTarget.style.boxShadow=`0 8px 24px ${color}22`; }}
      onMouseLeave={e => { e.currentTarget.style.transform="translateY(0)"; e.currentTarget.style.boxShadow="none"; }}>
      <div style={{ color:"#8e92a6", fontSize:11, fontWeight:600, textTransform:"uppercase", letterSpacing:1, marginBottom:6 }}>{label}</div>
      <div style={{ ...S.mono, fontSize:28, fontWeight:700, color }}>{value}</div>
      {sub && <div style={{ color:"#6b6f82", fontSize:11, marginTop:4 }}>{sub}</div>}
    </div>
  );
}

function SB({ items, total }) {
  return (
    <div>
      {items.map(({ name, value, color }, i) => {
        const pct = total ? ((value / total) * 100).toFixed(1) : 0;
        return (
          <div key={i} style={{ marginBottom:10 }}>
            <div style={{ display:"flex", justifyContent:"space-between", marginBottom:4 }}>
              <span style={{ fontSize:12, color:"#e8e9ed" }}>{name}</span>
              <span style={{ ...S.mono, fontSize:12, color:"#8e92a6" }}>{value} ({pct}%)</span>
            </div>
            <div style={{ height:5, background:"rgba(255,255,255,0.06)", borderRadius:3, overflow:"hidden" }}>
              <div style={{ height:"100%", width:`${pct}%`, borderRadius:3, background:`linear-gradient(90deg, ${color || CHART_COLORS[i % CHART_COLORS.length]}, ${color || CHART_COLORS[i % CHART_COLORS.length]}aa)`, transition:"width .5s" }} />
            </div>
          </div>
        );
      })}
    </div>
  );
}

function ChartCard({ title, children, chartRef, filename }) {
  return (
    <Cd>
      <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:16 }}>
        <div style={S.h3}>{title}</div>
        {chartRef && <Bt color="#8e92a6" onClick={() => { const svg = chartRef.current?.querySelector("svg"); if(svg) dlPng(svg, filename || "chart.png"); }} style={{ padding:"4px 10px", fontSize:11 }}>📷 PNG</Bt>}
      </div>
      {children}
    </Cd>
  );
}

function Highlight({ text, query }) {
  if (!query || !text) return <span>{String(text ?? "")}</span>;
  const t = String(text);
  const regex = buildAccentRegex(query);
  if (!regex) return <span>{t}</span>;
  const parts = [];
  let lastIdx = 0, match;
  while ((match = regex.exec(t)) !== null) {
    if (match[0].length === 0) { lastIdx++; continue; }
    if (match.index > lastIdx) parts.push({ t: t.slice(lastIdx, match.index), m: false });
    parts.push({ t: match[0], m: true });
    lastIdx = regex.lastIndex;
  }
  if (lastIdx < t.length) parts.push({ t: t.slice(lastIdx), m: false });
  if (!parts.length) return <span>{t}</span>;
  return (
    <span>{parts.map((p, i) => p.m
      ? <mark key={i} style={{ background:"rgba(99,102,241,0.35)", color:"#e8e9ed", borderRadius:3, padding:"0 2px" }}>{p.t}</mark>
      : p.t
    )}</span>
  );
}

const TT = ({ active, payload, label }) => {
  if (!active || !payload?.length) return null;
  return (
    <div style={{ background:"#0f1a2e", border:"1px solid rgba(255,255,255,0.1)", borderRadius:10, padding:"10px 14px", fontSize:12 }}>
      <div style={{ fontWeight:600, marginBottom:4 }}>{label}</div>
      {payload.map((p, i) => (
        <div key={i} style={{ color: p.color || p.fill, ...S.mono, fontSize:11 }}>
          {p.name || p.dataKey}: {p.value}
        </div>
      ))}
    </div>
  );
};

/* ═══════════════ MODAL COMPONENTS ═══════════════ */
function Modal({ onClose, children }) {
  return (
    <div style={S.overlay} onClick={onClose}>
      <div style={S.modal} onClick={e => e.stopPropagation()}>
        <button onClick={onClose} style={{ position:"absolute", top:16, right:16, background:"none", border:"none", color:"#8e92a6", fontSize:20, cursor:"pointer" }}>✕</button>
        {children}
      </div>
    </div>
  );
}

function StudentModal({ student, records, onClose }) {
  const sorted = [...records].sort((a, b) => (b.fecha || 0) - (a.fecha || 0));
  const asist = records.filter(r => r.estatus === "Asistencia").length;
  const services = [...new Set(records.map(r => r.servicio))];
  return (
    <Modal onClose={onClose}>
      <div style={{ marginBottom:20 }}>
        <div style={{ fontSize:20, fontWeight:700 }}>{student.nombre}</div>
        <div style={{ ...S.mono, color:"#8e92a6", fontSize:13, marginTop:4 }}>{student.matricula}</div>
        <div style={{ display:"flex", gap:8, marginTop:10, flexWrap:"wrap" }}>
          <span style={S.badge("#6366f1")}>{student.escuela}</span>
          <span style={S.badge("#10b981")}>{student.programa}</span>
          {student.comunidad && student.comunidad !== "Sin comunidad" && <span style={S.badge("#f59e0b")}>{student.comunidad}</span>}
          {student.interes && student.interes !== "Sin interés" && <span style={S.badge("#ec4899")}>{student.interes}</span>}
          {student.isCAGS  && <span style={{ ...S.badge("#a855f7"), fontWeight:700 }}>★ CAGS JUN26</span>}
          {student.isDIC25 && <span style={{ ...S.badge("#22d3ee"), fontWeight:700 }}>✓ DIC25</span>}
        </div>
      </div>
      <div style={S.grid(4)}>
        <KPI label="Sesiones" value={records.length} color="#6366f1" />
        <KPI label="Asistencias" value={asist} color="#10b981" />
        <KPI label="Servicios" value={services.length} color="#f59e0b" />
        <KPI label="Tasa asist." value={records.length ? `${((asist / records.length) * 100).toFixed(0)}%` : "—"} color="#8b5cf6" />
      </div>
      <div style={{ marginTop:20 }}>
        <div style={S.h3}>Servicios utilizados</div>
        <div style={{ display:"flex", gap:6, flexWrap:"wrap" }}>
          {services.map((s, i) => <span key={i} style={S.badge(CHART_COLORS[i % CHART_COLORS.length])}>{s}</span>)}
        </div>
      </div>
      <div style={{ marginTop:20 }}>
        <div style={S.h3}>Timeline</div>
        {sorted.map((r, i) => (
          <div key={i} style={{ display:"flex", gap:12, padding:"10px 0", borderBottom:"1px solid rgba(255,255,255,0.05)" }}>
            <div style={{ ...S.mono, fontSize:11, color:"#6b6f82", minWidth:80 }}>{fmtDate(r.fecha)}</div>
            <div style={{ fontSize:12 }}>{r.servicio}</div>
            <span style={{ ...S.badge(STATUS_COLORS[r.estatus] || "#6366f1"), marginLeft:"auto" }}>{r.estatus}</span>
          </div>
        ))}
      </div>
      <div style={{ marginTop:16 }}>
        <Bt color="#6366f1" onClick={() => dlXl(records.map(r => ({ Fecha:fmtDate(r.fecha), Servicio:r.servicio, Asesor:r.asesor, Estatus:r.estatus, Modalidad:r.modalidad })), `historial_${student.matricula}.xlsx`)}>
          ↓ Descargar historial
        </Bt>
      </div>
    </Modal>
  );
}

function AsesorModal({ asesor, records, onClose }) {
  const asist = records.filter(r => r.estatus === "Asistencia").length;
  const faltas = records.filter(r => r.estatus === "Falta").length;
  const express = records.filter(r => r.estatus === "Express").length;
  const weeks = new Set(records.map(r => r.semana));
  const srvData = countBy(records, "servicio").slice(0, 8);
  const escData = countBy(records, "escuela").slice(0, 6);
  const last30 = [...records].sort((a, b) => (b.fecha || 0) - (a.fecha || 0)).slice(0, 30);
  const chartRef1 = useRef(), chartRef2 = useRef();
  const base = records.length - express;

  const [aSearch, setASearch] = useState("");
  const [aEscuela, setAEscuela] = useState("");
  const [aEstatus, setAEstatus] = useState("");
  const [aInteres, setAInteres] = useState("");
  const [aPrograma, setAPrograma] = useState("");
  const [aServicio, setAServicio] = useState("");
  const [aComunidad, setAComunidad] = useState("");
  const [aCAGS, setACAGS] = useState(false);
  const [aDIC25, setADIC25] = useState(false);
  const dASearch = useDebounce(aSearch, 200);

  const alumnosList = useMemo(() => {
    const m = {};
    records.forEach(r => {
      if (!m[r.matricula]) m[r.matricula] = {
        matricula:r.matricula, nombre:r.nombre, sesiones:0, asistencias:0,
        servicios:new Set(), ultimaFecha:null,
        escuela:r.escuela, programa:r.programa, interes:r.interes, comunidad:r.comunidad,
        isCAGS:false, isDIC25:false, allEstatus:new Set(),
      };
      const s = m[r.matricula];
      s.sesiones++;
      if (r.estatus === "Asistencia") s.asistencias++;
      s.servicios.add(r.servicio);
      s.allEstatus.add(r.estatus);
      if (r.isCAGS) s.isCAGS = true;
      if (r.isDIC25) s.isDIC25 = true;
      if (!s.ultimaFecha || (r.fecha && r.fecha > s.ultimaFecha)) {
        s.ultimaFecha = r.fecha;
        s.escuela = r.escuela; s.programa = r.programa;
        s.interes = r.interes; s.comunidad = r.comunidad;
      }
    });
    return Object.values(m).sort((a, b) => b.sesiones - a.sesiones);
  }, [records]);

  const aEscuelas   = useMemo(() => [...new Set(alumnosList.map(s => s.escuela))].sort(),   [alumnosList]);
  const aEstatuses  = useMemo(() => [...new Set(records.map(r => r.estatus))].sort(),        [records]);
  const aIntereses  = useMemo(() => [...new Set(alumnosList.map(s => s.interes))].sort(),   [alumnosList]);
  const aProgramas  = useMemo(() => [...new Set(alumnosList.map(s => s.programa))].sort(),  [alumnosList]);
  const aServicios  = useMemo(() => [...new Set(records.map(r => r.servicio))].sort(),       [records]);
  const aComunidades = useMemo(() => [...new Set(alumnosList.map(s => s.comunidad))].sort(), [alumnosList]);

  const alumnosFiltrados = useMemo(() => {
    let f = alumnosList;
    if (dASearch) { const q = norm(dASearch); f = f.filter(s => norm(s.nombre).includes(q) || norm(s.matricula).includes(q)); }
    if (aEscuela)   f = f.filter(s => s.escuela   === aEscuela);
    if (aEstatus)   f = f.filter(s => s.allEstatus.has(aEstatus));
    if (aInteres)   f = f.filter(s => s.interes   === aInteres);
    if (aPrograma)  f = f.filter(s => s.programa  === aPrograma);
    if (aServicio)  f = f.filter(s => s.servicios.has(aServicio));
    if (aComunidad) f = f.filter(s => s.comunidad === aComunidad);
    if (aCAGS)  f = f.filter(s => s.isCAGS);
    if (aDIC25) f = f.filter(s => s.isDIC25);
    return f;
  }, [alumnosList, dASearch, aEscuela, aEstatus, aInteres, aPrograma, aServicio, aComunidad, aCAGS, aDIC25]);

  const clearAFiltros = () => { setASearch(""); setAEscuela(""); setAEstatus(""); setAInteres(""); setAPrograma(""); setAServicio(""); setAComunidad(""); setACAGS(false); setADIC25(false); };

  return (
    <Modal onClose={onClose}>
      <div style={{ fontSize:20, fontWeight:700, marginBottom:20 }}>{asesor}</div>
      <div style={S.grid(5)}>
        <KPI label="Total" value={records.length} color="#6366f1" />
        <KPI label="Asistencias" value={asist} color="#10b981" sub={base ? `${((asist / base) * 100).toFixed(1)}% de tasa` : "—"} />
        <KPI label="Faltas" value={faltas} color="#ef4444" sub={base ? `${((faltas / base) * 100).toFixed(1)}% de tasa` : "—"} />
        <KPI label="Prom/semana" value={weeks.size ? (records.length / weeks.size).toFixed(1) : "—"} color="#f59e0b" />
        <KPI label="Alumnos únicos" value={alumnosList.length} color="#8b5cf6" />
      </div>
      <div style={{ ...S.grid(2), marginTop:20 }}>
        <ChartCard title="Servicios" chartRef={chartRef1} filename={`${asesor}_servicios.png`}>
          <div ref={chartRef1}><ResponsiveContainer width="100%" height={200}>
            <BarChart data={srvData} layout="vertical" margin={{ left:90, right:20, top:5, bottom:5 }}>
              <XAxis type="number" tick={{ fill:"#6b6f82", fontSize:10 }} />
              <YAxis type="category" dataKey="name" tick={{ fill:"#8e92a6", fontSize:10 }} width={85} />
              <Tooltip content={<TT />} />
              <Bar dataKey="value" radius={[0,6,6,0]}>{srvData.map((_, i) => <Cell key={i} fill={CHART_COLORS[i % CHART_COLORS.length]} />)}</Bar>
            </BarChart>
          </ResponsiveContainer></div>
        </ChartCard>
        <ChartCard title="Escuelas" chartRef={chartRef2} filename={`${asesor}_escuelas.png`}>
          <div ref={chartRef2}><ResponsiveContainer width="100%" height={200}>
            <PieChart>
              <Pie data={escData} dataKey="value" nameKey="name" cx="50%" cy="50%" outerRadius={70} innerRadius={30} paddingAngle={3}>
                {escData.map((_, i) => <Cell key={i} fill={CHART_COLORS[i % CHART_COLORS.length]} />)}
              </Pie>
              <Tooltip content={<TT />} />
            </PieChart>
          </ResponsiveContainer></div>
        </ChartCard>
      </div>
      <div style={{ marginTop:20 }}>
        <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:8 }}>
          <div style={S.h3}>Alumnos atendidos ({alumnosFiltrados.length}{alumnosFiltrados.length !== alumnosList.length ? ` de ${alumnosList.length}` : ""})</div>
        </div>
        <div style={{ display:"flex", gap:6, flexWrap:"wrap", marginBottom:6 }}>
          <input style={{ ...S.input, maxWidth:200, padding:"6px 10px" }} placeholder="Buscar alumno..." value={aSearch} onChange={e => setASearch(e.target.value)} />
          <select style={{ ...S.select, fontSize:11 }} value={aEscuela} onChange={e => setAEscuela(e.target.value)}>
            <option value="">Escuela</option>
            {aEscuelas.map(v => <option key={v} value={v}>{v}</option>)}
          </select>
          <select style={{ ...S.select, fontSize:11 }} value={aEstatus} onChange={e => setAEstatus(e.target.value)}>
            <option value="">Estatus</option>
            {aEstatuses.map(v => <option key={v} value={v}>{v}</option>)}
          </select>
          <select style={{ ...S.select, fontSize:11 }} value={aInteres} onChange={e => setAInteres(e.target.value)}>
            <option value="">Interés</option>
            {aIntereses.map(v => <option key={v} value={v}>{v}</option>)}
          </select>
          <select style={{ ...S.select, fontSize:11 }} value={aPrograma} onChange={e => setAPrograma(e.target.value)}>
            <option value="">Programa</option>
            {aProgramas.map(v => <option key={v} value={v}>{v}</option>)}
          </select>
          <select style={{ ...S.select, fontSize:11 }} value={aServicio} onChange={e => setAServicio(e.target.value)}>
            <option value="">Servicio</option>
            {aServicios.map(v => <option key={v} value={v}>{v}</option>)}
          </select>
          <select style={{ ...S.select, fontSize:11 }} value={aComunidad} onChange={e => setAComunidad(e.target.value)}>
            <option value="">Comunidad</option>
            {aComunidades.map(v => <option key={v} value={v}>{v}</option>)}
          </select>
          <button onClick={() => { setACAGS(v => !v); setADIC25(false); }}
            style={{ ...S.btn(aCAGS ? "#a855f7" : "#8e92a6"), fontSize:11, padding:"6px 10px", opacity: aDIC25 ? 0.4 : 1 }}>
            ★ CAGS
          </button>
          <button onClick={() => { setADIC25(v => !v); setACAGS(false); }}
            style={{ ...S.btn(aDIC25 ? "#22d3ee" : "#8e92a6"), fontSize:11, padding:"6px 10px", opacity: aCAGS ? 0.4 : 1 }}>
            ✓ DIC25
          </button>
          <Bt color="#8e92a6" onClick={clearAFiltros} style={{ fontSize:11, padding:"6px 10px" }}>Limpiar</Bt>
        </div>
        <div style={{ overflowX:"auto", maxHeight:280, overflowY:"auto" }}>
          <table style={{ width:"100%", borderCollapse:"collapse", fontSize:12 }}>
            <thead style={{ position:"sticky", top:0, background:"#0f1628", zIndex:1 }}>
              <tr style={{ borderBottom:"1px solid rgba(255,255,255,0.08)" }}>
                {["Alumno","Matrícula","Sesiones","Asistencias","Servicios","Última visita"].map(h => (
                  <th key={h} style={{ textAlign:"left", padding:"8px 10px", fontSize:10, textTransform:"uppercase", letterSpacing:1, color:"#6b6f82", fontWeight:600, whiteSpace:"nowrap" }}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>{alumnosFiltrados.map((a, i) => (
              <tr key={i} style={{ borderBottom:"1px solid rgba(255,255,255,0.04)" }}
                onMouseEnter={e => e.currentTarget.style.background="rgba(99,102,241,0.06)"}
                onMouseLeave={e => e.currentTarget.style.background="transparent"}>
                <td style={{ padding:"8px 10px", fontWeight:500 }}>
                  <Highlight text={a.nombre} query={dASearch} />
                  {a.isCAGS  && <span style={{ ...S.badge("#a855f7"), marginLeft:6, fontSize:9 }}>CAGS</span>}
                  {a.isDIC25 && <span style={{ ...S.badge("#22d3ee"), marginLeft:6, fontSize:9 }}>DIC25</span>}
                </td>
                <td style={{ padding:"8px 10px", ...S.mono, fontSize:11, color:"#a5b4fc" }}><Highlight text={a.matricula} query={dASearch} /></td>
                <td style={{ padding:"8px 10px", textAlign:"center" }}><span style={S.badge("#6366f1")}>{a.sesiones}</span></td>
                <td style={{ padding:"8px 10px", textAlign:"center" }}><span style={S.badge("#10b981")}>{a.asistencias}</span></td>
                <td style={{ padding:"8px 10px", fontSize:11, color:"#8e92a6", maxWidth:160, overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }}>{[...a.servicios].join(", ")}</td>
                <td style={{ padding:"8px 10px", ...S.mono, fontSize:11, color:"#6b6f82" }}>{fmtDate(a.ultimaFecha)}</td>
              </tr>
            ))}</tbody>
          </table>
          {alumnosFiltrados.length === 0 && (
            <div style={{ textAlign:"center", padding:24, color:"#6b6f82", fontSize:12 }}>Sin resultados con los filtros aplicados</div>
          )}
        </div>
      </div>

      <div style={{ marginTop:20 }}>
        <div style={S.h3}>Últimas 30 asesorías</div>
        <div style={{ overflowX:"auto" }}>
          <table style={{ width:"100%", borderCollapse:"collapse", fontSize:12 }}>
            <thead><tr style={{ borderBottom:"1px solid rgba(255,255,255,0.08)" }}>
              {["Fecha","Alumno","Servicio","Estatus","Modalidad"].map(h => (
                <th key={h} style={{ textAlign:"left", padding:"8px 10px", fontSize:10, textTransform:"uppercase", letterSpacing:1, color:"#6b6f82", fontWeight:600 }}>{h}</th>
              ))}
            </tr></thead>
            <tbody>{last30.map((r, i) => (
              <tr key={i} style={{ borderBottom:"1px solid rgba(255,255,255,0.04)" }}
                onMouseEnter={e => e.currentTarget.style.background="rgba(99,102,241,0.06)"}
                onMouseLeave={e => e.currentTarget.style.background="transparent"}>
                <td style={{ padding:"8px 10px", ...S.mono, color:"#6b6f82", fontSize:11 }}>{fmtDate(r.fecha)}</td>
                <td style={{ padding:"8px 10px" }}>{r.nombre}</td>
                <td style={{ padding:"8px 10px" }}>{r.servicio}</td>
                <td style={{ padding:"8px 10px" }}><span style={S.badge(STATUS_COLORS[r.estatus]||"#6366f1")}>{r.estatus}</span></td>
                <td style={{ padding:"8px 10px" }}>{r.modalidad}</td>
              </tr>
            ))}</tbody>
          </table>
        </div>
      </div>
      <div style={{ marginTop:16, display:"flex", gap:10 }}>
        <Bt color="#6366f1" onClick={() => dlXl(records.map(r => ({ Fecha:fmtDate(r.fecha), Matrícula:r.matricula, Alumno:r.nombre, Servicio:r.servicio, Estatus:r.estatus, Escuela:r.escuela, Programa:r.programa, Modalidad:r.modalidad })), `asesor_${asesor}.xlsx`)}>↓ Descargar datos</Bt>
        <Bt color="#10b981" onClick={() => dlXl(alumnosFiltrados.map(a => ({ Matrícula:a.matricula, Alumno:a.nombre, Sesiones:a.sesiones, Asistencias:a.asistencias, Servicios:[...a.servicios].join(", "), CAGS:a.isCAGS?"Sí":"No", DIC25:a.isDIC25?"Sí":"No", "Última visita":fmtDate(a.ultimaFecha) })), `alumnos_${asesor}.xlsx`)}>↓ Lista alumnos</Bt>
      </div>
    </Modal>
  );
}

/* ═══════════════ TAB: DASHBOARD ═══════════════ */
function TabDashboard({ data }) {
  const total = data.length;
  const asist = data.filter(r => r.estatus === "Asistencia").length;
  const faltas = data.filter(r => r.estatus === "Falta").length;
  const express = data.filter(r => r.estatus === "Express").length;
  const cancel = data.filter(r => r.estatus === "Cancelación").length;
  const uniq = new Set(data.map(r => r.matricula)).size;
  const base = total - express;
  const tasaAsist = base ? ((asist / base) * 100).toFixed(1) : 0;
  const tasaFalta = base ? ((faltas / base) * 100).toFixed(1) : 0;

  const cagsUniq = useMemo(() => {
    const seen = new Set();
    data.forEach(r => { if (r.isCAGS) seen.add(r.matricula); });
    return seen.size;
  }, [data]);
  const dic25Uniq = useMemo(() => {
    const seen = new Set();
    data.forEach(r => { if (r.isDIC25) seen.add(r.matricula); });
    return seen.size;
  }, [data]);

  const weekData = useMemo(() => {
    const wm = {};
    data.forEach(r => { if (r.semana) { wm[r.semana] = (wm[r.semana] || 0) + 1; } });
    return Object.entries(wm).sort((a, b) => a[0] - b[0]).map(([w, v]) => ({ name:`S${w}`, value:v }));
  }, [data]);

  const statusData = [
    { name:"Asistencia", value:asist, color:"#10b981" },
    { name:"Falta", value:faltas, color:"#ef4444" },
    { name:"Express", value:express, color:"#f59e0b" },
    { name:"Cancelación", value:cancel, color:"#8b5cf6" }
  ];
  const escData = countBy(data, "escuela");
  const intData = countBy(data, "interes");
  const srvData = countBy(data, "servicio");
  const comData = countBy(data, "comunidad");
  const modData = countBy(data, "modalidad");

  const ref1=useRef(),ref2=useRef(),ref3=useRef(),ref4=useRef(),ref5=useRef();

  const dlAsist = () => {
    const rows = data.filter(r => r.estatus === "Asistencia").map(r => ({ Fecha:fmtDate(r.fecha), Matrícula:r.matricula, Nombre:r.nombre, Servicio:r.servicio, Asesor:r.asesor, Escuela:r.escuela, Programa:r.programa }));
    dlXl(rows, "asistencias_FJ26.xlsx");
  };
  const dlReporte = () => {
    const rows = data.map(r => ({ Fecha:fmtDate(r.fecha), Matrícula:r.matricula, Nombre:r.nombre, Servicio:r.servicio, Asesor:r.asesor, Escuela:r.escuela, Programa:r.programa, Estatus:r.estatus, Interés:r.interes, Modalidad:r.modalidad, Comunidad:r.comunidad, Semestre:r.semestre }));
    dlXl(rows, "reporte_completo_FJ26.xlsx");
  };

  return (
    <div>
      <div style={{ display:"flex", gap:10, marginBottom:20 }}>
        <Bt color="#10b981" onClick={dlAsist}>↓ Asistencias</Bt>
        <Bt color="#6366f1" onClick={dlReporte}>↓ Reporte completo</Bt>
      </div>
      <div style={S.grid(5)}>
        <KPI label="Total asesorías agendadas" value={total.toLocaleString()} color="#6366f1" sub={`${(total / Math.max(1, uniq)).toFixed(1)} por alumno`} />
        <KPI label="Alumnos únicos" value={uniq.toLocaleString()} color="#3b82f6" />
        <KPI label="Tasa asistencia" value={`${tasaAsist}%`} color="#10b981" sub={`${asist} asistencias`} />
        <KPI label="Faltas" value={faltas} color="#ef4444" sub={`${tasaFalta}% tasa`} />
        <KPI label="Express" value={express} color="#f59e0b" />
      </div>
      <div style={{ ...S.grid(2), marginBottom:4 }}>
        <KPI label="CAGS — Candidatos JUN 2026" value={cagsUniq} color="#a855f7" sub={`${uniq ? ((cagsUniq / uniq) * 100).toFixed(1) : 0}% de alumnos únicos`} />
        <KPI label="DIC 2025 — Generación graduada" value={dic25Uniq} color="#22d3ee" sub={`${uniq ? ((dic25Uniq / uniq) * 100).toFixed(1) : 0}% de alumnos únicos`} />
      </div>

      <ChartCard title="Asesorías por semana" chartRef={ref1} filename="asesorias_semana.png">
        <div ref={ref1}><ResponsiveContainer width="100%" height={260}>
          <BarChart data={weekData} margin={{ top:5, right:20, bottom:5, left:10 }}>
            <XAxis dataKey="name" tick={{ fill:"#6b6f82", fontSize:10 }} />
            <YAxis tick={{ fill:"#6b6f82", fontSize:10 }} />
            <Tooltip content={<TT />} />
            <Bar dataKey="value" fill="#6366f1" radius={[6,6,0,0]} />
          </BarChart>
        </ResponsiveContainer></div>
      </ChartCard>

      <div style={S.grid(3)}>
        <ChartCard title="Estatus">
          <SB items={statusData} total={total} />
        </ChartCard>
        <ChartCard title="Escuelas" chartRef={ref2} filename="escuelas.png">
          <div ref={ref2}><ResponsiveContainer width="100%" height={220}>
            <PieChart>
              <Pie data={escData} dataKey="value" nameKey="name" cx="50%" cy="50%" outerRadius={80} innerRadius={35} paddingAngle={3}>
                {escData.map((_, i) => <Cell key={i} fill={CHART_COLORS[i % CHART_COLORS.length]} />)}
              </Pie>
              <Tooltip content={<TT />} />
            </PieChart>
          </ResponsiveContainer></div>
        </ChartCard>
        <ChartCard title="Interés de asesoría">
          <SB items={intData.map((d, i) => ({ ...d, color: CHART_COLORS[i % CHART_COLORS.length] }))} total={total} />
        </ChartCard>
      </div>

      <div style={S.grid(3)}>
        <ChartCard title="Servicios" chartRef={ref3} filename="servicios.png">
          <div ref={ref3}><ResponsiveContainer width="100%" height={Math.max(200, srvData.length * 28)}>
            <BarChart data={srvData} layout="vertical" margin={{ left:120, right:20, top:5, bottom:5 }}>
              <XAxis type="number" tick={{ fill:"#6b6f82", fontSize:10 }} />
              <YAxis type="category" dataKey="name" tick={{ fill:"#8e92a6", fontSize:10 }} width={115} />
              <Tooltip content={<TT />} />
              <Bar dataKey="value" radius={[0,6,6,0]}>{srvData.map((_, i) => <Cell key={i} fill={CHART_COLORS[i % CHART_COLORS.length]} />)}</Bar>
            </BarChart>
          </ResponsiveContainer></div>
        </ChartCard>
        <ChartCard title="Comunidades" chartRef={ref4} filename="comunidades.png">
          <div ref={ref4}><ResponsiveContainer width="100%" height={260}>
            <BarChart data={comData.slice(0, 10)} margin={{ top:5, right:10, bottom:5, left:10 }}>
              <XAxis dataKey="name" tick={{ fill:"#6b6f82", fontSize:9 }} angle={-30} textAnchor="end" height={60} />
              <YAxis tick={{ fill:"#6b6f82", fontSize:10 }} />
              <Tooltip content={<TT />} />
              <Bar dataKey="value" radius={[6,6,0,0]}>{comData.slice(0,10).map((_, i) => <Cell key={i} fill={CHART_COLORS[i % CHART_COLORS.length]} />)}</Bar>
            </BarChart>
          </ResponsiveContainer></div>
        </ChartCard>
        <ChartCard title="Modalidad" chartRef={ref5} filename="modalidad.png">
          <div ref={ref5}><ResponsiveContainer width="100%" height={220}>
            <PieChart>
              <Pie data={modData} dataKey="value" nameKey="name" cx="50%" cy="50%" outerRadius={80} innerRadius={35} paddingAngle={5}>
                {modData.map((_, i) => <Cell key={i} fill={CHART_COLORS[i % CHART_COLORS.length]} />)}
              </Pie>
              <Tooltip content={<TT />} />
              <Legend formatter={(v) => <span style={{ color:"#8e92a6", fontSize:11 }}>{v}</span>} />
            </PieChart>
          </ResponsiveContainer></div>
        </ChartCard>
      </div>
    </div>
  );
}

/* ═══════════════ TAB: ASESORES ═══════════════ */
function TabAsesores({ data }) {
  const [selected, setSelected] = useState(null);
  const byAsesor = useMemo(() => {
    const m = {};
    data.forEach(r => { (m[r.asesor] = m[r.asesor] || []).push(r); });
    return Object.entries(m).sort((a, b) => b[1].length - a[1].length);
  }, [data]);
  const distData = byAsesor.map(([name, recs]) => ({ name, value: recs.length }));
  const distRef = useRef();

  return (
    <div>
      {selected && <AsesorModal asesor={selected} records={data.filter(r => r.asesor === selected)} onClose={() => setSelected(null)} />}
      <div style={{ ...S.grid(2), marginBottom:20 }}>
        {byAsesor.map(([name, recs], idx) => {
          const a = recs.filter(r => r.estatus === "Asistencia").length;
          const ex = recs.filter(r => r.estatus === "Express").length;
          const f = recs.filter(r => r.estatus === "Falta").length;
          const b = recs.length - ex;
          const weeks = new Set(recs.map(r => r.semana));
          const color = CHART_COLORS[idx % CHART_COLORS.length];
          const cagsCount = new Set(recs.filter(r => r.isCAGS).map(r => r.matricula)).size;
          const dic25Count = new Set(recs.filter(r => r.isDIC25).map(r => r.matricula)).size;
          return (
            <Cd key={name} style={{ cursor:"pointer", borderLeft:`4px solid ${color}` }} >
              <div onClick={() => setSelected(name)}>
                <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:10 }}>
                  <div style={{ fontSize:16, fontWeight:700 }}>{name}</div>
                  <span style={{ ...S.mono, fontSize:22, fontWeight:700, color }}>{recs.length}</span>
                </div>
                <div style={{ display:"flex", gap:16, marginBottom:8, fontSize:12, color:"#8e92a6" }}>
                  <span>Asistencia: <b style={{ color:"#10b981" }}>{b ? ((a / b) * 100).toFixed(0) : 0}%</b></span>
                  <span>Prom/sem: <b style={{ color:"#f59e0b" }}>{weeks.size ? (recs.length / weeks.size).toFixed(1) : "—"}</b></span>
                </div>
                <div style={{ display:"flex", gap:6, marginBottom:8 }}>
                  {cagsCount > 0 && <span style={S.badge("#a855f7")}>CAGS: {cagsCount}</span>}
                  {dic25Count > 0 && <span style={S.badge("#22d3ee")}>DIC25: {dic25Count}</span>}
                </div>
                <div style={{ display:"flex", height:6, borderRadius:3, overflow:"hidden", gap:2 }}>
                  {a > 0 && <div style={{ flex:a, background:"#10b981", borderRadius:3 }} />}
                  {f > 0 && <div style={{ flex:f, background:"#ef4444", borderRadius:3 }} />}
                  {ex > 0 && <div style={{ flex:ex, background:"#f59e0b", borderRadius:3 }} />}
                </div>
              </div>
              <div style={{ display:"flex", gap:8, marginTop:12 }}>
                <Bt color="#6366f1" onClick={(e) => { e.stopPropagation(); dlXl(recs.map(r => ({ Fecha:fmtDate(r.fecha), Matrícula:r.matricula, Alumno:r.nombre, Servicio:r.servicio, Estatus:r.estatus, Escuela:r.escuela, Programa:r.programa })), `${name}.xlsx`); }} style={{ fontSize:11, padding:"4px 10px" }}>↓ Excel</Bt>
              </div>
            </Cd>
          );
        })}
      </div>
      <ChartCard title="Distribución por asesor" chartRef={distRef} filename="distribucion_asesores.png">
        <div ref={distRef}><ResponsiveContainer width="100%" height={280}>
          <BarChart data={distData} margin={{ top:5, right:20, bottom:5, left:10 }}>
            <XAxis dataKey="name" tick={{ fill:"#6b6f82", fontSize:10 }} />
            <YAxis tick={{ fill:"#6b6f82", fontSize:10 }} />
            <Tooltip content={<TT />} />
            <Bar dataKey="value" radius={[6,6,0,0]}>{distData.map((_, i) => <Cell key={i} fill={CHART_COLORS[i % CHART_COLORS.length]} />)}</Bar>
          </BarChart>
        </ResponsiveContainer></div>
      </ChartCard>
    </div>
  );
}

/* ═══════════════ TAB: PIPELINE ═══════════════ */
function TabPipeline({ data }) {
  const srvData = useMemo(() => countBy(data, "servicio"), [data]);
  const top5 = srvData.slice(0, 5);
  const maxVal = top5[0]?.value || 1;
  const progData = countBy(data, "programa").slice(0, 15);
  const monthData = useMemo(() => {
    const m = {};
    data.forEach(r => { if (r.fecha) { const k = r.fecha.toLocaleDateString("es-MX", { month:"short", year:"numeric" }); m[k] = (m[k] || 0) + 1; } });
    return Object.entries(m).map(([name, value]) => ({ name, value }));
  }, [data]);
  const lineRef = useRef();
  const progRef = useRef();

  return (
    <div>
      <div style={S.grid(5)}>
        {top5.map((s, i) => <KPI key={s.name} label={s.name} value={s.value} color={CHART_COLORS[i]} sub={`${((s.value / data.length) * 100).toFixed(1)}% del total`} />)}
      </div>

      <Cd style={{ marginTop:16 }}>
        <div style={S.h3}>Funnel de servicios</div>
        <div style={{ display:"flex", gap:8, alignItems:"flex-end", height:200, padding:"20px 10px" }}>
          {srvData.map((s, i) => (
            <div key={s.name} style={{ flex:1, display:"flex", flexDirection:"column", alignItems:"center", gap:6 }}>
              <span style={{ ...S.mono, fontSize:10, color:"#8e92a6" }}>{s.value}</span>
              <div style={{ width:"100%", maxWidth:60, height:`${(s.value / maxVal) * 160}px`, background:`linear-gradient(180deg, ${CHART_COLORS[i % CHART_COLORS.length]}, ${CHART_COLORS[i % CHART_COLORS.length]}66)`, borderRadius:"8px 8px 4px 4px", minHeight:4, transition:"height .5s" }} />
              <span style={{ fontSize:8, color:"#6b6f82", textAlign:"center", lineHeight:1.2, maxWidth:60, overflow:"hidden" }}>{s.name}</span>
            </div>
          ))}
        </div>
      </Cd>

      <Cd>
        <div style={S.h3}>Todos los servicios</div>
        <SB items={srvData.map((d, i) => ({ ...d, color: CHART_COLORS[i % CHART_COLORS.length] }))} total={data.length} />
      </Cd>

      <div style={S.grid(2)}>
        <ChartCard title="Top programas académicos" chartRef={progRef} filename="programas.png">
          <div ref={progRef}><ResponsiveContainer width="100%" height={Math.max(200, progData.length * 24)}>
            <BarChart data={progData} layout="vertical" margin={{ left:50, right:20, top:5, bottom:5 }}>
              <XAxis type="number" tick={{ fill:"#6b6f82", fontSize:10 }} />
              <YAxis type="category" dataKey="name" tick={{ fill:"#8e92a6", fontSize:10 }} width={45} />
              <Tooltip content={<TT />} />
              <Bar dataKey="value" radius={[0,6,6,0]}>{progData.map((_, i) => <Cell key={i} fill={CHART_COLORS[i % CHART_COLORS.length]} />)}</Bar>
            </BarChart>
          </ResponsiveContainer></div>
        </ChartCard>
        <ChartCard title="Tendencia mensual" chartRef={lineRef} filename="tendencia_mensual.png">
          <div ref={lineRef}><ResponsiveContainer width="100%" height={260}>
            <LineChart data={monthData} margin={{ top:5, right:20, bottom:5, left:10 }}>
              <CartesianGrid strokeDasharray="3 3" stroke="rgba(255,255,255,0.06)" />
              <XAxis dataKey="name" tick={{ fill:"#6b6f82", fontSize:10 }} />
              <YAxis tick={{ fill:"#6b6f82", fontSize:10 }} />
              <Tooltip content={<TT />} />
              <Line type="monotone" dataKey="value" stroke="#6366f1" strokeWidth={2.5} dot={{ fill:"#6366f1", r:4 }} />
            </LineChart>
          </ResponsiveContainer></div>
        </ChartCard>
      </div>
    </div>
  );
}

/* ═══════════════ TAB: ALUMNOS ═══════════════ */
const PAGE_SIZE = 100;

function TabAlumnos({ data }) {
  const [search, setSearch] = useState("");
  const [fAsesor, setFAsesor] = useState("");
  const [fEscuela, setFEscuela] = useState("");
  const [fEstatus, setFEstatus] = useState("");
  const [fInteres, setFInteres] = useState("");
  const [fPrograma, setFPrograma] = useState("");
  const [fServicio, setFServicio] = useState("");
  const [fComunidad, setFComunidad] = useState("");
  const [fCAGS, setFCAGS] = useState(false);
  const [fDIC25, setFDIC25] = useState(false);
  const [selectedStudent, setSelectedStudent] = useState(null);
  const [sortCol, setSortCol] = useState("sesiones");
  const [sortDir, setSortDir] = useState("desc");
  const [page, setPage] = useState(0);

  const dSearch = useDebounce(search, 200);

  const students = useMemo(() => {
    const m = {};
    data.forEach(r => {
      if (!m[r.matricula]) m[r.matricula] = { matricula:r.matricula, records:[], servicesSet:new Set() };
      m[r.matricula].records.push(r);
      m[r.matricula].servicesSet.add(r.servicio);
    });
    return Object.values(m).map(s => {
      const sorted = [...s.records].sort((a, b) => (b.fecha || 0) - (a.fecha || 0));
      const latest = sorted[0];
      const bestNombre = sorted.find(r => r.nombre?.trim())?.nombre || "Sin nombre";
      return {
        matricula: s.matricula,
        nombre: bestNombre,
        sesiones: s.records.length,
        servicios: [...s.servicesSet].join(", "),
        ultimoServicio: latest?.servicio || "—",
        ultimoAsesor: latest?.asesor || "—",
        escuela: latest?.escuela || "—",
        programa: latest?.programa || "—",
        interes: latest?.interes || "—",
        comunidad: latest?.comunidad || "—",
        isCAGS: s.records.some(r => r.isCAGS),
        isDIC25: s.records.some(r => r.isDIC25),
        records: s.records,
      };
    });
  }, [data]);

  const asesores   = useMemo(() => [...new Set(data.map(r => r.asesor))].sort(),    [data]);
  const escuelas   = useMemo(() => [...new Set(data.map(r => r.escuela))].sort(),   [data]);
  const estatuses  = useMemo(() => [...new Set(data.map(r => r.estatus))].sort(),   [data]);
  const intereses  = useMemo(() => [...new Set(data.map(r => r.interes))].sort(),   [data]);
  const programas  = useMemo(() => [...new Set(data.map(r => r.programa))].sort(),  [data]);
  const servicios  = useMemo(() => [...new Set(data.map(r => r.servicio))].sort(),  [data]);
  const comunidades = useMemo(() => [...new Set(data.map(r => r.comunidad))].sort(), [data]);

  const filtered = useMemo(() => {
    let f = students;
    if (dSearch) {
      const q = norm(dSearch);
      f = f.filter(s => norm(s.nombre).includes(q) || norm(s.matricula).includes(q) || norm(s.programa).includes(q));
    }
    if (fAsesor)   f = f.filter(s => s.records.some(r => r.asesor    === fAsesor));
    if (fEscuela)  f = f.filter(s => s.escuela  === fEscuela);
    if (fEstatus)  f = f.filter(s => s.records.some(r => r.estatus   === fEstatus));
    if (fInteres)  f = f.filter(s => s.interes  === fInteres);
    if (fPrograma) f = f.filter(s => s.programa === fPrograma);
    if (fServicio)  f = f.filter(s => s.records.some(r => r.servicio === fServicio));
    if (fComunidad) f = f.filter(s => s.comunidad === fComunidad);
    if (fCAGS)  f = f.filter(s => s.isCAGS);
    if (fDIC25) f = f.filter(s => s.isDIC25);
    return [...f].sort((a, b) => {
      let av = a[sortCol] ?? "", bv = b[sortCol] ?? "";
      if (typeof av === "string") av = norm(av);
      if (typeof bv === "string") bv = norm(bv);
      if (av < bv) return sortDir === "asc" ? -1 : 1;
      if (av > bv) return sortDir === "asc" ? 1 : -1;
      return 0;
    });
  }, [students, dSearch, fAsesor, fEscuela, fEstatus, fInteres, fPrograma, fServicio, fComunidad, fCAGS, fDIC25, sortCol, sortDir]);

  useEffect(() => setPage(0), [dSearch, fAsesor, fEscuela, fEstatus, fInteres, fPrograma, fServicio, fComunidad, fCAGS, fDIC25, sortCol]);

  const totalPages = Math.ceil(filtered.length / PAGE_SIZE);
  const displayed = dSearch ? filtered : filtered.slice(page * PAGE_SIZE, (page + 1) * PAGE_SIZE);

  const handleSort = (col) => {
    if (col === sortCol) setSortDir(d => d === "asc" ? "desc" : "asc");
    else { setSortCol(col); setSortDir("desc"); }
  };

  const clearFilters = () => { setSearch(""); setFAsesor(""); setFEscuela(""); setFEstatus(""); setFInteres(""); setFPrograma(""); setFServicio(""); setFComunidad(""); setFCAGS(false); setFDIC25(false); setPage(0); };

  const dlFiltered = () => {
    if (!filtered.length) return;
    dlXl(filtered.map(s => ({
      Matrícula: s.matricula, Nombre: s.nombre, Sesiones: s.sesiones,
      CAGS: s.isCAGS ? "Sí" : "No", DIC25: s.isDIC25 ? "Sí" : "No",
      "Último servicio": s.ultimoServicio, "Último asesor": s.ultimoAsesor,
      Escuela: s.escuela, Programa: s.programa, Interés: s.interes, Comunidad: s.comunidad
    })), "alumnos_filtrados.xlsx");
  };

  const SortTh = ({ col, label, align }) => {
    const active = sortCol === col;
    return (
      <th onClick={() => handleSort(col)} style={{ textAlign:align||"left", padding:"8px 10px", fontSize:10, textTransform:"uppercase", letterSpacing:1, color:active?"#a5b4fc":"#6b6f82", fontWeight:600, whiteSpace:"nowrap", cursor:"pointer", userSelect:"none" }}>
        {label} {active ? (sortDir === "asc" ? "↑" : "↓") : <span style={{ opacity:0.3 }}>↕</span>}
      </th>
    );
  };

  return (
    <div>
      {selectedStudent && <StudentModal student={selectedStudent} records={selectedStudent.records} onClose={() => setSelectedStudent(null)} />}
      <div style={{ display:"flex", gap:8, marginBottom:10, flexWrap:"wrap", alignItems:"center" }}>
        <input style={{ ...S.input, maxWidth:280 }} placeholder="Buscar nombre, matrícula, programa..." value={search} onChange={e => setSearch(e.target.value)} />
        <select style={S.select} value={fAsesor} onChange={e => setFAsesor(e.target.value)}>
          <option value="">Todos los asesores</option>
          {asesores.map(a => <option key={a} value={a}>{a}</option>)}
        </select>
        <select style={S.select} value={fEscuela} onChange={e => setFEscuela(e.target.value)}>
          <option value="">Todas las escuelas</option>
          {escuelas.map(e => <option key={e} value={e}>{e}</option>)}
        </select>
        <select style={S.select} value={fEstatus} onChange={e => setFEstatus(e.target.value)}>
          <option value="">Todos los estatus</option>
          {estatuses.map(e => <option key={e} value={e}>{e}</option>)}
        </select>
      </div>
      <div style={{ display:"flex", gap:8, marginBottom:14, flexWrap:"wrap", alignItems:"center" }}>
        <select style={S.select} value={fInteres} onChange={e => setFInteres(e.target.value)}>
          <option value="">Todo interés</option>
          {intereses.map(v => <option key={v} value={v}>{v}</option>)}
        </select>
        <select style={S.select} value={fPrograma} onChange={e => setFPrograma(e.target.value)}>
          <option value="">Todos los programas</option>
          {programas.map(v => <option key={v} value={v}>{v}</option>)}
        </select>
        <select style={S.select} value={fServicio} onChange={e => setFServicio(e.target.value)}>
          <option value="">Todos los servicios</option>
          {servicios.map(v => <option key={v} value={v}>{v}</option>)}
        </select>
        <select style={S.select} value={fComunidad} onChange={e => setFComunidad(e.target.value)}>
          <option value="">Todas las comunidades</option>
          {comunidades.map(v => <option key={v} value={v}>{v}</option>)}
        </select>
        <button onClick={() => { setFCAGS(v => !v); setFDIC25(false); }}
          style={{ ...S.btn(fCAGS ? "#a855f7" : "#8e92a6"), fontSize:11, padding:"5px 12px", opacity: fDIC25 ? 0.4 : 1 }}>
          ★ Solo CAGS
        </button>
        <button onClick={() => { setFDIC25(v => !v); setFCAGS(false); }}
          style={{ ...S.btn(fDIC25 ? "#22d3ee" : "#8e92a6"), fontSize:11, padding:"5px 12px", opacity: fCAGS ? 0.4 : 1 }}>
          ✓ Solo DIC25
        </button>
        <Bt color="#8e92a6" onClick={clearFilters} style={{ fontSize:11 }}>Limpiar</Bt>
        <Bt color="#10b981" onClick={dlFiltered} style={{ fontSize:11, padding:"5px 12px" }}>↓ Excel</Bt>
        <span style={{ ...S.mono, fontSize:11, color:"#6b6f82", marginLeft:"auto" }}>
          {filtered.length} alumno{filtered.length !== 1 ? "s" : ""}{dSearch ? " encontrado" + (filtered.length !== 1 ? "s" : "") : ""}
        </span>
      </div>

      <div style={{ overflowX:"auto" }}>
        <table style={{ width:"100%", borderCollapse:"collapse", fontSize:12 }}>
          <thead><tr style={{ borderBottom:"1px solid rgba(255,255,255,0.08)" }}>
            <SortTh col="matricula" label="Matrícula" />
            <SortTh col="nombre" label="Nombre" />
            <SortTh col="sesiones" label="Sesiones" align="center" />
            <th style={{ textAlign:"left", padding:"8px 10px", fontSize:10, textTransform:"uppercase", letterSpacing:1, color:"#6b6f82", fontWeight:600, whiteSpace:"nowrap" }}>Servicios</th>
            <SortTh col="ultimoServicio" label="Último servicio" />
            <SortTh col="ultimoAsesor" label="Asesor" />
            <SortTh col="escuela" label="Escuela" />
            <SortTh col="programa" label="Programa" />
            <SortTh col="interes" label="Interés" />
            <SortTh col="comunidad" label="Comunidad" />
          </tr></thead>
          <tbody>{displayed.map(s => (
            <tr key={s.matricula} style={{ borderBottom:"1px solid rgba(255,255,255,0.04)", cursor:"pointer" }}
              onClick={() => setSelectedStudent(s)}
              onMouseEnter={e => e.currentTarget.style.background="rgba(99,102,241,0.06)"}
              onMouseLeave={e => e.currentTarget.style.background="transparent"}>
              <td style={{ padding:"8px 10px", ...S.mono, fontSize:11, color:"#a5b4fc" }}><Highlight text={s.matricula} query={dSearch} /></td>
              <td style={{ padding:"8px 10px", fontWeight:500 }}>
                <Highlight text={s.nombre} query={dSearch} />
                {s.isCAGS  && <span style={{ ...S.badge("#a855f7"), marginLeft:6, fontSize:9 }}>CAGS</span>}
                {s.isDIC25 && <span style={{ ...S.badge("#22d3ee"), marginLeft:6, fontSize:9 }}>DIC25</span>}
              </td>
              <td style={{ padding:"8px 10px", textAlign:"center" }}><span style={S.badge("#6366f1")}>{s.sesiones}</span></td>
              <td style={{ padding:"8px 10px", maxWidth:180, overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap", fontSize:11, color:"#8e92a6" }}>{s.servicios}</td>
              <td style={{ padding:"8px 10px", fontSize:11 }}>{s.ultimoServicio}</td>
              <td style={{ padding:"8px 10px", fontSize:11 }}>{s.ultimoAsesor}</td>
              <td style={{ padding:"8px 10px", fontSize:11, color:"#8e92a6" }}>{s.escuela}</td>
              <td style={{ padding:"8px 10px" }}><span style={S.badge("#3b82f6")}><Highlight text={s.programa} query={dSearch} /></span></td>
              <td style={{ padding:"8px 10px", fontSize:11, color:"#8e92a6" }}>{s.interes}</td>
              <td style={{ padding:"8px 10px", fontSize:11, color:"#8e92a6" }}>{s.comunidad}</td>
            </tr>
          ))}</tbody>
        </table>

        {dSearch && filtered.length === 0 && (
          <div style={{ textAlign:"center", padding:40, color:"#6b6f82", fontSize:13 }}>
            No se encontraron alumnos para <span style={{ color:"#a5b4fc" }}>"{search}"</span>
          </div>
        )}

        {!dSearch && totalPages > 1 && (
          <div style={{ display:"flex", justifyContent:"center", alignItems:"center", gap:10, padding:16 }}>
            <Bt color="#6366f1" onClick={() => setPage(p => Math.max(0, p - 1))} style={{ padding:"4px 14px", fontSize:12, opacity:page===0?0.3:1, pointerEvents:page===0?"none":"auto" }}>← Ant</Bt>
            <span style={{ ...S.mono, fontSize:12, color:"#8e92a6" }}>Página {page + 1} de {totalPages} · {filtered.length} alumnos</span>
            <Bt color="#6366f1" onClick={() => setPage(p => Math.min(totalPages - 1, p + 1))} style={{ padding:"4px 14px", fontSize:12, opacity:page===totalPages-1?0.3:1, pointerEvents:page===totalPages-1?"none":"auto" }}>Sig →</Bt>
          </div>
        )}
      </div>
    </div>
  );
}

/* ═══════════════ TAB: PERSONALIZADO ═══════════════ */
function TabCustom({ data }) {
  const [groupBy, setGroupBy] = useState("servicio");
  const [chartType, setChartType] = useState("barH");
  const [filterDim, setFilterDim] = useState("");
  const [filterVal, setFilterVal] = useState("");
  const [dateFrom, setDateFrom] = useState("");
  const [dateTo, setDateTo] = useState("");
  const chartRef = useRef();

  const filtered = useMemo(() => {
    let f = data;
    if (filterDim && filterVal) f = f.filter(r => r[filterDim] === filterVal);
    if (dateFrom) { const d = new Date(dateFrom + "T00:00:00"); f = f.filter(r => r.fecha && r.fecha >= d); }
    if (dateTo) { const d = new Date(dateTo + "T23:59:59"); f = f.filter(r => r.fecha && r.fecha <= d); }
    return f;
  }, [data, filterDim, filterVal, dateFrom, dateTo]);

  const grouped = useMemo(() => countBy(filtered, groupBy), [filtered, groupBy]);
  const total = filtered.length;

  const filterValues = useMemo(() => {
    if (!filterDim) return [];
    return [...new Set(data.map(r => r[filterDim]))].sort();
  }, [data, filterDim]);

  const dimLabels = { servicio:"Servicio", asesor:"Asesor", escuela:"Escuela", programa:"Programa", estatus:"Estatus", interes:"Interés", modalidad:"Modalidad", comunidad:"Comunidad", semestre:"Semestre" };

  return (
    <div>
      <Cd>
        <div style={{ display:"flex", gap:12, flexWrap:"wrap", alignItems:"center" }}>
          <div>
            <div style={{ fontSize:10, color:"#6b6f82", marginBottom:4, textTransform:"uppercase", letterSpacing:1 }}>Agrupar por</div>
            <select style={S.select} value={groupBy} onChange={e => setGroupBy(e.target.value)}>
              {DIMS.map(d => <option key={d} value={d}>{dimLabels[d]}</option>)}
            </select>
          </div>
          <div>
            <div style={{ fontSize:10, color:"#6b6f82", marginBottom:4, textTransform:"uppercase", letterSpacing:1 }}>Tipo de gráfica</div>
            <select style={S.select} value={chartType} onChange={e => setChartType(e.target.value)}>
              <option value="barH">Barras horizontal</option>
              <option value="barV">Barras vertical</option>
              <option value="pie">Pie / Dona</option>
            </select>
          </div>
          <div>
            <div style={{ fontSize:10, color:"#6b6f82", marginBottom:4, textTransform:"uppercase", letterSpacing:1 }}>Filtrar por</div>
            <select style={S.select} value={filterDim} onChange={e => { setFilterDim(e.target.value); setFilterVal(""); }}>
              <option value="">Sin filtro</option>
              {DIMS.map(d => <option key={d} value={d}>{dimLabels[d]}</option>)}
            </select>
          </div>
          {filterDim && (
            <div>
              <div style={{ fontSize:10, color:"#6b6f82", marginBottom:4, textTransform:"uppercase", letterSpacing:1 }}>Valor</div>
              <select style={S.select} value={filterVal} onChange={e => setFilterVal(e.target.value)}>
                <option value="">Todos</option>
                {filterValues.map(v => <option key={v} value={v}>{v}</option>)}
              </select>
            </div>
          )}
          <div>
            <div style={{ fontSize:10, color:"#6b6f82", marginBottom:4, textTransform:"uppercase", letterSpacing:1 }}>Desde</div>
            <input type="date" style={{ ...S.input, width:"auto" }} value={dateFrom} onChange={e => setDateFrom(e.target.value)} />
          </div>
          <div>
            <div style={{ fontSize:10, color:"#6b6f82", marginBottom:4, textTransform:"uppercase", letterSpacing:1 }}>Hasta</div>
            <input type="date" style={{ ...S.input, width:"auto" }} value={dateTo} onChange={e => setDateTo(e.target.value)} />
          </div>
          {(dateFrom || dateTo) && (
            <div style={{ alignSelf:"flex-end" }}>
              <Bt color="#ef4444" onClick={() => { setDateFrom(""); setDateTo(""); }} style={{ fontSize:11, padding:"8px 12px" }}>✕ Fechas</Bt>
            </div>
          )}
        </div>
      </Cd>

      <ChartCard title={`${dimLabels[groupBy] || groupBy} ${filterVal ? `(filtrado: ${filterVal})` : ""}`} chartRef={chartRef} filename={`custom_${groupBy}.png`}>
        <div ref={chartRef}>
          {chartType === "barH" && (
            <ResponsiveContainer width="100%" height={Math.max(200, grouped.length * 30)}>
              <BarChart data={grouped} layout="vertical" margin={{ left:120, right:20, top:5, bottom:5 }}>
                <XAxis type="number" tick={{ fill:"#6b6f82", fontSize:10 }} />
                <YAxis type="category" dataKey="name" tick={{ fill:"#8e92a6", fontSize:10 }} width={115} />
                <Tooltip content={<TT />} />
                <Bar dataKey="value" radius={[0,6,6,0]}>{grouped.map((_, i) => <Cell key={i} fill={CHART_COLORS[i % CHART_COLORS.length]} />)}</Bar>
              </BarChart>
            </ResponsiveContainer>
          )}
          {chartType === "barV" && (
            <ResponsiveContainer width="100%" height={300}>
              <BarChart data={grouped} margin={{ top:5, right:20, bottom:40, left:10 }}>
                <XAxis dataKey="name" tick={{ fill:"#6b6f82", fontSize:9 }} angle={-35} textAnchor="end" height={60} />
                <YAxis tick={{ fill:"#6b6f82", fontSize:10 }} />
                <Tooltip content={<TT />} />
                <Bar dataKey="value" radius={[6,6,0,0]}>{grouped.map((_, i) => <Cell key={i} fill={CHART_COLORS[i % CHART_COLORS.length]} />)}</Bar>
              </BarChart>
            </ResponsiveContainer>
          )}
          {chartType === "pie" && (
            <ResponsiveContainer width="100%" height={300}>
              <PieChart>
                <Pie data={grouped} dataKey="value" nameKey="name" cx="50%" cy="50%" outerRadius={100} innerRadius={40} paddingAngle={3}>
                  {grouped.map((_, i) => <Cell key={i} fill={CHART_COLORS[i % CHART_COLORS.length]} />)}
                </Pie>
                <Tooltip content={<TT />} />
                <Legend formatter={(v) => <span style={{ color:"#8e92a6", fontSize:11 }}>{v}</span>} />
              </PieChart>
            </ResponsiveContainer>
          )}
        </div>
      </ChartCard>

      <Cd>
        <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:12 }}>
          <div style={S.h3}>Datos</div>
          <div style={{ display:"flex", gap:8 }}>
            <Bt color="#8e92a6" onClick={() => { const svg = chartRef.current?.querySelector("svg"); if(svg) dlPng(svg, `custom_${groupBy}.png`); }} style={{ fontSize:11, padding:"4px 10px" }}>📷 PNG</Bt>
            <Bt color="#6366f1" onClick={() => dlXl(grouped.map(g => ({ [dimLabels[groupBy]]:g.name, Cantidad:g.value, Porcentaje:`${((g.value/total)*100).toFixed(1)}%` })), `custom_${groupBy}.xlsx`)} style={{ fontSize:11, padding:"4px 10px" }}>↓ Excel</Bt>
          </div>
        </div>
        <table style={{ width:"100%", borderCollapse:"collapse", fontSize:12 }}>
          <thead><tr style={{ borderBottom:"1px solid rgba(255,255,255,0.08)" }}>
            {[dimLabels[groupBy],"Cantidad","%"].map(h => (
              <th key={h} style={{ textAlign:"left", padding:"8px 10px", fontSize:10, textTransform:"uppercase", letterSpacing:1, color:"#6b6f82", fontWeight:600 }}>{h}</th>
            ))}
          </tr></thead>
          <tbody>{grouped.map((g, i) => (
            <tr key={i} style={{ borderBottom:"1px solid rgba(255,255,255,0.04)" }}
              onMouseEnter={e => e.currentTarget.style.background="rgba(99,102,241,0.06)"}
              onMouseLeave={e => e.currentTarget.style.background="transparent"}>
              <td style={{ padding:"8px 10px" }}><span style={{ display:"inline-block", width:8, height:8, borderRadius:4, background:CHART_COLORS[i%CHART_COLORS.length], marginRight:8 }} />{g.name}</td>
              <td style={{ padding:"8px 10px", ...S.mono }}>{g.value}</td>
              <td style={{ padding:"8px 10px", ...S.mono, color:"#8e92a6" }}>{((g.value / total) * 100).toFixed(1)}%</td>
            </tr>
          ))}</tbody>
        </table>
      </Cd>
    </div>
  );
}

/* ═══════════════ UPLOAD SCREEN ═══════════════ */
function UploadScreen({ onData, error, loading }) {
  const [dragOver, setDragOver] = useState(false);
  const inputRef = useRef();

  const handleFile = (file) => {
    if (!file) return;
    if (!file.name.endsWith(".xlsx") && !file.name.endsWith(".xls")) { onData(null, "Solo se aceptan archivos .xlsx o .xls"); return; }
    onData(file);
  };

  return (
    <div style={{ minHeight:"100vh", display:"flex", alignItems:"center", justifyContent:"center", background:"radial-gradient(ellipse at 50% 0%, rgba(99,102,241,0.12) 0%, rgba(56,189,248,0.05) 40%, #0b1120 70%)" }}>
      <div style={{ textAlign:"center", maxWidth:520 }}>
        <div style={{ width:72, height:72, borderRadius:18, background:"linear-gradient(135deg, #6366f1, #8b5cf6)", display:"flex", alignItems:"center", justifyContent:"center", margin:"0 auto 24px", fontSize:32, fontWeight:800, color:"#fff" }}>C</div>
        <h1 style={{ fontSize:28, fontWeight:800, marginBottom:4 }}>CVDP Empleabilidad</h1>
        <p style={{ color:"#8e92a6", fontSize:14, marginBottom:32 }}>CRM Dashboard — Periodo FJ26</p>

        <div
          style={{ border:`2px dashed ${dragOver ? "#6366f1" : "rgba(255,255,255,0.12)"}`, borderRadius:18, padding:"48px 32px", cursor:"pointer", transition:"all .2s", background: dragOver ? "rgba(99,102,241,0.08)" : "rgba(255,255,255,0.02)" }}
          onClick={() => inputRef.current?.click()}
          onDragOver={e => { e.preventDefault(); setDragOver(true); }}
          onDragLeave={() => setDragOver(false)}
          onDrop={e => { e.preventDefault(); setDragOver(false); handleFile(e.dataTransfer.files[0]); }}>
          <input ref={inputRef} type="file" accept=".xlsx,.xls" style={{ display:"none" }} onChange={e => handleFile(e.target.files[0])} />
          <div style={{ fontSize:40, marginBottom:16, opacity:0.5 }}>📊</div>
          <div style={{ fontSize:15, fontWeight:600, marginBottom:8 }}>
            {loading ? (
              <span style={{ display:"flex", alignItems:"center", gap:10, justifyContent:"center" }}>
                <span style={{ display:"inline-block", width:18, height:18, border:"3px solid rgba(99,102,241,0.3)", borderTopColor:"#6366f1", borderRadius:"50%", animation:"spin 0.8s linear infinite" }} />
                Procesando archivo…
              </span>
            ) : "Arrastra tu archivo Excel aquí"}
          </div>
          <div style={{ color:"#6b6f82", fontSize:12 }}>{loading ? "El archivo se procesa en segundo plano" : "o haz clic para seleccionar (.xlsx)"}</div>
          <style>{`@keyframes spin { to { transform: rotate(360deg); } }`}</style>
        </div>

        {error && (
          <div style={{ marginTop:16, background:"rgba(239,68,68,0.1)", border:"1px solid rgba(239,68,68,0.3)", borderRadius:12, padding:"12px 16px", color:"#ef4444", fontSize:13 }}>
            {error}
          </div>
        )}

        <div style={{ marginTop:32, textAlign:"left", background:"rgba(255,255,255,0.03)", borderRadius:14, padding:20 }}>
          <div style={{ fontSize:12, fontWeight:600, color:"#8e92a6", marginBottom:10 }}>Columnas esperadas:</div>
          <div style={{ display:"flex", flexWrap:"wrap", gap:6 }}>
            {["Matrícula","Nombre","AP","AM","Servicio","Atiende","Escuela","Programa","Estatus","Interés","Modalidad","Comunidad","Día"].map(c => (
              <span key={c} style={{ ...S.badge("#6366f1"), fontSize:10 }}>{c}</span>
            ))}
          </div>
          <div style={{ marginTop:12, fontSize:11, color:"#6b6f82" }}>Todo se procesa en tu navegador — ningún dato sale de tu computadora.</div>
        </div>
      </div>
    </div>
  );
}

/* ═══════════════ MAIN CRM COMPONENT ═══════════════ */
export default function CRM() {
  const [data, setData] = useState(null);
  const [tab, setTab] = useState("dashboard");
  const [error, setError] = useState(null);
  const [loading, setLoading] = useState(false);

  const handleFile = useCallback(async (file, errMsg) => {
    if (errMsg) { setError(errMsg); return; }
    setLoading(true);
    setError(null);
    try {
      const buffer = await file.arrayBuffer();
      const worker = new Worker(new URL("./parseWorker.js", import.meta.url), { type: "module" });
      worker.postMessage({ buffer }, [buffer]);
      worker.onmessage = ({ data: res }) => {
        worker.terminate();
        if (res.ok) { setData(res.data); setTab("dashboard"); }
        else setError(res.error);
        setLoading(false);
      };
      worker.onerror = (e) => {
        worker.terminate();
        setError(e.message || "Error procesando el archivo");
        setLoading(false);
      };
    } catch (e) {
      setError(e.message || "Error leyendo el archivo");
      setLoading(false);
    }
  }, []);

  const reset = () => { setData(null); setError(null); setTab("dashboard"); };

  if (!data) return <UploadScreen onData={handleFile} error={error} loading={loading} />;

  return (
    <div style={{ minHeight:"100vh" }}>
      {/* HEADER */}
      <header style={{ position:"sticky", top:0, zIndex:100, background:"rgba(11,17,32,0.92)", backdropFilter:"blur(16px)", borderBottom:"1px solid rgba(255,255,255,0.06)", padding:"0 24px" }}>
        <div style={{ display:"flex", alignItems:"center", height:56, maxWidth:1400, margin:"0 auto" }}>
          <div style={{ display:"flex", alignItems:"center", gap:10, marginRight:32 }}>
            <div style={{ width:30, height:30, borderRadius:8, background:"linear-gradient(135deg, #6366f1, #8b5cf6)", display:"flex", alignItems:"center", justifyContent:"center", fontSize:14, fontWeight:800, color:"#fff" }}>C</div>
            <span style={{ fontWeight:700, fontSize:14 }}>CVDP <span style={{ color:"#8e92a6", fontWeight:400 }}>Empleabilidad</span></span>
          </div>
          <nav style={{ display:"flex", gap:0, flex:1 }}>
            {TABS.map(t => (
              <button key={t.id} onClick={() => setTab(t.id)}
                style={{ background:"none", border:"none", borderBottom: tab === t.id ? "2px solid #6366f1" : "2px solid transparent", color: tab === t.id ? "#a5b4fc" : "#7d8296", padding:"16px 16px", fontSize:13, fontWeight: tab === t.id ? 600 : 400, cursor:"pointer", fontFamily:"'Plus Jakarta Sans'", transition:"all .2s", display:"flex", alignItems:"center", gap:6 }}>
                <span style={{ fontSize:14 }}>{t.icon}</span> {t.label}
              </button>
            ))}
          </nav>
          <div style={{ display:"flex", alignItems:"center", gap:12 }}>
            <span style={{ ...S.mono, fontSize:11, color:"#6b6f82" }}>{data.length.toLocaleString()} registros</span>
            <Bt color="#ef4444" onClick={reset} style={{ fontSize:11, padding:"5px 12px" }}>Nuevo archivo</Bt>
          </div>
        </div>
      </header>

      {/* CONTENT */}
      <main style={{ maxWidth:1400, margin:"0 auto", padding:"24px 24px 60px" }}>
        {tab === "dashboard" && <TabDashboard data={data} />}
        {tab === "asesores" && <TabAsesores data={data} />}
        {tab === "pipeline" && <TabPipeline data={data} />}
        {tab === "alumnos" && <TabAlumnos data={data} />}
        {tab === "custom" && <TabCustom data={data} />}
      </main>
    </div>
  );
}
