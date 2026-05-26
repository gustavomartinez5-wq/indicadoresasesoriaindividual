import * as XLSX from "xlsx";

/* ═══ NORMALIZACIÓN DE CAMPOS ═══ */
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
function norm(s) {
  return String(s ?? "").toLowerCase().normalize("NFD").replace(/[̀-ͯ]/g, "");
}

/* ═══ FECHAS ═══ */
function parseDate(v) {
  if (!v && v !== 0) return null;
  if (typeof v === "number") {
    const d = new Date(Math.round((v - 25569) * 86400000));
    return isNaN(d) ? null : d;
  }
  if (v instanceof Date) return isNaN(v) ? null : v;
  const d = new Date(String(v).trim());
  return isNaN(d) ? null : d;
}
function weekNum(d) {
  if (!d) return 0;
  const jan1 = new Date(d.getFullYear(), 0, 1);
  return Math.ceil((((d - jan1) / 86400000) + jan1.getDay() + 1) / 7);
}

/* ═══ PARSEO PRINCIPAL ═══ */
function parseBuffer(buffer) {
  // cellDates: false evita la conversión eager de fechas → 2-3x más rápido
  const wb = XLSX.read(buffer, { type: "array", cellDates: false });
  const sheetName = wb.SheetNames.find(n => n.toLowerCase().includes("agenda")) || wb.SheetNames[0];
  const ws = wb.Sheets[sheetName];
  const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });

  let headerIdx = -1;
  for (let i = 0; i < Math.min(15, raw.length); i++) {
    const row = raw[i].map(c => String(c).toLowerCase());
    if (row.some(c => c.includes("matrícula") || c.includes("matricula"))) {
      headerIdx = i; break;
    }
  }
  if (headerIdx === -1) throw new Error("No se encontró la fila de encabezado con 'Matrícula'");

  const headers = raw[headerIdx].map(c => String(c).toLowerCase().trim());
  const col = (keywords) => {
    for (const kw of keywords) {
      const idx = headers.findIndex(h => h.includes(kw));
      if (idx !== -1) return idx;
    }
    return -1;
  };

  const colAP = col(["ap"]);
  const colAM = headers.findIndex((h, i) => h === "am" && i > colAP);
  const cols = {
    dia:      col(["día", "dia", "fecha"]),
    matricula:col(["matrícula", "matricula"]),
    nombre:   col(["nombre"]),
    ap:       colAP,
    am:       colAM !== -1 ? colAM : -1,
    servicio: col(["servicio"]),
    atiende:  col(["atiende"]),
    escuela:  col(["escuela"]),
    programa: col(["programa"]),
    estatus:  col(["estatus"]),
    interes:  col(["interés", "interes"]),
    modalidad:col(["modalidad"]),
    comunidad:col(["comunidad"]),
    campus:   col(["campus"]),
    semestre: col(["semestre"]),
    cag:      col(["cag"]),
    exatec:   col(["exatec"]),
  };

  const data = [];
  for (let i = headerIdx + 1; i < raw.length; i++) {
    const r = raw[i];
    if (!r || r.length === 0) continue;
    const get = (k) => cols[k] >= 0 ? String(r[cols[k]] ?? "").trim() : "";
    const mat = get("matricula");
    if (!mat) continue;

    const fullName = [get("nombre"), get("ap"), get("am")].filter(Boolean).join(" ");
    const fecha = cols.dia >= 0 ? parseDate(r[cols.dia]) : null;
    const semVal = get("semestre");
    const cagVal = get("cag");
    const exatecVal = get("exatec");

    data.push({
      fecha,
      semana:    weekNum(fecha),
      matricula: mat,
      nombre:    fullName,
      servicio:  nSrv(get("servicio")),
      asesor:    get("atiende") || "Sin asesor",
      escuela:   nEsc(get("escuela")),
      programa:  get("programa") || "Sin programa",
      estatus:   get("estatus") || "Sin estatus",
      interes:   nInt(get("interes")),
      modalidad: nMod(get("modalidad")),
      comunidad: get("comunidad") || "Sin comunidad",
      campus:    get("campus") || "Sin campus",
      semestre:  semVal || "Sin semestre",
      isCAGS:    norm(semVal).startsWith("8") || norm(cagVal).startsWith("s"),
      isDIC25:   norm(exatecVal).includes("diciembre") && exatecVal.includes("2025"),
    });
  }

  if (data.length === 0) throw new Error("No se encontraron registros válidos");
  return data;
}

/* ═══ ENTRADA DEL WORKER ═══ */
self.onmessage = ({ data: { buffer } }) => {
  try {
    const data = parseBuffer(buffer);
    self.postMessage({ ok: true, data });
  } catch (err) {
    self.postMessage({ ok: false, error: err.message || String(err) });
  }
};
