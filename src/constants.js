export const ASESORES = [
  "Gustavo Martínez",
  "Claudia Toledano",
  "José Casas",
  "Lesly Sánchez",
  "Cecilia García",
  "Carolina Quesada",
];

export const ESCUELAS = [
  "Escuela de Ingeniería y Ciencias",
  "Escuela de Negocios",
  "Escuela de Arquitectura, Arte y Diseño",
  "Escuela de Humanidades y Educación",
  "Escuela de Ciencias Sociales y Gobierno",
  "Escuela de Medicina y Ciencias de la Salud",
];

export const SERVICIOS = [
  { label: "Asesoría Plan de Vida y Carrera", clave: "5.8" },
  { label: "Asesoría CV", clave: "5.5" },
  { label: "Asesoría LinkedIn", clave: "5.6" },
  { label: "Asesoría Bolsa de Trabajo", clave: "5.11" },
  { label: "Asesoría Entrevista Español", clave: "5.7" },
  { label: "Asesoría Entrevista Inglés", clave: "5.7" },
  { label: "Asesoría Portafolio", clave: "5.5" },
  { label: "Asesoría Evaluación Cartas Oferta", clave: "5.12" },
  { label: "Asesoría Cover Letter", clave: "5.5" },
  { label: "Asesorías individuales prácticas", clave: "5.1" },
  { label: "Sesiones Diagnóstico de CV para CAG", clave: "5.5" },
  { label: "Sesiones Diagnóstico de LinkedIn para CAG", clave: "5.6" },
  { label: "Sesión Diagnóstico - Bolsa de Trabajo", clave: "5.11" },
];

export const SERVICIO_CLAVE = Object.fromEntries(
  SERVICIOS.map((s) => [s.label, s.clave])
);

export const ESTATUSES = ["Agendado", "Asistencia", "Falta", "Cancelación", "Express"];

export const SEMESTRES = ["1","2","3","4","5","6","7","8","9","10","11","EXATEC","Posgrado"];

export const CAG_OPTS = ["Sí", "No"];

export const EXATEC_OPTS = ["Generación Diciembre 2025", "EXATEC Años Anteriores"];

export const MODALIDADES = ["Virtual", "Presencial"];

export const INTERESES = [
  "Prácticas profesionales",
  "Estancia Profesional",
  "On Campus Intern",
  "Empleo",
  "Empleo Internacional",
  "Posgrado",
  "Programa Internacional",
  "Servicio Social",
  "Grupos Estudiantiles",
  "Equipo Representativo Deporte",
  "Equipo Representativo Arte",
  "Otro",
];

export const COMUNIDADES = [
  "Ekvilibro","Energio","Forta","Krei","Kresko",
  "Pasio","Reflekto","Revo","Spirita","Talenta","NA",
];

export const STATUS_COLORS = {
  Asistencia: "#10b981",
  Falta: "#ef4444",
  Express: "#f59e0b",
  Cancelación: "#8b5cf6",
  Agendado: "#6366f1",
};

export const CHART_COLORS = [
  "#6366f1","#10b981","#f59e0b","#ef4444","#8b5cf6",
  "#ec4899","#f97316","#14b8a6","#3b82f6","#a855f7",
  "#84cc16","#06b6d4","#e11d48","#22d3ee",
];
