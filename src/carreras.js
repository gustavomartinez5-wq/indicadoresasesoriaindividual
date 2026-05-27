export const ESCUELAS_TEC = [
  { key: "EIC",  siglas: "EIC",  nombre: "Ingeniería y Ciencias" },
  { key: "EN",   siglas: "EN",   nombre: "Negocios" },
  { key: "EAAD", siglas: "EAAD", nombre: "Arquitectura, Arte y Diseño" },
  { key: "EHE",  siglas: "EHE",  nombre: "Humanidades y Educación" },
  { key: "ECSG", siglas: "ECSG", nombre: "Ciencias Sociales y Gobierno" },
  { key: "EMCS", siglas: "EMCS", nombre: "Medicina y Ciencias de la Salud" },
  { key: "POS",  siglas: "POS",  nombre: "Posgrado" },
];

export const CARRERAS = [
  { siglas: "ARQ",  nombre: "Arquitectura" },
  { siglas: "LAD",  nombre: "Animación y Arte Digital" },
  { siglas: "LDI",  nombre: "Diseño Industrial" },
  { siglas: "LUB",  nombre: "Urbanismo" },
  { siglas: "LEC",  nombre: "Economía" },
  { siglas: "LED",  nombre: "Derecho" },
  { siglas: "LRI",  nombre: "Relaciones Internacionales" },
  { siglas: "LTP",  nombre: "Gobierno y Transformación Pública" },
  { siglas: "LC",   nombre: "Comunicación" },
  { siglas: "LEI",  nombre: "Innovación Educativa" },
  { siglas: "LLE",  nombre: "Letras Hispánicas" },
  { siglas: "LPE",  nombre: "Periodismo" },
  { siglas: "LTM",  nombre: "Tecnología y Producción Musical" },
  { siglas: "IAL",  nombre: "Ing. en Alimentos" },
  { siglas: "IBT",  nombre: "Ing. en Biotecnología" },
  { siglas: "IC",   nombre: "Ingeniería Civil" },
  { siglas: "IDM",  nombre: "Ciencia de Datos y Matemáticas" },
  { siglas: "IDS",  nombre: "Ing. en Desarrollo Sustentable" },
  { siglas: "IE",   nombre: "Ing. en Electrónica" },
  { siglas: "IFI",  nombre: "Ing. en Física Industrial" },
  { siglas: "IIA",  nombre: "Ing. en Industrias Alimentarias" },
  { siglas: "IID",  nombre: "Ing. en Innovación y Desarrollo" },
  { siglas: "IIS",  nombre: "Ing. Industrial y de Sistemas" },
  { siglas: "IM",   nombre: "Ingeniería Mecánica" },
  { siglas: "IMD",  nombre: "Ing. Biomédica" },
  { siglas: "IMT",  nombre: "Ing. en Mecatrónica" },
  { siglas: "INA",  nombre: "Ing. en Nanotecnología" },
  { siglas: "IQ",   nombre: "Ingeniería Química" },
  { siglas: "IRS",  nombre: "Ing. en Robótica y Sistemas Digitales" },
  { siglas: "ITC",  nombre: "Ing. en Tecnologías Computacionales" },
  { siglas: "ITD",  nombre: "Ing. en Transformación Digital de Negocios" },
  { siglas: "LBC",  nombre: "Biociencias" },
  { siglas: "LNB",  nombre: "Nutrición y Bienestar Integral" },
  { siglas: "LPS",  nombre: "Psicología Clínica y de la Salud" },
  { siglas: "MC",   nombre: "Médico Cirujano" },
  { siglas: "MO",   nombre: "Médico Cirujano Odontólogo" },
  { siglas: "LAE",  nombre: "Estrategia y Transformación de Negocios" },
  { siglas: "LAF",  nombre: "Administración Financiera" },
  { siglas: "LCPF", nombre: "Contaduría Pública y Finanzas" },
  { siglas: "LDE",  nombre: "Emprendimiento" },
  { siglas: "LDO",  nombre: "Desarrollo de Talento y Cultura Organizacional" },
  { siglas: "LEM",  nombre: "Mercadotecnia" },
  { siglas: "LIN",  nombre: "Negocios Internacionales" },
  { siglas: "LIT",  nombre: "Inteligencia de Negocios" },
  { siglas: "LPO",  nombre: "Psicología Organizacional" },
  { siglas: "BGB",  nombre: "B.A. in International Business" },
  { siglas: "POS",  nombre: "Posgrado" },
];

export const CARRERA_ESCUELA = {
  ARQ: "EAAD", LAD: "EAAD", LDI: "EAAD", LUB: "EAAD",
  LEC: "ECSG", LED: "ECSG", LRI: "ECSG", LTP: "ECSG",
  LC:  "EHE",  LEI: "EHE",  LLE: "EHE",  LPE: "EHE",  LTM: "EHE",
  IAL: "EIC",  IBT: "EIC",  IC:  "EIC",  IDM: "EIC",  IDS: "EIC",
  IE:  "EIC",  IFI: "EIC",  IIA: "EIC",  IID: "EIC",  IIS: "EIC",
  IM:  "EIC",  IMD: "EIC",  IMT: "EIC",  INA: "EIC",  IQ:  "EIC",
  IRS: "EIC",  ITC: "EIC",  ITD: "EIC",
  LBC: "EMCS", LNB: "EMCS", LPS: "EMCS", MC:  "EMCS", MO:  "EMCS",
  LAE: "EN",   LAF: "EN",   LCPF: "EN",  LDE: "EN",   LDO: "EN",
  LEM: "EN",   LIN: "EN",   LIT: "EN",   LPO: "EN",   BGB: "EN",
  POS: "POS",
};

// Puente retrocompatible: nombre completo (registros viejos) → key
export const ESCUELA_A_KEY = {
  "Escuela de Ingeniería y Ciencias":          "EIC",
  "Escuela de Negocios":                       "EN",
  "Escuela de Arquitectura, Arte y Diseño":    "EAAD",
  "Escuela de Humanidades y Educación":        "EHE",
  "Escuela de Ciencias Sociales y Gobierno":   "ECSG",
  "Escuela de Medicina y Ciencias de la Salud":"EMCS",
  EIC: "EIC", EN: "EN", EAAD: "EAAD", EHE: "EHE", ECSG: "ECSG", EMCS: "EMCS", POS: "POS",
  "Posgrado": "POS",
};

// Nombre completo de la carrera por siglas (para tooltips)
export const CARRERA_NOMBRE = Object.fromEntries(CARRERAS.map(c => [c.siglas, c.nombre]));

// Nombre de la escuela por key (para tooltips)
export const ESCUELA_NOMBRE = Object.fromEntries(ESCUELAS_TEC.map(e => [e.key, e.nombre]));
