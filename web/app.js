// Consolidación de Actas PYM/PMT/DNT - DUSAKAWI
// App 100% cliente: toda la extracción, validación y consolidación corre en
// el navegador. El estado persiste en localStorage y se puede exportar/
// importar como el Excel maestro para moverlo entre computadores.

const PERIODO_MESES = {
  "OTRO SI": "Ene-Feb (adicion)",
  "1-TRIM": "Ene-Feb-Mar",
  "2-TRIM": "Abr-May-Jun",
  "3-TRIM": "Jul-Ago-Sep",
  "4-TRIM": "Oct-Nov-Dic",
};
const PERIODO_ORDEN = Object.keys(PERIODO_MESES);

const PROGRAMAS_TIPICOS = [
  "INDIVIDUALES PARA NIÑOS Y NIÑAS EN PRIMERA INFANCIA 1M - 5A",
  "INDIVIDUALES PARA NIÑOS Y NIÑAS EN INFANCIA 6 - 11 AÑOS",
  "INDIVIDUALES PARA LOS ADOLESCENTES 12 - 17 AÑOS",
  "INDIVIDUALES PARA LOS JOVENES 18 - 28 AÑOS",
  "INDIVIDUALES PARA LOS ADULTOS 29 - 59 AÑOS",
  "INDIVIDUALES PARA LOS ADULTOS MAYORES 60 A 80 Y MAS",
  "PAI",
  "MATERNO PERINATAL",
  "ATENCION POSPARTO",
  "DEMANDA INDUCIDA",
  "RUTA RIESGO CARDIOVASCULAR BAJA",
];

const STORAGE_KEY = "dusakawi_actas_estado_v1";
const STATE = { detalle: [], validaciones: [] };

// Ruta por defecto para el consolidado por prestador: promoción y
// mantenimiento por ciclo de vida + materno perinatal. Sin PAI, atención
// posparto, demanda inducida ni ruta cardiovascular (se pueden marcar
// aparte si se quieren incluir).
const ACTIVIDADES_RUTA_DEFECTO = new Set([
  "INDIVIDUALES PARA NIÑOS Y NIÑAS EN PRIMERA INFANCIA 1M - 5A",
  "INDIVIDUALES PARA NIÑOS Y NIÑAS EN INFANCIA 6 - 11 AÑOS",
  "INDIVIDUALES PARA LOS ADOLESCENTES 12 - 17 AÑOS",
  "INDIVIDUALES PARA LOS JOVENES 18 - 28 AÑOS",
  "INDIVIDUALES PARA LOS ADULTOS 29 - 59 AÑOS",
  "INDIVIDUALES PARA LOS ADULTOS MAYORES 60 A 80 Y MAS",
  "MATERNO PERINATAL",
]);
let actividadesSeleccionadas = new Set(ACTIVIDADES_RUTA_DEFECTO);

// Ordena las actividades como vienen en el acta original (PROGRAMAS_TIPICOS),
// no alfabéticamente. Cualquier actividad que no esté en la lista típica
// queda al final, ordenada alfabéticamente entre sí.
function compararProgramas(a, b) {
  const ia = PROGRAMAS_TIPICOS.indexOf(a);
  const ib = PROGRAMAS_TIPICOS.indexOf(b);
  const oa = ia === -1 ? PROGRAMAS_TIPICOS.length : ia;
  const ob = ib === -1 ? PROGRAMAS_TIPICOS.length : ib;
  if (oa !== ob) return oa - ob;
  return a.localeCompare(b);
}

// Municipios/prestadores conocidos, para adivinar el campo a partir de la
// ruta de carpeta o el nombre del archivo (más específico primero).
const MUNICIPIOS_CONOCIDOS = [
  "SAN JUAN DEL CESAR", "SAN JUAN",
  "AGUSTIN CODAZZI", "CODAZZI",
  "LA PAZ",
  "BECERRIL",
  "VALLEDUPAR",
  "RIOHACHA",
];

function norm(v) {
  if (v === null || v === undefined) return "";
  return String(v).trim();
}
function fmtMoney(v) {
  return (v || 0).toLocaleString("es-CO", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

// --------------------------------------------------------------------------
// Detección automática de periodo / municipio / contrato a partir de la ruta
// (carpeta de origen) o el nombre del archivo, para no pedirle al usuario
// que los seleccione a mano en cada carga.
// --------------------------------------------------------------------------
function detectarPeriodoDesdeTexto(texto) {
  const upper = (texto || "").toUpperCase();
  if (upper.includes("OTRO SI") || upper.includes("OTROSI")) return "OTRO SI";
  const m = upper.match(/(\d)\s*-?\s*TRIM/);
  if (m) return `${m[1]}-TRIM`;
  return null;
}

function detectarMunicipioDesdeTexto(texto) {
  const upper = (texto || "").toUpperCase();
  for (const m of MUNICIPIOS_CONOCIDOS) {
    if (upper.includes(m)) return m === "SAN JUAN" ? "SAN JUAN DEL CESAR" : m === "CODAZZI" ? "AGUSTIN CODAZZI" : m;
  }
  return null;
}

function detectarContratoDesdeTexto(texto) {
  const m = (texto || "").match(/\b(\d{4,6}-\d{1,4}(?:-[A-Z]{2,5})+)\b/i);
  return m ? m[1].toUpperCase() : null;
}

function rutaArchivo(file) {
  return file.webkitRelativePath || file.name;
}

function extraerAnio(texto) {
  const m = String(texto || "").match(/(20\d{2})/);
  return m ? m[1] : null;
}

// --------------------------------------------------------------------------
// Persistencia (localStorage)
// --------------------------------------------------------------------------
function guardarEstado() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(STATE));
}
function cargarEstado() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return;
    const parsed = JSON.parse(raw);
    STATE.detalle = parsed.detalle || [];
    STATE.validaciones = parsed.validaciones || [];
  } catch (e) {
    console.warn("No se pudo leer el estado guardado:", e);
  }
}

// Borra todo: el consolidado guardado (localStorage), lo que esté en
// revisión pendiente y las tarjetas de PDF cargadas. No queda nada alojado.
function limpiarTodo() {
  const total = STATE.detalle.length + STATE.validaciones.length;
  const ok = confirm(
    total > 0
      ? `Esto borra permanentemente el consolidado guardado en este navegador (${STATE.detalle.length} filas de detalle) y todo lo que esté en revisión pendiente. ¿Continuar?`
      : "Esto borra cualquier revisión pendiente y tarjetas de PDF cargadas. ¿Continuar?"
  );
  if (!ok) return;

  STATE.detalle = [];
  STATE.validaciones = [];
  localStorage.removeItem(STORAGE_KEY);

  STAGING_XLSX.clear();
  renderStagingXlsx();
  document.getElementById("confirmacion-xlsx").innerHTML = "";
  document.getElementById("pdf-items").innerHTML = "";

  const tabResumen = document.getElementById("tab-resumen");
  const tabVal = document.getElementById("tab-validaciones");
  if (tabResumen.classList.contains("active")) cargarResumen();
  if (tabVal.classList.contains("active")) cargarValidaciones();
  else document.getElementById("validaciones-contenido").innerHTML = "";
  if (!tabResumen.classList.contains("active")) document.getElementById("resumen-contenido").innerHTML = "";

  alert("Listo, no queda nada guardado.");
}

// --------------------------------------------------------------------------
// Extraccion desde Excel (hojas tipo "erev")
// --------------------------------------------------------------------------
function parseActaXlsx(workbook, filename) {
  let aoa = null;
  for (const sheetName of workbook.SheetNames) {
    const candidate = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1, defval: null, raw: true });
    const maxScan = Math.min(15, candidate.length);
    let found = false;
    for (let r = 0; r < maxScan; r++) {
      const row = candidate[r] || [];
      for (const cell of row) {
        if (norm(cell).toUpperCase().startsWith("ACTA N")) { found = true; break; }
      }
      if (found) break;
    }
    if (found) { aoa = candidate; break; }
  }
  if (!aoa) throw new Error(`No se encontro encabezado 'ACTA N*' en ${filename}`);

  const fields = {};
  let headerRowIdx = null;
  const maxScan = Math.min(30, aoa.length);
  for (let r = 0; r < maxScan; r++) {
    const row = aoa[r] || [];
    for (let c = 0; c < row.length; c++) {
      const label = norm(row[c]).toUpperCase();
      if (!label) continue;
      if (label.startsWith("ACTA N")) fields.acta_no = norm(row[c + 1]);
      else if (label.startsWith("PERIODO EVALUADO")) fields.meses = norm(row[c + 1]);
      else if (label.startsWith("VIGENCIA DEL CONTRATO")) fields.vigencia_contrato = norm(row[c + 1]);
      else if (label === "EMPRESA") fields.empresa = norm(row[c + 1]);
      else if (label === "NIT") fields.nit = norm(row[c + 1]);
      else if (label === "REGIMEN") fields.regimen = norm(row[c + 1]);
      else if (label === "MUNICIPIO") fields.municipio = norm(row[c + 1]);
      else if (label.startsWith("N") && label.includes("CONTRATO") && !label.includes("VIGENCIA")) fields.contrato = norm(row[c + 1]);
      else if (label === "PROGRAMA") headerRowIdx = r;
    }
  }
  if (headerRowIdx === null) throw new Error(`No se encontro la tabla 'Programa' en ${filename}`);
  if (!fields.contrato) throw new Error(`No se encontro 'Nº CONTRATO' en ${filename}`);

  const headerRow = aoa[headerRowIdx] || [];
  let colPrograma = null, colExigido = null, colReconocido = null, colDescuento = null;
  for (let c = 0; c < headerRow.length; c++) {
    const label = norm(headerRow[c]).toUpperCase();
    if (label === "PROGRAMA") colPrograma = c;
    else if (label === "VR EXIGIDO") colExigido = c;
    else if (label === "VR RECONOCIDO") colReconocido = c;
    else if (label === "DESCUENTO") colDescuento = c;
  }

  const programas = [];
  let totalEjecucion = null;
  let r = headerRowIdx + 1;
  while (r < aoa.length) {
    const row = aoa[r] || [];
    const name = norm(row[colPrograma]);
    if (!name) {
      r++;
      if (r - headerRowIdx > 40) break;
      continue;
    }
    const exigido = row[colExigido];
    const reconocido = row[colReconocido];
    const descuento = row[colDescuento];
    const upper = name.toUpperCase();
    if (upper.startsWith("TOTAL EJECUCION")) {
      totalEjecucion = { vr_exigido: exigido || 0, vr_reconocido: reconocido || 0, descuento: descuento || 0 };
      r++;
      continue;
    }
    if (upper.startsWith("TOTAL DESCUENTO")) break;
    programas.push({ programa: name, vr_exigido: exigido || 0, vr_reconocido: reconocido || 0, descuento: descuento || 0 });
    r++;
  }

  return {
    archivo_origen: filename,
    municipio: fields.municipio || "",
    contrato: fields.contrato || "",
    acta_no: fields.acta_no || "",
    meses: fields.meses || "",
    empresa: fields.empresa || "",
    nit: fields.nit || "",
    regimen: fields.regimen || "",
    vigencia_contrato: fields.vigencia_contrato || "",
    programas,
    total_ejecucion: totalEjecucion,
  };
}

// --------------------------------------------------------------------------
// Validacion
// --------------------------------------------------------------------------
function validar(acta) {
  const checks = [];
  const add = (nombre, ok, detalle) => checks.push({ chequeo: nombre, resultado: ok ? "OK" : "ERROR", detalle });

  const contrato = acta.contrato || "";
  const actaNo = acta.acta_no || "";
  add(
    "Acta_No coincide con Contrato",
    !!contrato && actaNo.startsWith(contrato),
    `contrato='${contrato}' acta_no='${actaNo}'`
  );

  // El prestador que presta el servicio puede variar por municipio (DUSAKAWI
  // IPSI, EZEQ SALUD IPSI, PALAIMA, etc.) — no se exige un nombre fijo, solo
  // que el acta traiga diligenciado quién es.
  const empresa = (acta.empresa || "").trim();
  add(
    "Empresa (prestador) diligenciada",
    empresa.length > 0,
    empresa ? `Prestador: ${empresa}` : "El acta no trae el campo Empresa diligenciado"
  );

  const tol = 1.0;
  (acta.programas || []).forEach((p) => {
    const exigido = parseFloat(p.vr_exigido) || 0;
    const reconocido = parseFloat(p.vr_reconocido) || 0;
    const descuento = parseFloat(p.descuento) || 0;
    const diff = Math.abs(exigido - descuento - reconocido);
    add(
      `Aritmetica fila: ${p.programa}`,
      diff <= tol,
      `Exigido-Descuento=${fmtMoney(exigido - descuento)} vs Reconocido=${fmtMoney(reconocido)} (dif=${fmtMoney(diff)})`
    );
    // Solo se reporta si de verdad hay un problema (para no llenar el
    // historial de chequeos "OK" repetidos sin valor informativo).
    if (exigido > 0 && reconocido > exigido + tol) {
      const pct = (reconocido / exigido) * 100;
      add(
        `Cumplimiento excede 100%: ${p.programa}`,
        false,
        `Vr Reconocido (${fmtMoney(reconocido)}) es mayor que Vr Exigido (${fmtMoney(exigido)}) → ${pct.toFixed(1)}%. Revisa esta fila en el acta original, seguramente hay un error de digitación.`
      );
    }
  });

  const total = acta.total_ejecucion;
  if (total) {
    const sumaExigido = (acta.programas || []).reduce((s, p) => s + (parseFloat(p.vr_exigido) || 0), 0);
    const sumaReconocido = (acta.programas || []).reduce((s, p) => s + (parseFloat(p.vr_reconocido) || 0), 0);
    const sumaDescuento = (acta.programas || []).reduce((s, p) => s + (parseFloat(p.descuento) || 0), 0);
    add(
      "Suma Vr_Exigido = Total Ejecucion",
      Math.abs(sumaExigido - (parseFloat(total.vr_exigido) || 0)) <= tol,
      `suma=${fmtMoney(sumaExigido)} vs total=${fmtMoney(parseFloat(total.vr_exigido) || 0)}`
    );
    add(
      "Suma Vr_Reconocido = Total Ejecucion",
      Math.abs(sumaReconocido - (parseFloat(total.vr_reconocido) || 0)) <= tol,
      `suma=${fmtMoney(sumaReconocido)} vs total=${fmtMoney(parseFloat(total.vr_reconocido) || 0)}`
    );
    add(
      "Suma Descuento = Total Ejecucion",
      Math.abs(sumaDescuento - (parseFloat(total.descuento) || 0)) <= tol,
      `suma=${fmtMoney(sumaDescuento)} vs total=${fmtMoney(parseFloat(total.descuento) || 0)}`
    );
  } else {
    add("Total Ejecucion presente", false, "No se encontro fila TOTAL EJECUCION");
  }

  return checks;
}

const MESES_POR_PERIODO = {
  "OTRO SI": ["ENE", "FEB"],
  "1-TRIM": ["ENE", "FEB", "MAR"],
  "2-TRIM": ["ABR", "MAY", "JUN"],
  "3-TRIM": ["JUL", "AGO", "SEP"],
  "4-TRIM": ["OCT", "NOV", "DIC"],
};

// Compara los meses que el propio documento dice haber evaluado contra el
// periodo asignado (detectado o elegido) — atrapa el caso de un archivo
// guardado o rotulado en la carpeta/periodo equivocado.
function checkConsistenciaMeses(acta, periodo) {
  const esperado = MESES_POR_PERIODO[periodo];
  const meses = (acta.meses || "").toUpperCase();
  if (!esperado || !meses) return null;
  const ok = esperado.some((m) => meses.includes(m));
  return {
    chequeo: "Meses del acta coinciden con el periodo asignado",
    resultado: ok ? "OK" : "ERROR",
    detalle: `periodo='${periodo}' meses_acta='${acta.meses}' (se esperaba alguno de: ${esperado.join(", ")})`,
  };
}

// --------------------------------------------------------------------------
// Consolidacion (en memoria + localStorage)
// --------------------------------------------------------------------------
function consolidar(acta, periodo, checks) {
  const municipio = (acta.municipio || "").trim();
  const contrato = (acta.contrato || "").trim();
  const actaNo = (acta.acta_no || "").trim();

  STATE.detalle = STATE.detalle.filter(
    (row) => !(row.Municipio === municipio && row.Contrato === contrato && row.Acta_No === actaNo && row.Periodo === periodo)
  );
  STATE.validaciones = STATE.validaciones.filter(
    (row) => !(row.Municipio === municipio && row.Contrato === contrato && row.Acta_No === actaNo && row.Periodo === periodo)
  );

  const fechaCarga = new Date().toISOString().slice(0, 10);
  const anio = extraerAnio(acta.vigencia_contrato) || extraerAnio(acta.anio) || String(new Date().getFullYear());
  (acta.programas || []).forEach((p) => {
    const exigido = parseFloat(p.vr_exigido) || 0;
    const reconocido = parseFloat(p.vr_reconocido) || 0;
    const descuento = parseFloat(p.descuento) || 0;
    const pct = exigido ? reconocido / exigido : null;
    STATE.detalle.push({
      Municipio: municipio, Contrato: contrato, Acta_No: actaNo, Periodo: periodo,
      Meses: acta.meses || "", Anio: anio, Empresa: acta.empresa || "", Nit: acta.nit || "", Regimen: acta.regimen || "",
      Programa: p.programa, Vr_Exigido: exigido, Vr_Reconocido: reconocido, Descuento: descuento,
      Pct_Cumplimiento: pct, Archivo_Origen: acta.archivo_origen || "", Fecha_Carga: fechaCarga,
    });
  });

  checks.forEach((c) => {
    STATE.validaciones.push({
      Municipio: municipio, Contrato: contrato, Acta_No: actaNo, Periodo: periodo,
      Chequeo: c.chequeo, Resultado: c.resultado, Detalle: c.detalle,
    });
  });

  guardarEstado();
}

function rebuildResumen() {
  const data = {};
  const periodosPresentes = new Set();
  STATE.detalle.forEach((row) => {
    const key = row.Municipio + "||" + row.Contrato;
    if (!data[key]) data[key] = { municipio: row.Municipio, contrato: row.Contrato, programas: {} };
    if (!data[key].programas[row.Programa]) data[key].programas[row.Programa] = {};
    data[key].programas[row.Programa][row.Periodo] = row.Pct_Cumplimiento;
    periodosPresentes.add(row.Periodo);
  });

  const periodosOrdenados = PERIODO_ORDEN.filter((p) => periodosPresentes.has(p)).concat(
    [...periodosPresentes].filter((p) => !PERIODO_ORDEN.includes(p)).sort()
  );

  const rows = [];
  Object.values(data)
    .sort((a, b) => (a.municipio + a.contrato).localeCompare(b.municipio + b.contrato))
    .forEach((entry) => {
      Object.keys(entry.programas).sort(compararProgramas).forEach((programa) => {
        const periodos = entry.programas[programa];
        const row = { Municipio: entry.municipio, Contrato: entry.contrato, Programa: programa };
        const valores = [];
        periodosOrdenados.forEach((per) => {
          const pct = periodos[per];
          row[per] = pct === undefined ? null : pct;
          if (pct !== undefined && pct !== null) valores.push(pct);
        });
        row.Promedio = valores.length ? valores.reduce((a, b) => a + b, 0) / valores.length : null;
        rows.push(row);
      });
    });

  return { rows, periodos: periodosOrdenados };
}

// Consolidado de cumplimiento POR PRESTADOR/CONTRATO: una sola fila por cada
// (Municipio, Contrato) con el promedio de TODA la ruta (todas las
// actividades juntas), por periodo y general, más Cumple/No cumple (>=80%).
// Pensado para ver de un vistazo decenas de prestadores con varios
// contratos cada uno, sin desglosar actividad por actividad.
function rebuildConsolidadoPrestadores(actividadesSel) {
  const grupos = {};
  const periodosPresentes = new Set();
  STATE.detalle.forEach((row) => {
    if (actividadesSel && !actividadesSel.has(row.Programa)) return;
    periodosPresentes.add(row.Periodo);
    const key = row.Municipio + "||" + row.Contrato;
    if (!grupos[key]) {
      grupos[key] = { municipio: row.Municipio, contrato: row.Contrato, empresas: new Set(), periodos: {} };
    }
    const g = grupos[key];
    if (row.Empresa) g.empresas.add(row.Empresa);
    if (!g.periodos[row.Periodo]) g.periodos[row.Periodo] = [];
    if (row.Pct_Cumplimiento !== null && row.Pct_Cumplimiento !== undefined) {
      g.periodos[row.Periodo].push(row.Pct_Cumplimiento);
    }
  });

  const periodosOrdenados = PERIODO_ORDEN.filter((p) => periodosPresentes.has(p)).concat(
    [...periodosPresentes].filter((p) => !PERIODO_ORDEN.includes(p)).sort()
  );

  const rows = Object.values(grupos)
    .sort((a, b) => (a.municipio + a.contrato).localeCompare(b.municipio + b.contrato))
    .map((g) => {
      const row = {
        Municipio: g.municipio,
        Contrato: g.contrato,
        Prestador: [...g.empresas].sort().join(", "),
      };
      const promediosPeriodo = [];
      periodosOrdenados.forEach((per) => {
        const arr = g.periodos[per];
        const avg = arr && arr.length ? arr.reduce((a, b) => a + b, 0) / arr.length : null;
        row[per] = avg;
        if (avg !== null) promediosPeriodo.push(avg);
      });
      row.Promedio = promediosPeriodo.length ? promediosPeriodo.reduce((a, b) => a + b, 0) / promediosPeriodo.length : null;
      row.Estado = row.Promedio === null ? "" : row.Promedio >= 0.8 ? "Cumple" : "No cumple";
      return row;
    });

  return { rows, periodos: periodosOrdenados };
}

// Consolidado semaforizado POR PRESTADOR (Empresa) — junta todos los
// municipios/contratos de una misma empresa en una sola fila con un solo %
// (ej. "DUSAKAWI 90%", "PALAIMA 100%"), usando solo las actividades
// marcadas en el checklist.
function rebuildConsolidadoPorEmpresa(actividadesSel) {
  const porEmpresa = {};
  STATE.detalle.forEach((row) => {
    if (actividadesSel && !actividadesSel.has(row.Programa)) return;
    const key = row.Empresa || "(sin empresa)";
    if (!porEmpresa[key]) porEmpresa[key] = { empresa: key, valores: [], municipios: new Set() };
    if (row.Pct_Cumplimiento !== null && row.Pct_Cumplimiento !== undefined) {
      porEmpresa[key].valores.push(row.Pct_Cumplimiento);
    }
    porEmpresa[key].municipios.add(row.Municipio);
  });

  return Object.values(porEmpresa)
    .sort((a, b) => a.empresa.localeCompare(b.empresa))
    .map((e) => {
      const promedio = e.valores.length ? e.valores.reduce((a, b) => a + b, 0) / e.valores.length : null;
      return {
        Empresa: e.empresa,
        Municipios: [...e.municipios].sort().join(", "),
        NumMunicipios: e.municipios.size,
        Promedio: promedio,
        Estado: promedio === null ? "" : promedio >= 0.8 ? "Cumple" : "No cumple",
      };
    });
}

function estadisticaCumplimiento(rowsEmpresa) {
  const total = rowsEmpresa.filter((r) => r.Estado).length;
  const cumplen = rowsEmpresa.filter((r) => r.Estado === "Cumple").length;
  return { total, cumplen, noCumplen: total - cumplen };
}

// Agregado "caso extremo": DUSAKAWI es el único prestador — esto junta
// todos los municipios en una sola fila por actividad, promediando entre
// todos los que la reporten.
function rebuildResumenGeneral() {
  const porPrograma = {};
  const periodosPresentes = new Set();
  STATE.detalle.forEach((row) => {
    periodosPresentes.add(row.Periodo);
    if (!porPrograma[row.Programa]) {
      porPrograma[row.Programa] = { periodos: {}, municipios: new Set(), prestadores: new Set() };
    }
    const p = porPrograma[row.Programa];
    if (!p.periodos[row.Periodo]) p.periodos[row.Periodo] = [];
    if (row.Pct_Cumplimiento !== null && row.Pct_Cumplimiento !== undefined) {
      p.periodos[row.Periodo].push(row.Pct_Cumplimiento);
    }
    p.municipios.add(row.Municipio);
    if (row.Empresa) p.prestadores.add(row.Empresa);
  });

  const periodosOrdenados = PERIODO_ORDEN.filter((p) => periodosPresentes.has(p)).concat(
    [...periodosPresentes].filter((p) => !PERIODO_ORDEN.includes(p)).sort()
  );

  const rows = Object.keys(porPrograma).sort(compararProgramas).map((programa) => {
    const p = porPrograma[programa];
    const row = {
      Programa: programa,
      Municipios: [...p.municipios].sort().join(", "),
      NumMunicipios: p.municipios.size,
      Prestadores: [...p.prestadores].sort().join(", "),
    };
    const valoresPromedio = [];
    periodosOrdenados.forEach((per) => {
      const arr = p.periodos[per];
      const avg = arr && arr.length ? arr.reduce((a, b) => a + b, 0) / arr.length : null;
      row[per] = avg;
      if (avg !== null) valoresPromedio.push(avg);
    });
    row.Promedio = valoresPromedio.length ? valoresPromedio.reduce((a, b) => a + b, 0) / valoresPromedio.length : null;
    return row;
  });

  return { rows, periodos: periodosOrdenados };
}

// Etiqueta de columna para un periodo: "OTRO SI (ENE-FEB 2026)" en vez de
// solo el código, usando los meses/año reales que traiga el acta.
function mapaPeriodoInfo() {
  const map = {};
  STATE.detalle.forEach((row) => {
    if (!map[row.Periodo]) map[row.Periodo] = { meses: row.Meses, anio: row.Anio };
  });
  return map;
}
function etiquetaPeriodo(periodo, info) {
  if (!info) return periodo;
  const meses = info.meses ? String(info.meses).replace(/\s*,\s*/g, "-") : "";
  const partes = [meses, info.anio].filter(Boolean).join(" ");
  return partes ? `${periodo} (${partes})` : periodo;
}

// --------------------------------------------------------------------------
// Render
// --------------------------------------------------------------------------
function renderResultado(container, payload) {
  if (payload.error) {
    container.innerHTML = `<div class="msg error">${payload.error}</div>`;
    return;
  }
  const { acta, checks, periodo, periodoAuto } = payload;
  const nErr = checks.filter((c) => c.resultado === "ERROR").length;
  const periodoTxt = periodoAuto ? `${periodo} <span class="pill ok" style="margin-left:4px;">detectado automático</span>` : periodo;
  const resumenMsg =
    nErr === 0
      ? `<div class="msg ok">✔ ${acta.municipio} — Contrato ${acta.contrato} — Acta ${acta.acta_no} — periodo ${periodoTxt}: ${checks.length} chequeos, todos OK.</div>`
      : `<div class="msg error">⚠ ${acta.municipio} — Contrato ${acta.contrato} — Acta ${acta.acta_no} — periodo ${periodoTxt}: ${nErr} de ${checks.length} chequeos con ERROR.</div>`;

  const filas = checks
    .map(
      (c) => `
    <tr class="${c.resultado === "ERROR" ? "row-error" : ""}">
      <td>${c.chequeo}</td>
      <td><span class="pill ${c.resultado === "OK" ? "ok" : "error"}">${c.resultado}</span></td>
      <td>${c.detalle}</td>
    </tr>`
    )
    .join("");

  container.innerHTML =
    resumenMsg +
    `<div class="table-wrap" style="margin-top:12px;">
      <table>
        <thead><tr><th>Chequeo</th><th>Resultado</th><th>Detalle</th></tr></thead>
        <tbody>${filas}</tbody>
      </table>
    </div>`;
}

// Umbrales: >=90% excelente, >=80% cumple, <80% no cumple.
function fmtPct(v) {
  if (v === null || v === undefined || v === "") return '<span style="color:#c0c4cc;">—</span>';
  const pct = v * 100;
  const cls = pct >= 90 ? "pct-pill-ok" : pct >= 80 ? "pct-pill-mid" : "pct-pill-bad";
  return `<span class="pct-pill ${cls}">${pct.toFixed(1)}%</span>`;
}

// headerCols/filas: la primera columna es texto (Programa); el resto son %
// (periodos + Promedio al final, que se resalta).
function tablaResumenHtml(headerCols, filas) {
  const thead = "<tr>" + headerCols.map((c) => `<th>${c}</th>`).join("") + "</tr>";
  const body = filas
    .map((celdas) => {
      const tds = celdas
        .map((v, i) => {
          if (i === 0) return `<td>${v}</td>`;
          const esPromedio = i === celdas.length - 1;
          return `<td class="pct-cell"${esPromedio ? ' style="background:#f7f9ff;"' : ""}>${fmtPct(v)}</td>`;
        })
        .join("");
      return `<tr>${tds}</tr>`;
    })
    .join("");
  return `<div class="table-wrap"><table><thead>${thead}</thead><tbody>${body}</tbody></table></div>`;
}

// Tabla del consolidado general: además del % por periodo, muestra los
// montos totales (Exigido/Reconocido/Descuento) y qué prestadores reportan
// cada actividad, para que no sea solo un número suelto.
function tablaResumenGeneralHtml(rows, periodos, infoPeriodos) {
  const headerCols = ["Programa / Actividad", ...periodos.map((p) => etiquetaPeriodo(p, infoPeriodos[p])), "Promedio", "Municipios", "Prestador (Empresa)"];
  const thead = "<tr>" + headerCols.map((c) => `<th>${c}</th>`).join("") + "</tr>";
  const body = rows
    .map((r) => {
      let tds = `<td>${r.Programa}</td>`;
      periodos.forEach((p) => { tds += `<td class="pct-cell">${fmtPct(r[p])}</td>`; });
      tds += `<td class="pct-cell" style="background:#f7f9ff;">${fmtPct(r.Promedio)}</td>`;
      tds += `<td class="hint" style="white-space:normal; max-width:220px;">${r.Municipios} <span class="pill ok">${r.NumMunicipios}</span></td>`;
      tds += `<td class="hint" style="white-space:normal; max-width:220px;">${r.Prestadores || "-"}</td>`;
      return `<tr>${tds}</tr>`;
    })
    .join("");
  return `<div class="table-wrap"><table><thead>${thead}</thead><tbody>${body}</tbody></table></div>`;
}

// Tabla del consolidado por prestador/contrato: promedio de TODA la ruta en
// un solo número por periodo, más Cumple/No cumple.
function tablaConsolidadoPrestadoresHtml(rows, periodos, infoPeriodos) {
  const headerCols = ["Municipio", "Nº Contrato", "Prestador (Empresa)", ...periodos.map((p) => etiquetaPeriodo(p, infoPeriodos[p])), "Promedio General", "Estado"];
  const thead = "<tr>" + headerCols.map((c) => `<th>${c}</th>`).join("") + "</tr>";
  const body = rows
    .map((r) => {
      let tds = `<td>${r.Municipio}</td><td>${r.Contrato}</td><td>${r.Prestador || "-"}</td>`;
      periodos.forEach((p) => { tds += `<td class="pct-cell">${fmtPct(r[p])}</td>`; });
      tds += `<td class="pct-cell" style="background:#f7f9ff;"><strong>${fmtPct(r.Promedio)}</strong></td>`;
      const estadoCls = r.Estado === "Cumple" ? "ok" : r.Estado === "No cumple" ? "error" : "";
      tds += `<td>${r.Estado ? `<span class="pill ${estadoCls}">${r.Estado}</span>` : "-"}</td>`;
      return `<tr>${tds}</tr>`;
    })
    .join("");
  return `<div class="table-wrap"><table><thead>${thead}</thead><tbody>${body}</tbody></table></div>`;
}

function tablaPorEmpresaHtml(rows) {
  const body = rows
    .map((r) => {
      const estadoCls = r.Estado === "Cumple" ? "ok" : r.Estado === "No cumple" ? "error" : "";
      return `<tr>
        <td>${r.Empresa}</td>
        <td class="hint" style="white-space:normal; max-width:260px;">${r.Municipios} <span class="pill ok">${r.NumMunicipios}</span></td>
        <td class="pct-cell"><strong>${fmtPct(r.Promedio)}</strong></td>
        <td>${r.Estado ? `<span class="pill ${estadoCls}">${r.Estado}</span>` : "-"}</td>
      </tr>`;
    })
    .join("");
  return `<div class="table-wrap"><table>
    <thead><tr><th>Prestador (Empresa)</th><th>Municipios</th><th>Promedio</th><th>Estado</th></tr></thead>
    <tbody>${body}</tbody></table></div>`;
}

function tarjetasEstadisticaHtml(stat) {
  return `<div class="stat-cards">
    <div class="stat-card"><div class="stat-num">${stat.total}</div><div class="stat-label">Prestadores cargados</div></div>
    <div class="stat-card stat-ok"><div class="stat-num">${stat.cumplen}</div><div class="stat-label">✅ Cumplen (≥80%)</div></div>
    <div class="stat-card stat-bad"><div class="stat-num">${stat.noCumplen}</div><div class="stat-label">⚠️ No cumplen</div></div>
  </div>`;
}

// Checklist de actividades que entran al consolidado por prestador/empresa.
function checklistActividadesHtml() {
  const items = PROGRAMAS_TIPICOS.map((p, i) => `
    <label class="check-item">
      <input type="checkbox" class="chk-actividad" data-idx="${i}" ${actividadesSeleccionadas.has(p) ? "checked" : ""}>
      <span>${p}</span>
    </label>`).join("");
  return `<div class="resumen-grupo">
    <h3>🗂️ Actividades incluidas en el consolidado por prestador</h3>
    <p class="hint">Marca qué actividades cuentan para el % general de cada prestador (abajo). Por defecto viene la ruta de promoción y mantenimiento; PAI, atención posparto, demanda inducida y ruta cardiovascular quedan fuera salvo que las marques.</p>
    <div class="checklist-actividades">${items}</div>
  </div>`;
}

function escapeXml(s) {
  return String(s).replace(/[<>&"]/g, (c) => ({ "<": "&lt;", ">": "&gt;", "&": "&amp;", '"': "&quot;" }[c]));
}

// Rasteriza un SVG (string) a bytes PNG, dibujándolo sobre fondo blanco
// (las celdas de Excel no manejan bien la transparencia).
function svgAPng(svgString, width, height, escala = 2) {
  return new Promise((resolve, reject) => {
    const blob = new Blob([svgString], { type: "image/svg+xml;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const img = new Image();
    img.onload = () => {
      const canvas = document.createElement("canvas");
      canvas.width = width * escala;
      canvas.height = height * escala;
      const ctx = canvas.getContext("2d");
      ctx.fillStyle = "#ffffff";
      ctx.fillRect(0, 0, canvas.width, canvas.height);
      ctx.drawImage(img, 0, 0, canvas.width, canvas.height);
      URL.revokeObjectURL(url);
      canvas.toBlob((blob2) => {
        if (!blob2) return reject(new Error("No se pudo generar el PNG de la gráfica"));
        blob2.arrayBuffer().then((buf) => resolve(new Uint8Array(buf)));
      }, "image/png");
    };
    img.onerror = () => { URL.revokeObjectURL(url); reject(new Error("No se pudo rasterizar la gráfica (SVG inválido)")); };
    img.src = url;
  });
}

// Inserta una imagen PNG como dibujo flotante en una hoja del libro, ya
// convertido a zip (JSZip). Ubica el borde superior izquierdo en la fila
// `filaAncla` (0-based) de la hoja llamada `sheetName`.
async function insertarImagenEnHoja(zip, sheetName, pngBytes, pxWidth, pxHeight, filaAncla, indice = 1) {
  const parser = new DOMParser();

  const wbXml = await zip.file("xl/workbook.xml").async("string");
  const wbDoc = parser.parseFromString(wbXml, "application/xml");
  const sheetEl = [...wbDoc.getElementsByTagName("sheet")].find((s) => s.getAttribute("name") === sheetName);
  if (!sheetEl) return false;
  const rId = sheetEl.getAttribute("r:id");

  const wbRelsXml = await zip.file("xl/_rels/workbook.xml.rels").async("string");
  const wbRelsDoc = parser.parseFromString(wbRelsXml, "application/xml");
  const rel = [...wbRelsDoc.getElementsByTagName("Relationship")].find((r) => r.getAttribute("Id") === rId);
  if (!rel) return false;
  const sheetFileName = rel.getAttribute("Target").split("/").pop();
  const sheetXmlPath = `xl/worksheets/${sheetFileName}`;

  zip.file(`xl/media/image${indice}.png`, pngBytes);

  const emu = 9525; // EMU por pixel
  const cx = Math.round(pxWidth * emu);
  const cy = Math.round(pxHeight * emu);
  zip.file(
    `xl/drawings/drawing${indice}.xml`,
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
<xdr:oneCellAnchor>
<xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${filaAncla}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
<xdr:ext cx="${cx}" cy="${cy}"/>
<xdr:pic>
<xdr:nvPicPr><xdr:cNvPr id="${indice}" name="GraficaCumplimiento${indice}"/><xdr:cNvPicPr/></xdr:nvPicPr>
<xdr:blipFill><a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rId1"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>
<xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr>
</xdr:pic>
<xdr:clientData/>
</xdr:oneCellAnchor>
</xdr:wsDr>`
  );
  zip.file(
    `xl/drawings/_rels/drawing${indice}.xml.rels`,
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image${indice}.png"/>
</Relationships>`
  );
  zip.file(
    `xl/worksheets/_rels/${sheetFileName}.rels`,
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing${indice}.xml"/>
</Relationships>`
  );

  let sheetXml = await zip.file(sheetXmlPath).async("string");
  sheetXml = sheetXml.replace("</worksheet>", '<drawing r:id="rId1"/></worksheet>');
  zip.file(sheetXmlPath, sheetXml);

  let ctXml = await zip.file("[Content_Types].xml").async("string");
  if (!ctXml.includes('Extension="png"')) {
    ctXml = ctXml.replace("</Types>", '<Default Extension="png" ContentType="image/png"/></Types>');
  }
  const drawingPart = `/xl/drawings/drawing${indice}.xml`;
  if (!ctXml.includes(drawingPart)) {
    ctXml = ctXml.replace("</Types>", `<Override PartName="${drawingPart}" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/></Types>`);
  }
  zip.file("[Content_Types].xml", ctXml);
  return true;
}

// Grafica de barras (SVG puro, sin librerias) con el % Promedio de cada
// actividad. Nunca puede pasar de 100% porque es un promedio de fracciones
// <=1 (si alguna actividad tuviera Reconocido > Exigido, la validacion ya lo
// marca como ERROR antes de llegar aqui).
function graficoBarrasSVG(rows) {
  const n = rows.length;
  if (!n) return "";
  const w = Math.max(640, n * 92);
  const h = 340;
  const padTop = 28, padBottom = 165, padLeft = 90, padRight = 20;
  const chartH = h - padTop - padBottom;
  const gap = (w - padLeft - padRight) / n;
  const barW = Math.min(48, gap * 0.55);

  let bars = "";
  rows.forEach((r, i) => {
    const pct = r.Promedio === null || r.Promedio === undefined ? 0 : Math.min(r.Promedio, 1) * 100;
    const barH = (pct / 100) * chartH;
    const x = padLeft + i * gap + (gap - barW) / 2;
    const y = padTop + (chartH - barH);
    const color = pct >= 90 ? "#16924f" : pct >= 80 ? "#b7791f" : "#d0342c";
    const cx = x + barW / 2;
    const labelY = padTop + chartH + 16;
    const nombreCorto = r.Programa.length > 30 ? r.Programa.slice(0, 28) + "…" : r.Programa;
    bars += `
      <rect x="${x.toFixed(1)}" y="${y.toFixed(1)}" width="${barW.toFixed(1)}" height="${Math.max(barH, 1).toFixed(1)}" fill="${color}" rx="4"/>
      <text x="${cx.toFixed(1)}" y="${(y - 7).toFixed(1)}" text-anchor="middle" font-size="12" font-weight="700" fill="${color}">${r.Promedio === null ? "-" : pct.toFixed(0) + "%"}</text>
      <text x="${cx.toFixed(1)}" y="${labelY}" font-size="10" fill="#475067" text-anchor="end" transform="rotate(-40 ${cx.toFixed(1)} ${labelY})">${escapeXml(nombreCorto)}</text>`;
  });

  const y100 = padTop;
  return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${w} ${h}" width="${w}" height="${h}" style="width:100%; max-width:${w}px; height:auto; display:block; margin:0 auto;">
    <line x1="${padLeft}" y1="${y100}" x2="${w - padRight}" y2="${y100}" stroke="#c7ccd6" stroke-dasharray="4 3"/>
    <text x="${padLeft}" y="${y100 - 6}" font-size="10" fill="#98a2b3">100%</text>
    ${bars}
  </svg>`;
}

function cargarResumen() {
  const container = document.getElementById("resumen-contenido");
  const { rows, periodos } = rebuildResumen();
  if (rows.length === 0) {
    container.innerHTML = '<div class="empty">Aún no hay actas consolidadas.</div>';
    return;
  }

  // Vista principal: DUSAKAWI es el único prestador — un solo consolidado
  // sin importar el municipio, una fila por actividad.
  const { rows: rowsGeneral, periodos: periodosGeneral } = rebuildResumenGeneral();
  const infoPeriodos = mapaPeriodoInfo();

  let html = `<div class="resumen-grupo resumen-general">
    <h3>🌐 Consolidado general — DUSAKAWI IPSI (todos los municipios)</h3>
    <p class="hint">Una fila por actividad — el % es el promedio entre todos los municipios que la reportan (nunca supera 100%).</p>
    ${graficoBarrasSVG(rowsGeneral)}
    <p class="leyenda-umbrales">
      <span class="pct-pill pct-pill-ok">≥ 90%</span> Excelente &nbsp;
      <span class="pct-pill pct-pill-mid">≥ 80%</span> Cumple &nbsp;
      <span class="pct-pill pct-pill-bad">&lt; 80%</span> No cumple
    </p>
    ${tablaResumenGeneralHtml(rowsGeneral, periodosGeneral, infoPeriodos)}
  </div>`;

  // Checklist de actividades para los dos consolidados por prestador.
  html += checklistActividadesHtml();

  // Consolidado semaforizado POR EMPRESA (ej. "DUSAKAWI 90%", "PALAIMA 100%").
  const rowsEmpresa = rebuildConsolidadoPorEmpresa(actividadesSeleccionadas);
  const statEmpresa = estadisticaCumplimiento(rowsEmpresa);
  const filasParaGrafico = rowsEmpresa.map((r) => ({ Programa: r.Empresa, Promedio: r.Promedio }));
  html += `<div class="resumen-grupo resumen-general">
    <h3>🚦 Consolidado por Prestador — Semaforizado</h3>
    <p class="hint">Cada prestador (empresa), con el promedio de las actividades marcadas arriba en todos sus municipios/contratos.</p>
    ${tarjetasEstadisticaHtml(statEmpresa)}
    ${graficoBarrasSVG(filasParaGrafico)}
    <p class="leyenda-umbrales">
      <span class="pct-pill pct-pill-ok">≥ 90%</span> Excelente &nbsp;
      <span class="pct-pill pct-pill-mid">≥ 80%</span> Cumple &nbsp;
      <span class="pct-pill pct-pill-bad">&lt; 80%</span> No cumple
    </p>
    ${tablaPorEmpresaHtml(rowsEmpresa)}
  </div>`;

  // Consolidado por prestador/contrato: toda la ruta en un solo % por cada
  // uno (pensado para ver de un vistazo decenas de prestadores).
  const { rows: rowsPrestadores, periodos: periodosPrestadores } = rebuildConsolidadoPrestadores(actividadesSeleccionadas);
  html += `<div class="resumen-grupo">
    <h3>📋 Consolidado de cumplimiento por prestador y contrato</h3>
    <p class="hint">Igual que arriba, pero desglosado por cada municipio/contrato en vez de agrupado por empresa.</p>
    ${tablaConsolidadoPrestadoresHtml(rowsPrestadores, periodosPrestadores, infoPeriodos)}
  </div>`;

  // Detalle: desglose por municipio, para quien necesite ver el origen.
  const grupos = new Map();
  rows.forEach((r) => {
    const key = r.Municipio + "||" + r.Contrato;
    if (!grupos.has(key)) grupos.set(key, { municipio: r.Municipio, contrato: r.Contrato, filas: [] });
    grupos.get(key).filas.push(r);
  });

  html += `<div class="resumen-detalle-toggle">
    <button type="button" class="secondary" id="btn-toggle-detalle-prestador">▸ Ver desglose por municipio (${grupos.size})</button>
  </div>
  <div id="detalle-por-prestador" style="display:none; margin-top:16px;">`;
  grupos.forEach((g) => {
    const headerCols = ["Programa", ...periodos.map((p) => etiquetaPeriodo(p, infoPeriodos[p])), "Promedio"];
    const filas = g.filas.map((r) => [r.Programa, ...periodos.map((p) => r[p]), r.Promedio]);
    html += `<div class="resumen-grupo">
      <h3>🏥 ${g.municipio} <span class="hint" style="display:inline; margin:0;">— Contrato ${g.contrato}</span></h3>
      ${tablaResumenHtml(headerCols, filas)}
    </div>`;
  });
  html += `</div>`;

  container.innerHTML = html;
  const btnToggle = document.getElementById("btn-toggle-detalle-prestador");
  btnToggle.addEventListener("click", () => {
    const det = document.getElementById("detalle-por-prestador");
    const abierto = det.style.display !== "none";
    det.style.display = abierto ? "none" : "block";
    btnToggle.textContent = `${abierto ? "▸" : "▾"} Ver desglose por municipio (${grupos.size})`;
  });

  container.querySelectorAll(".chk-actividad").forEach((chk) => {
    chk.addEventListener("change", () => {
      const programa = PROGRAMAS_TIPICOS[parseInt(chk.dataset.idx, 10)];
      if (chk.checked) actividadesSeleccionadas.add(programa);
      else actividadesSeleccionadas.delete(programa);
      cargarResumen();
    });
  });
}

function cargarValidaciones() {
  const container = document.getElementById("validaciones-contenido");
  if (STATE.validaciones.length === 0) {
    container.innerHTML = '<div class="empty">Aún no hay validaciones registradas.</div>';
    return;
  }
  const rows = STATE.validaciones
    .map(
      (r) => `
    <tr class="${r.Resultado === "ERROR" ? "row-error" : ""}">
      <td>${r.Municipio}</td><td>${r.Contrato}</td><td>${r.Acta_No}</td><td>${r.Periodo}</td>
      <td>${r.Chequeo}</td>
      <td><span class="pill ${r.Resultado === "OK" ? "ok" : "error"}">${r.Resultado}</span></td>
      <td>${r.Detalle}</td>
    </tr>`
    )
    .join("");
  container.innerHTML = `<table>
    <thead><tr><th>Municipio</th><th>Contrato</th><th>Acta</th><th>Periodo</th><th>Chequeo</th><th>Resultado</th><th>Detalle</th></tr></thead>
    <tbody>${rows}</tbody></table>`;
}

// --------------------------------------------------------------------------
// Exportar / Importar el Excel maestro
// --------------------------------------------------------------------------
// Cada columna define su clave interna (usada en STATE) y la etiqueta bonita
// que se muestra en el Excel. La importación se hace por POSICIÓN (no por
// texto de encabezado), así el archivo puede tener etiquetas legibles sin
// romper la relectura.
const DETALLE_COLUMNAS = [
  { key: "Municipio", label: "Municipio" },
  { key: "Contrato", label: "Nº Contrato" },
  { key: "Acta_No", label: "Nº Acta" },
  { key: "Periodo", label: "Periodo" },
  { key: "Meses", label: "Meses Evaluados" },
  { key: "Anio", label: "Año" },
  { key: "Empresa", label: "Empresa" },
  { key: "Nit", label: "NIT" },
  { key: "Regimen", label: "Régimen" },
  { key: "Programa", label: "Programa / Actividad" },
  { key: "Vr_Exigido", label: "Vr Exigido" },
  { key: "Vr_Reconocido", label: "Vr Reconocido" },
  { key: "Descuento", label: "Descuento" },
  { key: "Pct_Cumplimiento", label: "% Cumplimiento" },
  { key: "Archivo_Origen", label: "Archivo Origen" },
  { key: "Fecha_Carga", label: "Fecha de Carga" },
];
const VALIDACIONES_COLUMNAS = [
  { key: "Municipio", label: "Municipio" },
  { key: "Contrato", label: "Nº Contrato" },
  { key: "Acta_No", label: "Nº Acta" },
  { key: "Periodo", label: "Periodo" },
  { key: "Chequeo", label: "Chequeo" },
  { key: "Resultado", label: "Resultado" },
  { key: "Detalle", label: "Detalle" },
];

function aplicarFormatoNumero(ws, colIdx, numRows, formato) {
  for (let r = 1; r < numRows; r++) {
    const cell = ws[XLSX.utils.encode_cell({ r, c: colIdx })];
    if (cell && typeof cell.v === "number") cell.z = formato;
  }
}

// Ajusta el ancho de cada columna al contenido más largo (con topes
// razonables), ya que la librería gratuita no permite fijar colores/negrita
// de celda — esto es lo que realmente evita que los números se corten o se
// encimen con el encabezado.
function autoAjustarColumnas(ws, aoa) {
  const anchos = aoa[0].map((_, colIdx) => {
    let max = 8;
    for (const fila of aoa) {
      const v = fila[colIdx];
      const largo = v === null || v === undefined ? 0 : String(v).length;
      if (largo > max) max = largo;
    }
    return { wch: Math.min(Math.max(max + 2, 10), 45) };
  });
  ws["!cols"] = anchos;
}

function hojaConEstilo(aoa, opts = {}) {
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  autoAjustarColumnas(ws, aoa);
  (opts.columnasMoneda || []).forEach((c) => aplicarFormatoNumero(ws, c, aoa.length, "#,##0.00"));
  (opts.columnasPorcentaje || []).forEach((c) => aplicarFormatoNumero(ws, c, aoa.length, "0.0%"));
  ws["!autofilter"] = { ref: XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: aoa.length - 1, c: aoa[0].length - 1 } }) };
  return ws;
}

async function descargarMaestro() {
  const wb = XLSX.utils.book_new();

  const detalleHeaders = DETALLE_COLUMNAS.map((c) => c.label);
  const detalleAoa = [detalleHeaders, ...STATE.detalle.map((r) => DETALLE_COLUMNAS.map((c) => r[c.key] ?? null))];
  const wsDetalle = hojaConEstilo(detalleAoa, { columnasMoneda: [10, 11, 12], columnasPorcentaje: [13] });
  XLSX.utils.book_append_sheet(wb, wsDetalle, "Detalle");

  const valHeaders = VALIDACIONES_COLUMNAS.map((c) => c.label);
  const valAoa = [valHeaders, ...STATE.validaciones.map((r) => VALIDACIONES_COLUMNAS.map((c) => r[c.key] ?? null))];
  const wsVal = hojaConEstilo(valAoa, {});
  XLSX.utils.book_append_sheet(wb, wsVal, "Validaciones");

  // Resumen_General primero: es el consolidado principal — DUSAKAWI es el
  // único prestador, así que es una fila por actividad sin importar el
  // municipio, con las fechas reales de cada periodo en el encabezado.
  const { rows: rowsGeneral, periodos: periodosGeneral } = rebuildResumenGeneral();
  const infoPeriodosXlsx = mapaPeriodoInfo();
  const resGeneralHeaders = ["Programa / Actividad", ...periodosGeneral.map((p) => etiquetaPeriodo(p, infoPeriodosXlsx[p])), "Promedio", "Nº Municipios", "Municipios", "Prestador (Empresa)"];
  const resGeneralAoa = [
    resGeneralHeaders,
    ...rowsGeneral.map((r) => [
      r.Programa, ...periodosGeneral.map((p) => r[p] ?? null), r.Promedio ?? null, r.NumMunicipios, r.Municipios, r.Prestadores,
    ]),
  ];
  const wsResGeneral = hojaConEstilo(resGeneralAoa, {
    columnasPorcentaje: periodosGeneral.map((_, i) => 1 + i).concat([1 + periodosGeneral.length]),
  });
  XLSX.utils.book_append_sheet(wb, wsResGeneral, "Resumen_General");

  // Consolidado semaforizado POR EMPRESA (ej. "DUSAKAWI 90%"), solo con las
  // actividades marcadas en el checklist de la app.
  const rowsEmpresaXlsx = rebuildConsolidadoPorEmpresa(actividadesSeleccionadas);
  const empHeaders = ["Prestador (Empresa)", "Municipios", "Nº Municipios", "Promedio General", "Estado"];
  const empAoa = [
    empHeaders,
    ...rowsEmpresaXlsx.map((r) => [r.Empresa, r.Municipios, r.NumMunicipios, r.Promedio ?? null, r.Estado]),
  ];
  const wsEmp = hojaConEstilo(empAoa, { columnasPorcentaje: [3] });
  XLSX.utils.book_append_sheet(wb, wsEmp, "Consolidado_Empresas");

  // Consolidado por prestador/contrato: toda la ruta en un solo % por cada
  // uno, con Cumple/No cumple — pensado para decenas de prestadores.
  const { rows: rowsPrestadores, periodos: periodosPrestadores } = rebuildConsolidadoPrestadores(actividadesSeleccionadas);
  const presHeaders = ["Municipio", "Nº Contrato", "Prestador (Empresa)", ...periodosPrestadores.map((p) => etiquetaPeriodo(p, infoPeriodosXlsx[p])), "Promedio General", "Estado"];
  const presAoa = [
    presHeaders,
    ...rowsPrestadores.map((r) => [
      r.Municipio, r.Contrato, r.Prestador, ...periodosPrestadores.map((p) => r[p] ?? null), r.Promedio ?? null, r.Estado,
    ]),
  ];
  const wsPres = hojaConEstilo(presAoa, {
    columnasPorcentaje: periodosPrestadores.map((_, i) => 3 + i).concat([3 + periodosPrestadores.length]),
  });
  XLSX.utils.book_append_sheet(wb, wsPres, "Consolidado_Prestadores");

  const { rows, periodos } = rebuildResumen();
  const resHeaders = ["Municipio", "Nº Contrato", "Programa / Actividad", ...periodos.map((p) => etiquetaPeriodo(p, infoPeriodosXlsx[p])), "Promedio"];
  const resAoa = [
    resHeaders,
    ...rows.map((r) => [r.Municipio, r.Contrato, r.Programa, ...periodos.map((p) => r[p] ?? null), r.Promedio ?? null]),
  ];
  const wsRes = hojaConEstilo(resAoa, { columnasPorcentaje: periodos.map((_, i) => 3 + i).concat([3 + periodos.length]) });
  XLSX.utils.book_append_sheet(wb, wsRes, "Resumen_Detalle_Prestador");

  const nombreArchivo = "CONSOLIDADO_ACTAS_PYM_DUSAKAWI.xlsx";
  const bytesBase = XLSX.write(wb, { type: "array", bookType: "xlsx" });

  // Intenta insertar la gráfica del consolidado general como imagen en la
  // hoja Resumen_General. Si algo falla (SVG raro, navegador sin soporte),
  // se descarga igual el Excel sin la imagen en vez de fallar por completo.
  try {
    if (typeof JSZip === "undefined") throw new Error("JSZip no cargó");
    const svgChart = graficoBarrasSVG(rowsGeneral);
    const dims = svgChart.match(/viewBox="0 0 ([\d.]+) ([\d.]+)"/);
    const chartW = dims ? parseFloat(dims[1]) : 700;
    const chartH = dims ? parseFloat(dims[2]) : 340;
    const pngBytes = await svgAPng(svgChart, chartW, chartH);
    const zip = await JSZip.loadAsync(bytesBase);
    const filaAncla = rowsGeneral.length + 2;
    await insertarImagenEnHoja(zip, "Resumen_General", pngBytes, chartW, chartH, filaAncla, 1);

    // Gráfica semaforizada por empresa, en su propia hoja.
    const svgChartEmp = graficoBarrasSVG(rowsEmpresaXlsx.map((r) => ({ Programa: r.Empresa, Promedio: r.Promedio })));
    const dimsEmp = svgChartEmp.match(/viewBox="0 0 ([\d.]+) ([\d.]+)"/);
    const chartWEmp = dimsEmp ? parseFloat(dimsEmp[1]) : 700;
    const chartHEmp = dimsEmp ? parseFloat(dimsEmp[2]) : 340;
    const pngBytesEmp = await svgAPng(svgChartEmp, chartWEmp, chartHEmp);
    const filaAnclaEmp = rowsEmpresaXlsx.length + 2;
    await insertarImagenEnHoja(zip, "Consolidado_Empresas", pngBytesEmp, chartWEmp, chartHEmp, filaAnclaEmp, 2);

    const finalBlob = await zip.generateAsync({ type: "blob" });
    const url = URL.createObjectURL(finalBlob);
    const a = document.createElement("a");
    a.href = url;
    a.download = nombreArchivo;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
  } catch (e) {
    console.warn("No se pudo insertar la gráfica en el Excel, se descarga sin ella:", e);
    XLSX.writeFile(wb, nombreArchivo);
  }
}

function filasDesdeHoja(ws, columnas) {
  if (!ws) return [];
  const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, raw: true });
  return aoa.slice(1).filter((fila) => fila.some((v) => v !== null && v !== "")).map((fila) => {
    const obj = {};
    columnas.forEach((c, i) => { obj[c.key] = fila[i] ?? null; });
    return obj;
  });
}

function cargarMaestroDesdeArchivo(workbook) {
  STATE.detalle = filasDesdeHoja(workbook.Sheets["Detalle"], DETALLE_COLUMNAS);
  STATE.validaciones = filasDesdeHoja(workbook.Sheets["Validaciones"], VALIDACIONES_COLUMNAS);
  guardarEstado();
}

// --------------------------------------------------------------------------
// UI wiring
// --------------------------------------------------------------------------

function agregarFilaProgramaEn(tbody, valores) {
  const tr = document.createElement("tr");
  const v = valores || { programa: "", vr_exigido: "", vr_reconocido: "", descuento: "" };
  tr.innerHTML = `
    <td><input type="text" class="p-nombre" value="${v.programa}" placeholder="Nombre de la actividad"></td>
    <td><input type="number" step="0.01" class="p-exigido" value="${v.vr_exigido}"></td>
    <td><input type="number" step="0.01" class="p-reconocido" value="${v.vr_reconocido}"></td>
    <td><input type="number" step="0.01" class="p-descuento" value="${v.descuento}"></td>
    <td><button type="button" class="danger" onclick="this.closest('tr').remove()">✕</button></td>
  `;
  tbody.appendChild(tr);
}

// --------------------------------------------------------------------------
// Tarjetas dinámicas de PDF (previsualización + transcripción por archivo)
// --------------------------------------------------------------------------
function mostrarPDFEmbebido(file, container) {
  try {
    const url = URL.createObjectURL(file);
    container.innerHTML = "";
    const link = document.createElement("a");
    link.href = url;
    link.target = "_blank";
    link.rel = "noopener";
    link.className = "pdf-open-link";
    link.textContent = "↗ Abrir PDF en pestaña nueva (si no se ve abajo)";
    const embed = document.createElement("iframe");
    embed.className = "pdf-embed";
    embed.src = url;
    embed.title = file.name;
    container.appendChild(link);
    container.appendChild(embed);
  } catch (e) {
    container.innerHTML = `<div class="msg error">No se pudo previsualizar el PDF (${e.message}). Puedes transcribir igual usando el formulario de abajo.</div>`;
  }
}

function crearTarjetaPDF(file) {
  const ruta = rutaArchivo(file);
  const municipioDetectado = detectarMunicipioDesdeTexto(ruta) || "";
  const contratoDetectado = detectarContratoDesdeTexto(ruta) || "";
  const periodoDetectado = detectarPeriodoDesdeTexto(ruta);

  const card = document.createElement("div");
  card.className = "pdf-card";
  card.innerHTML = `
    <h3 style="margin-top:0;">${file.name}</h3>
    ${ruta !== file.name ? `<div class="hint" style="margin-top:-6px;">📁 ${ruta}</div>` : ""}
    <div class="pdf-pages"><div class="empty">Cargando previsualización...</div></div>
    <form class="pdf-form">
      <div class="grid2">
        <div><label>Municipio${municipioDetectado ? ' <span class="pill ok">auto</span>' : ""}</label><input type="text" name="municipio" value="${municipioDetectado}" required></div>
        <div><label>Nº Contrato${contratoDetectado ? ' <span class="pill ok">auto</span>' : ""}</label><input type="text" name="contrato" value="${contratoDetectado}" required placeholder="Ej. 20013-061-PMT"></div>
        <div><label>Nº Acta</label><input type="text" name="acta_no" required placeholder="Ej. 20013-061-PMT-5"></div>
        <div><label>Periodo${periodoDetectado ? ' <span class="pill ok">auto</span>' : ""}</label><select name="periodo" class="periodo-select" required></select></div>
        <div><label>Meses evaluados</label><input type="text" name="meses" placeholder="Ej. ENE,FEB"></div>
        <div><label>Año</label><input type="text" name="anio" value="${periodoDetectado ? (extraerAnio(ruta) || new Date().getFullYear()) : new Date().getFullYear()}"></div>
        <div><label>Empresa (prestador)</label><input type="text" name="empresa" placeholder="Ej. DUSAKAWI IPSI, EZEQ SALUD IPSI..."></div>
        <div><label>NIT</label><input type="text" name="nit"></div>
        <div><label>Régimen</label><input type="text" name="regimen" value="SUBSIDIADO"></div>
      </div>

      <h4 style="margin-top:18px;">Tabla de programas</h4>
      <div class="table-wrap">
        <table class="tabla-programas-pdf">
          <thead><tr><th>Programa</th><th>Vr Exigido</th><th>Vr Reconocido</th><th>Descuento</th><th></th></tr></thead>
          <tbody></tbody>
        </table>
      </div>
      <button type="button" class="secondary btn-agregar-fila-pdf">+ Agregar actividad</button>

      <h4 style="margin-top:18px;">Total ejecución (según el acta)</h4>
      <div class="grid2">
        <div><label>Total Vr Exigido</label><input type="number" step="0.01" name="total_exigido"></div>
        <div><label>Total Vr Reconocido</label><input type="number" step="0.01" name="total_reconocido"></div>
        <div><label>Total Descuento</label><input type="number" step="0.01" name="total_descuento"></div>
      </div>

      <button class="primary" type="submit">Validar y consolidar</button>
      <div class="resultado-pdf"></div>
    </form>
  `;

  const sel = card.querySelector(".periodo-select");
  sel.innerHTML = Object.entries(PERIODO_MESES).map(([k, v]) => `<option value="${k}">${k} (${v})</option>`).join("");
  if (periodoDetectado) sel.value = periodoDetectado;

  const tbody = card.querySelector(".tabla-programas-pdf tbody");
  PROGRAMAS_TIPICOS.forEach((nombre) => agregarFilaProgramaEn(tbody, { programa: nombre, vr_exigido: "", vr_reconocido: "", descuento: "" }));
  card.querySelector(".btn-agregar-fila-pdf").addEventListener("click", () => agregarFilaProgramaEn(tbody));

  card.querySelector("form").addEventListener("submit", (e) => {
    e.preventDefault();
    const form = e.target;
    const programas = [];
    form.querySelectorAll(".tabla-programas-pdf tbody tr").forEach((tr) => {
      const nombre = tr.querySelector(".p-nombre").value.trim();
      if (!nombre) return;
      programas.push({
        programa: nombre,
        vr_exigido: parseFloat(tr.querySelector(".p-exigido").value) || 0,
        vr_reconocido: parseFloat(tr.querySelector(".p-reconocido").value) || 0,
        descuento: parseFloat(tr.querySelector(".p-descuento").value) || 0,
      });
    });
    const fd = new FormData(form);
    const acta = {
      municipio: fd.get("municipio"),
      contrato: fd.get("contrato"),
      acta_no: fd.get("acta_no"),
      meses: fd.get("meses"),
      anio: fd.get("anio"),
      empresa: fd.get("empresa"),
      nit: fd.get("nit"),
      regimen: fd.get("regimen"),
      archivo_origen: file.name,
      programas,
      total_ejecucion: {
        vr_exigido: parseFloat(fd.get("total_exigido")) || 0,
        vr_reconocido: parseFloat(fd.get("total_reconocido")) || 0,
        descuento: parseFloat(fd.get("total_descuento")) || 0,
      },
    };
    const periodo = fd.get("periodo");
    const checks = validar(acta);
    const chkMeses = checkConsistenciaMeses(acta, periodo);
    if (chkMeses) checks.push(chkMeses);
    consolidar(acta, periodo, checks);
    renderResultado(form.querySelector(".resultado-pdf"), { acta, checks, periodo, periodoAuto: periodo === periodoDetectado });
  });

  mostrarPDFEmbebido(file, card.querySelector(".pdf-pages"));
  return card;
}

// --------------------------------------------------------------------------
// Dropzones (arrastrar y soltar + click para seleccionar)
// --------------------------------------------------------------------------
function setupDropzone(zoneEl, inputEl, onFiles) {
  zoneEl.addEventListener("click", () => inputEl.click());
  zoneEl.addEventListener("keydown", (e) => {
    if (e.key === "Enter" || e.key === " ") { e.preventDefault(); inputEl.click(); }
  });
  ["dragenter", "dragover"].forEach((evt) =>
    zoneEl.addEventListener(evt, (e) => { e.preventDefault(); zoneEl.classList.add("dragover"); })
  );
  ["dragleave", "drop"].forEach((evt) =>
    zoneEl.addEventListener(evt, (e) => { e.preventDefault(); zoneEl.classList.remove("dragover"); })
  );
  zoneEl.addEventListener("drop", (e) => {
    const files = e.dataTransfer.files;
    if (!files || !files.length) return;
    inputEl.files = files;
    inputEl.dispatchEvent(new Event("change"));
  });
  inputEl.addEventListener("click", (e) => e.stopPropagation());
  if (onFiles) inputEl.addEventListener("change", () => onFiles(Array.from(inputEl.files)));
}

// --------------------------------------------------------------------------
// Cola de revisión para Excel: cargar -> validar/previsualizar -> confirmar
// --------------------------------------------------------------------------
let stagingIdCounter = 0;
const STAGING_XLSX = new Map(); // id -> {id, file, ruta, periodo, periodoDetectado, acta, checks, error}

async function procesarArchivosXlsx(files) {
  const soloXlsx = files.filter((f) => f.name.toLowerCase().endsWith(".xlsx"));
  document.getElementById("confirmacion-xlsx").innerHTML = "";
  for (const file of files) {
    if (!file.name.toLowerCase().endsWith(".xlsx")) continue;
    const ruta = rutaArchivo(file);
    const periodoDetectado = detectarPeriodoDesdeTexto(ruta);
    const entry = { id: ++stagingIdCounter, file, ruta, periodo: periodoDetectado, periodoDetectado };
    try {
      const buf = await file.arrayBuffer();
      const workbook = XLSX.read(buf, { type: "array", cellDates: true });
      const acta = parseActaXlsx(workbook, file.name);
      const checks = validar(acta);
      if (periodoDetectado) {
        const chkMeses = checkConsistenciaMeses(acta, periodoDetectado);
        if (chkMeses) checks.push(chkMeses);
      }
      entry.acta = acta;
      entry.checks = checks;
    } catch (err) {
      entry.error = `${file.name}: ${err.message}`;
    }
    STAGING_XLSX.set(entry.id, entry);
  }
  renderStagingXlsx();
  if (!soloXlsx.length && files.length) {
    alert("La carpeta seleccionada no tiene archivos .xlsx en la raíz o subcarpetas.");
  }
}

function recalcularEntry(entry) {
  if (!entry.acta) return;
  entry.checks = validar(entry.acta);
  const chkMeses = checkConsistenciaMeses(entry.acta, entry.periodo);
  if (chkMeses) entry.checks.push(chkMeses);
}

function renderStagingCard(entry) {
  const div = document.createElement("div");
  div.className = "staging-item";

  if (entry.error) {
    div.innerHTML = `
      <div class="toolbar">
        <h3 style="margin:0;">${entry.file.name}</h3>
        <button type="button" class="danger">✕ Quitar</button>
      </div>
      <div class="msg error">${entry.error}</div>`;
    div.querySelector("button").addEventListener("click", () => { STAGING_XLSX.delete(entry.id); renderStagingXlsx(); });
    return div;
  }

  const { acta, checks, periodo, periodoDetectado, ruta } = entry;
  const nErr = checks.filter((c) => c.resultado === "ERROR").length;
  const sinPeriodo = !periodo;
  const filas = checks
    .map((c) => `
      <tr class="${c.resultado === "ERROR" ? "row-error" : ""}">
        <td>${c.chequeo}</td>
        <td><span class="pill ${c.resultado === "OK" ? "ok" : "error"}">${c.resultado}</span></td>
        <td>${c.detalle}</td>
      </tr>`)
    .join("");

  const estadoMsg = sinPeriodo
    ? `⚠ No pude detectar el periodo de este archivo — selecciónalo para poder consolidarlo:`
    : `${nErr === 0 ? "✔" : "⚠"} Se cargó la ruta de ${acta.municipio} — periodo`;

  div.innerHTML = `
    <div class="toolbar">
      <div>
        <h3 style="margin:0;">${entry.file.name}</h3>
        ${ruta !== entry.file.name ? `<div class="hint" style="margin-top:2px;">📁 ${ruta}</div>` : ""}
      </div>
      <button type="button" class="danger">✕ Quitar</button>
    </div>
    <div class="msg ${sinPeriodo ? "error" : nErr === 0 ? "ok" : "error"}">
      ${estadoMsg}
      <select class="select-inline periodo-override"></select>
      ${periodoDetectado ? '<span class="pill ok">detectado automático</span>' : ""}
    </div>
    ${!sinPeriodo ? `
    <div class="hint" style="margin-top:8px;">
      Contrato ${acta.contrato} · Acta ${acta.acta_no} ·
      ${nErr === 0 ? `<span class="pill ok">${checks.length} chequeos OK</span>` : `<span class="pill error">${nErr} de ${checks.length} con error</span>`}
      <button type="button" class="secondary btn-ver-detalle" style="margin-left:8px; padding:3px 10px;">Ver detalle</button>
    </div>
    <div class="table-wrap detalle-checks" style="margin-top:10px; display:${nErr > 0 ? "block" : "none"};">
      <table><thead><tr><th>Chequeo</th><th>Resultado</th><th>Detalle</th></tr></thead><tbody>${filas}</tbody></table>
    </div>` : ""}`;

  div.querySelector("button.danger").addEventListener("click", () => { STAGING_XLSX.delete(entry.id); renderStagingXlsx(); });
  const sel = div.querySelector(".periodo-override");
  const opcionesPeriodo = Object.keys(PERIODO_MESES).map((k) => `<option value="${k}" ${k === periodo ? "selected" : ""}>${k}</option>`).join("");
  sel.innerHTML = (sinPeriodo ? '<option value="" selected disabled>-- elegir periodo --</option>' : "") + opcionesPeriodo;
  sel.addEventListener("change", () => {
    entry.periodo = sel.value;
    entry.periodoDetectado = false;
    recalcularEntry(entry);
    renderStagingXlsx();
  });
  const btnDetalle = div.querySelector(".btn-ver-detalle");
  if (btnDetalle) {
    btnDetalle.addEventListener("click", () => {
      const detalle = div.querySelector(".detalle-checks");
      const abierto = detalle.style.display !== "none";
      detalle.style.display = abierto ? "none" : "block";
      btnDetalle.textContent = abierto ? "Ver detalle" : "Ocultar detalle";
    });
  }

  return div;
}

function renderStagingXlsx() {
  const card = document.getElementById("card-revision-xlsx");
  const container = document.getElementById("resultado-xlsx");
  if (STAGING_XLSX.size === 0) {
    card.style.display = "none";
    container.innerHTML = "";
    return;
  }
  card.style.display = "block";
  const entries = [...STAGING_XLSX.values()].reverse();
  const nValidos = entries.filter((e) => e.acta && e.periodo).length;
  document.getElementById("btn-confirmar-xlsx").textContent = `✅ Confirmar y consolidar todo (${nValidos})`;
  container.innerHTML = "";
  entries.forEach((entry) => container.appendChild(renderStagingCard(entry)));
}

function confirmarConsolidacionXlsx() {
  let count = 0;
  const pendientesPeriodo = [];
  STAGING_XLSX.forEach((entry, id) => {
    if (entry.acta && entry.periodo) {
      consolidar(entry.acta, entry.periodo, entry.checks);
      STAGING_XLSX.delete(id);
      count++;
    } else if (entry.acta && !entry.periodo) {
      pendientesPeriodo.push(entry.file.name);
    }
  });
  renderStagingXlsx();
  let msg = `<div class="msg ok">✔ Se consolidaron ${count} archivo(s) al maestro. Revísalo en "Resumen de cumplimiento".</div>`;
  if (pendientesPeriodo.length) {
    msg += `<div class="msg error">⚠ ${pendientesPeriodo.length} archivo(s) quedaron pendientes por falta de periodo: ${pendientesPeriodo.join(", ")}. Selecciónalo en su tarjeta arriba.</div>`;
  }
  document.getElementById("confirmacion-xlsx").innerHTML = msg;
}

function descartarTodoXlsx() {
  STAGING_XLSX.clear();
  renderStagingXlsx();
  document.getElementById("confirmacion-xlsx").innerHTML = "";
}

document.addEventListener("DOMContentLoaded", () => {
  cargarEstado();

  setupDropzone(document.getElementById("dropzone-xlsx"), document.getElementById("file-xlsx"), procesarArchivosXlsx);
  setupDropzone(document.getElementById("dropzone-pdfs"), document.getElementById("input-pdfs"));

  document.getElementById("btn-folder-xlsx").addEventListener("click", () => document.getElementById("folder-xlsx").click());
  document.getElementById("folder-xlsx").addEventListener("change", (e) => {
    procesarArchivosXlsx(Array.from(e.target.files));
    e.target.value = "";
  });

  document.getElementById("btn-folder-pdfs").addEventListener("click", () => document.getElementById("folder-pdfs").click());
  document.getElementById("folder-pdfs").addEventListener("change", (e) => {
    const files = Array.from(e.target.files).filter((f) => f.name.toLowerCase().endsWith(".pdf"));
    const container = document.getElementById("pdf-items");
    files.forEach((file) => container.appendChild(crearTarjetaPDF(file)));
    e.target.value = "";
  });

  document.getElementById("btn-cargar-maestro").addEventListener("click", () => document.getElementById("input-maestro").click());
  document.getElementById("btn-limpiar-todo").addEventListener("click", limpiarTodo);
  document.getElementById("btn-confirmar-xlsx").addEventListener("click", confirmarConsolidacionXlsx);
  document.getElementById("btn-descartar-xlsx").addEventListener("click", descartarTodoXlsx);

  document.querySelectorAll(".tab-btn").forEach((btn) => {
    btn.addEventListener("click", () => {
      document.querySelectorAll(".tab-btn").forEach((b) => b.classList.remove("active"));
      document.querySelectorAll(".tab").forEach((t) => t.classList.remove("active"));
      btn.classList.add("active");
      document.getElementById(btn.dataset.tab).classList.add("active");
      if (btn.dataset.tab === "tab-resumen") cargarResumen();
      if (btn.dataset.tab === "tab-validaciones") cargarValidaciones();
    });
  });

  document.getElementById("btn-refrescar-resumen").addEventListener("click", cargarResumen);
  document.getElementById("btn-refrescar-validaciones").addEventListener("click", cargarValidaciones);
  document.getElementById("download-btn").addEventListener("click", descargarMaestro);

  document.getElementById("input-pdfs").addEventListener("change", (e) => {
    const files = Array.from(e.target.files);
    const container = document.getElementById("pdf-items");
    files.forEach((file) => container.appendChild(crearTarjetaPDF(file)));
    e.target.value = "";
  });

  document.getElementById("input-maestro").addEventListener("change", async (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const buf = await file.arrayBuffer();
    const workbook = XLSX.read(buf, { type: "array", cellDates: true });
    cargarMaestroDesdeArchivo(workbook);
    alert(`Maestro cargado: ${STATE.detalle.length} filas de detalle, ${STATE.validaciones.length} validaciones.`);
    cargarResumen();
    cargarValidaciones();
    e.target.value = "";
  });
});
