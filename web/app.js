// Consolidación de Actas PYM/PMT/DNT - DUSAKAWI
// App 100% cliente: toda la extracción, validación y consolidación corre en
// el navegador. El estado persiste en localStorage y se puede exportar/
// importar como el Excel maestro para moverlo entre computadores.

const EMPRESA_ESPERADA = "DUSAKAWI IPSI";

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

  const empresa = (acta.empresa || "").trim().toUpperCase();
  add(
    "Empresa esperada",
    empresa.startsWith(EMPRESA_ESPERADA),
    `empresa='${acta.empresa || ""}' (esperado inicia con '${EMPRESA_ESPERADA}')`
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
  (acta.programas || []).forEach((p) => {
    const exigido = parseFloat(p.vr_exigido) || 0;
    const reconocido = parseFloat(p.vr_reconocido) || 0;
    const descuento = parseFloat(p.descuento) || 0;
    const pct = exigido ? reconocido / exigido : null;
    STATE.detalle.push({
      Municipio: municipio, Contrato: contrato, Acta_No: actaNo, Periodo: periodo,
      Meses: acta.meses || "", Empresa: acta.empresa || "", Nit: acta.nit || "", Regimen: acta.regimen || "",
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
      Object.keys(entry.programas).sort().forEach((programa) => {
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

// Agregado "caso extremo": todos los prestadores juntos, promediando cada
// actividad a través de todos los municipios/contratos que la reporten.
// Incluye montos totales y la lista de prestadores para que la vista no
// quede reducida a un solo número por actividad.
function rebuildResumenGeneral() {
  const porPrograma = {};
  const periodosPresentes = new Set();
  STATE.detalle.forEach((row) => {
    periodosPresentes.add(row.Periodo);
    if (!porPrograma[row.Programa]) {
      porPrograma[row.Programa] = { periodos: {}, exigido: 0, reconocido: 0, descuento: 0, prestadores: new Set() };
    }
    const p = porPrograma[row.Programa];
    if (!p.periodos[row.Periodo]) p.periodos[row.Periodo] = [];
    if (row.Pct_Cumplimiento !== null && row.Pct_Cumplimiento !== undefined) {
      p.periodos[row.Periodo].push(row.Pct_Cumplimiento);
    }
    p.exigido += Number(row.Vr_Exigido) || 0;
    p.reconocido += Number(row.Vr_Reconocido) || 0;
    p.descuento += Number(row.Descuento) || 0;
    p.prestadores.add(row.Municipio);
  });

  const periodosOrdenados = PERIODO_ORDEN.filter((p) => periodosPresentes.has(p)).concat(
    [...periodosPresentes].filter((p) => !PERIODO_ORDEN.includes(p)).sort()
  );

  const rows = Object.keys(porPrograma).sort().map((programa) => {
    const p = porPrograma[programa];
    const row = {
      Programa: programa,
      Vr_Exigido: p.exigido,
      Vr_Reconocido: p.reconocido,
      Descuento: p.descuento,
      Prestadores: [...p.prestadores].sort().join(", "),
      NumPrestadores: p.prestadores.size,
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

function fmtPct(v) {
  if (v === null || v === undefined || v === "") return '<span style="color:#c0c4cc;">—</span>';
  const pct = v * 100;
  const cls = pct >= 99.9 ? "pct-pill-ok" : pct >= 90 ? "pct-pill-mid" : "pct-pill-bad";
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
function tablaResumenGeneralHtml(rows, periodos) {
  const headerCols = ["Programa / Actividad", "Vr Exigido", "Vr Reconocido", "Descuento", ...periodos, "Promedio", "Prestadores que reportan"];
  const thead = "<tr>" + headerCols.map((c) => `<th>${c}</th>`).join("") + "</tr>";
  const body = rows
    .map((r) => {
      let tds = `<td>${r.Programa}</td>`;
      tds += `<td style="text-align:right; font-variant-numeric:tabular-nums;">${fmtMoney(r.Vr_Exigido)}</td>`;
      tds += `<td style="text-align:right; font-variant-numeric:tabular-nums;">${fmtMoney(r.Vr_Reconocido)}</td>`;
      tds += `<td style="text-align:right; font-variant-numeric:tabular-nums;">${fmtMoney(r.Descuento)}</td>`;
      periodos.forEach((p) => { tds += `<td class="pct-cell">${fmtPct(r[p])}</td>`; });
      tds += `<td class="pct-cell" style="background:#f7f9ff;">${fmtPct(r.Promedio)}</td>`;
      tds += `<td class="hint" style="white-space:normal; max-width:220px;">${r.Prestadores} <span class="pill ok">${r.NumPrestadores}</span></td>`;
      return `<tr>${tds}</tr>`;
    })
    .join("");
  return `<div class="table-wrap"><table><thead>${thead}</thead><tbody>${body}</tbody></table></div>`;
}

function escapeXml(s) {
  return String(s).replace(/[<>&"]/g, (c) => ({ "<": "&lt;", ">": "&gt;", "&": "&amp;", '"': "&quot;" }[c]));
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
    const color = pct >= 99.9 ? "#16924f" : pct >= 90 ? "#b7791f" : "#d0342c";
    const cx = x + barW / 2;
    const labelY = padTop + chartH + 16;
    const nombreCorto = r.Programa.length > 30 ? r.Programa.slice(0, 28) + "…" : r.Programa;
    bars += `
      <rect x="${x.toFixed(1)}" y="${y.toFixed(1)}" width="${barW.toFixed(1)}" height="${Math.max(barH, 1).toFixed(1)}" fill="${color}" rx="4"/>
      <text x="${cx.toFixed(1)}" y="${(y - 7).toFixed(1)}" text-anchor="middle" font-size="12" font-weight="700" fill="${color}">${r.Promedio === null ? "-" : pct.toFixed(0) + "%"}</text>
      <text x="${cx.toFixed(1)}" y="${labelY}" font-size="10" fill="#475067" text-anchor="end" transform="rotate(-40 ${cx.toFixed(1)} ${labelY})">${escapeXml(nombreCorto)}</text>`;
  });

  const y100 = padTop;
  return `<svg viewBox="0 0 ${w} ${h}" style="width:100%; max-width:${w}px; height:auto; display:block; margin:0 auto;">
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

  // Vista principal: UN SOLO consolidado, sin importar el municipio — una
  // fila por actividad, promediando entre todos los prestadores y periodos.
  const { rows: rowsGeneral, periodos: periodosGeneral } = rebuildResumenGeneral();

  let html = `<div class="resumen-grupo resumen-general">
    <h3>🌐 Consolidado general (todos los prestadores)</h3>
    <p class="hint">Una fila por actividad — el % es el promedio entre todos los municipios/contratos que la reportan (nunca supera 100%). Los montos son la suma de todos los prestadores.</p>
    ${graficoBarrasSVG(rowsGeneral)}
    ${tablaResumenGeneralHtml(rowsGeneral, periodosGeneral)}
  </div>`;

  // Detalle: desglose por prestador, para quien necesite ver el origen.
  const grupos = new Map();
  rows.forEach((r) => {
    const key = r.Municipio + "||" + r.Contrato;
    if (!grupos.has(key)) grupos.set(key, { municipio: r.Municipio, contrato: r.Contrato, filas: [] });
    grupos.get(key).filas.push(r);
  });

  html += `<div class="resumen-detalle-toggle">
    <button type="button" class="secondary" id="btn-toggle-detalle-prestador">▸ Ver desglose por prestador (${grupos.size})</button>
  </div>
  <div id="detalle-por-prestador" style="display:none; margin-top:16px;">`;
  grupos.forEach((g) => {
    const headerCols = ["Programa", ...periodos, "Promedio"];
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
    btnToggle.textContent = `${abierto ? "▸" : "▾"} Ver desglose por prestador (${grupos.size})`;
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

function descargarMaestro() {
  const wb = XLSX.utils.book_new();

  const detalleHeaders = DETALLE_COLUMNAS.map((c) => c.label);
  const detalleAoa = [detalleHeaders, ...STATE.detalle.map((r) => DETALLE_COLUMNAS.map((c) => r[c.key] ?? null))];
  const wsDetalle = hojaConEstilo(detalleAoa, { columnasMoneda: [9, 10, 11], columnasPorcentaje: [12] });
  XLSX.utils.book_append_sheet(wb, wsDetalle, "Detalle");

  const valHeaders = VALIDACIONES_COLUMNAS.map((c) => c.label);
  const valAoa = [valHeaders, ...STATE.validaciones.map((r) => VALIDACIONES_COLUMNAS.map((c) => r[c.key] ?? null))];
  const wsVal = hojaConEstilo(valAoa, {});
  XLSX.utils.book_append_sheet(wb, wsVal, "Validaciones");

  // Resumen_General primero: es el consolidado principal (una fila por
  // actividad, sin importar el municipio), con montos totales y qué
  // prestadores la reportan — no solo un porcentaje suelto.
  const { rows: rowsGeneral, periodos: periodosGeneral } = rebuildResumenGeneral();
  const resGeneralHeaders = ["Programa / Actividad", "Vr Exigido", "Vr Reconocido", "Descuento", ...periodosGeneral, "Promedio", "Nº Prestadores", "Prestadores"];
  const resGeneralAoa = [
    resGeneralHeaders,
    ...rowsGeneral.map((r) => [
      r.Programa, r.Vr_Exigido, r.Vr_Reconocido, r.Descuento,
      ...periodosGeneral.map((p) => r[p] ?? null), r.Promedio ?? null, r.NumPrestadores, r.Prestadores,
    ]),
  ];
  const wsResGeneral = hojaConEstilo(resGeneralAoa, {
    columnasMoneda: [1, 2, 3],
    columnasPorcentaje: periodosGeneral.map((_, i) => 4 + i).concat([4 + periodosGeneral.length]),
  });
  XLSX.utils.book_append_sheet(wb, wsResGeneral, "Resumen_General");

  const { rows, periodos } = rebuildResumen();
  const resHeaders = ["Municipio", "Nº Contrato", "Programa / Actividad", ...periodos, "Promedio"];
  const resAoa = [
    resHeaders,
    ...rows.map((r) => [r.Municipio, r.Contrato, r.Programa, ...periodos.map((p) => r[p] ?? null), r.Promedio ?? null]),
  ];
  const wsRes = hojaConEstilo(resAoa, { columnasPorcentaje: periodos.map((_, i) => 3 + i).concat([3 + periodos.length]) });
  XLSX.utils.book_append_sheet(wb, wsRes, "Resumen_Detalle_Prestador");

  XLSX.writeFile(wb, "CONSOLIDADO_ACTAS_PYM_DUSAKAWI.xlsx");
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
        <div><label>Empresa</label><input type="text" name="empresa" value="DUSAKAWI IPSI"></div>
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
    : `${nErr === 0 ? "✔" : "⚠"} ${acta.municipio} — Contrato ${acta.contrato} — Acta ${acta.acta_no} — periodo detectado:`;

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
      ${!sinPeriodo ? `: ${checks.length} chequeos, ${nErr === 0 ? "todos OK" : `${nErr} con ERROR`}.` : ""}
    </div>
    ${!sinPeriodo ? `<div class="table-wrap" style="margin-top:10px;">
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
