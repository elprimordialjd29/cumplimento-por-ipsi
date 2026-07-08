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

function norm(v) {
  if (v === null || v === undefined) return "";
  return String(v).trim();
}
function fmtMoney(v) {
  return (v || 0).toLocaleString("es-CO", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
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

// --------------------------------------------------------------------------
// Render
// --------------------------------------------------------------------------
function renderResultado(container, payload) {
  if (payload.error) {
    container.innerHTML = `<div class="msg error">${payload.error}</div>`;
    return;
  }
  const { acta, checks, periodo } = payload;
  const nErr = checks.filter((c) => c.resultado === "ERROR").length;
  const resumenMsg =
    nErr === 0
      ? `<div class="msg ok">✔ ${acta.municipio} — Contrato ${acta.contrato} — Acta ${acta.acta_no} — periodo ${periodo}: ${checks.length} chequeos, todos OK.</div>`
      : `<div class="msg error">⚠ ${acta.municipio} — Contrato ${acta.contrato} — Acta ${acta.acta_no} — periodo ${periodo}: ${nErr} de ${checks.length} chequeos con ERROR.</div>`;

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
  if (v === null || v === undefined || v === "") return "-";
  const pct = (v * 100).toFixed(1) + "%";
  const cls = v >= 0.999 ? "pct-100" : v < 0.9 ? "pct-bad" : "";
  return `<span class="${cls}">${pct}</span>`;
}

function cargarResumen() {
  const container = document.getElementById("resumen-contenido");
  const { rows, periodos } = rebuildResumen();
  if (rows.length === 0) {
    container.innerHTML = '<div class="empty">Aún no hay actas consolidadas.</div>';
    return;
  }
  const headerCols = ["Municipio", "Contrato", "Programa", ...periodos, "Promedio"];
  const thead = "<tr>" + headerCols.map((c) => `<th>${c}</th>`).join("") + "</tr>";
  const bodyRows = rows
    .map((r) => {
      let tds = `<td>${r.Municipio}</td><td>${r.Contrato}</td><td>${r.Programa}</td>`;
      periodos.forEach((p) => { tds += `<td class="pct-cell">${fmtPct(r[p])}</td>`; });
      tds += `<td class="pct-cell"><strong>${fmtPct(r.Promedio)}</strong></td>`;
      return `<tr>${tds}</tr>`;
    })
    .join("");
  container.innerHTML = `<table><thead>${thead}</thead><tbody>${bodyRows}</tbody></table>`;
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
function aplicarFormatoPorcentaje(ws, colIdx, numRows) {
  for (let r = 1; r < numRows; r++) {
    const addr = XLSX.utils.encode_cell({ r, c: colIdx });
    const cell = ws[addr];
    if (cell && typeof cell.v === "number") cell.z = "0.0%";
  }
}

function descargarMaestro() {
  const wb = XLSX.utils.book_new();

  const detalleHeaders = ["Municipio", "Contrato", "Acta_No", "Periodo", "Meses", "Empresa", "Nit", "Regimen",
    "Programa", "Vr_Exigido", "Vr_Reconocido", "Descuento", "Pct_Cumplimiento", "Archivo_Origen", "Fecha_Carga"];
  const detalleAoa = [detalleHeaders, ...STATE.detalle.map((r) => detalleHeaders.map((h) => (r[h] ?? null)))];
  const wsDetalle = XLSX.utils.aoa_to_sheet(detalleAoa);
  aplicarFormatoPorcentaje(wsDetalle, 12, detalleAoa.length);
  XLSX.utils.book_append_sheet(wb, wsDetalle, "Detalle");

  const valHeaders = ["Municipio", "Contrato", "Acta_No", "Periodo", "Chequeo", "Resultado", "Detalle"];
  const valAoa = [valHeaders, ...STATE.validaciones.map((r) => valHeaders.map((h) => (r[h] ?? null)))];
  const wsVal = XLSX.utils.aoa_to_sheet(valAoa);
  XLSX.utils.book_append_sheet(wb, wsVal, "Validaciones");

  const { rows, periodos } = rebuildResumen();
  const resHeaders = ["Municipio", "Contrato", "Programa", ...periodos, "Promedio"];
  const resAoa = [resHeaders, ...rows.map((r) => resHeaders.map((h) => (r[h] ?? null)))];
  const wsRes = XLSX.utils.aoa_to_sheet(resAoa);
  for (let c = 3; c < resHeaders.length; c++) aplicarFormatoPorcentaje(wsRes, c, resAoa.length);
  XLSX.utils.book_append_sheet(wb, wsRes, "Resumen_Cumplimiento");

  XLSX.writeFile(wb, "CONSOLIDADO_ACTAS_PYM_DUSAKAWI.xlsx");
}

function cargarMaestroDesdeArchivo(workbook) {
  const detalle = workbook.SheetNames.includes("Detalle")
    ? XLSX.utils.sheet_to_json(workbook.Sheets["Detalle"], { defval: null })
    : [];
  const validaciones = workbook.SheetNames.includes("Validaciones")
    ? XLSX.utils.sheet_to_json(workbook.Sheets["Validaciones"], { defval: null })
    : [];
  STATE.detalle = detalle;
  STATE.validaciones = validaciones;
  guardarEstado();
}

// --------------------------------------------------------------------------
// UI wiring
// --------------------------------------------------------------------------
function poblarSelectPeriodos(id) {
  const sel = document.getElementById(id);
  sel.innerHTML = Object.entries(PERIODO_MESES).map(([k, v]) => `<option value="${k}">${k} (${v})</option>`).join("");
}

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
  const card = document.createElement("div");
  card.className = "pdf-card";
  card.innerHTML = `
    <h3 style="margin-top:0;">${file.name}</h3>
    <div class="pdf-pages"><div class="empty">Cargando previsualización...</div></div>
    <form class="pdf-form">
      <div class="grid2">
        <div><label>Municipio</label><input type="text" name="municipio" required></div>
        <div><label>Nº Contrato</label><input type="text" name="contrato" required placeholder="Ej. 20013-061-PMT"></div>
        <div><label>Nº Acta</label><input type="text" name="acta_no" required placeholder="Ej. 20013-061-PMT-5"></div>
        <div><label>Periodo</label><select name="periodo" class="periodo-select" required></select></div>
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
    consolidar(acta, periodo, checks);
    renderResultado(form.querySelector(".resultado-pdf"), { acta, checks, periodo });
  });

  mostrarPDFEmbebido(file, card.querySelector(".pdf-pages"));
  return card;
}

document.addEventListener("DOMContentLoaded", () => {
  cargarEstado();
  poblarSelectPeriodos("periodo-xlsx");

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

  document.getElementById("form-xlsx").addEventListener("submit", async (e) => {
    e.preventDefault();
    const container = document.getElementById("resultado-xlsx");
    container.innerHTML = "";
    const files = Array.from(document.getElementById("file-xlsx").files);
    const periodo = document.getElementById("periodo-xlsx").value;
    if (!files.length) return;
    for (const file of files) {
      const wrapper = document.createElement("div");
      const titulo = document.createElement("h3");
      titulo.style.marginBottom = "4px";
      titulo.textContent = file.name;
      const resultDiv = document.createElement("div");
      wrapper.appendChild(titulo);
      wrapper.appendChild(resultDiv);
      container.appendChild(wrapper);
      try {
        const buf = await file.arrayBuffer();
        const workbook = XLSX.read(buf, { type: "array", cellDates: true });
        const acta = parseActaXlsx(workbook, file.name);
        const checks = validar(acta);
        consolidar(acta, periodo, checks);
        renderResultado(resultDiv, { acta, checks, periodo });
      } catch (err) {
        renderResultado(resultDiv, { error: `${file.name}: ${err.message}` });
      }
    }
  });

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
