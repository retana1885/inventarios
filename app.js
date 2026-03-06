const state = {
  fileName: "",
  rows: [],
  selectedFila: "TODAS",
  searchQuery: "",
  onlyDifferences: false,
  controlsCollapsed: false,
  collapseTimerId: null,
  isDirty: false,
  lastSavedAt: null,
  autosaveMs: 20000,
  detectedColumns: null,
};

const STORAGE_KEY = "inventario_conteo";

const ui = {
  fileInput: document.getElementById("fileInput"),
  fileDropZone: document.getElementById("fileDropZone"),
  chooseFileBtn: document.getElementById("chooseFileBtn"),
  selectedFileText: document.getElementById("selectedFileText"),
  controlsSection: document.getElementById("controlsSection"),
  toggleControlsTab: document.getElementById("toggleControlsTab"),
  fileMeta: document.getElementById("fileMeta"),
  saveBtn: document.getElementById("saveBtn"),
  restoreBtn: document.getElementById("restoreBtn"),
  clearBtn: document.getElementById("clearBtn"),
  exportBtn: document.getElementById("exportBtn"),
  globalStatsSection: document.getElementById("globalStatsSection"),
  globalStats: document.getElementById("globalStats"),
  filtersSection: document.getElementById("filtersSection"),
  filaButtons: document.getElementById("filaButtons"),
  nextPendingBtn: document.getElementById("nextPendingBtn"),
  diffOnlyBtn: document.getElementById("diffOnlyBtn"),
  searchInput: document.getElementById("searchInput"),
  clearSearchBtn: document.getElementById("clearSearchBtn"),
  activeFilter: document.getElementById("activeFilter"),
  autosaveStatus: document.getElementById("autosaveStatus"),
  progressSection: document.getElementById("progressSection"),
  progressStats: document.getElementById("progressStats"),
  tableSection: document.getElementById("tableSection"),
  tableBody: document.getElementById("tableBody"),
  emptyTableMessage: document.getElementById("emptyTableMessage"),
  floatingDots: document.getElementById("floatingDots"),
};

ui.fileInput.addEventListener("change", onFileSelected);
ui.chooseFileBtn.addEventListener("click", () => ui.fileInput.click());
ui.fileDropZone.addEventListener("click", (event) => {
  if (event.target.id !== "chooseFileBtn") {
    ui.fileInput.click();
  }
});
ui.fileDropZone.addEventListener("keydown", (event) => {
  if (event.key === "Enter" || event.key === " ") {
    event.preventDefault();
    ui.fileInput.click();
  }
});
ui.fileDropZone.addEventListener("dragover", onDragOver);
ui.fileDropZone.addEventListener("dragleave", onDragLeave);
ui.fileDropZone.addEventListener("drop", onDropFile);
ui.toggleControlsTab.addEventListener("click", toggleControls);
ui.saveBtn.addEventListener("click", saveProgress);
ui.restoreBtn.addEventListener("click", restoreProgress);
ui.clearBtn.addEventListener("click", clearProgress);
ui.exportBtn.addEventListener("click", exportResults);
ui.nextPendingBtn.addEventListener("click", goToNextPending);
ui.diffOnlyBtn.addEventListener("click", toggleOnlyDifferences);
ui.searchInput.addEventListener("input", onSearchChanged);
ui.clearSearchBtn.addEventListener("click", clearSearch);

initializeFloatingDots();
setInterval(runAutosave, state.autosaveMs);
setInterval(renderAutosaveStatus, 1000);

function onFileSelected(event) {
  const [file] = event.target.files;
  if (!file) {
    return;
  }

  processFile(file);
}

function processFile(file) {
  ui.selectedFileText.textContent = file.name;

  const reader = new FileReader();
  reader.onload = () => {
    try {
      parseWorkbook(reader.result, file.name);
    } catch (error) {
      console.error(error);
      alert("No se pudo leer el archivo. Verifica que sea un Excel valido.");
    }
  };
  reader.readAsArrayBuffer(file);
}

function onDragOver(event) {
  event.preventDefault();
  ui.fileDropZone.classList.add("is-dragging");
}

function onDragLeave() {
  ui.fileDropZone.classList.remove("is-dragging");
}

function onDropFile(event) {
  event.preventDefault();
  ui.fileDropZone.classList.remove("is-dragging");

  const [file] = event.dataTransfer.files || [];
  if (!file) {
    return;
  }
  if (!/\.xlsx?$/.test(file.name.toLowerCase())) {
    alert("Solo se permite archivo Excel (.xlsx o .xls).");
    return;
  }

  processFile(file);
}

function parseWorkbook(arrayBuffer, fileName) {
  const workbook = XLSX.read(arrayBuffer, { type: "array" });
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  const matrix = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });

  const detection = detectHeaderRow(matrix);
  if (!detection) {
    throw new Error("No se detectaron columnas clave.");
  }

  const parsedRows = [];
  for (let i = detection.headerRowIndex + 1; i < matrix.length; i += 1) {
    const line = matrix[i];
    const codigo = getCell(line, detection.columns.codigo);
    const nombre = getCell(line, detection.columns.nombre);
    const posicion = getCell(line, detection.columns.posicion);
    const existenciaRaw = getCell(line, detection.columns.existencia);

    if (!codigo && !nombre && !posicion) {
      continue;
    }

    const existencia = parseNumber(existenciaRaw);
    const fila = extractFila(posicion);

    parsedRows.push({
      id: parsedRows.length + 1,
      codigo,
      nombre,
      posicion,
      fila,
      existencia,
      conteoFisico: "",
      diferencia: null,
    });
  }

  if (!parsedRows.length) {
    throw new Error("No se encontraron filas de inventario en la hoja.");
  }

  state.fileName = fileName;
  state.rows = parsedRows;
  state.selectedFila = "TODAS";
  state.searchQuery = "";
  state.onlyDifferences = false;
  state.isDirty = false;
  state.detectedColumns = detection.columns;
  ui.searchInput.value = "";

  ui.fileMeta.textContent = `Archivo cargado: ${fileName} | Registros: ${parsedRows.length}. Se ocultara en 3s.`;
  ui.controlsSection.classList.add("loaded-feedback");
  ui.globalStatsSection.hidden = false;
  ui.filtersSection.hidden = false;
  ui.progressSection.hidden = false;
  ui.tableSection.hidden = false;

  ui.saveBtn.disabled = false;
  ui.restoreBtn.disabled = false;
  ui.clearBtn.disabled = false;
  ui.exportBtn.disabled = false;
  ui.nextPendingBtn.disabled = false;
  ui.diffOnlyBtn.disabled = false;

  renderFilaButtons();
  renderGlobalStats();
  updateQuickActionButtons();
  renderProgressStats();
  renderTable();

  state.controlsCollapsed = false;
  scheduleControlsAutoHide();
  renderControlsVisibility();
  renderAutosaveStatus();
}

function detectHeaderRow(matrix) {
  const candidates = {
    codigo: ["CODIGO", "CLAVE", "ARTICULO"],
    nombre: ["NOMBRE", "DESCRIPCION"],
    posicion: ["POSICION", "UBICACION", "LOCALIZACION"],
    existencia: ["EXISTENCIA", "EXISTENCIAS", "STOCK"],
  };

  for (let rowIndex = 0; rowIndex < Math.min(matrix.length, 60); rowIndex += 1) {
    const row = matrix[rowIndex] || [];
    const columns = {
      codigo: -1,
      nombre: -1,
      posicion: -1,
      existencia: -1,
    };

    row.forEach((cell, index) => {
      const text = normalizeHeader(cell);
      if (!text) {
        return;
      }

      Object.keys(candidates).forEach((key) => {
        if (columns[key] !== -1) {
          return;
        }

        if (candidates[key].some((keyword) => text.includes(keyword))) {
          columns[key] = index;
        }
      });
    });

    const foundCount = Object.values(columns).filter((column) => column !== -1).length;
    if (foundCount >= 4) {
      return {
        headerRowIndex: rowIndex,
        columns,
      };
    }
  }

  return null;
}

function renderGlobalStats() {
  const totalProductos = state.rows.length;
  const sumaPiezas = state.rows.reduce((sum, row) => sum + row.existencia, 0);
  const existenciaCero = state.rows.filter((row) => row.existencia === 0).length;
  const sinFila = state.rows.filter((row) => row.fila === "SIN_FILA").length;

  const porFila = computePerFilaStats(state.rows);
  const filasConDatos = Object.keys(porFila).length;

  ui.globalStats.innerHTML = "";
  const cards = [
    { label: "Productos", value: formatNumber(totalProductos), tone: "neutral" },
    { label: "Suma de piezas", value: formatNumber(sumaPiezas), tone: "info" },
    { label: "Existencia cero", value: formatNumber(existenciaCero), tone: "pending" },
    { label: "Filas detectadas", value: formatNumber(filasConDatos), tone: "info" },
    { label: "Registros sin fila", value: formatNumber(sinFila), tone: sinFila > 0 ? "missing" : "ok" },
  ];

  cards.forEach((card) => ui.globalStats.appendChild(createStatCard(card)));
}

function renderFilaButtons() {
  ui.filaButtons.innerHTML = "";
  const filaSet = new Set(state.rows.map((row) => row.fila));
  const filaList = Array.from(filaSet)
    .filter((fila) => /^F\d+$/.test(fila))
    .sort((a, b) => Number(a.slice(1)) - Number(b.slice(1)));

  const buttons = ["TODAS", ...filaList, ...(filaSet.has("SIN_FILA") ? ["SIN_FILA"] : [])];

  buttons.forEach((fila) => {
    const button = document.createElement("button");
    button.type = "button";
    button.className = "fila-btn";
    button.textContent = fila === "TODAS" ? "Todas" : fila;
    if (fila === state.selectedFila) {
      button.classList.add("active");
    }

    button.addEventListener("click", () => {
      state.selectedFila = fila;
      renderFilaButtons();
      renderTable();
      renderProgressStats();
      updateQuickActionButtons();
    });

    ui.filaButtons.appendChild(button);
  });
}

function renderProgressStats() {
  const rows = getFilteredRows();
  const total = rows.length;
  const capturados = rows.filter((row) => row.conteoFisico !== "").length;
  const pendientes = total - capturados;
  const ok = rows.filter((row) => row.conteoFisico !== "" && row.diferencia === 0).length;
  const faltantes = rows.filter((row) => row.conteoFisico !== "" && row.diferencia < 0).length;
  const sobrantes = rows.filter((row) => row.conteoFisico !== "" && row.diferencia > 0).length;
  const diferencias = faltantes + sobrantes;
  const porcentaje = total ? ((capturados / total) * 100).toFixed(1) : "0.0";
  const progressValue = Number(porcentaje);

  ui.progressStats.innerHTML = "";
  const topCards = [
    { label: "Registros", value: formatNumber(total), tone: "neutral" },
    {
      label: "Capturados",
      value: formatNumber(capturados),
      tone: capturados === 0 ? "neutral" : capturados === total ? "ok" : "info",
    },
    { label: "Pendientes", value: formatNumber(pendientes), tone: pendientes > 0 ? "pending" : "ok" },
    {
      label: "Avance",
      value: `${porcentaje}%`,
      tone: progressValue >= 100 ? "ok" : progressValue >= 50 ? "info" : "pending",
      hint: `${formatNumber(capturados)} de ${formatNumber(total)} capturados`,
      progress: progressValue,
    },
  ];

  const liveCards = [
    { label: "OK", value: formatNumber(ok), tone: ok > 0 ? "ok" : "neutral" },
    { label: "Faltantes", value: formatNumber(faltantes), tone: faltantes > 0 ? "missing" : "neutral" },
    { label: "Sobrantes", value: formatNumber(sobrantes), tone: sobrantes > 0 ? "surplus" : "neutral" },
    { label: "Con ajuste", value: formatNumber(diferencias), tone: diferencias > 0 ? "missing" : "ok" },
  ];

  ui.progressStats.appendChild(createStatsRow(topCards, "kpi-row top-row"));
  ui.progressStats.appendChild(createStatsRow(liveCards, "kpi-row live-row"));
}

function renderTable() {
  const rows = getFilteredRows();
  ui.tableBody.innerHTML = "";
  ui.emptyTableMessage.hidden = rows.length > 0;

  if (!rows.length) {
    return;
  }

  const fragment = document.createDocumentFragment();

  rows.forEach((row) => {
    const tr = document.createElement("tr");
    tr.className = getRowStatusClass(row);

    const codigoTd = document.createElement("td");
    codigoTd.textContent = row.codigo;
    tr.appendChild(codigoTd);

    const nombreTd = document.createElement("td");
    nombreTd.textContent = row.nombre;
    tr.appendChild(nombreTd);

    const posicionTd = document.createElement("td");
    posicionTd.textContent = row.posicion;
    tr.appendChild(posicionTd);

    const existenciaTd = document.createElement("td");
    existenciaTd.textContent = formatNumber(row.existencia);
    tr.appendChild(existenciaTd);

    const conteoTd = document.createElement("td");
    const input = document.createElement("input");
    input.className = "capture-input";
    input.dataset.rowId = String(row.id);
    input.type = "number";
    input.min = "0";
    input.step = "1";
    input.value = row.conteoFisico;
    input.placeholder = "Capturar";
    input.addEventListener("input", () => {
      const value = input.value.trim();
      row.conteoFisico = value;

      if (value === "") {
        row.diferencia = null;
      } else {
        row.diferencia = Number(value) - row.existencia;
      }

      state.isDirty = true;

      const { sobrante, faltante } = getSobranteFaltante(row);
      sobranteTd.textContent = sobrante === null ? "" : formatNumber(sobrante);
      faltanteTd.textContent = faltante === null ? "" : formatNumber(faltante);
      const { text, className } = getStatus(row);
      statusTd.textContent = text;
      statusTd.className = className;
      tr.className = getRowStatusClass(row);
      if (state.onlyDifferences && row.diferencia === 0) {
        renderTable();
      }
      renderProgressStats();
      updateQuickActionButtons();
    });

    input.addEventListener("keydown", (event) => {
      if (event.key !== "Enter") {
        return;
      }
      event.preventDefault();
      focusNextPendingInput(row.id);
    });

    conteoTd.appendChild(input);
    tr.appendChild(conteoTd);

    const sobranteTd = document.createElement("td");
    const faltanteTd = document.createElement("td");
    const { sobrante, faltante } = getSobranteFaltante(row);
    sobranteTd.textContent = sobrante === null ? "" : formatNumber(sobrante);
    faltanteTd.textContent = faltante === null ? "" : formatNumber(faltante);
    tr.appendChild(sobranteTd);
    tr.appendChild(faltanteTd);

    const statusTd = document.createElement("td");
    const { text, className } = getStatus(row);
    statusTd.textContent = text;
    statusTd.className = className;
    tr.appendChild(statusTd);

    fragment.appendChild(tr);
  });

  ui.tableBody.appendChild(fragment);
}

function saveProgress() {
  performSave({ silent: false });
}

function performSave({ silent }) {
  if (!state.rows.length) {
    return;
  }

  const payload = {
    signature: getFileSignature(),
    fileName: state.fileName,
    timestamp: new Date().toISOString(),
    counts: state.rows
      .filter((row) => row.conteoFisico !== "")
      .map((row) => ({ id: row.id, conteoFisico: row.conteoFisico })),
  };

  localStorage.setItem(STORAGE_KEY, JSON.stringify(payload));
  state.lastSavedAt = new Date();
  state.isDirty = false;
  renderAutosaveStatus();
  if (!silent) {
    alert("Avance guardado localmente en este navegador.");
  }
}

function restoreProgress() {
  if (!state.rows.length) {
    return;
  }

  const raw = localStorage.getItem(STORAGE_KEY);
  if (!raw) {
    alert("No hay avance guardado.");
    return;
  }

  const payload = JSON.parse(raw);
  if (payload.signature !== getFileSignature()) {
    alert("El avance guardado no corresponde al archivo actual.");
    return;
  }

  const countMap = new Map(payload.counts.map((item) => [item.id, item.conteoFisico]));
  state.rows.forEach((row) => {
    const restored = countMap.get(row.id);
    row.conteoFisico = restored ?? "";
    row.diferencia = row.conteoFisico === "" ? null : Number(row.conteoFisico) - row.existencia;
  });

  state.isDirty = false;
  state.lastSavedAt = payload.timestamp ? new Date(payload.timestamp) : new Date();

  renderTable();
  renderProgressStats();
  updateQuickActionButtons();
  renderAutosaveStatus();
  alert("Avance restaurado correctamente.");
}

function clearProgress() {
  localStorage.removeItem(STORAGE_KEY);
  state.rows.forEach((row) => {
    row.conteoFisico = "";
    row.diferencia = null;
  });

  state.isDirty = false;
  state.lastSavedAt = null;

  renderTable();
  renderProgressStats();
  updateQuickActionButtons();
  renderAutosaveStatus();
  alert("Se limpio el avance local.");
}

async function exportResults() {
  if (!state.rows.length) {
    return;
  }

  ui.exportBtn.disabled = true;
  try {
    const wbTemplate = await loadTemplateWorkbook();
    const rows = buildTemplateRows(state.rows);
    fillWorkbookFromTemplate(wbTemplate, rows);

    const stamp = buildTimestamp();
    const outName = `RESULTADO_CONTEO_${stamp}.xlsx`;
    XLSX.writeFile(wbTemplate, outName, { cellStyles: true, bookType: "xlsx" });
  } catch (error) {
    console.error(error);
    alert("No se pudo aplicar la plantilla de exportacion. Se exportara en formato estandar.");
    exportStandardWorkbook();
  } finally {
    ui.exportBtn.disabled = false;
  }
}

function exportStandardWorkbook() {
  const data = state.rows.map((row) => {
    const status = getStatus(row).text;
    const { sobrante, faltante } = getSobranteFaltante(row);
    return {
      CODIGO: row.codigo,
      NOMBRE: row.nombre,
      POSICION: row.posicion,
      FILA: row.fila,
      EXISTENCIA: row.existencia,
      "CONTEO FISICO": row.conteoFisico === "" ? null : Number(row.conteoFisico),
      SOBRANTE: sobrante,
      FALTANTE: faltante,
      ESTATUS: status,
    };
  });

  const wb = XLSX.utils.book_new();
  const wsData = XLSX.utils.json_to_sheet(data);
  XLSX.utils.book_append_sheet(wb, wsData, "CONTEO");

  const stamp = buildTimestamp();
  const outName = `RESULTADO_CONTEO_${stamp}.xlsx`;
  XLSX.writeFile(wb, outName);
}

async function loadTemplateWorkbook() {
  const response = await fetch("./Plantilla_exportar_datos.xlsx", { cache: "no-store" });
  if (!response.ok) {
    throw new Error("Plantilla no encontrada en la carpeta del proyecto.");
  }

  const buffer = await response.arrayBuffer();
  return XLSX.read(buffer, {
    type: "array",
    cellStyles: true,
    cellNF: true,
    cellDates: true,
  });
}

function buildTemplateRows(rows) {
  return [...rows]
    .sort((a, b) => compareRowsForExport(a, b))
    .map((row) => {
    const conteo = row.conteoFisico === "" ? "" : Number(row.conteoFisico);
    const diferencia = conteo === "" ? "" : conteo - row.existencia;
    return [
      row.codigo,
      row.nombre,
      row.posicion,
      row.existencia,
      conteo,
      diferencia,
    ];
    });
}

function fillWorkbookFromTemplate(workbook, rows) {
  const baseSheetName = workbook.SheetNames[0];
  if (!baseSheetName) {
    throw new Error("La plantilla no contiene hojas.");
  }

  const baseSheet = workbook.Sheets[baseSheetName];
  if (!baseSheet) {
    throw new Error("No se encontro la hoja base en la plantilla.");
  }

  while (workbook.SheetNames.length > 1) {
    const name = workbook.SheetNames.pop();
    delete workbook.Sheets[name];
  }

  clearTemplateDataRows(baseSheet, 7, Math.max(4000, rows.length + 20));
  if (rows.length) {
    XLSX.utils.sheet_add_aoa(baseSheet, rows, { origin: "A7" });
  }

  applyExportTableStyle(baseSheet, rows.length);

  baseSheet.F1 = {
    t: "s",
    v: `* ${formatTemplateDate(new Date())}`,
  };

  applyTemplateTitleStyle(baseSheet);
  applyTemplateLegendStyles(baseSheet);
}

function clearTemplateDataRows(ws, startRow, rowCount) {
  for (let row = startRow; row < startRow + rowCount; row += 1) {
    ["A", "B", "C", "D", "E", "F"].forEach((col) => {
      const addr = `${col}${row}`;
      if (ws[addr]) {
        ws[addr].v = "";
        ws[addr].w = "";
      }
    });
  }
}

function applyExportTableStyle(ws, dataRowCount) {
  const lastRow = Math.max(7, 6 + dataRowCount);
  const cols = ["A", "B", "C", "D", "E", "F"];

  const existingRange = ws["!ref"] ? XLSX.utils.decode_range(ws["!ref"]) : null;
  const endCol = existingRange ? Math.max(existingRange.e.c, 5) : 5;
  const endRow = existingRange ? Math.max(existingRange.e.r + 1, lastRow) : lastRow;
  ws["!ref"] = XLSX.utils.encode_range({ s: { c: 0, r: 0 }, e: { c: endCol, r: endRow - 1 } });
  ws["!autofilter"] = { ref: `A6:F${lastRow}` };
  ws["!cols"] = [
    { wch: 12 },
    { wch: 62 },
    { wch: 16 },
    { wch: 14 },
    { wch: 16 },
    { wch: 14 },
  ];

  const border = {
    top: { style: "thin", color: { rgb: "1A1A1A" } },
    bottom: { style: "thin", color: { rgb: "1A1A1A" } },
    left: { style: "thin", color: { rgb: "1A1A1A" } },
    right: { style: "thin", color: { rgb: "1A1A1A" } },
  };

  cols.forEach((col) => {
    const headerAddr = `${col}6`;
    if (!ws[headerAddr]) {
      ws[headerAddr] = { t: "s", v: "" };
    }
    ws[headerAddr].s = {
      font: { name: "Aptos Narrow", sz: 10, bold: true, color: { rgb: "FFFFFF" } },
      fill: { fgColor: { rgb: "000000" } },
      alignment: { horizontal: "center", vertical: "center" },
      border,
    };
  });

  for (let row = 7; row <= lastRow; row += 1) {
    const isAlt = row % 2 === 0;
    const fill = { fgColor: { rgb: isAlt ? "F2F2F2" : "FFFFFF" } };

    cols.forEach((col, index) => {
      const addr = `${col}${row}`;
      if (!ws[addr]) {
        ws[addr] = { t: "s", v: "" };
      }

      const align = index >= 3
        ? { horizontal: "right", vertical: "center" }
        : { horizontal: "left", vertical: "center" };

      ws[addr].s = {
        font: { name: "Aptos Narrow", sz: 10, color: { rgb: "111827" } },
        fill,
        alignment: align,
        border,
      };
    });
  }
}

function applyTemplateTitleStyle(ws) {
  const titleAddr = "B4";
  if (!ws[titleAddr]) {
    return;
  }

  const currentStyle = ws[titleAddr].s || {};
  const currentFont = currentStyle.font || {};
  ws[titleAddr].s = {
    ...currentStyle,
    font: {
      ...currentFont,
      bold: true,
    },
  };
}

function applyTemplateLegendStyles(ws) {
  applyCellFontStyle(ws, "A1", { bold: true, color: { rgb: "C00000" } });
  applyCellFontStyle(ws, "A2", { bold: true, color: { rgb: "153D64" } });
}

function applyCellFontStyle(ws, addr, fontPatch) {
  if (!ws[addr]) {
    return;
  }

  const currentStyle = ws[addr].s || {};
  const currentFont = currentStyle.font || {};
  ws[addr].s = {
    ...currentStyle,
    font: {
      ...currentFont,
      ...fontPatch,
    },
  };
}

function compareRowsForExport(a, b) {
  const filaA = filaSortValue(a.fila);
  const filaB = filaSortValue(b.fila);
  if (filaA !== filaB) {
    return filaA - filaB;
  }

  const posA = normalizeText(a.posicion);
  const posB = normalizeText(b.posicion);
  if (posA < posB) return -1;
  if (posA > posB) return 1;

  const codA = normalizeText(a.codigo);
  const codB = normalizeText(b.codigo);
  if (codA < codB) return -1;
  if (codA > codB) return 1;
  return 0;
}

function filaSortValue(fila) {
  const match = String(fila || "").toUpperCase().match(/^F(\d+)$/);
  if (!match) {
    return 999;
  }
  return Number(match[1]);
}

function formatTemplateDate(date) {
  const d = String(date.getDate()).padStart(2, "0");
  const m = String(date.getMonth() + 1).padStart(2, "0");
  const y = String(date.getFullYear()).slice(-2);
  return `${d}.${m}.${y}`;
}

function computePerFilaStats(rows) {
  const result = {};
  rows.forEach((row) => {
    if (!result[row.fila]) {
      result[row.fila] = { total: 0, capturados: 0, diferencias: 0 };
    }

    result[row.fila].total += 1;
    if (row.conteoFisico !== "") {
      result[row.fila].capturados += 1;
      if (row.diferencia !== 0) {
        result[row.fila].diferencias += 1;
      }
    }
  });
  return result;
}

function getFilteredRows() {
  return getScopedRows();
}

function getScopedRows(options = {}) {
  const { ignoreDiffMode = false } = options;

  let rows = state.selectedFila === "TODAS"
    ? state.rows
    : state.rows.filter((row) => row.fila === state.selectedFila);

  if (state.searchQuery) {
    const term = normalizeText(state.searchQuery);
    rows = rows
      .map((row) => ({ row, score: getSearchScore(row, term) }))
      .filter((item) => item.score > 0)
      .sort((a, b) => b.score - a.score || a.row.id - b.row.id)
      .map((item) => item.row);
  }

  if (!ignoreDiffMode && state.onlyDifferences) {
    rows = rows.filter((row) => row.conteoFisico !== "" && row.diferencia !== 0);
  }

  return rows;
}

function onSearchChanged(event) {
  state.searchQuery = event.target.value.trim();
  renderTable();
  renderProgressStats();
  updateQuickActionButtons();
}

function clearSearch() {
  if (!state.searchQuery) {
    return;
  }
  state.searchQuery = "";
  ui.searchInput.value = "";
  renderTable();
  renderProgressStats();
  updateQuickActionButtons();
}

function toggleOnlyDifferences() {
  state.onlyDifferences = !state.onlyDifferences;
  updateQuickActionButtons();
  renderTable();
  renderProgressStats();
}

function toggleControls() {
  if (!state.rows.length) {
    return;
  }
  if (state.collapseTimerId) {
    clearTimeout(state.collapseTimerId);
    state.collapseTimerId = null;
  }
  state.controlsCollapsed = !state.controlsCollapsed;
  renderControlsVisibility();
}

function goToNextPending() {
  focusNextPendingInput(null);
}

function focusNextPendingInput(currentRowId) {
  if (!state.rows.length) {
    return;
  }

  if (state.onlyDifferences) {
    state.onlyDifferences = false;
    updateQuickActionButtons();
    renderTable();
    renderProgressStats();
  }

  const scopedRows = getScopedRows({ ignoreDiffMode: true });
  if (!scopedRows.length) {
    alert("No hay registros con los filtros actuales.");
    return;
  }

  const nextRow = getNextPendingRow(scopedRows, currentRowId);
  if (!nextRow) {
    alert("No hay articulos pendientes en este filtro.");
    return;
  }

  const selector = `input.capture-input[data-row-id="${nextRow.id}"]`;
  const nextInput = document.querySelector(selector);
  if (!nextInput) {
    return;
  }

  nextInput.focus();
  nextInput.select();
  nextInput.scrollIntoView({ block: "center", behavior: "smooth" });
}

function getNextPendingRow(scopedRows, currentRowId) {
  const startIndex = currentRowId == null ? -1 : scopedRows.findIndex((row) => row.id === currentRowId);
  const tail = scopedRows.slice(startIndex + 1);
  const nextTail = tail.find((row) => row.conteoFisico === "");
  if (nextTail) {
    return nextTail;
  }
  return scopedRows.find((row) => row.conteoFisico === "");
}

function updateQuickActionButtons() {
  ui.diffOnlyBtn.textContent = state.onlyDifferences ? "Solo diferencias: ON" : "Solo diferencias: OFF";
  ui.diffOnlyBtn.classList.toggle("active-toggle", state.onlyDifferences);

  const scopedRows = getScopedRows({ ignoreDiffMode: true });
  const hasPending = scopedRows.some((row) => row.conteoFisico === "");
  ui.nextPendingBtn.disabled = !state.rows.length || !hasPending;
}

function renderControlsVisibility() {
  const loaded = state.rows.length > 0;
  ui.toggleControlsTab.hidden = !loaded;
  if (!loaded) {
    ui.controlsSection.hidden = false;
    return;
  }

  ui.controlsSection.hidden = state.controlsCollapsed;
  ui.controlsSection.classList.toggle("loaded-feedback", !state.controlsCollapsed && Boolean(state.collapseTimerId));
  ui.toggleControlsTab.textContent = state.controlsCollapsed
    ? "Carga [+]"
    : "Carga [-]";
}

function scheduleControlsAutoHide() {
  if (state.collapseTimerId) {
    clearTimeout(state.collapseTimerId);
  }

  state.collapseTimerId = setTimeout(() => {
    state.controlsCollapsed = true;
    renderControlsVisibility();
    ui.fileMeta.textContent = `Archivo cargado: ${state.fileName} | Registros: ${state.rows.length}`;
    state.collapseTimerId = null;
  }, 3000);
}

function runAutosave() {
  if (!state.rows.length || !state.isDirty) {
    return;
  }
  performSave({ silent: true });
}

function renderAutosaveStatus() {
  if (!ui.autosaveStatus) {
    return;
  }
  if (!state.rows.length) {
    ui.autosaveStatus.textContent = "Autoguardado inactivo.";
    return;
  }
  if (!state.lastSavedAt) {
    ui.autosaveStatus.textContent = state.isDirty
      ? `Cambios sin guardar. Autoguardado activo cada ${getAutosaveLabel()}.`
      : `Autoguardado activo cada ${getAutosaveLabel()}.`;
    return;
  }

  const seconds = Math.max(0, Math.floor((Date.now() - state.lastSavedAt.getTime()) / 1000));
  const suffix = state.isDirty ? " | Cambios pendientes" : "";
  ui.autosaveStatus.textContent = `Guardado hace ${formatElapsed(seconds)} (autoguardado ${getAutosaveLabel()})${suffix}`;
}

function getAutosaveLabel() {
  return `${Math.round(state.autosaveMs / 1000)}s`;
}

function formatElapsed(seconds) {
  if (seconds < 60) {
    return `${seconds}s`;
  }
  const min = Math.floor(seconds / 60);
  const sec = seconds % 60;
  return `${min}m ${String(sec).padStart(2, "0")}s`;
}

function getStatus(row) {
  if (row.conteoFisico === "") {
    return { text: "Pendiente", className: "status-pending" };
  }
  if (row.diferencia === 0) {
    return { text: "OK", className: "status-ok" };
  }
  if (row.diferencia > 0) {
    return { text: "Sobrante", className: "status-surplus" };
  }
  return { text: "Faltante", className: "status-missing" };
}

function getSobranteFaltante(row) {
  if (row.diferencia === null) {
    return { sobrante: null, faltante: null };
  }
  if (row.diferencia > 0) {
    return { sobrante: row.diferencia, faltante: 0 };
  }
  if (row.diferencia < 0) {
    return { sobrante: 0, faltante: Math.abs(row.diferencia) };
  }
  return { sobrante: 0, faltante: 0 };
}

function getFilterLabel(filter) {
  if (filter === "TODAS") {
    return "Todas las filas";
  }
  if (filter === "SIN_FILA") {
    return "Sin fila detectada";
  }
  return filter;
}

function extractFila(posicion) {
  const text = String(posicion || "").toUpperCase();
  const match = text.match(/F\s*([1-8])\b/) || text.match(/^([1-8])-/);
  if (!match) {
    return "SIN_FILA";
  }
  return `F${match[1]}`;
}

function getCell(row, index) {
  if (index === -1 || index == null) {
    return "";
  }
  return String(row[index] ?? "").trim();
}

function parseNumber(value) {
  const normalized = String(value ?? "")
    .replace(/,/g, "")
    .replace(/\s+/g, "")
    .trim();

  if (normalized === "") {
    return 0;
  }

  const number = Number(normalized);
  return Number.isFinite(number) ? number : 0;
}

function normalizeHeader(value) {
  return String(value ?? "")
    .toUpperCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^A-Z0-9 ]/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function normalizeText(value) {
  return String(value ?? "")
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .trim();
}

function getSearchScore(row, term) {
  const codigo = normalizeText(row.codigo);
  const nombre = normalizeText(row.nombre);
  const posicion = normalizeText(row.posicion);
  const combined = `${codigo} ${nombre} ${posicion}`;

  if (codigo === term) {
    return 100;
  }
  if (nombre === term) {
    return 95;
  }
  if (posicion === term) {
    return 92;
  }

  if (codigo.startsWith(term)) {
    return 80;
  }
  if (posicion.startsWith(term)) {
    return 72;
  }
  if (nombre.startsWith(term)) {
    return 70;
  }

  const wordExact = new RegExp(`(?:^|\\s)${escapeRegex(term)}(?:\\s|$)`);
  if (wordExact.test(nombre)) {
    return 60;
  }
  if (wordExact.test(posicion)) {
    return 56;
  }

  if (combined.includes(term)) {
    return 45;
  }
  return 0;
}

function escapeRegex(value) {
  return String(value).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function getRowStatusClass(row) {
  if (row.conteoFisico === "") {
    return "row-pending";
  }
  if (row.diferencia === 0) {
    return "row-ok";
  }
  if (row.diferencia > 0) {
    return "row-surplus";
  }
  return "row-missing";
}

function initializeFloatingDots() {
  if (!ui.floatingDots) {
    return;
  }

  const dotCount = Math.max(16, Math.min(28, Math.floor(window.innerWidth / 72)));
  const fragment = document.createDocumentFragment();
  for (let i = 0; i < dotCount; i += 1) {
    const dot = document.createElement("span");
    dot.className = "dot";
    dot.style.setProperty("--left", `${Math.random() * 100}%`);
    dot.style.setProperty("--size", `${Math.floor(Math.random() * 8) + 5}px`);
    dot.style.setProperty("--duration", `${Math.floor(Math.random() * 9) + 12}s`);
    dot.style.setProperty("--delay", `${Math.floor(Math.random() * 10)}s`);
    dot.style.setProperty("--drift", `${(Math.random() * 28 - 14).toFixed(1)}px`);
    dot.style.setProperty("--alpha", `${(Math.random() * 0.28 + 0.32).toFixed(2)}`);
    dot.style.setProperty("--hue", Math.random() > 0.75 ? "348" : "208");
    fragment.appendChild(dot);
  }

  ui.floatingDots.innerHTML = "";
  ui.floatingDots.appendChild(fragment);
}

function createStatCard(card) {
  const { label, value, tone = "neutral", hint = "", progress = null } = card;
  const item = document.createElement("div");
  item.className = `stat-item tone-${tone}`;

  const labelDiv = document.createElement("div");
  labelDiv.className = "label";
  labelDiv.textContent = label;

  const valueDiv = document.createElement("div");
  valueDiv.className = "value";
  valueDiv.textContent = value;

  item.appendChild(labelDiv);
  item.appendChild(valueDiv);

  if (hint) {
    const hintDiv = document.createElement("div");
    hintDiv.className = "stat-hint";
    hintDiv.textContent = hint;
    item.appendChild(hintDiv);
  }

  if (progress != null && Number.isFinite(progress)) {
    const track = document.createElement("div");
    track.className = "stat-progress";
    const fill = document.createElement("span");
    const width = Math.max(0, Math.min(100, progress));
    fill.style.width = `${width}%`;
    track.appendChild(fill);
    item.appendChild(track);
  }

  return item;
}

function createStatsRow(cards, className) {
  const row = document.createElement("div");
  row.className = className;
  cards.forEach((card) => row.appendChild(createStatCard(card)));
  return row;
}

function formatNumber(value) {
  const number = Number(value);
  if (!Number.isFinite(number)) {
    return "0";
  }
  return new Intl.NumberFormat("es-MX").format(number);
}

function getFileSignature() {
  if (!state.rows.length) {
    return "";
  }

  const first = state.rows[0];
  const last = state.rows[state.rows.length - 1];
  return [
    state.fileName,
    state.rows.length,
    first.codigo,
    first.posicion,
    last.codigo,
    last.posicion,
  ].join("|");
}

function buildTimestamp() {
  const now = new Date();
  const y = String(now.getFullYear());
  const m = String(now.getMonth() + 1).padStart(2, "0");
  const d = String(now.getDate()).padStart(2, "0");
  const hh = String(now.getHours()).padStart(2, "0");
  const mm = String(now.getMinutes()).padStart(2, "0");
  return `${y}${m}${d}_${hh}${mm}`;
}
