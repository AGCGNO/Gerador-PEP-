const pasteEtapas = document.getElementById("paste-etapas");
const btnPreview = document.getElementById("btn-preview-etapas");
const btnApplyAdjust = document.getElementById("btn-apply-adjust");
const btnCopy = document.getElementById("btn-copy-etapas");
const btnDownload = document.getElementById("btn-download-etapas");
const tableEtapas = document.getElementById("table-etapas");
const warnings = document.getElementById("warnings-etapas");
const adjustConcl = document.getElementById("adjust-concl");
const adjustEnc = document.getElementById("adjust-enc");
const adjustTotal = document.getElementById("adjust-total");
const adjustContrato = document.getElementById("adjust-contrato");
const adjustEmissao = document.getElementById("adjust-emissao");
const adjustTermo = document.getElementById("adjust-termo");
const adjustFornec = document.getElementById("adjust-fornec");
const filterId = document.getElementById("filter-id");

const OUTPUT_HEADERS = [
  "ID Etapa",
  "ID Projeto",
  "Marco",
  "Etapa",
  "Peso Etapa",
  "Fim Previsto",
  "Delta (dias)",
  "Peso Grupo",
  "Valor Total",
];

function normalizeKey(text) {
  return text
    .trim()
    .toUpperCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, " ");
}

function parseTsvAny(text) {
  const lines = text
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter((line) => line.length > 0);
  return lines.map((line) => line.split("\t").map((cell) => cell.trim()));
}

function parseNumber(value) {
  if (!value) return 0;
  const cleaned = value.replace(/\./g, "").replace(",", ".");
  const num = parseFloat(cleaned);
  return Number.isFinite(num) ? num : 0;
}

function parseMonthHeader(header) {
  const match = header.match(
    /(jan|fev|mar|abr|mai|jun|jul|ago|set|out|nov|dez)[/ -]?(\d{2,4})/i
  );
  if (!match) return null;
  const monthMap = {
    jan: 1,
    fev: 2,
    mar: 3,
    abr: 4,
    mai: 5,
    jun: 6,
    jul: 7,
    ago: 8,
    set: 9,
    out: 10,
    nov: 11,
    dez: 12,
  };
  const month = monthMap[match[1].toLowerCase()];
  const yearRaw = match[2];
  const year = yearRaw.length === 2 ? 2000 + parseInt(yearRaw, 10) : parseInt(yearRaw, 10);
  return { month, year };
}

function formatDate(date) {
  const dd = String(date.getDate()).padStart(2, "0");
  const mm = String(date.getMonth() + 1).padStart(2, "0");
  const yyyy = date.getFullYear();
  return `${dd}/${mm}/${yyyy}`;
}

function parseDate(value) {
  const parts = value.split("/");
  if (parts.length !== 3) return null;
  const [dd, mm, yyyy] = parts.map((p) => parseInt(p, 10));
  if (!dd || !mm || !yyyy) return null;
  return new Date(yyyy, mm - 1, dd);
}

function addDays(date, days) {
  const d = new Date(date);
  d.setDate(d.getDate() + days);
  return d;
}

function addMonths(date, months) {
  const d = new Date(date);
  d.setMonth(d.getMonth() + months);
  return d;
}

function diffDays(a, b) {
  return Math.round((a - b) / (1000 * 60 * 60 * 24));
}

function leadDaysByValue(total) {
  if (total <= 150000) return 35;
  if (total <= 500000) return 50;
  if (total <= 2000000) return 65;
  if (total <= 30000000) return 95;
  return 125;
}

function projectGroup(total) {
  if (total <= 500000) return { group: "Grupo 1", weight: "1" };
  if (total <= 2000000) return { group: "Grupo 2", weight: "2" };
  if (total <= 5000000) return { group: "Grupo 3", weight: "3" };
  if (total <= 20000000) return { group: "Grupo 4", weight: "5" };
  return { group: "Grupo 5", weight: "8" };
}

function detectColumns(headerRow) {
  const headers = headerRow.map((h) => normalizeKey(h));
  const find = (aliases) => {
    for (let i = 0; i < headers.length; i += 1) {
      const h = headers[i];
      if (aliases.some((alias) => h.includes(alias))) return i;
    }
    return -1;
  };

  const monthCols = [];
  const yearCols = [];
  headers.forEach((h, idx) => {
    if (parseMonthHeader(h)) monthCols.push(idx);
    if (/^20\d{2}$/.test(h)) yearCols.push(idx);
  });

  const idProjeto = find(["ID PROJETO"]);
  const idReal = find(["ID REAL", "ID SIGEP", "ID"]);
  const coletor = find(["COLETOR", "DEFINICAO", "DEFINIÇÃO", "PEP SIGEP"]);

  return {
    idProjeto,
    idReal,
    coletor,
    fimPrevisto: find(["FIM PREVISTO"]),
    total: find(["2026-2030", "2026 2030", "2026/2030"]),
    monthCols,
    yearCols,
  };
}

function renderTable(container, rows, highlightId = "") {
  if (!rows.length) {
    container.innerHTML = "<p class='hint'>Nada para exibir.</p>";
    return;
  }
  const table = document.createElement("table");
  const thead = document.createElement("thead");
  const headRow = document.createElement("tr");
  OUTPUT_HEADERS.forEach((h) => {
    const th = document.createElement("th");
    th.textContent = h;
    headRow.appendChild(th);
  });
  thead.appendChild(headRow);
  table.appendChild(thead);

  const tbody = document.createElement("tbody");
  rows.forEach((row) => {
    const tr = document.createElement("tr");
    if (highlightId && row[1] === highlightId) {
      tr.classList.add("row-highlight");
    }
    row.forEach((cell) => {
      const td = document.createElement("td");
      td.textContent = cell;
      td.setAttribute("contenteditable", "true");
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  });
  table.appendChild(tbody);
  container.innerHTML = "";
  container.appendChild(table);
}

function readTable(container) {
  const table = container.querySelector("table");
  if (!table) return [];
  return Array.from(table.querySelectorAll("tbody tr")).map((tr) =>
    Array.from(tr.querySelectorAll("td")).map((td) => td.textContent.trim())
  );
}

function rowType(row) {
  const marco = row[2] || "";
  const etapa = row[3] || "";
  if (marco.startsWith("1 - Levantamento")) return "levantamento";
  if (marco.startsWith("2 - Termo")) return "termo";
  if (marco.startsWith("3 - Emissão")) return "emissao";
  if (marco.startsWith("4 - Aprovação")) return "aprovacao";
  if (marco.startsWith("5 - Contratação")) return "contrato";
  if (marco.startsWith("6 - Fornecimento")) return "fornecimento";
  if (marco.startsWith("7 - Execução")) return "execucao";
  if (marco.startsWith("8 - Encerramento") && etapa.includes("Conclusão")) return "conclusao";
  if (marco.startsWith("8 - Encerramento") && etapa.includes("Encerramento")) return "encerramento";
  return "outro";
}

function computeDeltaByMarco(rows, dateIndex, deltaIndex) {
  const dateOf = (type) => {
    const r = rows.find((row) => rowType(row) === type);
    return r ? parseDate(r[dateIndex]) : null;
  };
  const fornecDates = rows
    .filter((row) => rowType(row) === "fornecimento")
    .map((row) => parseDate(row[dateIndex]))
    .filter(Boolean)
    .sort((a, b) => a - b);
  const firstFornec = fornecDates[0] || null;
  const lastFornec = fornecDates[fornecDates.length - 1] || null;

  const contrato = dateOf("contrato");
  const aprovacao = dateOf("aprovacao");
  const emissao = dateOf("emissao");
  const termo = dateOf("termo");
  const levantamento = dateOf("levantamento");
  const execucao = dateOf("execucao") || lastFornec;
  const conclusao = dateOf("conclusao");
  const encerramento = dateOf("encerramento");

  rows.forEach((row) => {
    const type = rowType(row);
    let delta = "";
    if (type === "fornecimento") delta = 0;
    if (type === "contrato" && firstFornec && contrato) delta = diffDays(firstFornec, contrato);
    if (type === "aprovacao" && contrato && aprovacao) delta = diffDays(contrato, aprovacao);
    if (type === "emissao" && aprovacao && emissao) delta = diffDays(aprovacao, emissao);
    if (type === "termo" && emissao && termo) delta = diffDays(emissao, termo);
    if (type === "levantamento" && termo && levantamento) delta = diffDays(termo, levantamento);
    if (type === "execucao" && lastFornec && execucao) delta = diffDays(execucao, lastFornec);
    if (type === "conclusao" && execucao && conclusao) delta = diffDays(conclusao, execucao);
    if (type === "encerramento" && conclusao && encerramento) delta = diffDays(encerramento, conclusao);
    row[deltaIndex] = delta;
  });
}

function buildEtapas(rows, header, overrides = {}) {
  const cols = detectColumns(header);
  const missing = [];
  if (cols.idProjeto === -1 && cols.idReal === -1 && cols.coletor === -1) {
    missing.push("ID Projeto/ID real/Coletor");
  }
  if (!cols.monthCols.length && !cols.yearCols.length && cols.fimPrevisto === -1) {
    missing.push("Meses/Anos ou FIM PREVISTO");
  }
  if (missing.length) {
    warnings.textContent = "Colunas não identificadas: " + missing.join(", ");
  } else {
    warnings.textContent = "";
  }

  const output = [];

  rows.forEach((row) => {
    const idProjeto =
      (cols.idProjeto >= 0 ? row[cols.idProjeto] : "") ||
      (cols.idReal >= 0 ? row[cols.idReal] : "") ||
      (cols.coletor >= 0 ? row[cols.coletor] : "");
    const idProjClean = (idProjeto || "").trim();
    if (!idProjClean) return;
    if (!/^COO\./i.test(idProjClean)) {
      warnings.textContent = "Existem IDs sem prefixo COO. (linhas ignoradas)";
      return;
    }

    const monthItems = cols.monthCols
      .map((idx) => {
        const info = parseMonthHeader(header[idx]);
        const value = parseNumber(row[idx]) * 1000;
        if (!info || value <= 0) return null;
        const date = new Date(info.year, info.month - 1, 26);
        return { date, value, kind: "month", year: info.year };
      })
      .filter(Boolean);

    const yearItemsRaw = cols.yearCols
      .map((idx) => {
        const year = parseInt(normalizeKey(header[idx]), 10);
        const value = parseNumber(row[idx]) * 1000;
        if (!Number.isFinite(year) || value <= 0) return null;
        const date = new Date(year, 0, 26);
        return { date, value, kind: "year", year };
      })
      .filter(Boolean);

    const monthYears = new Set(monthItems.map((m) => m.year));
    const yearItems = yearItemsRaw.filter((y) => !monthYears.has(y.year));

    let fornecimentoItems = [...monthItems, ...yearItems].sort(
      (a, b) => a.date - b.date
    );

    let fornecimentoDates = fornecimentoItems.map((m) => m.date);
    if (!fornecimentoDates.length && cols.fimPrevisto >= 0) {
      const raw = row[cols.fimPrevisto];
      const parts = raw ? raw.split(/[\/-]/) : [];
      if (parts.length === 3) {
        const d = new Date(`${parts[2]}-${parts[1]}-${parts[0]}`);
        fornecimentoDates = [d];
      }
    }

    if (!fornecimentoDates.length) return;

    let total = cols.total >= 0 ? parseNumber(row[cols.total]) * 1000 : 0;
    const override = overrides[idProjClean];
    if (override && override.total) {
      total = override.total;
    }
    const groupInfo = projectGroup(total);
    const leadDays = leadDaysByValue(total);

    const firstFornec = fornecimentoDates[0];
    const lastFornec = fornecimentoDates[fornecimentoDates.length - 1];

    const contratoDays = override?.contratoDays ?? 180;
    const emissaoDays = override?.emissaoDays ?? 15;
    const termoDays = override?.termoDays ?? 30;
    const fornecOffset = override?.fornecOffset ?? 0;

    if (fornecOffset && fornecimentoItems.length) {
      fornecimentoItems = fornecimentoItems.map((item) => ({
        ...item,
        date: addDays(item.date, fornecOffset),
      }));
      fornecimentoDates = fornecimentoItems.map((m) => m.date);
    }

    const contrato = addDays(firstFornec, -contratoDays);
    const aprovacao = addDays(contrato, -leadDays);
    const emissao = addDays(aprovacao, -emissaoDays);
    const termo = addDays(emissao, -termoDays);
    const levantamento = addDays(termo, -30);
    const execucao = lastFornec;
    const conclOffset = override?.conclOffset || 0;
    const encOffset = override?.encOffset || 0;
    const conclusao = addDays(addDays(execucao, 90), conclOffset);
    const encerramento = addDays(addDays(conclusao, 90), encOffset);

    let etapaCounter = 1;
    const etapasBase = [
      ["1 - Levantamento de Dados", "Levantamento de Dados", 5, levantamento],
      ["2 - Termo de Referência", "Termo de Referência", 5, termo],
      ["3 - Emissão da RC", "Emissão da RC", 5, emissao],
      ["4 - Aprovação da RC", "Aprovação da RC", 5, aprovacao],
      ["5 - Contratação", "Assinatura do Contrato", 5, contrato],
    ];

    const projectRows = [];
    etapasBase.forEach((e) => {
      const idEtapa = `${idProjClean}-${String(etapaCounter).padStart(2, "0")}`;
      etapaCounter += 1;
      projectRows.push([
        idEtapa,
        idProjClean,
        e[0],
        e[1],
        `${e[2].toFixed(2)}%`,
        formatDate(e[3]),
        "",
        groupInfo.weight,
        total.toLocaleString("pt-BR", { minimumFractionDigits: 2 }),
      ]);
    });

    const fornecPeso = 35 / fornecimentoDates.length;
    const monthByYear = {};
    fornecimentoItems.forEach((item) => {
      if (item.kind === "month") {
        monthByYear[item.year] = (monthByYear[item.year] || 0) + 1;
      }
    });
    const monthIndexByYear = {};

    fornecimentoItems.forEach((item) => {
      let etapaLabel = "Fornecimento";
      if (item.kind === "year") {
        etapaLabel = `Fornecimento ${item.year}`;
      } else {
        const count = monthByYear[item.year] || 0;
        if (count > 1) {
          const idx = (monthIndexByYear[item.year] || 0) + 1;
          monthIndexByYear[item.year] = idx;
          etapaLabel = `Fornecimento ${idx}ª Remessa`;
        } else {
          etapaLabel = `Fornecimento ${item.year}`;
        }
      }

      const idEtapa = `${idProjClean}-${String(etapaCounter).padStart(2, "0")}`;
      etapaCounter += 1;
      projectRows.push([
        idEtapa,
        idProjClean,
        "6 - Fornecimento",
        etapaLabel,
        `${fornecPeso.toFixed(2)}%`,
        formatDate(item.date),
        "",
        groupInfo.weight,
        total.toLocaleString("pt-BR", { minimumFractionDigits: 2 }),
      ]);
    });
    const execCount =
      groupInfo.group === "Grupo 4" || groupInfo.group === "Grupo 5"
        ? Math.min(fornecimentoItems.length, 10)
        : 1;
    const execPeso = 35 / execCount;
    for (let i = 0; i < execCount; i += 1) {
      projectRows.push([
        `${idProjClean}-${String(etapaCounter).padStart(2, "0")}`,
        idProjClean,
        "7 - Execução",
        execCount > 1 ? `Execução ${i + 1}` : "Execução",
        `${execPeso.toFixed(2)}%`,
        formatDate(execucao),
        "",
        groupInfo.weight,
        total.toLocaleString("pt-BR", { minimumFractionDigits: 2 }),
      ]);
      etapaCounter += 1;
    }
    projectRows.push([
      `${idProjClean}-${String(etapaCounter).padStart(2, "0")}`,
      idProjClean,
      "8 - Encerramento",
      "Conclusão do Projeto",
      "5.00%",
      formatDate(conclusao),
      "",
      groupInfo.weight,
      total.toLocaleString("pt-BR", { minimumFractionDigits: 2 }),
    ]);
    etapaCounter += 1;
    projectRows.push([
      `${idProjClean}-${String(etapaCounter).padStart(2, "0")}`,
      idProjClean,
      "8 - Encerramento",
      "Encerramento Técnico PEP",
      "0.00%",
      formatDate(encerramento),
      "",
      groupInfo.weight,
      total.toLocaleString("pt-BR", { minimumFractionDigits: 2 }),
    ]);

    const dateIndex = OUTPUT_HEADERS.indexOf("Fim Previsto");
    const deltaIndex = OUTPUT_HEADERS.indexOf("Delta (dias)");
    computeDeltaByMarco(projectRows, dateIndex, deltaIndex);
    projectRows.forEach((r) => output.push(r));
  });

  return output;
}

btnPreview.addEventListener("click", () => {
  const rows = parseTsvAny(pasteEtapas.value);
  if (!rows.length) return;
  const header = rows[0];
  const data = rows.slice(1);
  window.__etapasHeader = header;
  window.__etapasData = data;
  window.__etapasOverrides = window.__etapasOverrides || {};
  const output = buildEtapas(data, header, window.__etapasOverrides);
  renderTable(tableEtapas, output);
  populateFilter(output);
});

btnCopy.addEventListener("click", async () => {
  const table = tableEtapas.querySelector("table");
  if (!table) return;
  const rows = [OUTPUT_HEADERS];
  table.querySelectorAll("tbody tr").forEach((tr) => {
    const cells = Array.from(tr.querySelectorAll("td")).map((td) =>
      td.textContent.trim()
    );
    rows.push(cells);
  });
  const tsv = rows.map((row) => row.join("\t")).join("\n");
  await navigator.clipboard.writeText(tsv);
});

btnDownload.addEventListener("click", () => {
  const table = tableEtapas.querySelector("table");
  if (!table) return;
  const rows = [OUTPUT_HEADERS];
  table.querySelectorAll("tbody tr").forEach((tr) => {
    const cells = Array.from(tr.querySelectorAll("td")).map((td) =>
      td.textContent.trim()
    );
    rows.push(cells);
  });

  const worksheet = XLSX.utils.aoa_to_sheet(rows);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "Etapas");
  XLSX.writeFile(workbook, "etapas.xlsx");
});

renderTable(tableEtapas, []);
updateAdjustState();

function populateFilter(rows) {
  const ids = new Set(rows.map((r) => r[1]).filter(Boolean));
  const current = filterId.value;
  filterId.innerHTML = '<option value="">Todos</option>';
  Array.from(ids)
    .sort()
    .forEach((id) => {
      const opt = document.createElement("option");
      opt.value = id;
      opt.textContent = id;
      filterId.appendChild(opt);
    });
  if (current) filterId.value = current;
  updateAdjustState();
}

filterId.addEventListener("change", () => {
  const rows = readTable(tableEtapas);
  const id = filterId.value;
  updateAdjustState();
  if (!id) {
    renderTable(tableEtapas, rows);
    return;
  }
  const filtered = rows.filter((r) => r[1] === id);
  const rest = rows.filter((r) => r[1] !== id);
  renderTable(tableEtapas, [...filtered, ...rest], id);
});

function updateAdjustState() {
  const enabled = !!filterId.value;
  [adjustTotal, adjustContrato, adjustEmissao, adjustTermo, adjustFornec, adjustConcl, adjustEnc].forEach(
    (el) => {
      if (el) el.disabled = !enabled;
    }
  );
  if (btnApplyAdjust) btnApplyAdjust.disabled = !enabled;
}

btnApplyAdjust.addEventListener("click", () => {
  if (!window.__etapasHeader || !window.__etapasData) return;
  const id = filterId.value;
  if (!id) return;
  window.__etapasOverrides = window.__etapasOverrides || {};
  window.__etapasOverrides[id] = {
    total: parseNumber(adjustTotal.value || "") || undefined,
    contratoDays: parseInt(adjustContrato.value || "180", 10) || 180,
    emissaoDays: parseInt(adjustEmissao.value || "15", 10) || 15,
    termoDays: parseInt(adjustTermo.value || "30", 10) || 30,
    fornecOffset: parseInt(adjustFornec.value || "0", 10) || 0,
    conclOffset: parseInt(adjustConcl.value || "0", 10) || 0,
    encOffset: parseInt(adjustEnc.value || "0", 10) || 0,
  };
  const output = buildEtapas(
    window.__etapasData,
    window.__etapasHeader,
    window.__etapasOverrides
  );
  const filtered = output.filter((r) => r[1] === id);
  const rest = output.filter((r) => r[1] !== id);
  renderTable(tableEtapas, [...filtered, ...rest], id);
  populateFilter(output);
});

tableEtapas.addEventListener("focusin", (event) => {
  const cell = event.target.closest("td");
  if (!cell) return;
  cell.dataset.oldValue = cell.textContent.trim();
});

tableEtapas.addEventListener("focusout", (event) => {
  const cell = event.target.closest("td");
  if (!cell) return;
  const oldValue = cell.dataset.oldValue || "";
  const newValue = cell.textContent.trim();
  if (oldValue === newValue) return;

  const row = cell.parentElement;
  const cells = Array.from(row.querySelectorAll("td"));
  const dateIndex = OUTPUT_HEADERS.indexOf("Fim Previsto");
  const deltaIndex = OUTPUT_HEADERS.indexOf("Delta (dias)");
  const cellIndex = cells.indexOf(cell);
  if (cellIndex !== dateIndex && cellIndex !== deltaIndex) return;

  const projectId = cells[1]?.textContent.trim();
  if (!projectId) return;

  const table = tableEtapas.querySelector("table");
  if (!table) return;
  const projectRows = Array.from(table.querySelectorAll("tbody tr")).filter(
    (tr) => tr.querySelectorAll("td")[1]?.textContent.trim() === projectId
  );
  const indexInProject = projectRows.indexOf(row);
  if (indexInProject === -1) return;

  let deltaDays = 0;
  if (cellIndex === dateIndex) {
    const oldDate = parseDate(oldValue);
    const newDate = parseDate(newValue);
    if (!oldDate || !newDate) return;
    deltaDays = Math.round((newDate - oldDate) / (1000 * 60 * 60 * 24));
  } else if (cellIndex === deltaIndex) {
    const newDelta = parseInt(newValue, 10);
    if (!Number.isFinite(newDelta)) return;
    const rowsData = projectRows.map((tr) =>
      Array.from(tr.querySelectorAll("td")).map((td) => td.textContent.trim())
    );
    const rowData = rowsData[indexInProject];
    const type = rowType(rowData);
    const dates = rowsData.map((r) => parseDate(r[dateIndex]));
    const findDate = (t) => {
      const idx = rowsData.findIndex((r) => rowType(r) === t);
      return idx >= 0 ? dates[idx] : null;
    };
    const fornecDates = rowsData
      .filter((r) => rowType(r) === "fornecimento")
      .map((r) => parseDate(r[dateIndex]))
      .filter(Boolean)
      .sort((a, b) => a - b);
    const firstFornec = fornecDates[0] || null;
    const lastFornec = fornecDates[fornecDates.length - 1] || null;
    const contrato = findDate("contrato");
    const aprovacao = findDate("aprovacao");
    const emissao = findDate("emissao");
    const termo = findDate("termo");
    const execucao = findDate("execucao") || lastFornec;
    const conclusao = findDate("conclusao");

    let desiredDate = null;
    if (type === "contrato" && firstFornec) desiredDate = addDays(firstFornec, -newDelta);
    if (type === "aprovacao" && contrato) desiredDate = addDays(contrato, -newDelta);
    if (type === "emissao" && aprovacao) desiredDate = addDays(aprovacao, -newDelta);
    if (type === "termo" && emissao) desiredDate = addDays(emissao, -newDelta);
    if (type === "levantamento" && termo) desiredDate = addDays(termo, -newDelta);
    if (type === "execucao" && lastFornec) desiredDate = addDays(lastFornec, newDelta);
    if (type === "conclusao" && execucao) desiredDate = addDays(execucao, newDelta);
    if (type === "encerramento" && conclusao) desiredDate = addDays(conclusao, newDelta);
    if (!desiredDate) return;
    const currentDate = parseDate(
      row.querySelectorAll("td")[dateIndex].textContent.trim()
    );
    if (!currentDate) return;
    deltaDays = Math.round(
      (desiredDate - currentDate) / (1000 * 60 * 60 * 24)
    );
  }

  if (!deltaDays) return;

  // Aplica o deslocamento da linha atual para frente (ordem da tabela)
  projectRows.slice(indexInProject).forEach((tr) => {
    const tds = tr.querySelectorAll("td");
    const dateCell = tds[dateIndex];
    const current = parseDate(dateCell.textContent.trim());
    if (!current) return;
    dateCell.textContent = formatDate(addDays(current, deltaDays));
  });

  // Recalcula deltas por marco no projeto
  const rowsData = projectRows.map((tr) =>
    Array.from(tr.querySelectorAll("td")).map((td) => td.textContent.trim())
  );
  computeDeltaByMarco(rowsData, dateIndex, deltaIndex);
  projectRows.forEach((tr, idx) => {
    const tds = tr.querySelectorAll("td");
    tds[deltaIndex].textContent = rowsData[idx][deltaIndex];
  });
});
