const sapTextarea = document.getElementById("paste-sap");
const fileCapex = document.getElementById("file-capex");

const tableOutput = document.getElementById("table-output");
const warningsBox = document.getElementById("warnings");

const btnGenerate = document.getElementById("btn-generate");
const btnCopy = document.getElementById("btn-copy");
const btnDownload = document.getElementById("btn-download");
const btnRecalc = document.getElementById("btn-recalc");
const btnAddManual = document.getElementById("btn-add-manual");
const filterPep = document.getElementById("filter-pep");

const recalcUsina = document.getElementById("recalc-usina");
const recalcStart = document.getElementById("recalc-start");
const manualUsina = document.getElementById("manual-usina");
const manualColetor = document.getElementById("manual-coletor");
const manualIdReal = document.getElementById("manual-idreal");
const manualDesc = document.getElementById("manual-desc");
const manualObjeto = document.getElementById("manual-objeto");
let manualRows = [];

// Mapeamento obrigatório por usina.
const USINAS = {
  BALBINA: { prefix: "0161", lucro: "NO10101003", start: 1 },
  "COARACY NUNES": { prefix: "0159", lucro: "N010102001", start: 3 },
  "CURUA-UNA": { prefix: "0160", lucro: "N010103001", start: 4 },
  SAMUEL: { prefix: "0160", lucro: "N010104001", start: 2 },
  TUCURI: { prefix: "0162", lucro: "N010105001", start: 14 },
};

const OUTPUT_HEADERS = [
  "Usina",
  "Nível",
  "Elemento PEP",
  "Denominação",
  "ID real",
  "Tipo",
  "Pri",
  "Centro de lucro",
];

// Metadados para leitura das áreas de colagem.
const DATASETS = {
  projects: {
    cols: 5,
    headers: [
      "Coletor (Carga)",
      "ID real",
      "Descrição",
      "Local/Usina",
      "Objeto/Especificação",
    ],
  },
};

function normalizeKey(text) {
  return text
    .trim()
    .toUpperCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, " ");
}

function resolveUsinaKey(raw) {
  const key = normalizeKey(raw);
  if (USINAS[key]) return key;
  const candidates = Object.keys(USINAS);
  const found = candidates.find((candidate) => key.includes(candidate));
  if (found) return found;
  if (key.includes("TUCURUI")) return "TUCURI";
  return found || "";
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

  return {
    coletor: find(["COLETOR", "DEFINICAO", "DEFINIÇÃO", "PEP SIGEP"]),
    idReal: find(["ID REAL", "ID SIGEP", "ID"]),
    desc: find(["DESCRICAO", "DESCRIÇÃO", "DENOMINACAO", "DENOMINAÇÃO"]),
    objeto: find(["OBJETO", "ESPECIFICACAO", "ESPECIFICAÇÃO"]),
    usina: find(["LOCAL", "USINA", "INSTALACAO", "INSTALAÇÃO"]),
  };
}

function findHeaderIndex(rows) {
  for (let i = 0; i < rows.length; i += 1) {
    const cols = detectColumns(rows[i] || []);
    const hits = [cols.coletor, cols.idReal, cols.desc, cols.usina].filter(
      (v) => v !== -1
    ).length;
    if (hits >= 3) {
      // Confirma se há dados válidos abaixo do cabeçalho
      const sample = rows.slice(i + 1, i + 20);
      const hasData = sample.some((r) => {
        const coletor = cols.coletor >= 0 ? (r[cols.coletor] || "") : "";
        return /NGHI|COO/i.test(coletor);
      });
      if (hasData) return i;
    }
  }
  return -1;
}

function displayUsina(usinaKey) {
  return usinaKey || "";
}

function padSeq(seq) {
  return String(seq).padStart(2, "0");
}

// Converte texto colado do Excel em matriz de colunas fixas.
function parseTsv(text, cols) {
  const lines = text
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter((line) => line.length > 0);

  return lines.map((line) => {
    const parts = line.split("\t");
    if (parts.length > cols) {
      const head = parts.slice(0, cols - 1);
      const tail = parts.slice(cols - 1).join(" ");
      return [...head, tail].map((cell) => cell.trim());
    }
    if (parts.length < cols) {
      while (parts.length < cols) parts.push("");
    }
    return parts.map((cell) => cell.trim());
  });
}

function parseTsvAny(text) {
  const lines = text
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter((line) => line.length > 0);
  return lines.map((line) => line.split("\t").map((cell) => cell.trim()));
}

// Renderiza tabela editável.
function renderTable(container, headers, rows) {
  if (!rows || rows.length === 0) {
    container.innerHTML = "<p class='hint'>Sem dados.</p>";
    return;
  }

  const table = document.createElement("table");
  const thead = document.createElement("thead");
  const headRow = document.createElement("tr");
  headers.forEach((header) => {
    const th = document.createElement("th");
    th.textContent = header;
    headRow.appendChild(th);
  });
  thead.appendChild(headRow);
  table.appendChild(thead);

  const tbody = document.createElement("tbody");
  rows.forEach((row) => {
    const tr = document.createElement("tr");
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

function readTable(container, cols) {
  const table = container.querySelector("table");
  if (!table) return [];
  const rows = Array.from(table.querySelectorAll("tbody tr"));
  return rows.map((tr) => {
    const cells = Array.from(tr.querySelectorAll("td")).map((td) =>
      td.textContent.trim()
    );
    while (cells.length < cols) cells.push("");
    return cells.slice(0, cols);
  });
}

// Aplica abreviações e limita a 40 caracteres.
function shortenText(text) {
  let value = text.trim().toUpperCase();
  const rules = [
    [/MONITORAMENTO/g, "MONIT."],
    [/TRANSFORMADOR(?:ES)?/g, "TRAFO"],
    [/MODERNIZACAO|MODERNIZAÇÃO/g, "MODERN."],
    [/REVITALIZACAO|REVITALIZAÇÃO/g, "REVIT."],
    [/REGULACAO|REGULAÇÃO/g, "REGUL."],
    [/MANUTENCAO|MANUTENÇÃO/g, "MANUT."],
    [/SUBESTACAO|SUBESTAÇÃO/g, "SUBEST."],
    [/SUBSTITUICAO|SUBSTITUIÇÃO/g, "SUBST."],
    [/AQUISICAO|AQUISIÇÃO/g, "AQ."],
    [/IMPLANTACAO|IMPLANTAÇÃO/g, "IMPLANT."],
    [/FORNECIMENTO|FORNCEIMENTO/g, "FORN."],
    [/CONTRATACAO|CONTRATAÇÃO/g, "CONTRAT."],
    [/ALTERACAO|ALTERAÇÃO|AÇTERACA-?P/g, "ALT."],
    [/INSTALACAO|INSTALAÇÃO/g, "INST."],
    [/GERADOR(?:ES)?/g, "GER."],
    [/TURBINA(?:S)?/g, "TURB."],
    [/PROTECAO|PROTEÇÃO/g, "PROT."],
    [/COMUNICACAO|COMUNICAÇÃO/g, "COM."],
    [/TELECOMUNICACAO|TELECOMUNICAÇÃO/g, "TELECOM."],
    [/DISTRIBUICAO|DISTRIBUIÇÃO/g, "DIST."],
    [/SECCIONAMENTO/g, "SECC."],
    [/EQUIPAMENTO(?:S)?/g, "EQP."],
    [/ELETRICO|ELÉTRICO|ELETRICA|ELÉTRICA/g, "ELETR."],
    [/AUTOMATIZACAO|AUTOMATIZAÇÃO/g, "AUTOM."],
    [/OPERACAO|OPERAÇÃO/g, "OPER."],
    [/SUPERVISAO|SUPERVISÃO/g, "SUPERV."],
    [/CONTROLE/g, "CTRL."],
    [/REFORMA/g, "REF."],
    [/ADEQUACAO|ADEQUAÇÃO/g, "ADEQ."],
    [/MELHORIA/g, "MELH."],
    [/SISTEMA/g, "SIST."],
  ];
  rules.forEach(([regex, repl]) => {
    value = value.replace(regex, repl);
  });
  value = value.replace(/\s+/g, " ").trim();
  if (value.length > 40) value = value.slice(0, 40);
  return value;
}

function buildDenom403(idReal, objeto) {
  const obj = objeto ? shortenText(objeto) : "";
  const base = [idReal, obj].filter((item) => item && item.trim()).join(" - ");
  return base.slice(0, 40);
}

// Gera a estrutura completa de PEPs a partir das regras.
function buildOutput(projectRows) {
  const warnings = [];
  const currentSeqByPrefix = {};

  const output = [];
  const projectsMeta = [];

  projectRows.forEach((row, index) => {
    const [coletorRaw, idRealRaw, descRaw, usinaRaw, objetoRaw] = row;
    if (!coletorRaw) return;
    if (!usinaRaw && !idRealRaw && !descRaw) return;

    const usinaKey = resolveUsinaKey(usinaRaw);
    const mapping = USINAS[usinaKey];
    if (!mapping) {
      warnings.push(`Linha ${index + 1}: usina não reconhecida (${usinaRaw}).`);
      return;
    }

    const { prefix, lucro, start } = mapping;
    if (currentSeqByPrefix[prefix] == null) {
      currentSeqByPrefix[prefix] = start;
    }
    if (currentSeqByPrefix[prefix] < start) {
      currentSeqByPrefix[prefix] = start;
    }

    const seq = currentSeqByPrefix[prefix];
    currentSeqByPrefix[prefix] += 1;
    const base = `NGHI.${prefix}.${padSeq(seq)}`;
    const idReal = (idRealRaw || "").trim();
    const desc = shortenText(descRaw || "");
    const denom403 = buildDenom403(idReal, objetoRaw || "");

    const usinaDisplay = displayUsina(usinaKey);
    projectsMeta.push({
      usinaKey,
      usinaRaw: usinaDisplay,
      prefix,
      seq,
      idReal,
      desc,
      lucro,
      denom403,
    });

    output.push([usinaDisplay, "2", base, desc, idReal, "G6", "7", lucro]);
    output.push([
      usinaDisplay,
      "3",
      `${base}.00001`,
      desc,
      idReal,
      "G6",
      "7",
      lucro,
    ]);
    output.push([
      usinaDisplay,
      "4.001",
      `${base}.00001.001`,
      idReal ? `${idReal} - CUSTO COMUM` : "",
      idReal,
      "G6",
      "3",
      lucro,
    ]);
    output.push([
      usinaDisplay,
      "4.002",
      `${base}.00001.002`,
      idReal ? `${idReal} - SERVIÇO` : "",
      idReal,
      "G6",
      "3",
      lucro,
    ]);
    output.push([
      usinaDisplay,
      "4.003",
      `${base}.00001.003`,
      denom403,
      idReal,
      "G6",
      "2",
      lucro,
    ]);
  });

  return { output, warnings, projectsMeta };
}

function rebuildFromMeta(meta, prioritizeKey = "") {
  const ordered = prioritizeKey
    ? [
        ...meta.filter((item) => item.usinaKey === prioritizeKey),
        ...meta.filter((item) => item.usinaKey !== prioritizeKey),
      ]
    : meta;
  const output = [];
  ordered.forEach((item) => {
    const base = `NGHI.${item.prefix}.${padSeq(item.seq)}`;
    output.push([
      item.usinaRaw,
      "2",
      base,
      item.desc,
      item.idReal,
      "G6",
      "7",
      item.lucro,
    ]);
    output.push([
      item.usinaRaw,
      "3",
      `${base}.00001`,
      item.desc,
      item.idReal,
      "G6",
      "7",
      item.lucro,
    ]);
    output.push([
      item.usinaRaw,
      "4.001",
      `${base}.00001.001`,
      item.idReal ? `${item.idReal} - CUSTO COMUM` : "",
      item.idReal,
      "G6",
      "3",
      item.lucro,
    ]);
    output.push([
      item.usinaRaw,
      "4.002",
      `${base}.00001.002`,
      item.idReal ? `${item.idReal} - SERVIÇO` : "",
      item.idReal,
      "G6",
      "3",
      item.lucro,
    ]);
    output.push([
      item.usinaRaw,
      "4.003",
      `${base}.00001.003`,
      item.denom403,
      item.idReal,
      "G6",
      "2",
      item.lucro,
    ]);
  });
  return output;
}

function renderOutputTable(rows, highlightBase = "") {
  if (!rows.length) {
    tableOutput.innerHTML = "<p class='hint'>Nada para exibir.</p>";
    return;
  }

  const table = document.createElement("table");
  const thead = document.createElement("thead");
  const headRow = document.createElement("tr");
  OUTPUT_HEADERS.forEach((header) => {
    const th = document.createElement("th");
    th.textContent = header;
    headRow.appendChild(th);
  });
  thead.appendChild(headRow);
  table.appendChild(thead);

  const tbody = document.createElement("tbody");
  rows.forEach((row) => {
    const tr = document.createElement("tr");
    if (highlightBase && row[2]?.startsWith(highlightBase)) {
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

  tableOutput.innerHTML = "";
  tableOutput.appendChild(table);
}

btnGenerate.addEventListener("click", () => {
  const sapRows = parseTsvAny(sapTextarea.value);
  if (!sapRows.length && manualRows.length === 0) return;
  let projectRows = [];
  if (sapRows.length) {
    const headerIndex = findHeaderIndex(sapRows);
    if (headerIndex === -1) {
      warningsBox.textContent = "Cabeçalho não encontrado.";
      return;
    }
    const header = sapRows[headerIndex];
    const data = sapRows.slice(headerIndex + 1);
    const cols = detectColumns(header);
    const missing = [];
    if (cols.coletor === -1) missing.push("Coletor");
    if (cols.idReal === -1) missing.push("ID real");
    if (cols.desc === -1) missing.push("Descrição");
    if (cols.usina === -1) missing.push("Usina/Local");
    if (missing.length) {
      warningsBox.textContent =
        "Colunas não identificadas: " + missing.join(", ");
    } else {
      warningsBox.textContent = "";
    }
    projectRows = data
      .map((row) => [
        cols.coletor >= 0 ? row[cols.coletor] || "" : "",
        cols.idReal >= 0 ? row[cols.idReal] || "" : "",
        cols.desc >= 0 ? row[cols.desc] || "" : "",
        cols.usina >= 0 ? row[cols.usina] || "" : "",
        cols.objeto >= 0 ? row[cols.objeto] || "" : "",
      ])
      .filter((row) => {
        if (!row.some((cell) => cell && cell.trim())) return false;
        if (!row[0] || !row[0].trim()) return false;
        return true;
      });
  }

  projectRows = [...projectRows, ...manualRows];
  const { output, warnings, projectsMeta } = buildOutput(projectRows);
  warningsBox.textContent = warnings.length ? warnings.join(" ") : "";
  window.__pepOutput = output;
  renderOutputTable(output);
  window.__pepMeta = projectsMeta;
  populatePepFilter(output);
});

btnRecalc.addEventListener("click", () => {
  const usinaKey = normalizeKey(recalcUsina.value || "");
  const start = parseInt(recalcStart.value, 10);
  if (!window.__pepMeta || !usinaKey || !Number.isFinite(start)) return;
  let counter = start;
  window.__pepMeta.forEach((item) => {
    if (item.usinaKey === usinaKey) {
      item.seq = counter;
      counter += 1;
    }
  });
  const updated = rebuildFromMeta(window.__pepMeta, usinaKey);
  const target = normalizeKey(recalcUsina.value || "");
  const sorted = updated.slice().sort((a, b) => {
    const aMatch = resolveUsinaKey(a[0]) === target ? 0 : 1;
    const bMatch = resolveUsinaKey(b[0]) === target ? 0 : 1;
    if (aMatch !== bMatch) return aMatch - bMatch;
    return 0;
  });
  window.__pepOutput = sorted;
  renderOutputTable(sorted, filterPep.value);
});

recalcUsina.addEventListener("change", () => {
  if (!window.__pepMeta) return;
  const target = normalizeKey(recalcUsina.value || "");
  if (!target) return;
  const updated = rebuildFromMeta(window.__pepMeta);
  const sorted = updated.slice().sort((a, b) => {
    const aMatch = resolveUsinaKey(a[0]) === target ? 0 : 1;
    const bMatch = resolveUsinaKey(b[0]) === target ? 0 : 1;
    if (aMatch !== bMatch) return aMatch - bMatch;
    return 0;
  });
  window.__pepOutput = sorted;
  renderOutputTable(sorted, filterPep.value);
});


btnAddManual.addEventListener("click", () => {
  const row = [
    manualColetor.value.trim(),
    manualIdReal.value.trim(),
    manualDesc.value.trim(),
    manualUsina.value.trim(),
    manualObjeto.value.trim(),
  ];
  if (!row[0] || !row[3]) return;
  manualRows.push(row);
  manualColetor.value = "";
  manualIdReal.value = "";
  manualDesc.value = "";
  manualObjeto.value = "";
});

btnCopy.addEventListener("click", async () => {
  const table = tableOutput.querySelector("table");
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
  const table = tableOutput.querySelector("table");
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
  XLSX.utils.book_append_sheet(workbook, worksheet, "PEPs");
  XLSX.writeFile(workbook, "peps.xlsx");
});

renderOutputTable([]);

function populatePepFilter(rows) {
  const set = new Set(
    rows.filter((r) => r[1] === "2").map((r) => r[2]).filter(Boolean)
  );
  filterPep.innerHTML = '<option value="">Todos</option>';
  Array.from(set)
    .sort()
    .forEach((pep) => {
      const opt = document.createElement("option");
      opt.value = pep;
      opt.textContent = pep;
      filterPep.appendChild(opt);
    });
}

filterPep.addEventListener("change", () => {
  if (!window.__pepOutput) return;
  const base = filterPep.value;
  if (!base) {
    renderOutputTable(window.__pepOutput);
    return;
  }
  const selected = window.__pepOutput.filter((r) => r[2]?.startsWith(base));
  const rest = window.__pepOutput.filter((r) => !r[2]?.startsWith(base));
  renderOutputTable([...selected, ...rest], base);
});

function aoaToTsv(data) {
  return data.map((row) => row.join("\t")).join("\n");
}

fileCapex.addEventListener("change", async () => {
  const file = fileCapex.files?.[0];
  if (!file) return;
  let data = [];
  if (file.name.toLowerCase().endsWith(".csv")) {
    const text = await file.text();
    const wb = XLSX.read(text, { type: "string" });
    const sheet = wb.Sheets[wb.SheetNames[0]];
    data = XLSX.utils.sheet_to_json(sheet, { header: 1 });
  } else {
    const buf = await file.arrayBuffer();
    const wb = XLSX.read(buf, { type: "array" });
    const sheet = wb.Sheets[wb.SheetNames[0]];
    data = XLSX.utils.sheet_to_json(sheet, { header: 1 });
  }
  let rows = data;
  const headerIndex = findHeaderIndex(rows);
  if (headerIndex !== -1) {
    rows = rows.slice(headerIndex);
  }
  const tsv = aoaToTsv(rows);
  sapTextarea.value = tsv;
  localStorage.setItem("capex_tsv", tsv);
});

const savedCapex = localStorage.getItem("capex_tsv");
if (savedCapex && !sapTextarea.value) {
  sapTextarea.value = savedCapex;
}
