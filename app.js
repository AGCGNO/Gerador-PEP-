const sapTextarea = document.getElementById("paste-sap");

const tableOutput = document.getElementById("table-output");
const warningsBox = document.getElementById("warnings");
const warningsGenerate = document.getElementById("warnings-generate");

const btnGenerate = document.getElementById("btn-generate");
const btnCopy = document.getElementById("btn-copy");
const btnDownload = document.getElementById("btn-download");
const btnTemplatePep = document.getElementById("btn-template-pep");
const btnRecalc = document.getElementById("btn-recalc");
const filterPep = document.getElementById("filter-pep");

const recalcUsina = document.getElementById("recalc-usina");
const recalcStart = document.getElementById("recalc-start");
const manualUsina = document.getElementById("manual-usina");
const manualUsinaCustom = document.getElementById("manual-usina-custom");
const manualAdvanced = document.getElementById("manual-advanced");
const manualColetor = document.getElementById("manual-coletor");
const manualIdReal = document.getElementById("manual-idreal");
const manualDesc = document.getElementById("manual-desc");
const manualObjeto = document.getElementById("manual-objeto");
const manualCentroCusto = document.getElementById("manual-centro-custo");
const manualCentroLucro = document.getElementById("manual-centro-lucro");
const manualLocal = document.getElementById("manual-local");
const manualPerfil = document.getElementById("manual-perfil");
const manualEmpresa = document.getElementById("manual-empresa");

// Mapeamento obrigatório por usina.
const USINAS = {
  BALBINA: {
    prefix: "0158",
    centroCusto: "NO10101003",
    lucro: "N010101001",
    localInstalacao: "N-U-UHBA",
    start: 4,
  },
  "COARACY NUNES": {
    prefix: "0159",
    centroCusto: "NO10102003",
    lucro: "N010102001",
    localInstalacao: "N-U-UHCN",
    start: 3,
  },
  "CURUA-UNA": {
    prefix: "0160",
    centroCusto: "NO10103003",
    lucro: "N010103001",
    localInstalacao: "N-U-UHCU",
    start: 4,
  },
  SAMUEL: {
    prefix: "0161",
    centroCusto: "NO10104003",
    lucro: "N010104001",
    localInstalacao: "N-U-UHSU",
    start: 3,
  },
  TUCURI: {
    prefix: "0162",
    centroCusto: "NO10105003",
    lucro: "N010105001",
    localInstalacao: "N-U-UHTU",
    start: 7,
  },
};

const OUTPUT_HEADERS = [
  "Usina",
  "Coletor",
  "Nível",
  "Elemento PEP",
  "Denominação",
  "",
  "",
  "Pri",
  "Local de Instalação",
  "Perfil de Investimento",
  "ID real",
  "Tipo",
  "Centro de custo",
  "Centro de lucro",
  "PEP DE NIVEL ANTERIOR",
  "Área contábil",
  "Empresa",
  "Data de inicio",
  "Data de Fim",
  "Norma de Apropriação",
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
      "Equipamento/Especificação",
    ],
  },
};

const PEP_TEMPLATE_HEADERS = [
  "Usina",
  "Coletor de custo (NGHI)",
  "ID real",
  "Descrição",
  "Equipamento/Especificação",
];

function normalizeKey(text) {
  return text
    .trim()
    .toUpperCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^A-Z0-9]+/g, " ")
    .replace(/\s+/g, " ");
}

function compactKey(text) {
  return normalizeKey(text).replace(/\s+/g, "");
}

function resolveUsinaKey(raw) {
  const key = normalizeKey(raw);
  const compact = compactKey(raw);
  if (USINAS[key]) return key;
  const aliases = [
    ["BALBINA", ["BALBINA", "UHEBALBINA", "UHBA"]],
    ["COARACY NUNES", ["COARACYNUNES", "UHECOARACYNUNES", "UHCN"]],
    ["CURUA-UNA", ["CURUAUNA", "UHECURUAUNA", "CURUA", "UHCU"]],
    ["SAMUEL", ["SAMUEL", "UHESAMUEL", "UHSU"]],
    ["TUCURI", ["TUCURI", "TUCURUI", "UHETUCURI", "UHETUCURUI", "UHTU"]],
  ];

  const found = aliases.find(([, values]) =>
    values.some((alias) => compact.includes(alias))
  );
  return found ? found[0] : "";
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
    objeto: find(["EQUIPAMENTO", "OBJETO", "ESPECIFICACAO", "ESPECIFICAÇÃO"]),
    usina: find(["LOCAL", "USINA", "INSTALACAO", "INSTALAÇÃO"]),
  };
}

function detectColumnsByScan(rows) {
  const maxRows = Math.min(rows.length, 200);
  const fields = [
    {
      key: "coletor",
      aliases: ["COLETOR", "DEFINICAO", "DEFINIÇÃO", "PEP SIGEP"],
      data: /NGHI|COO/i,
    },
    {
      key: "idReal",
      aliases: ["ID REAL", "ID SIGEP", "ID"],
      data: /COO\./i,
    },
    {
      key: "desc",
      aliases: ["DESCRICAO", "DESCRIÇÃO", "DENOMINACAO", "DENOMINAÇÃO"],
    },
    {
      key: "objeto",
      aliases: ["EQUIPAMENTO", "OBJETO", "ESPECIFICACAO", "ESPECIFICAÇÃO"],
    },
    {
      key: "usina",
      aliases: ["LOCAL", "USINA", "INSTALACAO", "INSTALAÇÃO"],
    },
  ];

  const result = {
    coletor: -1,
    idReal: -1,
    desc: -1,
    objeto: -1,
    usina: -1,
    _rows: {},
    _score: 0,
    _headerIndex: -1,
  };

  const checkDataBelow = (col, startRow, regex) => {
    if (!regex) return true;
    const sample = rows.slice(startRow + 1, startRow + 30);
    return sample.some((r) => regex.test(String(r[col] || "")));
  };

  for (const field of fields) {
    for (let r = 0; r < maxRows; r += 1) {
      const row = rows[r] || [];
      for (let c = 0; c < row.length; c += 1) {
        const cell = normalizeKey(row[c] || "");
        if (!cell) continue;
        if (field.aliases.some((alias) => cell.includes(alias))) {
          if (checkDataBelow(c, r, field.data)) {
            result[field.key] = c;
            result._rows[field.key] = r;
            r = maxRows; // break outer
            break;
          }
        }
      }
    }
  }

  const foundCount = fields.filter((f) => result[f.key] !== -1).length;
  const headerIndex = Math.max(
    ...Object.values(result._rows || {}).filter((v) => typeof v === "number"),
    -1
  );
  const sample = headerIndex >= 0 ? rows.slice(headerIndex + 1, headerIndex + 30) : [];
  const hasData = sample.some((r) => {
    const coletor = result.coletor >= 0 ? (r[result.coletor] || "") : "";
    const id = result.idReal >= 0 ? (r[result.idReal] || "") : "";
    return /NGHI|COO/i.test(coletor) || /COO\\./i.test(id);
  });

  result._score = foundCount + (hasData ? 2 : 0);
  result._headerIndex = Number.isFinite(headerIndex) ? headerIndex : -1;
  return result;
}

function findHeaderAndColumns(rows) {
  let bestIndex = -1;
  let bestScore = -1;
  let bestCols = null;
  for (let i = 0; i < rows.length; i += 1) {
    const cols = detectColumns(rows[i] || []);
    const hits = [cols.coletor, cols.idReal, cols.desc, cols.usina].filter(
      (v) => v !== -1
    ).length;
    if (hits <= 0) continue;
    const sample = rows.slice(i + 1, i + 30);
    const hasData = sample.some((r) => {
      const coletor = cols.coletor >= 0 ? (r[cols.coletor] || "") : "";
      const id = cols.idReal >= 0 ? (r[cols.idReal] || "") : "";
      return /NGHI|COO/i.test(coletor) || /COO\./i.test(id);
    });
    const score = hits + (hasData ? 2 : 0);
    if (score > bestScore) {
      bestScore = score;
      bestIndex = i;
      bestCols = cols;
    }
  }

  if (bestIndex !== -1 && bestScore >= 2) {
    return { headerIndex: bestIndex, cols: bestCols, score: bestScore };
  }

  const scanned = detectColumnsByScan(rows);
  if (scanned._headerIndex !== -1) {
    return {
      headerIndex: scanned._headerIndex,
      cols: scanned,
      score: scanned._score || 0,
    };
  }

  return { headerIndex: -1, cols: detectColumns(rows[0] || []), score: -1 };
}

function displayUsina(usinaKey) {
  return usinaKey || "";
}

function padSeq(seq) {
  return String(seq).padStart(2, "0");
}

function formatSapDate(date) {
  const dd = String(date.getDate()).padStart(2, "0");
  const mm = String(date.getMonth() + 1).padStart(2, "0");
  const yyyy = date.getFullYear();
  return `${dd}.${mm}.${yyyy}`;
}

function normalizeCollectorBase(raw) {
  const value = String(raw || "").trim().toUpperCase();
  const match = value.match(/[A-Z]{3,12}(?:\.\d+)+/);
  return match ? match[0] : value;
}

function derivePrefixFromCollector(raw) {
  const base = normalizeCollectorBase(raw);
  const match = base.match(/[A-Z]{3,12}\.(\d+)/);
  return match ? match[1] : "";
}

function hasManualLevel2(raw) {
  const base = normalizeCollectorBase(raw);
  return /^[A-Z]{3,12}\.\d+\.\d+$/i.test(base);
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
  const obj = objeto ? shortenText(objeto) : "EQUIPAMENTO";
  return obj.slice(0, 40);
}

function resolveProjectMapping(usinaRaw, rowData = {}) {
  const usinaKey = resolveUsinaKey(usinaRaw);
  const knownMapping = USINAS[usinaKey];
  const advancedValues = [
    rowData.customColetor,
    rowData.customCentroCusto,
    rowData.customCentroLucro,
    rowData.customLocal,
  ]
    .map((value) => String(value || "").trim())
    .filter(Boolean);
  const inferredNgHi =
    advancedValues.find((value) => /[A-Z]{3,12}\.\d+/i.test(value)) || "";
  const manualLevel2 = inferredNgHi && hasManualLevel2(inferredNgHi);

  if (knownMapping) {
    return {
      usinaKey,
      usinaDisplay: displayUsina(usinaKey),
      ...knownMapping,
      prefix: inferredNgHi
        ? derivePrefixFromCollector(inferredNgHi) || knownMapping.prefix
        : knownMapping.prefix,
      centroCusto: rowData.customCentroCusto || knownMapping.centroCusto,
      lucro: rowData.customCentroLucro || knownMapping.lucro,
      localInstalacao: rowData.customLocal || knownMapping.localInstalacao,
      start: knownMapping.start,
      coletorBase: inferredNgHi
        ? normalizeCollectorBase(inferredNgHi)
        : `NGHI.${knownMapping.prefix}`,
      manualLevel2,
    };
  }

  const customPrefix = derivePrefixFromCollector(inferredNgHi);
  if (usinaRaw && customPrefix) {
    return {
      usinaKey: "",
      usinaDisplay: shortenText(String(usinaRaw).trim()),
      prefix: customPrefix,
      centroCusto: String(rowData.customCentroCusto || "").trim().toUpperCase(),
      lucro: String(rowData.customCentroLucro || "").trim().toUpperCase(),
      localInstalacao: String(rowData.customLocal || "").trim().toUpperCase(),
      start: 1,
      coletorBase: normalizeCollectorBase(inferredNgHi),
      manualLevel2,
    };
  }

  return null;
}

// Gera a estrutura completa de PEPs a partir das regras.
function buildOutput(projectRows) {
  const warnings = [];
  const currentSeqByPrefix = {};
  const defaultDataInicio = formatSapDate(new Date());
  const defaultDataFim = "31.12.2030";

  const output = [];
  const projectsMeta = [];

  projectRows.forEach((row, index) => {
    const rowData = Array.isArray(row)
      ? {
          coletor: row[0] || "",
          idReal: row[1] || "",
          desc: row[2] || "",
          usina: row[3] || "",
          objeto: row[4] || "",
        }
      : row || {};
    const coletorRaw = rowData.coletor || "";
    const idRealRaw = rowData.idReal || "";
    const descRaw = rowData.desc || "";
    const usinaRaw = rowData.usina || "";
    const objetoRaw = rowData.objeto || "";
    if (!usinaRaw && !idRealRaw && !descRaw) return;

    const mapping = resolveProjectMapping(usinaRaw, rowData);
    if (!mapping) {
      warnings.push(
        `Linha ${index + 1}: usina/mapeamento não reconhecido (${usinaRaw}).`
      );
      return;
    }

    const {
      usinaKey,
      usinaDisplay,
      prefix,
      coletorBase,
      centroCusto,
      lucro,
      localInstalacao,
      start,
      manualLevel2,
    } = mapping;
    if (currentSeqByPrefix[prefix] == null) {
      currentSeqByPrefix[prefix] = start;
    }
    if (currentSeqByPrefix[prefix] < start) {
      currentSeqByPrefix[prefix] = start;
    }

    const seq = currentSeqByPrefix[prefix];
    if (!manualLevel2) {
      currentSeqByPrefix[prefix] += 1;
    }
    const base = manualLevel2 ? coletorBase : `${coletorBase}.${padSeq(seq)}`;
    const idReal = (idRealRaw || "").trim();
    const desc = shortenText(descRaw || "");
    const denom403 = buildDenom403(idReal, objetoRaw || "");
    const recalcKey = normalizeKey(usinaKey || usinaDisplay);
    const perfil = String(rowData.customPerfil || "").trim().toUpperCase();
    const tipo = "G6";
    const areaContabil = "FCE1";
    const empresa = String(rowData.customEmpresa || "ENOR").trim().toUpperCase();
    const dataInicio = defaultDataInicio;
    const dataFim = defaultDataFim;
    const priTop = "7";
    const priCusto = "3";
    const priEquipamento = "2";
    projectsMeta.push({
      usinaKey,
      recalcKey,
      usinaRaw: usinaDisplay,
      prefix,
      seq,
      coletorBase,
      idReal,
      desc,
      centroCusto,
      lucro,
      localInstalacao,
      denom403,
      manualLevel2,
      perfil,
      tipo,
      areaContabil,
      empresa,
      dataInicio,
      dataFim,
      priTop,
      priCusto,
      priEquipamento,
    });

    output.push([
      usinaDisplay,
      coletorBase,
      "2",
      base,
      desc,
      "",
      "",
      priTop,
      localInstalacao,
      perfil,
      idReal,
      tipo,
      centroCusto,
      lucro,
      coletorBase,
      areaContabil,
      empresa,
      dataInicio,
      dataFim,
      "",
    ]);
    output.push([
      usinaDisplay,
      coletorBase,
      "3",
      `${base}.00001`,
      desc,
      "",
      "",
      priTop,
      localInstalacao,
      perfil,
      idReal,
      tipo,
      centroCusto,
      lucro,
      base,
      areaContabil,
      empresa,
      dataInicio,
      dataFim,
      "",
    ]);
    output.push([
      usinaDisplay,
      coletorBase,
      "4",
      `${base}.00001.001`,
      "CUSTO COMUM",
      "",
      "",
      priCusto,
      localInstalacao,
      perfil,
      idReal,
      tipo,
      centroCusto,
      lucro,
      `${base}.00001`,
      areaContabil,
      empresa,
      dataInicio,
      dataFim,
      `${base}.00001.003`,
    ]);
    output.push([
      usinaDisplay,
      coletorBase,
      "4",
      `${base}.00001.002`,
      "SERVIÇO",
      "",
      "",
      priCusto,
      localInstalacao,
      perfil,
      idReal,
      tipo,
      centroCusto,
      lucro,
      `${base}.00001`,
      areaContabil,
      empresa,
      dataInicio,
      dataFim,
      `${base}.00001.003`,
    ]);
    output.push([
      usinaDisplay,
      coletorBase,
      "4",
      `${base}.00001.003`,
      denom403,
      "",
      "",
      priEquipamento,
      localInstalacao,
      perfil,
      idReal,
      tipo,
      centroCusto,
      lucro,
      `${base}.00001`,
      areaContabil,
      empresa,
      dataInicio,
      dataFim,
      "",
    ]);
  });

  return { output, warnings, projectsMeta };
}

function rebuildFromMeta(meta, prioritizeKey = "") {
  const ordered = prioritizeKey
    ? [
        ...meta.filter((item) => item.recalcKey === prioritizeKey),
        ...meta.filter((item) => item.recalcKey !== prioritizeKey),
      ]
    : meta;
  const output = [];
  ordered.forEach((item) => {
    const base = item.manualLevel2
      ? item.coletorBase
      : `${item.coletorBase}.${padSeq(item.seq)}`;
    const perfil = item.perfil || "";
    const tipo = item.tipo || "G6";
    const areaContabil = item.areaContabil || "FCE1";
    const empresa = item.empresa || "ENOR";
    const dataInicio = item.dataInicio || formatSapDate(new Date());
    const dataFim = item.dataFim || "31.12.2030";
    const priTop = item.priTop || "7";
    const priCusto = item.priCusto || "3";
    const priEquipamento = item.priEquipamento || "2";
    output.push([
      item.usinaRaw,
      item.coletorBase,
      "2",
      base,
      item.desc,
      "",
      "",
      priTop,
      item.localInstalacao,
      perfil,
      item.idReal,
      tipo,
      item.centroCusto,
      item.lucro,
      item.coletorBase,
      areaContabil,
      empresa,
      dataInicio,
      dataFim,
      "",
    ]);
    output.push([
      item.usinaRaw,
      item.coletorBase,
      "3",
      `${base}.00001`,
      item.desc,
      "",
      "",
      priTop,
      item.localInstalacao,
      perfil,
      item.idReal,
      tipo,
      item.centroCusto,
      item.lucro,
      base,
      areaContabil,
      empresa,
      dataInicio,
      dataFim,
      "",
    ]);
    output.push([
      item.usinaRaw,
      item.coletorBase,
      "4",
      `${base}.00001.001`,
      "CUSTO COMUM",
      "",
      "",
      priCusto,
      item.localInstalacao,
      perfil,
      item.idReal,
      tipo,
      item.centroCusto,
      item.lucro,
      `${base}.00001`,
      areaContabil,
      empresa,
      dataInicio,
      dataFim,
      `${base}.00001.003`,
    ]);
    output.push([
      item.usinaRaw,
      item.coletorBase,
      "4",
      `${base}.00001.002`,
      "SERVIÇO",
      "",
      "",
      priCusto,
      item.localInstalacao,
      perfil,
      item.idReal,
      tipo,
      item.centroCusto,
      item.lucro,
      `${base}.00001`,
      areaContabil,
      empresa,
      dataInicio,
      dataFim,
      `${base}.00001.003`,
    ]);
    output.push([
      item.usinaRaw,
      item.coletorBase,
      "4",
      `${base}.00001.003`,
      item.denom403,
      "",
      "",
      priEquipamento,
      item.localInstalacao,
      perfil,
      item.idReal,
      tipo,
      item.centroCusto,
      item.lucro,
      `${base}.00001`,
      areaContabil,
      empresa,
      dataInicio,
      dataFim,
      "",
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
    if (highlightBase && row[3]?.startsWith(highlightBase)) {
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

function collectManualRows() {
  const customUsina = manualUsinaCustom.value.trim();
  const usina = (customUsina || manualUsina.value).trim();
  if (!usina) return [];
  return [
    {
      coletor: "",
      idReal: manualIdReal.value.trim(),
      desc: manualDesc.value.trim(),
      usina,
      objeto: manualObjeto.value.trim(),
      customColetor: manualColetor.value.trim(),
      customCentroCusto: manualCentroCusto.value.trim(),
      customCentroLucro: manualCentroLucro.value.trim(),
      customLocal: manualLocal.value.trim(),
      customPerfil: manualPerfil.value.trim(),
      customEmpresa: manualEmpresa.value.trim(),
    },
  ];
}

function validateManualInput() {
  const customUsina = manualUsinaCustom.value.trim();
  const selectedCustom = manualUsina.value === "__CUSTOM__";
  const hasManualData = [
    customUsina,
    manualColetor.value,
    manualIdReal.value,
    manualDesc.value,
    manualObjeto.value,
    manualCentroCusto.value,
    manualCentroLucro.value,
    manualLocal.value,
  ].some((value) => value.trim());

  if (!hasManualData) return "";
  if (!customUsina) {
    if (selectedCustom) {
      return "Informe o nome da usina personalizada nas opções avançadas.";
    }
    return "";
  }
  const coletor = manualColetor.value.trim();
  if (!coletor) {
    return "Para usina fora da lista, informe o Coletor de custo (NGHI) nas opções avançadas.";
  }
  if (!/^[A-Z]{3,12}\.\d+(?:\.\d+)?$/i.test(coletor)) {
    return "Coletor de custo inválido. Use o formato NGHI.0161 ou NGHI.0161.01.";
  }
  return "";
}

manualUsina.addEventListener("change", () => {
  const useCustomUsina = manualUsina.value === "__CUSTOM__";
  if (manualAdvanced && useCustomUsina) {
    manualAdvanced.open = true;
    manualUsinaCustom.focus();
  }
});

btnGenerate.addEventListener("click", () => {
  const sapRows = parseTsvAny(sapTextarea.value);
  const manualValidation = validateManualInput();
  if (manualValidation) {
    warningsBox.textContent = manualValidation;
    if (warningsGenerate) warningsGenerate.textContent = manualValidation;
    return;
  }
  const manualRows = collectManualRows();
  if (!sapRows.length && manualRows.length === 0) return;
  let projectRows = [];
  if (sapRows.length) {
    const headerInfo = findHeaderAndColumns(sapRows);
    const headerIndex = headerInfo.headerIndex;
    if (headerIndex === -1) {
      warningsBox.textContent = "Cabeçalho não encontrado.";
      if (warningsGenerate) warningsGenerate.textContent = "Cabeçalho não encontrado.";
      return;
    }
    const header = sapRows[headerIndex];
    const data = sapRows.slice(headerIndex + 1);
    const cols = headerInfo.cols || detectColumns(header);
    const missing = [];
    if (cols.idReal === -1) missing.push("ID real");
    if (cols.desc === -1) missing.push("Descrição");
    if (cols.usina === -1) missing.push("Usina/Local");
    if (missing.length) {
      const msg = "Colunas não identificadas: " + missing.join(", ");
      warningsBox.textContent = msg;
      if (warningsGenerate) warningsGenerate.textContent = msg;
    } else {
      warningsBox.textContent = "";
      if (warningsGenerate) warningsGenerate.textContent = "";
    }
    projectRows = data
      .map((row) => ({
        coletor: cols.coletor >= 0 ? row[cols.coletor] || "" : "",
        idReal: cols.idReal >= 0 ? row[cols.idReal] || "" : "",
        desc: cols.desc >= 0 ? row[cols.desc] || "" : "",
        usina: cols.usina >= 0 ? row[cols.usina] || "" : "",
        objeto: cols.objeto >= 0 ? row[cols.objeto] || "" : "",
      }))
      .filter((row) => {
        if (!Object.values(row).some((cell) => cell && String(cell).trim())) return false;
        if (!row.usina || !row.usina.trim()) return false;
        return true;
      });
  }

  projectRows = [...projectRows, ...manualRows];
  const { output, warnings, projectsMeta } = buildOutput(projectRows);
  warningsBox.textContent = warnings.length ? warnings.join(" ") : "";
  if (warningsGenerate) {
    warningsGenerate.textContent = warnings.length ? warnings.join(" ") : "";
  }
  window.__pepOutput = output;
  renderOutputTable(output);
  window.__pepMeta = projectsMeta;
  populateRecalcUsina(projectsMeta);
  populatePepFilter(output);
});

btnRecalc.addEventListener("click", () => {
  const usinaKey = normalizeKey(recalcUsina.value || "");
  const start = parseInt(recalcStart.value, 10);
  if (!window.__pepMeta || !usinaKey || !Number.isFinite(start)) return;
  let counter = start;
  window.__pepMeta.forEach((item) => {
    if (item.recalcKey === usinaKey) {
      item.seq = counter;
      counter += 1;
    }
  });
  const updated = rebuildFromMeta(window.__pepMeta, usinaKey);
  const target = normalizeKey(recalcUsina.value || "");
  const sorted = updated.slice().sort((a, b) => {
    const aMatch = normalizeKey(a[0] || "") === target ? 0 : 1;
    const bMatch = normalizeKey(b[0] || "") === target ? 0 : 1;
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
    const aMatch = normalizeKey(a[0] || "") === target ? 0 : 1;
    const bMatch = normalizeKey(b[0] || "") === target ? 0 : 1;
    if (aMatch !== bMatch) return aMatch - bMatch;
    return 0;
  });
  window.__pepOutput = sorted;
  renderOutputTable(sorted, filterPep.value);
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

btnTemplatePep.addEventListener("click", () => {
  const rows = [PEP_TEMPLATE_HEADERS];
  const worksheet = XLSX.utils.aoa_to_sheet(rows);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "Modelo PEP");
  XLSX.writeFile(workbook, "modelo_pep.xlsx");
});

renderOutputTable([]);

function populateRecalcUsina(meta) {
  const current = recalcUsina.value;
  const options = Array.from(
    new Set((meta || []).map((item) => item.usinaRaw).filter(Boolean))
  ).sort();
  recalcUsina.innerHTML = '<option value="">Selecione</option>';
  options.forEach((usina) => {
    const opt = document.createElement("option");
    opt.value = usina;
    opt.textContent = usina;
    recalcUsina.appendChild(opt);
  });
  if (current && options.includes(current)) {
    recalcUsina.value = current;
  }
}

function populatePepFilter(rows) {
  const set = new Set(
    rows.filter((r) => r[2] === "2").map((r) => r[3]).filter(Boolean)
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
  const selected = window.__pepOutput.filter((r) => r[3]?.startsWith(base));
  const rest = window.__pepOutput.filter((r) => !r[3]?.startsWith(base));
  renderOutputTable([...selected, ...rest], base);
});

const savedCapex = localStorage.getItem("capex_tsv");
if (savedCapex && !sapTextarea.value) {
  sapTextarea.value = savedCapex;
}
