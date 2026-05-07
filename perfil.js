const profileQuery = document.getElementById("profile-query");
const profileSummary = document.getElementById("profile-summary");
const profileResults = document.getElementById("profile-results");
const sapQuery = document.getElementById("sap-query");
const sapResults = document.getElementById("sap-results");
const btnClearProfile = document.getElementById("btn-clear-profile");
const btnCopyProfile = document.getElementById("btn-copy-profile");

const STOP_WORDS = new Set([
  "A",
  "AS",
  "COM",
  "DA",
  "DAS",
  "DE",
  "DO",
  "DOS",
  "E",
  "EM",
  "UM",
  "UMA",
  "CADA",
  "O",
  "OS",
  "OU",
  "PARA",
  "POR",
  "SISTEMA",
  "SIST",
  "UHE",
  "USINA",
]);

const SEARCH_ABBREVIATIONS = {
  AQ: ["AQUISICAO"],
  AQUIS: ["AQUISICAO"],
  AUTOM: ["AUTOMACAO"],
  AUX: ["AUXILIAR"],
  ELETR: ["ELETRICO", "ELETRICA"],
  GER: ["GERADOR"],
  IMPLANT: ["IMPLANTACAO"],
  INST: ["INSTALACAO"],
  MED: ["MEDICAO"],
  MODERN: ["MODERNIZACAO"],
  MODERNIZ: ["MODERNIZACAO"],
  PROT: ["PROTECAO"],
  RESFRI: ["RESFRIAMENTO"],
  REVIT: ["REVITALIZACAO"],
  SERV: ["SERVICO"],
  SIST: ["SISTEMA"],
  SUBEST: ["SUBESTACAO"],
  SUBST: ["SUBSTITUICAO"],
  SUPERV: ["SUPERVISAO", "SUPERVISORIO"],
  TRAFO: ["TRANSFORMADOR"],
  TURB: ["TURBINA"],
  UNID: ["UNIDADE"],
  UIDADES: ["UNIDADES"],
};

function normalizeKey(text) {
  return String(text || "")
    .trim()
    .toUpperCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^A-Z0-9]+/g, " ")
    .replace(/\s+/g, " ");
}

function normalizeSearchKey(text) {
  return normalizeKey(text)
    .split(" ")
    .flatMap((token) => [token, ...(SEARCH_ABBREVIATIONS[token] || [])])
    .join(" ")
    .replace(/\s+/g, " ")
    .trim();
}

function tokenizeSearch(text) {
  return normalizeSearchKey(text)
    .split(" ")
    .filter((token) => token.length >= 3 && !STOP_WORDS.has(token));
}

function hasTerm(textKey, term) {
  return textKey.includes(normalizeSearchKey(term));
}

function getProfilesByTuc(tuc) {
  return (window.PERFIL_INVESTIMENTO || []).filter((item) => item.tuc === tuc);
}

function getProfileByCode(code) {
  return (window.PERFIL_INVESTIMENTO || []).find((item) => item.codigo === code);
}

function getTucTitle(tuc) {
  const found = (window.TUC_REFERENCES || []).find((item) => item.codigo === tuc);
  return found ? found.descricao : "";
}

function getUarByTuc(tuc) {
  return (window.TUC_UAR || []).find((item) => item.codigo === tuc);
}

function getTucProfiles(tuc) {
  return getProfilesByTuc(tuc).filter(
    (profile) => !profile.codigo.includes("PR") && !/CONVERS/i.test(profile.descricao)
  );
}

function scoreProfile(profile, tokens) {
  const profileTokens = new Set(tokenizeSearch(profile.descricao));
  return tokens.reduce((score, token) => score + (profileTokens.has(token) ? 2 : 0), 0);
}

function bestProfileForTuc(tuc, tokens) {
  const profiles = getProfilesByTuc(tuc).filter(
    (profile) => !profile.codigo.includes("PR") && !/CONVERS/i.test(profile.descricao)
  );
  if (!profiles.length) return { profile: null, options: [] };

  const scored = profiles
    .map((profile) => ({
      profile,
      score: scoreProfile(profile, tokens),
    }))
    .sort((a, b) => b.score - a.score || a.profile.codigo.localeCompare(b.profile.codigo));

  const top = scored[0];
  const tied = scored.filter((item) => item.score === top.score);
  return {
    profile: tied.length === 1 ? top.profile : null,
    options: scored.map((item) => item.profile),
  };
}

function inferInvestmentProfile(text) {
  const textKey = normalizeSearchKey(text);
  const tokens = tokenizeSearch(text);
  if (!textKey || tokens.length === 0) return [];

  const byTuc = new Map();
  (window.PROFILE_EXAMPLES || []).forEach((example) => {
    const matched = (example.termos || []).some((term) => hasTerm(textKey, term));
    const profile = matched ? getProfileByCode(example.perfil) : null;
    if (!profile) return;
    const current = byTuc.get(profile.tuc) || {
      tuc: profile.tuc,
      score: 0,
      motivos: [],
      preferredProfiles: {},
    };
    current.score += 200;
    current.motivos.push(example.motivo || profile.descricao);
    current.preferredProfiles[profile.codigo] =
      (current.preferredProfiles[profile.codigo] || 0) + 1;
    byTuc.set(profile.tuc, current);
  });

  (window.TUC_RULES || []).forEach((rule) => {
    const matched = (rule.termos || []).some((term) => hasTerm(textKey, term));
    if (!matched) return;
    const current = byTuc.get(rule.tuc) || {
      tuc: rule.tuc,
      score: 0,
      motivos: [],
      preferredProfiles: {},
    };
    current.score += 100;
    current.motivos.push(rule.motivo);
    byTuc.set(rule.tuc, current);
  });

  (window.TUC_REFERENCES || []).forEach((tucRef) => {
    const titleTokens = new Set(tokenizeSearch(tucRef.descricao));
    const score = tokens.reduce(
      (total, token) => total + (titleTokens.has(token) ? 8 : 0),
      0
    );
    if (!score) return;
    const current = byTuc.get(tucRef.codigo) || {
      tuc: tucRef.codigo,
      score: 0,
      motivos: [],
      preferredProfiles: {},
    };
    current.score += score;
    current.motivos.push(tucRef.descricao);
    byTuc.set(tucRef.codigo, current);
  });

  (window.TUC_UAR || []).forEach((tucUar) => {
    let score = 0;
    const matches = [];
    const tokenKey = tokens.join(" ");
    (tucUar.itens || []).forEach((item) => {
      const itemTokenList = tokenizeSearch(item);
      const itemTokenKey = itemTokenList.join(" ");
      const itemTokens = new Set(itemTokenList);
      const overlap = tokens.filter((token) => itemTokens.has(token));
      const phraseMatch =
        itemTokenList.length > 1 && itemTokenKey && tokenKey.includes(itemTokenKey);
      const shortExact =
        itemTokens.size === 1 && overlap.length === 1 && textKey.includes(overlap[0]);
      const usefulOverlap =
        overlap.length >= 2 || overlap.some((token) => token.length >= 6);
      if (!phraseMatch && !shortExact && !usefulOverlap) return;
      score += phraseMatch ? 90 : Math.min(overlap.length * 18, 54);
      matches.push(item);
    });
    if (!score) return;
    const current = byTuc.get(tucUar.codigo) || {
      tuc: tucUar.codigo,
      score: 0,
      motivos: [],
      preferredProfiles: {},
      uarMatches: [],
    };
    current.score += score;
    current.motivos.push(`UAR MCPSE: ${matches.slice(0, 3).join("; ")}`);
    current.uarMatches = [
      ...(current.uarMatches || []),
      ...matches.filter((item) => !(current.uarMatches || []).includes(item)),
    ];
    byTuc.set(tucUar.codigo, current);
  });

  (window.PERFIL_INVESTIMENTO || []).forEach((profile) => {
    if (profile.codigo.includes("PR") || /CONVERS/i.test(profile.descricao)) return;
    const profileTokens = new Set(tokenizeSearch(profile.descricao));
    const score = tokens.reduce(
      (total, token) => total + (profileTokens.has(token) ? 5 : 0),
      0
    );
    if (!score) return;
    const current = byTuc.get(profile.tuc) || {
      tuc: profile.tuc,
      score: 0,
      motivos: [],
      preferredProfiles: {},
    };
    current.score += score;
    current.motivos.push(profile.descricao);
    byTuc.set(profile.tuc, current);
  });

  return Array.from(byTuc.values())
    .filter((item) => item.score >= 8)
    .map((item) => {
      const best = bestProfileForTuc(item.tuc, tokens);
      const preferredCode = Object.entries(item.preferredProfiles || {}).sort(
        (a, b) => b[1] - a[1] || a[0].localeCompare(b[0])
      )[0]?.[0];
      const preferredProfile = preferredCode ? getProfileByCode(preferredCode) : null;
      const profile = preferredProfile || best.profile;
      const options = preferredProfile
        ? [
            preferredProfile,
            ...best.options.filter((option) => option.codigo !== preferredProfile.codigo),
          ]
        : best.options;
      return {
        tuc: item.tuc,
        tucDescricao: getTucTitle(item.tuc),
        perfil: profile ? profile.codigo : "",
        perfilDescricao: profile ? profile.descricao : "",
        opcoesPerfil: options,
        score: item.score,
        motivos: Array.from(new Set(item.motivos)).slice(0, 3),
        uarMatches: item.uarMatches || [],
      };
    })
    .sort((a, b) => b.score - a.score || a.tuc.localeCompare(b.tuc));
}

function createEl(tag, className, text = "") {
  const el = document.createElement(tag);
  if (className) el.className = className;
  if (text) el.textContent = text;
  return el;
}

function renderProfileResults() {
  const query = profileQuery.value.trim();
  const candidates = inferInvestmentProfile(query);
  window.__profileCandidates = candidates;

  profileResults.innerHTML = "";
  if (!query) {
    profileSummary.textContent = "Digite um texto para iniciar a consulta.";
    return;
  }
  if (!candidates.length) {
    profileSummary.textContent = "Nenhum TUC provável encontrado.";
    profileResults.appendChild(
      createEl("p", "hint", "Tente usar o nome do equipamento principal do projeto.")
    );
    return;
  }

  const top = candidates[0];
  profileSummary.textContent = top.perfil
    ? `Mais provável: TUC ${top.tuc} / Perfil ${top.perfil}`
    : `Mais provável: TUC ${top.tuc}. Escolha uma das opções de perfil abaixo.`;

  candidates.forEach((candidate, index) => {
    const card = createEl("article", "profile-card");
    card.appendChild(createEl("span", "profile-rank", `#${index + 1}`));
    card.appendChild(
      createEl("h3", "", `TUC ${candidate.tuc} - ${candidate.tucDescricao || "Sem descrição"}`)
    );

    if (candidate.perfil) {
      card.appendChild(
        createEl(
          "p",
          "profile-code",
          `Perfil sugerido: ${candidate.perfil} - ${candidate.perfilDescricao}`
        )
      );
    } else {
      card.appendChild(
        createEl("p", "profile-code", "Perfil precisa de conferência manual.")
      );
    }

    const options = createEl("div", "profile-options");
    candidate.opcoesPerfil.forEach((profile) => {
      options.appendChild(
        createEl("span", "profile-option", `${profile.codigo} - ${profile.descricao}`)
      );
    });
    card.appendChild(options);

    if (candidate.motivos.length) {
      card.appendChild(createEl("p", "profile-reason", `Indícios: ${candidate.motivos.join("; ")}`));
    }

    if (candidate.uarMatches && candidate.uarMatches.length) {
      const uarBox = createEl("div", "profile-uar");
      uarBox.appendChild(createEl("strong", "", "Unidades de Adição e Retirada encontradas"));
      const list = createEl("ul", "");
      candidate.uarMatches.forEach((item) => {
        const li = createEl("li", "", item);
        list.appendChild(li);
      });
      uarBox.appendChild(list);
      card.appendChild(uarBox);
    }

    const uar = getUarByTuc(candidate.tuc);
    if (uar && uar.itens.length) {
      const details = createEl("details", "profile-uar-details");
      details.appendChild(
        createEl("summary", "", `Ver todas as UAR da TUC ${candidate.tuc}`)
      );
      const list = createEl("ul", "");
      uar.itens.forEach((item) => {
        list.appendChild(createEl("li", "", item));
      });
      details.appendChild(list);
      card.appendChild(details);
    }

    profileResults.appendChild(card);
  });
}

function renderSapResults() {
  const query = sapQuery.value.trim();
  const queryKey = normalizeKey(query);
  const matches = (window.PERFIL_INVESTIMENTO || [])
    .filter((profile) => {
      if (!queryKey) return true;
      return (
        normalizeKey(profile.codigo).includes(queryKey) ||
        normalizeKey(profile.descricao).includes(queryKey) ||
        normalizeKey(profile.tuc).includes(queryKey)
      );
    });

  const table = document.createElement("table");
  const thead = document.createElement("thead");
  const headRow = document.createElement("tr");
  ["Detalhes", "Perfil", "TUC", "Descrição"].forEach((header) => {
    const th = document.createElement("th");
    th.textContent = header;
    headRow.appendChild(th);
  });
  thead.appendChild(headRow);
  table.appendChild(thead);

  const tbody = document.createElement("tbody");
  matches.forEach((profile, index) => {
    const tr = document.createElement("tr");
    const detailId = `sap-detail-${index}`;
    const actionTd = document.createElement("td");
    const btn = document.createElement("button");
    btn.className = "sap-detail-toggle";
    btn.type = "button";
    btn.textContent = "Ver";
    btn.setAttribute("aria-expanded", "false");
    btn.setAttribute("aria-controls", detailId);
    actionTd.appendChild(btn);
    tr.appendChild(actionTd);

    [profile.codigo, profile.tuc, profile.descricao].forEach((value) => {
      const td = document.createElement("td");
      td.textContent = value;
      tr.appendChild(td);
    });
    tbody.appendChild(tr);

    const detailTr = document.createElement("tr");
    detailTr.id = detailId;
    detailTr.className = "sap-detail-row";
    detailTr.hidden = true;
    const detailTd = document.createElement("td");
    detailTd.colSpan = 4;

    const detailBox = createEl("div", "sap-detail-box");
    const detailHead = createEl("div", "sap-detail-head");
    detailHead.appendChild(
      createEl("h3", "", `TUC ${profile.tuc} - ${getTucTitle(profile.tuc) || "Sem descrição"}`)
    );
    detailBox.appendChild(detailHead);

    const detailGrid = createEl("div", "sap-detail-grid");
    const profilePanel = createEl("div", "sap-detail-panel");
    profilePanel.appendChild(createEl("h4", "", "Perfis SAP desta TUC"));

    const profileList = createEl("div", "profile-options");
    getTucProfiles(profile.tuc).forEach((item) => {
      profileList.appendChild(
        createEl("span", "profile-option", `${item.codigo} - ${item.descricao}`)
      );
    });
    profilePanel.appendChild(profileList);
    detailGrid.appendChild(profilePanel);

    const uarPanel = createEl("div", "sap-detail-panel");
    const uar = getUarByTuc(profile.tuc);
    if (uar && uar.itens.length) {
      const uarBox = createEl("div", "profile-uar");
      uarBox.appendChild(createEl("strong", "", "Unidades de Adição e Retirada"));
      const list = createEl("ul", "");
      uar.itens.forEach((item) => list.appendChild(createEl("li", "", item)));
      uarBox.appendChild(list);
      uarPanel.appendChild(uarBox);
    } else {
      uarPanel.appendChild(
        createEl("p", "hint", "Não há UAR extraída para esta TUC na base local.")
      );
    }
    detailGrid.appendChild(uarPanel);
    detailBox.appendChild(detailGrid);

    detailTd.appendChild(detailBox);
    detailTr.appendChild(detailTd);
    tbody.appendChild(detailTr);

    btn.addEventListener("click", () => {
      const expanded = btn.getAttribute("aria-expanded") === "true";
      btn.setAttribute("aria-expanded", String(!expanded));
      btn.textContent = expanded ? "Ver" : "Fechar";
      detailTr.hidden = expanded;
    });
  });
  table.appendChild(tbody);

  sapResults.innerHTML = "";
  sapResults.appendChild(table);
}

async function copyProfileResult() {
  const candidates = window.__profileCandidates || [];
  if (!candidates.length) return;
  const lines = candidates.map((candidate) => {
    const profiles = candidate.opcoesPerfil
      .map((profile) => `${profile.codigo} - ${profile.descricao}`)
      .join("; ");
    return `TUC ${candidate.tuc} - ${candidate.tucDescricao || ""}\t${
      candidate.perfil || "CONFERIR"
    }\t${candidate.perfilDescricao || profiles}`;
  });
  await navigator.clipboard.writeText(lines.join("\n"));
}

profileQuery.addEventListener("input", renderProfileResults);
sapQuery.addEventListener("input", renderSapResults);
btnClearProfile.addEventListener("click", () => {
  profileQuery.value = "";
  renderProfileResults();
  profileQuery.focus();
});
btnCopyProfile.addEventListener("click", copyProfileResult);

renderProfileResults();
renderSapResults();
