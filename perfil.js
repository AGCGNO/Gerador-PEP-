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

function normalizeKey(text) {
  return String(text || "")
    .trim()
    .toUpperCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^A-Z0-9]+/g, " ")
    .replace(/\s+/g, " ");
}

function tokenizeSearch(text) {
  return normalizeKey(text)
    .split(" ")
    .filter((token) => token.length >= 3 && !STOP_WORDS.has(token));
}

function hasTerm(textKey, term) {
  return textKey.includes(normalizeKey(term));
}

function getProfilesByTuc(tuc) {
  return (window.PERFIL_INVESTIMENTO || []).filter((item) => item.tuc === tuc);
}

function getTucTitle(tuc) {
  const found = (window.TUC_REFERENCES || []).find((item) => item.codigo === tuc);
  return found ? found.descricao : "";
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
    options: scored.slice(0, 6).map((item) => item.profile),
  };
}

function inferInvestmentProfile(text) {
  const textKey = normalizeKey(text);
  const tokens = tokenizeSearch(text);
  if (!textKey || tokens.length === 0) return [];

  const byTuc = new Map();
  (window.TUC_RULES || []).forEach((rule) => {
    const matched = (rule.termos || []).some((term) => hasTerm(textKey, term));
    if (!matched) return;
    const current = byTuc.get(rule.tuc) || { tuc: rule.tuc, score: 0, motivos: [] };
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
    };
    current.score += score;
    current.motivos.push(tucRef.descricao);
    byTuc.set(tucRef.codigo, current);
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
    };
    current.score += score;
    current.motivos.push(profile.descricao);
    byTuc.set(profile.tuc, current);
  });

  return Array.from(byTuc.values())
    .filter((item) => item.score >= 8)
    .map((item) => {
      const best = bestProfileForTuc(item.tuc, tokens);
      return {
        tuc: item.tuc,
        tucDescricao: getTucTitle(item.tuc),
        perfil: best.profile ? best.profile.codigo : "",
        perfilDescricao: best.profile ? best.profile.descricao : "",
        opcoesPerfil: best.options,
        score: item.score,
        motivos: Array.from(new Set(item.motivos)).slice(0, 3),
      };
    })
    .sort((a, b) => b.score - a.score || a.tuc.localeCompare(b.tuc))
    .slice(0, 8);
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
    })
    .slice(0, 80);

  const table = document.createElement("table");
  const thead = document.createElement("thead");
  const headRow = document.createElement("tr");
  ["Perfil", "TUC", "Descrição"].forEach((header) => {
    const th = document.createElement("th");
    th.textContent = header;
    headRow.appendChild(th);
  });
  thead.appendChild(headRow);
  table.appendChild(thead);

  const tbody = document.createElement("tbody");
  matches.forEach((profile) => {
    const tr = document.createElement("tr");
    [profile.codigo, profile.tuc, profile.descricao].forEach((value) => {
      const td = document.createElement("td");
      td.textContent = value;
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
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
