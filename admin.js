const REPO = "pauloricardoipa103-code/dashboard-programacao-ipora";
const DATA_PATH = "dados/anomalias.json";

let parsedRows = [];

const excelFile = document.getElementById("excelFile");
const githubToken = document.getElementById("githubToken");
const publishButton = document.getElementById("publishButton");
const statusBox = document.getElementById("statusBox");

const fmt = n => Number(n || 0).toLocaleString("pt-BR");
const norm = v => (v === null || v === undefined || v === "" || v === "#N/A") ? "Não informado" : String(v);

function setStatus(message) {
  statusBox.textContent = message;
}

function cleanCell(value) {
  if (value === null || value === undefined) return "";
  if (value instanceof Date) return value.toISOString().slice(0, 10);
  const text = String(value).trim();
  return text === "#NAME?" || text === "undefined" || text === "null" ? "" : text;
}

function normHeader(value) {
  return String(value || "").trim().toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
}

function toNumber(value) {
  if (value === "" || value === null || value === undefined) return null;
  const number = typeof value === "number" ? value : Number(String(value).replace(",", "."));
  return Number.isFinite(number) ? number : null;
}

function toDateText(value) {
  if (!value) return "";
  if (value instanceof Date) return value.toISOString().slice(0, 10);
  if (typeof value === "number") {
    const date = XLSX.SSF.parse_date_code(value);
    if (date) return `${date.y}-${String(date.m).padStart(2, "0")}-${String(date.d).padStart(2, "0")}`;
  }
  const text = String(value).trim();
  if (/^\d+(\.\d+)?$/.test(text) && Number(text) > 20000 && window.XLSX) {
    const date = XLSX.SSF.parse_date_code(Number(text));
    if (date) return `${date.y}-${String(date.m).padStart(2, "0")}-${String(date.d).padStart(2, "0")}`;
  }
  const br = text.match(/^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{2,4})$/);
  if (br) {
    const y = br[3].length === 2 ? `20${br[3]}` : br[3];
    return `${y}-${br[2].padStart(2, "0")}-${br[1].padStart(2, "0")}`;
  }
  const iso = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  return iso ? `${iso[1]}-${iso[2].padStart(2, "0")}-${iso[3].padStart(2, "0")}` : text;
}

function parseWorksheetRows(table) {
  if (!table.length) return [];
  const headers = table[0].map(v => normHeader(v));
  const index = name => headers.indexOf(normHeader(name));
  const required = ["Defeito", "SE", "Latitude", "Longitude", "Instalação", "ID Anomalia", "Execução", "DATA DE REGISTRO", "Tipo de trecho", "Taxonomia"];
  const missing = required.filter(name => index(name) === -1);
  if (missing.length) throw new Error(`Colunas obrigatorias ausentes: ${missing.join(", ")}`);
  const val = (row, name) => {
    const i = index(name);
    return i >= 0 ? cleanCell(row[i]) : "";
  };
  return table.slice(1).map(row => {
    const lat = toNumber(val(row, "Latitude"));
    const lon = toNumber(val(row, "Longitude"));
    return {
      defeito: val(row, "Defeito"),
      os: val(row, "OS"),
      se: val(row, "SE"),
      alimentador: val(row, "Alimentador"),
      lat,
      lon,
      crit: val(row, "Crit."),
      poste: val(row, "Poste"),
      projeto: val(row, "Projeto"),
      instalacao: val(row, "Instalação"),
      id: val(row, "ID Anomalia"),
      empresa: val(row, "Empresa"),
      mes: toDateText(val(row, "Mês")),
      ose: val(row, "OSE"),
      execucao: val(row, "Execução"),
      dataExecucao: toDateText(val(row, "Data de execução")),
      prioridade: val(row, "Prioridade"),
      tipoAnomalia: val(row, "Tipo de anomalia"),
      seccional: val(row, "Seccional"),
      dataRegistro: toDateText(val(row, "DATA DE REGISTRO")),
      tipoTrecho: val(row, "Tipo de trecho"),
      semana: val(row, "Semana"),
      pendente: val(row, "Anomalias pendentes"),
      taxonomia: val(row, "Taxonomia"),
      conjunto: val(row, "conjunto"),
      clientes: val(row, "Qtd Clientes"),
      statusEquipamento: val(row, "Status Equipamento"),
      prazo: val(row, "Prazo de execução")
    };
  }).filter(row => row.id || row.se || row.instalacao || row.defeito);
}

function statusOf(row) {
  const raw = norm(row.execucao);
  if (raw !== "Não informado") return raw;
  return norm(row.pendente).toLowerCase().includes("pend") ? "Pendente" : "Não informado";
}

function updatePreview(rows) {
  const executed = rows.filter(r => statusOf(r).toLowerCase().includes("execut")).length;
  const pending = rows.filter(r => statusOf(r).toLowerCase().includes("pend")).length;
  const projects = new Set(rows.map(r => norm(r.projeto)).filter(v => v !== "Não informado")).size;
  document.getElementById("previewMetrics").hidden = false;
  document.getElementById("mTotal").textContent = fmt(rows.length);
  document.getElementById("mExec").textContent = fmt(executed);
  document.getElementById("mPend").textContent = fmt(pending);
  document.getElementById("mProj").textContent = fmt(projects);
  publishButton.disabled = !rows.length;
}

excelFile.addEventListener("change", async event => {
  const file = event.target.files && event.target.files[0];
  if (!file) return;
  try {
    setStatus(`Lendo ${file.name}...`);
    const buffer = await file.arrayBuffer();
    const workbook = XLSX.read(buffer, { type: "array", cellDates: true });
    const sheetName = workbook.SheetNames.find(name => name.trim().toLowerCase() === "geral anomalias") || workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const table = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null, raw: true });
    parsedRows = parseWorksheetRows(table);
    if (!parsedRows.length) throw new Error("A planilha nao possui registros validos.");
    updatePreview(parsedRows);
    setStatus(`Planilha pronta para publicar.\nAba lida: ${sheetName}\nRegistros: ${fmt(parsedRows.length)}`);
  } catch (error) {
    parsedRows = [];
    publishButton.disabled = true;
    console.error(error);
    setStatus(`Erro ao ler planilha: ${error.message}`);
  }
});

async function getCurrentFileSha(token) {
  const url = `https://api.github.com/repos/${REPO}/contents/${DATA_PATH}?ref=main`;
  const response = await fetch(url, {
    headers: {
      Authorization: `Bearer ${token}`,
      Accept: "application/vnd.github+json",
      "X-GitHub-Api-Version": "2022-11-28"
    }
  });
  if (response.status === 404) return null;
  if (!response.ok) throw new Error(`Falha ao buscar arquivo atual: HTTP ${response.status}`);
  const json = await response.json();
  return json.sha;
}

function toBase64Utf8(value) {
  const bytes = new TextEncoder().encode(value);
  let binary = "";
  bytes.forEach(byte => binary += String.fromCharCode(byte));
  return btoa(binary);
}

async function publishData() {
  const token = githubToken.value.trim();
  if (!token) throw new Error("Informe o token GitHub.");
  if (!parsedRows.length) throw new Error("Selecione uma planilha valida antes de publicar.");
  const updatedAt = new Date().toISOString();
  const payload = {
    updatedAt,
    source: "Base publicada",
    rows: parsedRows
  };
  const sha = await getCurrentFileSha(token);
  const body = {
    message: `Atualiza base de anomalias em ${new Date(updatedAt).toLocaleString("pt-BR")}`,
    content: toBase64Utf8(JSON.stringify(payload)),
    branch: "main"
  };
  if (sha) body.sha = sha;
  const response = await fetch(`https://api.github.com/repos/${REPO}/contents/${DATA_PATH}`, {
    method: "PUT",
    headers: {
      Authorization: `Bearer ${token}`,
      Accept: "application/vnd.github+json",
      "X-GitHub-Api-Version": "2022-11-28",
      "Content-Type": "application/json"
    },
    body: JSON.stringify(body)
  });
  if (!response.ok) {
    const text = await response.text();
    throw new Error(`Falha ao publicar: HTTP ${response.status} ${text}`);
  }
  return updatedAt;
}

publishButton.addEventListener("click", async () => {
  publishButton.disabled = true;
  try {
    setStatus("Publicando base central no GitHub...");
    const updatedAt = await publishData();
    setStatus(`Base publicada com sucesso.\nAtualizado em ${new Date(updatedAt).toLocaleString("pt-BR")}.\nAtualize o painel com Ctrl+F5 para conferir.`);
  } catch (error) {
    console.error(error);
    setStatus(error.message);
  } finally {
    publishButton.disabled = !parsedRows.length;
  }
});
