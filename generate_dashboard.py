import base64
import io
import json
from datetime import datetime
from pathlib import Path

import openpyxl
from PIL import Image


ROOT = Path(__file__).resolve().parent
XLSX = Path(r"C:\Users\WIN_11\Downloads\BI teste\Pasta1.xlsx")
LOGO = Path(r"C:\Users\WIN_11\Downloads\logo Remo.png")
OUT = ROOT / "dashboard_anomalias_remo.html"


def clean(value):
    if value is None:
        return ""
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d")
    text = str(value).strip()
    return "" if text in {"#NAME?", "None"} else text


def row_value(row, index):
    return clean(row[index - 1])


def load_rows():
    workbook = openpyxl.load_workbook(XLSX, data_only=True)
    sheet = workbook["Geral anomalias"]
    rows = []
    for values in sheet.iter_rows(min_row=2, values_only=True):
        lat = values[5]
        lon = values[6]
        try:
            lat = float(lat)
            lon = float(lon)
        except (TypeError, ValueError):
            lat = None
            lon = None

        rows.append(
            {
                "defeito": row_value(values, 1),
                "os": row_value(values, 2),
                "se": row_value(values, 4),
                "alimentador": row_value(values, 5),
                "lat": lat,
                "lon": lon,
                "crit": row_value(values, 8),
                "poste": row_value(values, 9),
                "projeto": row_value(values, 10),
                "instalacao": row_value(values, 11),
                "id": row_value(values, 12),
                "empresa": row_value(values, 13),
                "mes": row_value(values, 14),
                "ose": row_value(values, 15),
                "execucao": row_value(values, 16),
                "dataExecucao": row_value(values, 17),
                "prioridade": row_value(values, 18),
                "tipoAnomalia": row_value(values, 19),
                "seccional": row_value(values, 20),
                "dataRegistro": row_value(values, 21),
                "tipoTrecho": row_value(values, 22),
                "semana": row_value(values, 23),
                "pendente": row_value(values, 24),
                "taxonomia": row_value(values, 25),
                "conjunto": row_value(values, 26),
                "clientes": row_value(values, 27),
                "statusEquipamento": row_value(values, 28),
                "prazo": row_value(values, 29),
            }
        )
    return rows


def encode_logo():
    image = Image.open(LOGO).convert("RGBA")
    bg = Image.new("RGBA", image.size, (255, 255, 255, 255))
    diff = Image.new("RGBA", image.size, (0, 0, 0, 0))
    pixels = []
    for pixel in image.getdata():
        r, g, b, a = pixel
        pixels.append((0, 0, 0, 255) if a > 0 and (r < 245 or g < 245 or b < 245) else (0, 0, 0, 0))
    diff.putdata(pixels)
    bbox = diff.getbbox()
    if bbox:
        pad = 18
        left = max(0, bbox[0] - pad)
        top = max(0, bbox[1] - pad)
        right = min(image.width, bbox[2] + pad)
        bottom = min(image.height, bbox[3] + pad)
        image = image.crop((left, top, right, bottom))
    image.thumbnail((520, 180), Image.LANCZOS)
    buffer = io.BytesIO()
    image.save(buffer, format="PNG", optimize=True)
    return f"data:image/png;base64,{base64.b64encode(buffer.getvalue()).decode('ascii')}"


def build_html(rows, logo_data):
    payload = json.dumps(rows, ensure_ascii=False, separators=(",", ":"))
    generated = datetime.now().strftime("%d/%m/%Y %H:%M")
    return f"""<!doctype html>
<html lang="pt-BR">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Dashboard de Anomalias | REMO Engenharia</title>
  <style>
    @import url("https://unpkg.com/leaflet@1.9.4/dist/leaflet.css");
    :root {{
      --remo-blue: #10496f;
      --remo-blue-2: #0b324e;
      --remo-green: #29c77b;
      --ink: #162331;
      --muted: #657382;
      --line: #d8e1e8;
      --soft: #f4f8fb;
      --yellow: #ffd84d;
      --danger: #d94f45;
      --panel: #ffffff;
      --shadow: 0 10px 28px rgba(19, 42, 60, .10);
    }}
    * {{ box-sizing: border-box; }}
    body {{
      margin: 0;
      background: #eef4f7;
      color: var(--ink);
      font: 14px/1.45 "Segoe UI", Arial, sans-serif;
    }}
    .topbar {{
      display: grid;
      grid-template-columns: 210px 1fr auto;
      gap: 20px;
      align-items: center;
      padding: 18px 28px;
      color: #fff;
      background: linear-gradient(110deg, var(--remo-blue-2), var(--remo-blue) 68%, #17745d);
      border-bottom: 4px solid var(--remo-green);
    }}
    .brand {{
      height: 70px;
      display: flex;
      align-items: center;
      justify-content: center;
      background: rgba(255,255,255,.96);
      border-radius: 8px;
      padding: 6px 10px;
      overflow: hidden;
    }}
    .brand img {{ width: 100%; height: 100%; object-fit: contain; }}
    h1 {{ margin: 0; font-size: 24px; font-weight: 750; letter-spacing: 0; }}
    .subtitle {{ margin-top: 4px; color: rgba(255,255,255,.82); }}
    .stamp {{ text-align: right; color: rgba(255,255,255,.82); font-size: 12px; }}
    .wrap {{ padding: 18px 22px 26px; max-width: 1680px; margin: 0 auto; }}
    .filters {{
      display: grid;
      grid-template-columns: repeat(7, minmax(130px, 1fr)) auto;
      gap: 10px;
      align-items: end;
      margin-bottom: 14px;
    }}
    .data-actions {{
      display: flex;
      flex-wrap: wrap;
      gap: 10px;
      align-items: center;
      margin-top: 12px;
      padding-top: 12px;
      border-top: 1px solid #e7edf2;
    }}
    .file-action {{
      position: relative;
      display: inline-flex;
      align-items: center;
      min-height: 32px;
      border-radius: 7px;
      padding: 7px 10px;
      background: var(--remo-green);
      color: #fff;
      font-weight: 800;
      font-size: 12px;
      cursor: pointer;
      overflow: hidden;
    }}
    .file-action input {{
      position: absolute;
      inset: 0;
      opacity: 0;
      cursor: pointer;
    }}
    .data-status {{ color: var(--muted); font-size: 12px; }}
    .data-actions button {{
      min-height: 32px;
      padding: 7px 10px;
      font-size: 12px;
    }}
    label {{ display: block; color: var(--muted); font-size: 11px; font-weight: 700; text-transform: uppercase; margin: 0 0 4px; }}
    select, input {{
      width: 100%;
      min-height: 38px;
      border: 1px solid var(--line);
      border-radius: 7px;
      background: #fff;
      color: var(--ink);
      padding: 8px 9px;
      outline: none;
    }}
    select:focus, input:focus {{ border-color: var(--remo-green); box-shadow: 0 0 0 3px rgba(41,199,123,.16); }}
    button {{
      border: 0;
      border-radius: 7px;
      min-height: 38px;
      padding: 8px 12px;
      cursor: pointer;
      font-weight: 700;
      color: #fff;
      background: var(--remo-blue);
    }}
    button.secondary {{ background: #6f7f8a; }}
    .active-filter {{
      display: flex;
      flex-wrap: wrap;
      gap: 8px;
      min-height: 24px;
      margin: 2px 0 14px;
    }}
    .chip {{
      display: inline-flex;
      align-items: center;
      gap: 7px;
      border-radius: 999px;
      padding: 4px 9px;
      background: #dff6eb;
      color: #11543a;
      font-size: 12px;
      font-weight: 700;
    }}
    .kpis {{ display: grid; grid-template-columns: repeat(5, 1fr); gap: 12px; margin-bottom: 14px; }}
    .kpi {{
      background: var(--panel);
      border: 1px solid var(--line);
      border-left: 6px solid var(--remo-green);
      border-radius: 8px;
      box-shadow: var(--shadow);
      padding: 13px 14px;
      min-height: 92px;
    }}
    .kpi:nth-child(2) {{ border-left-color: #29c77b; }}
    .kpi:nth-child(3) {{ border-left-color: #d94f45; }}
    .kpi:nth-child(4) {{ border-left-color: #ffd84d; }}
    .kpi:nth-child(5) {{ border-left-color: #4b8dbd; }}
    .kpi .label {{ color: var(--muted); font-size: 12px; font-weight: 700; text-transform: uppercase; }}
    .kpi .value {{ font-size: 30px; font-weight: 800; margin-top: 6px; color: var(--remo-blue-2); }}
    .kpi .note {{ color: var(--muted); font-size: 12px; margin-top: 1px; }}
    .grid {{
      display: grid;
      grid-template-columns: 1.12fr .88fr;
      gap: 14px;
    }}
    .charts {{
      display: grid;
      grid-template-columns: repeat(2, minmax(0, 1fr));
      gap: 14px;
    }}
    .panel {{
      background: var(--panel);
      border: 1px solid var(--line);
      border-radius: 8px;
      box-shadow: var(--shadow);
      padding: 14px;
      min-width: 0;
    }}
    .panel h2 {{
      margin: 0 0 10px;
      font-size: 15px;
      color: var(--remo-blue-2);
      display: flex;
      justify-content: space-between;
      gap: 10px;
      align-items: center;
    }}
    .tag-yellow {{
      background: #fff3bd;
      color: #6d5200;
      border: 1px solid #ffe17b;
      padding: 3px 8px;
      border-radius: 999px;
      font-size: 11px;
      white-space: nowrap;
    }}
    svg.chart {{ width: 100%; height: 260px; display: block; }}
    .bar-label, .axis, .legend {{ fill: #536373; font-size: 11px; }}
    .bar-value {{ fill: #223342; font-size: 11px; font-weight: 700; }}
    .clickable {{ cursor: pointer; }}
    .clickable:hover {{ opacity: .82; }}
    #geoMap {{ width: 100%; height: 545px; display: block; background: #edf4f7; border-radius: 7px; border: 1px solid var(--line); overflow: hidden; }}
    .leaflet-container {{ font: 12px/1.35 "Segoe UI", Arial, sans-serif; }}
    .leaflet-control-attribution {{ font-size: 10px; }}
    .marker-pin {{
      width: 13px;
      height: 13px;
      border-radius: 50%;
      border: 1.8px solid #59251f;
      box-shadow: 0 1px 4px rgba(0,0,0,.35);
    }}
    .marker-executado {{ background: #4f9d45; }}
    .marker-pendente {{ background: #9e160f; }}
    .map-fallback {{
      display: flex;
      align-items: center;
      justify-content: center;
      height: 100%;
      color: var(--muted);
      padding: 22px;
      text-align: center;
    }}
    .map-note {{ color: var(--muted); margin-top: 8px; font-size: 12px; }}
    .table-wrap {{ max-height: 255px; overflow: auto; border: 1px solid var(--line); border-radius: 7px; }}
    table {{ width: 100%; border-collapse: collapse; font-size: 12px; }}
    th, td {{ padding: 8px 9px; border-bottom: 1px solid #e7edf2; text-align: left; }}
    th {{ position: sticky; top: 0; background: #f7fafc; color: #516170; z-index: 1; }}
    tr:hover td {{ background: #f4fbf7; }}
    .tooltip {{
      position: fixed;
      pointer-events: none;
      z-index: 20;
      background: #0f2434;
      color: #fff;
      padding: 8px 10px;
      border-radius: 7px;
      font-size: 12px;
      max-width: 260px;
      display: none;
      box-shadow: var(--shadow);
    }}
    @media (max-width: 1250px) {{
      .filters {{ grid-template-columns: repeat(4, 1fr); }}
      .kpis {{ grid-template-columns: repeat(3, 1fr); }}
      .grid {{ grid-template-columns: 1fr; }}
    }}
    @media (max-width: 760px) {{
      .topbar {{ grid-template-columns: 1fr; }}
      .stamp {{ text-align: left; }}
      .filters, .charts, .kpis {{ grid-template-columns: 1fr; }}
      .data-actions, .data-actions button, .file-action {{ width: 100%; }}
      .data-actions button, .file-action {{ justify-content: center; }}
      .wrap {{ padding: 14px; }}
      svg.chart {{ height: 285px; }}
      #geoMap {{ height: 420px; }}
    }}
  </style>
</head>
<body>
  <header class="topbar">
    <div class="brand"><img src="{logo_data}" alt="REMO Engenharia"></div>
    <div>
      <h1>Acompanhamento de Anomalias Equatorial</h1>
      <div class="subtitle">Pendências, execução, taxonomia e distribuição geográfica das ocorrências encaminhadas à REMO</div>
    </div>
    <div class="stamp">Base: Geral anomalias<br>Gerado em {generated}</div>
  </header>

  <main class="wrap">
    <section class="filters" aria-label="Filtros">
      <div><label for="dateStart">Data inicial</label><input id="dateStart" type="date"></div>
      <div><label for="dateEnd">Data final</label><input id="dateEnd" type="date"></div>
      <div><label for="instalacao">Instalação</label><select id="instalacao"></select></div>
      <div><label for="se">SE</label><select id="se"></select></div>
      <div><label for="execucao">Execução</label><select id="execucao"></select></div>
      <div><label for="tipoTrecho">Tipo do trecho</label><select id="tipoTrecho"></select></div>
      <div><label for="taxonomia">Taxonomia</label><select id="taxonomia"></select></div>
      <button id="reset" type="button" class="secondary">Limpar</button>
    </section>
    <div id="activeFilters" class="active-filter"></div>

    <section class="kpis">
      <div class="kpi"><div class="label">Total de anomalias</div><div class="value" id="kpiTotal">0</div><div class="note">Registros filtrados</div></div>
      <div class="kpi"><div class="label">Executadas</div><div class="value" id="kpiExec">0</div><div class="note" id="noteExec">0%</div></div>
      <div class="kpi"><div class="label">Pendentes</div><div class="value" id="kpiPend">0</div><div class="note" id="notePend">0%</div></div>
      <div class="kpi"><div class="label">SEs atendidas</div><div class="value" id="kpiSe">0</div><div class="note">Subestações distintas</div></div>
      <div class="kpi"><div class="label">Instalações</div><div class="value" id="kpiInst">0</div><div class="note">Instalações distintas</div></div>
    </section>

    <section class="grid">
      <div class="charts">
        <div class="panel"><h2>Executadas x Pendentes</h2><svg id="chartStatus" class="chart"></svg></div>
        <div class="panel"><h2>Ocorrências por SE</h2><svg id="chartSe" class="chart"></svg></div>
        <div class="panel"><h2>Tipo de trecho</h2><svg id="chartTrecho" class="chart"></svg></div>
        <div class="panel"><h2>Taxonomia</h2><svg id="chartTaxonomia" class="chart"></svg></div>
        <div class="panel"><h2>Evolução por data de registro</h2><svg id="chartTempo" class="chart"></svg></div>
        <div class="panel"><h2>Criticidade</h2><svg id="chartCrit" class="chart"></svg></div>
      </div>
      <div class="panel">
        <h2>Mapa geográfico das ocorrências</h2>
        <div id="geoMap"></div>
        <div class="map-note">Arraste e use o zoom para navegar pelo mapa. Verde indica executada; vermelho indica pendente.</div>
        <section class="data-actions" aria-label="Atualização da base">
          <label class="file-action">Importar Excel<input id="excelUpload" type="file" accept=".xlsx,.xls"></label>
          <button id="exportCsv" type="button">Exportar CSV</button>
          <button id="restoreBase" type="button" class="secondary">Restaurar base</button>
          <span id="dataStatus" class="data-status"></span>
        </section>
      </div>
    </section>

    <section class="panel" style="margin-top:14px">
      <h2>Principais ocorrências e alimentadores</h2>
      <div class="table-wrap">
        <table>
          <thead><tr><th>ID</th><th>OS</th><th>Defeito</th><th>SE</th><th>Instalação</th><th>Trecho</th><th>Taxonomia</th><th>Execução</th><th>Registro</th></tr></thead>
          <tbody id="detailRows"></tbody>
        </table>
      </div>
    </section>
  </main>
  <div id="tooltip" class="tooltip"></div>

  <script src="https://unpkg.com/leaflet@1.9.4/dist/leaflet.js"></script>
  <script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
  <script>
    const ORIGINAL_DATA = {payload};
    let DATA = loadStoredData() || ORIGINAL_DATA;
    const state = {{
      dateStart: "", dateEnd: "", instalacao: "", se: "", execucao: "",
      tipoTrecho: "", taxonomia: "", crit: "", month: ""
    }};
    const palette = ["#10496f", "#29c77b", "#ffd84d", "#6f7f8a", "#d94f45", "#4b8dbd", "#7c5cc4"];
    const statusColor = value => String(value).toLowerCase().includes("pend") ? "#d94f45" : "#29c77b";
    const fmt = n => Number(n || 0).toLocaleString("pt-BR");
    const pct = (n, d) => d ? ((n / d) * 100).toLocaleString("pt-BR", {{maximumFractionDigits: 1}}) + "%" : "0%";
    const norm = v => (v === null || v === undefined || v === "" || v === "#N/A") ? "Não informado" : String(v);
    const els = id => document.getElementById(id);
    const tooltip = els("tooltip");

    function loadStoredData() {{
      try {{
        const raw = localStorage.getItem("remoDashboardData");
        return raw ? JSON.parse(raw) : null;
      }} catch {{
        return null;
      }}
    }}

    function saveStoredData(rows) {{
      localStorage.setItem("remoDashboardData", JSON.stringify(rows));
      localStorage.setItem("remoDashboardUpdatedAt", new Date().toISOString());
    }}

    function statusOf(row) {{
      const raw = norm(row.execucao);
      if (raw !== "Não informado") return raw;
      return norm(row.pendente).toLowerCase().includes("pend") ? "Pendente" : "Não informado";
    }}

    function initSelect(id, field, label = "Todos") {{
      const select = els(id);
      const values = [...new Set(DATA.map(row => field(row)).filter(Boolean).map(norm))].sort((a, b) => a.localeCompare(b, "pt-BR"));
      select.innerHTML = `<option value="">${{label}}</option>` + values.map(v => `<option value="${{escapeHtml(v)}}">${{escapeHtml(v)}}</option>`).join("");
      select.onchange = () => {{ state[id] = select.value; update(); }};
    }}

    function escapeHtml(value) {{
      return String(value).replace(/[&<>"']/g, ch => ({{"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}}[ch]));
    }}

    function setup() {{
      resetControls();
      ["dateStart", "dateEnd"].forEach(id => els(id).onchange = () => {{ state[id] = els(id).value; update(); }});
      els("reset").addEventListener("click", () => {{
        resetControls();
        update();
      }});
      els("excelUpload").addEventListener("change", handleExcelUpload);
      els("exportCsv").addEventListener("click", exportFilteredCsv);
      els("restoreBase").addEventListener("click", () => {{
        localStorage.removeItem("remoDashboardData");
        localStorage.removeItem("remoDashboardUpdatedAt");
        DATA = ORIGINAL_DATA;
        resetControls();
        update();
      }});
      window.addEventListener("resize", update);
      update();
    }}

    function resetControls() {{
      const dates = DATA.map(r => r.dataRegistro).filter(Boolean).sort();
      Object.assign(state, {{dateStart: dates[0] || "", dateEnd: dates[dates.length - 1] || "", instalacao:"", se:"", execucao:"", tipoTrecho:"", taxonomia:"", crit:"", month:""}});
      els("dateStart").value = state.dateStart;
      els("dateEnd").value = state.dateEnd;
      initSelect("instalacao", r => r.instalacao, "Todas");
      initSelect("se", r => r.se, "Todas");
      initSelect("execucao", statusOf, "Todos");
      initSelect("tipoTrecho", r => r.tipoTrecho, "Todos");
      initSelect("taxonomia", r => r.taxonomia, "Todas");
      updateDataStatus();
    }}

    function updateDataStatus() {{
      const storedAt = localStorage.getItem("remoDashboardUpdatedAt");
      els("dataStatus").textContent = storedAt
        ? `Base importada localmente em ${{new Date(storedAt).toLocaleString("pt-BR")}} · ${{fmt(DATA.length)}} registros`
        : `Base inicial embutida · ${{fmt(DATA.length)}} registros`;
    }}

    async function handleExcelUpload(event) {{
      const file = event.target.files && event.target.files[0];
      if (!file) return;
      if (!window.XLSX) {{
        alert("Nao foi possivel carregar o leitor de Excel. Verifique a conexao com a internet e tente novamente.");
        event.target.value = "";
        return;
      }}
      try {{
        els("dataStatus").textContent = `Lendo ${{file.name}}...`;
        const buffer = await file.arrayBuffer();
        const workbook = XLSX.read(buffer, {{ type: "array", cellDates: true }});
        const sheetName = workbook.SheetNames.find(name => name.trim().toLowerCase() === "geral anomalias") || workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];
        const table = XLSX.utils.sheet_to_json(sheet, {{ header: 1, defval: null, raw: true }});
        const rows = parseWorksheetRows(table);
        if (!rows.length) throw new Error("A planilha nao possui registros validos.");
        DATA = rows;
        saveStoredData(DATA);
        resetControls();
        update();
      }} catch (error) {{
        console.error(error);
        alert(`Nao consegui importar essa planilha: ${{error.message}}`);
        updateDataStatus();
      }} finally {{
        event.target.value = "";
      }}
    }}

    function parseWorksheetRows(table) {{
      if (!table.length) return [];
      const headers = table[0].map(v => normHeader(v));
      const index = name => headers.indexOf(normHeader(name));
      const required = ["Defeito", "SE", "Latitude", "Longitude", "Instalação", "ID Anomalia", "Execução", "DATA DE REGISTRO", "Tipo de trecho", "Taxonomia"];
      const missing = required.filter(name => index(name) === -1);
      if (missing.length) throw new Error(`Colunas obrigatorias ausentes: ${{missing.join(", ")}}`);
      const val = (row, name) => {{
        const i = index(name);
        return i >= 0 ? cleanCell(row[i]) : "";
      }};
      return table.slice(1).map(row => {{
        const lat = toNumber(val(row, "Latitude"));
        const lon = toNumber(val(row, "Longitude"));
        return {{
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
        }};
      }}).filter(row => row.id || row.se || row.instalacao || row.defeito);
    }}

    function normHeader(value) {{
      return String(value || "").trim().toLowerCase().normalize("NFD").replace(/[\\u0300-\\u036f]/g, "");
    }}

    function cleanCell(value) {{
      if (value === null || value === undefined) return "";
      if (value instanceof Date) return value.toISOString().slice(0, 10);
      const text = String(value).trim();
      return text === "#NAME?" || text === "undefined" || text === "null" ? "" : text;
    }}

    function toNumber(value) {{
      if (value === "" || value === null || value === undefined) return null;
      const number = typeof value === "number" ? value : Number(String(value).replace(",", "."));
      return Number.isFinite(number) ? number : null;
    }}

    function toDateText(value) {{
      if (!value) return "";
      if (value instanceof Date) return value.toISOString().slice(0, 10);
      if (typeof value === "number") {{
        const date = XLSX.SSF.parse_date_code(value);
        if (date) return `${{date.y}}-${{String(date.m).padStart(2, "0")}}-${{String(date.d).padStart(2, "0")}}`;
      }}
      const text = String(value).trim();
      if (/^\\d+(\\.\\d+)?$/.test(text) && Number(text) > 20000 && window.XLSX) {{
        const date = XLSX.SSF.parse_date_code(Number(text));
        if (date) return `${{date.y}}-${{String(date.m).padStart(2, "0")}}-${{String(date.d).padStart(2, "0")}}`;
      }}
      const br = text.match(/^(\\d{{1,2}})[\\/.-](\\d{{1,2}})[\\/.-](\\d{{2,4}})$/);
      if (br) {{
        const y = br[3].length === 2 ? `20${{br[3]}}` : br[3];
        return `${{y}}-${{br[2].padStart(2, "0")}}-${{br[1].padStart(2, "0")}}`;
      }}
      const iso = text.match(/^(\\d{{4}})-(\\d{{1,2}})-(\\d{{1,2}})/);
      return iso ? `${{iso[1]}}-${{iso[2].padStart(2, "0")}}-${{iso[3].padStart(2, "0")}}` : text;
    }}

    function filtered() {{
      return DATA.filter(r => {{
        if (state.dateStart && r.dataRegistro && r.dataRegistro < state.dateStart) return false;
        if (state.dateEnd && r.dataRegistro && r.dataRegistro > state.dateEnd) return false;
        if (state.instalacao && norm(r.instalacao) !== state.instalacao) return false;
        if (state.se && norm(r.se) !== state.se) return false;
        if (state.execucao && norm(statusOf(r)) !== state.execucao) return false;
        if (state.tipoTrecho && norm(r.tipoTrecho) !== state.tipoTrecho) return false;
        if (state.taxonomia && norm(r.taxonomia) !== state.taxonomia) return false;
        if (state.crit && norm(r.crit) !== state.crit) return false;
        if (state.month && (!r.dataRegistro || !r.dataRegistro.startsWith(state.month))) return false;
        return true;
      }});
    }}

    function update() {{
      const rows = filtered();
      renderActiveFilters();
      renderKpis(rows);
      renderBar("chartStatus", group(rows, statusOf), "execucao", statusColor);
      renderBar("chartSe", group(rows, r => norm(r.se)), "se", null, 10);
      renderDonut("chartTrecho", group(rows, r => norm(r.tipoTrecho)), "tipoTrecho");
      renderDonut("chartTaxonomia", group(rows, r => norm(r.taxonomia)), "taxonomia");
      renderLine("chartTempo", groupMonth(rows));
      renderBar("chartCrit", group(rows, r => norm(r.crit)), "crit", null, 8);
      renderMap(rows);
      renderTable(rows);
    }}

    function renderActiveFilters() {{
      const chips = [];
      const labels = {{instalacao:"Instalação", se:"SE", execucao:"Execução", tipoTrecho:"Trecho", taxonomia:"Taxonomia"}};
      Object.entries(labels).forEach(([key, label]) => {{ if (state[key]) chips.push(`${{label}}: ${{state[key]}}`); }});
      if (state.crit) chips.push(`Criticidade: ${{state.crit}}`);
      if (state.month) chips.push(`Mês: ${{state.month.slice(5)}}/${{state.month.slice(0,4)}}`);
      els("activeFilters").innerHTML = chips.map(c => `<span class="chip">${{escapeHtml(c)}}</span>`).join("");
    }}

    function renderKpis(rows) {{
      const total = rows.length;
      const executed = rows.filter(r => statusOf(r).toLowerCase().includes("execut")).length;
      const pending = rows.filter(r => statusOf(r).toLowerCase().includes("pend")).length;
      els("kpiTotal").textContent = fmt(total);
      els("kpiExec").textContent = fmt(executed);
      els("noteExec").textContent = pct(executed, total);
      els("kpiPend").textContent = fmt(pending);
      els("notePend").textContent = pct(pending, total);
      els("kpiSe").textContent = fmt(new Set(rows.map(r => norm(r.se)).filter(v => v !== "Não informado")).size);
      els("kpiInst").textContent = fmt(new Set(rows.map(r => norm(r.instalacao)).filter(v => v !== "Não informado")).size);
    }}

    function group(rows, accessor) {{
      const m = new Map();
      rows.forEach(r => {{ const key = norm(accessor(r)); m.set(key, (m.get(key) || 0) + 1); }});
      return [...m.entries()].sort((a,b) => b[1] - a[1]);
    }}

    function groupMonth(rows) {{
      const m = new Map();
      rows.forEach(r => {{
        if (!r.dataRegistro) return;
        const key = r.dataRegistro.slice(0, 7);
        m.set(key, (m.get(key) || 0) + 1);
      }});
      return [...m.entries()].sort((a,b) => a[0].localeCompare(b[0]));
    }}

    function svgRoot(id) {{
      const svg = els(id);
      const box = svg.getBoundingClientRect();
      const w = Math.max(320, box.width || 480), h = 260;
      svg.setAttribute("viewBox", `0 0 ${{w}} ${{h}}`);
      svg.innerHTML = "";
      return {{svg, w, h}};
    }}

    function toggleFilter(key, value) {{
      if (!key) return;
      state[key] = state[key] === value ? "" : value;
      const control = els(key);
      if (control) control.value = state[key];
      update();
    }}

    function renderBar(id, entries, filterKey, colorFn, limit = 6) {{
      const {{svg, w, h}} = svgRoot(id);
      const data = entries.slice(0, limit);
      const left = 116, right = 46, top = 12, rowH = (h - top - 18) / Math.max(data.length, 1);
      const max = Math.max(...data.map(d => d[1]), 1);
      data.forEach(([name, value], i) => {{
        const y = top + i * rowH + 6;
        const bw = (w - left - right) * value / max;
        const color = colorFn ? colorFn(name) : palette[i % palette.length];
        svg.insertAdjacentHTML("beforeend", `<text class="bar-label" x="4" y="${{y + rowH/2 + 4}}">${{escapeHtml(short(name, 18))}}</text>`);
        svg.insertAdjacentHTML("beforeend", `<rect class="clickable" data-filter="${{filterKey || ""}}" data-value="${{escapeHtml(name)}}" x="${{left}}" y="${{y}}" width="${{Math.max(2,bw)}}" height="${{Math.max(16,rowH-9)}}" rx="4" fill="${{color}}"></rect>`);
        svg.insertAdjacentHTML("beforeend", `<text class="bar-value" x="${{left + bw + 7}}" y="${{y + rowH/2 + 4}}">${{fmt(value)}}</text>`);
      }});
      svg.querySelectorAll("[data-filter]").forEach(el => el.addEventListener("click", () => {{
        const key = el.dataset.filter;
        if (!key) return;
        toggleFilter(key, el.dataset.value);
      }}));
    }}

    function renderDonut(id, entries, filterKey) {{
      const {{svg, w, h}} = svgRoot(id);
      const total = entries.reduce((s, d) => s + d[1], 0) || 1;
      const compact = w < 430;
      const cx = compact ? w / 2 : 106;
      const cy = compact ? 96 : h / 2;
      const r = compact ? 58 : 68;
      const sw = compact ? 22 : 25;
      let start = -Math.PI / 2;
      entries.slice(0, 6).forEach(([name, value], i) => {{
        const angle = (value / total) * Math.PI * 2;
        const end = start + angle;
        if (value === total) {{
          svg.insertAdjacentHTML("beforeend", `<circle class="clickable" data-filter="${{filterKey}}" data-value="${{escapeHtml(name)}}" cx="${{cx}}" cy="${{cy}}" r="${{r}}" fill="none" stroke="${{palette[i % palette.length]}}" stroke-width="${{sw}}"></circle>`);
        }} else {{
          const path = arc(cx, cy, r, start, end);
          svg.insertAdjacentHTML("beforeend", `<path class="clickable" data-filter="${{filterKey}}" data-value="${{escapeHtml(name)}}" d="${{path}}" fill="none" stroke="${{palette[i % palette.length]}}" stroke-width="${{sw}}" stroke-linecap="butt"></path>`);
        }}
        const ly = compact ? 188 + i * 22 : 32 + i * 28;
        const legendX = compact ? 40 : Math.max(cx + r + 28, Math.min(w - 210, 248));
        const labelSize = compact ? 30 : 28;
        svg.insertAdjacentHTML("beforeend", `<rect class="clickable" data-filter="${{filterKey}}" data-value="${{escapeHtml(name)}}" x="${{legendX}}" y="${{ly - 10}}" width="12" height="12" rx="2" fill="${{palette[i % palette.length]}}"></rect><text class="legend clickable" data-filter="${{filterKey}}" data-value="${{escapeHtml(name)}}" x="${{legendX + 18}}" y="${{ly}}">${{escapeHtml(short(name, labelSize))}} · ${{fmt(value)}} · ${{pct(value, total)}}</text>`);
        start = end;
      }});
      svg.insertAdjacentHTML("beforeend", `<text x="${{cx}}" y="${{cy - 4}}" text-anchor="middle" font-size="24" font-weight="800" fill="#0b324e">${{fmt(total)}}</text><text x="${{cx}}" y="${{cy + 17}}" text-anchor="middle" class="legend">ocorrências</text>`);
      svg.querySelectorAll("[data-filter]").forEach(el => el.addEventListener("click", () => {{
        toggleFilter(el.dataset.filter, el.dataset.value);
      }}));
    }}

    function arc(cx, cy, r, a0, a1) {{
      const x0 = cx + r * Math.cos(a0), y0 = cy + r * Math.sin(a0);
      const x1 = cx + r * Math.cos(a1), y1 = cy + r * Math.sin(a1);
      const large = a1 - a0 > Math.PI ? 1 : 0;
      return `M ${{x0}} ${{y0}} A ${{r}} ${{r}} 0 ${{large}} 1 ${{x1}} ${{y1}}`;
    }}

    function renderLine(id, entries) {{
      const {{svg, w, h}} = svgRoot(id);
      const left = 42, right = 22, top = 18, bottom = 34;
      const max = Math.max(...entries.map(d => d[1]), 1);
      const x = i => left + (w - left - right) * (entries.length <= 1 ? .5 : i / (entries.length - 1));
      const y = v => top + (h - top - bottom) * (1 - v / max);
      for (let i = 0; i <= 4; i++) {{
        const yy = top + (h - top - bottom) * i / 4;
        svg.insertAdjacentHTML("beforeend", `<line x1="${{left}}" y1="${{yy}}" x2="${{w-right}}" y2="${{yy}}" stroke="#e1e9ee"></line>`);
      }}
      const points = entries.map((d,i) => `${{x(i)}},${{y(d[1])}}`).join(" ");
      svg.insertAdjacentHTML("beforeend", `<polyline points="${{points}}" fill="none" stroke="#10496f" stroke-width="3"></polyline>`);
      entries.forEach((d,i) => {{
        const xx = x(i), yy = y(d[1]);
        svg.insertAdjacentHTML("beforeend", `<circle class="clickable" data-filter="month" data-value="${{escapeHtml(d[0])}}" cx="${{xx}}" cy="${{yy}}" r="5" fill="#29c77b"><title>${{d[0]}}: ${{d[1]}}</title></circle>`);
        svg.insertAdjacentHTML("beforeend", `<text class="bar-value" x="${{xx}}" y="${{Math.max(12, yy - 8)}}" text-anchor="middle">${{fmt(d[1])}}</text>`);
        if (i % Math.ceil(entries.length / 6 || 1) === 0) svg.insertAdjacentHTML("beforeend", `<text class="axis" x="${{x(i)}}" y="${{h-10}}" text-anchor="middle">${{d[0].slice(5)}}/${{d[0].slice(2,4)}}</text>`);
      }});
      svg.querySelectorAll("[data-filter]").forEach(el => el.addEventListener("click", () => toggleFilter(el.dataset.filter, el.dataset.value)));
    }}

    let map = null;
    let markerLayer = null;
    function renderMap(rows) {{
      const pts = rows.filter(r => typeof r.lat === "number" && typeof r.lon === "number");
      const container = els("geoMap");
      if (!window.L) {{
        container.innerHTML = `<div class="map-fallback">Mapa interativo indisponível porque a biblioteca online não carregou. Verifique a conexão à internet para habilitar navegação, zoom e base geográfica.</div>`;
        return;
      }}
      if (!map) {{
        map = L.map("geoMap", {{ zoomControl: true, preferCanvas: true }});
        L.tileLayer("https://{{s}}.basemaps.cartocdn.com/light_all/{{z}}/{{x}}/{{y}}{{r}}.png", {{
          maxZoom: 19,
          attribution: '&copy; OpenStreetMap &copy; CARTO'
        }}).addTo(map);
        markerLayer = L.layerGroup().addTo(map);
      }}
      markerLayer.clearLayers();
      if (!pts.length) return;
      const bounds = [];
      pts.forEach(r => {{
        const isPending = statusOf(r).toLowerCase().includes("pend");
        const marker = L.circleMarker([r.lat, r.lon], {{
          radius: 5,
          color: "#56231d",
          weight: 1,
          fillColor: isPending ? "#9e160f" : "#4f9d45",
          fillOpacity: .86
        }}).bindPopup(`<b>${{escapeHtml(statusOf(r))}}</b><br>SE: ${{escapeHtml(r.se)}}<br>Instalação: ${{escapeHtml(r.instalacao)}}<br>Trecho: ${{escapeHtml(r.tipoTrecho)}}<br>Taxonomia: ${{escapeHtml(r.taxonomia)}}<br>ID: ${{escapeHtml(r.id)}}`);
        marker.addTo(markerLayer);
        bounds.push([r.lat, r.lon]);
      }});
      if (bounds.length && !map._userMoved) {{
        map.fitBounds(bounds, {{ padding: [24, 24] }});
      }}
      map.on("dragstart zoomstart", () => map._userMoved = true);
    }}

    function topName(list, accessor) {{
      const items = group(list, accessor);
      return items[0] ? items[0][0] : "Não informado";
    }}

    function renderTable(rows) {{
      els("detailRows").innerHTML = rows.slice(0, 80).map(r => `<tr><td>${{escapeHtml(r.id)}}</td><td>${{escapeHtml(r.os)}}</td><td>${{escapeHtml(short(r.defeito, 52))}}</td><td>${{escapeHtml(r.se)}}</td><td>${{escapeHtml(r.instalacao)}}</td><td>${{escapeHtml(r.tipoTrecho)}}</td><td>${{escapeHtml(r.taxonomia)}}</td><td>${{escapeHtml(statusOf(r))}}</td><td>${{escapeHtml(r.dataRegistro)}}</td></tr>`).join("");
    }}

    function exportFilteredCsv() {{
      const rows = filtered();
      const columns = [
        ["id", "ID Anomalia"],
        ["os", "OS"],
        ["defeito", "Defeito"],
        ["se", "SE"],
        ["alimentador", "Alimentador"],
        ["instalacao", "Instalação"],
        ["poste", "Poste"],
        ["crit", "Criticidade"],
        ["tipoTrecho", "Tipo de trecho"],
        ["taxonomia", "Taxonomia"],
        ["tipoAnomalia", "Tipo de anomalia"],
        ["execucao", "Execução"],
        ["dataRegistro", "Data de registro"],
        ["dataExecucao", "Data de execução"],
        ["lat", "Latitude"],
        ["lon", "Longitude"]
      ];
      const lines = [columns.map(c => csvCell(c[1])).join(";")];
      rows.forEach(row => {{
        lines.push(columns.map(([key]) => csvCell(key === "execucao" ? statusOf(row) : row[key])).join(";"));
      }});
      const blob = new Blob(["\\ufeff" + lines.join("\\r\\n")], {{ type: "text/csv;charset=utf-8" }});
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = `anomalias_filtradas_${{new Date().toISOString().slice(0, 10)}}.csv`;
      document.body.appendChild(a);
      a.click();
      a.remove();
      URL.revokeObjectURL(url);
    }}

    function csvCell(value) {{
      const text = value === null || value === undefined ? "" : String(value);
      return `"${{text.replace(/"/g, '""')}}"`;
    }}

    function short(value, size) {{
      const text = norm(value);
      return text.length > size ? text.slice(0, size - 1) + "…" : text;
    }}

    setup();
  </script>
</body>
</html>"""


def main():
    rows = load_rows()
    OUT.write_text(build_html(rows, encode_logo()), encoding="utf-8")
    print(OUT)
    print(f"{len(rows)} registros exportados")


if __name__ == "__main__":
    main()
