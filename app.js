// app.js
// Validação de Folha por Verba (mensal) usando 6 Sigma (média ± 3σ).
// - Importa 1 Excel com colunas fixas (Código/Descrição etc) e colunas mensais no formato "MMM/AA - <Métrica>".
// - Suporta exports com blocos mensais (ex.: "JUN/25 - Hora", "JUN/25 - Valor", "JUN/25 - Dt Pgto").
// - Permite filtrar por verba, selecionar mês de referência, definir janela histórica e ignorar zeros.
// - Agrupa dinamicamente por colunas escolhidas (ex.: Estabelecimento, Centro de Custo, Matrícula).
// - Cria filtros automáticos para colunas extras marcadas.
// - Exibe grid com todos os meses e um "bandviz" com faixa ±1σ/±2σ/±3σ e marcador no mês de referência.
// - Exporta TXT e Excel conforme filtros aplicados.
// Igarapé Digital.

(() => {
  "use strict";

  const STORAGE_KEY = "igarape_verbas_sigma_v2";
  const STORAGE_META = "igarape_verbas_sigma_meta_v2";
  const STORAGE_COLS = "igarape_verbas_sigma_cols_v2";
  const STORAGE_GROUP = "igarape_verbas_sigma_group_v2";
  const STORAGE_METRIC = "igarape_verbas_sigma_metric_v2";

  // Configuração (conforme layout solicitado)
  // - Mantém o app inteiro, mas limita "Filtros extras" a C.R. e Clas. (quando existirem).
  // - Sugere defaults: colunas visíveis C.R., Clas., Nome, Processo; agrupamento por Empresa e CPF.
  const FIXED_FILTER_KEYS = ["cr","clas"];                // apenas filtros extras
  const FIXED_VISIBLE_KEYS_DEFAULT = ["cr","clas","nome","processo"]; // defaults de exibição
  const FIXED_GROUP_KEYS_DEFAULT = ["empresa","cpf"];     // defaults de agrupamento


  const el = (id) => document.getElementById(id);

  const ui = {
    fileInput: el("fileInput"),
    btnExportTxt: el("btnExportTxt"),
    btnExportXlsx: el("btnExportXlsx"),
    btnClear: el("btnClear"),
    btnApply: el("btnApply"),
    btnReset: el("btnReset"),

    fSearch: el("fSearch"),
    fVerba: el("fVerba"),
    fMetric: el("fMetric"),
    fRefMonth: el("fRefMonth"),
    fWindow: el("fWindow"),
    fIgnoreZeros: el("fIgnoreZeros"),
    fStatus: el("fStatus"),
    fMinZ: el("fMinZ"),
    fMaxZ: el("fMaxZ"),

    extraFilters: el("extraFilters"),
    colPickerBody: el("colPickerBody"),
    groupPickerBody: el("groupPickerBody"),

    dataInfo: el("dataInfo"),
    tableInfo: el("tableInfo"),

    kpiTotal: el("kpiTotal"),
    kpiOk: el("kpiOk"),
    kpiWarn: el("kpiWarn"),
    kpiOut: el("kpiOut"),
    kpiAvgZ: el("kpiAvgZ"),

    tblHead: el("tblHead"),
    tblBody: el("tblBody"),
    diagBox: el("diagBox"),
    buildInfo: el("buildInfo"),

    monthTotals: el("monthTotals"),
  };

  const state = {
    rawRows: [],
    meta: null,

    // meses detectados (ordem cronológica)
    baseMonths: [], // [{key:'2026-02', label:'FEV/26', metrics:{Valor:'FEV/26 - Valor', Hora:'FEV/26 - Hora'}}]
    months: [],     // derivado de baseMonths + métrica selecionada => [{key,label,header}]

    metricOptions: [],
    metric: "",

    // colunas base detectadas
    colCode: null,
    colDesc: null,

    // colunas disponíveis para agrupamento / extras
    baseColumns: [],
    extraColumns: [],
    visibleExtraColumns: [],
    groupByColumns: [],

    // linhas agregadas (uma por grupo)
    groups: [],
    filtered: [],

    extraFilterValues: {},
    extraFilterModes: {},
    sort: { key: null, dir: 1 },
  };

  function nowISO() { return new Date().toISOString(); }
  function safeText(v) { return (v === null || v === undefined) ? "" : String(v).trim(); }

  function normKey(s) {
    return safeText(s).toLowerCase().replace(/[^a-z0-9]/g, "");
  }

  function pickColsByKeys(cols, keys) {
    const out = [];
    const seen = new Set();
    for (const k of keys) {
      const found = (cols || []).find(c => normKey(c) === k);
      if (found && !seen.has(found)) { out.push(found); seen.add(found); }
    }
    return out;
  }


  function parseNumber(v) {
    if (v === null || v === undefined) return null;
    if (typeof v === "number") return Number.isFinite(v) ? v : null;
    const s = String(v).trim();
    if (!s) return null;
    const s1 = s.replace(/\s/g, "");
    const hasComma = s1.includes(",");
    const hasDot = s1.includes(".");
    let normalized = s1;

    if (hasComma && hasDot) {
      const lastComma = s1.lastIndexOf(",");
      const lastDot = s1.lastIndexOf(".");
      if (lastComma > lastDot) normalized = s1.replace(/\./g, "").replace(",", ".");
      else normalized = s1.replace(/,/g, "");
    } else if (hasComma && !hasDot) {
      normalized = s1.replace(/\./g, "").replace(",", ".");
    } else {
      normalized = s1.replace(/,/g, "");
    }

    const n = Number(normalized);
    return Number.isFinite(n) ? n : null;
  }

  function fmtNum(n, d = 2) {
    if (!Number.isFinite(n)) return "";
    return n.toFixed(d).replace(".", ",");
  }

  function fmtMoney(n) {
    if (!Number.isFinite(n)) return "";
    // como é folha (BR) e arquivo pode ter misturas, evita currency por while; mantém #,##0,00
    try {
      return new Intl.NumberFormat("pt-BR", { minimumFractionDigits: 2, maximumFractionDigits: 2 }).format(n);
    } catch {
      return String(n);
    }
  }

  function mean(arr) {
    const xs = arr.filter(Number.isFinite);
    if (!xs.length) return null;
    return xs.reduce((a,b)=>a+b,0) / xs.length;
  }

  function stdevSample(arr) {
    const xs = arr.filter(Number.isFinite);
    const n = xs.length;
    if (n < 2) return null;
    const m = mean(xs);
    let ss = 0;
    for (const x of xs) ss += (x - m) * (x - m);
    return Math.sqrt(ss / (n - 1));
  }

  function badgeNode(status) {
    const s = safeText(status);
    const span = document.createElement("span");
    span.className = "badge";
    span.textContent = s || "";
    if (s === "Aceitável") span.classList.add("ok");
    else if (s === "Alerta") span.classList.add("warn");
    else if (s === "Fora") span.classList.add("danger");
    else span.classList.add("na");
    return span;
  }

  function monthPill(val, cls) {
    const span = document.createElement("span");
    span.className = "monthpill " + cls;
    span.textContent = fmtMoney(val);
    return span;
  }

  // Band visual: eixo baseado em [mu-3s, mu+3s] com margens; desenha 3 ranges (1σ,2σ,3σ), linha no mu e marcador no ref.
  function buildSigmaViz(mu, sigma, refVal) {
    const w = 260;
    const h = 28;
    const pad = 6;

    if (!Number.isFinite(mu) || !Number.isFinite(sigma)) return "";
    const s = sigma;
    const left = mu - 3*s;
    const right = mu + 3*s;

    // margem visual quando sigma muito pequeno
    const span = Math.max(1, right - left);
    const L = left - span * 0.10;
    const R = right + span * 0.10;

    const x = (v) => {
      const t = (v - L) / (R - L);
      return pad + t * (w - pad*2);
    };

    const yMid = 14;
    const rect = (x1, x2, cls) => {
      const a = Math.min(x1, x2);
      const b = Math.max(x1, x2);
      return `<rect class="${cls}" x="${a}" y="${yMid-7}" width="${Math.max(1,b-a)}" height="14" rx="7" ry="7"></rect>`;
    };

    const axis = `<line class="axis" x1="${pad}" y1="${yMid}" x2="${w-pad}" y2="${yMid}"></line>`;
    const line = (xx, cls) => `<line class="${cls}" x1="${xx}" y1="${yMid-10}" x2="${xx}" y2="${yMid+10}"></line>`;

    const x1a = x(mu - 1*s), x1b = x(mu + 1*s);
    const x2a = x(mu - 2*s), x2b = x(mu + 2*s);
    const x3a = x(mu - 3*s), x3b = x(mu + 3*s);

    const xm = x(mu);
    const xr = Number.isFinite(refVal) ? x(refVal) : null;

    let out = `<div class="bandviz"><svg viewBox="0 0 ${w} ${h}" role="img">`;
    out += axis;
    out += rect(x3a, x3b, "range3");
    out += rect(x2a, x2b, "range2");
    out += rect(x1a, x1b, "range1");
    out += line(xm, "mid");
    if (xr !== null) {
      out += `<line class="ptline" x1="${xr}" y1="${yMid-10}" x2="${xr}" y2="${yMid+10}"></line>`;
      out += `<circle class="pt" cx="${xr}" cy="${yMid}" r="3.2"></circle>`;
    }
    out += `</svg></div>`;
    return out;
  }

  function guessColumn(headers, patterns) {
    const h = headers.map(safeText);
    for (const pat of patterns) {
      const re = pat instanceof RegExp ? pat : new RegExp(pat, "i");
      const found = h.find(col => re.test(col));
      if (found) return found;
    }
    return null;
  }

  function sheetToJson(wb, sheetName, opts = {}) {
    const ws = wb.Sheets[sheetName];
    if (!ws) return [];
    return XLSX.utils.sheet_to_json(ws, { defval: null, raw: true, ...opts });
  }

  function detectMainSheet(wb) {
    const names = wb.SheetNames || [];
    // tenta 1ª com cabeçalho padrão
    for (const name of names) {
      const rows = sheetToJson(wb, name, { range: 0 });
      const headers = Object.keys(rows[0] || {}).map(safeText).join(" | ").toLowerCase();
      if (headers.includes("código") || headers.includes("codigo")) return name;
    }
    return names[0] || null;
  }

  // Detecta colunas de mês no padrão "MMM/AA - <Métrica>" (PT-BR).
  const MONTHS_PT = {
    "JAN":1,"FEV":2,"MAR":3,"ABR":4,"MAI":5,"JUN":6,
    "JUL":7,"AGO":8,"SET":9,"OUT":10,"NOV":11,"DEZ":12
  };

  function normalizeMetricName(raw) {
    const s = safeText(raw).toUpperCase().replace(/\s+/g, " ").trim();
    if (!s) return "";
    if (s.includes("VALOR")) return "Valor";
    if (s.includes("HORA")) return "Hora";
    if (s.includes("DT PGTO") || s.includes("DT. PGTO") || s.includes("DATA PGTO") || s.includes("PAGTO")) return "Dt Pgto";
    const t = s.toLowerCase();
    return t.charAt(0).toUpperCase() + t.slice(1);
  }

  function parseMonthMetricHeader(h) {
    const s0 = safeText(h);
    if (!s0) return null;
    const s = s0.toUpperCase().replace(/\s+/g, " ").trim();

    // exemplos: "JUN/25 - HORA", "JUN/25 - VALOR", "JUN/25 - DT PGTO"
    const m1 = s.match(/\b(JAN|FEV|MAR|ABR|MAI|JUN|JUL|AGO|SET|OUT|NOV|DEZ)\s*\/\s*(\d{2})\s*-\s*(.+)$/i);
    if (m1) {
      const mon = MONTHS_PT[m1[1]];
      const yy = Number(m1[2]);
      const year = 2000 + yy;
      const key = String(year) + "-" + String(mon).padStart(2, "0");
      const label = m1[1] + "/" + m1[2];
      const metric = normalizeMetricName(m1[3]);
      return { key, label, month: mon, year, metric, header: s0 };
    }

    // fallback: "2026-02 - Valor"
    const m2 = s.match(/\b(20\d{2})\s*[-\/]\s*(\d{2})\s*-\s*(.+)$/i);
    if (m2) {
      const year = Number(m2[1]);
      const mon = Number(m2[2]);
      if (!(mon >= 1 && mon <= 12)) return null;
      const key = String(year) + "-" + String(mon).padStart(2, "0");
      const abbr = Object.keys(MONTHS_PT).find(k => MONTHS_PT[k] === mon) || String(mon).padStart(2, "0");
      const label = abbr + "/" + String(year).slice(2);
      const metric = normalizeMetricName(m2[3]);
      return { key, label, month: mon, year, metric, header: s0 };
    }

    return null;
  }

  function detectMonthBlocks(headers) {
    const map = new Map(); // key -> {key,label,month,year,metrics:{}}
    const metricSet = new Set();
    const monthHeaderSet = new Set();

    for (const h of headers) {
      const info = parseMonthMetricHeader(h);
      if (!info || !info.metric) continue;
      monthHeaderSet.add(info.header);
      metricSet.add(info.metric);
      if (!map.has(info.key)) {
        map.set(info.key, { key: info.key, label: info.label, month: info.month, year: info.year, metrics: {} });
      }
      map.get(info.key).metrics[info.metric] = info.header;
    }

    const months = Array.from(map.values()).sort((a,b) => a.key.localeCompare(b.key));
    const metrics = Array.from(metricSet.values()).sort((a,b)=>a.localeCompare(b, "pt-BR"));
    return { months, metrics, monthHeaderSet };
  }

  function detectNumericMetrics(rawRows, baseMonths, metrics) {
    // escolhe métricas que parecem numéricas (ex.: Valor, Hora)
    const out = [];
    const sampleRows = rawRows.slice(0, 80);
    const sampleMonths = baseMonths.slice(0, Math.min(baseMonths.length, 8));

    for (const m of metrics) {
      if (/\b(dt|data)\b/i.test(m) || /pgto/i.test(m)) continue; // evita datas
      let ok = 0;
      let tot = 0;
      for (const r of sampleRows) {
        for (const mon of sampleMonths) {
          const h = mon.metrics ? mon.metrics[m] : null;
          if (!h) continue;
          tot += 1;
          const v = parseNumber(r[h]);
          if (Number.isFinite(v)) ok += 1;
        }
      }
      if (tot === 0) continue;
      const ratio = ok / tot;
      if (ratio >= 0.50) out.push(m);
    }

    // prioridade: Valor
    if (out.includes("Valor")) {
      return ["Valor"].concat(out.filter(x => x !== "Valor"));
    }
    return out;
  }

  function deriveMonthsForMetric(baseMonths, metric) {
    return (baseMonths || []).map(m => ({
      key: m.key,
      label: m.label,
      month: m.month,
      year: m.year,
      header: (m.metrics || {})[metric] || null
    }));
  }

  function setSelectOptions(selectEl, values, allLabel = "Todos", opts = {}) {
    const keepOrder = !!opts.keepOrder;
    const uniq = [];
    const seen = new Set();
    for (const v0 of values.map(safeText).filter(Boolean)) {
      if (seen.has(v0)) continue;
      seen.add(v0);
      uniq.push(v0);
    }
    if (!keepOrder) uniq.sort((a,b)=>a.localeCompare(b,"pt-BR"));
    selectEl.innerHTML = "";
    const optAll = document.createElement("option");
    optAll.value = "";
    optAll.textContent = allLabel;
    selectEl.appendChild(optAll);
    for (const v of uniq) {
      const o = document.createElement("option");
      o.value = v;
      o.textContent = v;
      selectEl.appendChild(o);
    }
  }

  function saveVisibleColumns() {
    localStorage.setItem(STORAGE_COLS, JSON.stringify(state.visibleExtraColumns || []));
  }
  function loadVisibleColumns() {
    const raw = localStorage.getItem(STORAGE_COLS);
    if (!raw) return null;
    try {
      const v = JSON.parse(raw);
      return Array.isArray(v) ? v : null;
    } catch { return null; }
  }
  function saveGroupBy() {
    localStorage.setItem(STORAGE_GROUP, JSON.stringify(state.groupByColumns || []));
  }
  function loadGroupBy() {
    const raw = localStorage.getItem(STORAGE_GROUP);
    if (!raw) return null;
    try {
      const v = JSON.parse(raw);
      return Array.isArray(v) ? v : null;
    } catch { return null; }
  }

  function guessDefaultGroupBy(cols) {
    const fixed = pickColsByKeys(cols || [], FIXED_GROUP_KEYS_DEFAULT);
    if (fixed.length) return fixed;

    const wants = [/estabelecimento/i, /centro\s*de\s*custo/i, /c\.?r\.?/i, /matr/i, /empresa/i];
    const pick = [];
    for (const c of cols) {
      if (wants.some(re => re.test(c))) pick.push(c);
    }
    // default: tenta 2-3 colunas
    return pick.slice(0, 3);
  }

  function guessDefaultVisibleColumns(cols) {
    const fixed = pickColsByKeys(cols || [], FIXED_VISIBLE_KEYS_DEFAULT);
    if (fixed.length) return fixed;

    const wants = [/empresa/i, /estabelecimento/i, /centro\s*de\s*custo/i, /c\.?r\.?/i, /processo/i, /clas/i, /matr/i, /nome/i];
    const pick = [];
    for (const c of cols) {
      if (wants.some(re => re.test(c))) pick.push(c);
    }
    return pick.slice(0, 6);
  }

  function buildColPicker() {
    if (!ui.colPickerBody) return;
    ui.colPickerBody.innerHTML = "";

    const cols = state.extraColumns || [];
    const visible = new Set(state.visibleExtraColumns || []);

    for (const c of cols) {
      const label = document.createElement("label");
      label.className = "chk";

      const inp = document.createElement("input");
      inp.type = "checkbox";
      inp.checked = visible.has(c);

      inp.addEventListener("change", () => {
        const set = new Set(state.visibleExtraColumns || []);
        if (inp.checked) set.add(c);
        else set.delete(c);
        state.visibleExtraColumns = Array.from(set).sort((a,b)=>a.localeCompare(b,"pt-BR"));
        saveVisibleColumns();
        buildExtraFilters(state.groups);
        applyFilters();
      });

      const span = document.createElement("span");
      span.textContent = c;

      label.appendChild(inp);
      label.appendChild(span);

      ui.colPickerBody.appendChild(label);
    }
  }

  function buildGroupPicker() {
    if (!ui.groupPickerBody) return;
    ui.groupPickerBody.innerHTML = "";

    const cols = state.baseColumns || [];
    const selected = new Set(state.groupByColumns || []);

    for (const c of cols) {
      const label = document.createElement("label");
      label.className = "chk";

      const inp = document.createElement("input");
      inp.type = "checkbox";
      inp.checked = selected.has(c);

      inp.addEventListener("change", () => {
        const set = new Set(state.groupByColumns || []);
        if (inp.checked) set.add(c);
        else set.delete(c);
        state.groupByColumns = Array.from(set);
        saveGroupBy();
        rebuildAggregation(); // recalcula grupos
      });

      const span = document.createElement("span");
      span.textContent = c;

      label.appendChild(inp);
      label.appendChild(span);

      ui.groupPickerBody.appendChild(label);
    }
  }

  function wireClearButton(inputEl, btnEl) {
    if (!inputEl || !btnEl) return;

    const sync = () => {
      const has = safeText(inputEl.value).length > 0;
      btnEl.style.display = has ? "inline-flex" : "none";
    };

    btnEl.addEventListener("click", () => {
      inputEl.value = "";
      sync();
      applyFilters();
      inputEl.focus();
    });

    inputEl.addEventListener("input", sync);
    sync();
  }

  function buildExtraFilters(rows) {
    if (!ui.extraFilters) return;
    ui.extraFilters.innerHTML = "";

    const cols = state.extraColumns || [];
    const selected = new Set(state.visibleExtraColumns || []);

    // conforme layout solicitado: por padrão, "Filtros extras" mostra apenas C.R. e Clas. (quando existirem).
    // se essas colunas não existirem no arquivo, cai no comportamento original (colunas marcadas).
    const fixedCols = pickColsByKeys(cols, FIXED_FILTER_KEYS);
    const ordered = fixedCols.length ? fixedCols : Array.from(selected).filter(c => cols.includes(c));

    // remove filtros de colunas que sumiram
    for (const k of Object.keys(state.extraFilterValues || {})) {
      if (!cols.includes(k)) {
        delete state.extraFilterValues[k];
        delete state.extraFilterModes[k];
      }
    }

    const makeField = (labelText) => {
      const wrap = document.createElement("div");
      wrap.className = "field";
      const lab = document.createElement("label");
      lab.textContent = labelText;
      wrap.appendChild(lab);
      return wrap;
    };

    const valuesFor = (col) => rows
      .map(r => (r && r.extras) ? r.extras[col] : (r ? r[col] : null))
      .map(v => safeText(v))
      .filter(Boolean);

    for (const col of ordered) {
      const vals = valuesFor(col);
      const uniq = Array.from(new Set(vals));
      const maxLen = uniq.reduce((m, x) => Math.max(m, x.length), 0);

      const wrap = makeField(col);
      const isSmallCategorical = (uniq.length >= 2 && uniq.length <= 30 && maxLen <= 60);

      if (isSmallCategorical) {
        const sel = document.createElement("select");
        sel.setAttribute("data-col", col);
        sel.setAttribute("data-mode", "select");

        const optAll = document.createElement("option");
        optAll.value = "";
        optAll.textContent = "Todos";
        sel.appendChild(optAll);

        uniq.sort((a,b)=>a.localeCompare(b,"pt-BR"));
        for (const v of uniq) {
          const o = document.createElement("option");
          o.value = v;
          o.textContent = v;
          sel.appendChild(o);
        }

        const saved = safeText(state.extraFilterValues[col] || "");
        sel.value = saved;
        state.extraFilterModes[col] = "select";

        sel.addEventListener("change", () => {
          state.extraFilterValues[col] = sel.value;
          state.extraFilterModes[col] = "select";
          applyFilters();
        });

        const w2 = document.createElement("div");
        w2.className = "inputwrap";
        const b2 = document.createElement("button");
        b2.className = "clearbtn";
        b2.type = "button";
        b2.textContent = "X";
        w2.appendChild(sel);
        w2.appendChild(b2);

        const sync2 = () => { b2.style.display = sel.value ? "inline-flex" : "none"; };
        b2.addEventListener("click", () => { sel.value = ""; sync2(); state.extraFilterValues[col] = ""; applyFilters(); sel.focus(); });
        sel.addEventListener("change", sync2);
        sync2();

        wrap.appendChild(w2);
      } else {
        const inp = document.createElement("input");
        inp.type = "text";
        inp.placeholder = "Digite para filtrar";
        inp.setAttribute("data-col", col);
        inp.setAttribute("data-mode", "contains");

        const saved = safeText(state.extraFilterValues[col] || "");
        inp.value = saved;
        state.extraFilterModes[col] = "contains";

        inp.addEventListener("input", () => {
          state.extraFilterValues[col] = inp.value;
          state.extraFilterModes[col] = "contains";
          applyFilters();
        });

        const w3 = document.createElement("div");
        w3.className = "inputwrap";
        const b3 = document.createElement("button");
        b3.className = "clearbtn";
        b3.type = "button";
        b3.textContent = "X";
        w3.appendChild(inp);
        w3.appendChild(b3);
        wrap.appendChild(w3);
        wireClearButton(inp, b3);

        // datalist leve
        if (uniq.length > 0 && uniq.length <= 200) {
          const dlId = "dl_" + col.replace(/[^a-z0-9]/gi, "_").toLowerCase();
          inp.setAttribute("list", dlId);
          const dl = document.createElement("datalist");
          dl.id = dlId;
          uniq.slice(0, 200).sort((a,b)=>a.localeCompare(b,"pt-BR")).forEach(v => {
            const o = document.createElement("option");
            o.value = v;
            dl.appendChild(o);
          });
          wrap.appendChild(dl);
        }
      }

      ui.extraFilters.appendChild(wrap);
    }
  }

  function buildVerbaSelect(rawRows) {
    const colCode = state.colCode;
    const colDesc = state.colDesc;
    const vals = rawRows.map(r => {
      const c = safeText(r[colCode]);
      const d = safeText(r[colDesc]);
      if (!c && !d) return "";
      return (c ? c : "") + " - " + (d ? d : "");
    }).filter(Boolean);

    setSelectOptions(ui.fVerba, vals, "Todas");
  }

  function buildMonthSelect() {
    const labels = state.months.map(m => m.label);
    setSelectOptions(ui.fRefMonth, labels, "Último mês", { keepOrder: true });
  }

  function buildMetricSelect() {
    const vals = (state.metricOptions || []).slice();
    ui.fMetric.innerHTML = "";
    for (const v of vals) {
      const o = document.createElement("option");
      o.value = v;
      o.textContent = v;
      ui.fMetric.appendChild(o);
    }

    if (vals.length) {
      const saved = safeText(localStorage.getItem(STORAGE_METRIC) || "");
      const chosen = saved && vals.includes(saved) ? saved : (vals.includes("Valor") ? "Valor" : vals[0]);
      ui.fMetric.value = chosen;
      state.metric = chosen;
    }
  }

  function computeStatusFromZ(z) {
    if (!Number.isFinite(z)) return "Sem histórico";
    const a = Math.abs(z);
    if (a <= 2) return "Aceitável";
    if (a <= 3) return "Alerta";
    return "Fora";
  }

  function computeRowMonthClass(v, mu, sigma) {
    if (!Number.isFinite(v) || !Number.isFinite(mu) || !Number.isFinite(sigma) || sigma === 0) return "na";
    const z = (v - mu) / sigma;
    const a = Math.abs(z);
    if (a <= 2) return "ok";
    if (a <= 3) return "warn";
    return "danger";
  }

  function getRefMonthLabel() {
    // "Último mês" => value vazio => pega último da lista
    const sel = safeText(ui.fRefMonth.value);
    if (sel) return sel;
    const last = state.months[state.months.length - 1];
    return last ? last.label : "";
  }

  function getWindowN() {
    const n = parseNumber(ui.fWindow.value);
    if (Number.isFinite(n) && n >= 2) return Math.floor(n);
    return null; // null => usa todos
  }

  function rebuildAggregation() {
    // reconstroi os grupos agregados a partir dos rawRows e meses
    const rawRows = state.rawRows || [];
    const months = state.months || [];
    if (!rawRows.length || !months.length) {
      state.groups = [];
      state.filtered = [];
      applyFilters();
      return;
    }

    const colCode = state.colCode;
    const colDesc = state.colDesc;

    const groupBy = (state.groupByColumns && state.groupByColumns.length)
      ? state.groupByColumns.slice()
      : guessDefaultGroupBy(state.baseColumns);

    state.groupByColumns = groupBy;

    // Mapa: key => {keyParts, code, desc, extras, valuesByMonth}
    const map = new Map();

    for (const r of rawRows) {
      const code = safeText(r[colCode]);
      const desc = safeText(r[colDesc]);
      const verbaKey = (code ? code : "") + " - " + (desc ? desc : "");

      // key do grupo = verba + colunas selecionadas
      const parts = [];
      for (const c of groupBy) parts.push(safeText(r[c]));
      const key = verbaKey.toUpperCase() + "|" + parts.map(p=>p.toUpperCase()).join("|");

      if (!map.has(key)) {
        const extras = {};
        // extras: carrega apenas colunas extras marcadas, mas mantém todas p/ filtro e export
        for (const c of state.extraColumns) extras[c] = safeText(r[c]);
        map.set(key, {
          verbaKey,
          code,
          desc,
          groupParts: parts,
          groupBy,
          extras,
          values: {}, // label -> sum
        });
      }
      const obj = map.get(key);

      for (const m of months) {
        const v = parseNumber(r[m.header]);
        if (!Number.isFinite(v)) continue;
        obj.values[m.label] = (obj.values[m.label] || 0) + v;
      }
    }

    const groups = [];
    for (const obj of map.values()) {
      groups.push(obj);
    }

    // aplica estatística por grupo (depende de ref month / janela / ignore zeros); calculamos on-the-fly em applyFilters para refletir parâmetros.
    state.groups = groups;

    // UI (group picker) precisa refletir seleção atual
    buildGroupPicker();
    buildExtraFilters(state.rawRows);
    applyFilters();
  }

  function getSortValue(row, key) {
    if (!key) return null;
    if (key.startsWith("m:")) {
      const lab = key.slice(2);
      const v = row.values ? row.values[lab] : null;
      return Number.isFinite(v) ? v : null;
    }
    if (key.startsWith("extra:")) {
      const c = key.slice(6);
      return safeText((row.extras || {})[c]).toLowerCase();
    }
    const direct = row[key];
    if (typeof direct === "number") return Number.isFinite(direct) ? direct : null;
    return safeText(direct).toLowerCase();
  }

  function applySort(rows) {
    const key = state.sort && state.sort.key ? state.sort.key : null;
    const dir = state.sort && state.sort.dir ? state.sort.dir : 1;
    if (!key) return rows;

    const arr = rows.slice();
    arr.sort((a,b) => {
      const va = getSortValue(a, key);
      const vb = getSortValue(b, key);

      if (va === null && vb === null) return 0;
      if (va === null) return 1;
      if (vb === null) return -1;

      const na = typeof va === "number";
      const nb = typeof vb === "number";
      if (na && nb) return (va - vb) * dir;
      if (!na && !nb) return va.localeCompare(vb, "pt-BR") * dir;
      return String(va).localeCompare(String(vb), "pt-BR") * dir;
    });
    return arr;
  }

  function setSort(key) {
    if (!key) return;
    if (!state.sort) state.sort = { key: null, dir: 1 };
    if (state.sort.key === key) state.sort.dir = state.sort.dir * -1;
    else {
      state.sort.key = key;
      state.sort.dir = 1;
    }
  }

  function computeKPIs(rows) {
    const total = rows.length;
    const ok = rows.filter(r => r.status === "Aceitável").length;
    const warn = rows.filter(r => r.status === "Alerta").length;
    const out = rows.filter(r => r.status === "Fora").length;
    const zs = rows.map(r => r.z).filter(Number.isFinite);
    const avgZ = zs.length ? zs.reduce((a,b)=>a+b,0)/zs.length : null;

    ui.kpiTotal.textContent = String(total);
    ui.kpiOk.textContent = String(ok);
    ui.kpiWarn.textContent = String(warn);
    ui.kpiOut.textContent = String(out);
    ui.kpiAvgZ.textContent = Number.isFinite(avgZ) ? fmtNum(avgZ, 2) : "0,00";
  }

  function buildTableHeader() {
    if (!ui.tblHead) return;
    ui.tblHead.innerHTML = "";

    const metricHint = state.metric ? ` (métrica: ${state.metric})` : "";
    const helpFor = (label) => {
      const L = safeText(label);
      if (!L) return "";
      if (L === "Verba") return "Código - Descrição da verba.";
      if (L === "Grupo") return "Combinação das colunas marcadas em 'Agrupar por'.";
      if (L === "Ref") return `Valor do mês de referência${metricHint}.`;
      if (L === "Média") return "Média do histórico (μ).";
      if (L === "σ") return "Desvio padrão amostral do histórico (σ).";
      if (L.startsWith("LCL")) return "LCL (limite inferior) = μ − 3σ.";
      if (L.startsWith("UCL")) return "UCL (limite superior) = μ + 3σ.";
      if (L === "Z") return "Z-score = (Ref − μ) / σ.";
      if (L === "Status") return "Aceitável: |Z| ≤ 2 | Alerta: 2 < |Z| ≤ 3 | Fora: |Z| > 3 | Sem histórico: σ=0 ou histórico insuficiente.";
      if (L === "Faixa") return "Faixa 6σ: μ ± 3σ.";
      if (/^(JAN|FEV|MAR|ABR|MAI|JUN|JUL|AGO|SET|OUT|NOV|DEZ)\/\d{2}$/i.test(L)) return `Valor agregado do mês${metricHint}.`;
      return "";
    };

    const tr = document.createElement("tr");
    const th = (label, key, cls) => {
      const x = document.createElement("th");
      x.textContent = label;
      const tip = helpFor(label);
      if (tip) x.title = tip;
      if (cls) x.className = cls;
      if (key) {
        x.classList.add("sortable");
        x.setAttribute("data-key", key);
        x.addEventListener("click", () => {
          setSort(key);
          applyFilters();
        });
      }
      return x;
    };

    // grupo
    tr.appendChild(th("Verba","verbaKey"));
    tr.appendChild(th("Grupo","groupLabel"));

    // extras visíveis
    for (const c of (state.visibleExtraColumns || [])) {
      tr.appendChild(th(c, "extra:" + c));
    }

    tr.appendChild(th("Ref","refVal","num"));
    tr.appendChild(th("Média","mu","num"));
    tr.appendChild(th("σ","sigma","num"));
    tr.appendChild(th("LCL (-3σ)","lcl","num"));
    tr.appendChild(th("UCL (+3σ)","ucl","num"));
    tr.appendChild(th("Z","z","num"));
    tr.appendChild(th("Status","status"));
    tr.appendChild(th("Faixa", null));

    // todos os meses como colunas (grid completo)
    for (const m of state.months) {
      tr.appendChild(th(m.label, "m:" + m.label, "num"));
    }

    ui.tblHead.appendChild(tr);

    const ths = ui.tblHead.querySelectorAll("th.sortable");
    ths.forEach(h => {
      h.classList.remove("sort-asc");
      h.classList.remove("sort-desc");
      const k = h.getAttribute("data-key");
      if (state.sort && state.sort.key && k === state.sort.key) {
        if (state.sort.dir === 1) h.classList.add("sort-asc");
        else h.classList.add("sort-desc");
      }
    });
  }

  function renderTable(rows, refLabel) {
    buildTableHeader();
    ui.tblBody.innerHTML = "";
    const frag = document.createDocumentFragment();

    for (const r of rows) {
      const tr = document.createElement("tr");

      const td = (text, cls) => {
        const x = document.createElement("td");
        if (cls) x.className = cls;
        x.textContent = text;
        return x;
      };

      tr.appendChild(td(r.verbaKey || ""));
      tr.appendChild(td(r.groupLabel || ""));

      for (const c of (state.visibleExtraColumns || [])) {
        tr.appendChild(td(safeText((r.extras || {})[c] || "")));
      }

      tr.appendChild(td(fmtMoney(r.refVal), "num"));
      tr.appendChild(td(fmtMoney(r.mu), "num"));
      tr.appendChild(td(fmtMoney(r.sigma), "num"));
      tr.appendChild(td(fmtMoney(r.lcl), "num"));
      tr.appendChild(td(fmtMoney(r.ucl), "num"));
      tr.appendChild(td(Number.isFinite(r.z) ? fmtNum(r.z, 2) : "", "num"));

      const tdStatus = document.createElement("td");
      tdStatus.appendChild(badgeNode(r.status));
      tr.appendChild(tdStatus);

      const tdViz = document.createElement("td");
      tdViz.innerHTML = buildSigmaViz(r.mu, r.sigma, r.refVal);
      tr.appendChild(tdViz);

      // meses
      for (const m of state.months) {
        const v = Number.isFinite(r.values[m.label]) ? r.values[m.label] : null;
        const cls = (m.label === refLabel) ? "monthcell" : "monthcell";
        const cell = document.createElement("td");
        cell.className = cls + " num";

        if (!Number.isFinite(v)) {
          cell.appendChild(monthPill(0, "na"));
        } else {
          // classifica mês com relação à média/sigma do grupo (mesmo sigma)
          const pillCls = computeRowMonthClass(v, r.mu, r.sigma);
          const pill = monthPill(v, pillCls);
          // destaca ref
          if (m.label === refLabel) {
            pill.style.boxShadow = "0 0 0 3px rgba(10, 99, 198, 0.18)";
          }
          cell.appendChild(pill);
        }
        tr.appendChild(cell);
      }

      frag.appendChild(tr);
    }

    ui.tblBody.appendChild(frag);
  }

  

  function renderMonthTotals(rows, refLabel) {
    if (!ui.monthTotals) return;
    const months = state.months || [];
    if (!months.length) { ui.monthTotals.innerHTML = ""; return; }

    const sumsAll = {};
    const sumsFora = {};
    const sumsAlerta = {};

    for (const m of months) {
      sumsAll[m.label] = 0;
      sumsFora[m.label] = 0;
      sumsAlerta[m.label] = 0;
    }

    for (const r of (rows || [])) {
      for (const m of months) {
        const v = r && r.values ? r.values[m.label] : null;
        const num = Number.isFinite(v) ? v : 0;
        sumsAll[m.label] += num;
        if (r.status === "Fora") sumsFora[m.label] += num;
        if (r.status === "Alerta") sumsAlerta[m.label] += num;
      }
    }

    const arrFrom = (obj) => months.map(m => obj[m.label] || 0);

    const classByZ = (arr, val) => {
      if (!Number.isFinite(val) || val === 0) return "na";
      const xs = arr.filter(x => Number.isFinite(x) && x !== 0);
      if (xs.length < 2) return "ok";
      const mu = mean(xs);
      const sd = stdevSample(xs);
      if (!Number.isFinite(mu) || !Number.isFinite(sd) || sd === 0) return "ok";
      const z = (val - mu) / sd;
      const a = Math.abs(z);
      // Mantém apenas verde/laranja como no layout de referência (sem vermelho nessa faixa de totais)
      if (a > 2) return "warn";
      return "ok";
    };

    const allArr = arrFrom(sumsAll);
    const alertaArr = arrFrom(sumsAlerta);

    const grid = document.createElement("div");
    grid.className = "mt-grid";
    grid.style.gridTemplateColumns = `repeat(${months.length}, minmax(88px, 1fr))`;

    // Cabeçalho de meses
    for (const m of months) {
      const d = document.createElement("div");
      d.className = "mt-label";
      d.textContent = m.label;
      grid.appendChild(d);
    }

    const addRow = (obj, mode) => {
      for (const m of months) {
        const v = obj[m.label] || 0;
        const cell = document.createElement("div");
        cell.className = "mt-cell";

        let cls = "na";
        if (mode === "fora") {
          cls = (v === 0) ? "na" : "sum";
        } else if (mode === "all") {
          cls = classByZ(allArr, v);
        } else if (mode === "alerta") {
          cls = classByZ(alertaArr, v);
        }

        const pill = monthPill(v, cls);
        if (m.label === refLabel) pill.classList.add("ref");

        cell.appendChild(pill);
        grid.appendChild(cell);
      }
    };

    // Ordem conforme figura: fora (neutro), total (verde/laranja), alerta (verde/laranja)
    addRow(sumsFora, "fora");
    addRow(sumsAll, "all");
    addRow(sumsAlerta, "alerta");

    const scroll = document.createElement("div");
    scroll.className = "mt-scroll";
    scroll.appendChild(grid);

    ui.monthTotals.innerHTML = "";
    ui.monthTotals.appendChild(scroll);
  }


  function renderDiag(rows) {
    const total = rows.length;
    const sem = rows.filter(r => r.status === "Sem histórico").length;
    const fora = rows.filter(r => r.status === "Fora").length;

    const byVerba = new Map();
    for (const r of rows) {
      const k = safeText(r.verbaKey || "Sem verba");
      if (!byVerba.has(k)) byVerba.set(k, { k, total: 0, fora: 0 });
      const o = byVerba.get(k);
      o.total++;
      if (r.status === "Fora") o.fora++;
    }
    const top = Array.from(byVerba.values()).sort((a,b)=>b.fora-a.fora).slice(0,5);

    const box = (title, lines) => {
      const div = document.createElement("div");
      div.className = "box";
      const h = document.createElement("h4");
      h.textContent = title;
      const p = document.createElement("p");
      p.innerHTML = lines.join("<br>");
      div.appendChild(h);
      div.appendChild(p);
      return div;
    };

    ui.diagBox.innerHTML = "";
    ui.diagBox.appendChild(box("Leitura do arquivo", [
      `Grupos (após agregação): ${total}`,
      `Meses detectados: ${(state.months || []).length}`,
      `Agrupar por: ${(state.groupByColumns || []).join(" | ") || "não definido"}`,
      `Colunas extras detectadas: ${(state.extraColumns || []).length}`,
      `Colunas extras visíveis: ${(state.visibleExtraColumns || []).length}`,
      `Sem histórico: ${sem}`,
      `Fora: ${fora}`
    ]));

    ui.diagBox.appendChild(box("Top verbas com Fora", top.length
      ? top.map(x => `${x.k}: fora ${x.fora} | total ${x.total}`)
      : ["Sem dados suficientes."]));

    ui.diagBox.appendChild(box("Regra aplicada", [
      `Faixa (6σ): média ± 3σ (LCL/UCL).`,
      `Status: Aceitável (|Z| ≤ 2), Alerta (2 < |Z| ≤ 3), Fora (|Z| > 3).`,
      `Z = (Ref - média) / σ.`
    ]));
  }

  function applyFilters() {
    const q = safeText(ui.fSearch.value).toLowerCase();
    const verba = safeText(ui.fVerba.value);
    const status = safeText(ui.fStatus.value);
    const minZ = parseNumber(ui.fMinZ.value);
    const maxZ = parseNumber(ui.fMaxZ.value);
    const ignoreZeros = safeText(ui.fIgnoreZeros.value) === "1";
    const windowN = getWindowN();
    const refLabel = getRefMonthLabel();

    // calcula estatística e status em tempo real com base nos parâmetros
    let rows = state.groups.map(g => {
      const valsAll = state.months.map(m => Number.isFinite(g.values[m.label]) ? g.values[m.label] : 0);
      const idxRef = state.months.findIndex(m => m.label === refLabel);
      const refIdx = (idxRef >= 0) ? idxRef : (state.months.length - 1);

      const history = [];
      // meses antes do ref
      for (let i = 0; i < refIdx; i++) {
        const v = valsAll[i];
        if (!Number.isFinite(v)) continue;
        if (ignoreZeros && v === 0) continue;
        history.push(v);
      }
      // aplica janela (pega os últimos N do histórico)
      const hist = (windowN && history.length > windowN) ? history.slice(history.length - windowN) : history.slice();

      const mu = mean(hist);
      const sigma = stdevSample(hist);
      const refVal = Number.isFinite(valsAll[refIdx]) ? valsAll[refIdx] : 0;

      let z = null;
      if (Number.isFinite(mu) && Number.isFinite(sigma) && sigma !== 0) z = (refVal - mu) / sigma;

      const lcl = (Number.isFinite(mu) && Number.isFinite(sigma)) ? (mu - 3*sigma) : null;
      const ucl = (Number.isFinite(mu) && Number.isFinite(sigma)) ? (mu + 3*sigma) : null;

      const st = computeStatusFromZ(z);

      const groupLabel = (g.groupParts || []).filter(Boolean).join(" | ") || "(Sem grupo)";

      return {
        verbaKey: g.verbaKey,
        code: g.code,
        desc: g.desc,
        groupLabel,
        groupBy: g.groupBy,
        extras: g.extras,
        values: g.values,
        refVal,
        mu,
        sigma,
        lcl,
        ucl,
        z,
        status: st,
      };
    });

    // filtros
    if (verba) rows = rows.filter(r => r.verbaKey === verba);

    if (q) {
      rows = rows.filter(r => {
        const extraHay = (state.extraColumns || []).map(c => safeText((r.extras || {})[c])).join(" ");
        const hay = [
          r.verbaKey, r.groupLabel, r.code, r.desc, r.status, extraHay
        ].map(safeText).join(" ").toLowerCase();
        return hay.includes(q);
      });
    }

    if (status) rows = rows.filter(r => r.status === status);

    if (Number.isFinite(minZ)) rows = rows.filter(r => Number.isFinite(r.z) && r.z >= minZ);
    if (Number.isFinite(maxZ)) rows = rows.filter(r => Number.isFinite(r.z) && r.z <= maxZ);

    // filtros extras
    for (const col of Object.keys(state.extraFilterValues || {})) {
      const v = safeText(state.extraFilterValues[col]);
      if (!v) continue;

      const mode = safeText((state.extraFilterModes || {})[col]) || "select";
      if (mode === "contains") {
        const vv = v.toLowerCase();
        rows = rows.filter(r => safeText((r.extras || {})[col]).toLowerCase().includes(vv));
      } else {
        rows = rows.filter(r => safeText((r.extras || {})[col]) === v);
      }
    }

    // ordenação
    rows = applySort(rows);

    state.filtered = rows;

    computeKPIs(rows);
    renderMonthTotals(rows, refLabel);
    renderTable(rows, refLabel);
    renderDiag(rows);

    ui.tableInfo.textContent = `Exibindo ${rows.length} de ${state.groups.length}.`;
  }

  function resetFilters() {
    ui.fSearch.value = "";
    ui.fVerba.value = "";
    ui.fRefMonth.value = "";
    ui.fWindow.value = "";
    ui.fIgnoreZeros.value = "0";
    ui.fStatus.value = "";
    ui.fMinZ.value = "";
    ui.fMaxZ.value = "";

    state.extraFilterValues = {};
    if (ui.extraFilters) {
      const selects = ui.extraFilters.querySelectorAll("select[data-col]");
      selects.forEach(s => s.value = "");
    }

    applyFilters();
  }

  function fillStaticFilters() {
    setSelectOptions(ui.fStatus, ["Aceitável","Alerta","Fora","Sem histórico"], "Todos");
    buildMonthSelect();
  }

  function saveToStorage(rawRows, meta) {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(rawRows));
    localStorage.setItem(STORAGE_META, JSON.stringify(meta || {}));
  }

  function loadFromStorage() {
    const raw = localStorage.getItem(STORAGE_KEY);
    const meta = localStorage.getItem(STORAGE_META);
    if (!raw) return false;
    try {
      const rows = JSON.parse(raw);
      const m = meta ? JSON.parse(meta) : null;
      if (!Array.isArray(rows)) return false;
      state.rawRows = rows;
      state.meta = m;
      return true;
    } catch {
      return false;
    }
  }

  function clearStorage() {
    localStorage.removeItem(STORAGE_KEY);
    localStorage.removeItem(STORAGE_META);
    localStorage.removeItem(STORAGE_COLS);
    localStorage.removeItem(STORAGE_GROUP);
    localStorage.removeItem(STORAGE_METRIC);
  }

  function exportTxt(rows, meta) {
    const m = meta || {};
    const lines = [];
    lines.push("IGARAPE DIGITAL | VALIDACAO DE FOLHA POR VERBA | 6 SIGMA");
    lines.push(`Gerado em: ${new Date().toLocaleString("pt-BR")}`);
    if (m.sourceFile) lines.push(`Arquivo: ${m.sourceFile}`);
    if (m.sheet) lines.push(`Aba: ${m.sheet}`);
    if (state.metric) lines.push(`Métrica: ${state.metric}`);
    lines.push("");

    const extra = state.visibleExtraColumns || [];
    const months = state.months.map(x => x.label);

    const header = ["Verba","Grupo"].concat(extra).concat(["Ref","Media","Sigma","LCL","UCL","Z","Status","Faixa"]).concat(months);
    lines.push(header.join(" | "));
    lines.push("");

    const refLabel = getRefMonthLabel();

    for (const r of rows) {
      const extraVals = extra.map(c => safeText((r.extras || {})[c]));
      const monthVals = months.map(lab => {
        const v = r.values ? r.values[lab] : null;
        return Number.isFinite(v) ? String(v.toFixed(2)).replace(".", ",") : "0,00";
      });

      const faixa = (Number.isFinite(r.mu) && Number.isFinite(r.sigma)) ? `${fmtMoney(r.mu)} ± ${fmtMoney(3*r.sigma)}` : "";

      lines.push([
        safeText(r.verbaKey),
        safeText(r.groupLabel),
        ...extraVals,
        fmtMoney(r.refVal),
        fmtMoney(r.mu),
        fmtMoney(r.sigma),
        fmtMoney(r.lcl),
        fmtMoney(r.ucl),
        Number.isFinite(r.z) ? fmtNum(r.z, 4) : "",
        safeText(r.status),
        faixa,
        ...monthVals
      ].join(" | "));
    }

    const blob = new Blob([lines.join("\n")], { type: "text/plain;charset=utf-8" });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = `igarape_verbas_sigma_${new Date().toISOString().slice(0,10)}.txt`;
    document.body.appendChild(a);
    a.click();
    a.remove();
  }

  function exportXlsx(rows, meta) {
    if (!window.XLSX) return;

    const m = meta || {};
    const extra = (state.extraColumns || []).slice(); // exporta todas extras detectadas
    const months = state.months.map(x => x.label);

    const header = ["Verba","Grupo","Ref","Media","Sigma","LCL","UCL","Z","Status"].concat(extra).concat(months);

    const aoa = [];
    aoa.push(header);

    for (const r of rows) {
      const base = [
        safeText(r.verbaKey),
        safeText(r.groupLabel),
        Number.isFinite(r.refVal) ? r.refVal : 0,
        Number.isFinite(r.mu) ? r.mu : "",
        Number.isFinite(r.sigma) ? r.sigma : "",
        Number.isFinite(r.lcl) ? r.lcl : "",
        Number.isFinite(r.ucl) ? r.ucl : "",
        Number.isFinite(r.z) ? Number(r.z.toFixed(6)) : "",
        safeText(r.status)
      ];

      const extras = extra.map(c => safeText((r.extras || {})[c]));
      const mvals = months.map(lab => {
        const v = r.values ? r.values[lab] : null;
        return Number.isFinite(v) ? v : 0;
      });

      aoa.push(base.concat(extras).concat(mvals));
    }

    const ws = XLSX.utils.aoa_to_sheet(aoa);

    // formata numéricos
    const range = XLSX.utils.decode_range(ws["!ref"]);
    const idxRef = header.indexOf("Ref");
    const idxZ = header.indexOf("Z");
    for (let R = 1; R <= range.e.r; R++) {
      for (let C = idxRef; C < idxRef + 6; C++) { // Ref..UCL
        const addr = XLSX.utils.encode_cell({ r: R, c: C });
        if (ws[addr] && typeof ws[addr].v === "number") ws[addr].z = "#,##0.00";
      }
      const addrZ = XLSX.utils.encode_cell({ r: R, c: idxZ });
      if (ws[addrZ] && typeof ws[addrZ].v === "number") ws[addrZ].z = "0.00";
      // meses
      const startMonths = header.length - months.length;
      for (let C = startMonths; C < header.length; C++) {
        const addr = XLSX.utils.encode_cell({ r: R, c: C });
        if (ws[addr] && typeof ws[addr].v === "number") ws[addr].z = "#,##0.00";
      }
    }

    ws["!freeze"] = { xSplit: 0, ySplit: 1 };

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Export");

    const date = new Date().toISOString().slice(0,10);
    const slug = safeText(state.metric).toLowerCase().replace(/\s+/g, "_").replace(/[^a-z0-9_]/g, "");
    const fname = slug ? `igarape_verbas_sigma_${slug}_${date}.xlsx` : `igarape_verbas_sigma_export_${date}.xlsx`;
    XLSX.writeFile(wb, fname);
  }

  function sanitizeRows(rows) {
    // remove linhas vazias e normaliza chaves
    const out = [];
    for (const r of rows) {
      // se row não tem chaves (linha de título/blank), pula
      const keys = Object.keys(r || {});
      if (!keys.length) continue;
      out.push(r);
    }
    return out;
  }

  async function handleFile(file) {
    if (!window.XLSX) {
      ui.dataInfo.textContent = "Biblioteca XLSX não carregou. Verifique bloqueio de CDN.";
      return;
    }

    const buf = await file.arrayBuffer();
    const wb = XLSX.read(buf, { type: "array" });

    const sheet = detectMainSheet(wb);
    if (!sheet) throw new Error("Sem abas no arquivo.");

    // range 0: costuma pegar header errado em exports; então tentamos localizar cabeçalho real
    // abordagem: lê como JSON e filtra linhas sem colunas típicas
    let raw = sheetToJson(wb, sheet, { range: 0 });
    raw = sanitizeRows(raw);

    if (!raw.length) throw new Error("Aba sem dados.");

    const headers = Object.keys(raw[0] || {}).map(safeText);

    // colunas base
    state.colCode = guessColumn(headers, [/^c[oó]digo$/i, /^codigo$/i, /\bc[oó]d\b/i, /\bcod\b/i]);
    state.colDesc = guessColumn(headers, [/^descri[cç][aã]o$/i, /descric/i, /\bdescricao\b/i]);

    if (!state.colCode || !state.colDesc) {
      // tenta outro offset (quando a primeira linha vira título)
      raw = sheetToJson(wb, sheet, { range: 2 });
      raw = sanitizeRows(raw);
    }

    if (!raw.length) throw new Error("Aba sem dados após ajuste de cabeçalho.");

    const headers2 = Object.keys(raw[0] || {}).map(safeText);
    state.colCode = state.colCode || guessColumn(headers2, [/^c[oó]digo$/i, /^codigo$/i, /\bc[oó]d\b/i, /\bcod\b/i]);
    state.colDesc = state.colDesc || guessColumn(headers2, [/^descri[cç][aã]o$/i, /descric/i, /\bdescricao\b/i]);

    if (!state.colCode || !state.colDesc) {
      throw new Error("Não encontrei colunas Código e Descrição.");
    }

    // detecta blocos mensais (múltiplas métricas)
    const mb = detectMonthBlocks(headers2);
    state.baseMonths = mb.months;
    if (!state.baseMonths.length) {
      throw new Error("Não encontrei colunas mensais no formato 'MMM/AA - <Métrica>'.");
    }

    state.metricOptions = detectNumericMetrics(raw, state.baseMonths, mb.metrics);
    if (!state.metricOptions.length) {
      // fallback mínimo: tenta Valor se existir como nome
      state.metricOptions = mb.metrics.includes("Valor") ? ["Valor"] : mb.metrics.slice(0, 1);
    }

    // define métrica inicial e meses derivados
    const savedMetric = safeText(localStorage.getItem(STORAGE_METRIC) || "");
    const chosenMetric = (savedMetric && state.metricOptions.includes(savedMetric))
      ? savedMetric
      : (state.metricOptions.includes("Valor") ? "Valor" : state.metricOptions[0]);

    state.metric = chosenMetric;
    state.months = deriveMonthsForMetric(state.baseMonths, state.metric);

    // colunas base disponíveis para agrupamento: todas exceto colunas mensais
    const baseCols = headers2.filter(h => h && !mb.monthHeaderSet.has(h));

    // colunas extras: tudo menos código/descrição e menos as usadas para agrupamento; mas aqui tratamos como "colunas descritivas"
    const core = new Set([state.colCode, state.colDesc]);
    state.baseColumns = baseCols.filter(h => !core.has(h));

    // extras: colunas que podem ser exibidas/filtradas (todas descritivas, sem meses)
    state.extraColumns = state.baseColumns.slice().sort((a,b)=>a.localeCompare(b,"pt-BR"));

    const savedCols = loadVisibleColumns();
    state.visibleExtraColumns = savedCols ? savedCols.filter(c => state.extraColumns.includes(c)) : guessDefaultVisibleColumns(state.extraColumns);

    const savedGroup = loadGroupBy();
    state.groupByColumns = savedGroup ? savedGroup.filter(c => state.baseColumns.includes(c)) : guessDefaultGroupBy(state.baseColumns);

    // persiste seleção inicial
    saveVisibleColumns();
    saveGroupBy();

    state.rawRows = raw;

    // meta
    const meta = {
      importedAt: nowISO(),
      sourceFile: file.name,
      sheet,
      baseMonths: state.baseMonths,
      metricOptions: state.metricOptions,
      metric: state.metric,
      codeCol: state.colCode,
      descCol: state.colDesc,
      version: "v2.0"
    };
    state.meta = meta;

    saveToStorage(raw, meta);

    ui.dataInfo.textContent = `${meta.sourceFile} | Linhas: ${raw.length} | Meses: ${state.baseMonths.length} | Métrica: ${state.metric}`;
    ui.buildInfo.textContent = "Igarapé Digital | Verbas 6 Sigma | v2.0";

    // construir pickers e filtros
    buildColPicker();
    buildGroupPicker();
    fillStaticFilters();
    buildMetricSelect();
    buildVerbaSelect(raw);
    buildExtraFilters(state.rawRows);

    // agrega
    rebuildAggregation();
  }

  function initUI() {
    ui.buildInfo.textContent = "Igarapé Digital | Verbas 6 Sigma | v2.0";

    wireClearButton(ui.fSearch, el("btnClearSearch"));
    wireClearButton(ui.fWindow, el("btnClearWindow"));
    wireClearButton(ui.fMinZ, el("btnClearMinZ"));
    wireClearButton(ui.fMaxZ, el("btnClearMaxZ"));

    ui.btnApply.addEventListener("click", applyFilters);
    ui.btnReset.addEventListener("click", resetFilters);

    ui.fSearch.addEventListener("keydown", (e) => { if (e.key === "Enter") applyFilters(); });

    ui.fVerba.addEventListener("change", applyFilters);
    ui.fMetric.addEventListener("change", () => {
      const m = safeText(ui.fMetric.value);
      if (!m) return;
      state.metric = m;
      localStorage.setItem(STORAGE_METRIC, m);
      state.months = deriveMonthsForMetric(state.baseMonths, state.metric);
      buildMonthSelect();
      rebuildAggregation();
    });
    ui.fRefMonth.addEventListener("change", applyFilters);
    ui.fIgnoreZeros.addEventListener("change", applyFilters);
    ui.fStatus.addEventListener("change", applyFilters);

    ui.fWindow.addEventListener("input", () => { applyFilters(); });

    ui.fileInput.addEventListener("change", async (e) => {
      const f = e.target.files && e.target.files[0];
      if (!f) return;
      try {
        await handleFile(f);
      } catch (err) {
        console.error(err);
        ui.dataInfo.textContent = "Falha ao importar. Confirme que existem colunas 'Código', 'Descrição' e colunas mensais no formato 'MMM/AA - <Métrica>' (ex.: 'JUN/25 - Valor').";
      } finally {
        ui.fileInput.value = "";
      }
    });

    ui.btnClear.addEventListener("click", () => {
      clearStorage();
      state.rawRows = [];
      state.groups = [];
      state.filtered = [];
      state.meta = null;
      state.baseMonths = [];
      state.months = [];

      state.metricOptions = [];
      state.metric = "";
      state.baseColumns = [];
      state.extraColumns = [];
      state.visibleExtraColumns = [];
      state.groupByColumns = [];
      state.extraFilterValues = {};
      state.extraFilterModes = {};
      state.sort = { key: null, dir: 1 };

      ui.dataInfo.textContent = "Storage limpo. Importe um Excel.";
      ui.tableInfo.textContent = "";
      ui.tblHead.innerHTML = "";
      ui.tblBody.innerHTML = "";
      computeKPIs([]);
      if (ui.extraFilters) ui.extraFilters.innerHTML = "";
      if (ui.colPickerBody) ui.colPickerBody.innerHTML = "";
      if (ui.groupPickerBody) ui.groupPickerBody.innerHTML = "";
      if (ui.fMetric) ui.fMetric.innerHTML = "";
      renderDiag([]);
    });

    ui.btnExportTxt.addEventListener("click", () => {
      if (!state.filtered.length) return;
      exportTxt(state.filtered, state.meta);
    });

    if (ui.btnExportXlsx) {
      ui.btnExportXlsx.addEventListener("click", () => {
        if (!state.filtered.length) return;
        exportXlsx(state.filtered, state.meta);
      });
    }
  }

  function boot() {
    initUI();
    fillStaticFilters();

    const loaded = loadFromStorage();
    if (loaded) {
      const m = state.meta || {};
      ui.buildInfo.textContent = "Igarapé Digital | Verbas 6 Sigma | v2.0";

      // inferir headers
      const headers = Object.keys((state.rawRows[0] || {})).map(safeText);

      // colunas base
      state.colCode = m.codeCol || guessColumn(headers, [/^c[oó]digo$/i, /^codigo$/i, /\bc[oó]d\b/i, /\bcod\b/i]);
      state.colDesc = m.descCol || guessColumn(headers, [/^descri[cç][aã]o$/i, /descric/i, /\bdescricao\b/i]);

      // meses / métricas (recalcula a partir dos headers)
      const mb = detectMonthBlocks(headers);
      state.baseMonths = mb.months;
      state.metricOptions = detectNumericMetrics(state.rawRows, state.baseMonths, mb.metrics);
      if (!state.metricOptions.length) {
        state.metricOptions = mb.metrics.includes("Valor") ? ["Valor"] : mb.metrics.slice(0, 1);
      }

      const savedMetric = safeText(localStorage.getItem(STORAGE_METRIC) || "");
      const metaMetric = safeText(m.metric || "");
      const chosenMetric = (savedMetric && state.metricOptions.includes(savedMetric))
        ? savedMetric
        : (metaMetric && state.metricOptions.includes(metaMetric))
          ? metaMetric
          : (state.metricOptions.includes("Valor") ? "Valor" : state.metricOptions[0]);

      state.metric = chosenMetric;
      localStorage.setItem(STORAGE_METRIC, chosenMetric);
      state.months = deriveMonthsForMetric(state.baseMonths, state.metric);

      // colunas disponíveis para agrupamento / extras
      state.baseColumns = headers.filter(h => h && !mb.monthHeaderSet.has(h) && h !== state.colCode && h !== state.colDesc);
      state.extraColumns = state.baseColumns.slice().sort((a,b)=>a.localeCompare(b,"pt-BR"));

      const savedCols = loadVisibleColumns();
      state.visibleExtraColumns = savedCols ? savedCols.filter(c => state.extraColumns.includes(c)) : guessDefaultVisibleColumns(state.extraColumns);

      const savedGroup = loadGroupBy();
      state.groupByColumns = savedGroup ? savedGroup.filter(c => state.baseColumns.includes(c)) : guessDefaultGroupBy(state.baseColumns);

      ui.dataInfo.textContent = m.sourceFile
        ? `Storage carregado: ${m.sourceFile} | Linhas: ${(state.rawRows || []).length} | Meses: ${(state.baseMonths || []).length} | Métrica: ${state.metric}`
        : `Storage carregado | Linhas: ${(state.rawRows || []).length} | Meses: ${(state.baseMonths || []).length} | Métrica: ${state.metric}`;

      buildColPicker();
      buildGroupPicker();
      buildMetricSelect();
      buildVerbaSelect(state.rawRows);
      buildMonthSelect();
      buildExtraFilters(state.rawRows);

      rebuildAggregation();
    } else {
      ui.dataInfo.textContent = "Importe um Excel (.xlsx) com colunas Código/Descrição e colunas mensais no formato 'MMM/AA - <Métrica>' (ex.: 'JUN/25 - Valor', 'JUN/25 - Hora'). O sistema agrega e valida o último mês via 6 Sigma.";
      resetFilters();
    }
  }

  boot();
})();
