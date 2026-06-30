/* Integração do Dashboard com dados.xlsx (SheetJS carregado localmente). */
let PROJETOS_DB = [];
let ENTREGAS_DB = new Map();

const PERSISTENCE_DB = "brumas-do-amanha-dashboard";
const PERSISTENCE_STORE = "bases";
const PERSISTENCE_KEY = "base-ativa";
const LOCAL_STORAGE_KEY = "brumas-do-amanha:base-ativa";

const STATUS_ORDER = ["Não Iniciado", "Em Planejamento", "Em Execução", "Concluído", "Cancelado"];
const STATUS_COLORS = {
    "Não Iniciado": "#94a3b8",
    "Em Planejamento": "#a855f7",
    "Em Execução": "#f59e0b",
    "Concluído": "#10b981",
    "Cancelado": "#06b6d4"
};

function clean(value) {
    if (value === null || value === undefined) return "";
    return String(value).replace(/\s+/g, " ").trim();
}

function normalized(value) {
    return clean(value).normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase();
}

function normalizedHeader(value) {
    return normalized(value).replace(/[^a-z0-9]+/g, " ").trim();
}

function isMeaningful(value) {
    const text = normalized(value);
    return Boolean(text && text !== "na" && text !== "n a" && text !== "nao se aplica");
}

function escapeHTML(value) {
    return clean(value)
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#039;");
}

function toNumber(value) {
    if (typeof value === "number" && Number.isFinite(value)) return value;
    if (!isMeaningful(value)) return 0;
    let text = clean(value).replace(/R\$/gi, "").replace(/\s/g, "");
    if (/^-?\d{1,3}(\.\d{3})*,\d+$/.test(text)) text = text.replace(/\./g, "").replace(",", ".");
    else if (/^-?\d+,\d+$/.test(text)) text = text.replace(",", ".");
    else text = text.replace(/,/g, "");
    const number = Number(text);
    return Number.isFinite(number) ? number : 0;
}

function formatCurrency(value) {
    return toNumber(value).toLocaleString("pt-BR", {
        style: "currency",
        currency: "BRL",
        minimumFractionDigits: 0,
        maximumFractionDigits: 2
    });
}

function formatDate(value) {
    if (!isMeaningful(value) || value === 0) return "—";
    if (value instanceof Date && !Number.isNaN(value.getTime())) return value.toLocaleDateString("pt-BR");
    if (typeof value === "string" && /^\d{4}-\d{2}-\d{2}T/.test(value)) {
        const parsedDate = new Date(value);
        if (!Number.isNaN(parsedDate.getTime())) return parsedDate.toLocaleDateString("pt-BR");
    }
    if (typeof value === "number" && window.XLSX && XLSX.SSF) {
        const parsed = XLSX.SSF.parse_date_code(value);
        if (parsed) return `${String(parsed.d).padStart(2, "0")}/${String(parsed.m).padStart(2, "0")}/${parsed.y}`;
    }
    return clean(value);
}

function looksLikeStatus(value) {
    const text = normalized(value);
    return /conclu|execu|andamento|iniciado|planej|cancel|pausad|viabilidade|aprovacao|nao e projeto|verificar/.test(text);
}

function canonicalStatus(value) {
    const text = normalized(value);
    if (text.includes("nao e projeto")) return null;
    if (text.includes("conclu")) return "Concluído";
    if (text.includes("cancel")) return "Cancelado";
    if (text.includes("execu") || text.includes("andamento")) return "Em Execução";
    if (text.includes("nao iniciado")) return "Não Iniciado";
    return "Em Planejamento";
}

function projectKey(value) {
    return normalized(clean(value).replace(/^\d+(?:\.\d+)*\s*[-–—.:]?\s*/, ""))
        .replace(/[^a-z0-9]+/g, " ").trim();
}

function findSheet(workbook, expectedName) {
    const wanted = normalized(expectedName);
    const name = workbook.SheetNames.find(item => normalized(item) === wanted)
        || workbook.SheetNames.find(item => normalized(item).includes(wanted));
    return name ? workbook.Sheets[name] : null;
}

function boundedRange(sheet, lastColumn) {
    const source = XLSX.utils.decode_range(sheet["!ref"] || "A1:A1");
    return {
        s: { r: source.s.r, c: 0 },
        e: { r: source.e.r, c: Math.min(source.e.c, lastColumn) }
    };
}

function rowsWithAccessor(sheet) {
    const matrix = XLSX.utils.sheet_to_json(sheet, {
        header: 1,
        defval: null,
        raw: true,
        range: boundedRange(sheet, 29) // A:AD, campos efetivamente usados na base mestre
    });
    const headers = (matrix[0] || []).map(normalizedHeader);
    const indexes = new Map(headers.map((header, index) => [header, index]));
    return matrix.slice(1).map((row, index) => ({
        excelRow: index + 2,
        row,
        get(...names) {
            for (const name of names) {
                const position = indexes.get(normalizedHeader(name));
                if (position !== undefined && row[position] !== null && row[position] !== undefined) return row[position];
            }
            return null;
        }
    }));
}

function parseProjects(sheet) {
    let ignored = 0;
    let inferredStatus = 0;
    const projects = [];

    rowsWithAccessor(sheet).forEach(record => {
        const name = clean(record.get("Projeto Estratégico (Plano de Governo)"));
        if (!name) return;

        let leader = record.get("Líder do Projeto");
        let sourceStatus = record.get("Fase do Projeto");
        if (!isMeaningful(sourceStatus) && looksLikeStatus(leader)) {
            sourceStatus = leader;
            leader = null;
        }

        const status = canonicalStatus(sourceStatus);
        if (!status) {
            ignored += 1;
            return;
        }
        if (!isMeaningful(sourceStatus)) inferredStatus += 1;

        const eixo = clean(record.get("Eixo Estratégico")) || "Eixo não informado";
        const ppaProgram = record.get("Programa PPA");
        const programa = isMeaningful(ppaProgram)
            ? clean(ppaProgram)
            : `Sem programa PPA informado — ${eixo}`;
        const budgets = [2026, 2027, 2028, 2029].map(year => toNumber(record.get(`Orçamento ${year}`)));
        const investments = [2026, 2027, 2028, 2029].map(year => toNumber(record.get(`Investimento ${year}`)));

        projects.push({
            id: `projeto-${record.excelRow}`,
            excelRow: record.excelRow,
            nome: name,
            secretaria: clean(record.get("Secretaria Responsável")) || "Secretaria não informada",
            eixo,
            objetivoEixo: clean(record.get("Objetivo do Eixo")),
            meta: clean(record.get("Meta")) || "—",
            indicador: clean(record.get("Indicador")) || "—",
            programa,
            obj_prog: isMeaningful(record.get("Objetivo do Programa")) ? clean(record.get("Objetivo do Programa")) : "Não informado na planilha.",
            publico_prog: isMeaningful(record.get("Público - Alvo")) ? clean(record.get("Público - Alvo")) : "Não informado na planilha.",
            acaoPPA: isMeaningful(record.get("Projeto PPA")) ? clean(record.get("Projeto PPA")) : "Ação PPA não informada",
            objetivoAcao: isMeaningful(record.get("Objetivo da Ação")) ? clean(record.get("Objetivo da Ação")) : "Não informado na planilha.",
            produto: isMeaningful(record.get("Produto")) ? clean(record.get("Produto")) : "",
            lider: isMeaningful(leader) ? clean(leader) : "—",
            status,
            statusOriginal: isMeaningful(sourceStatus) ? clean(sourceStatus) : "Não informada",
            justificativa: isMeaningful(record.get("Justificativa")) ? clean(record.get("Justificativa")) : "Sem justificativa informada",
            inicio: record.get("Data de Início"),
            prazo: record.get("Prazo (Data de Conclusão)"),
            budgets,
            investments,
            orcado: budgets.reduce((sum, item) => sum + item, 0),
            real: investments.reduce((sum, item) => sum + item, 0)
        });
    });

    return { projects, ignored, inferredStatus };
}

function parseDeliveries(sheet) {
    const deliveries = new Map();
    if (!sheet) return deliveries;
    const rows = XLSX.utils.sheet_to_json(sheet, {
        header: 1,
        defval: null,
        raw: true,
        range: boundedRange(sheet, 23) // A:X, área de projetos e entregas
    });
    let currentProject = "";

    rows.slice(2).forEach(row => {
        if (isMeaningful(row[0])) currentProject = projectKey(row[0]);
        const description = clean(row[12]);
        if (!currentProject || !description) return;
        if (!deliveries.has(currentProject)) deliveries.set(currentProject, []);
        const list = deliveries.get(currentProject);
        list.push({
            id: `T${list.length + 1}`,
            descricao: description,
            responsavel: isMeaningful(row[13]) ? clean(row[13]) : "—",
            situacao: isMeaningful(row[14]) ? clean(row[14]) : "Não informada",
            prazo: formatDate(row[15])
        });
    });
    return deliveries;
}

function setImportState(kind, title, message) {
    const dot = document.getElementById("data-import-dot");
    const titleElement = document.getElementById("data-import-title");
    const statusElement = document.getElementById("data-import-status");
    if (dot) dot.className = `data-import-dot${kind ? ` ${kind}` : ""}`;
    if (titleElement) titleElement.textContent = title;
    if (statusElement) statusElement.textContent = message;
}

function openPersistenceDatabase() {
    return new Promise((resolve, reject) => {
        if (!window.indexedDB) {
            reject(new Error("IndexedDB indisponível"));
            return;
        }
        const request = window.indexedDB.open(PERSISTENCE_DB, 1);
        request.onupgradeneeded = () => {
            if (!request.result.objectStoreNames.contains(PERSISTENCE_STORE)) {
                request.result.createObjectStore(PERSISTENCE_STORE, { keyPath: "id" });
            }
        };
        request.onsuccess = () => resolve(request.result);
        request.onerror = () => reject(request.error || new Error("Falha ao abrir o armazenamento local"));
    });
}

async function saveStateInIndexedDB(record) {
    const database = await openPersistenceDatabase();
    try {
        await new Promise((resolve, reject) => {
            const transaction = database.transaction(PERSISTENCE_STORE, "readwrite");
            transaction.objectStore(PERSISTENCE_STORE).put(record);
            transaction.oncomplete = () => resolve();
            transaction.onerror = () => reject(transaction.error || new Error("Falha ao salvar a base"));
            transaction.onabort = () => reject(transaction.error || new Error("Gravação da base cancelada"));
        });
    } finally {
        database.close();
    }
}

async function readStateFromIndexedDB() {
    const database = await openPersistenceDatabase();
    try {
        return await new Promise((resolve, reject) => {
            const transaction = database.transaction(PERSISTENCE_STORE, "readonly");
            const request = transaction.objectStore(PERSISTENCE_STORE).get(PERSISTENCE_KEY);
            request.onsuccess = () => resolve(request.result || null);
            request.onerror = () => reject(request.error || new Error("Falha ao recuperar a base"));
        });
    } finally {
        database.close();
    }
}

function currentPersistenceRecord(sourceName) {
    return {
        id: PERSISTENCE_KEY,
        sourceName,
        savedAt: new Date().toISOString(),
        projects: PROJETOS_DB,
        deliveries: [...ENTREGAS_DB.entries()]
    };
}

async function persistDashboardState(sourceName) {
    const record = currentPersistenceRecord(sourceName);
    try {
        await saveStateInIndexedDB(record);
        return true;
    } catch (indexedDBError) {
        console.warn("IndexedDB indisponível; tentando localStorage:", indexedDBError);
        try {
            window.localStorage.setItem(LOCAL_STORAGE_KEY, JSON.stringify(record));
            return true;
        } catch (localStorageError) {
            console.warn("Não foi possível persistir a base:", localStorageError);
            return false;
        }
    }
}

async function readPersistedDashboardState() {
    try {
        const record = await readStateFromIndexedDB();
        if (record) return record;
    } catch (indexedDBError) {
        console.warn("Leitura pelo IndexedDB indisponível:", indexedDBError);
    }
    try {
        const value = window.localStorage.getItem(LOCAL_STORAGE_KEY);
        return value ? JSON.parse(value) : null;
    } catch (localStorageError) {
        console.warn("Leitura pelo localStorage indisponível:", localStorageError);
        return null;
    }
}

async function restorePersistedDashboardState() {
    const record = await readPersistedDashboardState();
    if (!record || !Array.isArray(record.projects) || !record.projects.length || !Array.isArray(record.deliveries)) {
        return false;
    }
    PROJETOS_DB = record.projects;
    ENTREGAS_DB = new Map(record.deliveries);
    refreshAllViews();
    const deliveryCount = [...ENTREGAS_DB.values()].reduce((sum, list) => sum + list.length, 0);
    const savedAt = record.savedAt ? new Date(record.savedAt) : null;
    const savedLabel = savedAt && !Number.isNaN(savedAt.getTime())
        ? ` • salvo em ${savedAt.toLocaleString("pt-BR")}`
        : "";
    setImportState(
        "success",
        `${PROJETOS_DB.length} projetos restaurados`,
        `${record.sourceName || "dados.xlsx"} • ${deliveryCount} entregas detalhadas${savedLabel}.`
    );
    return true;
}

async function importWorkbook(arrayBuffer, sourceName) {
    const button = document.getElementById("xlsx-import-button");
    if (button) button.disabled = true;
    setImportState("loading", "Importando planilha...", `Lendo ${sourceName}.`);
    try {
        if (!window.XLSX) throw new Error("O leitor de planilhas não foi carregado.");
        const workbook = XLSX.read(arrayBuffer, { type: "array", cellDates: true });
        const databaseSheet = findSheet(workbook, "BANCO DE DADOS");
        if (!databaseSheet) throw new Error('A aba obrigatória "BANCO DE DADOS" não foi encontrada.');

        const parsed = parseProjects(databaseSheet);
        if (!parsed.projects.length) throw new Error("Nenhum projeto válido foi encontrado na aba BANCO DE DADOS.");
        PROJETOS_DB = parsed.projects;
        ENTREGAS_DB = parseDeliveries(findSheet(workbook, "DESENVOLVIMENTO SOCIAL"));
        refreshAllViews();

        const persisted = await persistDashboardState(sourceName);

        const deliveryCount = [...ENTREGAS_DB.values()].reduce((sum, list) => sum + list.length, 0);
        const notes = [];
        if (parsed.ignored) notes.push(`${parsed.ignored} registro(s) “Não é projeto” ignorado(s)`);
        if (parsed.inferredStatus) notes.push(`${parsed.inferredStatus} fase(s) vazia(s) agrupada(s) em planejamento`);
        const suffix = notes.length ? ` ${notes.join("; ")}.` : "";
        setImportState(
            "success",
            `${PROJETOS_DB.length} projetos carregados`,
            `${sourceName} • ${deliveryCount} entregas detalhadas${persisted ? " • salvo neste navegador" : ""}.${suffix}`
        );
    } catch (error) {
        console.error(error);
        setImportState("error", "Não foi possível importar a planilha", error.message || "Arquivo inválido.");
    } finally {
        if (button) button.disabled = false;
    }
}

async function loadDefaultWorkbook() {
    try {
        setImportState("loading", "Carregando dados.xlsx...", "Buscando a planilha padrão ao lado do dashboard.");
        const response = await fetch("./dados.xlsx", { cache: "no-store" });
        if (!response.ok) throw new Error(`Falha HTTP ${response.status}`);
        await importWorkbook(await response.arrayBuffer(), "dados.xlsx");
    } catch (error) {
        console.warn("Carregamento automático indisponível:", error);
        setImportState("", "Selecione a planilha", "O carregamento automático não está disponível; use “Importar dados.xlsx”.");
    }
}

function fillSelect(elementId, values, firstOption) {
    const select = document.getElementById(elementId);
    if (!select) return;
    const previous = select.value;
    select.replaceChildren();
    if (firstOption) select.add(new Option(firstOption.label, firstOption.value));
    values.forEach(item => select.add(new Option(item.label ?? item, item.value ?? item)));
    const possible = [...select.options].some(option => option.value === previous);
    if (possible) select.value = previous;
    else if (select.options.length) select.value = select.options[0].value;
}

function uniqueSorted(values) {
    return [...new Set(values.filter(Boolean))].sort((a, b) => a.localeCompare(b, "pt-BR"));
}

function populateDashboardFilters() {
    fillSelect(
        "filter-secretaria",
        uniqueSorted(PROJETOS_DB.map(project => project.secretaria)),
        { label: "Todas as Secretarias (Geral)", value: "Todos" }
    );
    fillSelect(
        "filter-eixo",
        uniqueSorted(PROJETOS_DB.map(project => project.eixo)),
        { label: "Todos os Eixos Estratégicos", value: "Todos" }
    );
}

function populateFilters() {
    const programs = uniqueSorted(PROJETOS_DB.map(project => project.programa));
    fillSelect("programFilter", programs.map(item => ({ label: item, value: item })));
    filterProgramme(document.getElementById("programFilter")?.value || programs[0] || "");
}

function populateProjectFilter() {
    const projects = [...PROJETOS_DB].sort((a, b) => a.nome.localeCompare(b.nome, "pt-BR"));
    fillSelect("projectFilter", projects.map(project => ({ label: project.nome, value: project.id })));
    renderProjectReport(document.getElementById("projectFilter")?.value || projects[0]?.id || "");
}

function refreshAllViews() {
    populateDashboardFilters();
    populateFilters();
    populateProjectFilter();
    updateDashboard();
    renderEstrategico();
}

function switchTab(event, tabId) {
    document.querySelectorAll(".tab-content").forEach(content => content.classList.remove("active"));
    document.querySelectorAll(".tab-button").forEach(button => button.classList.remove("active"));
    document.getElementById(tabId)?.classList.add("active");
    event.currentTarget.classList.add("active");
}

function getSelectedTooltipVision() {
    return document.querySelector('input[name="tooltip-vision"]:checked')?.value || "eixo";
}

function updateTooltipVision() {
    updateDashboard();
}

function updateDashboard() {
    const selectedSecretaria = document.getElementById("filter-secretaria")?.value || "Todos";
    const selectedEixo = document.getElementById("filter-eixo")?.value || "Todos";
    const vision = getSelectedTooltipVision();
    const filtered = PROJETOS_DB.filter(project =>
        (selectedSecretaria === "Todos" || project.secretaria === selectedSecretaria)
        && (selectedEixo === "Todos" || project.eixo === selectedEixo)
    );

    const counts = Object.fromEntries(STATUS_ORDER.map(status => [status, filtered.filter(project => project.status === status).length]));
    const budget = filtered.reduce((sum, project) => sum + project.orcado, 0);
    const invested = filtered.reduce((sum, project) => sum + project.real, 0);
    const percentage = budget > 0 ? Math.round((invested / budget) * 100) : 0;

    document.getElementById("kpi-total-val").textContent = filtered.length;
    document.getElementById("kpi-nao-val").textContent = counts["Não Iniciado"];
    document.getElementById("kpi-plan-val").textContent = counts["Em Planejamento"];
    document.getElementById("kpi-exec-val").textContent = counts["Em Execução"];
    document.getElementById("kpi-concl-val").textContent = counts.Concluído;
    document.getElementById("kpi-canc-val").textContent = counts.Cancelado;
    document.getElementById("invest-real-txt").textContent = formatCurrency(invested);
    document.getElementById("invest-progress-bar").style.width = `${Math.min(percentage, 100)}%`;

    const footer = document.querySelector("#aba-kpi .invest-footer");
    if (footer) {
        footer.innerHTML = `Vr. Orçado: <span>${escapeHTML(formatCurrency(budget))}</span> <span style="color:#10b981; margin-left:8px;">(${percentage}%)</span><div style="font-size:11px; color:#64748b; margin-top:5px; text-transform:uppercase; font-weight:700;">Visão: ${vision === "eixo" ? "por Eixo Estratégico" : "por Secretaria"}</div>`;
    }

    renderTooltip("tip-total", filtered, "Todos", vision);
    renderTooltip("tip-nao", filtered, "Não Iniciado", vision);
    renderTooltip("tip-plan", filtered, "Em Planejamento", vision);
    renderTooltip("tip-exec", filtered, "Em Execução", vision);
    renderTooltip("tip-concl", filtered, "Concluído", vision);
    renderTooltip("tip-canc", filtered, "Cancelado", vision);
    renderInvestmentTooltip("tip-investimento", filtered, vision);
    renderBlockedTooltip("tip-matriz", filtered, vision);
    renderMatrix(filtered);
    renderBarCharts(filtered);
}

function grouped(dataset, keySelector, valueSelector = () => 1) {
    const result = new Map();
    dataset.forEach(item => {
        const key = keySelector(item) || "Não informado";
        result.set(key, (result.get(key) || 0) + valueSelector(item));
    });
    return result;
}

function renderTooltip(elementId, dataset, status, vision) {
    const target = document.getElementById(elementId);
    if (!target) return;
    const items = status === "Todos" ? dataset : dataset.filter(project => project.status === status);
    const values = grouped(items, project => vision === "eixo" ? project.eixo : project.secretaria);
    target.querySelector(".tooltip-title").textContent = vision === "eixo" ? "Distribuição por Eixo" : "Distribuição por Secretaria";
    target.querySelector(".tooltip-contents").innerHTML = values.size
        ? [...values].map(([label, value]) => `<div class="tooltip-row"><span>${escapeHTML(label)}</span><strong>${value}</strong></div>`).join("")
        : '<div class="tooltip-row"><span>Nenhum projeto</span><span>0</span></div>';
}

function renderInvestmentTooltip(elementId, dataset, vision) {
    const target = document.getElementById(elementId);
    if (!target) return;
    const values = grouped(dataset, project => vision === "eixo" ? project.eixo : project.secretaria, project => project.real);
    target.querySelector(".tooltip-title").textContent = vision === "eixo" ? "Investimento por Eixo" : "Investimento por Secretaria";
    const rows = [...values].filter(([, value]) => value > 0);
    target.querySelector(".tooltip-contents").innerHTML = rows.length
        ? rows.map(([label, value]) => `<div class="tooltip-row"><span>${escapeHTML(label)}</span><strong>${escapeHTML(formatCurrency(value))}</strong></div>`).join("")
        : '<div class="tooltip-row"><span>Sem investimentos</span><span>R$ 0</span></div>';
}

function renderBlockedTooltip(elementId, dataset, vision) {
    const target = document.getElementById(elementId);
    if (!target) return;
    const blocked = dataset.filter(project => project.status === "Não Iniciado" || project.status === "Cancelado");
    const values = grouped(blocked, project => vision === "eixo" ? project.eixo : project.secretaria);
    target.querySelector(".tooltip-title").textContent = vision === "eixo" ? "Parados/Cancelados por Eixo" : "Parados/Cancelados por Secretaria";
    target.querySelector(".tooltip-contents").innerHTML = values.size
        ? [...values].map(([label, value]) => `<div class="tooltip-row"><span>${escapeHTML(label)}</span><strong>${value} proj.</strong></div>`).join("")
        : '<div class="tooltip-row"><span>Nenhum projeto nessa condição</span><span>0</span></div>';
}

function renderMatrix(dataset) {
    const body = document.getElementById("matrix-body");
    const reasons = new Map();
    dataset.filter(project => project.status === "Não Iniciado" || project.status === "Cancelado").forEach(project => {
        const reason = project.justificativa || "Sem justificativa informada";
        if (!reasons.has(reason)) reasons.set(reason, { notStarted: 0, cancelled: 0 });
        const count = reasons.get(reason);
        if (project.status === "Cancelado") count.cancelled += 1;
        else count.notStarted += 1;
    });
    const rows = [...reasons].sort((a, b) => (b[1].notStarted + b[1].cancelled) - (a[1].notStarted + a[1].cancelled));
    body.innerHTML = rows.length
        ? rows.map(([reason, count]) => `<tr><td class="left-align">${escapeHTML(reason)}</td><td style="color:#64748b;font-weight:bold;">${count.notStarted}</td><td style="color:#06b6d4;font-weight:bold;">${count.cancelled}</td></tr>`).join("")
        : '<tr><td class="left-align">Nenhum registro nessa condição</td><td>0</td><td>0</td></tr>';
}

function renderChart(containerId, dataset, property) {
    const container = document.getElementById(containerId);
    const categories = uniqueSorted(dataset.map(project => project[property]));
    if (!categories.length) {
        container.innerHTML = '<div style="font-size:12px;color:#94a3b8;padding:12px;">Sem dados para exibir.</div>';
        return;
    }
    container.innerHTML = categories.map(category => {
        const items = dataset.filter(project => project[property] === category);
        const segments = STATUS_ORDER.map(status => {
            const percentage = items.length ? (items.filter(project => project.status === status).length / items.length) * 100 : 0;
            return `<div class="bar-segment" style="width:${percentage}%;background:${STATUS_COLORS[status]};" title="${escapeHTML(status)}"></div>`;
        }).join("");
        return `<div class="chart-row"><div class="chart-label">${escapeHTML(category)} (${items.length} Proj.)</div><div class="stacked-bar">${segments}</div></div>`;
    }).join("");
}

function renderBarCharts(dataset) {
    renderChart("axis-chart-container", dataset, "eixo");
    renderChart("secretaria-chart-container", dataset, "secretaria");
}

function filterProgramme(programName) {
    const list = PROJETOS_DB.filter(project => project.programa === programName);
    const first = list[0];
    if (!first) {
        document.getElementById("c2-table-projects").innerHTML = '<tr><td colspan="5">Nenhum programa carregado.</td></tr>';
        return;
    }
    document.getElementById("c2-nome-programa").textContent = programName;
    document.getElementById("c2-eixo").textContent = first.eixo;
    document.getElementById("c2-obj-programa").textContent = first.obj_prog;
    document.getElementById("c2-publico-programa").textContent = first.publico_prog;

    const notStarted = list.filter(project => project.status === "Não Iniciado").length;
    const cancelled = list.filter(project => project.status === "Cancelado").length;
    const executing = list.filter(project => project.status === "Em Execução").length;
    document.getElementById("c2-qtd-ni").textContent = notStarted;
    document.getElementById("c2-qtd-canc").textContent = cancelled;
    document.getElementById("c2-perc-exec").textContent = `${list.length ? Math.round((executing / list.length) * 100) : 0}%`;

    const budget = list.reduce((sum, project) => sum + project.orcado, 0);
    const invested = list.reduce((sum, project) => sum + project.real, 0);
    const percentage = budget ? Math.round((invested / budget) * 100) : 0;
    document.getElementById("c2-real").textContent = formatCurrency(invested);
    document.getElementById("c2-orcado").textContent = formatCurrency(budget);
    document.getElementById("c2-perc").textContent = `(${percentage}%)`;
    document.getElementById("c2-bar").style.width = `${Math.min(percentage, 100)}%`;

    document.getElementById("c2-table-projects").innerHTML = list.map(project => `<tr>
        <td class="left-align" style="padding-left:16px;">${escapeHTML(project.nome)}</td>
        <td>${escapeHTML(project.secretaria)}</td>
        <td><span style="color:${STATUS_COLORS[project.status]};font-weight:bold;">${escapeHTML(project.statusOriginal)}</span></td>
        <td>${escapeHTML(formatDate(project.prazo))}</td>
        <td style="text-align:right;padding-right:16px;font-weight:700;color:#334155;">${escapeHTML(formatCurrency(project.real))}</td>
    </tr>`).join("");
}

function renderProjectReport(projectId) {
    const project = PROJETOS_DB.find(item => item.id === projectId);
    if (!project) {
        document.getElementById("report-title").textContent = "Nenhum projeto carregado";
        document.getElementById("report-indicator-body").innerHTML = '<tr><td colspan="6">Importe dados.xlsx para visualizar o report.</td></tr>';
        document.getElementById("report-deliveries-body").innerHTML = '<tr><td colspan="5">Nenhuma entrega carregada.</td></tr>';
        return;
    }
    document.getElementById("report-eixo").textContent = project.eixo;
    document.getElementById("report-title").textContent = project.nome;
    document.getElementById("report-meta").textContent = project.meta;
    document.getElementById("report-publico").textContent = project.publico_prog;
    document.getElementById("report-ppa-title").textContent = project.acaoPPA;
    document.getElementById("report-ppa-desc").textContent = project.objetivoAcao;
    document.getElementById("report-indicator-body").innerHTML = `<tr>
        <td class="left-align" style="padding-left:16px;">${escapeHTML(project.indicador)}</td>
        <td class="badge-meta">${escapeHTML(project.meta)}</td>
        <td class="highlight-cell" style="color:${STATUS_COLORS[project.status]};font-weight:700;">${escapeHTML(project.statusOriginal)}</td>
        <td>${escapeHTML(formatDate(project.inicio))}</td>
        <td>${escapeHTML(formatDate(project.prazo))}</td>
        <td>${escapeHTML(project.lider)}</td>
    </tr>`;

    const deliveries = ENTREGAS_DB.get(projectKey(project.nome)) || [];
    document.getElementById("report-deliveries-body").innerHTML = deliveries.length
        ? deliveries.map(delivery => {
            const status = canonicalStatus(delivery.situacao) || "Em Planejamento";
            return `<tr><td>${escapeHTML(delivery.id)}</td><td class="left-align" style="padding-left:16px;">${escapeHTML(delivery.descricao)}</td><td style="color:#64748b;">${escapeHTML(delivery.responsavel)}</td><td><span class="badge-progress-gantt" style="background:${STATUS_COLORS[status]};">${escapeHTML(delivery.situacao)}</span></td><td>${escapeHTML(delivery.prazo)}</td></tr>`;
        }).join("")
        : '<tr><td colspan="5" style="padding:20px;color:#64748b;">A planilha não possui entregas detalhadas para este projeto.</td></tr>';
    document.getElementById("report-monitoring-date").textContent = "Fonte: dados.xlsx";
}

function renderEstrategico() {
    const body = document.getElementById("strat-table-body");
    if (!body) return;
    body.innerHTML = PROJETOS_DB.length
        ? PROJETOS_DB.map(project => `<tr>
            <td class="left-align" style="padding-left:16px;">${escapeHTML(project.programa)}</td>
            <td style="text-align:left;font-size:12px;color:#334155;">${escapeHTML(project.nome)}</td>
            <td style="text-align:left;font-size:12px;color:#475569;">${escapeHTML(project.indicador)}</td>
            <td style="text-align:left;font-size:12px;color:#475569;">${escapeHTML(project.meta)}</td>
            <td class="highlight-cell" style="color:${STATUS_COLORS[project.status]};font-weight:700;">${escapeHTML(project.statusOriginal)}</td>
            <td style="font-weight:600;">${escapeHTML(formatCurrency(project.orcado))}</td>
            <td style="font-weight:600;">${escapeHTML(formatCurrency(project.real))}</td>
        </tr>`).join("")
        : '<tr><td colspan="7">Importe dados.xlsx para visualizar o painel estratégico.</td></tr>';
}

document.addEventListener("DOMContentLoaded", async () => {
    const input = document.getElementById("xlsx-file-input");
    document.getElementById("xlsx-import-button").addEventListener("click", () => input.click());
    input.addEventListener("change", async event => {
        const file = event.target.files?.[0];
        if (!file) return;
        await importWorkbook(await file.arrayBuffer(), file.name);
        event.target.value = "";
    });

    refreshAllViews();
    if (await restorePersistedDashboardState()) return;
    if (window.location.protocol === "file:") {
        setImportState("", "Selecione a planilha", "Por segurança do navegador, use “Importar dados.xlsx” ao abrir este HTML diretamente.");
    } else {
        loadDefaultWorkbook();
    }
});
