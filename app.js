const CATEGORY_SEED = [
  {
    name: "Energia elétrica",
    fe: 0.0545,
    unit: "kWh/ano",
    method: "MME/SIN. Adaptado na metodologia (2025).",
    hasUsefulLife: false,
    lifeSpan: "",
  },
  {
    name: "Água",
    fe: 0.344,
    unit: "m³/ano",
    method: "IPCC (2006; 2019) e inventários nacionais de GEE (MCTI).",
    hasUsefulLife: false,
    lifeSpan: "",
  },
  {
    name: "Papel virgem",
    fe: 1.84,
    unit: "kg/ano",
    method: "Tabela de conversão de papel A4.",
    hasUsefulLife: false,
    lifeSpan: "",
  },
  {
    name: "Papel reciclado",
    fe: 0.61,
    unit: "kg/ano",
    method: "Tabela de conversão de papel A4.",
    hasUsefulLife: false,
    lifeSpan: "",
  },
  {
    name: "Áreas construídas",
    fe: 520,
    unit: "m²",
    method: "Amaral (2010).",
    hasUsefulLife: true,
    lifeSpan: "50",
  },
  {
    name: "Gasolina",
    fe: 2.21,
    unit: "L/ano",
    method: "CentroClima/COPPE/UFRJ; IPCC (2006; 2019).",
    hasUsefulLife: false,
    lifeSpan: "",
  },
  {
    name: "Diesel",
    fe: 2.63,
    unit: "L/ano",
    method: "CentroClima/COPPE/UFRJ; IPCC (2006; 2019).",
    hasUsefulLife: false,
    lifeSpan: "",
  },
  {
    name: "Refeições institucionais",
    fe: 2,
    unit: "refeições/ano",
    method: "Poore & Nemecek (2018); FAO (2013).",
    hasUsefulLife: false,
    lifeSpan: "",
  },
  {
    name: "Resíduos sólidos (aterro sanitário)",
    fe: 1200,
    unit: "t/ano",
    method: "IPCC (2006; 2019); WRI; WBCSD (2004).",
    hasUsefulLife: false,
    lifeSpan: "",
  },
];

const STORAGE_KEY = "pegada-ecologica-state-v1";
const STORAGE_SCHEMA = 2;

const categoriesBody = document.getElementById("categoriesBody");
const addCategoryButton = document.getElementById("addCategoryButton");
const resetButton = document.getElementById("resetButton");
const exportButton = document.getElementById("exportButton");
const exportJsonButton = document.getElementById("exportJsonButton");
const importJsonButton = document.getElementById("importJsonButton");
const importJsonInput = document.getElementById("importJsonInput");

const totalEmission = document.getElementById("totalEmission");
const totalArea = document.getElementById("totalArea");
const totalFootprint = document.getElementById("totalFootprint");
const perCapita = document.getElementById("perCapita");
const breakdown = document.getElementById("breakdown");
const warningBox = document.getElementById("warningBox");
const saveInfo = document.getElementById("saveInfo");

const baseYearInput = document.getElementById("baseYear");
const unitNameInput = document.getElementById("unitName");
const absorptionInput = document.getElementById("absorptionFactor");
const equivalenceInput = document.getElementById("equivalenceFactor");
const populationInput = document.getElementById("population");
const useGhaInput = document.getElementById("useGha");

const createId = () => {
  if (typeof crypto !== "undefined" && typeof crypto.randomUUID === "function") {
    return crypto.randomUUID();
  }
  return `id-${Date.now()}-${Math.random().toString(16).slice(2)}`;
};

const buildDefaultCategories = () =>
  CATEGORY_SEED.map((item) => ({
    ...item,
    id: createId(),
    enabled: true,
    consumption: "",
    custom: false,
  }));

let categories = buildDefaultCategories();

const formatNumber = (value, maxFraction = 2) =>
  new Intl.NumberFormat("pt-BR", {
    minimumFractionDigits: 2,
    maximumFractionDigits: maxFraction,
  }).format(Number.isFinite(value) ? value : 0);

const parseFieldNumber = (value) => {
  if (typeof value === "number") return value;
  if (value === null || value === undefined || value === "") return NaN;
  let raw = String(value).trim();
  if (!raw) return NaN;

  raw = raw.replace(/\s+/g, "");

  // Aceita formatos com virgula decimal, incluindo milhares no padrao pt-BR:
  // 1234,56 | 1.234,56 | 1234.56
  if (raw.includes(",")) {
    raw = raw.replace(/\./g, "").replace(/,/g, ".");
  }

  return Number(raw);
};

const normalizeText = (value) =>
  String(value ?? "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase();

const isBuiltAreaCategoryName = (name) => normalizeText(name).includes("area construida");

const safePositive = (value, fallback) => {
  const numeric = parseFieldNumber(value);
  return Number.isFinite(numeric) && numeric > 0 ? numeric : fallback;
};

const escapeHtml = (text) =>
  String(text)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;");

const escapeAttr = (text) => escapeHtml(text).replaceAll('"', "&quot;").replaceAll("'", "&#39;");

const nowLabel = () =>
  new Intl.DateTimeFormat("pt-BR", {
    dateStyle: "short",
    timeStyle: "medium",
  }).format(new Date());

const computeRow = (category, absorptionFactor, equivalenceFactor) => {
  const enabled = Boolean(category.enabled);
  const consumption = parseFieldNumber(category.consumption);
  const fe = parseFieldNumber(category.fe);
  const hasUsefulLife = Boolean(category.hasUsefulLife);
  const lifeSpan = parseFieldNumber(category.lifeSpan);

  if (
    !enabled ||
    !Number.isFinite(consumption) ||
    consumption <= 0 ||
    !Number.isFinite(fe) ||
    fe <= 0 ||
    (hasUsefulLife && (!Number.isFinite(lifeSpan) || lifeSpan <= 0))
  ) {
    return { kg: 0, ton: 0, area: 0, gha: 0, valid: false };
  }

  const annualizationDivisor = hasUsefulLife ? lifeSpan : 1;
  const kg = (consumption * fe) / annualizationDivisor;
  const ton = kg / 1000;
  const area = ton / absorptionFactor;
  const gha = area * equivalenceFactor;
  return { kg, ton, area, gha, valid: true };
};

const captureFocusState = () => {
  const active = document.activeElement;
  if (!(active instanceof HTMLInputElement)) return null;
  const dataIndex = active.dataset.index;
  if (typeof dataIndex === "undefined") return null;
  const rowClass = Array.from(active.classList).find((item) => item.startsWith("row-"));
  if (!rowClass) return null;
  return {
    index: dataIndex,
    rowClass,
    start: active.selectionStart,
    end: active.selectionEnd,
  };
};

const restoreFocusState = (focusState) => {
  if (!focusState) return;
  const selector = `.${focusState.rowClass}[data-index="${focusState.index}"]`;
  const input = categoriesBody.querySelector(selector);
  if (!(input instanceof HTMLInputElement)) return;
  input.focus();
  if (typeof focusState.start === "number" && typeof focusState.end === "number") {
    try {
      input.setSelectionRange(focusState.start, focusState.end);
    } catch (_error) {
      // setSelectionRange can fail on non-text inputs in some browsers.
    }
  }
};

const collectComputation = () => {
  const absorptionFactor = safePositive(absorptionInput.value, 6.27);
  const equivalenceFactor = safePositive(equivalenceInput.value, 1.37);
  const useGha = useGhaInput.checked;
  const population = parseFieldNumber(populationInput.value);

  const rows = categories.map((category) => ({
    category,
    metrics: computeRow(category, absorptionFactor, equivalenceFactor),
  }));

  const totalTon = rows.reduce((sum, row) => sum + row.metrics.ton, 0);
  const totalHa = rows.reduce((sum, row) => sum + row.metrics.area, 0);
  const totalGha = rows.reduce((sum, row) => sum + row.metrics.gha, 0);
  const perCapitaValue =
    Number.isFinite(population) && population > 0 ? (useGha ? totalGha : totalHa) / population : NaN;

  return {
    absorptionFactor,
    equivalenceFactor,
    useGha,
    population,
    rows,
    totalTon,
    totalHa,
    totalGha,
    perCapitaValue,
  };
};

const rowTemplate = (category, metrics, index) => {
  const deleteDisabled = category.custom ? "" : "disabled";
  const deleteLabel = category.custom ? "Remover" : "Padrão";
  const usefulLifeCell = category.hasUsefulLife
    ? `<input class="row-life" data-index="${index}" type="text" inputmode="decimal" value="${escapeAttr(
        category.lifeSpan
      )}" placeholder="50" />`
    : '<span class="muted-cell">-</span>';

  return `
    <tr data-id="${escapeAttr(category.id)}">
      <td data-label="Ativar">
        <input class="row-enabled" data-index="${index}" type="checkbox" ${
    category.enabled ? "checked" : ""
  } />
      </td>
      <td data-label="Categoria">
        <input class="row-name" data-index="${index}" type="text" value="${escapeAttr(category.name)}" />
      </td>
      <td data-label="Consumo anual">
        <input class="row-input" data-index="${index}" type="text" inputmode="decimal" value="${escapeAttr(
    category.consumption
  )}" placeholder="0" />
      </td>
      <td data-label="Vida útil (anos)">${usefulLifeCell}</td>
      <td data-label="Unidade">
        <input class="row-unit" data-index="${index}" type="text" value="${escapeAttr(category.unit)}" />
      </td>
      <td data-label="FE">
        <input class="row-fe" data-index="${index}" type="text" inputmode="decimal" value="${escapeAttr(
    category.fe
  )}" placeholder="FE" />
      </td>
      <td data-label="Emissão (kg CO₂)"><output>${formatNumber(metrics.kg, 4)}</output></td>
      <td data-label="Emissão (t CO₂)"><output>${formatNumber(metrics.ton, 6)}</output></td>
      <td data-label="Área (ha/ano)"><output>${formatNumber(metrics.area, 6)}</output></td>
      <td data-label="Pegada (gha/ano)"><output>${formatNumber(metrics.gha, 6)}</output></td>
      <td data-label="Base metodológica">
        <input class="row-method" data-index="${index}" type="text" value="${escapeAttr(category.method)}" />
      </td>
      <td data-label="Ação">
        <button class="remove-btn" data-index="${index}" type="button" ${deleteDisabled}>${deleteLabel}</button>
      </td>
    </tr>
  `;
};

const renderBreakdown = (rows, useGha) => {
  const total = rows.reduce((sum, item) => sum + (useGha ? item.metrics.gha : item.metrics.area), 0);

  if (total <= 0) {
    breakdown.innerHTML = "<small>Nenhuma contribuição para exibir. Informe consumos válidos.</small>";
    return;
  }

  const bars = rows
    .filter((item) => item.metrics.valid)
    .sort((a, b) => (useGha ? b.metrics.gha - a.metrics.gha : b.metrics.area - a.metrics.area))
    .map((row) => {
      const amount = useGha ? row.metrics.gha : row.metrics.area;
      const percent = (amount / total) * 100;
      const amountLabel = `${formatNumber(amount, 6)} ${useGha ? "gha/ano" : "ha/ano"}`;
      return `
        <div class="bar-row">
          <div class="bar-label">${escapeHtml(row.category.name)}</div>
          <div class="bar-track">
            <div class="bar-fill" style="width: ${Math.max(percent, 1.5)}%"></div>
          </div>
          <div class="bar-value">${amountLabel}</div>
        </div>
      `;
    })
    .join("");

  breakdown.innerHTML = bars || "<small>Nenhuma contribuição para exibir.</small>";
};

const updateWarnings = () => {
  const missingFe = categories.filter((category) => {
    const consumption = parseFieldNumber(category.consumption);
    const fe = parseFieldNumber(category.fe);
    return (
      category.enabled &&
      Number.isFinite(consumption) &&
      consumption > 0 &&
      (!Number.isFinite(fe) || fe <= 0)
    );
  });

  const invalidUsefulLife = categories.filter((category) => {
    const consumption = parseFieldNumber(category.consumption);
    const fe = parseFieldNumber(category.fe);
    const lifeSpan = parseFieldNumber(category.lifeSpan);
    return (
      category.enabled &&
      category.hasUsefulLife &&
      Number.isFinite(consumption) &&
      consumption > 0 &&
      Number.isFinite(fe) &&
      fe > 0 &&
      (!Number.isFinite(lifeSpan) || lifeSpan <= 0)
    );
  });

  if (missingFe.length === 0 && invalidUsefulLife.length === 0) {
    warningBox.hidden = true;
    warningBox.textContent = "";
    return;
  }

  warningBox.hidden = false;
  const messages = [];

  if (missingFe.length > 0) {
    const names = missingFe.map((category) => category.name).join(", ");
    messages.push(
      `As categorias ${names} possuem consumo informado, mas FE inválido ou vazio. Elas não entraram no total até o FE ser corrigido.`
    );
  }

  if (invalidUsefulLife.length > 0) {
    const names = invalidUsefulLife.map((category) => category.name).join(", ");
    messages.push(
      `As categorias ${names} possuem consumo e FE válidos, mas vida útil inválida. Corrija a vida útil (anos) para incluir essas categorias no total.`
    );
  }

  warningBox.textContent = messages.join(" ");
};

const setSaveInfo = (message, isError = false) => {
  saveInfo.textContent = message;
  saveInfo.classList.toggle("error", isError);
};

const toSerializableCategory = (category) => ({
  id: category.id,
  name: category.name,
  fe: category.fe,
  unit: category.unit,
  method: category.method,
  enabled: Boolean(category.enabled),
  consumption: category.consumption,
  custom: Boolean(category.custom),
  hasUsefulLife: Boolean(category.hasUsefulLife),
  lifeSpan:
    category.lifeSpan === null || category.lifeSpan === undefined ? "" : category.lifeSpan,
});

const buildSnapshot = () => ({
  schema: STORAGE_SCHEMA,
  savedAt: new Date().toISOString(),
  baseYear: baseYearInput.value || "",
  unitName: unitNameInput.value || "",
  absorptionFactor: absorptionInput.value || "",
  equivalenceFactor: equivalenceInput.value || "",
  population: populationInput.value || "",
  useGha: useGhaInput.checked,
  categories: categories.map(toSerializableCategory),
});

const normalizeCategory = (rawCategory, index) => {
  const fallbackName = `Categoria ${index + 1}`;
  const name =
    typeof rawCategory?.name === "string" && rawCategory.name.trim()
      ? rawCategory.name.trim()
      : fallbackName;
  const hasUsefulLife =
    typeof rawCategory?.hasUsefulLife === "boolean"
      ? rawCategory.hasUsefulLife
      : isBuiltAreaCategoryName(name);

  return {
    id:
      typeof rawCategory?.id === "string" && rawCategory.id.trim()
        ? rawCategory.id
        : createId(),
    name,
    fe:
      rawCategory?.fe === "" || rawCategory?.fe === null || rawCategory?.fe === undefined
        ? ""
        : rawCategory.fe,
    unit:
      typeof rawCategory?.unit === "string" && rawCategory.unit.trim()
        ? rawCategory.unit.trim()
        : "unidade/ano",
    method:
      typeof rawCategory?.method === "string" && rawCategory.method.trim()
        ? rawCategory.method.trim()
        : "Base metodológica não informada.",
    enabled: rawCategory?.enabled !== false,
    consumption:
      rawCategory?.consumption === null || rawCategory?.consumption === undefined
        ? ""
        : rawCategory.consumption,
    custom: typeof rawCategory?.custom === "boolean" ? rawCategory.custom : true,
    hasUsefulLife,
    lifeSpan:
      rawCategory?.lifeSpan === null || rawCategory?.lifeSpan === undefined || rawCategory.lifeSpan === ""
        ? hasUsefulLife
          ? "50"
          : ""
        : rawCategory.lifeSpan,
  };
};

const mergeLegacyWithMethodology = (rawCategories) => {
  const normalizedLegacy = Array.isArray(rawCategories)
    ? rawCategories.map((category, index) => normalizeCategory(category, index))
    : [];

  const legacyByName = new Map(
    normalizedLegacy.map((category) => [normalizeText(category.name), category])
  );

  const migratedDefaults = CATEGORY_SEED.map((seed, index) => {
    const key = normalizeText(seed.name);
    const legacy = legacyByName.get(key);
    return normalizeCategory(
      {
        ...seed,
        id:
          typeof legacy?.id === "string" && legacy.id.trim()
            ? legacy.id
            : createId(),
        enabled: legacy?.enabled !== false,
        consumption:
          legacy?.consumption === null || legacy?.consumption === undefined
            ? ""
            : legacy.consumption,
        custom: false,
        lifeSpan:
          legacy?.lifeSpan === null || legacy?.lifeSpan === undefined || legacy.lifeSpan === ""
            ? seed.lifeSpan ?? ""
            : legacy.lifeSpan,
      },
      index
    );
  });

  const customLegacy = normalizedLegacy.filter((category) => category.custom);
  return [...migratedDefaults, ...customLegacy];
};

const applySnapshot = (snapshot) => {
  const source =
    snapshot && typeof snapshot === "object" && snapshot.state && typeof snapshot.state === "object"
      ? snapshot.state
      : snapshot;
  const snapshotSchema =
    snapshot && typeof snapshot === "object" && snapshot.schema !== undefined
      ? Number(snapshot.schema)
      : source && typeof source === "object" && source.schema !== undefined
        ? Number(source.schema)
        : NaN;

  if (!source || typeof source !== "object") return false;

  if (source.baseYear !== undefined) baseYearInput.value = String(source.baseYear ?? "");
  if (source.unitName !== undefined) unitNameInput.value = String(source.unitName ?? "");
  if (source.absorptionFactor !== undefined) {
    absorptionInput.value = String(source.absorptionFactor ?? "");
  }
  if (source.equivalenceFactor !== undefined) {
    equivalenceInput.value = String(source.equivalenceFactor ?? "");
  }
  if (source.population !== undefined) populationInput.value = String(source.population ?? "");
  if (source.useGha !== undefined) useGhaInput.checked = Boolean(source.useGha);

  if (Array.isArray(source.categories) && source.categories.length > 0) {
    categories =
      Number.isFinite(snapshotSchema) && snapshotSchema >= STORAGE_SCHEMA
        ? source.categories.map((category, index) => normalizeCategory(category, index))
        : mergeLegacyWithMethodology(source.categories);
  } else {
    categories = buildDefaultCategories();
  }

  return true;
};

const saveToLocalStorage = () => {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(buildSnapshot()));
    setSaveInfo(`Dados salvos automaticamente. Última gravação: ${nowLabel()}`);
  } catch (_error) {
    setSaveInfo("Não foi possível salvar automaticamente neste navegador.", true);
  }
};

const loadFromLocalStorage = () => {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return false;
    const parsed = JSON.parse(raw);
    const ok = applySnapshot(parsed);
    if (!ok) return false;
    return true;
  } catch (_error) {
    setSaveInfo("Falha ao ler dados salvos. O estado padrão será usado.", true);
    return false;
  }
};

const render = (options = {}) => {
  const focusState = options.preserveFocus ? captureFocusState() : null;
  const computation = collectComputation();

  categoriesBody.innerHTML = computation.rows
    .map((row, index) => rowTemplate(row.category, row.metrics, index))
    .join("");

  restoreFocusState(focusState);

  totalEmission.textContent = `${formatNumber(computation.totalTon, 6)} t CO₂/ano`;
  totalArea.textContent = `${formatNumber(computation.totalHa, 6)} ha/ano`;
  totalFootprint.textContent = computation.useGha
    ? `${formatNumber(computation.totalGha, 6)} gha/ano`
    : `${formatNumber(computation.totalHa, 6)} ha/ano`;

  if (Number.isFinite(computation.population) && computation.population > 0) {
    const unit = computation.useGha ? "gha/pessoa/ano" : "ha/pessoa/ano";
    perCapita.textContent = `${formatNumber(computation.perCapitaValue, 6)} ${unit}`;
  } else {
    perCapita.textContent = "Informe a população";
  }

  renderBreakdown(computation.rows, computation.useGha);
  updateWarnings();
  saveToLocalStorage();
};

const updateCategory = (index, field, value) => {
  if (index < 0 || index >= categories.length) return;
  categories[index][field] = value;
  render({ preserveFocus: true });
};

document.addEventListener("input", (event) => {
  const target = event.target;
  if (!(target instanceof HTMLElement)) return;

  const indexedInput = target.dataset.index;
  const index = Number(indexedInput);
  const hasIndex = Number.isInteger(index);

  if (hasIndex && target.classList.contains("row-enabled") && target instanceof HTMLInputElement) {
    updateCategory(index, "enabled", target.checked);
    return;
  }
  if (hasIndex && target.classList.contains("row-name") && target instanceof HTMLInputElement) {
    updateCategory(index, "name", target.value.trim() || "Categoria sem nome");
    return;
  }
  if (hasIndex && target.classList.contains("row-input") && target instanceof HTMLInputElement) {
    updateCategory(index, "consumption", target.value);
    return;
  }
  if (hasIndex && target.classList.contains("row-unit") && target instanceof HTMLInputElement) {
    updateCategory(index, "unit", target.value.trim() || "unidade/ano");
    return;
  }
  if (hasIndex && target.classList.contains("row-fe") && target instanceof HTMLInputElement) {
    updateCategory(index, "fe", target.value);
    return;
  }
  if (hasIndex && target.classList.contains("row-life") && target instanceof HTMLInputElement) {
    updateCategory(index, "lifeSpan", target.value);
    return;
  }
  if (hasIndex && target.classList.contains("row-method") && target instanceof HTMLInputElement) {
    updateCategory(index, "method", target.value.trim() || "Base metodológica não informada.");
    return;
  }

  if (
    target.id === "absorptionFactor" ||
    target.id === "equivalenceFactor" ||
    target.id === "population" ||
    target.id === "baseYear" ||
    target.id === "unitName"
  ) {
    render();
  }
});

document.addEventListener("change", (event) => {
  const target = event.target;
  if (!(target instanceof HTMLElement)) return;
  if (target.id === "useGha") {
    render();
  }
});

document.addEventListener("click", (event) => {
  const target = event.target;
  if (!(target instanceof HTMLElement)) return;
  if (!target.classList.contains("remove-btn")) return;

  const index = Number(target.dataset.index);
  if (!Number.isInteger(index) || !categories[index] || !categories[index].custom) return;

  categories.splice(index, 1);
  render();
});

addCategoryButton.addEventListener("click", () => {
  categories.push({
    id: createId(),
    name: "Nova categoria",
    fe: "",
    unit: "unidade/ano",
    method: "Categoria personalizada.",
    enabled: true,
    consumption: "",
    custom: true,
    hasUsefulLife: false,
    lifeSpan: "",
  });
  render();
});

resetButton.addEventListener("click", () => {
  categories = categories.map((category) => ({
    ...category,
    consumption: "",
    enabled: true,
  }));
  render();
});

const downloadBlob = (content, filename, mimeType) => {
  const blob = new Blob([content], { type: mimeType });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = filename;
  anchor.click();
  URL.revokeObjectURL(url);
};

const buildExcelCategoryRows = (computation) =>
  computation.rows.map((row) => ({
    Categoria: row.category.name,
    Ativa: row.category.enabled ? "Sim" : "Não",
    "Consumo anual":
      row.category.consumption === "" ? "" : parseFieldNumber(row.category.consumption),
    "Vida útil (anos)": row.category.hasUsefulLife
      ? parseFieldNumber(row.category.lifeSpan) || ""
      : "",
    Unidade: row.category.unit,
    "Fator de emissão (kg CO₂/unidade)":
      row.category.fe === "" ? "" : parseFieldNumber(row.category.fe),
    "Emissão (kg CO₂)": row.metrics.kg,
    "Emissão (t CO₂)": row.metrics.ton,
    "Área (ha/ano)": row.metrics.area,
    "Pegada (gha/ano)": row.metrics.gha,
    "Base metodológica": row.category.method,
  }));

const buildExcelSummaryRows = (computation) => {
  const perCapitaUnit = computation.useGha ? "gha/pessoa/ano" : "ha/pessoa/ano";
  return [
    { Campo: "Ano-base", Valor: baseYearInput.value || "" },
    { Campo: "Unidade", Valor: unitNameInput.value || "" },
    { Campo: "Fator de absorção (t/ha/ano)", Valor: computation.absorptionFactor },
    { Campo: "Fator de equivalência", Valor: computation.equivalenceFactor },
    { Campo: "População", Valor: populationInput.value || "" },
    { Campo: "Consolidado em", Valor: computation.useGha ? "gha" : "ha" },
    { Campo: "Emissão total (t CO₂/ano)", Valor: computation.totalTon },
    { Campo: "Área total (ha/ano)", Valor: computation.totalHa },
    { Campo: "Pegada total (gha/ano)", Valor: computation.totalGha },
    {
      Campo: `Pegada per capita (${perCapitaUnit})`,
      Valor: Number.isFinite(computation.perCapitaValue) ? computation.perCapitaValue : "",
    },
    { Campo: "Data de exportação", Valor: new Date().toISOString() },
  ];
};

exportButton.addEventListener("click", () => {
  if (typeof XLSX === "undefined") {
    setSaveInfo("Biblioteca de Excel não carregada. Verifique a internet e tente de novo.", true);
    return;
  }

  const computation = collectComputation();
  const categoryRows = buildExcelCategoryRows(computation);
  const summaryRows = buildExcelSummaryRows(computation);

  const workbook = XLSX.utils.book_new();
  const summarySheet = XLSX.utils.json_to_sheet(summaryRows);
  const categoriesSheet = XLSX.utils.json_to_sheet(categoryRows);
  XLSX.utils.book_append_sheet(workbook, summarySheet, "Resumo");
  XLSX.utils.book_append_sheet(workbook, categoriesSheet, "Categorias");

  const year = baseYearInput.value || "ano-base";
  XLSX.writeFile(workbook, `pegada-ecologica-${year}.xlsx`);
  setSaveInfo(`Planilha Excel exportada com sucesso em ${nowLabel()}.`);
});

exportJsonButton.addEventListener("click", () => {
  const payload = {
    ...buildSnapshot(),
    exportedAt: new Date().toISOString(),
  };
  const year = baseYearInput.value || "ano-base";
  downloadBlob(JSON.stringify(payload, null, 2), `pegada-ecologica-${year}.json`, "application/json");
  setSaveInfo(`JSON exportado com sucesso em ${nowLabel()}.`);
});

importJsonButton.addEventListener("click", () => {
  importJsonInput.click();
});

importJsonInput.addEventListener("change", async (event) => {
  const file = event.target.files?.[0];
  if (!file) return;

  try {
    const text = await file.text();
    const parsed = JSON.parse(text);
    const ok = applySnapshot(parsed);
    if (!ok) throw new Error("Arquivo JSON inválido.");
    render();
    setSaveInfo(`JSON importado com sucesso em ${nowLabel()}.`);
  } catch (_error) {
    setSaveInfo("Falha ao importar JSON. Verifique o formato do arquivo.", true);
  } finally {
    importJsonInput.value = "";
  }
});

loadFromLocalStorage();
render();
