(function () {
  "use strict";

  const FUNCTIONS = ["Finance", "Sales", "Marketing", "HR", "Supply", "D&T"];
  const CATEGORY_ORDER = [
    "Staff Cost - Indirect",
    "Staff Cost - SG&A",
    "Staff Cost - SG&A or Indirect",
    "External Personnel Costs",
    "IT Costs",
    "Outside Consulting, Fees",
    "Other Costs"
  ];
  const CATEGORY_COLORS = ["#0b4f3a","#176b4d","#2f8664","#58a17e","#517c91","#b08448","#7b847f"];
  const FRANCE_SPLIT = [{ cluster: "France - Long Life", weight: .48 }, { cluster: "France - Traiteur", weight: .38 }, { cluster: "Champiland", weight: .14 }];
  const CE_SPLIT = [{ country: "POL", weight: .24 }, { country: "CZE", weight: .18 }, { country: "HUN", weight: .13 }, { country: "ROU", weight: .15 }, { country: "SVK", weight: .07 }, { country: "HRV", weight: .08 }, { country: "GRC", weight: .05 }, { country: "UKR", weight: .10 }];
  const state = {
    facts: [], sourceType: "demo", sourceFile: "Données de démonstration",
    grain: "function", comparison: "quarter",
    selectedFunctions: new Set(FUNCTIONS), selectedClusters: new Set(), selectedCountries: new Set()
  };
  const el = (id) => document.getElementById(id);
  const safe = (v) => String(v == null ? "" : v).replace(/[&<>'"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", "'": "&#39;", '"': "&quot;" }[c]));
  const unique = (values) => [...new Set(values.filter(Boolean))].sort((a, b) => a.localeCompare(b, "fr"));
  const sum = (rows, field = "amount") => rows.reduce((total, row) => total + (Number(row[field]) || 0), 0);
  const number = (value) => {
    if (typeof value === "number") return Number.isFinite(value) ? value : 0;
    if (value == null || value === "" || value === "-") return 0;
    const normalized = String(value).replace(/[€\s]/g, "").replace(/\((.*)\)/, "-$1").replace(/\.(?=\d{3}(\D|$))/g, "").replace(",", ".");
    const parsed = Number(normalized); return Number.isFinite(parsed) ? parsed : 0;
  };
  const yes = (value) => /^(yes|oui|true|1)$/i.test(String(value == null ? "" : value).trim());
  const text = (value, fallback = "Non renseigné") => String(value == null ? "" : value).replace(/\uFFFD/g, "").replace(/\s+/g, " ").trim() || fallback;
  const at = (row, index) => row && row[index] != null ? row[index] : "";

  function normalizeFunction(value) {
    const raw = text(value, "").toLowerCase().replace(/[^a-z&]/g, "");
    if (raw.includes("financ")) return "Finance";
    if (raw.includes("sale") || raw.includes("commercial")) return "Sales";
    if (raw.includes("market")) return "Marketing";
    if (raw === "hr" || raw.includes("humanresource") || raw.includes("ressourcehumaine")) return "HR";
    if (raw.includes("supply")) return "Supply";
    if (raw.includes("d&t") || raw === "dt" || raw.includes("digital") || raw === "it") return "D&T";
    return text(value, "Other");
  }
  function normalizeCluster(value) {
    const cleaned = text(value).replace(/[–—]/g, "-");
    if (/France\s*-\s*Long Life/i.test(cleaned)) return "France - Long Life";
    if (/France\s*-\s*(Traiteur|Fresh)/i.test(cleaned)) return "France - Traiteur";
    if (/Champiland/i.test(cleaned)) return "Champiland";
    return /^France$/i.test(cleaned) ? "France" : /Central Europe/i.test(cleaned) ? "CE" : cleaned;
  }
  function normalizeCountry(value) {
    const raw = text(value).toUpperCase();
    return ({ DE: "DEU", HU: "HUN", FRANCE: "FRA", SPAIN: "ESP", ITALY: "ITA", BELGIUM: "BEL", NETHERLANDS: "NLD", PORTUGAL: "PRT", NL: "NLD", CZ: "CZE" })[raw] || raw;
  }
  function normalizeEntity(value) {
    const match = String(value || "").match(/\d+/); return match ? String(Number(match[0])).padStart(5, "0") : "";
  }
  function normalizeToken(value) { return text(value, "").toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/[^a-z0-9]/g, ""); }
  function normalizeDate(value) {
    if (value instanceof Date && !Number.isNaN(value.getTime())) return new Date(value.getFullYear(), value.getMonth(), 1);
    if (typeof value === "number" && window.XLSX && XLSX.SSF) {
      const parsed = XLSX.SSF.parse_date_code(value); if (parsed) return new Date(parsed.y, parsed.m - 1, 1);
    }
    const french = /^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{4})$/.exec(String(value || "").trim());
    const date = french ? new Date(Number(french[3]), Number(french[2]) - 1, 1) : new Date(value);
    return Number.isNaN(date.getTime()) ? null : new Date(date.getFullYear(), date.getMonth(), 1);
  }
  function addMonths(date, count) { return new Date(date.getFullYear(), date.getMonth() + count, 1); }
  function monthKey(date) { return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}`; }
  function fiscalYear(date) { const start = date.getMonth() >= 6 ? date.getFullYear() : date.getFullYear() - 1; return `FY${String(start).slice(-2)}/${String(start + 1).slice(-2)}`; }
  function quarterLabel(date) { return `Q${Math.floor(date.getMonth() / 3) + 1}-${String(date.getFullYear()).slice(-2)}`; }
  function monthsBetween(start, end) { const months = []; for (let date = new Date(start); date <= end; date = addMonths(date, 1)) months.push(new Date(date)); return months; }
  function formatMoney(value, compact = true) {
    if (value == null || !Number.isFinite(value)) return "N/D";
    const abs = Math.abs(value); let divisor = 1, suffix = " €";
    if (compact && abs >= 1000000) { divisor = 1000000; suffix = " M€"; } else if (compact && abs >= 1000) { divisor = 1000; suffix = " k€"; }
    const decimals = divisor === 1 ? 0 : abs / divisor >= 100 ? 0 : 1;
    return `${new Intl.NumberFormat("fr-FR", { minimumFractionDigits: decimals, maximumFractionDigits: decimals }).format(value / divisor)}${suffix}`;
  }
  function formatPercent(value) { return value == null || !Number.isFinite(value) ? "N/D" : `${value > 0 ? "+" : ""}${new Intl.NumberFormat("fr-FR", { maximumFractionDigits: 1 }).format(value)} %`; }

  function categoryInfo(detail) {
    const raw = text(detail, "Other");
    if (/Staff Cost - SG&A or Indirect|SG&A or Indirect/i.test(raw)) return { family: "Personnel", category: "Staff Cost - SG&A or Indirect" };
    if (/Indus Contrib|industrial contribution|Staff Cost - Indirect/i.test(raw)) return { family: "Personnel", category: "Staff Cost - Indirect" };
    if (/SG&A.*Internal|Internal personnel.*SG&A|Staff Cost - SG&A$/i.test(raw)) return { family: "Personnel", category: "Staff Cost - SG&A" };
    if (/External Personnel/i.test(raw)) return { family: "Externes", category: "External Personnel Costs" };
    if (/Temporary workforce/i.test(raw)) return { family: "Externes", category: "External Personnel Costs" };
    if (/Consulting|Fees/i.test(raw)) return { family: "Externes", category: "Outside Consulting, Fees" };
    if (/IT Spend|IT Local|Recharge dédié|Function-Specific|IT Recharge general|General Recharge|IT Costs/i.test(raw)) return { family: "IT", category: "IT Costs" };
    return { family: "Autres", category: "Other Costs" };
  }

  function worksheetRows(workbook, name) { const sheet = workbook.Sheets[name]; return sheet ? XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: "" }) : []; }

  function entityGeographyMap(workbook) {
    const map = new Map();
    const rows = worksheetRows(workbook, "Mapping");
    rows.slice(1).forEach((row) => {
      const entity = normalizeEntity(at(row, 0)); if (!entity) return;
      const region = text(at(row, 1)); const clusterRaw = text(at(row, 2));
      const cluster = /holding/i.test(clusterRaw) ? "Holding" : /BDNE/i.test(clusterRaw) ? "BDNE" : /CE|Central/i.test(clusterRaw) ? "CE" : /Iberia/i.test(clusterRaw) ? "Iberia" : /Italy/i.test(clusterRaw) ? "Italy" : /USA|United/i.test(clusterRaw) ? "USA" : /Russia/i.test(clusterRaw) ? "Russia" : /CIS/i.test(clusterRaw) ? "CIS" : /Mercosur/i.test(clusterRaw) ? "Mercosur" : /France/i.test(clusterRaw) ? "France" : clusterRaw;
      map.set(entity, { region, cluster, country: normalizeCountry(at(row, 3)) || "Unallocated" });
    });
    return map;
  }

  function allocateBaseline(workbook) {
    const rows = worksheetRows(workbook, "Cost Baseline");
    if (!rows.length) return [];
    const geoMap = entityGeographyMap(workbook);
    const functionColumns = [[4, "D&T"], [5, "Finance"], [6, "HR"], [7, "Supply"], [8, "Marketing"], [9, "Sales"]];
    const source = [];
    for (let i = 2; i < rows.length; i += 1) {
      const row = rows[i], detail = text(at(row, 3), "");
      if (!detail || /^ETP$/i.test(detail)) continue;
      const entity = normalizeEntity(at(row, 0)); const geo = geoMap.get(entity) || { region: text(at(row, 1)), cluster: text(at(row, 1)), country: "Unallocated" };
      const category = categoryInfo(detail);
      functionColumns.forEach(([column, func]) => {
        const amount = number(at(row, column)); if (amount <= 0) return;
        source.push({ entity, function: func, region: geo.region, cluster: geo.cluster, country: geo.country, ...category, annualAmount: amount });
      });
    }
    return expandGeography(source).flatMap((item) => monthsBetween(new Date(2024, 0, 1), new Date(2029, 5, 1)).map((date) => ({
      ...item, month: monthKey(date), date, fiscalYear: fiscalYear(date), scenario: "Baseline", component: "baseline", amount: item.annualAmount / 12,
      allocationMethod: date < new Date(2025, 6, 1) ? "RETROPOLATED_BASELINE" : "STRAIGHT_LINE_FY", assumed: true
    })));
  }

  function expandGeography(rows) {
    return rows.flatMap((row) => {
      if (row.cluster === "France") return FRANCE_SPLIT.map((split) => ({ ...row, cluster: split.cluster, country: "FRA", annualAmount: row.annualAmount * split.weight, assumed: true, allocationMethod: "ESTIMATED_FRANCE_SPLIT" }));
      if (row.cluster === "CE" && row.country === "CE_MULTI") return CE_SPLIT.map((split) => ({ ...row, country: split.country, annualAmount: row.annualAmount * split.weight, assumed: true, allocationMethod: "ESTIMATED_CE_FTE_WEIGHT" }));
      return [row];
    });
  }

  function expandMonthlyGeography(rows) {
    return rows.flatMap((row) => {
      if (row.cluster === "France") return FRANCE_SPLIT.map((split) => ({ ...row, cluster: split.cluster, country: "FRA", amount: row.amount * split.weight, assumed: true, allocationMethod: "ESTIMATED_FRANCE_SPLIT" }));
      if (row.cluster === "CE" && ["CE_MULTI", "NON RENSEIGNÉ", ""].includes(row.country)) return CE_SPLIT.map((split) => ({ ...row, country: split.country, amount: row.amount * split.weight, assumed: true, allocationMethod: "ESTIMATED_CE_FTE_WEIGHT" }));
      if (row.cluster === "BDNE" && row.country === "NL+BE+LUX") return [{ ...row, country: "BEL", amount: row.amount * .35, assumed: true, allocationMethod: "ESTIMATED_BDNE_COST_WEIGHT" }, { ...row, country: "NLD", amount: row.amount * .50, assumed: true, allocationMethod: "ESTIMATED_BDNE_COST_WEIGHT" }, { ...row, country: "LUX", amount: row.amount * .15, assumed: true, allocationMethod: "ESTIMATED_BDNE_COST_WEIGHT" }];
      if (row.cluster === "BDNE" && ["NON RENSEIGNÉ", ""].includes(row.country)) return [{ ...row, country: "BEL", amount: row.amount * .20, assumed: true, allocationMethod: "ESTIMATED_BDNE_COST_WEIGHT" }, { ...row, country: "NLD", amount: row.amount * .25, assumed: true, allocationMethod: "ESTIMATED_BDNE_COST_WEIGHT" }, { ...row, country: "DEU", amount: row.amount * .35, assumed: true, allocationMethod: "ESTIMATED_BDNE_COST_WEIGHT" }, { ...row, country: "AUT", amount: row.amount * .12, assumed: true, allocationMethod: "ESTIMATED_BDNE_COST_WEIGHT" }, { ...row, country: "LUX", amount: row.amount * .08, assumed: true, allocationMethod: "ESTIMATED_BDNE_COST_WEIGHT" }];
      if (row.cluster === "Iberia" && ["NON RENSEIGNÉ", ""].includes(row.country)) return [{ ...row, country: "ESP", amount: row.amount * .78, assumed: true, allocationMethod: "ESTIMATED_IBERIA_COST_WEIGHT" }, { ...row, country: "PRT", amount: row.amount * .22, assumed: true, allocationMethod: "ESTIMATED_IBERIA_COST_WEIGHT" }];
      return [row];
    });
  }

  function prorateUnallocated(rows) {
    const clear = rows.filter((row) => row.country !== "Unallocated");
    return rows.flatMap((row) => {
      if (row.country !== "Unallocated") return [row];
      let peers = clear.filter((peer) => peer.function === row.function && peer.category === row.category && peer.cluster === row.cluster);
      if (!peers.length) peers = clear.filter((peer) => peer.function === row.function && peer.category === row.category);
      if (!peers.length) peers = clear.filter((peer) => peer.function === row.function);
      const weights = new Map(); peers.forEach((peer) => weights.set(peer.country, (weights.get(peer.country) || 0) + Math.abs(peer.annualAmount)));
      const total = [...weights.values()].reduce((a, b) => a + b, 0);
      if (!total) return [{ ...row, country: "Unallocated", assumed: true }];
      return [...weights.entries()].map(([country, weight]) => ({ ...row, country, annualAmount: row.annualAmount * weight / total, assumed: true, allocationMethod: "PRORATA_FUNCTION_CATEGORY" }));
    });
  }

  function rampWeights(start, end, fyStart, fyEnd) {
    const months = monthsBetween(fyStart, fyEnd), result = new Map();
    const rampStart = start || fyStart, rampEnd = end && end >= rampStart ? end : rampStart;
    const rampMonths = Math.max(1, (rampEnd.getFullYear() - rampStart.getFullYear()) * 12 + rampEnd.getMonth() - rampStart.getMonth() + 1);
    months.forEach((date) => {
      let weight = 0;
      if (date >= rampStart && date <= rampEnd) weight = ((date.getFullYear() - rampStart.getFullYear()) * 12 + date.getMonth() - rampStart.getMonth() + 1) / rampMonths;
      else if (date > rampEnd) weight = 1;
      result.set(monthKey(date), weight);
    });
    return result;
  }

  function spreadAnnual(amount, dates, rawWeights, allowFallback = true) {
    const totalWeight = dates.reduce((total, date) => total + (rawWeights.get(monthKey(date)) || 0), 0);
    if (!totalWeight) return allowFallback ? dates.map((date) => [date, amount / dates.length]) : [];
    return dates.map((date) => [date, amount * (rawWeights.get(monthKey(date)) || 0) / totalWeight]);
  }

  function allocateLeverComponent(base, amount, fy, component, recurring, start, end, budgetFlag, explicit = false) {
    if (!amount && !explicit) return [];
    const match = /^FY(\d{2})\/(\d{2})$/.exec(fy); if (!match) return [];
    const startYear = 2000 + Number(match[1]), fyStart = new Date(startYear, 6, 1), fyEnd = new Date(startYear + 1, 5, 1), dates = monthsBetween(fyStart, fyEnd);
    let allocations, method;
    if (recurring) {
      const outside = start && start > fyEnd;
      allocations = outside ? dates.map((date) => [date, amount / dates.length]) : spreadAnnual(amount, dates, rampWeights(start, end, fyStart, fyEnd));
      method = outside ? "OUT_OF_PERIOD_FALLBACK" : start ? "LINEAR_RAMP" : "STRAIGHT_LINE_FY";
    }
    else {
      const periodStart = start && start > fyStart ? start : fyStart, periodEnd = end && end < fyEnd ? end : fyEnd;
      const active = periodEnd >= periodStart ? monthsBetween(periodStart, periodEnd) : dates;
      allocations = active.map((date) => [date, amount / active.length]); method = periodEnd >= periodStart && start && end ? "IMPLEMENTATION_PERIOD_SPREAD" : start || end ? "OUT_OF_PERIOD_FALLBACK" : "STRAIGHT_LINE_FY";
    }
    return allocations.map(([date, monthlyAmount]) => ({ ...base, month: monthKey(date), date, fiscalYear: fy, amount: monthlyAmount, component, scenario: budgetFlag ? "Budget" : "Forecast", allocationMethod: method, assumed: method !== "DIRECT_MONTH" }));
  }

  function allocateLevers(workbook) {
    const rows = worksheetRows(workbook, "Lever BCase - Updated"), facts = [];
    const years = ["FY25/26", "FY26/27", "FY27/28", "FY28/29"];
    for (let i = 6; i < rows.length; i += 1) {
      const row = rows[i], func = normalizeFunction(at(row, 5)); if (!FUNCTIONS.includes(func)) continue;
      const start = normalizeDate(at(row, 26)), end = normalizeDate(at(row, 27)), costType = text(at(row, 20)), category = categoryInfo(costType);
      const base = { sourceRow: i + 1, leverId: text(at(row, 1), ""), leverName: text(at(row, 2), ""), implementationStart: start, implementationEnd: end, function: func, region: text(at(row, 8)), cluster: normalizeCluster(at(row, 9)), country: normalizeCountry(at(row, 10)), ...category, includedInBudget: yes(at(row, 23)) };
      years.forEach((fy, yi) => {
        const savings = number(at(row, 28 + yi));
        const recurring = number(at(row, 36 + yi)); const oneOff = number(at(row, 40 + yi)); const capex = number(at(row, 44 + yi));
        facts.push(...allocateLeverComponent(base, -savings, fy, "recurring-change", true, start, end, false));
        facts.push(...allocateLeverComponent(base, recurring, fy, "recurring-opex", true, start, end, false));
        facts.push(...allocateLeverComponent({ ...base, family: "Autres", category: "One-Off OPEX" }, oneOff, fy, "one-off", false, start, end, false));
        facts.push(...allocateLeverComponent({ ...base, family: "Autres", category: "CAPEX" }, capex, fy, "capex", false, start, end, false));
      });
    }
    return facts;
  }

  function allocateBudget(workbook, leverFacts) {
    const rows = worksheetRows(workbook, "Budget FY26-27"), facts = [];
    const leverCandidates = new Map(); leverFacts.filter((f) => f.fiscalYear === "FY26/27").forEach((fact) => {
      if (!leverCandidates.has(fact.sourceRow)) leverCandidates.set(fact.sourceRow, fact);
    });
    for (let i = 6; i < rows.length; i += 1) {
      const row = rows[i], func = normalizeFunction(at(row, 1)); if (!FUNCTIONS.includes(func) || !yes(at(row, 11))) continue;
      const category = categoryInfo(at(row, 8)), cluster = normalizeCluster(at(row, 4)), country = normalizeCountry(at(row, 5));
      const base = { sourceRow: i + 1, function: func, region: text(at(row, 2)), cluster, country, ...category, includedInBudget: true };
      const commentToken = normalizeToken(at(row, 18));
      const nameMatches = [...leverCandidates.values()].filter((fact) => {
        if (fact.function !== func || !commentToken) return false;
        const id = normalizeToken(fact.leverId), name = normalizeToken(fact.leverName);
        return (id && commentToken.includes(id)) || (name && (commentToken.includes(name) || name.includes(commentToken)));
      });
      const nameMatch = nameMatches.length === 1 ? { start: nameMatches[0].implementationStart, end: nameMatches[0].implementationEnd } : null;
      const matched = nameMatch;
      const sourceDates = matched || { start: null, end: null };
      const savings = number(at(row, 12)), recurring = number(at(row, 14)), oneOff = number(at(row, 15)), capex = number(at(row, 16));
      const explicit = true;
      facts.push(...allocateLeverComponent(base, -savings, "FY26/27", "recurring-change", true, sourceDates.start, sourceDates.end, true, explicit));
      facts.push(...allocateLeverComponent(base, recurring, "FY26/27", "recurring-opex", true, sourceDates.start, sourceDates.end, true, explicit));
      facts.push(...allocateLeverComponent({ ...base, family: "Autres", category: "One-Off OPEX" }, oneOff, "FY26/27", "one-off", false, sourceDates.start, sourceDates.end, true, explicit));
      facts.push(...allocateLeverComponent({ ...base, family: "Autres", category: "CAPEX" }, capex, "FY26/27", "capex", false, sourceDates.start, sourceDates.end, true, explicit));
    }
    return facts;
  }

  function demoFacts() {
    const facts = [], geos = [["BE", "France - Long Life", "FRA"], ["BE", "France - Traiteur", "FRA"], ["BE", "Champiland", "FRA"], ["BE", "BDNE", "BEL"], ["BE", "BDNE", "NLD"], ["BE", "BDNE", "DEU"], ["BE", "Italy", "ITA"], ["BE", "Iberia", "ESP"], ["BE", "Iberia", "PRT"], ["BE", "CE", "POL"], ["BE", "CE", "CZE"], ["BE", "CE", "HUN"], ["BE", "CE", "ROU"], ["BE", "Export", "EXPORT"], ["BEEM", "Russia", "RUS"], ["BEEM", "CIS", "KAZ"], ["BEEM", "Mercosur", "BRA"], ["BA", "USA", "USA"], ["BSA", "Holding", "HOL"]];
    const categories = [{ family: "Personnel", category: "Staff Cost - SG&A", weight: .62 }, { family: "Externes", category: "External Personnel Costs", weight: .12 }, { family: "IT", category: "IT Costs", weight: .15 }, { family: "Autres", category: "Other Costs", weight: .11 }];
    FUNCTIONS.forEach((func, fi) => geos.forEach(([region, cluster, country], gi) => categories.forEach((cat, ci) => monthsBetween(new Date(2024, 0, 1), new Date(2028, 11, 1)).forEach((date, mi) => {
        const baseline = (430000 + fi * 117000) * (.09 + gi * .013) * cat.weight;
        const change = date >= new Date(2025, 6, 1) ? -(mi - 17) * (470 + ci * 175) : 0;
      facts.push({ date, month: monthKey(date), fiscalYear: fiscalYear(date), scenario: "Baseline", component: "baseline", function: func, region, cluster, country, ...cat, amount: baseline, allocationMethod: "DEMO", assumed: true });
      facts.push({ date, month: monthKey(date), fiscalYear: fiscalYear(date), scenario: "Forecast", component: "recurring-change", function: func, region, cluster, country, ...cat, amount: change, includedInBudget: !["Finance", "Supply"].includes(func), allocationMethod: "DEMO", assumed: true });
      if (date >= new Date(2026, 6, 1) && date <= new Date(2027, 5, 1) && !["Finance", "Supply"].includes(func)) facts.push({ date, month: monthKey(date), fiscalYear: fiscalYear(date), scenario: "Budget", component: "recurring-change", function: func, region, cluster, country, ...cat, amount: change * .94, includedInBudget: true, allocationMethod: "DEMO", assumed: true });
    }))));
    return facts;
  }

  function initializeControls() {
    state.selectedFunctions = new Set(FUNCTIONS.filter((func) => state.facts.some((fact) => fact.function === func)));
    state.selectedClusters = new Set(unique(state.facts.map((fact) => fact.cluster)));
    state.selectedCountries = new Set(unique(state.facts.map((fact) => fact.country)));
    renderFilterLists();
  }

  function renderFilterList(id, values, selected, dimension) {
    const visibleSelected = values.filter((value) => selected.has(value)), isolated = visibleSelected.length === 1;
    el(id).innerHTML = values.map((value) => `<div class="check-item"><input type="checkbox" data-dimension="${dimension}" value="${safe(value)}" ${selected.has(value) ? "checked" : ""}><button class="filter-value ${isolated && visibleSelected[0] === value ? "isolated" : ""}" type="button" data-isolate="${dimension}" data-value="${safe(value)}">${safe(value)}</button></div>`).join("") || `<span class="filter-help">Aucune valeur</span>`;
  }

  function availableValues(dimension) {
    if (dimension === "functions") return FUNCTIONS.filter((func) => state.facts.some((fact) => fact.function === func));
    if (dimension === "clusters") return unique(state.facts.filter((fact) => state.selectedFunctions.has(fact.function)).map((fact) => fact.cluster));
    return unique(state.facts.filter((fact) => state.selectedFunctions.has(fact.function) && state.selectedClusters.has(fact.cluster)).map((fact) => fact.country));
  }

  function renderFilterLists() {
    renderFilterList("functionFilters", availableValues("functions"), state.selectedFunctions, "functions");
    renderFilterList("clusterFilters", availableValues("clusters"), state.selectedClusters, "clusters");
    renderFilterList("countryFilters", availableValues("countries"), state.selectedCountries, "countries");
  }

  function factsInScope() {
    return state.facts.filter((fact) => state.selectedFunctions.has(fact.function) && state.selectedClusters.has(fact.cluster) && state.selectedCountries.has(fact.country));
  }

  function dimensionValue(fact, grain = state.grain) { return grain === "function" ? fact.function : fact.cluster; }
  function grainLabel(grain) { return grain === "function" ? "Fonctions" : "Clusters"; }

  function periodWindows() {
    const today = new Date(), currentQuarterStart = new Date(today.getFullYear(), Math.floor(today.getMonth() / 3) * 3, 1);
    const currentStart = addMonths(currentQuarterStart, -12);
    return {
      current: { start: currentStart, end: currentQuarterStart, label: `${quarterLabel(currentStart)} → ${quarterLabel(addMonths(currentQuarterStart,-3))}` },
      quarter: { start: addMonths(currentStart, -3), end: addMonths(currentQuarterStart, -3), label: `${quarterLabel(addMonths(currentStart,-3))} → ${quarterLabel(addMonths(currentQuarterStart,-6))}` },
      year: { start: addMonths(currentStart, -12), end: addMonths(currentQuarterStart, -12), label: `${quarterLabel(addMonths(currentStart,-12))} → ${quarterLabel(addMonths(currentQuarterStart,-15))}` },
      budget: { start: new Date(2026, 6, 1), end: new Date(2027, 6, 1), label: "FY26/27 · juil. 2026–juin 2027" }
    };
  }

  function inWindow(fact, window) { return fact.date >= window.start && fact.date < window.end; }

  function costFacts(scope, window, scenario = "Forecast") {
    const baseline = scope.filter((fact) => fact.scenario === "Baseline" && fact.component === "baseline" && inWindow(fact, window));
    const changes = scope.filter((fact) => fact.scenario === scenario && inWindow(fact, window) && ["recurring-change", "recurring-opex"].includes(fact.component));
    return baseline.concat(changes);
  }

  function categoriesForDisplay(facts, detailed = false) {
    return detailed ? unique(facts.map((fact) => fact.category)) : unique(facts.map((fact) => fact.family));
  }

  function segmentKey(fact, detailed = false) { return detailed ? fact.category : fact.family; }

  function nonNegativeSegments(facts, detailed = false) {
    const keys = categoriesForDisplay(facts, detailed), result = {};
    keys.forEach((key) => { result[key] = Math.max(0, sum(facts.filter((fact) => segmentKey(fact, detailed) === key))); });
    return result;
  }

  function buildBudgetSegments(group, currentSegments, detailed = false) {
    const windows = periodWindows(), segments = {}, sources = new Set();
    const functions = unique(group.map((fact) => fact.function));
    functions.forEach((func) => {
      const functionGroup = group.filter((fact) => fact.function === func);
      const functionCurrent = nonNegativeSegments(costFacts(functionGroup, windows.current), detailed);
      const budgetChanges = functionGroup.filter((fact) => fact.scenario === "Budget" && inWindow(fact, windows.budget) && ["recurring-change","recurring-opex"].includes(fact.component));
      let functionBudget;
      if (budgetChanges.length) {
        const baseline = functionGroup.filter((fact) => fact.scenario === "Baseline" && fact.component === "baseline" && inWindow(fact, windows.budget));
        functionBudget = nonNegativeSegments(baseline.concat(budgetChanges), detailed); sources.add("Excel");
      } else {
        functionBudget = {};
        sources.add("Budget manquant");
      }
      Object.entries(functionBudget).forEach(([key, value]) => { segments[key] = (segments[key] || 0) + value; });
    });
    Object.keys(currentSegments).forEach((key) => { if (segments[key] == null) segments[key] = 0; });
    return { segments, source: sources.size > 1 ? [...sources].join(" + ") : sources.has("Excel") ? "Excel FY26/27" : "Budget manquant" };
  }

  function aggregateRows() {
    const windows = periodWindows(), scope = factsInScope();
    const categories = state.grain === "function" ? FUNCTIONS.filter((func) => state.selectedFunctions.has(func) && scope.some((fact) => fact.function === func)) : unique(scope.map((fact) => dimensionValue(fact)));
    return categories.map((category) => {
      const group = scope.filter((fact) => dimensionValue(fact) === category);
      const currentFacts = costFacts(group, windows.current), quarterFacts = costFacts(group, windows.quarter), yearFacts = costFacts(group, windows.year);
      const currentSegments = nonNegativeSegments(currentFacts), quarterSegments = nonNegativeSegments(quarterFacts), yearSegments = nonNegativeSegments(yearFacts);
      const budgetData = buildBudgetSegments(group, currentSegments);
      const allSegments = unique([...Object.keys(currentSegments), ...Object.keys(quarterSegments), ...Object.keys(yearSegments), ...Object.keys(budgetData.segments)]);
      const segments = {};
      allSegments.forEach((key) => {
        segments[key] = {
          current: Math.max(0, currentSegments[key] || 0),
          quarter: Math.max(0, quarterSegments[key] || 0),
          year: Math.max(0, yearSegments[key] || 0),
          budget: Math.max(0, budgetData.segments[key] || 0)
        };
      });
      const seriesTotal = (series) => Object.values(segments).reduce((total, item) => total + item[series], 0);
      return { category, current: seriesTotal("current"), quarter: seriesTotal("quarter"), year: seriesTotal("year"), budget: seriesTotal("budget"), budgetSource: budgetData.source, segments, assumed: currentFacts.concat(quarterFacts, yearFacts).some((fact) => fact.assumed) };
    });
  }

  function render() {
    const data = detailedSeries(), windows = periodWindows();
    el("dashboardSubtitle").textContent = `Période actuelle terminée en ${addMonths(windows.current.end, -1).toLocaleDateString("fr-FR", { month: "long", year: "numeric" })}`;
    const reference = state.comparison === "quarter" ? windows.quarter : windows.year;
    el("referenceHeader").textContent = reference.label;
    el("chartTitle").textContent = `12 mois actuels vs ${state.comparison === "quarter" ? "fenêtre 12 mois décalée d’un trimestre" : "même période un an plus tôt"}`;
    el("chartContext").textContent = `Actuel : ${dateRange(windows.current)} · Référence : ${dateRange(reference)} · Budget : ${dateRange(windows.budget)}`;
    renderScope(); renderChart(data); renderTable(data);
  }

  function dateRange(window) {
    const end = addMonths(window.end, -1), format = (date) => date.toLocaleDateString("fr-FR", { month: "short", year: "numeric" }).replace(".", "");
    return `${format(window.start)}–${format(end)}`;
  }

  function scopeSummary() {
    const allF=availableValues("functions"),allC=availableValues("clusters"),allP=availableValues("countries");
    const f=state.selectedFunctions.size===allF.length?"Toutes fonctions":[...state.selectedFunctions].join(", ");
    const c=state.selectedClusters.size===allC.length?"Tous clusters":[...state.selectedClusters].join(", ");
    const p=state.selectedCountries.size===allP.length?"Tous pays":[...state.selectedCountries].join(", ");
    return `${f} → ${c} → ${p}`;
  }

  function renderScope() {
    const dimensions=[
      ["1 · FONCTION",state.selectedFunctions,availableValues("functions")],
      ["2 · CLUSTER",state.selectedClusters,availableValues("clusters")],
      ["3 · PAYS",state.selectedCountries,availableValues("countries")]
    ];
    el("scopeFlow").innerHTML=dimensions.map(([label,selected,available],index)=>{
      const all=selected.size===available.length,value=all?`Tous (${available.length})`:[...selected].join(", ");
      return `${index?'<div class="scope-arrow">→</div>':''}<div class="scope-node"><span>${label}</span><strong title="${safe(value)}">${safe(value)}</strong><small>${all?"Périmètre complet":`${selected.size} sélectionné(s)`}</small></div>`;
    }).join("");
  }

  function categoryColor(key) { const index=CATEGORY_ORDER.indexOf(key); return CATEGORY_COLORS[index>=0?index:CATEGORY_COLORS.length-1]; }

  function detailedSeries() {
    const scope=factsInScope(),windows=periodWindows(),referenceKey=state.comparison;
    const current=nonNegativeSegments(costFacts(scope,windows.current),true);
    const reference=nonNegativeSegments(costFacts(scope,windows[referenceKey]),true);
    const budget=buildBudgetSegments(scope,current,true).segments;
    const categories=CATEGORY_ORDER.filter((key)=>(current[key]||0)>0||(reference[key]||0)>0||(budget[key]||0)>0);
    return { categories, current, reference, budget, windows, referenceKey };
  }

  function renderChart(data) {
    if(!data.categories.length){el("comparisonChart").innerHTML=`<div class="empty-chart">Aucune donnée pour cette sélection.</div>`;return;}
    const series=[data.reference,data.current,data.budget],totals=series.map((set)=>data.categories.reduce((total,key)=>total+(set[key]||0),0)),max=Math.max(1,...totals),width=1120,height=590,baseY=485,plotH=380,columnW=132,x=[120,494,868];
    const boundaries=series.map((set)=>{let running=0;const result={};data.categories.forEach((key)=>{const value=set[key]||0,start=running,end=running+value;result[key]={value,start,end,center:(start+end)/2};running=end;});return result;});
    const refDeltas=data.categories.map((key)=>({key,value:(data.current[key]||0)-(data.reference[key]||0)})).sort((a,b)=>Math.abs(b.value)-Math.abs(a.value));
    const budgetDeltas=data.categories.map((key)=>({key,value:(data.current[key]||0)-(data.budget[key]||0)})).sort((a,b)=>Math.abs(b.value)-Math.abs(a.value));
    const topRef=new Set(refDeltas.filter((item)=>Math.abs(item.value)>0.01).slice(0,5).map((item)=>item.key)),topBudget=new Set(budgetDeltas.filter((item)=>Math.abs(item.value)>0.01).slice(0,5).map((item)=>item.key)),maxDelta=Math.max(1,...refDeltas.concat(budgetDeltas).map((item)=>Math.abs(item.value)));
    const y=(value)=>baseY-value/max*plotH;
    let svg=`<svg viewBox="0 0 ${width} ${height}" role="img"><defs><marker id="arrow-green" markerWidth="5" markerHeight="5" refX="4.5" refY="2.5" orient="auto" markerUnits="userSpaceOnUse"><path d="M0,0 L5,2.5 L0,5 z" fill="#15805a"/></marker><marker id="arrow-red" markerWidth="5" markerHeight="5" refX="4.5" refY="2.5" orient="auto" markerUnits="userSpaceOnUse"><path d="M0,0 L5,2.5 L0,5 z" fill="#c94c52"/></marker></defs>`;
    const titles=[[state.comparison==="quarter"?"TRIMESTRE PRÉCÉDENT":"ANNÉE PRÉCÉDENTE",dateRange(data.windows[data.referenceKey])],["12 MOIS ACTUELS",dateRange(data.windows.current)],["BUDGET FY26/27",dateRange(data.windows.budget)]];
    series.forEach((set,si)=>{const total=totals[si];svg+=`<text class="flow-column-title" x="${x[si]+columnW/2}" y="26" text-anchor="middle">${titles[si][0]}</text><text class="flow-column-date" x="${x[si]+columnW/2}" y="42" text-anchor="middle">${titles[si][1]}</text><text class="flow-column-total" x="${x[si]+columnW/2}" y="64" text-anchor="middle">${formatMoney(total)}</text>`;data.categories.forEach((key)=>{const item=boundaries[si][key],yy=y(item.end),hh=Math.max(0,y(item.start)-y(item.end)),color=categoryColor(key);svg+=`<rect class="flow-segment" x="${x[si]}" y="${yy}" width="${columnW}" height="${hh}" fill="${color}" data-category="${safe(key)}" data-series="${si}" data-value="${item.value}"></rect>`;if(hh>=19)svg+=`<text class="flow-segment-value" x="${x[si]+columnW/2}" y="${yy+hh/2+3}" text-anchor="middle">${formatMoney(item.value)}</text>`;});});
    const connector=(key,leftSeries,rightSeries,leftX,rightX,delta,showLabel)=>{const left=boundaries[leftSeries][key],right=boundaries[rightSeries][key];if(!left||!right||(!left.value&&!right.value))return"";const y1=y(left.center),y2=y(right.center),color=delta>0?"#c94c52":"#15805a",stroke=.7+Math.abs(delta)/maxDelta*2.2,marker=color==="#c94c52"?"arrow-red":"arrow-green",cx=(leftX+rightX)/2,denominator=left.value,pct=denominator?delta/Math.abs(denominator)*100:null;let out=`<path class="flow-arrow" d="M${leftX},${y1} C${cx-35},${y1} ${cx+35},${y2} ${rightX},${y2}" stroke="${color}" stroke-width="${stroke}" marker-end="url(#${marker})"><title>${safe(key)} : ${formatMoney(delta,false)} (${pct==null?"N/D":formatPercent(pct)})</title></path>`;if(showLabel&&Math.abs(delta)>0.01){const arrow=delta<=0?"↓":"↑";out+=`<text class="flow-arrow-label" x="${cx}" y="${(y1+y2)/2-7}" text-anchor="middle" fill="${color}">${arrow} ${formatMoney(Math.abs(delta))} · ${pct==null?"N/D":formatPercent(pct)}</text>`;}return out;};
    data.categories.forEach((key)=>{const refDelta=(data.current[key]||0)-(data.reference[key]||0);const budgetDelta=(data.current[key]||0)-(data.budget[key]||0);svg+=connector(key,0,1,x[0]+columnW,x[1],refDelta,topRef.has(key));svg+=connector(key,2,1,x[2],x[1]+columnW,budgetDelta,topBudget.has(key));});
    svg+=`<line x1="70" x2="1050" y1="${baseY}" y2="${baseY}" stroke="#9eaaa5"/><text class="axis-label" x="62" y="${baseY+3}" text-anchor="end">0 €</text></svg>`;
    el("comparisonChart").innerHTML=svg;
    el("chartLegend").innerHTML=data.categories.map((key)=>`<div class="category-legend-item" title="${safe(key)}"><i style="background:${categoryColor(key)}"></i><span>${safe(key)}</span></div>`).join("");
    el("comparisonChart").querySelectorAll(".flow-segment").forEach((segment)=>{segment.addEventListener("mousemove",(event)=>showTooltip(event,`<strong>${safe(segment.dataset.category)}</strong><br>${titles[Number(segment.dataset.series)][0]} : ${formatMoney(Number(segment.dataset.value),false)}`));segment.addEventListener("mouseleave",hideTooltip);});
  }
  function seriesLabel(key) { const windows = periodWindows(); return ({ current: windows.current.label, quarter: windows.quarter.label, year: windows.year.label, budget: "Budget FY26/27" })[key]; }


  function renderTable(data) {
    const sorted=data.categories.map((key)=>{const reference=data.reference[key]||0,current=data.current[key]||0,budget=data.budget[key]||0;return{key,reference,current,budget,delta:current-reference,budgetDelta:current-budget};}).sort((a,b)=>Math.abs(b.delta)-Math.abs(a.delta));
    const cls=(value)=>value<=0?"value-good":"value-bad";
    el("varianceTableBody").innerHTML=sorted.map((row)=>`<tr><td><span class="category-legend-item"><i style="background:${categoryColor(row.key)}"></i><span>${safe(row.key)}</span></span></td><td class="numeric">${formatMoney(row.reference)}</td><td class="numeric">${formatMoney(row.current)}</td><td class="numeric ${cls(row.delta)}">${formatMoney(row.delta)}</td><td class="numeric ${cls(row.delta)}">${row.reference?formatPercent(row.delta/row.reference*100):"N/D"}</td><td class="numeric">${formatMoney(row.budget)}</td><td class="numeric ${cls(row.budgetDelta)}">${formatMoney(row.budgetDelta)}</td><td class="numeric ${cls(row.budgetDelta)}">${row.budget?formatPercent(row.budgetDelta/row.budget*100):"N/D"}</td></tr>`).join("")||`<tr><td colspan="8">Aucune donnée</td></tr>`;
  }

  async function importWorkbook(file) {
    showLoading(true);
    try {
      const data = await file.arrayBuffer(), workbook = XLSX.read(data, {
        type: "array", cellDates: true, cellFormula: true, cellNF: false,
        sheets: ["Cost Baseline", "Mapping", "Lever BCase - Updated", "Budget FY26-27"]
      });
      const baseline = allocateBaseline(workbook), levers = allocateLevers(workbook), budget = allocateBudget(workbook, levers);
      if (!baseline.length) throw new Error("La baseline détaillée est introuvable dans 'Cost Baseline'.");
      state.facts = baseline.concat(expandMonthlyGeography(levers), expandMonthlyGeography(budget)); state.sourceType = "monthly-model"; state.sourceFile = file.name;
      initializeControls(); updateSourceUI(); render();
      showToast(`${state.facts.length.toLocaleString("fr-FR")} lignes mensuelles générées. Baseline, forecast et budget chargés.`);
    } catch (error) { console.error(error); showToast(error.message || "Impossible de charger le classeur.", true); }
    finally { showLoading(false); el("workbookInput").value = ""; }
  }

  function updateSourceUI() {
    el("sourceName").textContent = state.sourceFile;
    el("sourceDetail").textContent = state.sourceType === "demo" ? "Modèle mensuel de démonstration" : "Baseline + Updated + Budget mensualisés";
    el("dataStatus").innerHTML = `<span class="status-dot ${state.sourceType === "demo" ? "demo" : ""}"></span><span>${state.sourceType === "demo" ? "Mode démonstration" : "Classeur chargé"}</span>`;
    el("refreshInfo").textContent = state.sourceType === "demo" ? "Données de démonstration" : `Source : ${state.sourceFile}`;
  }

  function selectOnly(dimension, value) {
    const key = `selected${dimension[0].toUpperCase()}${dimension.slice(1)}`, available = availableValues(dimension), visible = available.filter((item) => state[key].has(item));
    const restoring = visible.length === 1 && visible[0] === value;
    state[key] = restoring ? new Set(available) : new Set([value]);
    if (dimension === "functions") {
      state.selectedClusters = new Set(availableValues("clusters"));
      state.selectedCountries = new Set(availableValues("countries"));
    }
    if (dimension === "clusters") state.selectedCountries = new Set(availableValues("countries"));
    renderFilterLists(); render();
  }
  function resetAll() { state.grain = "function"; initializeControls(); render(); }

  function showLoading(visible) { el("loadingOverlay").classList.toggle("hidden", !visible); }
  let toastTimer;
  function showToast(message, error = false) { clearTimeout(toastTimer); el("toast").textContent = message; el("toast").classList.toggle("error", error); el("toast").classList.remove("hidden"); toastTimer = setTimeout(() => el("toast").classList.add("hidden"), 5500); }
  function showTooltip(event, html) { el("tooltip").innerHTML = html; el("tooltip").style.left = `${event.clientX + 12}px`; el("tooltip").style.top = `${event.clientY + 12}px`; el("tooltip").classList.remove("hidden"); }
  function hideTooltip() { el("tooltip").classList.add("hidden"); }

  function bindEvents() {
    el("workbookInput").addEventListener("change", (event) => { if (event.target.files[0]) importWorkbook(event.target.files[0]); });
    document.querySelector(".import-button").addEventListener("keydown", (event) => { if (event.key === "Enter" || event.key === " ") el("workbookInput").click(); });
    el("comparisonControl").addEventListener("click",(event)=>{const button=event.target.closest("[data-comparison]");if(!button)return;state.comparison=button.dataset.comparison;document.querySelectorAll("[data-comparison]").forEach((item)=>{const active=item===button;item.classList.toggle("active",active);item.setAttribute("aria-pressed",String(active));});render();});
    document.querySelector(".sidebar").addEventListener("change", (event) => {
      const input = event.target.closest("input[data-dimension]"); if (!input) return;
      const key = `selected${input.dataset.dimension[0].toUpperCase()}${input.dataset.dimension.slice(1)}`; if (input.checked) state[key].add(input.value); else state[key].delete(input.value);
      if (input.dataset.dimension === "functions") { state.selectedClusters = new Set(availableValues("clusters")); state.selectedCountries = new Set(availableValues("countries")); }
      else if (input.dataset.dimension === "clusters") state.selectedCountries = new Set(availableValues("countries"));
      if (input.dataset.dimension !== "countries") renderFilterLists(); render();
    });
    document.querySelector(".sidebar").addEventListener("click", (event) => {
      const isolate = event.target.closest("[data-isolate]"); if (isolate) { selectOnly(isolate.dataset.isolate, isolate.dataset.value); return; }
      const all = event.target.closest("[data-select-all]"); if (!all) return; const dimension = all.dataset.selectAll; state[`selected${dimension[0].toUpperCase()}${dimension.slice(1)}`] = new Set(availableValues(dimension));
      if(dimension==="functions"){state.selectedClusters=new Set(availableValues("clusters"));state.selectedCountries=new Set(availableValues("countries"));}else if(dimension==="clusters")state.selectedCountries=new Set(availableValues("countries"));renderFilterLists();render();
    });
    el("resetFilters").addEventListener("click", resetAll); el("scopeReset").addEventListener("click",resetAll);
  }

  state.facts = demoFacts(); bindEvents(); initializeControls(); updateSourceUI(); render();
})();
