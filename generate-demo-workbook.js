const fs = require("fs");
const XLSX = require("./xlsx.full.min.js");

const FUNCTIONS = ["D&T", "Finance", "HR", "Supply", "Marketing", "Sales"];
const CATEGORIES = [
  "Staff Cost - Indirect",
  "Staff Cost - SG&A",
  "Staff Cost - SG&A or Indirect",
  "External Personnel Costs",
  "IT Costs",
  "Outside Consulting, Fees",
  "Other Costs"
];
const GEOGRAPHIES = [
  ["50001", "EU", "France - Long Life", "FRA"], ["50002", "EU", "France - Traiteur", "FRA"], ["50003", "EU", "Champiland", "FRA"],
  ["50004", "EU", "BDNE", "BEL"], ["50005", "EU", "BDNE", "NLD"], ["50006", "EU", "BDNE", "DEU"], ["50007", "EU", "BDNE", "AUT"], ["50008", "EU", "BDNE", "LUX"],
  ["50009", "EU", "Italy", "ITA"], ["50010", "EU", "Iberia", "ESP"], ["50011", "EU", "Iberia", "PRT"],
  ["50012", "EU", "CE", "POL"], ["50013", "EU", "CE", "CZE"], ["50014", "EU", "CE", "HUN"], ["50015", "EU", "CE", "ROU"], ["50016", "EU", "CE", "SVK"], ["50017", "EU", "CE", "HRV"], ["50018", "EU", "CE", "GRC"], ["50019", "EU", "CE", "UKR"],
  ["50020", "EU", "Export", "EXPORT"], ["50021", "EAST", "Russia", "RUS"], ["50022", "EAST", "CIS", "KAZ"], ["50023", "AMERICAS", "Mercosur", "BRA"], ["50024", "AMERICAS", "USA", "USA"], ["50025", "CORPORATE", "Holding", "HOL"]
];

function hash(key) {
  let value = 2166136261;
  for (const char of `gooduelle-performance-demo-v1|${key}`) {
    value ^= char.charCodeAt(0);
    value = Math.imul(value, 16777619);
  }
  return value >>> 0;
}
function unit(key) { return hash(key) / 4294967296; }
function amount(key, min, max, step = 100) { return Math.round((min + unit(key) * (max - min)) / step) * step; }

const workbook = XLSX.utils.book_new();

const mapping = [["Legal Entity", "Region", "Cluster", "Country"], ...GEOGRAPHIES];
XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(mapping), "Mapping");

const baseline = [["Synthetic Gooduelle baseline"], ["Legal Entity", "Region", "Unused", "Cost detail", ...FUNCTIONS]];
for (const [entity, region, cluster, country] of GEOGRAPHIES) {
  for (const category of CATEGORIES) {
    const row = [entity, region, cluster, category];
    FUNCTIONS.forEach((fn, index) => {
      const functionScale = [1.25, 1.05, .72, 1.45, .86, 1.65][index];
      const categoryScale = { "Staff Cost - Indirect": 1.0, "Staff Cost - SG&A": 1.35, "Staff Cost - SG&A or Indirect": .28, "External Personnel Costs": .42, "IT Costs": .55, "Outside Consulting, Fees": .24, "Other Costs": .32 }[category];
      row.push(amount(`${entity}|${category}|${fn}`, 55000, 310000, 1000) * functionScale * categoryScale);
    });
    baseline.push(row);
  }
}
XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(baseline), "Cost Baseline");

const leverRows = Array.from({ length: 6 }, (_, index) => index === 4 ? ["", "Lever ID", "Lever name", "", "", "Function", "", "", "Region", "Cluster", "Country", "", "", "", "", "", "", "", "", "", "Cost type", "", "", "Included in FY26/27 budget", "", "", "Implementation start", "Implementation end"] : []);
let leverNumber = 1;
for (const fn of ["Finance", "Sales", "Marketing", "HR", "Supply", "D&T"]) {
  const geos = GEOGRAPHIES.filter((_, index) => index % 6 === leverNumber % 6).slice(0, 3);
  for (const [, region, cluster, country] of geos) {
    const category = CATEGORIES[(leverNumber + FUNCTIONS.indexOf(fn)) % CATEGORIES.length];
    const row = Array(48).fill("");
    row[1] = `GD-L${String(leverNumber).padStart(3, "0")}`;
    row[2] = `Synthetic efficiency initiative ${leverNumber}`;
    row[5] = fn; row[8] = region; row[9] = cluster; row[10] = country; row[20] = category; row[23] = "Yes";
    row[26] = new Date(2026, (leverNumber * 2) % 12, 1); row[27] = new Date(2026 + (leverNumber % 2), (leverNumber * 2 + 5) % 12, 1);
    for (let year = 0; year < 4; year++) {
      row[28 + year] = amount(`saving|${leverNumber}|${year}`, 25000, 240000, 1000) * (1 + year * .35);
      row[36 + year] = amount(`opex|${leverNumber}|${year}`, 2000, 42000, 500);
      row[40 + year] = year < 2 ? amount(`oneoff|${leverNumber}|${year}`, 0, 55000, 500) : 0;
      row[44 + year] = year < 2 ? amount(`capex|${leverNumber}|${year}`, 0, 70000, 500) : 0;
    }
    leverRows.push(row); leverNumber++;
  }
}
XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(leverRows), "Lever BCase - Updated");

const budgetRows = Array.from({ length: 6 }, (_, index) => index === 4 ? ["", "Function", "Region", "", "Cluster", "Country", "", "", "Cost type", "", "", "Included", "Savings", "", "Recurring OPEX", "One-Off OPEX", "CAPEX", "", "Comment"] : []);
let budgetIndex = 0;
for (const fn of ["Finance", "Sales", "Marketing", "HR", "Supply", "D&T"]) {
  const [, region, cluster, country] = GEOGRAPHIES[(budgetIndex * 4 + 3) % GEOGRAPHIES.length];
  const row = Array(19).fill("");
  row[1] = fn; row[2] = region; row[4] = cluster; row[5] = country; row[8] = CATEGORIES[(budgetIndex + 1) % CATEGORIES.length]; row[11] = "Yes";
  row[12] = amount(`budget-saving|${fn}`, 64000, 287000, 1000); row[14] = amount(`budget-opex|${fn}`, 6500, 48500, 500); row[15] = amount(`budget-oneoff|${fn}`, 0, 33500, 500); row[16] = amount(`budget-capex|${fn}`, 0, 43500, 500); row[18] = `GD-L${String(budgetIndex * 3 + 1).padStart(3, "0")}`;
  budgetRows.push(row); budgetIndex++;
}
XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(budgetRows), "Budget FY26-27");

workbook.Props = { Title: "Gooduelle Performance Demo", Subject: "Synthetic demonstration data", Author: "Demo Generator", Company: "Gooduelle", Comments: "All operational values are fictional." };
const output = XLSX.write(workbook, { type: "buffer", bookType: "xlsx", compression: true });
fs.writeFileSync("Gooduelle_Performance_Demo.xlsx", output);
