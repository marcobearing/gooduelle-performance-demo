const fs=require("fs"),XLSX=require("./xlsx.full.min.js"),w=XLSX.read(fs.readFileSync("Gooduelle_Performance_Demo.xlsx"),{type:"buffer",cellDates:true}),rows=n=>XLSX.utils.sheet_to_json(w.Sheets[n],{header:1,raw:true,defval:""});
const required=["Mapping","Cost Baseline","Lever BCase - Updated","Budget FY26-27","Monthly_Data","KPI_Catalog","KPI_History","OD_Movements","OD_Monthly","Data Dictionary","Generation Log"];for(const n of required)if(!w.Sheets[n])throw Error(`Missing ${n}`);
const mapping=rows("Mapping"),baseline=rows("Cost Baseline"),budget=rows("Budget FY26-27"),monthly=rows("Monthly_Data"),catalog=rows("KPI_Catalog"),history=rows("KPI_History"),odMovements=rows("OD_Movements"),odMonthly=rows("OD_Monthly"),log=rows("Generation Log");
const functions=["D&T","Finance","HR","Supply","Marketing","Sales"],entities=mapping.slice(1).map(r=>r[0]),categories=["Staff Cost - Indirect","Staff Cost - SG&A","Staff Cost - SG&A or Indirect","External Personnel Costs","IT Costs","Outside Consulting, Fees","Other Costs"];
const expected=entities.length*functions.length*categories.length;if(budget.length-6!==expected)throw Error(`Budget coverage ${budget.length-6}/${expected}`);for(const entity of entities)for(const fn of functions)for(const cat of categories)if(!budget.slice(6).some(r=>r[3]===entity&&r[1]===fn&&r[8]===cat))throw Error(`Missing budget ${entity}/${fn}/${cat}`);
for(const r of baseline.slice(2))for(const v of r.slice(4,10))if(Number(v)<0)throw Error("Negative baseline");
const monthlyTotals=new Map();for(const r of monthly.slice(1)){const key=[r[4],r[8],r[9]].join("|");monthlyTotals.set(key,(monthlyTotals.get(key)||0)+Number(r[11]))}for(const r of baseline.slice(2)){functions.forEach((fn,i)=>{const key=[r[0],fn,r[3]].join("|");if(Math.abs((monthlyTotals.get(key)||0)-Number(r[4+i]))>.01)throw Error(`Monthly reconciliation ${key}`)})}
if(catalog.length<13||history.length<500||log.length<100)throw Error("Traceability datasets too small");
const odFiscalYears=new Set(["FY26/27","FY27/28","FY28/29"]),odScenarios=new Set(["Baseline","Budget","Plan","Forecast","Actual"]),transferBalance=new Map(),transferPairs=new Map();
for(const r of odMovements.slice(1)){
  if(!odFiscalYears.has(r[3]))throw Error(`Unexpected OD movement fiscal year ${r[3]}`);
  if(!odScenarios.has(r[4]))throw Error(`Unexpected OD movement scenario ${r[4]}`);
  if(r[4]==="Actual"&&new Date(r[2])>new Date(2026,7,31))throw Error("Actual OD movement after August 2026");
  if(r[5]==="Transfer In"||r[5]==="Transfer Out"){
    const key=[r[4],new Date(r[2]).toISOString().slice(0,7)].join("|");transferBalance.set(key,(transferBalance.get(key)||0)+Number(r[6]));
    const pair=r[14];if(!pair)throw Error("Transfer missing pair ID");transferPairs.set(pair,(transferPairs.get(pair)||0)+Number(r[6]));
  }
  if(Math.abs(Number(r[15])+Number(r[16])-Number(r[17]))>.01)throw Error(`OD movement cost reconciliation ${r[1]}`);
}
for(const [key,value] of transferBalance)if(Math.abs(value)>.0001)throw Error(`Transfer group imbalance ${key}: ${value}`);
for(const [key,value] of transferPairs)if(Math.abs(value)>.0001)throw Error(`Transfer pair imbalance ${key}: ${value}`);
const monthlyCoverage=new Set(),monthlyKeys=new Set(),odLedger=new Map(),odCumulative=new Map(),movementLedger=new Map();
for(const r of odMovements.slice(1)){const key=[r[4],new Date(r[2]).toISOString().slice(0,7),r[7],r[11]].join("|");const item=movementLedger.get(key)||{hire:0,exit:0,transferIn:0,transferOut:0};if(r[5]==="Hire")item.hire+=Math.abs(Number(r[6]));if(r[5]==="Exit")item.exit+=Math.abs(Number(r[6]));if(r[5]==="Transfer In")item.transferIn+=Math.abs(Number(r[6]));if(r[5]==="Transfer Out")item.transferOut+=Math.abs(Number(r[6]));movementLedger.set(key,item)}
for(const r of odMonthly.slice(1)){
  const date=new Date(r[1]),fy=r[2],scenario=r[3],key=[scenario,date.toISOString().slice(0,7),r[4],r[8]].join("|");
  if(monthlyKeys.has(key))throw Error(`Duplicate OD monthly row ${key}`);monthlyKeys.add(key);
  if(!odFiscalYears.has(fy))throw Error(`Unexpected OD monthly fiscal year ${fy}`);monthlyCoverage.add(fy);
  if(!odScenarios.has(scenario))throw Error(`Unexpected OD monthly scenario ${scenario}`);
  if(scenario==="Actual"&&date>new Date(2026,7,31))throw Error("Actual OD monthly row after August 2026");
  const signed=Number(r[10])-Number(r[11])+Number(r[12])-Number(r[13]);
  if(Math.abs(signed-Number(r[14]))>.0001||Math.abs(Number(r[9])+signed-Number(r[15]))>.0001)throw Error(`OD FTE reconciliation ${key}`);
  if(Math.abs(Number(r[16])+Number(r[17])-Number(r[18]))>.01)throw Error(`OD monthly cost reconciliation ${key}`);
  if(Math.abs(Number(r[19])-Number(r[18])-Number(r[20]))>.01)throw Error(`OD monthly savings reconciliation ${key}`);
  const chainKey=[scenario,r[4],r[8]].join("|"),previous=odLedger.get(chainKey);if(previous&&date>previous.date&&Math.abs(Number(r[9])-previous.closing)>.0001)throw Error(`OD opening/closing continuity ${key}`);odLedger.set(chainKey,{date,closing:Number(r[15])});
  const previousCumulative=odCumulative.get(chainKey);if(previousCumulative!=null&&Math.abs(previousCumulative+Number(r[20])-Number(r[21]))>.01)throw Error(`OD cumulative savings reconciliation ${key}`);odCumulative.set(chainKey,Number(r[21]));
  if(scenario!=="Baseline"){const movements=movementLedger.get(key)||{hire:0,exit:0,transferIn:0,transferOut:0};if(Math.abs(movements.hire-Number(r[10]))>.0001||Math.abs(movements.exit-Number(r[11]))>.0001||Math.abs(movements.transferIn-Number(r[12]))>.0001||Math.abs(movements.transferOut-Number(r[13]))>.0001)throw Error(`OD movement-to-ledger reconciliation ${key}`)}
}
for(const entity of entities)for(const fn of functions){const actual=odMonthly.slice(1).filter(r=>r[3]==="Actual"&&r[4]===entity&&r[8]===fn).sort((a,b)=>new Date(a[1])-new Date(b[1])),forecast=odMonthly.slice(1).filter(r=>r[3]==="Forecast"&&r[4]===entity&&r[8]===fn).sort((a,b)=>new Date(a[1])-new Date(b[1]));if(actual.length&&forecast.length){if(Math.abs(Number(actual.at(-1)[15])-Number(forecast[0][9]))>.0001)throw Error(`Actual/Forecast OD continuity ${entity}/${fn}`);if(Math.abs(Number(actual.at(-1)[21])+Number(forecast[0][20])-Number(forecast[0][21]))>.01)throw Error(`Actual/Forecast cumulative savings continuity ${entity}/${fn}`)}}
for(const fy of odFiscalYears)if(!monthlyCoverage.has(fy))throw Error(`Missing OD fiscal year ${fy}`);
for(const scenario of ["Baseline","Budget","Plan"])for(const fy of odFiscalYears)if(!odMonthly.slice(1).some(r=>r[2]===fy&&r[3]===scenario))throw Error(`Missing OD coverage ${scenario}/${fy}`);
if(!odMonthly.slice(1).some(r=>r[3]==="Actual"&&new Date(r[1]).getMonth()===7)||!odMonthly.slice(1).some(r=>r[3]==="Forecast"&&r[2]==="FY28/29"))throw Error("Incomplete Actual/Forecast OD coverage");
const text=JSON.stringify({mapping,baseline,budget,monthly,catalog,history,odMovements,odMonthly,log});for(const term of ["Bonduelle","F4P","GELT","Control Tower","Delivery Board","Marco","Guillaume"])if(text.toLowerCase().includes(term.toLowerCase()))throw Error(`Forbidden ${term}`);
console.log(JSON.stringify({sheets:w.SheetNames,budgetRows:budget.length-6,expectedBudgetRows:expected,monthlyRows:monthly.length-1,kpiRows:catalog.length-1,kpiHistoryRows:history.length-1,odMovementRows:odMovements.length-1,odMonthlyRows:odMonthly.length-1,odFiscalYears:[...monthlyCoverage],traceRows:log.length-1}));
