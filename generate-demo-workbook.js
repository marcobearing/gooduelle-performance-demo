const fs = require("fs");
const XLSX = require("./xlsx.full.min.js");

const SEED = "gooduelle-traceable-demo-v2";
const FUNCTIONS = ["D&T", "Finance", "HR", "Supply", "Marketing", "Sales"];
const CATEGORIES = ["Staff Cost - Indirect", "Staff Cost - SG&A", "Staff Cost - SG&A or Indirect", "External Personnel Costs", "IT Costs", "Outside Consulting, Fees", "Other Costs"];
const FYS = ["FY25/26", "FY26/27", "FY27/28", "FY28/29"];
const OD_FYS = ["FY26/27", "FY27/28", "FY28/29"];
const OD_SCENARIOS = ["Baseline", "Budget", "Plan", "Forecast", "Actual"];
const GEOS = [
  ["50001","EU","France - Long Life","FRA"],["50002","EU","France - Traiteur","FRA"],["50003","EU","Champiland","FRA"],
  ["50004","EU","BDNE","BEL"],["50005","EU","BDNE","NLD"],["50006","EU","BDNE","DEU"],["50007","EU","BDNE","AUT"],["50008","EU","BDNE","LUX"],
  ["50009","EU","Italy","ITA"],["50010","EU","Iberia","ESP"],["50011","EU","Iberia","PRT"],
  ["50012","EU","CE","POL"],["50013","EU","CE","CZE"],["50014","EU","CE","HUN"],["50015","EU","CE","ROU"],["50016","EU","CE","SVK"],["50017","EU","CE","HRV"],["50018","EU","CE","GRC"],["50019","EU","CE","UKR"],
  ["50020","EU","Export","EXPORT"],["50021","EAST","Russia","RUS"],["50022","EAST","CIS","KAZ"],["50023","AMERICAS","Mercosur","BRA"],["50024","AMERICAS","USA","USA"],["50025","CORPORATE","Holding","HOL"]
];
const KPI_BENCHMARKS = {
  Finance:[["Taux d'externalisation Finance","%",72.5],["ETP Accounts Payable / 300k factures"," ETP",14]],
  Sales:[["Chiffre d'affaires par ETP Sales"," M€",4],["Chiffre d'affaires par KAM"," M€",30]],
  Marketing:[["Budget Marketing / CA de marque","%",7.25],["Gain de productivité Marketing","%",10]],
  HR:[["Employés par ETP HR","",61],["Coût HR par employé"," k€",1.85]],
  Supply:[["Taux de service client","%",98],["Précision des prévisions","%",80]],
  "D&T":[["Part d'effectifs externes D&T","%",60],["Économies nettes / coûts D&T","%",8.5]]
};

function hash(key){let value=2166136261;for(const char of `${SEED}|${key}`){value^=char.charCodeAt(0);value=Math.imul(value,16777619)}return value>>>0}
function unit(key){return hash(key)/4294967296}
function amount(key,min,max,step=100){return Math.round((min+unit(key)*(max-min))/step)*step}
function pct(key,min,max){return +(min+unit(key)*(max-min)).toFixed(2)}
function addMonths(date,count){return new Date(date.getFullYear(),date.getMonth()+count,1)}
function fyDates(fy){const year=2000+Number(fy.slice(2,4));return[new Date(year,6,1),new Date(year+1,5,1)]}
function fiscalYear(date){const year=date.getMonth()>=6?date.getFullYear():date.getFullYear()-1;return`FY${String(year).slice(-2)}/${String(year+1).slice(-2)}`}
function traceId(type,...parts){return`${type}-${hash(parts.join("|")).toString(16).toUpperCase().padStart(8,"0")}`}
function appendBook(workbook,name,rows){XLSX.utils.book_append_sheet(workbook,XLSX.utils.aoa_to_sheet(rows),name)}

const workbook=XLSX.utils.book_new(), trace=[];
appendBook(workbook,"Mapping",[["Legal Entity","Region","Cluster","Country","Data Status","Trace ID"],...GEOS.map(g=>[...g,"Synthetic",traceId("MAP",...g)])]);

const baseline=[["Synthetic Gooduelle baseline"],["Legal Entity","Region","Cluster","Cost detail",...FUNCTIONS,"Country","Data Status","Generation Method","Trace ID"]];
const baselineCells=[];
for(const [entity,region,cluster,country] of GEOS)for(const category of CATEGORIES){const values=FUNCTIONS.map((fn,index)=>{const f=[1.18,1.02,.75,1.40,.88,1.55][index],c={"Staff Cost - Indirect":.92,"Staff Cost - SG&A":1.28,"Staff Cost - SG&A or Indirect":.25,"External Personnel Costs":.39,"IT Costs":.61,"Outside Consulting, Fees":.22,"Other Costs":.34}[category],value=amount(`baseline|${entity}|${fn}|${category}`,48000,295000,1000)*f*c;baselineCells.push({entity,region,cluster,country,fn,category,value});return value});const id=traceId("BASE",entity,category);baseline.push([entity,region,cluster,category,...values,country,"Synthetic","Keyed deterministic model",id]);trace.push([id,"Cost Baseline",entity,country,"ALL",category,"FY25/26","Keyed deterministic model",SEED])}
appendBook(workbook,"Cost Baseline",baseline);

const levers=Array.from({length:6},(_,i)=>i===4?["","Lever ID","Lever name","","","Function","","","Region","Cluster","Country","","","","","","","","","","Cost type","","","Included in FY26/27 budget","Data Status","Generation Method","Implementation start","Implementation end"]:[]);
const leverIndex=new Map();let leverNo=1;
for(const fn of FUNCTIONS)for(const [entity,region,cluster,country] of GEOS)for(const category of CATEGORIES){if(unit(`lever-present|${fn}|${entity}|${category}`)>.34)continue;const id=`GD-L${String(leverNo++).padStart(4,"0")}`,leverTrace=traceId("LEV",fn,entity,category),row=Array(48).fill("");row[1]=id;row[2]=`Synthetic ${category.toLowerCase()} initiative`;row[5]=fn;row[8]=region;row[9]=cluster;row[10]=country;row[20]=category;row[23]="Yes";row[24]="Synthetic";row[25]="Keyed annual initiative model";row[26]=new Date(2025+Math.floor(unit(`start-y|${id}`)*2),Math.floor(unit(`start-m|${id}`)*12),1);row[27]=addMonths(row[26],6+Math.floor(unit(`duration|${id}`)*12));for(let y=0;y<4;y++){row[28+y]=amount(`saving|${id}|${y}`,8000,85000,1000)*(1+y*.28);row[36+y]=amount(`recurring|${id}|${y}`,1000,19000,500);row[40+y]=y<2?amount(`oneoff|${id}|${y}`,0,18000,500):0;row[44+y]=y<2?amount(`capex|${id}|${y}`,0,24000,500):0}levers.push(row);leverIndex.set(`${fn}|${entity}|${category}`,id);trace.push([leverTrace,"Lever BCase",entity,country,fn,category,"FY25/26-FY28/29","Keyed annual initiative model",SEED])}
appendBook(workbook,"Lever BCase - Updated",levers);

const budget=Array.from({length:6},(_,i)=>i===4?["","Function","Region","Legal Entity","Cluster","Country","Data Status","Generation Method","Cost type","Trace ID","","Included","Savings","","Recurring OPEX","One-Off OPEX","CAPEX","","Comment"]:[]);
for(const [entity,region,cluster,country] of GEOS)for(const fn of FUNCTIONS)for(const category of CATEGORIES){const id=traceId("BUD",fn,entity,category),row=Array(19).fill("");row[1]=fn;row[2]=region;row[3]=entity;row[4]=cluster;row[5]=country;row[6]="Synthetic";row[7]="Explicit country-function-category budget";row[8]=category;row[9]=id;row[11]="Yes";const base=baselineCells.find(x=>x.entity===entity&&x.fn===fn&&x.category===category).value;row[12]=Math.round(base*pct(`budget-saving-rate|${fn}|${entity}|${category}`,.025,.095)/500)*500;row[14]=Math.round(base*pct(`budget-opex-rate|${fn}|${entity}|${category}`,.002,.018)/500)*500;row[15]=amount(`budget-oneoff|${fn}|${entity}|${category}`,0,6500,500);row[16]=amount(`budget-capex|${fn}|${entity}|${category}`,0,9000,500);row[18]=leverIndex.get(`${fn}|${entity}|${category}`)||"Explicit synthetic budget";budget.push(row);trace.push([id,"Budget FY26-27",entity,country,fn,category,"FY26/27","Explicit country-function-category budget",SEED])}
appendBook(workbook,"Budget FY26-27",budget);

const monthly=[["Trace ID","Month","Fiscal Year","Scenario","Legal Entity","Region","Cluster","Country","Function","Category","Component","Amount EUR","Data Status","Generation Method"]];
for(const cell of baselineCells){const [start]=fyDates("FY25/26"),weights=Array.from({length:12},(_,m)=>1+pct(`season|${cell.fn}|${cell.category}|${m}`,-.08,.08)),sum=weights.reduce((a,b)=>a+b,0);let allocated=0;weights.forEach((weight,m)=>{const value=m===11?cell.value-allocated:Math.round(cell.value*weight/sum);allocated+=value;monthly.push([traceId("MON",cell.entity,cell.fn,cell.category,m),addMonths(start,m),"FY25/26","Baseline",cell.entity,cell.region,cell.cluster,cell.country,cell.fn,cell.category,"Baseline",value,"Synthetic","Normalized monthly seasonality"])});}
appendBook(workbook,"Monthly_Data",monthly);

const kpiCatalog=[["KPI ID","Function","KPI Name","Unit","Direction","Current","Target","Global Benchmark","Data Status","Source Label","Generation Method"]],kpiHistory=[["KPI ID","Month","Function","Cluster","Country","Value","Data Status","Generation Method"]];
for(const [fn,kpis] of Object.entries(KPI_BENCHMARKS))for(const [name,unitLabel,benchmark] of kpis){const id=traceId("KPI",fn,name),direction=/coût|ETP Accounts|externes/i.test(name)?"Lower":"Higher",current=unitLabel.includes("M€")?pct(`kpi-current|${id}`,2.1,26):unitLabel.includes("k€")?pct(`kpi-current|${id}`,1.5,2.5):pct(`kpi-current|${id}`,8,96),target=direction==="Lower"?current*.84:current*1.12;kpiCatalog.push([id,fn,name,unitLabel,direction,current,target,benchmark,"Synthetic","Illustrative benchmark","Independent KPI model"]);for(const [,region,cluster,country] of GEOS)for(let m=0;m<30;m++){const date=addMonths(new Date(2024,0,1),m),trend=(m/29-.5)*(direction==="Lower"?-1:1)*current*.08,noise=pct(`kpi-hist|${id}|${cluster}|${country}|${m}`,-current*.018,current*.018);kpiHistory.push([id,date,fn,cluster,country,Math.max(0,current+trend+noise),"Synthetic","Deterministic trend plus bounded noise"])}trace.push([id,"KPI Catalog","","",fn,name,"Current/Target","Independent KPI model",SEED])}
appendBook(workbook,"KPI_Catalog",kpiCatalog);appendBook(workbook,"KPI_History",kpiHistory);

const odDimensions=[];
for(const [entity,region,cluster,country] of GEOS)for(const fn of FUNCTIONS){
  const baseFte=12+Math.floor(unit(`od-base-fte|${entity}|${fn}`)*74);
  const annualStaffCost=amount(`od-staff-cost|${entity}|${fn}`,42000,112000,500);
  const socialRate=pct(`od-social-rate|${country}`,.17,.34);
  odDimensions.push({entity,region,cluster,country,fn,baseFte,monthlyStaffCost:annualStaffCost/12,socialRate});
}
const odMovements=[["Trace ID","Movement ID","Month","Fiscal Year","Scenario","Movement Type","Signed FTE","Legal Entity","Region","Cluster","Country","Function","Counterparty Legal Entity","Counterparty Function","Transfer Pair ID","Monthly Staff Cost Impact","Employer Social Charges Impact","Total People Cost Impact","Data Status","Generation Method"]];
const odMonthly=[["Trace ID","Month","Fiscal Year","Scenario","Legal Entity","Region","Cluster","Country","Function","Opening FTE","Hire FTE","Exit FTE","Transfer In FTE","Transfer Out FTE","Signed Movements FTE","Closing FTE","Staff Cost","Employer Social Charges","Total People Cost","Baseline Total People Cost","Monthly Savings vs Baseline","Cumulative Savings vs Baseline","Data Status","Generation Method"]];
const actualClosing=new Map(),actualCumulative=new Map();
function odMonthsForScenario(scenario){
  if(scenario==="Actual")return[0,1].map(offset=>addMonths(new Date(2026,6,1),offset));
  if(scenario==="Forecast")return Array.from({length:34},(_,offset)=>addMonths(new Date(2026,8,1),offset));
  return Array.from({length:36},(_,offset)=>addMonths(new Date(2026,6,1),offset));
}
function addOdMovement(events,scenario,month,type,index,fte,counterpartyIndex=-1,pairId=""){
  const dimension=odDimensions[index],counterparty=counterpartyIndex>=0?odDimensions[counterpartyIndex]:null;
  const sign=type==="Exit"||type==="Transfer Out"?-1:1,signedFte=sign*fte;
  const staffImpact=Math.round(signedFte*dimension.monthlyStaffCost),socialImpact=Math.round(staffImpact*dimension.socialRate),totalImpact=staffImpact+socialImpact;
  const movementId=traceId("ODM",scenario,month.toISOString(),type,index,pairId),method=type.startsWith("Transfer")?"Deterministic paired transfer":"Deterministic scenario movement";
  events.push({index,type,fte,signedFte});
  odMovements.push([movementId,movementId,month,fiscalYear(month),scenario,type,signedFte,dimension.entity,dimension.region,dimension.cluster,dimension.country,dimension.fn,counterparty?.entity||"",counterparty?.fn||"",pairId,staffImpact,socialImpact,totalImpact,"Synthetic",method]);
  trace.push([movementId,"OD_Movements",dimension.entity,dimension.country,dimension.fn,type,fiscalYear(month),method,SEED]);
}
function generateOdEvents(scenario,month,state){
  const events=[],monthKey=`${month.getFullYear()}-${month.getMonth()+1}`,scenarioRate={Actual:.055,Budget:.035,Plan:.045,Forecast:.04}[scenario]||0;
  for(let index=0;index<odDimensions.length;index++){
    if(unit(`od-event|${scenario}|${monthKey}|${index}`)>=scenarioRate)continue;
    const type=unit(`od-event-type|${scenario}|${monthKey}|${index}`)<({Actual:.46,Budget:.58,Plan:.50,Forecast:.43}[scenario]||.5)?"Hire":"Exit";
    const fte=1+(unit(`od-event-size|${scenario}|${monthKey}|${index}`)>.82?1:0);
    if(type==="Exit"&&state[index]<fte)continue;
    addOdMovement(events,scenario,month,type,index,fte);
  }
  const transferCount=scenario==="Actual"?3:scenario==="Baseline"?0:4;
  for(let transfer=0;transfer<transferCount;transfer++){
    let source=hash(`od-transfer-source|${scenario}|${monthKey}|${transfer}`)%odDimensions.length;
    let target=hash(`od-transfer-target|${scenario}|${monthKey}|${transfer}`)%odDimensions.length;
    if(target===source)target=(target+1)%odDimensions.length;
    const fte=1+(unit(`od-transfer-size|${scenario}|${monthKey}|${transfer}`)>.88?1:0);
    if(state[source]<fte)continue;
    const pairId=traceId("ODT",scenario,monthKey,transfer);
    addOdMovement(events,scenario,month,"Transfer Out",source,fte,target,pairId);
    addOdMovement(events,scenario,month,"Transfer In",target,fte,source,pairId);
  }
  return events;
}
function buildOdScenario(scenario){
  const state=odDimensions.map((dimension,index)=>scenario==="Forecast"?actualClosing.get(index):dimension.baseFte);
  const cumulative=odDimensions.map((_,index)=>scenario==="Forecast"?(actualCumulative.get(index)||0):0);
  for(const month of odMonthsForScenario(scenario)){
    const opening=[...state],events=generateOdEvents(scenario,month,state),movementTotals=odDimensions.map(()=>({Hire:0,Exit:0,"Transfer In":0,"Transfer Out":0,signed:0}));
    for(const event of events){movementTotals[event.index][event.type]+=event.fte;movementTotals[event.index].signed+=event.signedFte;state[event.index]+=event.signedFte}
    odDimensions.forEach((dimension,index)=>{
      const movements=movementTotals[index],averageFte=(opening[index]+state[index])/2,staffCost=Math.round(averageFte*dimension.monthlyStaffCost),social=Math.round(staffCost*dimension.socialRate),total=staffCost+social;
      const baselineStaff=Math.round(dimension.baseFte*dimension.monthlyStaffCost),baselineSocial=Math.round(baselineStaff*dimension.socialRate),baselineTotal=baselineStaff+baselineSocial,monthlySavings=baselineTotal-total;
      cumulative[index]+=monthlySavings;
      const id=traceId("ODMTH",scenario,month.toISOString(),dimension.entity,dimension.fn);
      odMonthly.push([id,month,fiscalYear(month),scenario,dimension.entity,dimension.region,dimension.cluster,dimension.country,dimension.fn,opening[index],movements.Hire,movements.Exit,movements["Transfer In"],movements["Transfer Out"],movements.signed,state[index],staffCost,social,total,baselineTotal,monthlySavings,cumulative[index],"Synthetic",scenario==="Baseline"?"Frozen June 2026 baseline":"Deterministic opening-movement-closing ledger"]);
      trace.push([id,"OD_Monthly",dimension.entity,dimension.country,dimension.fn,scenario,fiscalYear(month),"Deterministic OD monthly ledger",SEED]);
    });
  }
  if(scenario==="Actual")odDimensions.forEach((_,index)=>{actualClosing.set(index,state[index]);actualCumulative.set(index,cumulative[index])});
}
for(const scenario of ["Baseline","Budget","Plan","Actual","Forecast"])buildOdScenario(scenario);
appendBook(workbook,"OD_Movements",odMovements);
appendBook(workbook,"OD_Monthly",odMonthly);

appendBook(workbook,"Data Dictionary",[["Sheet","Field / Range","Definition","Status"],["Mapping","A:F","Synthetic legal entity geography reference","Synthetic"],["Cost Baseline","A:P","FY25/26 annual recurring cost baseline","Synthetic"],["Lever BCase - Updated","A:AV","Annual synthetic transformation initiatives","Synthetic"],["Budget FY26-27","A:S","Explicit budget impacts for every entity/function/category","Synthetic"],["Monthly_Data","A:N","Reconciled monthly baseline facts","Synthetic"],["KPI_Catalog","A:K","KPI definitions, current, target and global benchmark","Synthetic except illustrative benchmarks"],["KPI_History","A:H","Monthly KPI observations by geography","Synthetic"],["OD_Movements","A:T","Hire, Exit, Transfer In and Transfer Out events; signed FTE and monthly people-cost impact by organisation dimension and scenario","Synthetic"],["OD_Monthly","A:X","Monthly OD ledger from frozen June 2026 baseline, with actuals through August 2026 and forecast thereafter; opening, movement, closing FTE, people costs and savings","Synthetic"],["Generation Log","A:I","Traceability metadata for generated facts, including OD movement and monthly records","Synthetic"]]);
appendBook(workbook,"Generation Log",[["Trace ID","Dataset","Legal Entity","Country","Function","Category / KPI","Period","Generation Method","Seed"],...trace]);

workbook.Props={Title:"Gooduelle Performance Demo - Traceable Synthetic Data",Subject:"Fully synthetic and traceable demonstration dataset",Author:"Gooduelle Demo Generator",Company:"Gooduelle",Comments:`Generated deterministically with seed ${SEED}. No operational source data used.`};
fs.writeFileSync("Gooduelle_Performance_Demo.xlsx",XLSX.write(workbook,{type:"buffer",bookType:"xlsx",compression:true}));
