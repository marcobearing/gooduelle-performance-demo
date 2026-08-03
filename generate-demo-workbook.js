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

// Program scope: 6 support functions with fixed June-2026 baseline FTE totals, plus a virtual "Operations"
// function that carries the 8450 FTE out of scope. Country/cluster split is preserved proportionally.
const OD_INSCOPE_FUNCTIONS=FUNCTIONS;
const OD_OUTSCOPE_FUNCTION="Operations";
const OD_FUNCTIONS_ALL=[...OD_INSCOPE_FUNCTIONS,OD_OUTSCOPE_FUNCTION];
const OD_BASELINE_TARGETS={"D&T":200,Finance:300,HR:200,Supply:350,Marketing:100,Sales:400,[OD_OUTSCOPE_FUNCTION]:8450};
const OD_INSCOPE_TOTAL=OD_INSCOPE_FUNCTIONS.reduce((sum,fn)=>sum+OD_BASELINE_TARGETS[fn],0); // 1550
const OD_OUTSCOPE_TOTAL=OD_BASELINE_TARGETS[OD_OUTSCOPE_FUNCTION]; // 8450
const OD_GROUP_TOTAL=OD_INSCOPE_TOTAL+OD_OUTSCOPE_TOTAL; // 10000
const odDimensions=[];
for(const [entity,region,cluster,country] of GEOS)for(const fn of OD_FUNCTIONS_ALL){
  const rawWeight=12+Math.floor(unit(`od-base-fte|${entity}|${fn}`)*74);
  const annualStaffCost=amount(`od-staff-cost|${entity}|${fn}`,42000,112000,500);
  const socialRate=pct(`od-social-rate|${country}`,.17,.34);
  odDimensions.push({entity,region,cluster,country,fn,rawWeight,baseFte:0,monthlyStaffCost:annualStaffCost/12,socialRate,scope:OD_INSCOPE_FUNCTIONS.includes(fn)?"In-scope":"Out-of-scope"});
}
const odFunctionIndexes=new Map(OD_FUNCTIONS_ALL.map(fn=>[fn,odDimensions.map((d,index)=>d.fn===fn?index:-1).filter(index=>index>=0)]));
// Rescale raw weights to hit exact per-function baseline totals while preserving proportional split
for(const fn of OD_FUNCTIONS_ALL){
  const indexes=odFunctionIndexes.get(fn),weights=indexes.map(index=>odDimensions[index].rawWeight),allocated=allocateInteger(OD_BASELINE_TARGETS[fn],weights);
  indexes.forEach((index,localIndex)=>{odDimensions[index].baseFte=allocated[localIndex]});
}
const odBaselineFte=Object.fromEntries(OD_FUNCTIONS_ALL.map(fn=>[fn,odFunctionIndexes.get(fn).reduce((sum,index)=>sum+odDimensions[index].baseFte,0)]));
const odBaselineStaff=Object.fromEntries(OD_FUNCTIONS_ALL.map(fn=>[fn,odFunctionIndexes.get(fn).reduce((sum,index)=>sum+Math.round(odDimensions[index].baseFte*odDimensions[index].monthlyStaffCost),0)]));
const odBaselineGroupInScope=OD_INSCOPE_FUNCTIONS.reduce((sum,fn)=>sum+odBaselineFte[fn],0);
const odBaselineStaffGroupInScope=OD_INSCOPE_FUNCTIONS.reduce((sum,fn)=>sum+odBaselineStaff[fn],0);
// Per-function FTE ratios: Finance -50%, Sales/Marketing -2.5%, HR -5%, D&T +2%.
// Supply is derived so that total in-scope monthly staff cost lands at -10% at FY28/29.
// D&T also carries -10% per-FTE recurring staff cost (nearshoring).
const OD_FTE_RATIO_TARGETS={"D&T":1.02,Finance:.50,HR:.95,Marketing:.975,Sales:.975};
const OD_STAFF_COST_PER_FTE_TARGETS={"D&T":.90,Finance:1.00,HR:1.00,Marketing:1.00,Sales:1.00,Supply:1.00};
// Compute Supply's final FTE so that Sum(final staff cost) = 0.90 * Sum(baseline staff cost) over the 6 in-scope functions.
const odTargetGroupStaff=Math.round(odBaselineStaffGroupInScope*.90);
const odFixedFinalStaff=OD_INSCOPE_FUNCTIONS.filter(fn=>fn!=="Supply").reduce((sum,fn)=>sum+odBaselineStaff[fn]*OD_FTE_RATIO_TARGETS[fn]*OD_STAFF_COST_PER_FTE_TARGETS[fn],0);
const odSupplyFinalStaff=odTargetGroupStaff-odFixedFinalStaff;
const odSupplyFinalFteRatio=odSupplyFinalStaff/odBaselineStaff.Supply; // Supply per-FTE cost unchanged, so ratio equals FTE ratio
OD_FTE_RATIO_TARGETS.Supply=odSupplyFinalFteRatio;
const odPlanFinal=Object.fromEntries(OD_INSCOPE_FUNCTIONS.map(fn=>[fn,Math.round(odBaselineFte[fn]*OD_FTE_RATIO_TARGETS[fn])]));
odPlanFinal[OD_OUTSCOPE_FUNCTION]=odBaselineFte[OD_OUTSCOPE_FUNCTION]; // Out-of-scope FTE unchanged
const odScenarioFinal={
  Plan:odPlanFinal,
  Budget:Object.fromEntries(OD_FUNCTIONS_ALL.map(fn=>[fn,fn===OD_OUTSCOPE_FUNCTION?odBaselineFte[fn]:odBaselineFte[fn]+Math.round((odPlanFinal[fn]-odBaselineFte[fn])*(fn==="D&T"?.85:.93))])),
  Forecast:Object.fromEntries(OD_FUNCTIONS_ALL.map(fn=>[fn,fn===OD_OUTSCOPE_FUNCTION?odBaselineFte[fn]:odBaselineFte[fn]+Math.round((odPlanFinal[fn]-odBaselineFte[fn])*({"D&T":.93,Finance:.98,HR:.90,Supply:1.01,Marketing:.90,Sales:1.05}[fn]))]))
};
const odMovements=[["Trace ID","Movement ID","Month","Fiscal Year","Scenario","Movement Type","Exit Subtype","Signed FTE","Legal Entity","Region","Cluster","Country","Function","Program Scope","Counterparty Legal Entity","Counterparty Function","Transfer Pair ID","Recurring Staff Cost Impact","Recurring Employer Social Charges Impact","Transformation Social Cost (ENR)","Recurring Total People Cost Impact","Total Cash Cost Impact Including ENR","Data Status","Generation Method"]];
const odMonthly=[["Trace ID","Month","Fiscal Year","Scenario","Legal Entity","Region","Cluster","Country","Function","Program Scope","Opening FTE","Hire FTE","Natural Attrition FTE","Forced Exit FTE","Total Exit FTE","Transfer In FTE","Transfer Out FTE","Signed Movements FTE","Closing FTE","Recurring Staff Cost","Recurring Employer Social Charges","Transformation Social Cost (ENR)","Recurring Total People Cost","Total Cash Cost Including ENR","Baseline Recurring Total People Cost","Recurring Savings vs Baseline","Net Savings After ENR","Cumulative Recurring Savings","Cumulative Net Savings","Data Status","Generation Method"]];
const actualClosing=new Map(),actualRecurringCumulative=new Map(),actualNetCumulative=new Map();
function allocateInteger(total,weights){
  if(total<=0)return weights.map(()=>0);
  const weightTotal=weights.reduce((sum,value)=>sum+Math.max(0,value),0);
  if(!weightTotal){const result=weights.map(()=>0);result[0]=total;return result}
  const exact=weights.map(value=>total*Math.max(0,value)/weightTotal),result=exact.map(Math.floor);
  let remainder=total-result.reduce((sum,value)=>sum+value,0);
  exact.map((value,index)=>({index,fraction:value-result[index]})).sort((a,b)=>b.fraction-a.fraction||a.index-b.index).slice(0,remainder).forEach(item=>result[item.index]++);
  return result;
}
function finalVector(start,targetTotal){
  const delta=targetTotal-start.reduce((sum,value)=>sum+value,0),changes=allocateInteger(Math.abs(delta),start);
  return start.map((value,index)=>value+(delta<0?-changes[index]:changes[index]));
}
function trajectoryVector(start,final,targetTotal,key){
  const startTotal=start.reduce((sum,value)=>sum+value,0),delta=targetTotal-startTotal;
  if(!delta)return [...start];
  const capacities=final.map((value,index)=>Math.abs(value-start[index])),schedule=[];
  for(let level=0;level<Math.max(...capacities);level++)capacities.map((capacity,index)=>({capacity,index})).filter(item=>item.capacity>level).sort((a,b)=>hash(`${key}|${level}|${a.index}`)-hash(`${key}|${level}|${b.index}`)).forEach(item=>schedule.push(item.index));
  const changes=start.map(()=>0);schedule.slice(0,Math.abs(delta)).forEach(index=>changes[index]++);
  return start.map((value,index)=>value+(delta<0?-changes[index]:changes[index]));
}
function scenarioMonths(scenario){
  if(scenario==="Baseline")return Array.from({length:37},(_,offset)=>addMonths(new Date(2026,5,1),offset));
  if(scenario==="Actual")return[0,1].map(offset=>addMonths(new Date(2026,6,1),offset));
  if(scenario==="Forecast")return Array.from({length:34},(_,offset)=>addMonths(new Date(2026,8,1),offset));
  return Array.from({length:36},(_,offset)=>addMonths(new Date(2026,6,1),offset));
}
function rampProgress(scenario,index){
  if(scenario==="Baseline")return 0;
  if(scenario==="Actual")return[index===0?.015:.03][0];
  const segmentLengths=scenario==="Forecast"?[10,12,12]:[12,12,12],milestones=[.30,.70,1],monthNumber=index+1;
  let elapsed=0,previous=0;
  for(let segment=0;segment<segmentLengths.length;segment++){
    const length=segmentLengths[segment],target=milestones[segment];
    if(monthNumber<=elapsed+length)return previous+(target-previous)*(monthNumber-elapsed)/length;
    elapsed+=length;previous=target;
  }
  return 1;
}
function aggregateTarget(scenario,fn,index,startTotal){
  if(scenario==="Baseline")return startTotal;
  if(scenario==="Actual")return odBaselineFte[fn]+Math.round((odPlanFinal[fn]-odBaselineFte[fn])*rampProgress(scenario,index));
  const final=odScenarioFinal[scenario][fn],progress=rampProgress(scenario,index);
  return startTotal+Math.round((final-startTotal)*progress);
}
function addOdMovement(events,scenario,month,type,index,fte,counterpartyIndex=-1,pairId=""){
  if(!fte)return;
  const dimension=odDimensions[index],counterparty=counterpartyIndex>=0?odDimensions[counterpartyIndex]:null;
  const sign=type==="Natural Attrition"||type==="Forced Exit"||type==="Transfer Out"?-1:1,signedFte=sign*fte;
  const staffImpact=Math.round(signedFte*dimension.monthlyStaffCost),socialImpact=Math.round(staffImpact*dimension.socialRate),recurringImpact=staffImpact+socialImpact;
  const enr=type==="Forced Exit"?Math.round(Math.abs(recurringImpact)*9*(.90+unit(`od-enr-country|${dimension.country}`)*.20)):0,totalCashImpact=recurringImpact+enr;
  const movementId=traceId("ODM",scenario,month.toISOString(),type,index,pairId,fte),exitSubtype=type==="Natural Attrition"||type==="Forced Exit"?type:"";
  const method=type.startsWith("Transfer")?"Paired D&T nearshore transfer":type==="Forced Exit"?"Target-driven forced exit with nine-month employer-cost ENR":"Target-driven deterministic movement";
  events.push({index,type,fte,signedFte,enr});
  odMovements.push([movementId,movementId,month,fiscalYear(month),scenario,type,exitSubtype,signedFte,dimension.entity,dimension.region,dimension.cluster,dimension.country,dimension.fn,dimension.scope,counterparty?.entity||"",counterparty?.fn||"",pairId,staffImpact,socialImpact,enr,recurringImpact,totalCashImpact,"Synthetic",method]);
  trace.push([movementId,"OD_Movements",dimension.entity,dimension.country,dimension.fn,type,fiscalYear(month),method,SEED]);
}
// Staff-cost per-FTE ratio for scenario/function at a given monthly progress in [0,1].
// In-scope functions ramp from 1.00 to OD_STAFF_COST_PER_FTE_TARGETS[fn]. Out-of-scope keeps 1.00 always.
function costPerFteRatio(scenario,fn,progress){
  if(fn===OD_OUTSCOPE_FUNCTION)return 1;
  const finalRatio=OD_STAFF_COST_PER_FTE_TARGETS[fn]??1;
  if(scenario==="Baseline")return 1;
  if(scenario==="Budget")return 1+(finalRatio-1)*progress*.85;
  return 1+(finalRatio-1)*progress;
}
function buildOdScenario(scenario){
  const startState=odDimensions.map((dimension,index)=>scenario==="Forecast"?actualClosing.get(index):dimension.baseFte),state=[...startState],organicState=[...startState],transferOffset=odDimensions.map(()=>0);
  const recurringCumulative=odDimensions.map((_,index)=>scenario==="Forecast"?(actualRecurringCumulative.get(index)||0):0),netCumulative=odDimensions.map((_,index)=>scenario==="Forecast"?(actualNetCumulative.get(index)||0):0);
  const startsByFunction={},finalsByFunction={};
  for(const fn of OD_FUNCTIONS_ALL){
    const indexes=odFunctionIndexes.get(fn),start=indexes.map(index=>startState[index]),startTotal=start.reduce((sum,value)=>sum+value,0);
    let finalTotal;
    if(fn===OD_OUTSCOPE_FUNCTION)finalTotal=startTotal;
    else if(scenario==="Actual")finalTotal=odBaselineFte[fn]+Math.round((odPlanFinal[fn]-odBaselineFte[fn])*.03);
    else if(scenario==="Baseline")finalTotal=startTotal;
    else finalTotal=odScenarioFinal[scenario][fn];
    startsByFunction[fn]=start;finalsByFunction[fn]=finalVector(start,finalTotal);
  }
  const forcedRunning=Object.fromEntries(OD_INSCOPE_FUNCTIONS.map(fn=>[fn,{total:0,forced:0}]));
  const dtIndexes=odFunctionIndexes.get("D&T"),dtSource=[...dtIndexes].sort((a,b)=>odDimensions[b].monthlyStaffCost-odDimensions[a].monthlyStaffCost)[0],dtTarget=[...dtIndexes].sort((a,b)=>odDimensions[a].monthlyStaffCost-odDimensions[b].monthlyStaffCost)[0];
  for(const [monthIndex,month] of scenarioMonths(scenario).entries()){
    const opening=[...state],events=[];
    // Only in-scope functions generate movements. Out-of-scope Operations stays flat.
    for(const fn of OD_INSCOPE_FUNCTIONS){
      const indexes=odFunctionIndexes.get(fn),start=startsByFunction[fn],final=finalsByFunction[fn],startTotal=start.reduce((sum,value)=>sum+value,0),targetTotal=aggregateTarget(scenario,fn,monthIndex,startTotal),target=trajectoryVector(start,final,targetTotal,`${scenario}|${fn}`);
      indexes.forEach((index,localIndex)=>{
        const delta=target[localIndex]-organicState[index];
        if(delta>0)addOdMovement(events,scenario,month,"Hire",index,delta);
        if(delta<0){
          const exitFte=-delta,running=forcedRunning[fn],share=fn==="Finance"?.95:.30,newTotal=running.total+exitFte,desiredForced=Math.round(newTotal*share),forced=Math.max(0,Math.min(exitFte,desiredForced-running.forced));
          addOdMovement(events,scenario,month,"Forced Exit",index,forced);addOdMovement(events,scenario,month,"Natural Attrition",index,exitFte-forced);running.total=newTotal;running.forced+=forced;
        }
        organicState[index]=target[localIndex];
      });
    }
    const transferDue=scenario!=="Baseline"&&((scenario==="Actual"&&monthIndex===1)||(scenario!=="Actual"&&(monthIndex+1)%6===0));
    if(transferDue&&organicState[dtSource]+transferOffset[dtSource]>1){const pairId=traceId("ODT",scenario,month.toISOString());addOdMovement(events,scenario,month,"Transfer Out",dtSource,1,dtTarget,pairId);addOdMovement(events,scenario,month,"Transfer In",dtTarget,1,dtSource,pairId);transferOffset[dtSource]--;transferOffset[dtTarget]++}
    odDimensions.forEach((_,index)=>state[index]=organicState[index]+transferOffset[index]);
    const movementTotals=odDimensions.map(()=>({Hire:0,"Natural Attrition":0,"Forced Exit":0,"Transfer In":0,"Transfer Out":0,signed:0,enr:0}));
    for(const event of events){movementTotals[event.index][event.type]+=event.fte;movementTotals[event.index].signed+=event.signedFte;movementTotals[event.index].enr+=event.enr}
    const staffByIndex=odDimensions.map(()=>0);
    for(const fn of OD_FUNCTIONS_ALL){
      const indexes=odFunctionIndexes.get(fn);
      // Baseline scenario: freeze at June-2026 baseline for every function including Operations.
      if(scenario==="Baseline"){indexes.forEach(index=>staffByIndex[index]=Math.round(odDimensions[index].baseFte*odDimensions[index].monthlyStaffCost));continue}
      // Out-of-scope stays at baseline staff cost in every scenario.
      if(fn===OD_OUTSCOPE_FUNCTION){indexes.forEach(index=>staffByIndex[index]=Math.round(odDimensions[index].baseFte*odDimensions[index].monthlyStaffCost));continue}
      // Progress towards the final per-FTE cost reduction, aligned with FTE trajectory milestones.
      let progress;
      if(scenario==="Actual")progress=rampProgress("Actual",monthIndex);
      else if(scenario==="Forecast"){const actualProgress=rampProgress("Actual",1);progress=actualProgress+(1-actualProgress)*rampProgress("Forecast",monthIndex)}
      else progress=rampProgress(scenario,monthIndex);
      const finalRatio=OD_STAFF_COST_PER_FTE_TARGETS[fn]??1,perFteMultiplier=1+(finalRatio-1)*progress;
      // Distribute the function total across entities based on FTE-weighted expected cost (with D&T country mix).
      const raw=indexes.map(index=>((opening[index]+state[index])/2)*odDimensions[index].monthlyStaffCost*(fn==="D&T"?.96+unit(`od-dt-cost-mix|${odDimensions[index].country}`)*.08:1));
      const closingTotal=indexes.reduce((sum,index)=>sum+state[index],0);
      const functionTarget=Math.round(closingTotal*(odBaselineStaff[fn]/odBaselineFte[fn])*perFteMultiplier);
      const allocated=allocateInteger(functionTarget,raw);indexes.forEach((index,localIndex)=>staffByIndex[index]=allocated[localIndex]);
    }
    odDimensions.forEach((dimension,index)=>{
      const movements=movementTotals[index],staffCost=staffByIndex[index],social=Math.round(staffCost*dimension.socialRate),enr=movements.enr,recurringTotal=staffCost+social,totalCash=recurringTotal+enr;
      const baselineStaff=Math.round(dimension.baseFte*dimension.monthlyStaffCost),baselineSocial=Math.round(baselineStaff*dimension.socialRate),baselineTotal=baselineStaff+baselineSocial,recurringSavings=baselineTotal-recurringTotal,netSavings=recurringSavings-enr;
      recurringCumulative[index]+=recurringSavings;netCumulative[index]+=netSavings;
      const id=traceId("ODMTH",scenario,month.toISOString(),dimension.entity,dimension.fn);
      const method=scenario==="Baseline"?"Frozen June 2026 baseline ledger":(dimension.scope==="Out-of-scope"?"Out-of-scope Operations ledger frozen at June 2026 baseline":"Target-driven opening-movement-closing and transformation economics ledger");
      odMonthly.push([id,month,fiscalYear(month),scenario,dimension.entity,dimension.region,dimension.cluster,dimension.country,dimension.fn,dimension.scope,opening[index],movements.Hire,movements["Natural Attrition"],movements["Forced Exit"],movements["Natural Attrition"]+movements["Forced Exit"],movements["Transfer In"],movements["Transfer Out"],movements.signed,state[index],staffCost,social,enr,recurringTotal,totalCash,baselineTotal,recurringSavings,netSavings,recurringCumulative[index],netCumulative[index],"Synthetic",method]);
      trace.push([id,"OD_Monthly",dimension.entity,dimension.country,dimension.fn,scenario,fiscalYear(month),method,SEED]);
    });
  }
  if(scenario==="Actual")odDimensions.forEach((_,index)=>{actualClosing.set(index,state[index]);actualRecurringCumulative.set(index,recurringCumulative[index]);actualNetCumulative.set(index,netCumulative[index])});
}
for(const scenario of ["Baseline","Budget","Plan","Actual","Forecast"])buildOdScenario(scenario);
appendBook(workbook,"OD_Movements",odMovements);
appendBook(workbook,"OD_Monthly",odMonthly);

trace.push([traceId("ODA","SCOPE"),"OD_Assumptions","","","ALL","Program scope","June 2026 baseline",`In-scope program: 6 support functions totalling ${OD_INSCOPE_TOTAL} FTE (Sales ${OD_BASELINE_TARGETS.Sales}, Supply ${OD_BASELINE_TARGETS.Supply}, Finance ${OD_BASELINE_TARGETS.Finance}, HR ${OD_BASELINE_TARGETS.HR}, D&T ${OD_BASELINE_TARGETS["D&T"]}, Marketing ${OD_BASELINE_TARGETS.Marketing}); virtual Operations function carries ${OD_OUTSCOPE_TOTAL} FTE out of scope for a group total of ${OD_GROUP_TOTAL} FTE`,SEED]);
trace.push([traceId("ODA","TARGETS"),"OD_Assumptions","","","ALL","Plan FTE and staff-cost targets","June 2026-June 2029",`Group in-scope staff-cost milestones -3% end FY26/27, -7% end FY27/28, -10% end FY28/29; final FTE Finance -50%, Sales -2.5%, Marketing -2.5%, HR -5%, D&T +2%, Supply calibrated (${odPlanFinal.Supply} FTE, ${((odPlanFinal.Supply/odBaselineFte.Supply-1)*100).toFixed(1)}%); Operations out-of-scope FTE frozen at ${OD_OUTSCOPE_TOTAL}`,SEED]);
trace.push([traceId("ODA","ENR"),"OD_Assumptions","","","ALL","Transformation Social Cost (ENR)","Movement month","Forced Exit only; nine months of employer cost with deterministic country factor 90%-110%",SEED]);
trace.push([traceId("ODA","DT"),"OD_Assumptions","","","D&T","Nearshoring/internalization economics","FY26/27-FY28/29","D&T +2% final FTE and -10% final recurring staff cost per FTE (approx -8.2% total staff cost); paired transfers remain group neutral",SEED]);
trace.push([traceId("ODA","FORCED"),"OD_Assumptions","","","ALL","Forced exit share of total exits","FY26/27-FY28/29","Finance ~95%; Sales, Supply, HR, D&T, Marketing ~30%",SEED]);
appendBook(workbook,"Data Dictionary",[["Sheet","Field / Range","Definition","Status"],["Mapping","A:F","Synthetic legal entity geography reference","Synthetic"],["Cost Baseline","A:P","FY25/26 annual recurring cost baseline for the 6 in-scope support functions","Synthetic"],["Lever BCase - Updated","A:AV","Annual synthetic transformation initiatives","Synthetic"],["Budget FY26-27","A:S","Explicit budget impacts for every entity/function/category","Synthetic"],["Monthly_Data","A:N","Reconciled monthly baseline facts","Synthetic"],["KPI_Catalog","A:K","KPI definitions, current, target and global benchmark","Synthetic except illustrative benchmarks"],["KPI_History","A:H","Monthly KPI observations by geography","Synthetic"],["OD_Movements","A:X","Hire, Natural Attrition, Forced Exit, Transfer In and Transfer Out events; exit subtype, signed FTE, Program Scope, recurring cost impact and one-off ENR impact","Synthetic"],["OD_Movements","N:N","Program Scope flag - In-scope (6 support functions) or Out-of-scope (Operations)","Synthetic"],["OD_Movements","T:T","Transformation Social Cost (ENR): one-off cash cost only for Forced Exit, equal to nine months of employer cost with deterministic country variation","Synthetic"],["OD_Monthly","A:AE","Monthly OD ledger from frozen June 2026 baseline, Plan/Budget from July 2026, Actual July-August 2026 and Forecast September 2026-June 2029; includes out-of-scope Operations ledger frozen at 8450 FTE","Synthetic"],["OD_Monthly","J:J","Program Scope flag - In-scope (6 support functions covering 1550 FTE) or Out-of-scope (Operations, 8450 FTE frozen across scenarios)","Synthetic"],["OD_Monthly","T:X","Recurring Staff Cost, Recurring Employer Social Charges, Transformation Social Cost (ENR), Recurring Total People Cost and Total Cash Cost Including ENR","Synthetic"],["OD_Monthly","Y:AC","Baseline Recurring Total People Cost, Recurring Savings vs Baseline, Net Savings After ENR, Cumulative Recurring Savings and Cumulative Net Savings","Synthetic"],["OD_Assumptions","Program scope","In-scope program covers 1550 FTE across 6 support functions; virtual Operations function carries 8450 FTE out of scope for a group total of 10000 FTE","Synthetic"],["OD_Assumptions","Staff cost milestones","Group in-scope staff cost declines -3% by end FY26/27, -7% by end FY27/28 and -10% by end FY28/29","Synthetic"],["Generation Log","A:I","Traceability metadata plus explicit OD target, ENR, D&T economics and program scope assumptions","Synthetic"]]);
appendBook(workbook,"Generation Log",[["Trace ID","Dataset","Legal Entity","Country","Function","Category / KPI","Period","Generation Method","Seed"],...trace]);

workbook.Props={Title:"Gooduelle Performance Demo - Traceable Synthetic Data",Subject:"Fully synthetic and traceable demonstration dataset",Author:"Gooduelle Demo Generator",Company:"Gooduelle",Comments:`Generated deterministically with seed ${SEED}. No operational source data used.`};
fs.writeFileSync("Gooduelle_Performance_Demo.xlsx",XLSX.write(workbook,{type:"buffer",bookType:"xlsx",compression:true}));
