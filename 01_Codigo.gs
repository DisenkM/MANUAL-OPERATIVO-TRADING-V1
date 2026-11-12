
/* ============================================================
   CUADERNO DEL TRADER V5.4
   Backend Google Apps Script
   Binance Spot · Sin apalancamiento
   Capital general · Asignaciones · Transferencias
   Rentabilidad · Riesgo configurable · Auditoría
   ============================================================ */

const V54 = {
  VERSION: "5.4",
  TZ: "America/Bogota",
  SHEETS: {
    CONFIG: "V54_CONFIG",
    STRATEGIES: "V54_ESTRATEGIAS",
    RULES: "V54_REGLAS",
    TRADES: "V54_OPERACIONES",
    TRANSFERS: "V54_TRANSFERENCIAS",
    RISKLOG: "V54_HISTORIAL_RIESGO",
    AUDIT: "V54_AUDITORIA"
  },
  HEADERS: {
    V54_CONFIG: ["clave","valor","actualizado"],
    V54_ESTRATEGIAS: ["id","nombre","simbolo","mercado","estado","version","creada"],
    V54_REGLAS: ["id","strategy_id","fase","nombre","descripcion","tipo","condicion","obligatoria","activa","version","creada"],
    V54_OPERACIONES: [
      "id","strategy_id","strategy_name","strategy_version","symbol",
      "direction","status","mode","entry","stop","target","quantity",
      "notional","risk_usd","risk_pct","entry_fee","exit","exit_fee",
      "gross_pnl","net_pnl","r_multiple","emotion","notes","opened_at","closed_at"
    ],
    V54_TRANSFERENCIAS: ["id","fecha","origen","destino","importe","motivo"],
    V54_HISTORIAL_RIESGO: ["id","fecha","clave","anterior","nuevo","motivo"],
    V54_AUDITORIA: ["id","fecha","evento","detalle"]
  }
};

const V54_DEFAULTS = {
  CAPITAL_USD: 80,
  RIESGO_REAL_PCT: 0.5,
  RIESGO_PRUEBAS_PCT: 1,
  COMISION_SPOT: 0.001,
  MIN_NOTIONAL_USD: 5,
  TRM_COP_USD: 4000,
  MODO: "SIMULACION",
  LIMITE_MENSUAL_PCT: 0,
  TECHO_SEGURIDAD_PCT: 0
};

/* ---------- INICIO Y BASE DE DATOS ---------- */

function doGet() {
  return HtmlService.createHtmlOutputFromFile("V1_Index")
    .setTitle("Cuaderno del Trader V5.4")
    .addMetaTag("viewport","width=device-width, initial-scale=1");
}

function setupV54() {
  const ss = SpreadsheetApp.getActive();
  Object.keys(V54.HEADERS).forEach(name => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    const headers = V54.HEADERS[name];
    if (sh.getLastRow() === 0) {
      sh.getRange(1,1,1,headers.length).setValues([headers]);
      sh.setFrozenRows(1);
    } else {
      const existing = sh.getRange(1,1,1,headers.length).getValues()[0];
      if (headers.some((h,i) => existing[i] !== h)) {
        throw new Error("Encabezados incompatibles en "+name+
          ". No se modificó la hoja. Haz una copia y revisa su estructura.");
      }
    }
  });

  const cfg = getConfig_();
  Object.keys(V54_DEFAULTS).forEach(k => {
    if (cfg[k] === undefined || cfg[k] === "") {
      setConfigValue_(k,V54_DEFAULTS[k]);
    }
  });

  if (getRows_(V54.SHEETS.STRATEGIES).length === 0) {
    append_(V54.SHEETS.STRATEGIES,[
      "EST-001","Estrategia A","BTC/USDT","SPOT","BORRADOR",1,now_()
    ]);
  }

  audit_("SETUP_V54","Base V5.4 inicializada o verificada.");
  return "V5.4 lista. Las hojas V5.3 no fueron borradas.";
}

/* ---------- UTILIDADES ---------- */

function ss_(){ return SpreadsheetApp.getActive(); }
function sh_(name){
  const sh = ss_().getSheetByName(name);
  if (!sh) throw new Error("Falta la hoja "+name+". Ejecuta setupV54().");
  return sh;
}
function now_(){ return Utilities.formatDate(new Date(),V54.TZ,"yyyy-MM-dd HH:mm:ss"); }
function uid_(prefix){ return prefix+"-"+Utilities.getUuid().slice(0,8).toUpperCase(); }
function num_(v){ const n=Number(v); return Number.isFinite(n)?n:0; }
function bool_(v){ return v===true || String(v).toUpperCase()==="TRUE"; }
function append_(sheet,row){ sh_(sheet).appendRow(row); }
function getRows_(sheet){
  const s=sh_(sheet), data=s.getDataRange().getValues();
  if(data.length<2) return [];
  const headers=data[0];
  return data.slice(1).filter(r=>r.some(v=>v!=="")).map(r=>{
    const o={}; headers.forEach((h,i)=>o[h]=r[i]); return o;
  });
}
function audit_(event,detail){
  append_(V54.SHEETS.AUDIT,[uid_("AUD"),now_(),event,String(detail)]);
}
function configMap_(){
  const out={};
  getRows_(V54.SHEETS.CONFIG).forEach(r=>out[r.clave]=r.valor);
  return out;
}
function getConfig_(){ return configMap_(); }
function setConfigValue_(key,value){
  const sh=sh_(V54.SHEETS.CONFIG), rows=getRows_(V54.SHEETS.CONFIG);
  const idx=rows.findIndex(r=>r.clave===key);
  if(idx<0) sh.appendRow([key,value,now_()]);
  else sh.getRange(idx+2,2,1,2).setValues([[value,now_()]]);
}
function round_(n,d=8){ return Math.round((n+Number.EPSILON)*Math.pow(10,d))/Math.pow(10,d); }
function requireText_(v,label){
  const s=String(v||"").trim();
  if(!s) throw new Error(label+" es obligatorio.");
  return s;
}

/* ---------- DATOS INICIALES ---------- */

function api_getInitialData(){
  setupGuard_();
  const config=configMap_();
  const strategies=getRows_(V54.SHEETS.STRATEGIES);
  const rules=getRows_(V54.SHEETS.RULES);
  const trades=getRows_(V54.SHEETS.TRADES);
  const transfers=getRows_(V54.SHEETS.TRANSFERS);
  const riskLog=getRows_(V54.SHEETS.RISKLOG);
  const allocations=getAllocations_();
  return {
    success:true,version:V54.VERSION,config,
    strategies,rules,trades,transfers,riskLog,allocations,
    summary:summary_(strategies,trades,allocations,config)
  };
}
function setupGuard_(){
  if(!ss_().getSheetByName(V54.SHEETS.CONFIG)) setupV54();
}

/* ---------- CONFIGURACIÓN Y RIESGO ---------- */

function api_saveConfig(p){
  setupGuard_();
  p=p||{};
  const old=configMap_();
  const editable=[
    "CAPITAL_USD","RIESGO_REAL_PCT","RIESGO_PRUEBAS_PCT",
    "COMISION_SPOT","MIN_NOTIONAL_USD","TRM_COP_USD",
    "MODO","LIMITE_MENSUAL_PCT","TECHO_SEGURIDAD_PCT"
  ];
  const changes=[];
  editable.forEach(k=>{
    if(p[k]===undefined) return;
    const v=(k==="MODO")?String(p[k]):num_(p[k]);
    if(k!=="MODO" && v<0) throw new Error(k+" no puede ser negativo.");
    if(k==="CAPITAL_USD" && v<=0) throw new Error("El capital debe ser mayor que cero.");
    if(k==="RIESGO_REAL_PCT" && v<=0) throw new Error("El riesgo real debe ser mayor que cero.");
    if(k==="RIESGO_PRUEBAS_PCT" && v<=0) throw new Error("El riesgo de pruebas debe ser mayor que cero.");
    if(k==="COMISION_SPOT" && v>=0.1) throw new Error("Revisa la comisión: debe expresarse como decimal, por ejemplo 0.001.");
    if(k==="MODO" && !["REAL","SIMULACION"].includes(v)) throw new Error("Modo no válido.");
    if(String(old[k])!==String(v)) changes.push({k,old:old[k]??"",v});
  });

  const riskIncrease=changes.filter(c=>
    c.k==="RIESGO_REAL_PCT" && num_(c.v)>num_(c.old)
  );
  if(riskIncrease.length){
    if(p.confirmacionAumento!==true) throw new Error("Aumentar el riesgo requiere confirmación.");
    if(String(p.motivo||"").trim().length<8) throw new Error("Escribe un motivo de al menos 8 caracteres.");
    const cap=num_(p.TECHO_SEGURIDAD_PCT ?? old.TECHO_SEGURIDAD_PCT);
    if(cap>0 && num_(riskIncrease[0].v)>cap)
      throw new Error("El riesgo supera tu techo de seguridad configurado.");
  }

  changes.forEach(c=>{
    setConfigValue_(c.k,c.v);
    if(c.k==="RIESGO_REAL_PCT" || c.k==="RIESGO_PRUEBAS_PCT"){
      append_(V54.SHEETS.RISKLOG,[
        uid_("RISK"),now_(),c.k,c.old,c.v,
        String(p.motivo||"Ajuste de configuración")
      ]);
    }
  });

  if(changes.length) audit_("CONFIG_CAMBIADA",JSON.stringify({
    changes,motivo:p.motivo||"No aplica"
  }));
  return {success:true};
}

/* ---------- ESTRATEGIAS Y VERSIONES ---------- */

function api_createStrategy(p){
  setupGuard_(); p=p||{};
  const name=requireText_(p.nombre,"Nombre");
  const symbol=requireText_(p.simbolo||"BTC/USDT","Símbolo");
  const id=uid_("EST");
  append_(V54.SHEETS.STRATEGIES,[id,name,symbol,"SPOT","BORRADOR",1,now_()]);
  audit_("ESTRATEGIA_CREADA",id+" "+name);
  return {success:true,id};
}
function api_saveRule(p){
  setupGuard_(); p=p||{};
  const strategy=requireText_(p.strategy_id,"Estrategia");
  const s=findStrategy_(strategy);
  if(!s) throw new Error("Estrategia no encontrada.");
  const phase=requireText_(p.fase,"Fase");
  if(!["ANTES","DURANTE","DESPUES"].includes(phase)) throw new Error("Fase inválida.");
  const type=p.tipo==="AUTO"?"AUTO":"MANUAL";
  const condition=p.condicion||"MANUAL_CONFIRM";
  const allowed=["MANUAL_CONFIRM","LONG_ONLY","SL_SIDE_VALID","RISK_WITHIN_LIMIT","TP_SIDE_VALID"];
  if(!allowed.includes(condition)) throw new Error("Condición automática no reconocida.");
  const next=Number(s.version)+1;
  updateStrategyVersion_(strategy,next);
  append_(V54.SHEETS.RULES,[
    uid_("RUL"),strategy,phase,requireText_(p.nombre,"Nombre de regla"),
    String(p.descripcion||""),type,condition,p.obligatoria!==false,
    true,next,now_()
  ]);
  audit_("REGLA_CREADA",strategy+" v"+next);
  return {success:true,version:next};
}
function api_deleteRule(id){
  setupGuard_();
  const sh=sh_(V54.SHEETS.RULES), data=sh.getDataRange().getValues();
  for(let i=1;i<data.length;i++){
    if(data[i][0]===id && bool_(data[i][8])){
      const strategy=data[i][1], s=findStrategy_(strategy);
      const next=Number(s.version)+1;
      sh.getRange(i+1,9).setValue(false);
      sh.getRange(i+1,10).setValue(Number(s.version));
      updateStrategyVersion_(strategy,next);
      audit_("REGLA_DESACTIVADA",id+" nueva versión "+next);
      return {success:true,version:next};
    }
  }
  throw new Error("Regla activa no encontrada.");
}
function findStrategy_(id){
  return getRows_(V54.SHEETS.STRATEGIES).find(s=>s.id===id);
}
function updateStrategyVersion_(id,v){
  const sh=sh_(V54.SHEETS.STRATEGIES), data=sh.getDataRange().getValues();
  for(let i=1;i<data.length;i++) if(data[i][0]===id){
    sh.getRange(i+1,6).setValue(v); return;
  }
}

/* ---------- CAPITAL Y ASIGNACIONES ---------- */

function getAllocations_(){
  const strategies=getRows_(V54.SHEETS.STRATEGIES);
  const transfers=getRows_(V54.SHEETS.TRANSFERS);
  const map={"SIN_ASIGNAR":num_(configMap_().CAPITAL_USD)};
  strategies.forEach(s=>map[s.id]=0);
  transfers.forEach(t=>{
    const amount=num_(t.importe);
    if(map[t.origen]===undefined) map[t.origen]=0;
    if(map[t.destino]===undefined) map[t.destino]=0;
    map[t.origen]=round_(map[t.origen]-amount,8);
    map[t.destino]=round_(map[t.destino]+amount,8);
  });
  return map;
}
function api_transferCapital(p){
  setupGuard_(); p=p||{};
  const from=requireText_(p.origen,"Origen");
  const to=requireText_(p.destino,"Destino");
  const amount=num_(p.importe);
  if(from===to) throw new Error("Origen y destino deben ser diferentes.");
  if(amount<=0) throw new Error("El importe debe ser mayor que cero.");
  if(from!=="SIN_ASIGNAR" && !findStrategy_(from)) throw new Error("Estrategia origen inválida.");
  if(to!=="SIN_ASIGNAR" && !findStrategy_(to)) throw new Error("Estrategia destino inválida.");
  const balances=getAllocations_();
  if(num_(balances[from])<amount) throw new Error("No hay capital suficiente en el origen.");
  append_(V54.SHEETS.TRANSFERS,[
    uid_("MOV"),now_(),from,to,round_(amount,8),
    requireText_(p.motivo,"Motivo")
  ]);
  audit_("TRANSFERENCIA_CAPITAL",from+" -> "+to+" : "+amount);
  return {success:true,allocations:getAllocations_()};
}

/* ---------- CÁLCULO DE POSICIÓN ---------- */

function api_calculatePosition(p){
  setupGuard_(); p=p||{};
  const strategy=findStrategy_(p.strategy_id);
  if(!strategy) throw new Error("Selecciona una estrategia válida.");
  const cfg=configMap_(), allocations=getAllocations_();
  const capital=num_(allocations[strategy.id]);
  if(capital<=0) throw new Error("Asigna capital a esta estrategia antes de calcular.");
  const entry=num_(p.entrada), stop=num_(p.stop_loss), target=num_(p.take_profit);
  if(entry<=0 || stop<=0 || target<=0) throw new Error("Entrada, stop y objetivo deben ser mayores que cero.");
  if(!(stop<entry && target>entry)) throw new Error("En Spot LONG, el stop debe estar debajo de la entrada y el objetivo encima.");
  const mode=String(cfg.MODO||"SIMULACION");
  const riskPct=num_(mode==="REAL"?cfg.RIESGO_REAL_PCT:cfg.RIESGO_PRUEBAS_PCT)/100;
  const fee=num_(cfg.COMISION_SPOT);
  const riskBudget=capital*riskPct;
  const unitLoss=(entry-stop)+(entry+stop)*fee;
  if(unitLoss<=0) throw new Error("Distancia de stop o costos inválidos.");
  const qty=Math.floor((riskBudget/unitLoss)*1e6)/1e6;
  const notional=qty*entry;
  const grossLoss=qty*(entry-stop);
  const fees=qty*(entry+stop)*fee;
  const totalLoss=grossLoss+fees;
  const tpFees=qty*(entry+target)*fee;
  const netProfit=qty*(target-entry)-tpFees;
  const rr=totalLoss>0?netProfit/totalLoss:0;
  return {
    success:true,quantity:qty,notional_usd:round_(notional,8),
    risk_budget_usd:round_(riskBudget,8),
    estimated_loss_usd:round_(totalLoss,8),
    estimated_fees_usd:round_(fees,8),
    estimated_net_profit_usd:round_(netProfit,8),
    net_rr:round_(rr,4),
    risk_pct:round_(totalLoss/capital*100,4),
    minimum_ok:notional>=num_(cfg.MIN_NOTIONAL_USD),
    authorized:qty>0 && totalLoss<=riskBudget+1e-8 &&
      notional>=num_(cfg.MIN_NOTIONAL_USD)
  };
}

/* ---------- AUTORIZAR Y REGISTRAR OPERACIÓN ---------- */

function api_authorizeTrade(p){
  setupGuard_(); p=p||{};
  const strategy=findStrategy_(p.strategy_id);
  if(!strategy) throw new Error("Estrategia no encontrada.");
  const calc=api_calculatePosition({
    strategy_id:p.strategy_id,entrada:p.entrada,
    stop_loss:p.stop_loss,take_profit:p.take_profit
  });
  if(!calc.authorized) throw new Error("La operación no cumple el presupuesto de riesgo o el mínimo de referencia.");

  const rules=getRows_(V54.SHEETS.RULES).filter(r=>
    r.strategy_id===p.strategy_id && r.fase==="ANTES" &&
    bool_(r.activa) && Number(r.version)===Number(strategy.version) &&
    bool_(r.obligatoria)
  );
  const answers=p.answers||{};
  const failures=[];
  rules.forEach(r=>{
    let ok=false;
    if(r.tipo==="MANUAL") ok=answers[r.id]===true;
    else {
      switch(r.condicion){
        case "LONG_ONLY": ok=p.direccion==="LONG"; break;
        case "SL_SIDE_VALID": ok=num_(p.stop_loss)<num_(p.entrada); break;
        case "TP_SIDE_VALID": ok=num_(p.take_profit)>num_(p.entrada); break;
        case "RISK_WITHIN_LIMIT": ok=calc.authorized; break;
        case "MANUAL_CONFIRM": ok=answers[r.id]===true; break;
        default: ok=false;
      }
    }
    if(!ok) failures.push({id:r.id,nombre:r.nombre,motivo:"Regla obligatoria no cumplida"});
  });
  if(failures.length) return {authorized:false,failures};

  const cfg=configMap_(), id=uid_("TRD"), now=now_();
  const entry=num_(p.entrada), stop=num_(p.stop_loss), target=num_(p.take_profit);
  const qty=calc.quantity, fee=qty*entry*num_(cfg.COMISION_SPOT);
  append_(V54.SHEETS.TRADES,[
    id,strategy.id,strategy.nombre,Number(strategy.version),
    strategy.simbolo,p.direccion||"LONG","ABIERTA",cfg.MODO,
    entry,stop,target,qty,calc.notional_usd,
    calc.estimated_loss_usd,calc.risk_pct,
    round_(fee,8),"",0,0,0,0,
    p.emocion||"NEUTRAL",String(p.notas||""),now,""
  ]);
  audit_("OPERACION_ABIERTA",id+" "+strategy.nombre+" v"+strategy.version);
  return {authorized:true,id_trade:id,calculation:calc};
}

/* ---------- CIERRE DE OPERACIÓN ---------- */

function api_closeTrade(p){
  setupGuard_(); p=p||{};
  const id=requireText_(p.id,"ID de operación");
  const exit=num_(p.exit);
  if(exit<=0) throw new Error("El precio de salida debe ser mayor que cero.");
  const sh=sh_(V54.SHEETS.TRADES), data=sh.getDataRange().getValues();
  for(let i=1;i<data.length;i++){
    if(data[i][0]!==id) continue;
    const row=data[i];
    if(row[6]!=="ABIERTA") throw new Error("La operación ya está cerrada.");
    const direction=row[5], qty=num_(row[11]), entry=num_(row[8]);
    if(direction!=="LONG") throw new Error("Esta versión solo calcula cierres LONG de Spot.");
    const cfg=configMap_();
    const exitFee=(p.exit_fee!==undefined && p.exit_fee!=="")
      ?num_(p.exit_fee):qty*exit*num_(cfg.COMISION_SPOT);
    const gross=qty*(exit-entry);
    const net=gross-num_(row[15])-exitFee;
    const risk=num_(row[13]);
    const rMultiple=risk>0?net/risk:0;
    sh.getRange(i+1,7).setValue("CERRADA");
    sh.getRange(i+1,17,1,9).setValues([[
      exit,round_(exitFee,8),round_(gross,8),round_(net,8),
      round_(rMultiple,4),row[20],row[21],row[22],now_()
    ]]);
    audit_("OPERACION_CERRADA",id+" neto="+round_(net,8));
    return {success:true,net_pnl:round_(net,8),gross_pnl:round_(gross,8),r_multiple:round_(rMultiple,4)};
  }
  throw new Error("Operación no encontrada.");
}

/* ---------- ESTADÍSTICAS Y RENTABILIDAD ---------- */

function strategyStats_(strategyId,trades){
  const closed=trades.filter(t=>t.strategy_id===strategyId && t.status==="CERRADA");
  const wins=closed.filter(t=>num_(t.net_pnl)>0);
  const losses=closed.filter(t=>num_(t.net_pnl)<0);
  const grossWins=wins.reduce((a,t)=>a+num_(t.gross_pnl),0);
  const grossLosses=Math.abs(losses.reduce((a,t)=>a+num_(t.gross_pnl),0));
  const fees=closed.reduce((a,t)=>a+num_(t.entry_fee)+num_(t.exit_fee),0);
  const net=closed.reduce((a,t)=>a+num_(t.net_pnl),0);
  const allocated=num_(getAllocations_()[strategyId]);
  return {
    strategy_id:strategyId,trades:closed.length,winners:wins.length,
    losers:losses.length,win_rate:closed.length?wins.length/closed.length*100:0,
    gross_wins:round_(grossWins,8),gross_losses:round_(grossLosses,8),
    fees:round_(fees,8),net:round_(net,8),
    allocated,return_pct:allocated?round_(net/allocated*100,4):0
  };
}
function summary_(strategies,trades,allocations,config){
  const closed=trades.filter(t=>t.status==="CERRADA");
  const wins=closed.filter(t=>num_(t.net_pnl)>0);
  const losses=closed.filter(t=>num_(t.net_pnl)<0);
  const grossWins=wins.reduce((a,t)=>a+num_(t.gross_pnl),0);
  const grossLosses=Math.abs(losses.reduce((a,t)=>a+num_(t.gross_pnl),0));
  const fees=closed.reduce((a,t)=>a+num_(t.entry_fee)+num_(t.exit_fee),0);
  const net=closed.reduce((a,t)=>a+num_(t.net_pnl),0);
  const allocated=Object.keys(allocations).reduce((a,k)=>a+num_(allocations[k]),0);
  return {
    capital:num_(config.CAPITAL_USD),allocated:round_(allocated,8),
    unassigned:round_(num_(allocations.SIN_ASIGNAR),8),
    gross_wins:round_(grossWins,8),gross_losses:round_(grossLosses,8),
    fees:round_(fees,8),net:round_(net,8),
    open_trades:trades.filter(t=>t.status==="ABIERTA").length,
    closed_trades:closed.length,winners:wins.length,losers:losses.length,
    by_strategy:strategies.map(s=>({
      id:s.id,name:s.nombre,...strategyStats_(s.id,trades)
    }))
  };
}
function api_getStrategyStats(id){
  const trades=getRows_(V54.SHEETS.TRADES);
  return strategyStats_(id,trades);
}

/* ===== V5.4 PATCH: VERSIONADO Y CIERRE SEGURO ===== */

function avanzarVersion_(strategyId, excludeRuleId) {
  const strategy=findStrategy_(strategyId);
  if(!strategy) throw new Error("Estrategia no encontrada.");

  const oldVersion=Number(strategy.version);
  const newVersion=oldVersion+1;
  const active=getRows_(V54.SHEETS.RULES).filter(r=>
    r.strategy_id===strategyId &&
    Number(r.version)===oldVersion &&
    bool_(r.activa) &&
    r.id!==excludeRuleId
  );

  active.forEach(r=>{
    append_(V54.SHEETS.RULES,[
      uid_("RUL"),r.strategy_id,r.fase,r.nombre,r.descripcion,
      r.tipo,r.condicion,r.obligatoria,true,newVersion,now_()
    ]);
  });

  updateStrategyVersion_(strategyId,newVersion);
  return newVersion;
}

function api_saveRule(p) {
  setupGuard_(); p=p||{};
  const strategy=requireText_(p.strategy_id,"Estrategia");
  if(!findStrategy_(strategy)) throw new Error("Estrategia no encontrada.");
  const phase=requireText_(p.fase,"Fase");
  if(!["ANTES","DURANTE","DESPUES"].includes(phase))
    throw new Error("Fase inválida.");

  const type=p.tipo==="AUTO"?"AUTO":"MANUAL";
  const condition=p.condicion||"MANUAL_CONFIRM";
  const allowed=[
    "MANUAL_CONFIRM","LONG_ONLY","SL_SIDE_VALID",
    "RISK_WITHIN_LIMIT","TP_SIDE_VALID"
  ];
  if(!allowed.includes(condition))
    throw new Error("Condición automática no reconocida.");

  const version=avanzarVersion_(strategy,null);
  append_(V54.SHEETS.RULES,[
    uid_("RUL"),strategy,phase,
    requireText_(p.nombre,"Nombre de regla"),
    String(p.descripcion||""),type,condition,
    p.obligatoria!==false,true,version,now_()
  ]);
  audit_("REGLA_CREADA",strategy+" v"+version);
  return {success:true,version};
}

function api_deleteRule(id) {
  setupGuard_();
  const rule=getRows_(V54.SHEETS.RULES).find(r=>
    r.id===id && bool_(r.activa)
  );
  if(!rule) throw new Error("Regla activa no encontrada.");

  const strategy=findStrategy_(rule.strategy_id);
  if(Number(rule.version)!==Number(strategy.version))
    throw new Error("La regla no pertenece a la versión activa.");

  const version=avanzarVersion_(rule.strategy_id,id);
  audit_("REGLA_DESACTIVADA",id+"; nueva versión "+version);
  return {success:true,version};
}

function api_closeTrade(p) {
  setupGuard_(); p=p||{};
  const id=requireText_(p.id,"ID de operación");
  const exit=num_(p.exit);
  if(exit<=0) throw new Error("El precio de salida debe ser mayor que cero.");

  const sh=sh_(V54.SHEETS.TRADES);
  const data=sh.getDataRange().getValues();

  for(let i=1;i<data.length;i++){
    const row=data[i];
    if(row[0]!==id) continue;
    if(row[6]!=="ABIERTA") throw new Error("La operación ya está cerrada.");

    const direction=row[5];
    if(direction!=="LONG")
      throw new Error("Esta versión solo admite cierres LONG en Spot.");

    const qty=num_(row[11]);
    const entry=num_(row[8]);
    const cfg=configMap_();
    const exitFee=(p.exit_fee!==undefined && p.exit_fee!=="")
      ?num_(p.exit_fee)
      :qty*exit*num_(cfg.COMISION_SPOT);

    if(exitFee<0) throw new Error("La comisión no puede ser negativa.");

    const gross=qty*(exit-entry);
    const net=gross-num_(row[15])-exitFee;
    const risk=num_(row[13]);
    const rMultiple=risk>0?net/risk:0;

    // Columnas Q-U: salida, comisión salida, bruto, neto y múltiplo R.
    sh.getRange(i+1,17,1,5).setValues([[
      exit,round_(exitFee,8),round_(gross,8),
      round_(net,8),round_(rMultiple,4)
    ]]);

    // Columna Y: fecha de cierre.
    sh.getRange(i+1,25).setValue(now_());

    audit_("OPERACION_CERRADA",id+" neto="+round_(net,8));
    return {
      success:true,
      net_pnl:round_(net,8),
      gross_pnl:round_(gross,8),
      r_multiple:round_(rMultiple,4)
    };
  }
  throw new Error("Operación no encontrada.");
}
