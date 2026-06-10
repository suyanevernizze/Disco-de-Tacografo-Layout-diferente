(function(){
  if(localStorage.getItem('binotto-theme')==='light')
    document.documentElement.classList.add('light');
})();

document.addEventListener('DOMContentLoaded', function() {

// ── TEMA ──
window.toggleTheme = function(){
  const light = document.documentElement.classList.toggle('light');
  localStorage.setItem('binotto-theme', light ? 'light' : 'dark');
  if(mensalAll.length){ renderMCharts(); renderDCharts(); renderVCharts(); }
};

// ── ESTADO ──
let mensalAll=[], mensalFilt=[], mPage=1, mSortCol='mes', mSortDir='asc';
let discosAll=[], discosFilt=[], dPage=1, dSortCol='mes', dSortDir='asc';
let vdoAll=[],    vdoFilt=[],    vPage=1, vSortCol='mes', vSortDir='asc';
let mActiveKpi=null, dActiveKpi=null, vActiveKpi=null;
const PS=50;
let CH={};
const COLORS=['#6C63FF','#1D9E75','#E24B4A','#378ADD','#BA7517','#D4537E','#5DCAA5','#EF9F27','#7F77DD','#D85A30','#85B7EB','#97C459','#F09595','#FAC775','#9FE1CB'];
const MES_ORD=['JANEIRO','FEVEREIRO','MARÇO','ABRIL','MAIO','JUNHO','JULHO','AGOSTO','SETEMBRO','OUTUBRO','NOVEMBRO','DEZEMBRO'];
const n=v=>(typeof v==='number'?v:parseFloat(String(v||'0').replace(',','.'))||0);
const fmt=d=>{if(!d)return'—';if(d instanceof Date&&!isNaN(d))return d.toLocaleDateString('pt-BR');if(typeof d==='string'&&d.includes('T'))return new Date(d).toLocaleDateString('pt-BR');return String(d);};
const g=id=>document.getElementById(id);
const v=id=>g(id)?.value||'';
function gc(){return document.documentElement.classList.contains('light');}
function gridC(){return gc()?'rgba(0,0,0,.07)':'rgba(42,52,71,.6)';}
function tickC(){return gc()?'#57606a':'#6e7681';}
function legC(){return gc()?'#57606a':'#8b949e';}

// ── TABS ──
window.switchTab = function(t){
  document.querySelectorAll('.tab').forEach((el,i)=>el.classList.toggle('active',['mensal','discos','vdo'][i]===t));
  document.querySelectorAll('.tab-panel').forEach(el=>el.classList.remove('active'));
  g('tab-'+t).classList.add('active');
};

// ── DRAG & DROP ──
const dz=g('dropZone');
dz.addEventListener('dragover',e=>{e.preventDefault();dz.classList.add('drag-over');});
dz.addEventListener('dragleave',()=>dz.classList.remove('drag-over'));
dz.addEventListener('drop',e=>{e.preventDefault();dz.classList.remove('drag-over');if(e.dataTransfer.files[0])processFile(e.dataTransfer.files[0]);});
window.loadFile=i=>{if(i.files[0])processFile(i.files[0]);};

// ── LOAD ──
function processFile(file){
  const badge=g('fileInfo');
  badge.textContent='⏳ Carregando '+file.name+'…';
  const reader=new FileReader();
  reader.onerror=()=>{badge.textContent='❌ Erro ao ler o arquivo';};
  reader.onload=e=>{
    try{
      const wb=XLSX.read(new Uint8Array(e.target.result),{type:'array',raw:false,cellDates:true});
      function findSheet(target){
        if(wb.Sheets[target])return target;
        const t=target.toLowerCase().replace(/\s+/g,' ').trim();
        return wb.SheetNames.find(n=>n.toLowerCase().replace(/\s+/g,' ').trim()===t)||null;
      }
      const nameM=findSheet('MENSAL DISCOS'),nameD=findSheet('Discos'),nameV=findSheet('Leitura VDO');
      const missing=[nameM?null:'MENSAL DISCOS',nameD?null:'Discos',nameV?null:'Leitura VDO'].filter(Boolean);
      if(missing.length)throw new Error('Aba(s) não encontrada(s): '+missing.join(', ')+'\n\nAbas na planilha:\n• '+wb.SheetNames.join('\n• '));

      const rowsM=XLSX.utils.sheet_to_json(wb.Sheets[nameM],{defval:'',raw:false});
      mensalAll=rowsM.filter(r=>r['MOTORISTA']&&String(r['MOTORISTA']).trim()).map(r=>({
        mes:String(r['MÊS']||'').trim().toUpperCase(),
        mesOrd:MES_ORD.indexOf(String(r['MÊS']||'').trim().toUpperCase()),
        motorista:String(r['MOTORISTA']||'').trim().toUpperCase(),
        placa:String(r['PLACA']||'').trim().toUpperCase(),
        contrato:String(r['CONTRATO']||'').trim().toUpperCase(),
        filial:String(r['FILIAL']||'').trim().toUpperCase(),
        quin1:n(r['1ª QUIN']),quin2:n(r['2ª QUIN']),meta:n(r['META MENSAL']),
        pendencia:n(r['PENDÊNCIA']),statusPend:String(r['STATUS PENDÊNCIA']||'').trim(),
        picos1:n(r['PICOS 1ª QUIN']),picos2:n(r['PICOS 2ª QUIN']),
        totalPicos:n(r['TOTAL PICOS']),obs:String(r['OBSERVAÇÃO']||'').trim(),
      }));

      const rowsD=XLSX.utils.sheet_to_json(wb.Sheets[nameD],{defval:'',raw:false});
      discosAll=rowsD.filter(r=>r['MOTORISTA']&&String(r['MOTORISTA']).trim()).map(r=>({
        motorista:String(r['MOTORISTA']||'').trim().toUpperCase(),
        placa:String(r['PLACA']||'').trim().toUpperCase(),
        contrato:String(r['CONTRATO']||'').trim().toUpperCase(),
        mes:String(r['MÊS']||'').trim().toUpperCase(),
        mesOrd:MES_ORD.indexOf(String(r['MÊS']||'').trim().toUpperCase()),
        quinzena:n(r['QUINZENA']),
        dataInicial:r['DATA INICIAL']||'',dataFinal:r['DATA FINAL']||'',
        filial:String(r['FILIAL']||'').trim().toUpperCase(),
        picos:n(r['QTDE PICOS']),dataReceb:r['DATA RECEBIMENTO']||'',
        obs:String(r['OBSERVAÇÃO']||'').trim(),
      }));

      const rowsV=XLSX.utils.sheet_to_json(wb.Sheets[nameV],{defval:'',raw:false});
      vdoAll=rowsV.filter(r=>r['MOTORISTA']&&String(r['MOTORISTA']).trim()).map(r=>({
        motorista:String(r['MOTORISTA']||'').trim().toUpperCase(),
        placa:String(r['PLACA']||'').trim().toUpperCase(),
        contrato:String(r['CONTRATO']||'').trim().toUpperCase(),
        mes:String(r['MÊS']||'').trim().toUpperCase(),
        mesOrd:MES_ORD.indexOf(String(r['MÊS']||'').trim().toUpperCase()),
        quinzena:n(r['QUINZENA']),
        dataInicial:r['DATA INICIAL']||'',dataFinal:r['DATA FINAL']||'',
        filial:String(r['FILIAL']||'').trim().toUpperCase(),
        picos:n(r['QTDE PICOS']),obs:String(r['OBSERVAÇÃO']||'').trim(),
      }));

      badge.textContent='✅ '+file.name+' — '+mensalAll.length+' mensal / '+discosAll.length+' discos / '+vdoAll.length+' VDO';
      g('dropZone').classList.add('hidden');
      g('dashboard').classList.add('visible');
      const hdrTabs=g('hdrTabs'); if(hdrTabs)hdrTabs.classList.add('visible');
      populateFilters();
      applyMensal(); applyDiscos(); applyVdo();
    }catch(err){badge.textContent='❌ '+err.message;alert('Erro: '+err.message);}
  };
  reader.readAsArrayBuffer(file);
}

// ── FILTROS ──
function fillSel(id,vals){
  const sel=g(id),cur=sel.value;
  while(sel.options.length>1)sel.remove(1);
  [...new Set(vals)].filter(Boolean).sort().forEach(v=>{const o=document.createElement('option');o.value=v;o.text=v;sel.appendChild(o);});
  sel.value=cur;
}
function fillSelOrdered(id,vals){
  const sel=g(id),cur=sel.value;
  while(sel.options.length>1)sel.remove(1);
  [...new Set(vals)].filter(Boolean).sort((a,b)=>MES_ORD.indexOf(a)-MES_ORD.indexOf(b)).forEach(v=>{const o=document.createElement('option');o.value=v;o.text=v;sel.appendChild(o);});
  sel.value=cur;
}
function populateFilters(){
  fillSelOrdered('mSelMes',mensalAll.map(r=>r.mes));
  fillSel('mSelMot',mensalAll.map(r=>r.motorista));
  fillSel('mSelCon',mensalAll.map(r=>r.contrato));
  fillSelOrdered('dSelMes',discosAll.map(r=>r.mes));
  fillSel('dSelMot',discosAll.map(r=>r.motorista));
  fillSel('dSelCon',discosAll.map(r=>r.contrato));
  fillSelOrdered('vSelMes',vdoAll.map(r=>r.mes));
  fillSel('vSelMot',vdoAll.map(r=>r.motorista));
  fillSel('vSelCon',vdoAll.map(r=>r.contrato));
}

window.applyMensal=function(){
  const mes=v('mSelMes'),mot=v('mSelMot'),con=v('mSelCon'),
        sta=v('mSelStatus'),placa=v('mSelPlaca').toUpperCase(),
        dIni=v('mDateIni'),dFim=v('mDateFim');
  const tIni=dIni?new Date(dIni).getTime():null;
  const tFim=dFim?new Date(dFim+'T23:59:59').getTime():null;
  mensalFilt=mensalAll.filter(r=>{
    if(mes&&r.mes!==mes)return false;
    if(mot&&r.motorista!==mot)return false;
    if(con&&r.contrato!==con)return false;
    if(sta&&r.statusPend!==sta)return false;
    if(placa&&!r.placa.includes(placa))return false;
    if(tIni||tFim){
      const mesIdx=r.mesOrd>=0?r.mesOrd:0;
      const ano=2026;
      const tMes=new Date(ano,mesIdx,1).getTime();
      const tMesFim=new Date(ano,mesIdx+1,0,23,59,59).getTime();
      if(tIni&&tMesFim<tIni)return false;
      if(tFim&&tMes>tFim)return false;
    }
    return true;
  });
  mPage=1; renderMKpis(); renderMCharts(); renderMTable();
};
window.resetMensal=function(){['mSelMes','mSelMot','mSelCon','mSelStatus'].forEach(i=>g(i).value='');['mSelPlaca','mTableSearch','mDateIni','mDateFim'].forEach(i=>{const el=g(i);if(el)el.value='';});mActiveKpi=null;applyMensal();};

window.applyDiscos=function(){
  const mes=v('dSelMes'),mot=v('dSelMot'),con=v('dSelCon'),
        quin=v('dSelQuin'),obs=v('dSelObs'),placa=v('dSelPlaca').toUpperCase(),
        dIni=v('dDateIni'),dFim=v('dDateFim');
  const tIni=dIni?new Date(dIni).getTime():null;
  const tFim=dFim?new Date(dFim+'T23:59:59').getTime():null;
  discosFilt=discosAll.filter(r=>{
    if(mes&&r.mes!==mes)return false;
    if(mot&&r.motorista!==mot)return false;
    if(con&&r.contrato!==con)return false;
    if(quin&&String(r.quinzena)!==quin)return false;
    if(placa&&!r.placa.includes(placa))return false;
    if(obs==='EXCESSO'&&!r.obs.includes('EXCESSO'))return false;
    if(obs==='CONFORME'&&!r.obs.toUpperCase().includes('CONFORME'))return false;
    if(obs==='FILIAL'&&!r.obs.toUpperCase().includes('FILIAL'))return false;
    if(tIni||tFim){
      const di=r.dataInicial?new Date(r.dataInicial).getTime():null;
      const df=r.dataFinal?new Date(r.dataFinal).getTime():null;
      const ref=di||df;
      if(ref){
        if(tIni&&ref<tIni)return false;
        if(tFim&&(di||df)>tFim)return false;
      }
    }
    return true;
  });
  dPage=1; renderDKpis(); renderDCharts(); renderDTable();
};
window.resetDiscos=function(){['dSelMes','dSelMot','dSelCon','dSelQuin','dSelObs'].forEach(i=>g(i).value='');['dSelPlaca','dTableSearch','dDateIni','dDateFim'].forEach(i=>{const el=g(i);if(el)el.value='';});dActiveKpi=null;applyDiscos();};

window.applyVdo=function(){
  const mes=v('vSelMes'),mot=v('vSelMot'),con=v('vSelCon'),
        quin=v('vSelQuin'),cp=v('vSelPicos'),placa=v('vSelPlaca').toUpperCase(),
        dIni=v('vDateIni'),dFim=v('vDateFim');
  const tIni=dIni?new Date(dIni).getTime():null;
  const tFim=dFim?new Date(dFim+'T23:59:59').getTime():null;
  vdoFilt=vdoAll.filter(r=>{
    if(mes&&r.mes!==mes)return false;
    if(mot&&r.motorista!==mot)return false;
    if(con&&r.contrato!==con)return false;
    if(quin&&String(r.quinzena)!==quin)return false;
    if(placa&&!r.placa.includes(placa))return false;
    if(cp==='sim'&&r.picos===0)return false;
    if(cp==='nao'&&r.picos>0)return false;
    if(tIni||tFim){
      const di=r.dataInicial?new Date(r.dataInicial).getTime():null;
      const df=r.dataFinal?new Date(r.dataFinal).getTime():null;
      const ref=di||df;
      if(ref){
        if(tIni&&ref<tIni)return false;
        if(tFim&&(di||df)>tFim)return false;
      }
    }
    return true;
  });
  vPage=1; renderVKpis(); renderVCharts(); renderVTable();
};
window.resetVdo=function(){['vSelMes','vSelMot','vSelCon','vSelQuin','vSelPicos'].forEach(i=>g(i).value='');['vSelPlaca','vTableSearch','vDateIni','vDateFim'].forEach(i=>{const el=g(i);if(el)el.value='';});vActiveKpi=null;applyVdo();};

// ── KPIs INTERATIVOS ──
function kpiCard(id,activeId,icon,color,label,val,sub,hint,onclick){
  const isActive = activeId===id;
  return `<div class="kpi-card${isActive?' kpi-active':''}" onclick="${onclick}" role="button" tabindex="0" style="--kpi-color:${color}">
    <div class="kpi-accent" style="background:${color}"></div>
    <div class="kpi-header">
      <svg class="kpi-icon" width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="${isActive?color:'currentColor'}" stroke-width="1.8">${icon}</svg>
      <div class="kpi-label">${label}</div>
    </div>
    <div class="kpi-value">${val}</div>
    <div class="kpi-sub" style="color:${color}">${sub}</div>
    <div class="kpi-hint">▸ ${hint}</div>
  </div>`;
}

const SVG_TRUCK   = '<rect x="1" y="3" width="15" height="13" rx="1"/><path d="M16 8h4l3 4v4h-7V8z"/><circle cx="5.5" cy="18.5" r="2.5"/><circle cx="18.5" cy="18.5" r="2.5"/>';
const SVG_CHECK   = '<path d="M22 11.08V12a10 10 0 11-5.93-9.14"/><polyline points="22 4 12 14.01 9 11.01"/>';
const SVG_WARN    = '<path d="M10.29 3.86L1.82 18a2 2 0 001.71 3h16.94a2 2 0 001.71-3L13.71 3.86a2 2 0 00-3.42 0z"/><line x1="12" y1="9" x2="12" y2="13"/><line x1="12" y1="17" x2="12.01" y2="17"/>';
const SVG_BOLT    = '<polygon points="13 2 3 14 12 14 11 22 21 10 12 10 13 2"/>';
const SVG_CHART   = '<line x1="18" y1="20" x2="18" y2="10"/><line x1="12" y1="20" x2="12" y2="4"/><line x1="6" y1="20" x2="6" y2="14"/>';
const SVG_DISC    = '<circle cx="12" cy="12" r="10"/><circle cx="12" cy="12" r="3"/>';
const SVG_EXCL    = '<circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/>';
const SVG_STAR    = '<polygon points="12 2 15.09 8.26 22 9.27 17 14.14 18.18 21.02 12 17.77 5.82 21.02 7 14.14 2 9.27 8.91 8.26 12 2"/>';
const SVG_LAYERS  = '<polygon points="12 2 2 7 12 12 22 7 12 2"/><polyline points="2 17 12 22 22 17"/><polyline points="2 12 12 17 22 12"/>';
const SVG_EYE     = '<path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"/><circle cx="12" cy="12" r="3"/>';

function renderMKpis(){
  const d=mensalFilt;
  if(!d.length){g('mKpiGrid').innerHTML='';g('mKpiDetail').innerHTML='';return;}
  const total=d.length;
  const entregues=d.filter(r=>r.statusPend==='ENTREGUES').length;
  const pend1=d.filter(r=>r.statusPend==='PENDÊNCIA 1ª QUIN').length;
  const pend2=d.filter(r=>r.statusPend==='PENDÊNCIA 2ª QUIN').length;
  const totalPend=pend1+pend2;
  const somaPicos=d.reduce((s,r)=>s+r.totalPicos,0);
  const somaEnt=d.reduce((s,r)=>s+r.quin1+r.quin2,0);
  const somaMeta=d.reduce((s,r)=>s+r.meta,0);
  const taxa=total?Math.round(entregues/total*100):0;
  const meses=[...new Set(d.map(r=>r.mes))].length;

  // Pendências Agregado: contrato AGREGADO (ou TERCEIRO) com pendência
  const pendAgrRows=d.filter(r=>r.statusPend&&r.statusPend!=='ENTREGUES'&&r.contrato!=='FROTA');
  const pendAgrUniq=[...new Set(pendAgrRows.map(r=>r.placa))].length;
  const pendAgrMot=[...new Set(pendAgrRows.map(r=>r.motorista))].length;

  // Pendências Frota: contrato FROTA com pendência
  const pendFrotaRows=d.filter(r=>r.statusPend&&r.statusPend!=='ENTREGUES'&&r.contrato==='FROTA');
  const pendFrotaUniq=[...new Set(pendFrotaRows.map(r=>r.placa))].length;
  const pendFrotaMot=[...new Set(pendFrotaRows.map(r=>r.motorista))].length;

  g('mKpiGrid').innerHTML=[
    kpiCard('total',      mActiveKpi,SVG_TRUCK, '#6C63FF','Registros',         total,                    meses+' mês(es)',                     'Detalhar tudo',      "mClickKpi('total')"),
    kpiCard('entregue',   mActiveKpi,SVG_CHECK, '#1D9E75','Entregues',         entregues,                taxa+'% do total',                    'Ver por mês',        "mClickKpi('entregue')"),
    kpiCard('pend-agr',   mActiveKpi,SVG_WARN,  '#E24B4A','Pendência Agregado',    pendAgrUniq+' placas',    pendAgrMot+' motoristas',             'Ver na aba Discos',  "mClickKpi('pend-agr')"),
    kpiCard('pend-frota', mActiveKpi,SVG_WARN,  '#D4537E','Pendência Frota',       pendFrotaUniq+' placas',  pendFrotaMot+' motoristas',           'Ver na aba VDO',     "mClickKpi('pend-frota')"),
    kpiCard('picos',      mActiveKpi,SVG_BOLT,  '#BA7517','Total picos',       somaPicos.toLocaleString('pt-BR'),'eventos acumulados',         'Ranking picos',      "mClickKpi('picos')"),
    kpiCard('taxa',       mActiveKpi,SVG_CHART, '#378ADD','Taxa entrega',      taxa+'%',                 'meta: 100%',                         'Evolução mensal',    "mClickKpi('taxa')"),
    kpiCard('discos',     mActiveKpi,SVG_DISC,  '#7F77DD','Discos entregues',  somaEnt.toLocaleString('pt-BR'),'meta: '+somaMeta.toLocaleString('pt-BR'),'Ver déficit',"mClickKpi('discos')"),
  ].join('');
  renderMKpiDetail(d,{total,entregues,totalPend,pend1,pend2,somaPicos,somaEnt,somaMeta,taxa,meses,pendAgrRows,pendAgrUniq,pendAgrMot,pendFrotaRows,pendFrotaUniq,pendFrotaMot});
}

window.mFilterPend=function(quin){
  window._mPendFilter = window._mPendFilter===quin ? 'all' : quin;
  renderMKpiDetail(mensalFilt, (function(){
    const d=mensalFilt;
    const total=d.length;
    const entregues=d.filter(r=>r.statusPend==='ENTREGUES').length;
    const pend1=d.filter(r=>r.statusPend==='PENDÊNCIA 1ª QUIN').length;
    const pend2=d.filter(r=>r.statusPend==='PENDÊNCIA 2ª QUIN').length;
    const totalPend=pend1+pend2;
    const somaPicos=d.reduce((s,r)=>s+r.totalPicos,0);
    const somaEnt=d.reduce((s,r)=>s+r.quin1+r.quin2,0);
    const somaMeta=d.reduce((s,r)=>s+r.meta,0);
    const taxa=total?Math.round(entregues/total*100):0;
    const meses=[...new Set(d.map(r=>r.mes))].length;
    const pendAgrRows=d.filter(r=>r.statusPend&&r.statusPend!=='ENTREGUES'&&r.contrato!=='FROTA');
    const pendAgrUniq=[...new Set(pendAgrRows.map(r=>r.placa))].length;
    const pendAgrMot=[...new Set(pendAgrRows.map(r=>r.motorista))].length;
    const pendFrotaRows=d.filter(r=>r.statusPend&&r.statusPend!=='ENTREGUES'&&r.contrato==='FROTA');
    const pendFrotaUniq=[...new Set(pendFrotaRows.map(r=>r.placa))].length;
    const pendFrotaMot=[...new Set(pendFrotaRows.map(r=>r.motorista))].length;
    return {total,entregues,totalPend,pend1,pend2,somaPicos,somaEnt,somaMeta,taxa,meses,pendAgrRows,pendAgrUniq,pendAgrMot,pendFrotaRows,pendFrotaUniq,pendFrotaMot};
  })());
};

window.mClickKpi=function(id){
  window._mPendFilter='all';
  mActiveKpi = mActiveKpi===id ? null : id;
  const chartsArea=g('mChartsArea');
  if(mActiveKpi){chartsArea.style.display='none';}else{chartsArea.style.display='block';}
  renderMKpis();
};

function renderMKpiDetail(d,stats){
  const dp=g('mKpiDetail');
  if(!mActiveKpi){dp.innerHTML='';return;}
  const ca=g('mChartsArea'); if(ca)ca.style.display='none';
  const titles={total:'Todos os registros',entregue:'Discos entregues por mês','pend-agr':'Pendências Agregado — base: Discos','pend-frota':'Pendências Frota — base: Leitura VDO',picos:'Ranking de picos',taxa:'Evolução da taxa de entrega',discos:'Discos entregues vs meta'};
  let body='';

  if(mActiveKpi==='total'){
    const cMap={};d.forEach(r=>{if(r.contrato)cMap[r.contrato]=(cMap[r.contrato]||0)+1;});
    const cK=Object.keys(cMap).sort((a,b)=>cMap[b]-cMap[a]);
    body=miniKpis([{l:'Total',v:d.length},{l:'Motoristas',v:[...new Set(d.map(r=>r.motorista))].length},{l:'Placas',v:[...new Set(d.map(r=>r.placa))].length},{l:'Meses',v:stats.meses}])+
    '<div class="card-title" style="margin-bottom:8px">Por contrato</div>'+
    barRows(cK,k=>cMap[k],Math.max(...Object.values(cMap)),'#6C63FF');
  }
  else if(mActiveKpi==='entregue'){
    const meses=[...new Set(d.sort((a,b)=>a.mesOrd-b.mesOrd).map(r=>r.mes))];
    const entM=meses.map(m=>d.filter(r=>r.mes===m&&r.statusPend==='ENTREGUES').length);
    const maxE=Math.max(...entM)||1;
    body=miniKpis([{l:'Total entregues',v:stats.entregues},{l:'Taxa',v:stats.taxa+'%'},{l:'Melhor mês',v:meses[entM.indexOf(Math.max(...entM))]||'—'},{l:'Média/mês',v:Math.round(stats.entregues/stats.meses)}])+
    '<div class="card-title" style="margin-bottom:8px">Entregues por mês</div>'+
    meses.map((m,i)=>`<div class="bar-row"><span class="bar-lbl">${m}</span><div class="bar-track"><div class="bar-fill" style="width:${Math.round(entM[i]/maxE*100)}%;background:#1D9E75"></div></div><span class="bar-num">${entM[i]}</span></div>`).join('');
  }
  else if(mActiveKpi==='pend-agr'){
    const allPendMensal=d.filter(r=>r.statusPend&&r.statusPend!=='ENTREGUES'&&r.contrato!=='FROTA');
    const mPendFilter=window._mPendFilter||'all';
    const filteredMensal=mPendFilter==='q1'?allPendMensal.filter(r=>r.statusPend.includes('1ª'))
                        :mPendFilter==='q2'?allPendMensal.filter(r=>r.statusPend.includes('2ª'))
                        :allPendMensal;
    const pend1=allPendMensal.filter(r=>r.statusPend.includes('1ª')).length;
    const pend2=allPendMensal.filter(r=>r.statusPend.includes('2ª')).length;
    const q1Active=mPendFilter==='q1', q2Active=mPendFilter==='q2';
    const rows=filteredMensal.map(r=>{
      const disco=discosAll.find(d=>d.placa===r.placa&&d.mes===r.mes)||{};
      return {...r, picosDiscos:disco.picos||0, dataReceb:disco.dataReceb||'—'};
    });
    body=`<div class="detail-kpis">
      <div class="detail-kpi">
        <div class="dk-label">Total pendente</div>
        <div class="dk-value">${allPendMensal.length}</div>
      </div>
      <div class="detail-kpi">
        <div class="dk-label">Motoristas</div>
        <div class="dk-value">${[...new Set(allPendMensal.map(r=>r.motorista))].length}</div>
      </div>
      <div class="detail-kpi pend-quin-card${q1Active?' pqc-active':''}" onclick="mFilterPend('q1')" style="cursor:pointer;border-color:${q1Active?'#BA7517':'var(--bd)'};background:${q1Active?'rgba(186,117,23,.1)':'var(--sf2)'};--kpi-color:#BA7517">
        <div class="dk-label" style="color:#BA7517">1ª Quinzena</div>
        <div class="dk-value" style="color:#BA7517">${pend1}</div>
        <div class="dk-sub" style="color:#BA7517;font-size:10px;margin-top:2px">${q1Active?'▸ clique para ver todos':'▸ clique para filtrar'}</div>
      </div>
      <div class="detail-kpi pend-quin-card${q2Active?' pqc-active':''}" onclick="mFilterPend('q2')" style="cursor:pointer;border-color:${q2Active?'#E24B4A':'var(--bd)'};background:${q2Active?'rgba(226,75,74,.1)':'var(--sf2)'};--kpi-color:#E24B4A">
        <div class="dk-label" style="color:#E24B4A">2ª Quinzena</div>
        <div class="dk-value" style="color:#E24B4A">${pend2}</div>
        <div class="dk-sub" style="color:#E24B4A;font-size:10px;margin-top:2px">${q2Active?'▸ clique para ver todos':'▸ clique para filtrar'}</div>
      </div>
    </div>
    <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:8px">
      <div class="card-title" style="margin:0">${mPendFilter==='q1'?'Pendências 1ª Quinzena':mPendFilter==='q2'?'Pendências 2ª Quinzena':'Todas as pendências'} — Agregado <span style="font-family:var(--font-mono);font-size:10px;color:var(--tx3)">(${rows.length})</span></div>
      ${mPendFilter!=='all'?`<span style="font-size:11px;color:var(--acc);cursor:pointer" onclick="mFilterPend('all')">✕ limpar filtro</span>`:''}
    </div>
    <div class="table-wrap" style="${rows.length>25?'max-height:'+Math.round(25*34+34)+'px;overflow-y:auto;':''}"><table class="modal-table">
      <thead><tr><th>Motorista</th><th>Placa</th><th>Mês</th><th>Status</th><th>Picos (Discos)</th><th>Recebimento</th><th>Total Picos</th></tr></thead>
      <tbody>${rows.map(r=>`<tr>
        <td>${r.motorista}</td>
        <td class="vc">${r.placa}</td>
        <td>${r.mes}</td>
        <td><span class="badge ${r.statusPend.includes('1ª')?'ba':'br'}" style="border-color:${r.statusPend.includes('1ª')?'#BA7517':'#E24B4A'};color:${r.statusPend.includes('1ª')?'#BA7517':'#E24B4A'};background:${r.statusPend.includes('1ª')?'rgba(186,117,23,.12)':'rgba(226,75,74,.12)'}">${r.statusPend}</span></td>
        <td class="${r.picosDiscos>0?'vr':'vc'}">${r.picosDiscos}</td>
        <td>${r.dataReceb}</td>
        <td class="${r.totalPicos>0?'vr':'vc'}">${r.totalPicos}</td>
      </tr>`).join('')}</tbody>
    </table></div>`;
  }
  else if(mActiveKpi==='pend-frota'){
    const allPendMensal=d.filter(r=>r.statusPend&&r.statusPend!=='ENTREGUES'&&r.contrato==='FROTA');
    const mPendFilter=window._mPendFilter||'all';
    const filteredMensal=mPendFilter==='q1'?allPendMensal.filter(r=>r.statusPend.includes('1ª'))
                        :mPendFilter==='q2'?allPendMensal.filter(r=>r.statusPend.includes('2ª'))
                        :allPendMensal;
    const pend1=allPendMensal.filter(r=>r.statusPend.includes('1ª')).length;
    const pend2=allPendMensal.filter(r=>r.statusPend.includes('2ª')).length;
    const q1Active=mPendFilter==='q1', q2Active=mPendFilter==='q2';
    const rows=filteredMensal.map(r=>{
      const vdo=vdoAll.find(v=>v.placa===r.placa&&v.mes===r.mes)||{};
      return {...r, picosVdo:vdo.picos||0, obsVdo:vdo.obs||'—'};
    });
    body=`<div class="detail-kpis">
      <div class="detail-kpi">
        <div class="dk-label">Total pendente</div>
        <div class="dk-value">${allPendMensal.length}</div>
      </div>
      <div class="detail-kpi">
        <div class="dk-label">Motoristas</div>
        <div class="dk-value">${[...new Set(allPendMensal.map(r=>r.motorista))].length}</div>
      </div>
      <div class="detail-kpi pend-quin-card${q1Active?' pqc-active':''}" onclick="mFilterPend('q1')" style="cursor:pointer;border-color:${q1Active?'#BA7517':'var(--bd)'};background:${q1Active?'rgba(186,117,23,.1)':'var(--sf2)'};--kpi-color:#BA7517">
        <div class="dk-label" style="color:#BA7517">1ª Quinzena</div>
        <div class="dk-value" style="color:#BA7517">${pend1}</div>
        <div class="dk-sub" style="color:#BA7517;font-size:10px;margin-top:2px">${q1Active?'▸ clique para ver todos':'▸ clique para filtrar'}</div>
      </div>
      <div class="detail-kpi pend-quin-card${q2Active?' pqc-active':''}" onclick="mFilterPend('q2')" style="cursor:pointer;border-color:${q2Active?'#E24B4A':'var(--bd)'};background:${q2Active?'rgba(226,75,74,.1)':'var(--sf2)'};--kpi-color:#E24B4A">
        <div class="dk-label" style="color:#E24B4A">2ª Quinzena</div>
        <div class="dk-value" style="color:#E24B4A">${pend2}</div>
        <div class="dk-sub" style="color:#E24B4A;font-size:10px;margin-top:2px">${q2Active?'▸ clique para ver todos':'▸ clique para filtrar'}</div>
      </div>
    </div>
    <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:8px">
      <div class="card-title" style="margin:0">${mPendFilter==='q1'?'Pendências 1ª Quinzena':mPendFilter==='q2'?'Pendências 2ª Quinzena':'Todas as pendências'} — Frota <span style="font-family:var(--font-mono);font-size:10px;color:var(--tx3)">(${rows.length})</span></div>
      ${mPendFilter!=='all'?`<span style="font-size:11px;color:var(--acc);cursor:pointer" onclick="mFilterPend('all')">✕ limpar filtro</span>`:''}
    </div>
    <div class="table-wrap" style="${rows.length>25?'max-height:'+Math.round(25*34+34)+'px;overflow-y:auto;':''}"><table class="modal-table">
      <thead><tr><th>Motorista</th><th>Placa</th><th>Mês</th><th>Status</th><th>Picos (VDO)</th><th>Obs VDO</th><th>Total Picos</th></tr></thead>
      <tbody>${rows.map(r=>`<tr>
        <td>${r.motorista}</td>
        <td class="vc">${r.placa}</td>
        <td>${r.mes}</td>
        <td><span class="badge ${r.statusPend.includes('1ª')?'ba':'br'}" style="border-color:${r.statusPend.includes('1ª')?'#BA7517':'#E24B4A'};color:${r.statusPend.includes('1ª')?'#BA7517':'#E24B4A'};background:${r.statusPend.includes('1ª')?'rgba(186,117,23,.12)':'rgba(226,75,74,.12)'}">${r.statusPend}</span></td>
        <td class="${r.picosVdo>0?'vr':'vc'}">${r.picosVdo}</td>
        <td style="max-width:160px;overflow:hidden;text-overflow:ellipsis;font-size:11px">${r.obsVdo}</td>
        <td class="${r.totalPicos>0?'vr':'vc'}">${r.totalPicos}</td>
      </tr>`).join('')}</tbody>
    </table></div>`;
  }
  else if(mActiveKpi==='picos'){
    const mMap={};d.forEach(r=>{if(r.motorista)mMap[r.motorista]=(mMap[r.motorista]||0)+r.totalPicos;});
    const top=Object.entries(mMap).sort((a,b)=>b[1]-a[1]).slice(0,15);
    const maxV=top[0]?top[0][1]:1;
    body=miniKpis([{l:'Total picos',v:stats.somaPicos.toLocaleString('pt-BR')},{l:'Motoristas',v:[...new Set(d.map(r=>r.motorista))].length},{l:'Média/motorista',v:Math.round(stats.somaPicos/([...new Set(d.map(r=>r.motorista))].length||1))},{l:'Meses',v:stats.meses}])+
    '<div class="card-title" style="margin-bottom:8px">Top motoristas — picos</div>'+
    top.map(([nome,val])=>`<div class="bar-row"><span class="bar-lbl">${nome}</span><div class="bar-track"><div class="bar-fill" style="width:${Math.round(val/maxV*100)}%;background:#E24B4A;opacity:${0.5+val/maxV*.5}"></div></div><span class="bar-num" style="color:var(--red)">${val.toLocaleString('pt-BR')}</span></div>`).join('');
  }
  else if(mActiveKpi==='taxa'){
    const meses=[...new Set(d.sort((a,b)=>a.mesOrd-b.mesOrd).map(r=>r.mes))];
    const taxM=meses.map(m=>{const md=d.filter(r=>r.mes===m);const tot=md.length;return tot?Math.round(md.filter(r=>r.statusPend==='ENTREGUES').length/tot*100):0;});
    body=miniKpis([{l:'Taxa atual',v:(taxM[taxM.length-1]||0)+'%'},{l:'Melhor mês',v:Math.max(...taxM)+'%'},{l:'Pior mês',v:Math.min(...taxM)+'%'},{l:'Média geral',v:Math.round(taxM.reduce((a,b)=>a+b,0)/(taxM.length||1))+'%'}])+
    '<div class="card-title" style="margin-bottom:8px">Taxa por mês</div>'+
    meses.map((m,i)=>`<div class="bar-row"><span class="bar-lbl">${m}</span><div class="bar-track"><div class="bar-fill" style="width:${taxM[i]}%;background:${taxM[i]>=90?'#1D9E75':taxM[i]>=70?'#BA7517':'#E24B4A'}"></div></div><span class="bar-num" style="color:${taxM[i]>=90?'var(--grn)':taxM[i]>=70?'var(--amb)':'var(--red)'}">${taxM[i]}%</span></div>`).join('');
  }
  else if(mActiveKpi==='discos'){
    const meses=[...new Set(d.sort((a,b)=>a.mesOrd-b.mesOrd).map(r=>r.mes))];
    const entM=meses.map(m=>d.filter(r=>r.mes===m).reduce((s,r)=>s+r.quin1+r.quin2,0));
    const metM=meses.map(m=>d.filter(r=>r.mes===m).reduce((s,r)=>s+r.meta,0));
    const maxV=Math.max(...metM,1);
    const deficit=stats.somaEnt-stats.somaMeta;
    body=miniKpis([{l:'Total entregues',v:stats.somaEnt.toLocaleString('pt-BR')},{l:'Meta total',v:stats.somaMeta.toLocaleString('pt-BR')},{l:'Déficit',v:deficit>=0?'+'+deficit:String(deficit)},{l:'Média/placa',v:(stats.somaEnt/([...new Set(d.map(r=>r.placa))].length||1)).toFixed(1)}])+
    '<div class="card-title" style="margin-bottom:8px">Entregues vs meta por mês</div>'+
    meses.map((m,i)=>`<div class="bar-row"><span class="bar-lbl">${m}</span><div class="bar-track" style="position:relative"><div class="bar-fill" style="width:${Math.round(entM[i]/maxV*100)}%;background:#6C63FF;opacity:.85"></div></div><span class="bar-num">${entM[i]} / ${metM[i]}</span></div>`).join('');
  }

  dp.innerHTML=`<div class="kpi-detail">
    <div class="kpi-detail-header">
      <span class="kpi-detail-title">${titles[mActiveKpi]||''}</span>
      <span class="kpi-back" onclick="mClickKpi('${mActiveKpi}')">← Fechar</span>
    </div>${body}</div>`;
}

function renderDKpis(){
  const d=discosFilt;
  if(!d.length){g('dKpiGrid').innerHTML='';g('dKpiDetail').innerHTML='';return;}
  const total=d.length;
  const exc=d.filter(r=>r.obs.toUpperCase().includes('EXCESSO')).length;
  const conf=d.filter(r=>r.obs.toUpperCase().includes('CONFORME')).length;
  const fil=d.filter(r=>r.obs.toUpperCase().includes('FILIAL')).length;
  const somaPicos=d.reduce((s,r)=>s+r.picos,0);
  const meses=[...new Set(d.map(r=>r.mes))].length;
  const mMap={};d.forEach(r=>{if(r.motorista)mMap[r.motorista]=(mMap[r.motorista]||0)+r.picos;});
  const topMot=Object.entries(mMap).sort((a,b)=>b[1]-a[1])[0];
  g('dKpiGrid').innerHTML=[
    kpiCard('total', dActiveKpi,SVG_DISC,  '#6C63FF','Discos analisados',total,       meses+' mês(es)',          'Ver todos',       "dClickKpi('total')"),
    kpiCard('exc',   dActiveKpi,SVG_EXCL,  '#E24B4A','Com excesso',      exc,         Math.round(exc/total*100)+'% dos discos','Detalhar',  "dClickKpi('exc')"),
    kpiCard('conf',  dActiveKpi,SVG_CHECK, '#1D9E75','Conformes',        conf,        Math.round(conf/total*100)+'% dos discos','Detalhar', "dClickKpi('conf')"),
    kpiCard('fil',   dActiveKpi,SVG_LAYERS,'#BA7517','Entregue filial',  fil,         Math.round(fil/total*100)+'% dos discos','Detalhar',  "dClickKpi('fil')"),
    kpiCard('picos', dActiveKpi,SVG_BOLT,  '#378ADD','Total picos',      somaPicos.toLocaleString('pt-BR'),'eventos','Ranking motoristas',"dClickKpi('picos')"),
    kpiCard('top',   dActiveKpi,SVG_STAR,  '#D4537E','Mais picos',       topMot?topMot[0]:'—', topMot?topMot[1].toLocaleString('pt-BR')+' picos':'—','Ver detalhes', "dClickKpi('top')"),
  ].join('');
  renderDKpiDetail(d,{total,exc,conf,fil,somaPicos,topMot,meses});
}

window.dClickKpi=function(id){
  dActiveKpi = dActiveKpi===id ? null : id;
  const ca=g('dChartsArea');
  if(dActiveKpi){ca.style.display='none';}else{ca.style.display='block';}
  renderDKpis();
};

function renderDKpiDetail(d,stats){
  const dp=g('dKpiDetail');
  if(!dActiveKpi){dp.innerHTML='';return;}
  const titles={total:'Todos os discos analisados',exc:'Discos com excesso de velocidade',conf:'Discos conformes',fil:'Entregues na filial',picos:'Ranking de picos por motorista',top:'Detalhes do motorista'};
  let body='';
  const rows = dActiveKpi==='total'?d : dActiveKpi==='exc'?d.filter(r=>r.obs.toUpperCase().includes('EXCESSO')) : dActiveKpi==='conf'?d.filter(r=>r.obs.toUpperCase().includes('CONFORME')) : dActiveKpi==='fil'?d.filter(r=>r.obs.toUpperCase().includes('FILIAL')) : d;

  if(dActiveKpi==='picos'||dActiveKpi==='top'){
    const mMap={};d.forEach(r=>{if(r.motorista)mMap[r.motorista]=(mMap[r.motorista]||0)+r.picos;});
    const top=Object.entries(mMap).sort((a,b)=>b[1]-a[1]).slice(0,15);
    const maxV=top[0]?top[0][1]:1;
    body=miniKpis([{l:'Total picos',v:stats.somaPicos.toLocaleString('pt-BR')},{l:'Motoristas',v:[...new Set(d.map(r=>r.motorista))].length},{l:'Maior pico único',v:Math.max(...d.map(r=>r.picos)).toLocaleString('pt-BR')},{l:'Média/disco',v:(stats.somaPicos/stats.total).toFixed(1)}])+
    '<div class="card-title" style="margin-bottom:8px">Top motoristas — picos</div>'+
    top.map(([nome,val])=>`<div class="bar-row"><span class="bar-lbl">${nome}</span><div class="bar-track"><div class="bar-fill" style="width:${Math.round(val/maxV*100)}%;background:#E24B4A;opacity:${0.5+val/maxV*.5}"></div></div><span class="bar-num" style="color:var(--red)">${val.toLocaleString('pt-BR')}</span></div>`).join('');
  } else {
    body=miniKpis([{l:'Registros',v:rows.length},{l:'Motoristas',v:[...new Set(rows.map(r=>r.motorista))].length},{l:'Total picos',v:rows.reduce((s,r)=>s+r.picos,0).toLocaleString('pt-BR')},{l:'Média picos',v:(rows.reduce((s,r)=>s+r.picos,0)/(rows.length||1)).toFixed(1)}])+
    `<div class="table-wrap" style="max-height:280px;overflow-y:auto"><table class="modal-table">
      <thead><tr><th>Motorista</th><th>Placa</th><th>Mês</th><th>Quin</th><th>Picos</th><th>Observação</th></tr></thead>
      <tbody>${rows.slice(0,50).map(r=>`<tr><td>${r.motorista}</td><td class="vc">${r.placa}</td><td>${r.mes}</td><td>${r.quinzena}ª</td><td class="${r.picos>0?'vr':'vc'}">${r.picos}</td><td style="max-width:160px;overflow:hidden;text-overflow:ellipsis">${r.obs||'—'}</td></tr>`).join('')}</tbody>
    </table></div>`;
  }
  dp.innerHTML=`<div class="kpi-detail"><div class="kpi-detail-header"><span class="kpi-detail-title">${titles[dActiveKpi]||''}</span><span class="kpi-back" onclick="dClickKpi('${dActiveKpi}')">← Fechar</span></div>${body}</div>`;
}

function renderVKpis(){
  const d=vdoFilt;
  if(!d.length){g('vKpiGrid').innerHTML='';g('vKpiDetail').innerHTML='';return;}
  const total=d.length;
  const comPicos=d.filter(r=>r.picos>0).length;
  const semPicos=d.filter(r=>r.picos===0).length;
  const somaPicos=d.reduce((s,r)=>s+r.picos,0);
  const meses=[...new Set(d.map(r=>r.mes))].length;
  const mMap={};d.forEach(r=>{if(r.motorista)mMap[r.motorista]=(mMap[r.motorista]||0)+r.picos;});
  const topMot=Object.entries(mMap).sort((a,b)=>b[1]-a[1])[0];
  const pMap={};d.forEach(r=>{if(r.placa)pMap[r.placa]=(pMap[r.placa]||0)+r.picos;});
  const topPlaca=Object.entries(pMap).sort((a,b)=>b[1]-a[1])[0];
  g('vKpiGrid').innerHTML=[
    kpiCard('total',  vActiveKpi,SVG_EYE,  '#6C63FF','Leituras VDO',  total,        meses+' mês(es)',          'Ver todas',       "vClickKpi('total')"),
    kpiCard('picos',  vActiveKpi,SVG_BOLT, '#E24B4A','Total picos',   somaPicos.toLocaleString('pt-BR'),'eventos acumulados','Ranking', "vClickKpi('picos')"),
    kpiCard('com',    vActiveKpi,SVG_WARN, '#BA7517','Com picos',     comPicos,     Math.round(comPicos/total*100)+'% das leituras','Detalhar', "vClickKpi('com')"),
    kpiCard('sem',    vActiveKpi,SVG_CHECK,'#1D9E75','Sem picos',     semPicos,     Math.round(semPicos/total*100)+'% das leituras','Detalhar', "vClickKpi('sem')"),
    kpiCard('topMot', vActiveKpi,SVG_STAR, '#D4537E','Mais picos',   topMot?topMot[0]:'—',topMot?topMot[1].toLocaleString('pt-BR')+' picos':'—','Detalhes', "vClickKpi('topMot')"),
    kpiCard('topPlaca',vActiveKpi,SVG_TRUCK,'#378ADD','Placa crítica',topPlaca?topPlaca[0]:'—',topPlaca?topPlaca[1].toLocaleString('pt-BR')+' picos':'—','Detalhes',"vClickKpi('topPlaca')"),
  ].join('');
  renderVKpiDetail(d,{total,comPicos,semPicos,somaPicos,meses,topMot,topPlaca});
}

window.vClickKpi=function(id){
  vActiveKpi = vActiveKpi===id ? null : id;
  const ca=g('vChartsArea');
  if(vActiveKpi){ca.style.display='none';}else{ca.style.display='block';}
  renderVKpis();
};

function renderVKpiDetail(d,stats){
  const dp=g('vKpiDetail');
  if(!vActiveKpi){dp.innerHTML='';return;}
  const titles={total:'Todas as leituras VDO',picos:'Ranking de picos',com:'Leituras com picos',sem:'Leituras sem picos',topMot:'Motorista com mais picos',topPlaca:'Placa com mais picos'};
  let body='';
  const rows=vActiveKpi==='total'?d:vActiveKpi==='com'?d.filter(r=>r.picos>0):vActiveKpi==='sem'?d.filter(r=>r.picos===0):d;

  if(vActiveKpi==='picos'||vActiveKpi==='topMot'){
    const mMap={};d.forEach(r=>{if(r.motorista)mMap[r.motorista]=(mMap[r.motorista]||0)+r.picos;});
    const top=Object.entries(mMap).sort((a,b)=>b[1]-a[1]).slice(0,15);
    const maxV=top[0]?top[0][1]:1;
    body=miniKpis([{l:'Total picos',v:stats.somaPicos.toLocaleString('pt-BR')},{l:'Motoristas',v:[...new Set(d.map(r=>r.motorista))].length},{l:'Média/leitura',v:(stats.somaPicos/stats.total).toFixed(1)},{l:'Meses',v:stats.meses}])+
    '<div class="card-title" style="margin-bottom:8px">Top motoristas — picos</div>'+
    top.map(([nome,val])=>`<div class="bar-row"><span class="bar-lbl">${nome}</span><div class="bar-track"><div class="bar-fill" style="width:${Math.round(val/maxV*100)}%;background:#E24B4A;opacity:${0.5+val/maxV*.5}"></div></div><span class="bar-num" style="color:var(--red)">${val.toLocaleString('pt-BR')}</span></div>`).join('');
  } else if(vActiveKpi==='topPlaca'){
    const pMap={};d.forEach(r=>{if(r.placa)pMap[r.placa]=(pMap[r.placa]||0)+r.picos;});
    const top=Object.entries(pMap).sort((a,b)=>b[1]-a[1]).slice(0,15);
    const maxV=top[0]?top[0][1]:1;
    body=miniKpis([{l:'Total picos',v:stats.somaPicos.toLocaleString('pt-BR')},{l:'Placas únicas',v:[...new Set(d.map(r=>r.placa))].length},{l:'Maior placa',v:top[0]?top[0][0]:'—'},{l:'Maior valor',v:top[0]?top[0][1].toLocaleString('pt-BR'):'—'}])+
    '<div class="card-title" style="margin-bottom:8px">Top placas — picos</div>'+
    top.map(([placa,val])=>`<div class="bar-row"><span class="bar-lbl">${placa}</span><div class="bar-track"><div class="bar-fill" style="width:${Math.round(val/maxV*100)}%;background:#378ADD;opacity:${0.5+val/maxV*.5}"></div></div><span class="bar-num">${val.toLocaleString('pt-BR')}</span></div>`).join('');
  } else {
    body=miniKpis([{l:'Registros',v:rows.length},{l:'Motoristas',v:[...new Set(rows.map(r=>r.motorista))].length},{l:'Total picos',v:rows.reduce((s,r)=>s+r.picos,0).toLocaleString('pt-BR')},{l:'Média picos',v:(rows.reduce((s,r)=>s+r.picos,0)/(rows.length||1)).toFixed(1)}])+
    `<div class="table-wrap" style="max-height:280px;overflow-y:auto"><table class="modal-table">
      <thead><tr><th>Motorista</th><th>Placa</th><th>Mês</th><th>Quin</th><th>Picos</th><th>Observação</th></tr></thead>
      <tbody>${rows.slice(0,50).map(r=>`<tr><td>${r.motorista}</td><td class="vc">${r.placa}</td><td>${r.mes}</td><td>${r.quinzena}ª</td><td class="${r.picos>0?'vr':'vc'}">${r.picos}</td><td style="max-width:160px;overflow:hidden;text-overflow:ellipsis">${r.obs||'—'}</td></tr>`).join('')}</tbody>
    </table></div>`;
  }
  dp.innerHTML=`<div class="kpi-detail"><div class="kpi-detail-header"><span class="kpi-detail-title">${titles[vActiveKpi]||''}</span><span class="kpi-back" onclick="vClickKpi('${vActiveKpi}')">← Fechar</span></div>${body}</div>`;
}

// ── HELPERS ──
function miniKpis(items){
  return `<div class="detail-kpis">${items.map(i=>`<div class="detail-kpi"><div class="dk-label">${i.l}</div><div class="dk-value">${i.v}</div>${i.s?`<div class="dk-sub">${i.s}</div>`:''}</div>`).join('')}</div>`;
}
function barRows(keys,valFn,maxV,color){
  return keys.map(k=>`<div class="bar-row"><span class="bar-lbl">${k}</span><div class="bar-track"><div class="bar-fill" style="width:${Math.round(valFn(k)/maxV*100)}%;background:${color};opacity:.8"></div></div><span class="bar-num">${valFn(k)}</span></div>`).join('');
}

// ── CHARTS ──
function dc(id){if(CH[id]){CH[id].destroy();delete CH[id];}}
function mesOrdered(d){return[...new Set(d.slice().sort((a,b)=>a.mesOrd-b.mesOrd).map(r=>r.mes))];}
function mkScales(y2=false){
  const sc={x:{ticks:{color:tickC(),font:{size:11}},grid:{color:gridC()}},y:{ticks:{color:tickC(),font:{size:9}},grid:{color:gridC()},beginAtZero:true}};
  if(y2)sc.y2={position:'right',ticks:{color:'#E24B4A',font:{size:9}},grid:{display:false},beginAtZero:true};
  return sc;
}
function mkDonut(id,labels,vals,colors,onClick){
  dc(id);
  CH[id]=new Chart(g(id),{type:'doughnut',
    data:{labels,datasets:[{data:vals,backgroundColor:colors,borderWidth:2,borderColor:'transparent',hoverOffset:8}]},
    options:{responsive:true,maintainAspectRatio:false,cutout:'62%',
      plugins:{legend:{display:false},tooltip:{callbacks:{label:ctx=>ctx.label+': '+ctx.raw.toLocaleString('pt-BR')}}},
      onClick:(_,el)=>{if(el.length&&onClick)onClick(el[0].index);}}});
}

function renderMCharts(){
  const d=mensalFilt; const ms=mesOrdered(d);
  const q1Mes=ms.map(m=>d.filter(r=>r.mes===m).reduce((s,r)=>s+r.quin1,0));
  const q2Mes=ms.map(m=>d.filter(r=>r.mes===m).reduce((s,r)=>s+r.quin2,0));
  const metaMes=ms.map(m=>d.filter(r=>r.mes===m).reduce((s,r)=>s+r.meta,0));
  dc('mCMensal');
  CH['mCMensal']=new Chart(g('mCMensal'),{data:{labels:ms,datasets:[
    {type:'bar',label:'1ª Quinzena',data:q1Mes,backgroundColor:'rgba(108,99,255,.85)',borderColor:'#6C63FF',borderWidth:1,borderRadius:4,yAxisID:'y'},
    {type:'bar',label:'2ª Quinzena',data:q2Mes,backgroundColor:'rgba(127,119,221,.6)',borderColor:'#7F77DD',borderWidth:1,borderRadius:4,yAxisID:'y'},
    {type:'line',label:'Meta',data:metaMes,borderColor:'#BA7517',borderDash:[5,3],pointBackgroundColor:'#BA7517',pointRadius:4,tension:.3,fill:false,yAxisID:'y'}
  ]},options:{responsive:true,maintainAspectRatio:false,interaction:{mode:'index',intersect:false},
    plugins:{legend:{display:true,labels:{color:legC(),boxWidth:11,font:{size:11}}},tooltip:{callbacks:{label:ctx=>ctx.dataset.label+': '+ctx.raw.toLocaleString('pt-BR')}}},
    scales:mkScales(),
    onClick:(_,el)=>{if(!el.length)return;const mes=ms[el[0].index];const rows=d.filter(r=>r.mes===mes);openModalRows('📅 '+mes+' — Mensal',mTheads,rows.map(mRow),[{l:'Registros',v:rows.length},{l:'1ª Quin',v:rows.reduce((s,r)=>s+r.quin1,0)},{l:'2ª Quin',v:rows.reduce((s,r)=>s+r.quin2,0)},{l:'Meta',v:rows.reduce((s,r)=>s+r.meta,0)}]);}}});
  const mMap={};d.forEach(r=>{if(r.motorista)mMap[r.motorista]=(mMap[r.motorista]||0)+r.quin1+r.quin2;});
  const motK=Object.keys(mMap).sort((a,b)=>mMap[b]-mMap[a]).slice(0,15),motV=motK.map(k=>mMap[k]);
  dc('mCMotoristas');if(motK.length){
    g('mCMotoristas').parentElement.style.height=Math.max(motK.length*22+60,180)+'px';
    CH['mCMotoristas']=new Chart(g('mCMotoristas'),{type:'bar',data:{labels:motK,datasets:[{data:motV,backgroundColor:motK.map((_,i)=>COLORS[i%COLORS.length]+'bb'),borderColor:motK.map((_,i)=>COLORS[i%COLORS.length]),borderWidth:1,borderRadius:3}]},
      options:{indexAxis:'y',responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false},tooltip:{callbacks:{label:ctx=>ctx.raw+' discos'}}},
        scales:{x:{ticks:{color:tickC(),font:{size:9}},grid:{color:gridC()}},y:{ticks:{color:legC(),font:{size:10}}}},
        onClick:(_,el)=>{if(el.length)openModalMMotorista(motK[el[0].index]);}}});}
  const ent=d.filter(r=>r.statusPend==='ENTREGUES').length;
  const p1=d.filter(r=>r.statusPend==='PENDÊNCIA 1ª QUIN').length;
  const p2=d.filter(r=>r.statusPend==='PENDÊNCIA 2ª QUIN').length;
  g('mLgStatus').innerHTML=['Entregues','Pend 1ª','Pend 2ª'].map((l,i)=>`<span><span class="ld" style="background:${['#1D9E75','#BA7517','#E24B4A'][i]}"></span>${l}</span>`).join('');
  mkDonut('mCStatus',['Entregues','Pend 1ª','Pend 2ª'],[ent,p1,p2],['#1D9E75','#BA7517','#E24B4A'],idx=>{ if(idx===0)mClickKpi('entregue'); else mClickKpi('pendente'); });
  const cMap={};d.forEach(r=>{if(r.contrato)cMap[r.contrato]=(cMap[r.contrato]||0)+1;});
  const cK=Object.keys(cMap).sort((a,b)=>cMap[b]-cMap[a]),cC=cK.map((_,i)=>COLORS[i%COLORS.length]);
  g('mLgContrato').innerHTML=cK.map((k,i)=>`<span><span class="ld" style="background:${cC[i]}"></span>${k}</span>`).join('');
  mkDonut('mCContrato',cK,cK.map(k=>cMap[k]),cC);
  const fMap={};d.forEach(r=>{if(r.filial)fMap[r.filial]=(fMap[r.filial]||0)+1;});
  const fK=Object.keys(fMap).sort((a,b)=>fMap[b]-fMap[a]),fC=fK.map((_,i)=>COLORS[(i+3)%COLORS.length]);
  g('mLgFilial').innerHTML=fK.map((k,i)=>`<span><span class="ld" style="background:${fC[i]}"></span>${k}</span>`).join('');
  mkDonut('mCFilial',fK,fK.map(k=>fMap[k]),fC);
  const picMes=ms.map(m=>d.filter(r=>r.mes===m).reduce((s,r)=>s+r.totalPicos,0));
  dc('mCPicos');
  CH['mCPicos']=new Chart(g('mCPicos'),{type:'bar',data:{labels:ms,datasets:[{data:picMes,backgroundColor:'rgba(226,75,74,.65)',borderColor:'#E24B4A',borderWidth:1,borderRadius:4}]},
    options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false},tooltip:{callbacks:{label:ctx=>ctx.raw.toLocaleString('pt-BR')+' picos'}}},scales:mkScales()}});
}

function renderDCharts(){
  const d=discosFilt; const ms=mesOrdered(d);
  const q1Mes=ms.map(m=>d.filter(r=>r.mes===m&&r.quinzena===1).length);
  const q2Mes=ms.map(m=>d.filter(r=>r.mes===m&&r.quinzena===2).length);
  const picMes=ms.map(m=>d.filter(r=>r.mes===m).reduce((s,r)=>s+r.picos,0));
  dc('dCMensal');
  CH['dCMensal']=new Chart(g('dCMensal'),{data:{labels:ms,datasets:[
    {type:'bar',label:'1ª Quinzena',data:q1Mes,backgroundColor:'rgba(108,99,255,.85)',borderColor:'#6C63FF',borderWidth:1,borderRadius:4,yAxisID:'y'},
    {type:'bar',label:'2ª Quinzena',data:q2Mes,backgroundColor:'rgba(127,119,221,.6)',borderColor:'#7F77DD',borderWidth:1,borderRadius:4,yAxisID:'y'},
    {type:'line',label:'Picos',data:picMes,borderColor:'#E24B4A',backgroundColor:'rgba(226,75,74,.07)',pointBackgroundColor:'#E24B4A',pointRadius:4,tension:.3,fill:true,yAxisID:'y2'}
  ]},options:{responsive:true,maintainAspectRatio:false,interaction:{mode:'index',intersect:false},
    plugins:{legend:{display:true,labels:{color:legC(),boxWidth:11,font:{size:11}}},tooltip:{callbacks:{label:ctx=>ctx.dataset.label+': '+ctx.raw.toLocaleString('pt-BR')}}},
    scales:mkScales(true),
    onClick:(_,el)=>{if(!el.length)return;const mes=ms[el[0].index];const rows=d.filter(r=>r.mes===mes);openModalRows('📅 '+mes+' — Discos',dTheads,rows.map(dRow),[{l:'Registros',v:rows.length},{l:'1ª Quin',v:rows.filter(r=>r.quinzena===1).length},{l:'2ª Quin',v:rows.filter(r=>r.quinzena===2).length},{l:'Total picos',v:rows.reduce((s,r)=>s+r.picos,0).toLocaleString('pt-BR')}]);}}});
  const mMap={};d.forEach(r=>{if(r.motorista)mMap[r.motorista]=(mMap[r.motorista]||0)+r.picos;});
  const motK=Object.keys(mMap).sort((a,b)=>mMap[b]-mMap[a]).slice(0,15),motV=motK.map(k=>mMap[k]);
  dc('dCTop');if(motK.length){
    g('dCTop').parentElement.style.height=Math.max(motK.length*22+60,180)+'px';
    CH['dCTop']=new Chart(g('dCTop'),{type:'bar',data:{labels:motK,datasets:[{data:motV,backgroundColor:motK.map((_,i)=>COLORS[i%COLORS.length]+'bb'),borderColor:motK.map((_,i)=>COLORS[i%COLORS.length]),borderWidth:1,borderRadius:3}]},
      options:{indexAxis:'y',responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false},tooltip:{callbacks:{label:ctx=>ctx.raw.toLocaleString('pt-BR')+' picos'}}},
        scales:{x:{ticks:{color:tickC(),font:{size:9}},grid:{color:gridC()}},y:{ticks:{color:legC(),font:{size:10}}}},
        onClick:(_,el)=>{if(el.length)openModalDMotorista(motK[el[0].index]);}}});}
  const exc=d.filter(r=>r.obs.toUpperCase().includes('EXCESSO')).length;
  const conf=d.filter(r=>r.obs.toUpperCase().includes('CONFORME')).length;
  const fil=d.filter(r=>r.obs.toUpperCase().includes('FILIAL')).length;
  const out=d.length-exc-conf-fil;
  const oL=['Excesso','Conforme','Filial',...(out?['Outros']:[])];
  const oV=[exc,conf,fil,...(out?[out]:[])];
  const oC=['#E24B4A','#1D9E75','#BA7517','#7F77DD'];
  g('dLgObs').innerHTML=oL.map((l,i)=>`<span><span class="ld" style="background:${oC[i]}"></span>${l}</span>`).join('');
  mkDonut('dCObs',oL,oV,oC,idx=>{ dClickKpi(['exc','conf','fil','total'][idx]||'total'); });
  const cMap={};d.forEach(r=>{if(r.contrato)cMap[r.contrato]=(cMap[r.contrato]||0)+1;});
  const cK=Object.keys(cMap).sort((a,b)=>cMap[b]-cMap[a]),cC=cK.map((_,i)=>COLORS[i%COLORS.length]);
  g('dLgContrato').innerHTML=cK.map((k,i)=>`<span><span class="ld" style="background:${cC[i]}"></span>${k}</span>`).join('');
  mkDonut('dCContrato',cK,cK.map(k=>cMap[k]),cC);
  const fMap={};d.forEach(r=>{if(r.filial)fMap[r.filial]=(fMap[r.filial]||0)+1;});
  const fK=Object.keys(fMap).sort((a,b)=>fMap[b]-fMap[a]),fC=fK.map((_,i)=>COLORS[(i+3)%COLORS.length]);
  g('dLgFilial').innerHTML=fK.map((k,i)=>`<span><span class="ld" style="background:${fC[i]}"></span>${k}</span>`).join('');
  mkDonut('dCFilial',fK,fK.map(k=>fMap[k]),fC);
  const q1=ms.map(m=>d.filter(r=>r.mes===m&&r.quinzena===1).reduce((s,r)=>s+r.picos,0));
  const q2=ms.map(m=>d.filter(r=>r.mes===m&&r.quinzena===2).reduce((s,r)=>s+r.picos,0));
  dc('dCQuin');
  CH['dCQuin']=new Chart(g('dCQuin'),{type:'bar',data:{labels:ms,datasets:[{label:'1ª Quin',data:q1,backgroundColor:'rgba(108,99,255,.7)',borderRadius:3},{label:'2ª Quin',data:q2,backgroundColor:'rgba(55,138,221,.7)',borderRadius:3}]},
    options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:true,labels:{color:legC(),boxWidth:10,font:{size:10}}}},scales:mkScales()}});
}

function renderVCharts(){
  const d=vdoFilt; const ms=mesOrdered(d);
  const q1Mes=ms.map(m=>d.filter(r=>r.mes===m&&r.quinzena===1).length);
  const q2Mes=ms.map(m=>d.filter(r=>r.mes===m&&r.quinzena===2).length);
  const picMes=ms.map(m=>d.filter(r=>r.mes===m).reduce((s,r)=>s+r.picos,0));
  dc('vCMensal');
  CH['vCMensal']=new Chart(g('vCMensal'),{data:{labels:ms,datasets:[
    {type:'bar',label:'1ª Quinzena',data:q1Mes,backgroundColor:'rgba(55,138,221,.85)',borderColor:'#378ADD',borderWidth:1,borderRadius:4,yAxisID:'y'},
    {type:'bar',label:'2ª Quinzena',data:q2Mes,backgroundColor:'rgba(133,183,235,.6)',borderColor:'#85B7EB',borderWidth:1,borderRadius:4,yAxisID:'y'},
    {type:'line',label:'Picos',data:picMes,borderColor:'#E24B4A',backgroundColor:'rgba(226,75,74,.07)',pointBackgroundColor:'#E24B4A',pointRadius:4,tension:.3,fill:true,yAxisID:'y2'}
  ]},options:{responsive:true,maintainAspectRatio:false,interaction:{mode:'index',intersect:false},
    plugins:{legend:{display:true,labels:{color:legC(),boxWidth:11,font:{size:11}}},tooltip:{callbacks:{label:ctx=>ctx.dataset.label+': '+ctx.raw.toLocaleString('pt-BR')}}},
    scales:mkScales(true),
    onClick:(_,el)=>{if(!el.length)return;const mes=ms[el[0].index];const rows=d.filter(r=>r.mes===mes);openModalRows('📅 '+mes+' — VDO',vTheads,rows.map(vRow),[{l:'Registros',v:rows.length},{l:'1ª Quin',v:rows.filter(r=>r.quinzena===1).length},{l:'2ª Quin',v:rows.filter(r=>r.quinzena===2).length},{l:'Total picos',v:rows.reduce((s,r)=>s+r.picos,0).toLocaleString('pt-BR')}]);}}});
  const mMap={};d.forEach(r=>{if(r.motorista)mMap[r.motorista]=(mMap[r.motorista]||0)+r.picos;});
  const motK=Object.keys(mMap).sort((a,b)=>mMap[b]-mMap[a]).slice(0,15),motV=motK.map(k=>mMap[k]);
  dc('vCTop');if(motK.length){
    g('vCTop').parentElement.style.height=Math.max(motK.length*22+60,180)+'px';
    CH['vCTop']=new Chart(g('vCTop'),{type:'bar',data:{labels:motK,datasets:[{data:motV,backgroundColor:motK.map((_,i)=>COLORS[i%COLORS.length]+'bb'),borderColor:motK.map((_,i)=>COLORS[i%COLORS.length]),borderWidth:1,borderRadius:3}]},
      options:{indexAxis:'y',responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false},tooltip:{callbacks:{label:ctx=>ctx.raw.toLocaleString('pt-BR')+' picos'}}},
        scales:{x:{ticks:{color:tickC(),font:{size:9}},grid:{color:gridC()}},y:{ticks:{color:legC(),font:{size:10}}}},
        onClick:(_,el)=>{if(el.length)openModalVMotorista(motK[el[0].index]);}}});}
  const cp=d.filter(r=>r.picos>0).length,sp=d.filter(r=>r.picos===0).length;
  g('vLgDonut').innerHTML=['Com picos','Sem picos'].map((l,i)=>`<span><span class="ld" style="background:${['#E24B4A','#1D9E75'][i]}"></span>${l}</span>`).join('');
  mkDonut('vCDonut',['Com picos','Sem picos'],[cp,sp],['#E24B4A','#1D9E75'],idx=>vClickKpi(idx===0?'com':'sem'));
  const fMap={};d.forEach(r=>{if(r.filial)fMap[r.filial]=(fMap[r.filial]||0)+1;});
  const fK=Object.keys(fMap).sort((a,b)=>fMap[b]-fMap[a]),fC=fK.map((_,i)=>COLORS[(i+3)%COLORS.length]);
  g('vLgFilial').innerHTML=fK.map((k,i)=>`<span><span class="ld" style="background:${fC[i]}"></span>${k}</span>`).join('');
  mkDonut('vCFilial',fK,fK.map(k=>fMap[k]),fC);
  const q1=ms.map(m=>d.filter(r=>r.mes===m&&r.quinzena===1).reduce((s,r)=>s+r.picos,0));
  const q2=ms.map(m=>d.filter(r=>r.mes===m&&r.quinzena===2).reduce((s,r)=>s+r.picos,0));
  dc('vCQuin');
  CH['vCQuin']=new Chart(g('vCQuin'),{type:'bar',data:{labels:ms,datasets:[{label:'1ª Quin',data:q1,backgroundColor:'rgba(55,138,221,.65)',borderRadius:3},{label:'2ª Quin',data:q2,backgroundColor:'rgba(139,124,246,.65)',borderRadius:3}]},
    options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:true,labels:{color:legC(),boxWidth:10,font:{size:10}}}},scales:mkScales()}});
  const pMap={};d.forEach(r=>{if(r.placa)pMap[r.placa]=(pMap[r.placa]||0)+r.picos;});
  const pK=Object.keys(pMap).sort((a,b)=>pMap[b]-pMap[a]).slice(0,10),pV=pK.map(k=>pMap[k]);
  dc('vCPlacas');
  if(pK.length){
    CH['vCPlacas']=new Chart(g('vCPlacas'),{type:'bar',data:{labels:pK,datasets:[{data:pV,backgroundColor:'rgba(55,138,221,.7)',borderRadius:3}]},
      options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false},tooltip:{callbacks:{label:ctx=>ctx.raw.toLocaleString('pt-BR')+' picos'}}},
        scales:{x:{ticks:{color:tickC(),font:{size:10},maxRotation:35},grid:{color:gridC()}},y:{ticks:{color:tickC(),font:{size:9}},grid:{color:gridC()},beginAtZero:true}}}});
  }
}

// ── TABELAS ──
const mTheads=['Mês','Motorista','Placa','Contrato','Filial','1ª Quin','2ª Quin','Meta','Pendência','Status','Picos','Obs'];
const mRow=r=>`<tr><td>${r.mes}</td><td>${r.motorista}</td><td class="vc">${r.placa}</td><td>${r.contrato}</td><td>${r.filial}</td><td class="vc">${r.quin1}</td><td class="vc">${r.quin2}</td><td class="vc">${r.meta}</td><td class="${r.pendencia<0?'vr':r.pendencia>0?'vc':''}">${r.pendencia}</td><td><span class="badge ${r.statusPend==='ENTREGUES'?'bg':r.statusPend.includes('1ª')?'ba':'br'}">${r.statusPend||'—'}</span></td><td class="${r.totalPicos>0?'vr':'vc'}">${r.totalPicos.toLocaleString('pt-BR')}</td><td style="max-width:160px;overflow:hidden;text-overflow:ellipsis;font-size:11px">${r.obs||'—'}</td></tr>`;
const dTheads=['Motorista','Placa','Contrato','Mês','Quin','Data Inicial','Data Final','Filial','Picos','Recebimento','Observação'];
const dRow=r=>`<tr><td>${r.motorista}</td><td class="vc">${r.placa}</td><td>${r.contrato}</td><td>${r.mes}</td><td class="vc">${r.quinzena}ª</td><td>${fmt(r.dataInicial)}</td><td>${fmt(r.dataFinal)}</td><td>${r.filial}</td><td class="${r.picos>0?'vr':'vc'}">${r.picos.toLocaleString('pt-BR')}</td><td>${fmt(r.dataReceb)}</td><td style="max-width:180px;overflow:hidden;text-overflow:ellipsis;font-size:11px">${r.obs||'—'}</td></tr>`;
const vTheads=['Motorista','Placa','Contrato','Mês','Quin','Data Inicial','Data Final','Filial','Picos','Observação'];
const vRow=r=>`<tr><td>${r.motorista}</td><td class="vc">${r.placa}</td><td>${r.contrato}</td><td>${r.mes}</td><td class="vc">${r.quinzena}ª</td><td>${fmt(r.dataInicial)}</td><td>${fmt(r.dataFinal)}</td><td>${r.filial}</td><td class="${r.picos>0?'vr':'vc'}">${r.picos.toLocaleString('pt-BR')}</td><td style="max-width:180px;overflow:hidden;text-overflow:ellipsis;font-size:11px">${r.obs||'—'}</td></tr>`;

function renderMTable(){sortArr(mensalFilt,mSortCol,mSortDir);renderTable('mTbody','mCount','mPag',mensalFilt,mPage,p=>{mPage=p;renderMTable();},mRow);}
function renderDTable(){sortArr(discosFilt,dSortCol,dSortDir);renderTable('dTbody','dCount','dPag',discosFilt,dPage,p=>{dPage=p;renderDTable();},dRow);}
function renderVTable(){sortArr(vdoFilt,vSortCol,vSortDir);renderTable('vTbody','vCount','vPag',vdoFilt,vPage,p=>{vPage=p;renderVTable();},vRow);}

function renderTable(tbId,cntId,pagId,data,cur,goFn,rowFn){
  const total=data.length,pages=Math.ceil(total/PS),start=(cur-1)*PS;
  g(cntId).textContent=total.toLocaleString('pt-BR')+' registros';
  g(tbId).innerHTML=data.slice(start,start+PS).map(rowFn).join('');
  renderPag(pagId,cur,pages,total,start,goFn);
}
function renderPag(pgId,cur,pages,total,start,goFn){
  const pg=g(pgId);
  if(pages<=1){pg.innerHTML=`<span>${total} registro${total!==1?'s':''}</span>`;return;}
  let btns='',sp=[],prev=0;
  for(let i=1;i<=pages;i++)if(i===1||i===pages||Math.abs(i-cur)<=2)sp.push(i);
  sp.forEach(p=>{if(prev&&p-prev>1)btns+=`<button class="pgb" disabled>…</button>`;btns+=`<button class="pgb${p===cur?' active':''}" onclick="(${goFn.toString()})(${p})">${p}</button>`;prev=p;});
  pg.innerHTML=`<span>${start+1}–${Math.min(start+PS,total)} de ${total}</span><div class="pg-btns"><button class="pgb" onclick="(${goFn.toString()})(${cur-1})" ${cur===1?'disabled':''}>‹</button>${btns}<button class="pgb" onclick="(${goFn.toString()})(${cur+1})" ${cur===pages?'disabled':''}>›</button></div>`;
}

// ── SORT ──
const numCols=['quin1','quin2','meta','pendencia','totalPicos','picos1','picos2','quinzena','picos'];
function sortArr(arr,col,dir){arr.sort((a,b)=>{let va=a[col],vb=b[col];if(col==='mes'){va=a.mesOrd;vb=b.mesOrd;}else if(numCols.includes(col)){va=Number(va)||0;vb=Number(vb)||0;}else{va=String(va||'').toLowerCase();vb=String(vb||'').toLowerCase();}return va<vb?(dir==='asc'?-1:1):va>vb?(dir==='asc'?1:-1):0;});}
function mkSort(colVar,dirVar,colsArr,tSel,applyFn){
  return function(col){
    if(window[colVar]===col)window[dirVar]=window[dirVar]==='asc'?'desc':'asc';
    else{window[colVar]=col;window[dirVar]=numCols.includes(col)?'desc':'asc';}
    document.querySelectorAll(tSel+' thead th').forEach(th=>th.classList.remove('sa','sd'));
    const idx=colsArr.indexOf(col);if(idx>=0){const ths=document.querySelectorAll(tSel+' thead th');if(ths[idx])ths[idx].classList.add(window[dirVar]==='asc'?'sa':'sd');}
    applyFn();
  };
}
const mSort=mkSort('mSortCol','mSortDir',['mes','motorista','placa','contrato','filial','quin1','quin2','meta','pendencia','statusPend','totalPicos','obs'],'#tab-mensal',()=>{mPage=1;renderMTable();});
const dSort=mkSort('dSortCol','dSortDir',['motorista','placa','contrato','mes','quinzena','dataInicial','dataFinal','filial','picos','dataReceb','obs'],'#tab-discos',()=>{dPage=1;renderDTable();});
const vSort=mkSort('vSortCol','vSortDir',['motorista','placa','contrato','mes','quinzena','dataInicial','dataFinal','filial','picos','obs'],'#tab-vdo',()=>{vPage=1;renderVTable();});

// ── MODAL ──
function openModal(title,metrics,theads,rows){
  g('modalTitle').textContent=title;
  const mm=g('modalMetrics');
  if(metrics&&metrics.length){mm.innerHTML=metrics.map(m=>`<div class="modal-metric"><div class="mm-label">${m.l}</div><div class="mm-val">${m.v}</div></div>`).join('');mm.style.display='grid';}
  else{mm.innerHTML='';mm.style.display='none';}
  g('modalThead').innerHTML='<tr>'+theads.map(c=>`<th>${c}</th>`).join('')+'</tr>';
  g('modalTbody').innerHTML=rows.slice(0,300).join('');
  g('modalOverlay').classList.add('open');
}
window.closeModal=function(e){if(!e||e.target===g('modalOverlay'))g('modalOverlay').classList.remove('open');};
function openModalRows(title,theads,rows,metrics){openModal(title,metrics||[],theads,rows);}
function openModalMMotorista(nome){const rows=mensalFilt.filter(r=>r.motorista===nome);openModal('🚛 '+nome,[{l:'Registros',v:rows.length},{l:'Discos',v:rows.reduce((s,r)=>s+r.quin1+r.quin2,0)},{l:'Picos',v:rows.reduce((s,r)=>s+r.totalPicos,0).toLocaleString('pt-BR')},{l:'Meses',v:[...new Set(rows.map(r=>r.mes))].join(', ')}],mTheads,rows.map(mRow));}
function openModalDMotorista(nome){const rows=discosFilt.filter(r=>r.motorista===nome);openModal('🚛 '+nome,[{l:'Registros',v:rows.length},{l:'Total picos',v:rows.reduce((s,r)=>s+r.picos,0).toLocaleString('pt-BR')},{l:'Excessos',v:rows.filter(r=>r.obs.includes('EXCESSO')).length},{l:'Meses',v:[...new Set(rows.map(r=>r.mes))].join(', ')}],dTheads,rows.map(dRow));}
function openModalVMotorista(nome){const rows=vdoFilt.filter(r=>r.motorista===nome);openModal('🚛 '+nome,[{l:'Leituras',v:rows.length},{l:'Total picos',v:rows.reduce((s,r)=>s+r.picos,0).toLocaleString('pt-BR')},{l:'Meses',v:[...new Set(rows.map(r=>r.mes))].join(', ')}],vTheads,rows.map(vRow));}

// ── EXPORTAR ──
window.exportCSV=function(data,filename,headers){
  const rows=data.map(r=>headers.map((_,i)=>`"${String(Object.values(r)[i]||'').replace(/"/g,'""')}"`).join(','));
  dl(filename,'\uFEFF'+headers.join(',')+'\n'+rows.join('\n'),'text/csv;charset=utf-8');
};
window.exportXLSX=function(data,filename,sheetname,headers){
  const ws=XLSX.utils.json_to_sheet(data.map(r=>{const o={};headers.forEach((h,i)=>{o[h]=Object.values(r)[i];});return o;}));
  const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,sheetname);XLSX.writeFile(wb,filename);
};
function dl(name,content,type){const a=document.createElement('a');a.href=URL.createObjectURL(new Blob([content],{type}));a.download=name;a.click();}

// ── EXPÕE FUNÇÕES PARA O HTML ──
window.switchTab = switchTab;
window.loadFile = loadFile;
window.applyMensal = applyMensal;
window.resetMensal = resetMensal;
window.applyDiscos = applyDiscos;
window.resetDiscos = resetDiscos;
window.applyVdo = applyVdo;
window.resetVdo = resetVdo;
window.mSort = mSort;
window.dSort = dSort;
window.vSort = vSort;
window.exportCSV = exportCSV;
window.exportXLSX = exportXLSX;
window.mClickKpi = mClickKpi;
window.dClickKpi = dClickKpi;
window.vClickKpi = vClickKpi;

}); // end DOMContentLoaded
