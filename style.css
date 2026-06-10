@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&family=DM+Mono:wght@400;500&display=swap');

/* ── TEMA ESCURO ── */
:root{
  --bg:#0d1117;--sf:#161b22;--sf2:#1c2333;--sf3:#21293b;--sf4:#2a3447;
  --bd:#2a3447;--bd2:#344059;--bd3:#3f4f6e;
  --tx:#e6edf3;--tx2:#8b949e;--tx3:#6e7681;--tx4:#3d444d;
  --acc:#6C63FF;--acc2:#534AB7;--acc3:#EEEDFE;
  --red:#E24B4A;--red2:#FCEBEB;--red3:#791F1F;
  --grn:#1D9E75;--grn2:#E1F5EE;--grn3:#085041;
  --amb:#BA7517;--amb2:#FAEEDA;--amb3:#633806;
  --blu:#378ADD;--blu2:#E6F1FB;--blu3:#0C447C;
  --pur:#7F77DD;--pur2:#EEEDFE;--pur3:#3C3489;
  --hdr:#0d1117;--nav:#0a0e16;
  --r:8px;--rl:12px;--rxl:16px;
  --shadow:rgba(0,0,0,.5);
}
html.light{
  --bg:#f0f2f5;--sf:#ffffff;--sf2:#f6f8fa;--sf3:#eef0f3;--sf4:#e2e5ea;
  --bd:#d0d7de;--bd2:#b8c0cc;--bd3:#9aa3b0;
  --tx:#1c2333;--tx2:#57606a;--tx3:#8c959f;--tx4:#c6cbd6;
  --acc:#6C63FF;--acc2:#534AB7;--acc3:#EEEDFE;
  --hdr:#1a1f36;--nav:#13172b;
  --shadow:rgba(0,0,0,.1);
}

*{box-sizing:border-box;margin:0;padding:0;}
html{scroll-behavior:smooth;}
body{font-family:'Inter',sans-serif;background:var(--bg);color:var(--tx);min-height:100vh;font-size:14px;line-height:1.5;}
::-webkit-scrollbar{width:5px;height:5px;}
::-webkit-scrollbar-track{background:var(--sf);}
::-webkit-scrollbar-thumb{background:var(--bd3);border-radius:3px;}

/* ── HEADER ── */
header{background:var(--hdr);border-bottom:1px solid rgba(255,255,255,.06);padding:0;position:sticky;top:0;z-index:200;display:flex;flex-direction:column;}
.hdr-top{display:flex;align-items:center;justify-content:space-between;padding:0 1.75rem;height:68px;}
.logo-wrap{display:flex;align-items:center;gap:12px;}
.logo-icon{width:32px;height:32px;background:var(--acc);border-radius:8px;display:flex;align-items:center;justify-content:center;font-size:16px;font-weight:700;color:#fff;flex-shrink:0;box-shadow:0 0 10px rgba(108,99,255,.55),0 0 22px rgba(108,99,255,.25);transition:box-shadow .3s;}
.logo-icon:hover{box-shadow:0 0 16px rgba(108,99,255,.9),0 0 38px rgba(108,99,255,.5),0 0 60px rgba(108,99,255,.2);}
.logo-text{font-size:14px;font-weight:700;color:#fff;letter-spacing:.08em;line-height:1;}
.logo-dot{color:var(--acc);}
.logo-sub{font-size:10px;color:rgba(255,255,255,.35);font-weight:400;letter-spacing:.06em;text-transform:uppercase;margin-top:2px;}
.hdr-right{display:flex;align-items:center;gap:8px;}
.file-badge{font-size:11px;color:rgba(255,255,255,.45);font-family:'DM Mono',monospace;background:rgba(255,255,255,.07);border:1px solid rgba(255,255,255,.1);border-radius:6px;padding:4px 10px;max-width:260px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;}
.btn-upload{display:flex;align-items:center;gap:6px;padding:7px 16px;background:var(--acc);color:#fff;border:none;border-radius:var(--r);font-size:12px;font-family:'Inter',sans-serif;font-weight:600;cursor:pointer;transition:all .2s;letter-spacing:.02em;}
.btn-upload:hover{background:var(--acc2);transform:translateY(-1px);}
.btn-upload svg{flex-shrink:0;}
#fileInput{display:none;}
.theme-toggle{display:flex;align-items:center;justify-content:center;width:32px;height:32px;border-radius:7px;border:1px solid rgba(255,255,255,.15);background:rgba(255,255,255,.07);cursor:pointer;color:#fff;flex-shrink:0;transition:all .2s;}
.theme-toggle:hover{background:rgba(255,255,255,.15);}
.theme-toggle svg{width:16px;height:16px;fill:none;stroke:currentColor;stroke-width:2;}
.icon-moon{display:block;} .icon-sun{display:none;}
html.light .icon-moon{display:none;} html.light .icon-sun{display:block;}

/* ── MAIN ── */
.main{padding:1.5rem 1.75rem;max-width:1600px;margin:0 auto;}

/* ── DROP ZONE ── */
.drop-zone{border:2px dashed var(--bd3);border-radius:var(--rxl);padding:5rem 2rem;text-align:center;background:var(--sf);transition:all .2s;cursor:pointer;margin-bottom:1.5rem;}
.drop-zone.drag-over{border-color:var(--acc);background:var(--sf2);}
.drop-zone.hidden{display:none;}
.drop-icon{width:68px;height:68px;background:var(--sf2);border-radius:16px;display:flex;align-items:center;justify-content:center;margin:0 auto 1.25rem;border:1px solid var(--bd);}
.drop-icon svg{width:32px;height:32px;stroke:var(--acc);fill:none;stroke-width:1.5;}
.drop-title{font-size:20px;font-weight:700;margin-bottom:.4rem;letter-spacing:-.02em;}
.drop-sub{font-size:13px;color:var(--tx2);}
.drop-abas{display:flex;gap:8px;justify-content:center;margin-top:.85rem;flex-wrap:wrap;}
.drop-aba{font-size:11px;font-family:'DM Mono',monospace;background:var(--sf2);border:1px solid var(--bd);border-radius:6px;padding:4px 10px;color:var(--tx2);}

#dashboard{display:none;}
#dashboard.visible{display:block;}

/* ── TABS no HEADER ── */
.hdr-tabs{display:none;gap:0;padding:0 1.25rem;border-top:1px solid rgba(255,255,255,.05);}
.hdr-tabs.visible{display:flex;}
.tab{padding:11px 20px;font-size:13px;font-weight:500;color:rgba(255,255,255,.4);cursor:pointer;border-bottom:2px solid transparent;transition:all .2s;white-space:nowrap;position:relative;}
.tab.active{color:#fff;border-color:var(--acc);}
.tab:hover:not(.active){color:rgba(255,255,255,.75);text-shadow:0 0 12px rgba(108,99,255,.6),0 0 28px rgba(108,99,255,.3);}
.tab.active{text-shadow:0 0 10px rgba(108,99,255,.5);}
.tab-panel{display:none;}
.tab-panel.active{display:block;}

/* ── SECTION LABEL ── */
.slabel{font-size:10px;font-weight:700;color:var(--tx3);text-transform:uppercase;letter-spacing:.12em;margin-bottom:.75rem;display:flex;align-items:center;gap:10px;}
.slabel::after{content:'';flex:1;height:1px;background:var(--bd);}

/* ── FILTROS ── */
.controls{background:var(--sf);border:1px solid var(--bd);border-radius:var(--rl);padding:1rem 1.25rem;margin-bottom:1.25rem;display:flex;flex-wrap:wrap;gap:10px;align-items:flex-end;}
.cg{display:flex;flex-direction:column;gap:3px;}
.cg label{font-size:10px;font-weight:600;color:var(--tx3);text-transform:uppercase;letter-spacing:.07em;}
select,input[type=text]{padding:6px 10px;border-radius:var(--r);border:1px solid var(--bd2);background:var(--sf2);color:var(--tx);font-family:'Inter',sans-serif;font-size:12px;outline:none;min-width:120px;transition:border-color .15s;}
select:focus,input[type=text]:focus{border-color:var(--acc);}
.btn-reset{padding:7px 14px;border-radius:var(--r);border:1px solid var(--bd2);background:transparent;color:var(--tx2);font-family:'Inter',sans-serif;font-size:12px;cursor:pointer;transition:all .15s;white-space:nowrap;}
.btn-reset:hover{background:var(--sf2);color:var(--tx);}

/* ── KPI GRID ── */
.kpi-grid{display:grid;grid-template-columns:repeat(7,minmax(0,1fr));gap:8px;margin-bottom:1.25rem;}
.kpi-card{background:var(--sf);border:1px solid var(--bd);border-radius:var(--rl);padding:14px 15px;cursor:pointer;transition:border-color .25s,transform .2s,box-shadow .25s;position:relative;overflow:hidden;}
.kpi-card:hover{border-color:var(--kpi-color,var(--acc));transform:translateY(-3px);box-shadow:0 0 0 1px var(--kpi-color,var(--acc)),0 4px 20px color-mix(in srgb,var(--kpi-color,var(--acc)) 35%,transparent),0 8px 32px color-mix(in srgb,var(--kpi-color,var(--acc)) 15%,transparent);}
.kpi-card.kpi-active{border-color:var(--kpi-color,var(--acc));background:var(--sf2);box-shadow:0 0 0 1px var(--kpi-color,var(--acc)),0 4px 24px color-mix(in srgb,var(--kpi-color,var(--acc)) 40%,transparent);}
.kpi-accent{position:absolute;left:0;top:0;bottom:0;width:3px;}
.kpi-header{display:flex;align-items:center;gap:8px;margin-bottom:8px;}
.kpi-icon{font-size:22px;color:var(--tx3);flex-shrink:0;transition:color .2s;}
.kpi-card:hover .kpi-icon,.kpi-card.kpi-active .kpi-icon{color:var(--kpi-color,var(--acc));}
.kpi-label{font-size:11px;font-weight:700;color:var(--tx2);text-transform:uppercase;letter-spacing:.06em;line-height:1.2;}
.kpi-value{font-size:22px;font-weight:700;color:var(--tx);letter-spacing:-.02em;line-height:1.1;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}
.kpi-sub{font-size:11px;color:var(--tx3);margin-top:4px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}
.kpi-hint{font-size:10px;color:var(--acc);margin-top:6px;opacity:0;transition:opacity .15s;display:flex;align-items:center;gap:3px;}
.kpi-card:hover .kpi-hint{opacity:1;}

/* ── KPI DETAIL PANEL ── */
.kpi-detail{background:var(--sf);border:1px solid var(--bd);border-radius:var(--rl);padding:1.25rem 1.5rem;margin-bottom:1.25rem;animation:fadeIn .2s ease;}
@keyframes fadeIn{from{opacity:0;transform:translateY(-6px)}to{opacity:1;transform:translateY(0)}}
.kpi-detail-header{display:flex;align-items:center;justify-content:space-between;margin-bottom:1rem;}
.kpi-detail-title{font-size:13px;font-weight:600;color:var(--tx);}
.kpi-back{font-size:12px;color:var(--acc);cursor:pointer;display:flex;align-items:center;gap:4px;transition:opacity .15s;}
.kpi-back:hover{opacity:.7;}
.detail-kpis{display:grid;grid-template-columns:repeat(4,1fr);gap:10px;margin-bottom:1.25rem;}
.detail-kpi{background:var(--sf2);border-radius:var(--r);padding:10px 14px;border:1px solid var(--bd);}
.dk-label{font-size:10px;font-weight:600;color:var(--tx3);text-transform:uppercase;letter-spacing:.06em;margin-bottom:3px;}
.dk-value{font-size:18px;font-weight:700;color:var(--tx);}
.dk-sub{font-size:11px;color:var(--tx3);margin-top:2px;}

/* ── CARDS / LAYOUT ── */
.card{background:var(--sf);border:1px solid var(--bd);border-radius:var(--rl);padding:1.25rem 1.5rem;}
.card-title{font-size:10px;font-weight:700;color:var(--tx3);text-transform:uppercase;letter-spacing:.08em;margin-bottom:.85rem;}
.g2{display:grid;grid-template-columns:1fr 1fr;gap:1.1rem;margin-bottom:1.1rem;}
.g3{display:grid;grid-template-columns:1fr 1fr 1fr;gap:1.1rem;margin-bottom:1.1rem;}
.legend{display:flex;flex-wrap:wrap;gap:10px;margin-bottom:6px;font-size:11px;color:var(--tx2);}
.legend span{display:flex;align-items:center;gap:5px;}
.ld{width:8px;height:8px;border-radius:2px;flex-shrink:0;}

/* ── STATUS ROWS (detail panel) ── */
.status-row-item{display:flex;align-items:center;justify-content:space-between;padding:8px 12px;border-radius:var(--r);border:1px solid var(--bd);cursor:pointer;transition:background .15s;margin-bottom:6px;}
.status-row-item:hover{background:var(--sf2);}
.status-row-item.sr-active{background:var(--sf3);border-color:var(--acc);}
.sr-left{display:flex;align-items:center;gap:8px;font-size:13px;color:var(--tx);}
.sr-dot{width:7px;height:7px;border-radius:50%;flex-shrink:0;}
.pill{font-size:10px;font-weight:700;padding:2px 9px;border-radius:20px;}
.pill-ok{background:var(--grn2);color:var(--grn3);}
.pill-bad{background:var(--red2);color:var(--red3);}
.pill-warn{background:var(--amb2);color:var(--amb3);}
.pill-blue{background:var(--blu2);color:var(--blu3);}
.pill-pur{background:var(--pur2);color:var(--pur3);}

/* ── BAR ROWS ── */
.bar-row{display:flex;align-items:center;gap:8px;margin-bottom:7px;}
.bar-lbl{font-size:12px;color:var(--tx);width:140px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}
.bar-track{flex:1;height:7px;background:var(--sf3);border-radius:4px;overflow:hidden;}
.bar-fill{height:100%;border-radius:4px;transition:width .4s ease;}
.bar-num{font-size:11px;color:var(--tx2);font-family:'DM Mono',monospace;width:50px;text-align:right;}

/* ── EXPORT BAR ── */
.export-bar{display:flex;gap:8px;margin-bottom:1.25rem;flex-wrap:wrap;}
.btn-export{display:flex;align-items:center;gap:6px;padding:7px 14px;border:1px solid var(--bd2);border-radius:var(--r);background:var(--sf2);color:var(--tx2);font-family:'Inter',sans-serif;font-size:12px;font-weight:500;cursor:pointer;transition:all .15s;}
.btn-export:hover{background:var(--sf3);color:var(--tx);border-color:var(--acc);}
.chart-clickable{cursor:pointer;}

/* ── MODAL ── */
.modal-overlay{display:none;position:fixed;inset:0;background:rgba(0,0,0,.7);z-index:500;align-items:center;justify-content:center;backdrop-filter:blur(4px);}
.modal-overlay.open{display:flex;}
.modal{background:var(--sf);border:1px solid var(--bd2);border-radius:var(--rxl);padding:1.75rem;width:96%;max-width:1100px;max-height:90vh;overflow-y:auto;position:relative;}
.modal-close{position:absolute;top:1rem;right:1rem;background:var(--sf2);border:1px solid var(--bd2);border-radius:6px;color:var(--tx2);cursor:pointer;padding:3px 10px;font-size:16px;}
.modal-close:hover{color:var(--tx);}
.modal-title{font-size:15px;font-weight:700;margin-bottom:1rem;color:var(--tx);}
.modal-metrics{display:grid;grid-template-columns:repeat(auto-fit,minmax(130px,1fr));gap:8px;margin-bottom:1rem;}
.modal-metric{background:var(--sf2);border:1px solid var(--bd);border-radius:var(--r);padding:.75rem 1rem;}
.mm-label{font-size:10px;font-weight:600;color:var(--tx3);text-transform:uppercase;letter-spacing:.05em;margin-bottom:2px;}
.mm-val{font-size:17px;font-weight:700;color:var(--tx);}

/* ── TABLE ── */
.table-section{background:var(--sf);border:1px solid var(--bd);border-radius:var(--rl);overflow:hidden;margin-bottom:1.5rem;}
.table-hdr{padding:.85rem 1.25rem;border-bottom:1px solid var(--bd);display:flex;align-items:center;justify-content:space-between;gap:10px;flex-wrap:wrap;background:var(--sf2);}
.table-title{font-size:13px;font-weight:600;color:var(--tx);}
.table-count{font-size:11px;color:var(--tx3);font-family:'DM Mono',monospace;background:var(--sf3);padding:2px 9px;border-radius:20px;border:1px solid var(--bd);}
.table-search{padding:6px 12px;border-radius:var(--r);border:1px solid var(--bd2);background:var(--sf3);color:var(--tx);font-family:'Inter',sans-serif;font-size:12px;outline:none;width:260px;}
.table-search::placeholder{color:var(--tx3);}
.table-search:focus{border-color:var(--acc);}
.table-wrap{overflow-x:auto;}
table{width:100%;border-collapse:collapse;}
thead th{padding:8px 12px;text-align:left;font-size:10px;font-weight:700;color:var(--tx3);text-transform:uppercase;letter-spacing:.06em;background:var(--sf2);border-bottom:1px solid var(--bd);white-space:nowrap;cursor:pointer;user-select:none;}
thead th:hover{color:var(--tx);}
thead th.sa::after{content:' ↑';color:var(--acc);}
thead th.sd::after{content:' ↓';color:var(--acc);}
tbody td{padding:8px 12px;border-bottom:1px solid var(--bd);font-size:12px;white-space:nowrap;color:var(--tx);}
tbody tr:last-child td{border-bottom:none;}
tbody tr:hover td{background:var(--sf2);}
.badge{display:inline-flex;align-items:center;padding:2px 8px;border-radius:20px;font-size:10px;font-weight:700;}
.bg{background:rgba(29,158,117,.15);color:#5DCAA5;border:1px solid rgba(29,158,117,.3);}
.br{background:rgba(226,75,74,.12);color:#F09595;border:1px solid rgba(226,75,74,.25);}
.ba{background:rgba(186,117,23,.12);color:#EF9F27;border:1px solid rgba(186,117,23,.25);}
.vc{font-family:'DM Mono',monospace;font-weight:500;color:var(--acc);}
.vr{font-family:'DM Mono',monospace;font-weight:500;color:var(--red);}
.vo{font-family:'DM Mono',monospace;font-weight:500;color:var(--amb);}

/* ── PAGINATION ── */
.pagination{display:flex;align-items:center;justify-content:space-between;padding:.65rem 1.25rem;border-top:1px solid var(--bd);font-size:12px;color:var(--tx3);flex-wrap:wrap;gap:8px;background:var(--sf2);}
.pg-btns{display:flex;gap:4px;}
.pgb{padding:4px 10px;border:1px solid var(--bd2);border-radius:6px;background:var(--sf3);cursor:pointer;font-size:11px;color:var(--tx2);transition:all .15s;}
.pgb:hover:not(:disabled){background:var(--bd);color:var(--tx);}
.pgb:disabled{opacity:.3;cursor:default;}
.pgb.active{background:var(--acc);color:#fff;border-color:var(--acc);}

/* ── MODAL TABLE ── */
.modal-table{width:100%;border-collapse:collapse;font-size:12px;}
.modal-table th{padding:7px 10px;text-align:left;font-size:10px;font-weight:700;color:var(--tx3);text-transform:uppercase;letter-spacing:.05em;background:var(--sf2);border-bottom:1px solid var(--bd);}
.modal-table td{padding:7px 10px;border-bottom:1px solid var(--bd);color:var(--tx);}
.modal-table tr:last-child td{border-bottom:none;}
.modal-table tr:hover td{background:var(--sf2);}

footer{text-align:center;padding:1.75rem;font-size:12px;color:var(--tx4);border-top:1px solid var(--bd);margin-top:1.5rem;}

@media(max-width:1400px){.kpi-grid{grid-template-columns:repeat(7,minmax(0,1fr));}}
@media(max-width:1200px){.kpi-grid{grid-template-columns:repeat(4,1fr);}}
@media(max-width:900px){.g2,.g3{grid-template-columns:1fr;}.kpi-grid{grid-template-columns:repeat(2,1fr);}}
@media(max-width:600px){.main{padding:1rem;}.kpi-grid{grid-template-columns:repeat(2,1fr);}}

.pend-quin-card{
  transition:border-color .2s,background .2s,box-shadow .2s,transform .15s;
}
.pend-quin-card:hover{
  transform:translateY(-2px);
  box-shadow:0 0 0 1px var(--kpi-color),0 4px 18px color-mix(in srgb,var(--kpi-color) 30%,transparent);
}
.pend-quin-card.pqc-active{
  box-shadow:0 0 0 1px var(--kpi-color),0 4px 20px color-mix(in srgb,var(--kpi-color) 38%,transparent);
}

