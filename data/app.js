// ============================================================
// app.js — Dashboard MPCH Egresos · Municipalidad Provincial de Chiclayo
// Versión 2.0 — Julio 2026
// Bloques: B1 KPIs+Donuts · B2 Ranking · B3 Top Proyectos · B4 Riesgo
//          B5 Rubro · B6 Fuente · B7 Función · B8 Histórico
// Fuente: MEF Consulta Amigable de Gastos — Solo Proyectos
// ============================================================
"use strict";

// ── Utilidades ───────────────────────────────────────────────
const $ = id => document.getElementById(id);

function fmtNum(n) {
  if (n === null || n === undefined || isNaN(n)) return "—";
  return "S/ " + Math.round(n).toLocaleString("es-PE");
}
// Número sin prefijo S/ — para tablas (evitar redundancia con header)
function fmtN(n) {
  if (n === null || n === undefined || isNaN(n)) return "—";
  return Math.round(n).toLocaleString("es-PE");
}
function fmtCompacto(n) {
  if (!n && n !== 0) return "—";
  if (Math.abs(n) >= 1e6) return "S/ " + (n/1e6).toFixed(1) + " M";
  if (Math.abs(n) >= 1e3) return "S/ " + (n/1e3).toFixed(0) + " K";
  return "S/ " + Math.round(n).toLocaleString("es-PE");
}
function fmtPct(n) {
  if (n === null || isNaN(n)) return "—";
  return n.toFixed(1) + "%";
}
function parseNum(s) {
  if (!s) return 0;
  return parseFloat(String(s).replace(/,/g,"")) || 0;
}
function fechaHoy() {
  const d=new Date(), m=["enero","febrero","marzo","abril","mayo","junio",
  "julio","agosto","setiembre","octubre","noviembre","diciembre"],
  ds=["domingo","lunes","martes","miércoles","jueves","viernes","sábado"];
  return d.getDate()+" de "+m[d.getMonth()]+" de "+d.getFullYear();
}
// Rellena todos los spans de fecha en los subbands
function rellenarFechas() {
  const f = fechaHoy();
  ["sb1fecha","sb2fecha","sb3fecha","sb4fecha",
   "sb5fecha","sb6fecha","sb7fecha","sb8fecha"].forEach(id=>{
    const el=$(id); if(el) el.textContent=f;
  });
}
function semaforo(pct) {
  if (pct===null) return "#888";
  if (pct>=70) return "var(--verde)";
  if (pct>=40) return "var(--amarillo-s)";
  return "var(--rojo-s)";
}
function barraHTML(pct) {
  if (pct===null) return '<span style="color:#aaa;font-size:.9em">—</span>';
  const col=semaforo(pct), w=Math.min(pct,100).toFixed(1);
  return `<div style="display:flex;align-items:center;gap:6px">
    <div style="flex:1;background:#e5e7eb;border-radius:4px;height:7px;min-width:55px">
      <div style="width:${w}%;background:${col};height:7px;border-radius:4px"></div>
    </div>
    <span style="font-family:'Barlow Condensed';font-weight:800;font-size:12.5px;color:${col};min-width:38px">${pct.toFixed(1)}%</span>
  </div>`;
}

// ── Estado global ────────────────────────────────────────────
let datos = {};
let cargados = new Set();
let b8ChartInstance = null;
let B8_HIST = {};
let donutInstances = {};

const ARCHIVOS_ESPERADOS = [
  "rubro.xls","categoria.xls","proyecto.xls","ranking.xls","fuente.xls","funcion.xls"
];

// ── Parseo HTML/XLS ──────────────────────────────────────────
function extraerFilas(buffer) {
  const text = new TextDecoder("utf-8").decode(buffer);
  const trRe = /<tr[^>]*>([\s\S]*?)<\/tr>/gi;
  const tdRe = /<t[dh][^>]*>([\s\S]*?)<\/t[dh]>/gi;
  const rows = [];
  let trM;
  while ((trM=trRe.exec(text))!==null) {
    const cells=[]; let tdM; const inner=trM[1]; tdRe.lastIndex=0;
    while ((tdM=tdRe.exec(inner))!==null) {
      const val=tdM[1].replace(/<[^>]+>/g,"").replace(/&nbsp;/g,"").replace(/\r\n/g,"").trim();
      cells.push(val);
    }
    if (cells.some(c=>c)) rows.push(cells);
  }
  return rows;
}

// Columnas MEF Gastos:
// 0:Desc  1:PIA  2:PIM  3:Cert  4:CompAnual  5:ACM  6:Devengado  7:Girado  8:Avance%
function detectarTipo(rows) {
  for (const r of rows) {
    const d=(r[0]||"").trim();
    if (/^Rubro$/i.test(d))               return "rubro";
    if (/^Fuente de Financiamiento/i.test(d)) return "fuente";
    if (/^Funci/i.test(d))                return "funcion";
    if (/^Proyecto$/i.test(d))            return "proyecto";
    if (/^Categor/i.test(d))              return "categoria";
    if (/^Municipalidad$/i.test(d))       return "ranking";
  }
  return null;
}

function esFilaDato(r) {
  if (!r || r.length < 7) return false;
  const d=(r[0]||"").trim();
  if (!d) return false;
  if (/^(Rubro|Fuente|Funci|Proyecto|Municipalidad|Categor|Mes|Nivel|Gob\.|Departamento|Consulta|Atenci|TOTAL$)/i.test(d)) return false;
  return [1,2,3,4,5,6,7].some(i=>/[\d,]+/.test(r[i]||""));
}

function buscarTotalMPL(rows) {
  for (const r of rows) {
    if ((r[0]||"").includes("140101-301212") ||
        (r[0]||"").toUpperCase().includes("MUNICIPALIDAD PROVINCIAL DE CHICLAYO")) {
      return {
        pia:parseNum(r[1]),pim:parseNum(r[2]),cert:parseNum(r[3]),
        comp:parseNum(r[4]),dev:parseNum(r[6]),girado:parseNum(r[7])
      };
    }
  }
  return null;
}

function parsearArchivo(buffer, nombre) {
  const rows = extraerFilas(buffer);
  const tipo = detectarTipo(rows);
  if (!tipo) { console.warn("[MPCH-EG] Tipo no detectado:", nombre); return null; }

  const mpl = buscarTotalMPL(rows);
  const registros = [];
  for (const r of rows) {
    if (!esFilaDato(r)) continue;
    const pim=parseNum(r[2]), dev=parseNum(r[6]), cert=parseNum(r[3]), comp=parseNum(r[4]);
    registros.push({
      desc:r[0]||"",
      pia:parseNum(r[1]),pim,cert,comp,
      acm:parseNum(r[5]),dev,girado:parseNum(r[7]),
      pct:pim>0?dev/pim*100:null,
      pctCert:pim>0?cert/pim*100:null,
      pctComp:pim>0?comp/pim*100:null
    });
  }

  // Totales
  let tot = mpl;
  if (!tot || tot.pim===0) {
    tot = registros.reduce((a,r)=>({
      pia:a.pia+r.pia,pim:a.pim+r.pim,cert:a.cert+r.cert,
      comp:a.comp+r.comp,dev:a.dev+r.dev,girado:a.girado+r.girado
    }),{pia:0,pim:0,cert:0,comp:0,dev:0,girado:0});
  }

  console.log(`[MPCH-EG] ${nombre} → tipo=${tipo} | regs=${registros.length} | PIM=${tot.pim} | Dev=${tot.dev}`);
  return { tipo, nombre, registros, ...tot };
}

function actualizarFileList() {
  const fl=$("fileList"); if(!fl) return;
  fl.innerHTML=ARCHIVOS_ESPERADOS.map(f=>{
    const ok=cargados.has(f);
    return `<div class="file-item">
      <span class="${ok?"fi-ok":"fi-wait"}">${ok?"✓":"○"}</span>
      <span class="fi-name">${f}</span>
    </div>`;
  }).join("");
}

async function autoCargar() {
  for (const nombre of ARCHIVOS_ESPERADOS) {
    try {
      const r=await fetch("xls/"+nombre+"?"+Date.now());
      if(!r.ok) continue;
      const buf=await r.arrayBuffer();
      const res=parsearArchivo(buf,nombre);
      if(res){ datos[res.tipo]=res; cargados.add(nombre); actualizarFileList(); render(); }
    } catch(e){}
  }
}

function procesarArchivos(files) {
  Array.from(files).forEach(file=>{
    const reader=new FileReader();
    reader.onload=e=>{
      const res=parsearArchivo(e.target.result,file.name);
      if(res){ datos[res.tipo]=res; cargados.add(file.name); actualizarFileList(); render(); }
    };
    reader.readAsArrayBuffer(file);
  });
}

// ── Render principal ──────────────────────────────────────────
function render() {
  const hoy=fechaHoy();
  ["b2fecha","b3fecha","b4fecha","b5fecha","b6fecha","b7fecha"].forEach(id=>{
    const el=$(id); if(el) el.textContent=hoy;
  });
  rellenarFechas();
  renderB1(); renderB2(); renderB3(); renderB4();
  renderB5(); renderB6(); renderB7(); renderB8();
}

// ══════════════════════════════════════════════════════════════
// B1 — KPIs + 3 Donuts
// ══════════════════════════════════════════════════════════════
function renderDonut(canvasId, pct, colorFill, colorEmpty, small=false) {
  const canvas=$(canvasId); if(!canvas) return;
  if(donutInstances[canvasId]){ donutInstances[canvasId].destroy(); delete donutInstances[canvasId]; }
  const safe=Math.min(Math.max(pct||0,0),100);
  donutInstances[canvasId]=new Chart(canvas,{
    type:"doughnut",
    data:{ datasets:[{ data:[safe,100-safe], backgroundColor:[colorFill,colorEmpty], borderWidth:0, hoverOffset:0 }] },
    options:{
      cutout:small?"68%":"72%", responsive:false, animation:{duration:500},
      plugins:{legend:{display:false},tooltip:{enabled:false}}
    }
  });
}

function renderB1() {
  const d=datos.rubro; if(!d) return;
  const pim=d.pim, cert=d.cert, comp=d.comp, dev=d.dev, porDev=pim-dev;
  const pctCert=pim>0?cert/pim*100:0;
  const pctDev =pim>0?dev /pim*100:0;
  const pctComp=pim>0?comp/pim*100:0;

  $("kpi-pim").textContent    =fmtNum(pim);
  $("kpi-pim-sub").textContent=d.registros.length+" registros cargados";
  $("kpi-cert").textContent   =fmtNum(cert);
  $("kpi-cert-sub").textContent=pctCert.toFixed(1)+"%";
  $("kpi-comp").textContent   =fmtNum(comp);
  $("kpi-comp-sub").textContent=pctComp.toFixed(1)+"%";
  $("kpi-dev").textContent    =fmtNum(dev);
  $("kpi-dev-sub").textContent=pctDev.toFixed(1)+"%";
  $("kpi-pordev").textContent =fmtNum(porDev);

  // Donuts grandes
  renderDonut("donutCert",pctCert,"#1e7e34","#a8d5b0");
  $("donutCertPct").textContent=pctCert.toFixed(1)+"%";
  $("legCert").textContent   =fmtNum(cert);
  $("legPorCert").textContent=fmtNum(pim-cert);

  renderDonut("donutDev",pctDev,"#1f5f91","#c3dcef");
  $("donutDevPct").textContent=pctDev.toFixed(1)+"%";
  $("legDev").textContent   =fmtNum(dev);
  $("legPorDev").textContent=fmtNum(porDev);

  renderDonut("donutComp",pctComp,"#f5c518","#faeab0");
  $("donutCompPct").textContent=pctComp.toFixed(1)+"%";
  $("legComp").textContent   =fmtNum(comp);
  $("legPorComp").textContent=fmtNum(pim-comp);
}

// ══════════════════════════════════════════════════════════════
// B2 — Ranking Departamental
// ══════════════════════════════════════════════════════════════
function renderB2() {
  const d=datos.ranking, tbody=$("b2tbody");
  if(!tbody) return;
  if(!d){ tbody.innerHTML=`<tr><td colspan="6" class="vacio">Carga ranking.xls para ver los datos.</td></tr>`; return; }

  const muns=d.registros.map(r=>{
    const esMPL=r.desc.includes("140101-301212")||r.desc.toUpperCase().includes("PROVINCIAL DE CHICLAYO");
    // Quitar código "140301-XXXXXX: " del inicio del nombre
    let nombre=r.desc.replace(/^\d{6}-\d{6}:\s*/,"").replace(/^\d+-\d+:\s*/,"");
    // Normalizar Ñ → N (encoding issues del MEF)
    nombre=nombre.replace(/�/g,"N").replace(/Ñ/g,"N").replace(/ñ/g,"n");
    return {...r,esMPL,nombre};
  });
  const sorted=[...muns].sort((a,b)=>(b.pct??-1)-(a.pct??-1));
  const mpl=sorted.find(m=>m.esMPL);
  const pos=mpl?sorted.indexOf(mpl)+1:null;
  const porMonto=[...muns].sort((a,b)=>b.dev-a.dev);
  const posMonto=mpl?porMonto.indexOf(mpl)+1:null;

  // Bloque highlight eliminado por diseño

  tbody.innerHTML=sorted.map((m,i)=>{
    const pos=i+1;
    const posStyle=pos<=3?'style="color:var(--dorado-osc);font-weight:800"':'style="font-weight:700"';
    return `<tr ${m.esMPL?'class="mpl-row"':''}>
      <td class="ctr" ${posStyle}>${pos}°</td>
      <td style="font-size:.95em;font-weight:${m.esMPL?700:600}">${m.nombre}</td>
      <td class="num">${fmtNum(m.pia)}</td>
      <td class="num">${fmtNum(m.pim)}</td>
      <td class="num">${fmtNum(m.dev)}</td>
      <td>${barraHTML(m.pct)}</td>
    </tr>`;
  }).join("");
}

// ══════════════════════════════════════════════════════════════
// B3 — Top 10 Proyectos por Devengado
// ══════════════════════════════════════════════════════════════
function renderB3() {
  const d=datos.proyecto, tbody=$("b3tbody"), tfoot=$("b3tfoot"), prom=$("b3prom");
  if(!tbody) return;
  if(!d){ tbody.innerHTML=`<tr><td colspan="9" class="vacio">Carga proyecto.xls para ver los datos.</td></tr>`; return; }

  const sorted=[...d.registros].sort((a,b)=>b.dev-a.dev);
  const top10=sorted.slice(0,10);
  const promPct=top10.length?top10.reduce((s,r)=>s+(r.pct||0),0)/top10.length:0;
  if(prom) prom.innerHTML=`PROM. DE AVANCE (TOP 10): <strong>${promPct.toFixed(1)}%</strong>`;

  tbody.innerHTML=top10.map((r,i)=>{
    const m=r.desc.match(/^(\d+):\s*(.+)$/);
    const nom=m?m[2]:r.desc;
    const porCert=r.pim-r.cert, porDev=r.pim-r.dev;
    const col=semaforo(r.pct);
    return `<tr>
      <td class="ctr" style="font-weight:700;color:${i<3?"var(--dorado-osc)":"var(--texto-mut)"}">${i+1}</td>
      <td style="font-size:.93em;font-weight:600;line-height:1.3">${nom}</td>
      <td class="num">${fmtN(r.pia)}</td>
      <td class="num">${fmtN(r.pim)}</td>
      <td class="num">${fmtN(r.cert)}</td>
      <td class="num col-highlight" style="color:${porCert>0?"#92400e":"#aaa"}">${fmtN(porCert)}</td>
      <td class="num" style="color:var(--rojo-osc);font-weight:700">${fmtN(r.dev)}</td>
      <td class="num col-highlight" style="color:${porDev>0?"#92400e":"#aaa"}">${fmtN(porDev)}</td>
      <td>${barraHTML(r.pct)}</td>
    </tr>`;
  }).join("");

  const tot={
    pia:top10.reduce((s,r)=>s+r.pia,0),
    pim:top10.reduce((s,r)=>s+r.pim,0),
    cert:top10.reduce((s,r)=>s+r.cert,0),
    dev:top10.reduce((s,r)=>s+r.dev,0)
  };
  const totPct=tot.pim>0?tot.dev/tot.pim*100:null;
  tfoot.innerHTML=`<tr>
    <td style="font-weight:800;font-family:'Barlow Condensed';text-align:center;color:var(--texto-mut)">—</td>
    <td style="font-weight:800;font-family:'Barlow Condensed';text-transform:uppercase">Total Top 10</td>
    <td class="num" style="color:var(--rojo-osc);font-weight:800">${fmtN(tot.pia)}</td>
    <td class="num" style="color:var(--rojo-osc);font-weight:800">${fmtN(tot.pim)}</td>
    <td class="num" style="color:var(--rojo-osc);font-weight:800">${fmtN(tot.cert)}</td>
    <td class="num col-highlight" style="color:#92400e;font-weight:800">${fmtN(tot.pim-tot.cert)}</td>
    <td class="num" style="color:var(--rojo-osc);font-weight:800">${fmtN(tot.dev)}</td>
    <td class="num col-highlight" style="color:#92400e;font-weight:800">${fmtN(tot.pim-tot.dev)}</td>
    <td>${barraHTML(totPct)}</td>
  </tr>`;
}

// ══════════════════════════════════════════════════════════════
// B4 — Proyectos con Ejecución en Riesgo (dev = 0)
// ══════════════════════════════════════════════════════════════
function renderB4() {
  const d=datos.proyecto, tbody=$("b4tbody"), tfoot=$("b4tfoot");
  if(!tbody) return;
  if(!d){ tbody.innerHTML=`<tr><td colspan="8" class="vacio">Carga proyecto.xls para ver los datos.</td></tr>`; return; }

  const riesgo=d.registros.filter(r=>r.dev===0&&r.pim>0).sort((a,b)=>b.pim-a.pim).slice(0,10);
  const totalRiesgo=riesgo.reduce((s,r)=>s+r.pim,0);
  // La fecha ya la rellena rellenarFechas() — el monto en riesgo va en la nota interna

  tbody.innerHTML=riesgo.map((r,i)=>{
    const m=r.desc.match(/^(\d+):\s*(.+)$/);
    const nom=m?m[2]:r.desc;
    const pctCertStr=r.pim>0?(r.cert/r.pim*100).toFixed(1)+"%":"0.0%";
    return `<tr class="riesgo-row">
      <td class="ctr" style="font-weight:700">${i+1}</td>
      <td style="font-size:.93em;font-weight:600;line-height:1.3">${nom}</td>
      <td class="num">${fmtN(r.pia)}</td>
      <td class="num">${fmtN(r.pim)}</td>
      <td class="num">${fmtN(r.cert)}</td>
      <td class="num">${pctCertStr}</td>
      <td class="num" style="font-weight:800;color:#374151">0</td>
      <td class="num pim-sin-ej">${fmtN(r.pim)}</td>
    </tr>`;
  }).join("");

  const totPIA=riesgo.reduce((s,r)=>s+r.pia,0);
  const totPIM=riesgo.reduce((s,r)=>s+r.pim,0);
  const totCert=riesgo.reduce((s,r)=>s+r.cert,0);
  tfoot.innerHTML=`<tr>
    <td style="font-weight:800;font-family:'Barlow Condensed';text-align:center;color:var(--texto-mut)">—</td>
    <td style="font-weight:800;font-family:'Barlow Condensed';text-transform:uppercase">Total Top 10</td>
    <td class="num" style="color:var(--rojo-osc);font-weight:800">${fmtN(totPIA)}</td>
    <td class="num" style="color:var(--rojo-osc);font-weight:800">${fmtN(totPIM)}</td>
    <td class="num" style="color:var(--rojo-osc);font-weight:800">${fmtN(totCert)}</td>
    <td class="num">${totPIM>0?(totCert/totPIM*100).toFixed(1)+"%":"—"}</td>
    <td class="num" style="font-weight:800;color:#374151">0</td>
    <td class="num pim-sin-ej">${fmtN(totPIM)}</td>
  </tr>`;
}

// ══════════════════════════════════════════════════════════════
// Render donuts pequeños (rubro/fuente)
// ══════════════════════════════════════════════════════════════
const PALETA_DONUTS = [
  ["#2563eb","#bfdbfe"], // azul
  ["#d97706","#fde68a"], // ámbar
  ["#059669","#a7f3d0"], // esmeralda
  ["#7c3aed","#ddd6fe"], // violeta
  ["#dc2626","#fecaca"], // rojo
  ["#0891b2","#a5f3fc"], // cian
  ["#c2410c","#fed7aa"], // naranja
];

function renderDonutsSeccion(contenedorId, registros, totalPIM, totalDev) {
  const cont=$(contenedorId); if(!cont) return;
  // Destruir instancias anteriores en este contenedor
  cont.querySelectorAll("canvas").forEach(c=>{
    if(donutInstances[c.id]){ donutInstances[c.id].destroy(); delete donutInstances[c.id]; }
  });

  let html="";
  registros.forEach((r,i)=>{
    const pct=r.pct!==null?r.pct:0;
    const [cf,ce]=PALETA_DONUTS[i%PALETA_DONUTS.length];
    const cid=`${contenedorId}_d${i}`;
    // Extraer nombre corto
    const m=r.desc.match(/^\d+:\s*(.+)$/);
    const nom=m?m[1]:r.desc;
    const shortNom=nom.length>22?nom.substring(0,20)+"…":nom;
    html+=`<div class="rub-donut-card">
      <div class="rub-donut-label">${shortNom}</div>
      <div class="rub-donut-wrap">
        <canvas id="${cid}" width="90" height="90"></canvas>
        <div class="rub-donut-center">
          <div class="rub-donut-pct" style="color:${cf}">${pct.toFixed(1)}%</div>
          <div class="rub-donut-sub">Dev.</div>
        </div>
      </div>
      <div class="rub-donut-pim">PIM: ${fmtCompacto(r.pim)}</div>
    </div>`;
  });

  // Donut total
  const totPct=totalPIM>0?totalDev/totalPIM*100:0;
  const tidTotal=`${contenedorId}_total`;
  html+=`<div class="rub-donut-card total-card">
    <div class="rub-donut-label">DEVENGADO TOTAL %</div>
    <div class="rub-donut-wrap" style="width:110px;height:110px">
      <canvas id="${tidTotal}" width="110" height="110"></canvas>
      <div class="rub-donut-center">
        <div class="rub-donut-pct" style="color:var(--rojo-osc);font-size:22px;font-weight:800">${totPct.toFixed(1)}%</div>
        <div class="rub-donut-sub">TOTAL</div>
      </div>
    </div>
    <div class="rub-donut-pim"><strong style="color:var(--rojo-osc)">${fmtCompacto(totalDev)}</strong></div>
  </div>`;

  cont.innerHTML=html;

  // Dibujar después de insertar en DOM
  registros.forEach((r,i)=>{
    const pct=r.pct!==null?r.pct:0;
    const [cf,ce]=PALETA_DONUTS[i%PALETA_DONUTS.length];
    const cid=`${contenedorId}_d${i}`;
    renderDonut(cid,pct,cf,ce,true);
  });
  renderDonut(tidTotal,totPct,"#1f5f91","#c3dcef",false);
}

// ══════════════════════════════════════════════════════════════
// B5 — Rubro (donuts + tabla completa)
// ══════════════════════════════════════════════════════════════
function renderB5() {
  const d=datos.rubro, tbody=$("b5tbody"), tfoot=$("b5tfoot");
  if(!tbody) return;
  if(!d){ tbody.innerHTML=`<tr><td colspan="13" class="vacio">Carga rubro.xls para ver los datos.</td></tr>`; return; }

  renderDonutsSeccion("b5donuts", d.registros, d.pim, d.dev);

  tbody.innerHTML=d.registros.map(r=>{
    const porCert=r.pim-r.cert, porDev=r.pim-r.dev;
    const pctPIM=d.pim>0?r.pim/d.pim*100:0;
    const m=r.desc.match(/^(\d+):\s*(.+)$/);
    const cod=m?m[1]:"", nom=m?m[2]:r.desc;
    return `<tr>
      <td class="cod">${cod}</td>
      <td style="font-size:.93em;font-weight:600">${nom}</td>
      <td class="num">${fmtN(r.pia)}</td>
      <td class="num">${fmtN(r.pim)}</td>
      <td class="num" style="color:var(--texto-mut)">${pctPIM.toFixed(1)}%</td>
      <td class="num">${fmtN(r.cert)}</td>
      <td class="num" style="color:#1e7e34">${fmtPct(r.pctCert)}</td>
      <td class="num col-highlight">${fmtN(porCert)}</td>
      <td class="num">${fmtN(r.comp)}</td>
      <td class="num" style="color:var(--dorado-osc)">${fmtPct(r.pctComp)}</td>
      <td class="num" style="color:var(--rojo-osc);font-weight:700">${fmtN(r.dev)}</td>
      <td class="num">${fmtN(porDev)}</td>
      <td>${barraHTML(r.pct)}</td>
    </tr>`;
  }).join("");

  const totPct=d.pim>0?d.dev/d.pim*100:null;
  tfoot.innerHTML=`<tr>
    <td colspan="2" style="font-weight:800;font-family:'Barlow Condensed';text-transform:uppercase">TOTAL</td>
    <td class="num" style="color:var(--rojo-osc);font-weight:800">${fmtN(d.pia)}</td>
    <td class="num" style="color:var(--rojo-osc);font-weight:800">${fmtN(d.pim)}</td>
    <td class="num">100%</td>
    <td class="num" style="color:var(--rojo-osc);font-weight:800">${fmtN(d.cert)}</td>
    <td class="num" style="color:#1e7e34;font-weight:800">${fmtPct(d.pim>0?d.cert/d.pim*100:null)}</td>
    <td class="num col-highlight" style="font-weight:800">${fmtN(d.pim-d.cert)}</td>
    <td class="num" style="color:var(--rojo-osc);font-weight:800">${fmtN(d.comp)}</td>
    <td class="num" style="font-weight:800">${fmtPct(d.pim>0?d.comp/d.pim*100:null)}</td>
    <td class="num" style="color:var(--rojo-osc);font-weight:800">${fmtN(d.dev)}</td>
    <td class="num" style="font-weight:800">${fmtN(d.pim-d.dev)}</td>
    <td>${barraHTML(totPct)}</td>
  </tr>`;
}

// ══════════════════════════════════════════════════════════════
// B6 — Fuente de Financiamiento (donuts + tabla)
// ══════════════════════════════════════════════════════════════
function renderB6() {
  const d=datos.fuente, tbody=$("b6tbody"), tfoot=$("b6tfoot");
  if(!tbody) return;
  if(!d){ tbody.innerHTML=`<tr><td colspan="13" class="vacio">Carga fuente.xls para ver los datos.</td></tr>`; return; }

  renderDonutsSeccion("b6donuts", d.registros, d.pim, d.dev);

  tbody.innerHTML=d.registros.map(r=>{
    const porCert=r.pim-r.cert, porDev=r.pim-r.dev;
    const pctPIM=d.pim>0?r.pim/d.pim*100:0;
    const m=r.desc.match(/^(\d+):\s*(.+)$/);
    const cod=m?m[1]:"", nom=m?m[2]:r.desc;
    return `<tr>
      <td class="cod">${cod}</td>
      <td style="font-size:.93em;font-weight:600">${nom}</td>
      <td class="num">${fmtN(r.pia)}</td>
      <td class="num">${fmtN(r.pim)}</td>
      <td class="num" style="color:var(--texto-mut)">${pctPIM.toFixed(1)}%</td>
      <td class="num">${fmtN(r.cert)}</td>
      <td class="num" style="color:#1e7e34">${fmtPct(r.pctCert)}</td>
      <td class="num col-highlight">${fmtN(porCert)}</td>
      <td class="num">${fmtN(r.comp)}</td>
      <td class="num" style="color:var(--dorado-osc)">${fmtPct(r.pctComp)}</td>
      <td class="num" style="color:var(--rojo-osc);font-weight:700">${fmtN(r.dev)}</td>
      <td class="num">${fmtN(porDev)}</td>
      <td>${barraHTML(r.pct)}</td>
    </tr>`;
  }).join("");

  const totPct=d.pim>0?d.dev/d.pim*100:null;
  tfoot.innerHTML=`<tr>
    <td colspan="2" style="font-weight:800;font-family:'Barlow Condensed';text-transform:uppercase">TOTAL MPCH</td>
    <td class="num" style="color:var(--rojo-osc);font-weight:800">${fmtN(d.pia)}</td>
    <td class="num" style="color:var(--rojo-osc);font-weight:800">${fmtN(d.pim)}</td>
    <td class="num">100%</td>
    <td class="num" style="color:var(--rojo-osc);font-weight:800">${fmtN(d.cert)}</td>
    <td class="num" style="color:#1e7e34;font-weight:800">${fmtPct(d.pim>0?d.cert/d.pim*100:null)}</td>
    <td class="num col-highlight" style="font-weight:800">${fmtN(d.pim-d.cert)}</td>
    <td class="num" style="color:var(--rojo-osc);font-weight:800">${fmtN(d.comp)}</td>
    <td class="num" style="font-weight:800">${fmtPct(d.pim>0?d.comp/d.pim*100:null)}</td>
    <td class="num" style="color:var(--rojo-osc);font-weight:800">${fmtN(d.dev)}</td>
    <td class="num" style="font-weight:800">${fmtN(d.pim-d.dev)}</td>
    <td>${barraHTML(totPct)}</td>
  </tr>`;
}

// ══════════════════════════════════════════════════════════════
// B7 — Funciones a nivel pliego (solo tabla)
// ══════════════════════════════════════════════════════════════
function renderB7() {
  const d=datos.funcion, tbody=$("b7tbody"), tfoot=$("b7tfoot");
  if(!tbody) return;
  if(!d){ tbody.innerHTML=`<tr><td colspan="11" class="vacio">Carga funcion.xls para ver los datos.</td></tr>`; return; }

  tbody.innerHTML=d.registros.map(r=>{
    const porDev=r.pim-r.dev;
    const pctPIM=d.pim>0?r.pim/d.pim*100:0;
    const m=r.desc.match(/^(\d+):\s*(.+)$/);
    const nom=m?m[2]:r.desc;
    return `<tr>
      <td style="font-size:.93em;font-weight:700">${nom}</td>
      <td class="num">${fmtN(r.pia)}</td>
      <td class="num">${fmtN(r.pim)}</td>
      <td class="num" style="color:var(--texto-mut)">${pctPIM.toFixed(1)}%</td>
      <td class="num">${fmtN(r.cert)}</td>
      <td class="num" style="color:#1e7e34">${fmtPct(r.pctCert)}</td>
      <td class="num">${fmtN(r.comp)}</td>
      <td class="num" style="color:var(--dorado-osc)">${fmtPct(r.pctComp)}</td>
      <td class="num" style="color:var(--rojo-osc);font-weight:700">${fmtN(r.dev)}</td>
      <td class="num">${fmtN(porDev)}</td>
      <td>${barraHTML(r.pct)}</td>
    </tr>`;
  }).join("");

  const totPct=d.pim>0?d.dev/d.pim*100:null;
  tfoot.innerHTML=`<tr>
    <td style="font-weight:800;font-family:'Barlow Condensed';text-transform:uppercase">TOTAL</td>
    <td class="num" style="color:var(--rojo-osc);font-weight:800">${fmtN(d.pia)}</td>
    <td class="num" style="color:var(--rojo-osc);font-weight:800">${fmtN(d.pim)}</td>
    <td class="num">100%</td>
    <td class="num" style="color:var(--rojo-osc);font-weight:800">${fmtN(d.cert)}</td>
    <td class="num" style="color:#1e7e34;font-weight:800">${fmtPct(d.pim>0?d.cert/d.pim*100:null)}</td>
    <td class="num" style="color:var(--rojo-osc);font-weight:800">${fmtN(d.comp)}</td>
    <td class="num" style="font-weight:800">${fmtPct(d.pim>0?d.comp/d.pim*100:null)}</td>
    <td class="num" style="color:var(--rojo-osc);font-weight:800">${fmtN(d.dev)}</td>
    <td class="num" style="font-weight:800">${fmtN(d.pim-d.dev)}</td>
    <td>${barraHTML(totPct)}</td>
  </tr>`;
}

// ══════════════════════════════════════════════════════════════
// B8 — Histórico Ene–Jul (diseño mejorado)
// ══════════════════════════════════════════════════════════════
function renderB8() {
  const años=Object.keys(B8_HIST).map(Number).sort();
  const dev2026=datos.rubro?datos.rubro.dev:null;
  if(dev2026!==null) años.push(2026);
  const IDX_2026=años.length-1;
  const valores=años.map(a=>a===2026?dev2026:(B8_HIST[a]??0));

  // KPI cards
  const kpiWrap=$("b8kpis");
  if(kpiWrap) {
    kpiWrap.innerHTML=años.map((a,i)=>{
      const v=valores[i], es2026=(a===2026);
      let vari="";
      if(i>0&&valores[i-1]>0){
        const delta=(v-valores[i-1])/valores[i-1]*100;
        const col=delta>=0?"#2a7d46":"#c0392b", sym=delta>=0?"▲":"▼";
        vari=`<div class="hist-var" style="color:${col}">${sym} ${Math.abs(delta).toFixed(1)}% vs ${a-1}</div>`;
      }
      return `<div class="hist-card ${es2026?"cur":""}">
        <div class="hist-label">${a}${es2026?" · Acum.":" · Ene–Set"}</div>
        <div class="hist-val">${fmtNum(v)}</div>
        ${vari}
      </div>`;
    }).join("");
  }

  // Gráfico
  const canvas=$("b8chart"); if(!canvas) return;
  if(b8ChartInstance){ b8ChartInstance.destroy(); b8ChartInstance=null; }

  const bgColors=años.map((_,i)=>i===IDX_2026?"#FFC526":"#1f5f91");
  const borderColors=años.map((_,i)=>i===IDX_2026?"#d9a000":"#123a5c");
  const labelColors=años.map((_,i)=>i===IDX_2026?"#92400e":"#123a5c");
  const tickColors=años.map((_,i)=>i===IDX_2026?"#92400e":"#374151");

  const labelsPlugin={
    id:"b8Labels",
    afterDatasetsDraw(chart){
      const{ctx,data,scales:{y}}=chart; ctx.save();
      data.datasets[0].data.forEach((val,i)=>{
        if(!val) return;
        const meta=chart.getDatasetMeta(0), bar=meta.data[i];
        ctx.fillStyle=labelColors[i];
        ctx.font="bold 11px 'Barlow Condensed',sans-serif";
        ctx.textAlign="center";
        ctx.fillText(fmtNum(val),bar.x,bar.y-6);
      });
      ctx.restore();
    }
  };

  b8ChartInstance=new Chart(canvas,{
    type:"bar",
    data:{
      labels:años.map(String),
      datasets:[{data:valores,backgroundColor:bgColors,borderColor:borderColors,borderWidth:1.5,borderRadius:6}]
    },
    options:{
      responsive:true,maintainAspectRatio:false,
      plugins:{
        legend:{display:false},
        tooltip:{callbacks:{label:ctx=>"Devengado: "+fmtNum(ctx.parsed.y)}}
      },
      scales:{
        x:{grid:{display:false},ticks:{color:ctx=>tickColors[ctx.index]||"#374151",
           font:{family:"'Barlow Condensed',sans-serif",size:13,weight:"600"}}},
        y:{grid:{color:"#f3f4f6"},ticks:{color:"#6b7280",
           font:{family:"Barlow,sans-serif",size:11},callback:v=>fmtCompacto(v)}}
      }
    },
    plugins:[labelsPlugin]
  });
}

// ── Eventos ───────────────────────────────────────────────────
document.getElementById("dropzone").addEventListener("click",()=>document.getElementById("file").click());
document.getElementById("dropzone").addEventListener("keydown",e=>{
  if(e.key==="Enter"||e.key===" "){ e.preventDefault(); document.getElementById("file").click(); }
});
document.getElementById("file").addEventListener("change",e=>{procesarArchivos(e.target.files);e.target.value="";});
document.getElementById("pick").addEventListener("click",()=>document.getElementById("file").click());
document.getElementById("clear").addEventListener("click",()=>{
  datos={};cargados=new Set();actualizarFileList();render();
});
const dz=document.getElementById("dropzone");
["dragenter","dragover"].forEach(ev=>dz.addEventListener(ev,e=>{e.preventDefault();dz.classList.add("drag");}));
["dragleave","drop"].forEach(ev=>dz.addEventListener(ev,e=>{e.preventDefault();dz.classList.remove("drag");}));
dz.addEventListener("drop",e=>{ if(e.dataTransfer.files.length) procesarArchivos(e.dataTransfer.files); });

// ── PDF ───────────────────────────────────────────────────────
let pdfLibsCargadas=false;
async function cargarLibsPDF(){
  if(pdfLibsCargadas) return;
  await new Promise((res,rej)=>{const s=document.createElement("script");s.src="https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js";s.onload=res;s.onerror=rej;document.head.appendChild(s);});
  await new Promise((res,rej)=>{const s=document.createElement("script");s.src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js";s.onload=res;s.onerror=rej;document.head.appendChild(s);});
  pdfLibsCargadas=true;
}
async function exportarPDF(){
  const btn=document.querySelector(".btn-pdf");
  if(btn){btn.innerHTML='<span>⏳ Generando PDF...</span>';btn.disabled=true;}
  try{
    await cargarLibsPDF();
    const{jsPDF}=window.jspdf;

    // A4 landscape en mm y en px a 96dpi
    const PW_MM=297, PH_MM=210;
    const MARGIN_MM=8;
    const AREA_W=PW_MM-MARGIN_MM*2, AREA_H=PH_MM-MARGIN_MM*2;

    const pdf=new jsPDF({orientation:"landscape",unit:"mm",format:"a4"});
    const slides=document.querySelectorAll(".slide");
    let primera=true;

    for(const slide of slides){
      // Capturar el slide a escala 2x para buena resolución
      const canvas=await html2canvas(slide,{
        scale:2,
        useCORS:true,
        logging:false,
        backgroundColor:"#ffffff",
        // Forzar el ancho del slide al ancho real del elemento
        width:slide.scrollWidth,
        height:slide.scrollHeight,
        windowWidth:slide.scrollWidth,
      });

      const imgData=canvas.toDataURL("image/jpeg",0.93);

      // Calcular dimensiones para ajustar el slide al área útil de la página
      // manteniendo proporción y sin deformar
      const imgW=canvas.width, imgH=canvas.height;
      const scaleW=AREA_W/imgW, scaleH=AREA_H/imgH;
      const scale=Math.min(scaleW,scaleH);
      const drawW=imgW*scale, drawH=imgH*scale;

      // Centrar en la página
      const offsetX=MARGIN_MM+(AREA_W-drawW)/2;
      const offsetY=MARGIN_MM+(AREA_H-drawH)/2;

      if(!primera) pdf.addPage("a4","landscape");
      pdf.addImage(imgData,"JPEG",offsetX,offsetY,drawW,drawH);
      primera=false;
    }

    const hoy=new Date();
    const fecha=`${hoy.getFullYear()}${String(hoy.getMonth()+1).padStart(2,"0")}${String(hoy.getDate()).padStart(2,"0")}`;
    pdf.save(`MPCH_Egresos_${fecha}.pdf`);

  }catch(e){
    console.error("PDF:",e);
    alert("Error al generar el PDF. Intente nuevamente.");
  }finally{
    if(btn){
      btn.innerHTML=`<svg width="14" height="14" viewBox="0 0 24 24" fill="none"
        stroke="currentColor" stroke-width="2.2">
        <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
        <polyline points="14 2 14 8 20 8"/>
        <line x1="16" y1="13" x2="8" y2="13"/>
        <line x1="16" y1="17" x2="8" y2="17"/>
      </svg> Exportar PDF`;
      btn.disabled=false;
    }
  }
}

// ── Inicio ────────────────────────────────────────────────────
rellenarFechas(); // fechas inmediatas al cargar la página
fetch("data/historico_egresos.json?"+Date.now())
  .then(r=>r.json())
  .then(data=>{B8_HIST=data;renderB8();})
  .catch(()=>console.warn("[MPCH-EG] historico_egresos.json no encontrado"));

autoCargar();
