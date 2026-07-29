const pptxgen = require("pptxgenjs");
const fs = require("fs");
const DATA = JSON.parse(fs.readFileSync("proceso/_data.json", "utf8"));

// --- Paleta tomada del PDF modelo ---
const VERDE = "A2AD00", AZUL = "284973", CAB = "002060",
      TIT = "3B5A9E", TITBORDE = "6B88C0", ZEBRA = "E3E7F0";
const IMG = "modelo-archivos/imagenes/";
const FONT_T = "FS EMERIC";      // títulos (fuente de marca)
const FONT_B = "Aptos Narrow";   // tablas (según especificación)

const meses = ["enero","febrero","marzo","abril","mayo","junio","julio","agosto","septiembre","octubre","noviembre","diciembre"];
function fechaLarga(){ const p=DATA.fecha.split("-"); const d=new Date(+p[2],+p[1]-1,+p[0]); return `${d.getDate()} de ${meses[d.getMonth()]} de ${d.getFullYear()}`; }

function barraInferior(pres, s){
  s.addShape(pres.ShapeType.rect, { x:0, y:7.15, w:2.2, h:0.35, fill:{color:VERDE} });
  s.addShape(pres.ShapeType.rect, { x:2.2, y:7.15, w:11.133, h:0.35, fill:{color:CAB} });
}

function baseDeco(pres, s, titulo){
  s.background = { color:"FFFFFF" };
  s.addImage({ path: IMG+"LOGO_metrogas.png", x:0.35, y:0.28, w:1.6, h:0.47 });
  s.addImage({ path: IMG+"LOGO_nttdata.png",  x:11.35, y:0.30, w:1.55, h:0.42 });
  if (titulo){
    s.addShape(pres.ShapeType.roundRect, { x:4.0, y:0.28, w:5.33, h:0.62, rectRadius:0.1,
      fill:{color:"FFFFFF"}, line:{color:TITBORDE, width:1.5} });
    s.addText(titulo, { x:4.0, y:0.28, w:5.33, h:0.62, align:"center", valign:"middle",
      fontFace:FONT_T, fontSize:22, bold:true, color:TIT });
  }
  barraInferior(pres, s);
}

function cell(text, opt){ return { text:String(text), options:Object.assign({ fontFace:FONT_B, valign:"middle" }, opt) }; }

function portada(pres, subtitulo){
  const s = pres.addSlide(); s.background={color:"FFFFFF"};
  s.addImage({ path: IMG+"LOGO_metrogas.png", x:0.4, y:0.4, w:2.1, h:0.62 });
  s.addImage({ path: IMG+"LOGO_nttdata.png", x:6.35, y:0.45, w:1.55, h:0.42 });
  s.addImage({ path: IMG+"IMG_corporativa.png", x:8.6, y:0, w:4.733, h:7.5, sizing:{type:"cover",w:4.733,h:7.5} });
  s.addText("Seguimiento Semanal", { x:0.6, y:2.5, w:7.6, h:1.0, fontFace:FONT_T, fontSize:40, bold:true, color:VERDE });
  s.addText(subtitulo, { x:0.6, y:3.5, w:7.6, h:1.2, fontFace:FONT_T, fontSize:32, bold:true, color:AZUL });
  s.addText(fechaLarga(), { x:0.6, y:6.6, w:6.0, h:0.4, fontFace:FONT_B, fontSize:14, color:AZUL });
  barraInferior(pres, s);
}

function agenda(pres, torre){
  const s = pres.addSlide(); baseDeco(pres, s, "Agenda");
  const items = [`Volumetría tickets ${torre}`, "Detalle tickets por Módulo", "Detalle Desarrollos por Módulo"];
  const rows = items.map((t,i)=>[
    cell(i+1, { fill:{color:CAB}, color:"FFFFFF", bold:true, align:"center", fontSize:16 }),
    cell(t,   { fill:{color: i%2? "FFFFFF":ZEBRA}, color:CAB, align:"left", fontSize:16, margin:6 }),
  ]);
  s.addTable(rows, { x:1.2, y:2.4, w:8.5, colW:[0.7,7.8], rowH:0.7, border:{type:"solid",color:"FFFFFF",pt:2} });
  s.addImage({ path: IMG+"IMG_corporativa.png", x:10.2, y:2.5, w:2.6, h:2.6, rounding:true });
}

function volumetria(pres, torre, modulos){
  const s = pres.addSlide(); baseDeco(pres, s, "Volumetría Tickets");
  const head = ["Módulo","Reunión Anterior","Reunión Actual","Diferencia"].map(t=>
    cell(t, { fill:{color:CAB}, color:"FFFFFF", bold:true, align:"center", fontSize:12 }));
  const rows=[head]; let total=0;
  modulos.forEach((m,i)=>{ total+=m.tickets_total; const bg=i%2? "FFFFFF":ZEBRA;
    rows.push([
      cell(m.nombre, { fill:{color:CAB}, color:"FFFFFF", align:"left", fontSize:12, margin:4 }),
      cell("N/A", { fill:{color:bg}, color:"888888", align:"center", fontSize:12 }),
      cell(m.tickets_total, { fill:{color:bg}, align:"center", fontSize:12 }),
      cell("N/A", { fill:{color:bg}, color:"888888", align:"center", fontSize:12 }),
    ]); });
  rows.push([
    cell("TOTAL", { fill:{color:AZUL}, color:"FFFFFF", bold:true, align:"left", fontSize:12, margin:4 }),
    cell("N/A", { fill:{color:AZUL}, color:"FFFFFF", align:"center", fontSize:12 }),
    cell(total, { fill:{color:AZUL}, color:"FFFFFF", bold:true, align:"center", fontSize:12 }),
    cell("N/A", { fill:{color:AZUL}, color:"FFFFFF", align:"center", fontSize:12 }),
  ]);
  s.addTable(rows, { x:2.6, y:2.2, w:8.1, colW:[2.4,1.9,1.9,1.9], border:{type:"solid",color:"D9D9D9",pt:1}, rowH:0.38 });
  s.addText("N/A en 'Reunión Anterior' por ser la primera ejecución (sin log previo).",
    { x:2.6, y:6.55, w:8.1, h:0.3, fontFace:FONT_B, fontSize:10, italic:true, color:"888888" });
}

function chunk(a,n){ const o=[]; for(let i=0;i<a.length;i+=n) o.push(a.slice(i,i+n)); return o; }

// navyCols = nº de columnas iniciales pintadas en navy (Estado/Sociedad o Status)
function tablaSlides(pres, titulo, cols, colW, filas, fontSize, maxRows, navyCols){
  const paginas = filas.length ? chunk(filas, maxRows) : [[]];
  let last=null;
  paginas.forEach((pg, idx)=>{
    const s = pres.addSlide();
    baseDeco(pres, s, titulo + (paginas.length>1?` (${idx+1}/${paginas.length})`:""));
    const head = cols.map(c=> cell(c, { fill:{color:CAB}, color:"FFFFFF", bold:true, align:"center", fontSize:fontSize }));
    const rows=[head];
    pg.forEach((f, ri)=>{
      const bg = ri%2? "FFFFFF": ZEBRA;
      rows.push(f.map((v,ci)=>{
        if (ci < navyCols) return cell(v, { fill:{color:CAB}, color:"FFFFFF", bold:true, align:"center", fontSize:fontSize, margin:3 });
        return cell(v, { fill:{color:bg}, color:"111111", align:"left", fontSize:fontSize, margin:3 });
      }));
    });
    s.addTable(rows, { x:0.4, y:1.15, w:12.53, colW:colW, border:{type:"solid",color:"FFFFFF",pt:1}, rowH:0.32, autoPage:false });
    last=s;
  });
  return last;
}

// paleta de colores por estado (para dona + leyenda)
const STATUS_COLOR = {
  "Ingresado":"284973", "En validación de evaluación":"5A9E2F", "Bloqueado":"E8663B",
  "En PRD":"6785C1", "Entregado a Qas":"8FA9D0", "En ejecución":"F0A93B",
  "En planificación":"9DC3E6", "Enviado a evaluación":"C55A11", "En ajustes":"7F7F7F",
};
const PAL_FALLBACK = ["284973","5A9E2F","E8663B","6785C1","F0A93B","8FA9D0","C55A11","7F7F7F","9DC3E6","A2AD00"];
function colorFor(status, i){ return STATUS_COLOR[status] || PAL_FALLBACK[i % PAL_FALLBACK.length]; }

function donut(pres, slide, resumen){
  const data = resumen.filter(([e,c])=> e!=="Total general" && c>0);
  if (!data.length) return;
  const labels = data.map(d=>d[0]);
  const values = data.map(d=>d[1]);
  const colors = data.map((d,i)=> colorFor(d[0], i));
  slide.addChart(pres.ChartType.doughnut,
    [{ name:"Status", labels, values }],
    { x:4.6, y:4.95, w:3.4, h:2.0, holeSize:55,
      chartColors: colors, showLegend:false,
      showValue:false, showPercent:true, dataLabelPosition:"outEnd",
      dataLabelColor:"333333", dataLabelFontFace:FONT_B, dataLabelFontSize:9,
      showTitle:false });
  // leyenda (infográfico) a la derecha
  const leg = labels.map((l,i)=>[
    { text:"", options:{ fill:{color:colors[i]} } },
    { text:l, options:{ fontFace:FONT_B, fontSize:9, color:"111111", align:"left", valign:"middle", margin:2 } },
  ]);
  slide.addTable(leg, { x:8.2, y:5.0, w:4.4, colW:[0.22,4.18], rowH:0.26, border:{type:"none"} });
}

function resumenBox(slide, resumen){
  const rows=[[
    cell("Estado", { fill:{color:CAB}, color:"FFFFFF", bold:true, fontSize:10, align:"left", margin:3 }),
    cell("Cant",   { fill:{color:CAB}, color:"FFFFFF", bold:true, fontSize:10, align:"center" }),
  ]];
  resumen.forEach(([e,c],i)=>{
    const tot = (i===resumen.length-1);
    const bg = tot? AZUL : (i%2? "FFFFFF": ZEBRA);
    rows.push([
      cell(e, { fill:{color:bg}, color:tot?"FFFFFF":"111111", bold:tot, fontSize:10, align:"left", margin:3 }),
      cell(c, { fill:{color:bg}, color:tot?"FFFFFF":"111111", bold:tot, fontSize:10, align:"center" }),
    ]);
  });
  // anclar por abajo: el cuadro crece hacia arriba y nunca cruza la barra inferior (y=7.15)
  const nRows = rows.length;              // header + estados + total
  const rowH = 0.28;
  const bottomY = 6.85;
  const topY = bottomY - nRows * rowH;
  slide.addText("Resumen por Status", { x:0.5, y:topY-0.3, w:3.4, h:0.26, fontFace:FONT_B, fontSize:11, bold:true, color:AZUL });
  slide.addTable(rows, { x:0.5, y:topY, w:3.4, colW:[2.7,0.7], border:{type:"solid",color:"FFFFFF",pt:1}, rowH:rowH, autoPage:false });
}

function cierre(pres){
  const s=pres.addSlide(); s.background={color:CAB};
  s.addImage({ path: IMG+"LOGO_metrogas.png", x:4.4, y:3.1, w:4.5, h:1.3 });
}

function buildDeck(deckKey, subtitulo, torre, outfile){
  const pres = new pptxgen();
  pres.defineLayout({ name:"W", width:13.333, height:7.5 });
  pres.layout = "W";
  portada(pres, subtitulo);
  agenda(pres, torre);
  const mods = DATA.decks[deckKey].modulos;
  volumetria(pres, torre, mods);
  const tCols=["Estado","Sociedad","Número","Resumen","Fecha de apertura","Usuario MTGS/GS00","Días Tranc"];
  const tW=[1.1,1.3,1.63,4.0,1.6,1.9,1.0];
  const dCols=["Status","Creado","Código","Tipo de req.","Título","Fecha plan QA","Fecha entrega QA","Solicitante"];
  const dW=[1.5,1.4,1.3,1.4,3.33,1.2,1.2,1.2];
  mods.forEach(m=>{
    const tf = m.tickets.map(t=>[t.estado,t.sociedad,t.numero,t.resumen,t.apertura,t.usuario,t.dias]);
    tablaSlides(pres, `Tickets módulo ${m.nombre}`, tCols, tW, tf, 9, 13, 2);   // 2 columnas navy
    const df = m.desarrollos.map(d=>[d.status,d.creado,d.codigo,d.tipo,d.titulo,d.plan_qa,d.entrega_qa,d.solicitante]);
    const last = tablaSlides(pres, `Desarrollos módulo ${m.nombre}`, dCols, dW, df, 9, 10, 1); // 1 columna navy
    resumenBox(last, m.resumen);
    donut(pres, last, m.resumen);
  });
  cierre(pres);
  return pres.writeFile({ fileName: outfile });
}

(async ()=>{
  const F = DATA.fecha;
  // carpeta de salida: 1er argumento; por defecto Ejecuciones/<fecha> (flujo real).
  // Para pruebas: node proceso/generar_ppt.js test/<fecha>
  const OUT = (process.argv[2] || `Ejecuciones/${F}`).replace(/\/$/, "");
  await buildDeck("comercial","Torre Comercial y Técnica","Torre Comercial y Técnica",
     `${OUT}/Seguimiento Semanal Torre Comercial y Técnico.pptx`);
  await buildDeck("corporativa","Torre Corporativa","Torre Corporativa",
     `${OUT}/Seguimiento Semanal Torre Corporativa.pptx`);
  console.log("PPTX generados en", OUT);
})();
