const pptxgen = require("pptxgenjs");
const { execSync } = require("child_process");
const fs = require("fs");
const path = require("path");

// ── Configuración ──────────────────────────────────────────────
const config = JSON.parse(fs.readFileSync("config_mes.json", "utf8"));

const datos = JSON.parse(fs.readFileSync("docs/datos_presentacion.json", "utf8"));

// ── Trayectoria histórica ─────────────────────────────────────
const trayectoriaHistorica = [
  {mes:"Dic-24",toe:67.55},{mes:"Ene-25",toe:64.72},{mes:"Feb-25",toe:50.56},
  {mes:"Mar-25",toe:69.31},{mes:"Abr-25",toe:67.08},{mes:"May-25",toe:76.97},
  {mes:"Jun-25",toe:81.05},{mes:"Jul-25",toe:92.31},{mes:"Ago-25",toe:89.08},
  {mes:"Sep-25",toe:92.15},{mes:"Oct-25",toe:94.26},{mes:"Nov-25",toe:94.61},
  {mes:"Dic-25",toe:91.11},
];
const trayectoria = [...trayectoriaHistorica, datos.trayectoria_mes_actual];

// ── Colores Kinexperience ──────────────────────────────────────
const C = {
  marino:    "080B3D",
  azul:      "009CDE",
  verde:     "85C4B3",
  palido:    "EAF1FA",
  blanco:    "FFFFFF",
  gris:      "64748B",
  grisCiaro: "F1F5F9",
  rojo:      "E53E3E",
  amarillo:  "F6AD55",
  verdeOk:   "38A169",
};

// ── Logo ───────────────────────────────────────────────────────
const logoClaro = path.resolve("LOGO-FONDO-CLARO.png");
const logoBlanco = path.resolve("LOGO-BLANCO.png");
const logoNegro = path.resolve("LOGO-NEGRO.png");

// ── Helper: color fila según TOE ──────────────────────────────
function colorFila(toe, meta) {
  if (toe >= meta)           return "E6F4EA"; // verde suave
  if (toe >= meta * 0.85)   return "FFF8E1"; // amarillo suave
  return "FFEBEE";                            // rojo suave
}

function colorToe(toe, meta) {
  if (toe >= meta)           return C.verdeOk;
  if (toe >= meta * 0.85)   return "D97706";
  return C.rojo;
}

// ── Helpers shadow/layout ──────────────────────────────────────
const makeShadow = () => ({ type: "outer", blur: 8, offset: 3, angle: 135, color: "000000", opacity: 0.10 });

// ══════════════════════════════════════════════════════════════
// CONSTRUIR PRESENTACIÓN
// ══════════════════════════════════════════════════════════════
let pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.3" × 7.5"

// ── SLIDE 1: Portada ──────────────────────────────────────────
{
  let s = pres.addSlide();
  s.background = { color: C.marino };

  // Banda lateral izquierda azul
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 0.18, h: 7.5, fill: { color: C.azul }, line: { color: C.azul } });

  // Logo blanco
  s.addImage({ path: logoBlanco, x: 0.5, y: 0.4, w: 1.8, h: 1.8 });

  // Título
  s.addText("Termómetro", { x: 0.5, y: 2.4, w: 8, h: 1.4, fontSize: 72, bold: true, color: C.blanco, fontFace: "Calibri" });
  s.addText("Kinexperience", { x: 0.5, y: 3.7, w: 8, h: 0.9, fontSize: 40, bold: false, color: C.azul, fontFace: "Calibri" });
  s.addText(config.presentacion.mes_label, { x: 0.5, y: 4.6, w: 8, h: 0.6, fontSize: 22, color: C.verde, fontFace: "Calibri" });

  // Línea decorativa
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 5.5, w: 4, h: 0.05, fill: { color: C.verde }, line: { color: C.verde } });
}

// ── SLIDE 2: El mes en números (SIN TOE grupal) ───────────────
{
  let s = pres.addSlide();
  s.background = { color: C.palido };

  s.addText("El mes en números", { x: 0.5, y: 0.3, w: 9, h: 0.7, fontSize: 28, bold: true, color: C.marino, fontFace: "Calibri" });

  // Tarjeta 1 — Capacidad del equipo
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.8, w: 3.8, h: 3.5, fill: { color: C.blanco }, line: { color: C.blanco }, shadow: makeShadow() });
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.8, w: 3.8, h: 0.18, fill: { color: C.marino }, line: { color: C.marino } });
  s.addText("CAPACIDAD DEL EQUIPO", { x: 0.5, y: 1.8, w: 3.8, h: 0.18, fontSize: 9, bold: true, color: C.blanco, align: "center", fontFace: "Calibri", margin: 0 });
  s.addText(datos.potencial_total.toString(), { x: 0.5, y: 2.0, w: 3.8, h: 1.6, fontSize: 80, bold: true, color: C.marino, align: "center", fontFace: "Calibri" });
  s.addText("sesiones posibles este mes", { x: 0.5, y: 3.6, w: 3.8, h: 0.5, fontSize: 11, color: C.gris, align: "center", fontFace: "Calibri" });

  // Tarjeta 2 — Sesiones programadas
  s.addShape(pres.shapes.RECTANGLE, { x: 4.75, y: 1.8, w: 3.8, h: 3.5, fill: { color: C.blanco }, line: { color: C.blanco }, shadow: makeShadow() });
  s.addShape(pres.shapes.RECTANGLE, { x: 4.75, y: 1.8, w: 3.8, h: 0.18, fill: { color: C.azul }, line: { color: C.azul } });
  s.addText("SESIONES PROGRAMADAS (TOP)", { x: 4.75, y: 1.8, w: 3.8, h: 0.18, fontSize: 9, bold: true, color: C.blanco, align: "center", fontFace: "Calibri", margin: 0 });
  s.addText(datos.potencial_top.toString(), { x: 4.75, y: 2.0, w: 3.8, h: 1.6, fontSize: 80, bold: true, color: C.azul, align: "center", fontFace: "Calibri" });
  s.addText("sesiones agendadas", { x: 4.75, y: 3.6, w: 3.8, h: 0.5, fontSize: 11, color: C.gris, align: "center", fontFace: "Calibri" });

  // Tarjeta 3 — Frase del mes
  s.addShape(pres.shapes.RECTANGLE, { x: 9.0, y: 1.8, w: 3.8, h: 3.5, fill: { color: C.marino }, line: { color: C.marino }, shadow: makeShadow() });
  s.addText("META DEL MES", { x: 9.0, y: 1.6, w: 3.8, h: 0.4, fontSize: 11, bold: true, color: C.verde, align: "center", fontFace: "Calibri" });
  s.addText(config.presentacion.frase_motivacional, { x: 9.0, y: 2.2, w: 3.8, h: 2.0, fontSize: 22, bold: true, color: C.blanco, align: "center", valign: "middle", fontFace: "Calibri" });
  s.addImage({ path: logoClaro, x: 11.8, y: 0.1, w: 1.3, h: 1.3 });
}

// ── SLIDE 3: Tabla del equipo ─────────────────────────────────
{
  let s = pres.addSlide();
  s.background = { color: C.palido };
  s.addText("Tabla del equipo", { x: 0.5, y: 0.2, w: 9, h: 0.6, fontSize: 28, bold: true, color: C.marino, fontFace: "Calibri" });
  s.addText(config.presentacion.mes_label, { x: 0.5, y: 0.75, w: 9, h: 0.35, fontSize: 13, color: C.gris, fontFace: "Calibri" });

  const meta = config.meta_TOE * 100;
  // 11 columnas en 12.7": Kine(2.5) + Pacientes(1.1) + Eval(1.1) + Realizadas(1.0) + Susp(1.05) + Recup(1.05) + Cancel(1.0) + Efectivas(1.0) + TOP(0.95) + TOE(0.95) + Tendencia(1.0)
  const colW = [2.5, 1.1, 1.1, 1.0, 1.05, 1.05, 1.0, 1.0, 0.95, 0.95, 1.0];
  const hdr = { fill: { color: C.marino }, color: C.blanco, bold: true, fontSize: 8.5, fontFace: "Calibri", align: "center", valign: "middle" };

  let rows = [[
    { text: "KINE",             options: { ...hdr, align: "left" } },
    { text: "PACIENTES",        options: hdr },
    { text: "EVAL. INGRESO",    options: hdr },
    { text: "REALIZADAS",       options: hdr },
    { text: "SUSPENDIDAS",      options: hdr },
    { text: "RECUPERADAS",      options: hdr },
    { text: "CANCELADAS",       options: hdr },
    { text: "EFECTIVAS",        options: hdr },
    { text: "TOP",              options: hdr },
    { text: "TOE",              options: hdr },
    { text: "TENDENCIA",        options: hdr },
  ]];

  // Ordenar por TOE descendente
  const kinesOrdenados = [...datos.kines].sort((a, b) => b.toe - a.toe);

  kinesOrdenados.forEach(k => {
    const bg = colorFila(k.toe, meta);
    const tc = colorToe(k.toe, meta);
    const cell = (txt, extra = {}) => ({ text: txt, options: { fontSize: 11, fontFace: "Calibri", align: "center", valign: "middle", fill: { color: bg }, color: C.marino, ...extra } });
    rows.push([
      cell(k.nombre,            { align: "left",  bold: true }),
      cell(k.pacientes.toString()),
      cell(k.evaluaciones.toString()),
      cell(k.realizadas.toString()),
      cell(k.suspendidas.toString()),
      cell(k.recuperadas.toString()),
      cell(k.canceladas.toString(), { color: k.canceladas < 0 ? C.verde : C.marino }),
      cell(k.efectivas.toString(),  { bold: true }),
      cell(k.top.toFixed(1) + "%"),
      cell(k.toe.toFixed(1) + "%", { bold: true, color: tc }),
      cell(k.tendencia,         { fontSize: 14, bold: true, color: k.tendencia === "↑" ? C.verdeOk : C.rojo }),
    ]);
  });

  s.addTable(rows, { x: 0.3, y: 1.6, w: 12.7, colW, rowH: 0.52, border: { pt: 0.5, color: "E2E8F0" } });
  s.addImage({ path: logoClaro, x: 11.8, y: 0.2, w: 1.0, h: 1.0 });
}

// ── SLIDE 4: TOP vs TOE individual ───────────────────────────
{
  let s = pres.addSlide();
  s.background = { color: C.palido };
  s.addImage({ path: logoClaro, x: 11.8, y: 0.1, w: 1.3, h: 1.3 });
  s.addText("Programación vs Efectivas", { x: 0.5, y: 0.2, w: 10, h: 0.6, fontSize: 28, bold: true, color: C.marino, fontFace: "Calibri" });
  s.addText("TOP vs TOE · " + config.presentacion.mes_label, { x: 0.5, y: 0.75, w: 10, h: 0.35, fontSize: 13, color: C.gris, fontFace: "Calibri" });

  const kinesOrdenados = [...datos.kines].sort((a, b) => b.toe - a.toe);
  s.addChart(pres.charts.BAR, [
    { name: "TOP (Programadas)", labels: kinesOrdenados.map(k => k.nombre), values: kinesOrdenados.map(k => k.top / 100) },
    { name: "TOE (Efectivas)",   labels: kinesOrdenados.map(k => k.nombre), values: kinesOrdenados.map(k => k.toe / 100) },
  ], {
    x: 0.4, y: 1.5, w: 12.5, h: 5.6,
    barDir: "bar",
    barGrouping: "clustered",
    chartColors: [C.marino, C.azul],
    chartArea: { fill: { color: C.palido } },
    dataLabelFormatCode: "0%",
    showValue: true,
    valAxisMaxVal: 1.1,
    valAxisNumFmt: "0%",
    catAxisLabelColor: C.marino,
    valAxisLabelColor: C.gris,
    valGridLine: { color: "E2E8F0", size: 0.5 },
    catGridLine: { style: "none" },
    showLegend: true,
    legendPos: "t",
    legendFontSize: 11,
    dataLabelColor: C.marino,
    dataLabelFontSize: 12,
  });
}

// ── SLIDE 5: Evolución TOE ────────────────────────────────────
{
  let s = pres.addSlide();
  s.background = { color: C.palido };
  s.addImage({ path: logoClaro, x: 11.8, y: 0.1, w: 1.3, h: 1.3 });
  s.addText("Evolución TOE individual", { x: 0.5, y: 0.2, w: 10, h: 0.6, fontSize: 28, bold: true, color: C.marino, fontFace: "Calibri" });
  s.addText("Mes anterior vs " + config.presentacion.mes_label, { x: 0.5, y: 0.75, w: 10, h: 0.35, fontSize: 13, color: C.gris, fontFace: "Calibri" });

  // Datos enero como "mes anterior" para prueba
  const anterior = { "Patricio Orrego": 97.6, "José Aguilar": 100.0, "Mauricio Arce": 106.3,
    "Guillermo Silva": 71.4, "Katalina Correa": 48.6, "Daniela Jaque": 77.1,
    "Marcia Reveco": 30.0, "Sebastián de la Peña": 68.6 };

  const kinesOrdenados = [...datos.kines].sort((a, b) => b.toe - a.toe);
  s.addChart(pres.charts.BAR, [
    { name: "Mes anterior", labels: kinesOrdenados.map(k => k.nombre), values: kinesOrdenados.map(k => (anterior[k.nombre] || 0) / 100) },
    { name: config.presentacion.mes_label,     labels: kinesOrdenados.map(k => k.nombre), values: kinesOrdenados.map(k => k.toe / 100) },
  ], {
    x: 0.4, y: 1.5, w: 12.5, h: 5.6,
    barDir: "bar",
    barGrouping: "clustered",
    chartColors: [C.marino, C.azul],
    chartArea: { fill: { color: C.blanco } },
    showValue: true,
    valAxisMaxVal: 1.2,
    valAxisNumFmt: "0%",
    dataLabelFormatCode: "0%",
    catAxisLabelColor: C.marino,
    valAxisLabelColor: C.gris,
    valGridLine: { color: "E2E8F0", size: 0.5 },
    catGridLine: { style: "none" },
    showLegend: true,
    legendPos: "t",
    legendFontSize: 11,
    dataLabelColor: C.marino,
    dataLabelFontSize: 12,
  });
}

// ── SLIDE 6: Suspendidas / Recuperadas ───────────────────────
{
  let s = pres.addSlide();
  s.background = { color: C.palido };
  s.addImage({ path: logoClaro, x: 11.8, y: 0.1, w: 1.3, h: 1.3 });
  s.addText("Suspendidas y Recuperadas", { x: 0.5, y: 0.2, w: 10, h: 0.6, fontSize: 28, bold: true, color: C.marino, fontFace: "Calibri" });
  s.addText(config.presentacion.mes_label + "  ·  Valor negativo = recuperaciones superan suspensiones", { x: 0.5, y: 0.75, w: 12, h: 0.35, fontSize: 11, color: C.gris, fontFace: "Calibri" });

  const nombres = datos.kines.map(k => k.nombre.split(" ")[0]);

  s.addChart(pres.charts.BAR, [
    { name: "Suspendidas", labels: nombres, values: datos.kines.map(k => k.suspendidas) },
    { name: "Canceladas",  labels: nombres, values: datos.kines.map(k => k.canceladas) },
  ], {
    x: 0.3, y: 1.5, w: 6.3, h: 5.6,
    barDir: "col", barGrouping: "clustered",
    chartColors: [C.marino, C.rojo],
    chartArea: { fill: { color: C.palido } },
    showValue: true, showLegend: true, legendPos: "t",
    valGridLine: { color: "E2E8F0", size: 0.5 },
    catGridLine: { style: "none" },
    dataLabelFontSize: 12,
  });

  s.addChart(pres.charts.BAR, [
    { name: "Recuperadas", labels: nombres, values: datos.kines.map(k => k.recuperadas) },
  ], {
    x: 6.7, y: 1.5, w: 6.3, h: 5.6,
    barDir: "col",
    chartColors: [C.verde],
    chartArea: { fill: { color: C.palido } },
    showValue: true, showLegend: true, legendPos: "t",
    valGridLine: { color: "E2E8F0", size: 0.5 },
    catGridLine: { style: "none" },
    dataLabelFontSize: 12,
  });
}

// ── SLIDE 7: Distribución de carga ───────────────────────────
{
  let s = pres.addSlide();
  s.background = { color: C.palido };
  s.addImage({ path: logoClaro, x: 11.8, y: 0.1, w: 1.3, h: 1.3 });
  s.addText("Distribución de carga", { x: 0.5, y: 0.2, w: 10, h: 0.6, fontSize: 28, bold: true, color: C.marino, fontFace: "Calibri" });
  s.addText("Sesiones efectivas por kine · " + config.presentacion.mes_label, { x: 0.5, y: 0.75, w: 10, h: 0.35, fontSize: 13, color: C.gris, fontFace: "Calibri" });

  const kinesOrdenados = [...datos.kines].sort((a, b) => b.efectivas - a.efectivas);
  const totalSesiones = kinesOrdenados.reduce((s, k) => s + k.efectivas, 0);
  const ideal = totalSesiones / kinesOrdenados.length;

  s.addChart(pres.charts.BAR, [
    { name: "Sesiones efectivas", labels: kinesOrdenados.map(k => k.nombre), values: kinesOrdenados.map(k => k.efectivas) },
  ], {
    x: 0.4, y: 1.5, w: 12.5, h: 5.5,
    barDir: "bar",
    chartColors: [C.azul],
    chartArea: { fill: { color: C.blanco } },
    showValue: true,
    catAxisLabelColor: C.marino,
    valAxisLabelColor: C.gris,
    valGridLine: { color: "E2E8F0", size: 0.5 },
    catGridLine: { style: "none" },
    showLegend: false,
    dataLabelFontSize: 10,
    dataLabelColor: C.marino,
  });

  s.addText(`Distribución ideal: ${Math.round(ideal)} sesiones / kine`, {
    x: 0.4, y: 6.8, w: 8, h: 0.4, fontSize: 11, color: C.gris, italic: true, fontFace: "Calibri"
  });
}

// ── SLIDE 8: Meta del mes (primero la meta, sin resultado) ────
{
  let s = pres.addSlide();
  s.background = { color: C.marino };

  s.addImage({ path: logoBlanco, x: 11.8, y: 0.1, w: 1.3, h: 1.3 });
  s.addText("Meta " + config.presentacion.mes_label, { x: 0.5, y: 0.4, w: 9, h: 0.7, fontSize: 32, bold: true, color: C.blanco, fontFace: "Calibri" });

  // Tarjeta TOE meta
  s.addShape(pres.shapes.RECTANGLE, { x: 1.5, y: 1.4, w: 4.2, h: 4.5, fill: { color: C.blanco }, line: { color: C.blanco }, shadow: makeShadow() });
  s.addText("TOE", { x: 1.5, y: 1.4, w: 4.2, h: 0.6, fontSize: 18, bold: true, color: C.marino, align: "center", valign: "middle", fontFace: "Calibri" });
  s.addText((config.meta_TOE * 100) + "%", { x: 1.5, y: 2.2, w: 4.2, h: 2.5, fontSize: 90, bold: true, color: C.azul, align: "center", fontFace: "Calibri" });
  s.addText("tasa de ocupación efectiva", { x: 1.5, y: 4.8, w: 4.2, h: 0.5, fontSize: 11, color: C.gris, align: "center", fontFace: "Calibri" });

  // Tarjeta TOP meta
  s.addShape(pres.shapes.RECTANGLE, { x: 7.6, y: 1.4, w: 4.2, h: 4.5, fill: { color: C.marino }, line: { color: C.azul, pt: 2 }, shadow: makeShadow() });
  s.addText("TOP", { x: 7.6, y: 1.4, w: 4.2, h: 0.6, fontSize: 18, bold: true, color: C.verde, align: "center", valign: "middle", fontFace: "Calibri" });
  s.addText((config.meta_TOP * 100) + "%", { x: 7.6, y: 2.2, w: 4.2, h: 2.5, fontSize: 90, bold: true, color: C.blanco, align: "center", fontFace: "Calibri" });
  s.addText("tasa de ocupación programada", { x: 7.6, y: 4.8, w: 4.2, h: 0.5, fontSize: 11, color: C.verde, align: "center", fontFace: "Calibri" });
}

// ── SLIDE 9: Resultado del mes (el reveal) ────────────────────
{
  let s = pres.addSlide();
  s.background = { color: C.marino };

  const toeOk  = datos.toe_grupal  >= config.meta_TOE * 100;
  const topOk  = datos.top_grupal  >= config.meta_TOP * 100;

  s.addImage({ path: logoBlanco, x: 11.8, y: 0.1, w: 1.3, h: 1.3 });
  s.addText("Resultado " + config.presentacion.mes_label, { x: 0.5, y: 0.4, w: 9, h: 0.7, fontSize: 32, bold: true, color: C.blanco, fontFace: "Calibri" });

  // Tarjeta TOE resultado
  const toeBg = toeOk ? C.verdeOk : C.rojo;
  s.addShape(pres.shapes.RECTANGLE, { x: 1.5, y: 1.4, w: 4.2, h: 4.8, fill: { color: C.blanco }, line: { color: C.blanco }, shadow: makeShadow() });
  s.addShape(pres.shapes.RECTANGLE, { x: 1.5, y: 1.4, w: 4.2, h: 0.5, fill: { color: toeBg }, line: { color: toeBg } });
  s.addText("TOE", { x: 1.5, y: 1.4, w: 4.2, h: 0.5, fontSize: 16, bold: true, color: C.blanco, align: "center", valign: "middle", fontFace: "Calibri", margin: 0 });
  s.addText(datos.toe_grupal.toFixed(1) + "%", { x: 1.5, y: 2.1, w: 4.2, h: 2.2, fontSize: 86, bold: true, color: C.marino, align: "center", fontFace: "Calibri" });
  s.addText(toeOk ? "✓ Meta alcanzada" : "✗ Bajo la meta (" + config.meta_TOE * 100 + "%)", {
    x: 1.5, y: 4.4, w: 4.2, h: 0.6, fontSize: 13, bold: true, color: toeOk ? C.verdeOk : C.rojo, align: "center", fontFace: "Calibri"
  });

  // Tarjeta TOP resultado
  const topBg = topOk ? C.verdeOk : C.rojo;
  s.addShape(pres.shapes.RECTANGLE, { x: 7.6, y: 1.4, w: 4.2, h: 4.8, fill: { color: C.blanco }, line: { color: C.blanco }, shadow: makeShadow() });
  s.addShape(pres.shapes.RECTANGLE, { x: 7.6, y: 1.4, w: 4.2, h: 0.5, fill: { color: topBg }, line: { color: topBg } });
  s.addText("TOP", { x: 7.6, y: 1.4, w: 4.2, h: 0.5, fontSize: 16, bold: true, color: C.blanco, align: "center", valign: "middle", fontFace: "Calibri", margin: 0 });
  s.addText(datos.top_grupal.toFixed(1) + "%", { x: 7.6, y: 2.1, w: 4.2, h: 2.2, fontSize: 86, bold: true, color: C.marino, align: "center", fontFace: "Calibri" });
  s.addText(topOk ? "✓ Meta alcanzada" : "✗ Bajo la meta (" + config.presentacion.meta_top_mes + "%)", {
    x: 7.6, y: 4.4, w: 4.2, h: 0.6, fontSize: 13, bold: true, color: topOk ? C.verdeOk : C.rojo, align: "center", fontFace: "Calibri"
  });
}

// ── SLIDE 10: Trayectoria histórica ──────────────────────────
{
  let s = pres.addSlide();
  s.background = { color: C.palido };
  s.addImage({ path: logoClaro, x: 11.8, y: 0.1, w: 1.3, h: 1.3 });
  s.addText("Trayectoria histórica", { x: 0.5, y: 0.2, w: 10, h: 0.6, fontSize: 28, bold: true, color: C.marino, fontFace: "Calibri" });
  s.addText("TOE mensual del equipo", { x: 0.5, y: 0.75, w: 10, h: 0.35, fontSize: 13, color: C.gris, fontFace: "Calibri" });

  s.addChart(pres.charts.LINE, [
    { name: "TOE", labels: trayectoria.map(d => d.mes), values: trayectoria.map(d => d.toe / 100) },
  ], {
    x: 0.4, y: 1.5, w: 12.5, h: 5.6,
    lineSize: 3,
    lineSmooth: false,
    chartColors: [C.azul],
    chartArea: { fill: { color: C.palido } },
    showValue: true,
    valAxisMinVal: 0.4,
    valAxisMaxVal: 1.1,
    valAxisNumFmt: "0%",
    dataLabelFormatCode: "0%",
    dataLabelPosition: "t",
    catAxisLabelColor: C.marino,
    valAxisLabelColor: C.gris,
    valGridLine: { color: "E2E8F0", size: 0.5 },
    catGridLine: { style: "none" },
    showLegend: false,
    dataLabelFontSize: 12,
    dataLabelColor: C.marino,
  });
}

// ── SLIDE 11: Zoom — Últimos meses ───────────────────────────
{
  let s = pres.addSlide();
  s.background = { color: C.palido };
  s.addImage({ path: logoClaro, x: 11.8, y: 0.1, w: 1.3, h: 1.3 });
  s.addText("Zoom — Últimos meses", { x: 0.5, y: 0.2, w: 10, h: 0.6, fontSize: 28, bold: true, color: C.marino, fontFace: "Calibri" });
  s.addText("Comparativa mismo período año anterior", { x: 0.5, y: 0.75, w: 10, h: 0.35, fontSize: 13, color: C.gris, fontFace: "Calibri" });

  const zoomLabels = ["Dic", "Ene", "Feb", "Mar"];
  s.addChart(pres.charts.LINE, [
    {
      name: "Mismo período 2024-25",
      labels: zoomLabels,
      values: [91.11 / 100, 64.72 / 100, 50.56 / 100, 69.31 / 100],
    },
    {
      name: "2025-26",
      labels: zoomLabels,
      values: [91.11 / 100, 73.97 / 100, 58.93 / 100, datos.trayectoria_mes_actual.toe / 100],
    },
  ], {
    x: 0.4, y: 1.5, w: 12.5, h: 5.6,
    lineSize: 3,
    lineSmooth: false,
    chartColors: [C.verde, C.azul],
    chartArea: { fill: { color: C.palido } },
    showValue: true,
    valAxisMinVal: 0.4,
    valAxisMaxVal: 1.1,
    valAxisNumFmt: "0%",
    dataLabelFormatCode: "0%",
    dataLabelPosition: "t",
    catAxisLabelColor: C.marino,
    valAxisLabelColor: C.gris,
    valGridLine: { color: "E2E8F0", size: 0.5 },
    catGridLine: { style: "none" },
    showLegend: true,
    legendPos: "t",
    legendFontSize: 11,
    dataLabelFontSize: 12,
    dataLabelColor: C.marino,
  });
}

// ── SLIDE 12: Meta próximo mes ────────────────────────────────
{
  let s = pres.addSlide();
  s.background = { color: C.marino };

  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 13.3, h: 0.18, fill: { color: C.azul }, line: { color: C.azul } });
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 7.32, w: 13.3, h: 0.18, fill: { color: C.verde }, line: { color: C.verde } });

  s.addImage({ path: logoBlanco, x: 11.8, y: 0.3, w: 1.3, h: 1.3 });
  s.addText("Meta", { x: 0.5, y: 0.5, w: 6, h: 0.8, fontSize: 42, bold: true, color: C.blanco, fontFace: "Calibri" });
  s.addText(config.presentacion.meta_siguiente_nombre, { x: 0.5, y: 1.2, w: 6, h: 0.7, fontSize: 30, color: C.azul, fontFace: "Calibri" });

  s.addText(config.presentacion.meta_toe_siguiente + "%", { x: 1.0, y: 2.2, w: 4.5, h: 3.0, fontSize: 110, bold: true, color: C.azul, align: "center", fontFace: "Calibri" });
  s.addText("TOE", { x: 1.0, y: 5.1, w: 4.5, h: 0.6, fontSize: 18, color: C.gris, align: "center", fontFace: "Calibri" });

  s.addText(config.presentacion.meta_top_siguiente + "%", { x: 7.3, y: 2.2, w: 4.5, h: 3.0, fontSize: 110, bold: true, color: C.verde, align: "center", fontFace: "Calibri" });
  s.addText("TOP", { x: 7.3, y: 5.1, w: 4.5, h: 0.6, fontSize: 18, color: C.gris, align: "center", fontFace: "Calibri" });
}

// ── SLIDE 13: Cierre ──────────────────────────────────────────
{
  let s = pres.addSlide();
  s.background = { color: C.marino };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 0.18, h: 7.5, fill: { color: C.azul }, line: { color: C.azul } });
  s.addShape(pres.shapes.RECTANGLE, { x: 13.12, y: 0, w: 0.18, h: 7.5, fill: { color: C.verde }, line: { color: C.verde } });
  s.addImage({ path: logoBlanco, x: 4.65, y: 1.8, w: 4.0, h: 4.0 });
}

// ── Exportar ───────────────────────────────────────────────────
const nombreArchivo = `Termometro_${config.presentacion.mes_label.replace(/ /g, "_")}.pptx`;
pres.writeFile({ fileName: nombreArchivo }).then(() => {
  console.log(`\n✅ Presentación generada: ${nombreArchivo}`);
});
