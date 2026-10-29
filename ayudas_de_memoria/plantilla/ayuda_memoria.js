// Genera una "Ayuda de memoria" con el formato 162,14,15-60 de la Unidad para las Víctimas.
// Uso: node plantilla/ayuda_memoria.js actividades/<archivo>.js
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, ImageRun,
  Header, AlignmentType, WidthType, ShadingType, BorderStyle, VerticalAlign,
  LevelFormat, PageNumber, HeightRule, PageBreak,
} = require("docx");

const BASE = path.join(__dirname, "..");
const GRIS = "BFBFBF";
const GRIS_OSCURO = "7F7F7F";
const borde = { style: BorderStyle.SINGLE, size: 4, color: "000000" };
const bordes = { top: borde, bottom: borde, left: borde, right: borde };

// ---------- Encabezado institucional ----------
function celdaHeader(children, opts = {}) {
  return new TableCell({
    borders: bordes,
    width: { size: opts.w, type: WidthType.DXA },
    verticalAlign: opts.va || VerticalAlign.CENTER,
    shading: opts.fill ? { fill: opts.fill, type: ShadingType.CLEAR, color: "auto" } : undefined,
    rowSpan: opts.rowSpan,
    margins: { left: 80, right: 80, top: 0, bottom: 0 },
    children,
  });
}
function txtH(text, o = {}) {
  return new Paragraph({
    alignment: o.align || AlignmentType.LEFT,
    children: Array.isArray(text) ? text : [new TextRun({ text, font: "Verdana", size: 16, bold: o.bold, color: o.color })],
  });
}

function encabezado() {
  const W = [3510, 5236, 2014];
  const logo = fs.readFileSync(path.join(BASE, "recursos/logo_uariv.png"));
  const v = (t) => new TextRun({ text: t, font: "Verdana", size: 16 });
  const tabla = new Table({
    width: { size: 10760, type: WidthType.DXA },
    columnWidths: W,
    indent: { size: -690, type: WidthType.DXA },
    rows: [
      new TableRow({
        height: { value: 780, rule: HeightRule.ATLEAST },
        children: [
          celdaHeader([new Paragraph({ alignment: AlignmentType.CENTER, children: [
            new ImageRun({ type: "png", data: logo, transformation: { width: 92, height: 87 } })] })],
            { w: W[0], fill: GRIS, rowSpan: 4 }),
          celdaHeader([txtH("FORMATO AYUDA DE MEMORIA", { bold: true, color: "FFFFFF", align: AlignmentType.CENTER })], { w: W[1], fill: GRIS }),
          celdaHeader([txtH("Código: 162,14,15-60")], { w: W[2] }),
        ],
      }),
      new TableRow({
        height: { value: 500, rule: HeightRule.ATLEAST },
        children: [
          celdaHeader([txtH("PROCESO DE GESTIÓN DOCUMENTAL", { align: AlignmentType.CENTER })], { w: W[1] }),
          celdaHeader([txtH("Versión: 01")], { w: W[2] }),
        ],
      }),
      new TableRow({
        height: { value: 200, rule: HeightRule.ATLEAST },
        children: [
          celdaHeader([txtH("MANUAL PARA LA ORGANIZACIÓN DE DOCUMENTOS FÍSICOS Y ELECTRÓNICOS", { align: AlignmentType.CENTER })],
            { w: W[1], rowSpan: 2 }),
          celdaHeader([txtH("Fecha: 16/11/2023")], { w: W[2] }),
        ],
      }),
      new TableRow({
        height: { value: 560, rule: HeightRule.ATLEAST },
        children: [
          celdaHeader([txtH([v("Página "), new TextRun({ children: [PageNumber.CURRENT], font: "Verdana", size: 16, bold: true }),
            v(" de "), new TextRun({ children: [PageNumber.TOTAL_PAGES], font: "Verdana", size: 16, bold: true })])],
            { w: W[2], va: VerticalAlign.TOP }),
        ],
      }),
    ],
  });
  return new Header({ children: [tabla, new Paragraph({ children: [] })] });
}

// ---------- Utilidades de cuerpo ----------
const A = (text, o = {}) => new TextRun({ text, font: "Arial", size: 24, ...o });
const parrafo = (runs, o = {}) => new Paragraph({
  alignment: o.align || AlignmentType.JUSTIFIED,
  spacing: { after: o.after ?? 160, line: 276 },
  indent: o.indent,
  keepNext: o.keepNext,
  children: (Array.isArray(runs) ? runs : [runs]).map((r) => (typeof r === "string" ? A(r) : r)),
});
// Texto con **negrita** simple
function rico(text) {
  return text.split(/(\*\*[^*]+\*\*)/).filter(Boolean).map((s) =>
    s.startsWith("**") ? A(s.slice(2, -2), { bold: true }) : A(s));
}
const seccion = (n, titulo) => new Paragraph({
  numbering: { reference: "secciones", level: 0 },
  spacing: { before: 240, after: 160 },
  keepNext: true,
  children: [A(titulo, { bold: true })],
});
const vineta = (text, nivel = 0) => new Paragraph({
  numbering: { reference: "vinetas", level: nivel },
  alignment: AlignmentType.JUSTIFIED,
  spacing: { after: 80, line: 276 },
  children: rico(text),
});

function tablaAsistentes(asistentes) {
  const W = [2600, 3700, 1900];
  const cel = (t, w, head) => new TableCell({
    borders: bordes,
    width: { size: w, type: WidthType.DXA },
    verticalAlign: VerticalAlign.CENTER,
    shading: head ? { fill: GRIS, type: ShadingType.CLEAR, color: "auto" } : undefined,
    margins: { left: 80, right: 80, top: 20, bottom: 20 },
    children: [new Paragraph({ alignment: head ? AlignmentType.CENTER : AlignmentType.LEFT,
      children: [A(t, { bold: head })] })],
  });
  return new Table({
    width: { size: 8200, type: WidthType.DXA },
    columnWidths: W,
    indent: { size: 280, type: WidthType.DXA },
    rows: [
      new TableRow({ tableHeader: true, children: [
        cel("Nombre", W[0], true), cel("Organización / entidad a la que pertenece", W[1], true), cel("Cargo", W[2], true)] }),
      ...asistentes.map((a) => new TableRow({ cantSplit: true, children: [
        cel(a.nombre, W[0]), cel(a.entidad, W[1]), cel(a.cargo, W[2])] })),
    ],
  });
}

function tablaControlCambios() {
  const W = [1500, 2200, 5100];
  const cel = (t, w, head) => new TableCell({
    borders: bordes,
    width: { size: w, type: WidthType.DXA },
    shading: head ? { fill: GRIS_OSCURO, type: ShadingType.CLEAR, color: "auto" } : undefined,
    margins: { left: 80, right: 80 },
    children: [new Paragraph({ alignment: AlignmentType.CENTER,
      children: [new TextRun({ text: t, font: "Calibri", size: 20, color: head ? "FFFFFF" : "000000" })] })],
  });
  return new Table({
    width: { size: 8800, type: WidthType.DXA },
    columnWidths: W,
    rows: [
      new TableRow({ children: [cel("Versión", W[0], true), cel("Fecha", W[1], true), cel("Descripción de la modificación", W[2], true)] }),
      new TableRow({ children: [cel("1", W[0]), cel("16/11/2023", W[1]), cel("Creación del documento", W[2])] }),
    ],
  });
}

function anexoImagen(an) {
  const buf = fs.readFileSync(path.join(BASE, an.archivo));
  // ancho/alto en píxeles del PNG
  const w = buf.readUInt32BE(16), h = buf.readUInt32BE(20);
  const ancho = an.ancho || 560; // ~6.25 pulgadas
  return [
    parrafo([A(an.titulo, { bold: true })], { align: AlignmentType.LEFT, after: 80, keepNext: true }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 120 }, keepNext: true, children: [
      new ImageRun({ type: "png", data: buf, transformation: { width: ancho, height: Math.round((ancho * h) / w) },
        altText: { title: an.titulo, description: an.titulo, name: an.titulo } })] }),
    ...(an.descripcion ? [parrafo([A(an.descripcion, { italics: true, size: 20 })], { align: AlignmentType.CENTER, after: 240 })] : []),
  ];
}

// ---------- Documento ----------
function construir(d) {
  const cuerpo = [];
  cuerpo.push(parrafo([A("Nombre del espacio:  ", { bold: true }), A(d.nombreEspacio)], { after: 200 }));
  cuerpo.push(parrafo([A("Objeto del espacio: ", { bold: true }), A(d.objetoEspacio || "")], { after: 0 }));
  cuerpo.push(parrafo([A("Fecha:  ", { bold: true }), A(d.fecha)], { after: 0, align: AlignmentType.LEFT }));
  cuerpo.push(parrafo([A("Hora:  ", { bold: true }), A(d.hora)], { after: 0, align: AlignmentType.LEFT }));
  cuerpo.push(parrafo([A("Lugar: ", { bold: true }), A(d.lugar)], { after: 120, align: AlignmentType.LEFT }));

  cuerpo.push(seccion(1, "Asistentes:"));
  cuerpo.push(tablaAsistentes(d.asistentes));

  cuerpo.push(seccion(2, "Objetivo:"));
  cuerpo.push(parrafo(rico(d.objetivo), { indent: { left: 720 } }));

  cuerpo.push(seccion(3, "Desarrollo de la actividad:"));
  d.introduccion.forEach((p) => cuerpo.push(parrafo(rico(p))));
  d.puntos.forEach((pt, i) => {
    cuerpo.push(parrafo([A(`${i + 1}. ${pt.titulo}`, { bold: true })], { align: AlignmentType.LEFT, keepNext: true, after: 120 }));
    (pt.parrafos || []).forEach((p) => cuerpo.push(parrafo(rico(p))));
    (pt.vinetas || []).forEach((v) => (typeof v === "string" ? cuerpo.push(vineta(v)) : (cuerpo.push(vineta(v.texto)), (v.sub || []).forEach((s) => cuerpo.push(vineta(s, 1))))));
    (pt.cierre || []).forEach((p) => cuerpo.push(parrafo(rico(p), { after: 160 })));
    cuerpo.push(new Paragraph({ children: [] }));
  });

  cuerpo.push(new Paragraph({ children: [new PageBreak()] }));
  cuerpo.push(seccion(4, "Anexos:"));
  d.anexos.forEach((an) => cuerpo.push(...anexoImagen(an)));

  // Firma
  const f = d.firma;
  // Bloque de firma en una tabla sin bordes para que firma y datos no se separen de página
  const sinBorde = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
  const lf = (k, v) => new Paragraph({ children: [new TextRun({ text: `${k}${v}`, font: "Arial Narrow", size: 20, bold: true })] });
  cuerpo.push(new Paragraph({ spacing: { before: 240 }, children: [] }));
  cuerpo.push(new Table({
    width: { size: 5000, type: WidthType.DXA },
    columnWidths: [5000],
    borders: { top: sinBorde, bottom: sinBorde, left: sinBorde, right: sinBorde, insideHorizontal: sinBorde, insideVertical: sinBorde },
    rows: [new TableRow({ cantSplit: true, children: [new TableCell({
      width: { size: 5000, type: WidthType.DXA },
      borders: { top: sinBorde, bottom: sinBorde, left: sinBorde, right: sinBorde },
      margins: { left: 0, right: 0 },
      children: [
        new Paragraph({ children: f.imagen ? [new ImageRun({ type: "png", data: fs.readFileSync(path.join(BASE, f.imagen)), transformation: { width: 74, height: 86 } })] : [] }),
        lf("Nombre: ", f.nombre), lf("Cargo/Número de Contrato:", f.contrato), lf("Cargo: ", f.cargo), lf("Dependencia: ", f.dependencia),
      ] })] })],
  }));

  cuerpo.push(new Paragraph({ numbering: { reference: "secciones", level: 0 }, spacing: { before: 360, after: 160 }, keepNext: true,
    children: [new TextRun({ text: "CONTROL DE CAMBIOS:", font: "Calibri", size: 22 })] }));
  cuerpo.push(tablaControlCambios());

  return new Document({
    creator: f.nombre,
    title: d.nombreEspacio,
    styles: { default: { document: { run: { font: "Arial", size: 24 } } } },
    numbering: { config: [
      { reference: "secciones", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 720, hanging: 360 } }, run: { bold: true, font: "Arial" } } }] },
      { reference: "vinetas", levels: [
        { level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } },
        { level: 1, format: LevelFormat.BULLET, text: "◦", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 1440, hanging: 360 } } } },
      ] },
    ] },
    sections: [{
      properties: { page: { size: { width: 12240, height: 15840 },
        margin: { top: 2600, bottom: 1300, left: 1700, right: 1250, header: 700, footer: 700 } } },
      headers: { default: encabezado() },
      children: cuerpo,
    }],
  });
}

if (require.main === module) {
  const archivo = path.resolve(process.argv[2]);
  const datos = require(archivo);
  const salida = path.join(BASE, "salida", `${datos.archivoSalida}.docx`);
  fs.mkdirSync(path.dirname(salida), { recursive: true });
  Packer.toBuffer(construir(datos)).then((b) => { fs.writeFileSync(salida, b); console.log(salida); });
}
module.exports = { construir };
