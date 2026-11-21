/* Reporte diario de ejecución — INABIF
 * Arma en el navegador el PPTX "REPORTE_EJECUCIÓN_INABIF_dd.mm.aaaa.pptx" (19 láminas):
 *   1 carátula · 2 histórico PIM vs devengado · 3 MIMP por UE · 4 INABIF ordinario y FEN ·
 *   5 MIMP FEN por UE · 6 unidades funcionales · 7-18 una por unidad · 19 recursos FEN
 * Láminas 2-5: datos/consulta_amigable.json (lo actualiza GitHub Actions).
 * Láminas 6-19: el Excel del SIAF enriquecido (hoja SheetGasto), que se lee aquí y no se sube a ningún lado.
 * Mismo formato que reporte_consulta_amigable.py y reporte_unidades.py.
 */
(() => {
  "use strict";

  // ── Estilo (scraper_inabif.py · clase C) ──────────────────────────────
  const C = { rojo: "B51F35", navy: "0A3A70", texto: "45536D", gris: "8A94A6", linea: "E6EAF0",
    blanco: "FFFFFF", fondo: "F0F3F7", borde: "D9DEE6", negro: "000000" };
  const FUENTE = "Open Sans";
  const X0 = 0.35, ANCHO = 9.3;            // margen y ancho de la nota al pie
  const XT = 0.2, ANCHO_T = 9.6;           // tabla y franja roja
  const Y_FIN = 5.3;                       // borde inferior de la tabla
  const PESOS_U = [3.3, 1.0, 1.05, 0.74, 1.05, 0.74, 1.05, 0.74, 1.2];      // reporte_unidades.py
  const PESOS_CA = [2.4, 1.25, 1.25, 0.62, 1.25, 0.62, 1.25, 0.62, 1.25];   // reporte_consulta_amigable.py
  const CABECERAS = ["PIM", "CERTIFICADO", "%", "COMPROMISO", "%", "DEVENGADO", "%", "SALDO POR CERTIFICAR"];
  // Imágenes: si está img/logos.js (versión local, doble clic) van embebidas; si no, se leen de img/
  const BASE = new URL("img/", document.currentScript.src).href;
  const IMG = Object.fromEntries(Object.entries({
    fondo: "fondo_portada.jpg", logo_color: "logo_mimp_inabif_color.png",
    pr_color: "logo_peru_responde_color.png", mimp_gris: "logo_mimp_gris.png",
    inabif_gris: "logo_inabif_gris.png", pr_gris: "logo_peru_responde_gris.png",
  }).map(([k, f]) => [k, (window.LOGOS_DIARIO && window.LOGOS_DIARIO[k]) ? { data: window.LOGOS_DIARIO[k] } : { path: BASE + f }]));
  const imagen = (s, k, x, y, w, h) => s.addImage({ ...IMG[k], x, y, w, h });

  // ── Datos del Excel (reporte_unidades.py) ─────────────────────────────
  const UNIDADES = [
    ["UPPM", "Unidad de Planeamiento, Presupuesto y Modernización"],
    ["DE", "Dirección Ejecutiva"],
    ["UA", "Unidad de Administración"],
    ["URH", "Unidad de Recursos Humanos"],
    ["UCIM", "Unidad de Comunicación e Imagen"],
    ["OCI", "Órgano de Control Institucional"],
    ["UTI", "Unidad de Tecnologías de la Información"],
    ["UAJ", "Unidad de Asesoría Jurídica"],
    ["UAE", "Unidad de Asistencia Económica y Acompañamiento"],
    ["UDIF", "Unidad de Desarrollo Integral de las Familias"],
    ["UFCT", "Unidad de Fortalecimiento de Servicios y Coordinación Territorial"],
    ["USPNNA", "Unidad de Servicios de Protección de Niñas, Niños y Adolescentes"],
  ];
  const NOMBRES = Object.fromEntries(UNIDADES);
  const CONCEPTOS_VISIBLES = new Set([
    "ASISTENCIA ECONOMICA", "CAS", "CAP", "LOCADORES", "ALIMENTOS", "SEGURIDAD Y LIMPIEZA",
    "PASAJES Y VIATICOS", "SUMIN. MANTENIMIENTO", "DERECHOS ADMINISTRATIVOS / SENTENCIAS JUDICIALES",
    "INVERSIONES", "VEHICULOS",
  ]);
  const FEN_CONCEPTOS = {
    "213.- MANTENIMIENTO E INFRAESTRUCTURA": "Mantenimiento e infraestructura",
    "214.- FORTALECIMIENTO DE CAPACIDADES": "Fortalecimiento de capacidades",
    "215.- BIENES Y SERVICIOS": "Bienes y servicios",
  };
  const UE = "1160";

  // ── Montos ────────────────────────────────────────────────────────────
  const fila = (nombre, m = {}) => ({ nombre, pia: m.pia || 0, pim: m.pim || 0, cert: m.cert || 0, comp: m.comp || 0, dev: m.dev || 0 });
  const MONTOS = ["pia", "pim", "cert", "comp", "dev"];
  const sumarA = (a, b) => { for (const k of MONTOS) a[k] += b[k]; return a; };
  const sumar = (filas, nombre) => filas.reduce(sumarA, fila(nombre));
  const restar = (a, b, nombre) => { const r = fila(nombre); for (const k of MONTOS) r[k] = a[k] - b[k]; return r; };
  const pct = (f, x) => (f.pim ? x / f.pim : 0);
  const nf0 = new Intl.NumberFormat("en-US", { maximumFractionDigits: 0 });
  const miles = v => nf0.format(Math.abs(v) < 0.5 ? 0 : v);
  const fpct = v => `${(v * 100).toFixed(1)}%`;
  const p2 = n => String(n).padStart(2, "0");
  const fechaTxt = d => `${p2(d.getDate())}/${p2(d.getMonth() + 1)}/${d.getFullYear()}`;
  const mill = (v, dec = 2) => Math.abs(v) < 1e6 ? `S/ ${miles(v)}` : `S/ ${(v / 1e6).toFixed(dec).replace(/\.?0+$/, "")} mill`;
  const celdas = (f, etq) => [etq, miles(f.pim), miles(f.cert), fpct(pct(f, f.cert)), miles(f.comp),
    fpct(pct(f, f.comp)), miles(f.dev), fpct(pct(f, f.dev)), miles(f.pim - f.cert)];
  const porDevengado = filas => [...filas].sort((a, b) => pct(b.f, b.f.dev) - pct(a.f, a.f.dev));

  // ── Lectura del Excel ─────────────────────────────────────────────────
  function leerExcel(buf) {
    const wb = XLSX.read(buf, { type: "array", cellDates: true, dense: true });
    const hoja = wb.Sheets.SheetGasto;
    if (!hoja) throw new Error("El archivo no tiene la hoja SheetGasto (el Excel del SIAF con las columnas del área).");
    const aoa = XLSX.utils.sheet_to_json(hoja, { header: 1, defval: null, raw: true });
    const h = aoa.slice(0, 30).findIndex(r => r && r.some(c => String(c ?? "").trim() === "mto_pim"));
    if (h < 0) throw new Error("No encontré la fila de títulos (columna mto_pim).");
    const ix = {};
    aoa[h].forEach((c, i) => { const k = String(c ?? "").trim(); if (k && !(k in ix)) ix[k] = i; });
    const faltan = ["UNIDAD FUNCIONAL (SIGLAS)", "CONDICION 2", "mto_pia", "mto_pim", "mto_certificado", "mto_compro_anual"]
      .filter(c => !(c in ix));
    if (faltan.length) throw new Error("Al Excel le faltan las columnas: " + faltan.join(", "));
    const num = v => (typeof v === "number" && isFinite(v) ? v : 0);
    const meses = Array.from({ length: 12 }, (_, i) => ix["mto_devenga_" + p2(i + 1)]).filter(i => i != null);
    const unidades = new Map(), conceptos = new Map(), fen = new Map();
    let fecha = null, sinUnidad = 0;
    for (let i = h + 1; i < aoa.length; i++) {
      const r = aoa[i];
      if (!r || (r[ix.mto_pim] == null && r[ix.mto_pia] == null)) continue;
      let sig = String(r[ix["UNIDAD FUNCIONAL (SIGLAS)"]] ?? "").trim();
      const con = String(r[ix["CONDICION 2"]] ?? "").trim() || "OTROS";
      const m = { pia: num(r[ix.mto_pia]), pim: num(r[ix.mto_pim]), cert: num(r[ix.mto_certificado]),
        comp: num(r[ix.mto_compro_anual]), dev: meses.reduce((a, j) => a + num(r[j]), 0) };
      if (!fecha && "FEC_REP" in ix && r[ix.FEC_REP] instanceof Date) fecha = r[ix.FEC_REP];
      if (!sig) { sinUnidad += m.pim; sig = "SIN UNIDAD"; }
      if (!unidades.has(sig)) { unidades.set(sig, fila(sig)); conceptos.set(sig, new Map()); }
      sumarA(unidades.get(sig), m);
      const cs = conceptos.get(sig);
      if (!cs.has(con)) cs.set(con, fila(con));
      sumarA(cs.get(con), m);
      const tc = "TIPO CENTRO" in ix ? String(r[ix["TIPO CENTRO"]] ?? "").trim() : "";
      if (tc in FEN_CONCEPTOS) {
        if (!fen.has(tc)) fen.set(tc, fila(FEN_CONCEPTOS[tc]));
        sumarA(fen.get(tc), m);
      }
    }
    if (!unidades.size) throw new Error("La hoja SheetGasto no tiene filas de datos.");
    const avisos = [];
    if (sinUnidad) avisos.push(`Hay registros sin UNIDAD FUNCIONAL (PIM S/ ${miles(sinUnidad)}).`);
    const otras = [...unidades.keys()].filter(s => !(s in NOMBRES));
    if (otras.length) avisos.push(`Unidades no previstas: ${otras.join(", ")} (van al final).`);
    if (!fen.size) avisos.push("El Excel no trae los conceptos FEN (TIPO CENTRO 213/214/215): no habrá lámina FEN.");
    return {
      unidades, conceptos, fecha,
      fen: Object.keys(FEN_CONCEPTOS).filter(k => fen.has(k)).map(k => fen.get(k)),
      avisos,
    };
  }

  function agruparConceptos(cs) {
    const out = new Map();
    for (const [nombre, f] of cs) {
      const k = CONCEPTOS_VISIBLES.has(nombre) ? nombre : "OTROS";
      if (!out.has(k)) out.set(k, fila(k));
      sumarA(out.get(k), f);
    }
    return porDevengado([...out.values()].filter(f => Math.abs(f.pim) > 0.5 || Math.abs(f.dev) > 0.5).map(f => ({ f, nombre: f.nombre })));
  }

  // ── Componentes de lámina ─────────────────────────────────────────────
  function banda(s, alto) {      // franja blanca superior con las esquinas de abajo redondeadas
    s.addShape("roundRect", { x: 0, y: -0.25, w: 10, h: alto + 0.25, fill: { color: C.blanco }, rectRadius: 0.16667 * alto });
  }

  function interior(pptx) {
    const s = pptx.addSlide();
    s.background = { color: C.fondo };
    banda(s, 0.563);
    imagen(s, "mimp_gris", 0.591, -0.031, 1.823, 0.682);
    imagen(s, "inabif_gris", 2.544, 0.096, 1.078, 0.371);
    imagen(s, "pr_gris", 7.439, -0.058, 1.672, 0.691);
    return s;
  }

  function portada(pptx, hoy) {
    const s = pptx.addSlide();
    s.background = { color: C.rojo };
    imagen(s, "fondo", 0, 0, 10, 5.625);
    banda(s, 0.681);
    imagen(s, "logo_color", 0.511, 0.029, 3.391, 0.681);
    imagen(s, "pr_color", 7.152, -0.079, 2.065, 0.853);
    s.addShape("line", { x: 0.885, y: 3.708, w: 6.468, h: 0, line: { color: C.blanco, width: 1.5 } });
    s.addText([
      { text: "AVANCE DE EJECUCIÓN PRESUPUESTAL", options: { breakLine: true } },
      { text: `AL ${hoy}` },
    ], { x: 0.753, y: 1.51, w: 7.185, h: 2.02, fontFace: FUENTE, fontSize: 36, bold: true, color: C.blanco,
      align: "center", valign: "top", margin: [7.2, 7.2, 7.2, 7.2] });
    s.addText([
      { text: "Programa Integral Nacional para el Bienestar Familiar - INABIF", options: { bold: true, breakLine: true } },
      { text: "Unidad de Planeamiento, Presupuesto y Modernización", options: { bold: false } },
    ], { x: 0.403, y: 3.846, w: 9.077, h: 0.875, fontFace: FUENTE, fontSize: 20, color: C.blanco,
      align: "left", valign: "top", margin: [7.2, 7.2, 7.2, 7.2] });
  }

  function encabezadoRojo(s, texto) {
    s.addText(texto, { x: XT, y: 0.68, w: ANCHO_T, h: 0.42, fill: { color: C.rojo }, fontFace: "Arial",
      fontSize: texto.length <= 62 ? 18 : 15, bold: true, color: C.blanco, align: "center", valign: "middle",
      margin: [7.2, 7.2, 3.6, 3.6] });
  }

  function nota(s, texto) {
    s.addText(texto, { x: X0, y: 5.36, w: ANCHO, h: 0.16, fontFace: FUENTE, fontSize: 5.5, color: C.gris,
      valign: "top", margin: 0 });
  }

  function tarjeta(s, x, y, w, h, radio = 0.03) {
    s.addShape("roundRect", { x, y, w, h, fill: { color: C.blanco }, rectRadius: radio * Math.min(w, h) });
  }

  const SIN = { type: "none" };
  const linea = (pt, color) => ({ type: "solid", pt, color });

  /* Tabla que llena el alto libre: con pocas filas, filas más altas y letra más grande.
     filas: [{f, nombre}] · cabFija: alto fijo de la cabecera (el sobrante va a las filas) */
  function tabla(s, etiqueta, filas, total, y, { filaMax = 0.44, fontMax = 10.5, pesos = PESOS_U, cabFija = null, yFin = Y_FIN } = {}) {
    const n = filas.length + 2;
    const rh = Math.min(filaMax, (yFin - y - 0.28) / n);
    const font = Math.max(7, Math.min(fontMax, Math.round(rh * 27 * 2) / 2));
    const fCab = Math.min(font + 0.5, 8.5);
    tarjeta(s, XT, y, ANCHO_T, 0.1 + n * rh + 0.18);
    const rowH = Array(n).fill(rh);
    if (cabFija && rh > cabFija) {
      const extra = (rh - cabFija) / (n - 1);
      rowH[0] = cabFija;
      for (let i = 1; i < n; i++) rowH[i] = rh + extra;
    }
    const tw = ANCHO_T - 0.3, suma = pesos.reduce((a, b) => a + b, 0);
    const base = { fontFace: FUENTE, valign: "middle", margin: [0, 0.03, 0, 0.03] };
    const cab = [etiqueta, ...CABECERAS].map(t => ({ text: t, options: { ...base, fontFace: "Arial", fontSize: fCab,
      bold: true, color: C.blanco, fill: { color: C.navy }, align: "center", border: [SIN, SIN, linea(0.5, C.linea), SIN] } }));
    const cuerpo = filas.map(({ f, nombre }, i) => {
      const ultima = i === filas.length - 1;
      return celdas(f, nombre).map((t, j) => ({ text: t, options: { ...base, fontSize: font, bold: j === 6,
        color: j ? C.navy : C.texto, align: j ? "right" : "left",
        border: [SIN, SIN, ultima ? linea(0.75, C.negro) : linea(0.5, C.linea), SIN] } }));
    });
    const tot = celdas(total, "Total").map((t, j) => ({ text: t, options: { ...base, fontSize: font + 0.5, bold: true,
      color: j ? C.navy : C.texto, align: j ? "right" : "left", border: [linea(0.75, C.negro), SIN, SIN, SIN] } }));
    s.addTable([cab, ...cuerpo, tot], { x: XT + 0.18, y: y + 0.1, w: tw, colW: pesos.map(p => tw * p / suma), rowH, autoPage: false });
  }

  // tablas de la Consulta Amigable (reporte_consulta_amigable.py · tabla_llena)
  function tablaLlena(s, etiqueta, filas, total, yFin = Y_FIN) {
    const grande = ["pim", "cert", "comp", "dev"].some(k => Math.abs(total[k]) >= 1e9);
    tabla(s, etiqueta, porDevengado(filas), total, 1.18,
      { filaMax: 1.2, fontMax: grande ? 9.5 : 10.5, pesos: PESOS_CA, cabFija: 0.42, yFin });
  }

  // Nota de la lámina "Presupuesto ordinario y FEN": de dónde sale el FEN (NM 330)
  const NOTA_FEN = [
    { text: "Nota: ", options: { bold: true } },
    { text: "La RM N.° 545-2026-MIMP asignó a INABIF " },
    { text: "S/ 13,615,627", options: { bold: true } },
    { text: ", de los cuales " },
    { text: "S/ 10,900,474", options: { bold: true } },
    { text: " se destinan a la actividad FEN, para acciones de contingencia, y " },
    { text: "S/ 2,715,153", options: { bold: true } },
    { text: " forman parte del presupuesto ordinario y se distribuyen de la siguiente manera: S/ 2,463,076 para el proyecto de inversión con CUI 2197584 «Mejoramiento y ampliación de los servicios de protección integral a adolescentes mujeres en el CAR Ermelinda Carrera» y S/ 252,077 para atención al ciudadano y a la familia en emergencia social." },
  ];
  function notaFen(s, y = 4.76) {
    s.addText(NOTA_FEN.map(p => ({ text: p.text, options: { ...(p.options || {}) } })),
      { x: XT, y, w: ANCHO_T, h: Y_FIN - y + 0.02, fontFace: FUENTE, fontSize: 8, color: C.texto, valign: "middle",
        fill: { color: C.blanco }, line: { color: C.borde, width: 0.75 }, margin: [7.2, 7.2, 3.6, 3.6] });
  }

  function indicadoresFen(s, t, y = 1.2, h = 0.76) {
    const datos = [[mill(t.pim, 1), "PIM", `S/ ${miles(t.pim)}`], [fpct(pct(t, t.cert)), "Certificado", mill(t.cert)],
      [fpct(pct(t, t.comp)), "Compromiso", mill(t.comp)], [fpct(pct(t, t.dev)), "Devengado", mill(t.dev)]];
    const w = ANCHO_T / datos.length;
    datos.forEach(([valor, nombre, monto], i) => {
      s.addText([
        { text: valor, options: { fontSize: 24, bold: true, color: C.navy, breakLine: true } },
        { text: nombre, options: { fontSize: 11, color: C.texto, breakLine: true } },
        { text: monto, options: { fontSize: 8, color: C.gris } },
      ], { x: XT + i * w, y, w, h, fontFace: "Arial", align: "center", valign: "middle", fill: { color: C.blanco },
        line: { color: C.borde, width: 1 }, margin: [7.2, 7.2, 1.44, 1.44], paraSpaceBefore: 0, paraSpaceAfter: 0 });
    });
  }

  const ROJO_ETQ = "C0FFEE";      // marca temporal: color de las etiquetas de la 2.a serie (se corrige al final)
  function grafico(s, serie, y = 1.2) {
    const h = Y_FIN - y;
    tarjeta(s, XT, y, ANCHO_T, h);
    const etiquetas = serie.map(a => `${a.anio}\nEjec. ${fpct(a.pim ? a.dev / a.pim : 0)}`);
    const r1 = v => Math.round(v / 1e5) / 10;
    s.addChart("bar", [
      { name: "PIM", labels: etiquetas, values: serie.map(a => r1(a.pim)) },
      { name: "Devengado", labels: etiquetas, values: serie.map(a => r1(a.dev)) },
    ], { x: XT + 0.15, y: y + 0.1, w: ANCHO_T - 0.3, h: h - 0.2, barDir: "col", barGrouping: "clustered",
      barGapWidthPct: 60, barOverlapPct: -10, chartColors: [C.navy, C.rojo],
      showLegend: true, legendPos: "t", legendFontSize: 10, legendFontFace: FUENTE, legendColor: C.texto,
      showValue: true, dataLabelFormatCode: "#,##0.0", dataLabelPosition: "outEnd", dataLabelFontSize: 11,
      dataLabelFontBold: true, dataLabelFontFace: FUENTE, dataLabelColor: ROJO_ETQ,
      valAxisHidden: true, valGridLine: { style: "none" }, catGridLine: { style: "none" },
      catAxisLabelFontSize: 11, catAxisLabelFontBold: true, catAxisLabelFontFace: FUENTE, catAxisLabelColor: C.texto,
      catAxisLineShow: true, catAxisLineColor: C.linea, catAxisMajorTickMark: "none", valAxisMajorTickMark: "none" });
  }

  // PptxGenJS pone el mismo color a las etiquetas de todas las series: se separa navy / rojo en el XML
  async function colorearEtiquetas(buf) {
    const zip = await JSZip.loadAsync(buf);
    for (const nombre of Object.keys(zip.files).filter(n => /^ppt\/charts\/chart\d+\.xml$/.test(n))) {
      let xml = await zip.file(nombre).async("string");
      if (!xml.includes(ROJO_ETQ)) continue;
      let i = 0;
      xml = xml.replace(/<c:ser>[\s\S]*?<\/c:ser>/g, ser => ser.split(ROJO_ETQ).join(i++ === 0 ? C.navy : C.rojo));
      xml = xml.split(ROJO_ETQ).join(C.navy);
      zip.file(nombre, xml);
    }
    return zip.generateAsync({ type: "blob", mimeType: "application/vnd.openxmlformats-officedocument.presentationml.presentation" });
  }

  // ── Presentación completa ─────────────────────────────────────────────
  const nombreUE = f => (f.codigo ? `${f.codigo}: ${f.nombre}` : f.nombre);
  const esInabif = f => f.codigo === UE || String(f.codigo).endsWith("-" + UE);

  function construir(ex, ca, fecha) {
    const pptx = new PptxGenJS();
    pptx.layout = "LAYOUT_16x9";
    pptx.title = "Avance de ejecución presupuestal INABIF";
    const hoy = fechaTxt(fecha);
    const notaU = `Fuente: SIAF – MEF (Reporte de Gasto) · ${hoy}`;

    portada(pptx, hoy);

    // 2-5 · Consulta Amigable
    if (ca) {
      const c = ca.consultado ? new Date(ca.consultado) : null;
      const cuando = c ? `${fechaTxt(c)} ${p2(c.getHours())}:${p2(c.getMinutes())}` : hoy;
      const notaCA = `Fuente: Consulta Amigable – MEF · consultado el ${cuando}`;
      const ues = ca.ues.map(u => ({ ...fila(u.nombre, u), codigo: u.codigo }));
      const inabif = ues.find(esInabif) || fila("INABIF");
      const fen = fila("Presupuesto FEN", ca.fen || {});

      let s = interior(pptx);
      const serie = ca.historico || [];
      if (serie.length) {
        encabezadoRojo(s, `INABIF – PIM VS. DEVENGADO ${serie[0].anio}–${serie[serie.length - 1].anio} (S/ MILLONES)`);
        grafico(s, serie);
        nota(s, `${notaCA}. ${serie[serie.length - 1].anio}: ejecución acumulada.`);
      }

      s = interior(pptx);
      encabezadoRojo(s, `MIMP – EJECUCIÓN POR UNIDAD EJECUTORA – ${hoy}`);
      tablaLlena(s, "UNIDAD EJECUTORA", ues.map(f => ({ f, nombre: nombreUE(f) })), sumar(ues, "Total"));
      nota(s, notaCA);

      s = interior(pptx);
      encabezadoRojo(s, `INABIF – PRESUPUESTO ORDINARIO Y FEN – ${hoy}`);
      const ordinario = restar(inabif, fen, "Presupuesto ordinario");
      tablaLlena(s, "PRESUPUESTO", [{ f: ordinario, nombre: ordinario.nombre }, { f: fen, nombre: fen.nombre }], inabif, 4.66);
      notaFen(s);
      nota(s, notaCA);

      if (ca.fen_ues && ca.fen_ues.length) {
        const fu = ca.fen_ues.map(u => ({ ...fila(u.nombre, u), codigo: u.codigo }));
        s = interior(pptx);
        encabezadoRojo(s, `MIMP – RECURSOS FEN POR UNIDAD EJECUTORA – ${hoy}`);
        tablaLlena(s, "UNIDAD EJECUTORA", fu.map(f => ({ f, nombre: nombreUE(f) })), sumar(fu, "Total"));
        nota(s, `Actividad 5006412: Intervenciones extraordinarias del Fenómeno El Niño - FEN. ${notaCA}`);
      }
    }

    // 6 · unidades funcionales
    const orden = [...UNIDADES.map(u => u[0]).filter(u => ex.unidades.has(u)), ...[...ex.unidades.keys()].filter(u => !(u in NOMBRES))];
    let s = interior(pptx);
    encabezadoRojo(s, `EJECUCIÓN PRESUPUESTAL POR UNIDADES FUNCIONALES – ${hoy}`);
    const filasU = orden.map(u => ({ f: ex.unidades.get(u), nombre: NOMBRES[u] || u }));
    tabla(s, "UNIDAD FUNCIONAL", filasU, sumar(filasU.map(x => x.f), "Total"), 1.18);
    nota(s, notaU);

    // 7-18 · una lámina por unidad
    for (const u of orden) {
      s = interior(pptx);
      encabezadoRojo(s, `${u} – EJECUCIÓN PRESUPUESTAL – ${hoy}`);
      s.addText(NOMBRES[u] || u, { x: XT + 0.02, y: 1.12, w: ANCHO_T, h: 0.16, fontFace: FUENTE, fontSize: 8.5,
        bold: true, color: C.navy, valign: "top", margin: 0 });
      tabla(s, "CONCEPTO", agruparConceptos(ex.conceptos.get(u)), ex.unidades.get(u), 1.34);
      nota(s, notaU);
    }

    // 19 · recursos FEN
    if (ex.fen.length) {
      s = interior(pptx);
      encabezadoRojo(s, `EJECUCIÓN DE RECURSOS FEN – ${hoy}`);
      const t = sumar(ex.fen, "Total");
      indicadoresFen(s, t);
      tabla(s, "CONCEPTO", ex.fen.map(f => ({ f, nombre: f.nombre })), t, 2.08, { filaMax: 0.6 });
      nota(s, notaU);
    }
    return pptx;
  }

  async function generar(ex, ca, fecha) {
    const pptx = construir(ex, ca, fecha);
    const buf = await pptx.write({ outputType: "arraybuffer" });
    return colorearEtiquetas(buf);
  }

  const nombreArchivo = fecha => `REPORTE_EJECUCIÓN_INABIF_${p2(fecha.getDate())}.${p2(fecha.getMonth() + 1)}.${fecha.getFullYear()}.pptx`;

  window.ReporteDiario = { leerExcel, generar, nombreArchivo, fechaTxt };
})();
