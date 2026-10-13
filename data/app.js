// ── CORTE HISTÓRICO (Rubro 08 y Rubro 09; el ranking usa CORTE_LABEL) ─ ÚNICO lugar a cambiar cada mes ──────────
// CORTE_LABEL   : texto que se ve en tarjetas y gráficos.
// CORTE_MES_MAX : último mes (0=Ene … 9=Oct … 11=Dic) en que 2026 sigue siendo
//                 la barra "a la fecha". Si queda por debajo del mes actual, la
//                 barra 2026 DESAPARECE (así pasó al entrar a octubre).
const CORTE_LABEL   = "Ene\u2013Oct";
const CORTE_MES_MAX = 9;

const $ = id => document.getElementById(id);

const fmt  = n => Math.round(n || 0).toLocaleString("es-PE");
const fmtS = n => "S/ " + fmt(n);

function fechaHoy() {
  const d = new Date();
  const dias  = ["domingo","lunes","martes","miércoles","jueves","viernes","sábado"];
  const meses = ["enero","febrero","marzo","abril","mayo","junio","julio",
                 "agosto","septiembre","octubre","noviembre","diciembre"];
  return `${dias[d.getDay()]}, ${d.getDate()} de ${meses[d.getMonth()]} de ${d.getFullYear()}`;
}

function fechaCorta() {
  const d = new Date();
  const meses = ["enero","febrero","marzo","abril","mayo","junio","julio",
                 "agosto","septiembre","octubre","noviembre","diciembre"];
  return `${d.getDate()} de ${meses[d.getMonth()]} de ${d.getFullYear()}`;
}

function colorSem(pct) {
  if (pct === null) return "#888";
  if (pct >= 70)   return "var(--verde)";
  if (pct >= 40)   return "var(--amarillo-s)";
  return "var(--rojo-s)";
}

function barraHTML(pct, colorFijo) {
  if (pct === null) {
    return '<span style="font-family:\'Barlow Condensed\';font-weight:700;color:#888;font-size:13px">N/A</span>';
  }
  const col = colorFijo || colorSem(pct);
  const w   = Math.min(100, pct).toFixed(1);
  return `<div class="av-wrap">
    <div class="av-bar"><span style="width:${w}%;background:${col}"></span></div>
    <span class="av-pct" style="color:${col}">${pct.toFixed(1)}%</span>
  </div>`;
}

function toNum(v) {
  if (v == null || v === "") return 0;
  const s = String(v).replace(/,/g, "").trim();
  const n = parseFloat(s);
  return isNaN(n) ? 0 : n;
}

function esFilaDato(row) {
  const c0 = String(row[0] || "").trim();
  if (!c0) return false;
  if (/^total\b/i.test(c0))                           return false;
  if (/^nivel de gobierno/i.test(c0))                 return false;
  if (/^gob\.loc/i.test(c0))                          return false;
  if (/^departamento/i.test(c0))                      return false;
  if (/^municipalidad\b/i.test(c0) &&
      !/^\d{6}-/.test(c0))                            return false;
  if (/^fuente de financiamiento/i.test(c0))          return false;
  if (/^rubro$/i.test(c0))                            return false;
  if (/^genérica$/i.test(c0))                         return false;
  if (/^consulta amigable/i.test(c0))                 return false;
  const pim = toNum(row[2]);
  const rec = toNum(row[3]);
  return pim > 0 || rec > 0;
}

function detectarTipo(registros) {
  if (!registros.length) return "desconocido";
  const c0 = String(registros[0].descripcion || "");
  if (/^\d{6}-\d+:/.test(c0)) return "ranking";
  if (/^0\d:/.test(c0))       return "rubro";
  const todas = registros.map(r => String(r.descripcion || "").toUpperCase()).join(" ");
  if (/RECURSOS DIRECTAMENTE|RECURSOS DETERMINADOS|OPERACIONES OFICIALES/.test(todas))
    return "fuente";
  return "generica";
}

function parsearIngresos(arrayBuffer, nombreArchivo) {
  try {
    const u8 = new Uint8Array(arrayBuffer);
    const wb = XLSX.read(u8, { type: "array", codepage: 1252 });

    const registros = [];
    let totalPIA = 0, totalPIM = 0, totalRec = 0;
    let totalEncontrado = false;

    for (const sheetName of wb.SheetNames) {
      const aoa = XLSX.utils.sheet_to_json(
        wb.Sheets[sheetName],
        { header: 1, blankrows: false, raw: false }
      );

      for (const row of aoa) {
        const c0 = String(row[0] || "").trim();

        if (/Municipalidad\s+140101/i.test(c0) && !totalEncontrado) {
          totalPIA = toNum(row[1]);
          totalPIM = toNum(row[2]);
          totalRec = toNum(row[3]);
          totalEncontrado = true;
          continue;
        }

        if (!esFilaDato(row)) continue;

        registros.push({
          descripcion: c0,
          pia: toNum(row[1]),
          pim: toNum(row[2]),
          rec: toNum(row[3])
        });
      }
      if (registros.length) break;
    }

    if (!registros.length) return null;

    const tipo = detectarTipo(registros);

    if (!totalEncontrado) {
      totalPIA = registros.reduce((s, r) => s + r.pia, 0);
      totalPIM = registros.reduce((s, r) => s + r.pim, 0);
      totalRec = registros.reduce((s, r) => s + r.rec, 0);
    }

    return { tipo, totalPIA, totalPIM, totalRec, registros };
  } catch (e) {
    console.warn("[MPC] parsearIngresos error:", nombreArchivo, e.message);
    return null;
  }
}

const ARCHIVOS_ESPERADOS = ["fuente.xls", "generica.xls", "rubro.xls", "ranking.xls"];
let datos = {};
let cargados = new Set();

async function autoCargar() {
  for (const nombre of ARCHIVOS_ESPERADOS) {
    try {
      const r = await fetch("xls/" + nombre + "?" + Date.now());
      if (!r.ok) continue;
      const buf = await r.arrayBuffer();
      const res = parsearIngresos(buf, nombre);
      if (res) {
        datos[res.tipo] = res;
        cargados.add(nombre);
        actualizarFileList();
        render();
        console.log("[MPC] Auto-cargado: xls/" + nombre, "→ tipo:", res.tipo,
                    "| registros:", res.registros.length,
                    "| PIM:", res.totalPIM, "| Rec:", res.totalRec);
      }
    } catch (e) { /* archivo no disponible */ }
  }
}

function procesarArchivos(files) {
  Array.from(files).forEach(file => {
    const reader = new FileReader();
    reader.onload = e => {
      const res = parsearIngresos(e.target.result, file.name);
      if (res) {
        datos[res.tipo] = res;
        cargados.add(file.name);
        actualizarFileList();
        render();
        console.log("[MPC] Manual:", file.name, "→ tipo:", res.tipo,
                    "| registros:", res.registros.length);
      } else {
        console.warn("[MPC] No se pudo parsear:", file.name);
      }
    };
    reader.readAsArrayBuffer(file);
  });
}

function actualizarFileList() {
  const fl = $("fileList");
  if (!fl) return;
  fl.innerHTML = ARCHIVOS_ESPERADOS.map(f => {
    const ok = cargados.has(f);
    return `<div class="file-item">
      <span class="${ok ? "fi-ok" : "fi-wait"}">${ok ? "✓" : "○"}</span>
      <span class="fi-name">${f}</span>
    </div>`;
  }).join("");
}

function render() {
  const hoy = fechaHoy();
  ["b1fecha","b2fecha","b3fecha","b4fecha"].forEach(id => {
    const el = $(id); if (el) el.textContent = hoy;
  });
  renderB1();
  renderB2();
  renderB3();
  renderB4();
  renderB5();
  renderB7();
  renderB8();
}

function renderB1() {
  const d = datos.rubro || datos.fuente;
  if (!d) return;

  const pia = d.totalPIA, pim = d.totalPIM, rec = d.totalRec;
  const pct    = pim > 0 ? rec / pim * 100 : null;
  const pend   = pim - rec;
  const difPIA = pim - pia;
  const pctPIA = pia > 0 ? difPIA / pia * 100 : 0;

  const hoy   = new Date();
  const meses = ["Ene","Feb","Mar","Abr","May","Jun","Jul","Ago","Sep","Oct","Nov","Dic"];

  $("kpi-pia").textContent = fmtS(pia);
  $("kpi-pim").textContent = fmtS(pim);
  $("kpi-pim-sub").textContent =
    (pctPIA >= 0 ? "+" : "") + pctPIA.toFixed(1) +
    "% sobre PIA · Modificaciones al " + hoy.getDate() + "-" + meses[hoy.getMonth()];
  $("kpi-rec").textContent = fmtS(rec);
  $("kpi-pct").textContent = pct !== null ? pct.toFixed(1) + "%" : "N/A";
  $("kpi-pct-sub").textContent = pct !== null
    ? "Pendiente de recaudar: " + fmtS(Math.max(0, pend))
    : "PIM = 0";
}

// Instancias de donas B2
let b2DonaInstancias = [];

function renderB2() {
  const d     = datos.fuente;
  const tbody = $("b2tbody"), tfoot = $("b2tfoot"), nota = $("b2nota");

  if (!d || !d.registros.length) {
    tbody.innerHTML = '<tr><td colspan="6" class="vacio">Carga fuente.xls para ver los datos.</td></tr>';
    const w = $("b2donawrap"); if (w) w.style.display = "none";
    return;
  }

  // Paleta B2 — fuentes de financiamiento
  const PALETA_B2 = [
    { bg: "#0067a6", light: "#e0f0f9" },  // azul — RDR (2)
    { bg: "#7c3aed", light: "#ede9fe" },  // violeta — Operaciones Crédito (3)
    { bg: "#059669", light: "#d1fae5" },  // verde — Donaciones (4)
    { bg: "#9a1820", light: "#fce8e9" },  // rojo institucional — Recursos Determinados (5)
    { bg: "#d9a000", light: "#fef3c7" },  // dorado — TOTAL
  ];
  renderDonas(d.registros, d.totalPIM, d.totalRec, "b2dona", b2DonaInstancias, PALETA_B2);

  // ── Tabla ──────────────────────────────────────────────────────
  const notas = [];
  tbody.innerHTML = d.registros.map(r => {
    const pend  = r.pim - r.rec;
    const pct   = r.pim > 0 ? r.rec / r.pim * 100 : (r.rec > 0 ? null : 0);
    const nombre  = r.descripcion.replace(/^\d+:\s*/, "");
    const cod     = (r.descripcion.match(/^(\d+)/) || ["",""])[1];
    // Pendiente: nunca negativo — si rec >= pim mostrar S/ 0
    const pendMostrar = Math.max(0, pend);
    // % recaudación: 100% si PIM=0 y rec>0; normal si PIM>0; 0 si ambos 0
    const pctTabla = r.pim === 0 ? (r.rec > 0 ? 100 : 0) : Math.min(r.rec / r.pim * 100, 100);
    return `<tr>
      <td style="font-weight:600">${cod}: ${nombre}</td>
      <td class="num">${fmtS(r.pia)}</td>
      <td class="num">${fmtS(r.pim)}</td>
      <td class="num">${fmtS(r.rec)}</td>
      <td class="num">${fmtS(pendMostrar)}</td>
      <td>${barraHTML(pctTabla)}</td>
    </tr>`;
  }).join("");

  const tp = d.totalPIM, tr_ = d.totalRec, tpia = d.totalPIA;
  const tpct = tp > 0 ? Math.min(tr_ / tp * 100, 100) : (tr_ > 0 ? 100 : 0);
  const tpend = Math.max(0, tp - tr_);
  tfoot.innerHTML = `<tr>
    <td style="font-family:'Barlow Condensed';font-weight:800;text-transform:uppercase">TOTAL</td>
    <td class="num">${fmtS(tpia)}</td>
    <td class="num">${fmtS(tp)}</td>
    <td class="num">${fmtS(tr_)}</td>
    <td class="num">${fmtS(tpend)}</td>
    <td>${barraHTML(tpct)}</td>
  </tr>`;

  nota.style.display = "none";
}

// Instancias de donas B3 y B4
let b3DonaInstancias = [];
let b4DonaInstancias = [];

// Helper genérico: crea el HTML + Chart.js para un conjunto de donas
// prefix: "b3dona" | "b4dona"
// instArr: array donde se guardan las instancias para destruirlas luego
// PALETA_FN: función(index) → {bg, light}
function renderDonas(registros, totalPIM, totalRec, prefix, instArr, paleta) {
  const contenedor = $(prefix + "wrap");
  if (!contenedor) return;

  instArr.forEach(c => c.destroy());
  instArr.length = 0;

  // pct: si PIM=0 y rec>0 → 100%; si PIM>0 → min(rec/pim,100); si ambos 0 → 0
  function calcPct(pim, rec) {
    if (pim === 0) return rec > 0 ? 100 : 0;
    return Math.min(rec / pim * 100, 100);
  }

  const items = registros.map((r, i) => {
    const cod         = (r.descripcion.match(/^(\d+)/) || ["",""])[1];
    const nombre      = r.descripcion.replace(/^\d+:\s*/, "").trim();
    const nombreCorto = nombre.length > 22 ? nombre.substring(0, 20).trimEnd() + "…" : nombre;
    const pct         = calcPct(r.pim, r.rec);
    return { cod, nombreCorto, pct, rec: r.rec, pim: r.pim,
             color: paleta[i % (paleta.length - 1)] };
  });

  const pctTotal = calcPct(totalPIM, totalRec);
  items.push({ cod: "∑", nombreCorto: "TOTAL", pct: pctTotal,
               rec: totalRec, pim: totalPIM,
               color: paleta[paleta.length - 1], esTotal: true });

  // Layout: items normales flex-start con gap, TOTAL separada a la derecha con borde izquierdo
  contenedor.style.display = "flex";
  contenedor.style.justifyContent = "flex-start";
  contenedor.style.gap = "55px";
  contenedor.style.alignItems = "flex-start";
  contenedor.innerHTML = items.map((it, idx) => {
    const esTotal    = !!it.esTotal;
    const marginLeft = esTotal ? "auto" : "0";
    const borderLeft = esTotal ? "1px solid #e0d8d0" : "none";
    const paddingLeft = esTotal ? "24px" : "0";
    return `<div style="min-width:120px;max-width:155px;text-align:center;
                        margin-left:${marginLeft};border-left:${borderLeft};padding-left:${paddingLeft}">
      <div style="font-family:'Barlow Condensed';font-size:11px;font-weight:700;
                  color:#6b7280;text-transform:uppercase;letter-spacing:.04em;
                  margin-bottom:6px;min-height:28px;display:flex;align-items:center;
                  justify-content:center;line-height:1.3">${it.nombreCorto}</div>
      <div style="position:relative;width:110px;height:110px;margin:0 auto">
        <canvas id="${prefix}${idx}" width="110" height="110"></canvas>
        <div style="position:absolute;inset:0;display:flex;flex-direction:column;
                    align-items:center;justify-content:center;pointer-events:none">
          <span style="font-family:'Barlow Condensed';font-weight:800;
                       font-size:${esTotal ? "20px" : "18px"};
                       color:${it.color.bg};line-height:1">
            ${it.pct.toFixed(1)}%
          </span>
          <span style="font-size:9px;font-weight:600;color:#9ca3af;letter-spacing:.03em;
                       text-transform:uppercase;margin-top:2px">REC.</span>
        </div>
      </div>
      <div style="margin-top:7px;font-family:'Barlow Condensed';font-size:11px;
                  color:#6b7280;font-weight:600">
        PIM: ${it.pim >= 1e6
          ? "S/ " + (it.pim / 1e6).toLocaleString("es-PE",{minimumFractionDigits:1,maximumFractionDigits:1}) + " M"
          : "S/ " + Math.round(it.pim).toLocaleString("es-PE")}
      </div>
    </div>`;
  }).join("");

  items.forEach((it, idx) => {
    const canvas = $(prefix + idx);
    if (!canvas) return;
    const pctVal = it.pct;
    const inst = new Chart(canvas, {
      type: "doughnut",
      data: { datasets: [{
        data: [pctVal, Math.max(0, 100 - pctVal)],
        backgroundColor: [it.color.bg, it.color.light],
        borderWidth: 0,
        borderRadius: pctVal > 0 && pctVal < 100 ? 4 : 0,
      }]},
      options: {
        responsive: false, cutout: "72%",
        plugins: { legend: { display: false }, tooltip: { enabled: false } },
        animation: { duration: 500 }, events: [],
      }
    });
    instArr.push(inst);
  });
}

function renderB3() {
  const d     = datos.rubro;
  const tbody = $("b3tbody"), tfoot = $("b3tfoot"), nota = $("b3nota");
  const wrap  = $("b3donawrap");

  if (!d || !d.registros.length) {
    tbody.innerHTML = '<tr><td colspan="7" class="vacio">Carga rubro.xls para ver los datos.</td></tr>';
    if (wrap) wrap.style.display = "none";
    return;
  }

  // Paleta B3 — tonos verdes/azules/cálidos por rubro
  const PALETA_B3 = [
    { bg: "#0067a6", light: "#e0f0f9" },  // azul — FONCOMUN (07)
    { bg: "#9a1820", light: "#fce8e9" },  // rojo institucional — Impuestos (08)
    { bg: "#059669", light: "#d1fae5" },  // verde — RDR (09)
    { bg: "#7c3aed", light: "#ede9fe" },  // violeta — Donaciones (13)
    { bg: "#b45309", light: "#fef3c7" },  // ámbar — Canon (18)
    { bg: "#6b7280", light: "#f3f4f6" },  // gris — Operaciones crédito (19)
    { bg: "#d9a000", light: "#fef3c7" },  // dorado — TOTAL
  ];
  renderDonas(d.registros, d.totalPIM, d.totalRec, "b3dona", b3DonaInstancias, PALETA_B3);

  // ── Tabla ──────────────────────────────────────────────────────
  tbody.innerHTML = d.registros.map(r => {
    const pend       = r.pim - r.rec;
    const pendMostrar = Math.max(0, pend);
    // % recaudación: 100% si PIM=0 y rec>0, o si rec>=pim; normal en resto
    const pctTabla   = r.pim === 0 ? (r.rec > 0 ? 100 : 0) : Math.min(r.rec / r.pim * 100, 100);
    const cod    = (r.descripcion.match(/^(\d+)/) || ["",""])[1];
    const nombre = r.descripcion.replace(/^\d+:\s*/, "").trim();
    return `<tr>
      <td class="cod">${cod}</td>
      <td style="font-weight:600">${nombre}</td>
      <td class="num">${fmtS(r.pia)}</td>
      <td class="num">${fmtS(r.pim)}</td>
      <td class="num">${fmtS(r.rec)}</td>
      <td class="num">${fmtS(pendMostrar)}</td>
      <td>${barraHTML(pctTabla)}</td>
    </tr>`;
  }).join("");

  const tp = d.totalPIM, tr_ = d.totalRec, tpia = d.totalPIA;
  const tpct  = tp > 0 ? Math.min(tr_ / tp * 100, 100) : (tr_ > 0 ? 100 : 0);
  const tpend = Math.max(0, tp - tr_);
  tfoot.innerHTML = `<tr>
    <td class="cod">—</td>
    <td style="font-family:'Barlow Condensed';font-weight:800;text-transform:uppercase">TOTAL</td>
    <td class="num">${fmtS(tpia)}</td>
    <td class="num">${fmtS(tp)}</td>
    <td class="num">${fmtS(tr_)}</td>
    <td class="num">${fmtS(tpend)}</td>
    <td>${barraHTML(tpct)}</td>
  </tr>`;

  nota.style.display = "none";
}

function renderB4() {
  const d     = datos.generica;
  const tbody = $("b4tbody"), tfoot = $("b4tfoot"), nota = $("b4nota");
  const wrap  = $("b4donawrap");

  if (!d || !d.registros.length) {
    tbody.innerHTML = '<tr><td colspan="7" class="vacio">Carga generica.xls para ver los datos.</td></tr>';
    if (wrap) wrap.style.display = "none";
    return;
  }

  // Paleta B4 — genéricas (1, 3, 4, 5, 9)
  const PALETA_B4 = [
    { bg: "#9a1820", light: "#fce8e9" },  // rojo — Impuestos y contribuciones (1)
    { bg: "#0067a6", light: "#e0f0f9" },  // azul — Venta bienes/servicios (3)
    { bg: "#059669", light: "#d1fae5" },  // verde — Donaciones (4)
    { bg: "#b45309", light: "#fef3c7" },  // ámbar — Otros ingresos (5)
    { bg: "#7c3aed", light: "#ede9fe" },  // violeta — Saldos de balance (9)
    { bg: "#d9a000", light: "#fef3c7" },  // dorado — TOTAL
  ];
  renderDonas(d.registros, d.totalPIM, d.totalRec, "b4dona", b4DonaInstancias, PALETA_B4);

  // ── Tabla ──────────────────────────────────────────────────────
  tbody.innerHTML = d.registros.map(r => {
    const pend        = r.pim - r.rec;
    const pendMostrar = Math.max(0, pend);
    const pctTabla    = r.pim === 0 ? (r.rec > 0 ? 100 : 0) : Math.min(r.rec / r.pim * 100, 100);
    const cod    = (r.descripcion.match(/^(\d+)/) || ["",""])[1];
    const nombre = r.descripcion.replace(/^\d+:\s*/, "").trim();
    return `<tr>
      <td class="cod">${cod}</td>
      <td style="font-weight:600">${nombre}</td>
      <td class="num">${fmtS(r.pia)}</td>
      <td class="num">${fmtS(r.pim)}</td>
      <td class="num">${fmtS(r.rec)}</td>
      <td class="num">${fmtS(pendMostrar)}</td>
      <td>${barraHTML(pctTabla)}</td>
    </tr>`;
  }).join("");

  const tp = d.totalPIM, tr_ = d.totalRec, tpia = d.totalPIA;
  const tpct  = tp > 0 ? Math.min(tr_ / tp * 100, 100) : (tr_ > 0 ? 100 : 0);
  const tpend = Math.max(0, tp - tr_);
  tfoot.innerHTML = `<tr>
    <td class="cod">—</td>
    <td style="font-family:'Barlow Condensed';font-weight:800;text-transform:uppercase">TOTAL</td>
    <td class="num">${fmtS(tpia)}</td>
    <td class="num">${fmtS(tp)}</td>
    <td class="num">${fmtS(tr_)}</td>
    <td class="num">${fmtS(tpend)}</td>
    <td>${barraHTML(tpct)}</td>
  </tr>`;

  nota.style.display = "none";
}

// ═══════════════════════════════════════════════════════════════
//  BLOQUE 5 — RANKING DE MUNICIPALIDADES (comparativo por año 2015–2026)
//  Años cerrados: data/historico_ranking.json (generado por scripts/generar_historico_ranking.py)
//  Año en curso (2026): xls/ranking.xls cargado en vivo (datos.ranking)
// ═══════════════════════════════════════════════════════════════
const MPC_ID      = "140101-301212";   // ubigeo-ejecutora de la Municipalidad Provincial de Chiclayo
const ANIO_ACTUAL = 2026;
let B5_HIST = null, b5Anio = ANIO_ACTUAL, b5Orden = "monto";
const b5Muni = MPC_ID;   // el gráfico lineal es solo de la Municipalidad Provincial de Chiclayo
let b5ChartInstance = null, b5Listo = false;

fetch("data/historico_ranking.json?" + Date.now())
  .then(r => { if (!r.ok) throw new Error("HTTP " + r.status); return r.json(); })
  .then(j => { B5_HIST = j; b5Init(); renderB5(); })
  .catch(e => console.warn("[MPC] No se pudo cargar historico_ranking.json:", e.message));

function b5Init() {
  if (b5Listo || !B5_HIST) return;
  b5Listo = true;

  $("b5tabs").innerHTML = B5_HIST.años.map(a =>
    `<button type="button" class="tab" data-anio="${a}" aria-pressed="${a === b5Anio}">${a}` +
    (a === ANIO_ACTUAL ? '<span class="en-curso">en curso</span>' : "") + "</button>"
  ).join("");
  $("b5tabs").addEventListener("click", e => {
    const b = e.target.closest(".tab"); if (!b) return;
    b5Anio = +b.getAttribute("data-anio");
    document.querySelectorAll("#b5tabs .tab").forEach(t =>
      t.setAttribute("aria-pressed", String(+t.getAttribute("data-anio") === b5Anio)));
    renderB5();
  });

  const setOrden = o => {
    b5Orden = o;
    $("b5ordMonto").setAttribute("aria-pressed", String(o === "monto"));
    $("b5ordPct").setAttribute("aria-pressed", String(o === "pct"));
    renderB5();
  };
  $("b5ordMonto").addEventListener("click", () => setOrden("monto"));
  $("b5ordPct").addEventListener("click", () => setOrden("pct"));
}

// Filas del ranking para un año. Los nombres salen del JSON (el .xls pierde la Ñ al leerlo con SheetJS).
function b5Filas(anio) {
  const nombres = B5_HIST ? B5_HIST.municipalidades : {};
  let base = null;
  if (anio === ANIO_ACTUAL && datos.ranking && datos.ranking.registros.length) {
    base = datos.ranking.registros.map(r => {
      const m  = r.descripcion.match(/^(\d{6}-\d{6})/);
      const id = m ? m[1] : r.descripcion;
      return { id, nombre: (nombres[id] || r.descripcion.replace(/^\d{6}-\d+:\s*/, "")).toUpperCase(),
               pia: r.pia, pim: r.pim, rec: r.rec };
    });
  } else if (B5_HIST && B5_HIST.datos[String(anio)]) {
    const d = B5_HIST.datos[String(anio)].munis;
    base = Object.keys(d).map(id => ({ id, nombre: (nombres[id] || id).toUpperCase(),
                                       pia: d[id].pia, pim: d[id].pim, rec: d[id].eneoct }));
  }
  if (!base) return null;
  return base.map(f => ({ ...f, pct: f.pim > 0 ? f.rec / f.pim * 100 : null, esMPC: f.id === MPC_ID }));
}

function renderB5() {
  const tbody = $("b5tbody");
  const filas = b5Filas(b5Anio);
  if (!filas) {
    tbody.innerHTML = '<tr><td colspan="6" class="vacio">Carga ranking.xls para ver los datos.</td></tr>';
    $("b5tfoot").innerHTML = "";
    return;
  }
  const enCurso  = b5Anio === ANIO_ACTUAL;
  const porMonto = [...filas].sort((a, b) => b.rec - a.rec);
  const porPct   = [...filas].sort((a, b) => (b.pct ?? -1) - (a.pct ?? -1));
  const lista    = b5Orden === "monto" ? porMonto : porPct;

  $("b5thRec").textContent = `Recaudado ${CORTE_LABEL} (S/)`;
  $("b5meta").textContent  = `${filas.length} municipalidades · ordenado por ${b5Orden === "monto" ? "monto recaudado" : "% de avance"} · ${CORTE_LABEL} ${b5Anio}${enCurso ? " (a la fecha)" : ""}`;

  tbody.innerHTML = lista.map((m, i) => {
    const pos      = i + 1;
    const posStyle = (pos <= 3 && !m.esMPC) ? 'style="color:var(--dorado-osc);font-weight:800"' : 'style="font-weight:700"';
    return `<tr ${m.esMPC ? 'class="mpl-row"' : ""}>
      <td style="text-align:center" ${posStyle}>${pos}°</td>
      <td style="font-weight:700">${m.nombre}</td>
      <td class="num">${fmtS(m.pia)}</td>
      <td class="num">${fmtS(m.pim)}</td>
      <td class="num">${fmtS(m.rec)}</td>
      <td>${barraHTML(m.pct, m.esMPC ? "var(--rojo-osc)" : null)}</td>
    </tr>`;
  }).join("");

  const tPia = filas.reduce((s, f) => s + f.pia, 0), tPim = filas.reduce((s, f) => s + f.pim, 0),
        tRec = filas.reduce((s, f) => s + f.rec, 0);
  $("b5tfoot").innerHTML = `<tr><td></td><td>TOTAL DPTO. LAMBAYEQUE (${filas.length} MUNICIPALIDADES)</td>
    <td class="num">${fmtS(tPia)}</td><td class="num">${fmtS(tPim)}</td><td class="num">${fmtS(tRec)}</td>
    <td>${barraHTML(tPim > 0 ? tRec / tPim * 100 : null)}</td></tr>`;

  renderB5Linea();
}

function renderB5Linea() {
  const canvas = $("b5chart");
  if (!canvas || !B5_HIST) return;
  const años = B5_HIST.años;
  const valores = años.map(a => {
    if (a === ANIO_ACTUAL) {
      const f = (b5Filas(a) || []).find(x => x.id === b5Muni);
      return f ? f.rec : null;
    }
    const m = B5_HIST.datos[String(a)].munis[b5Muni];
    return m ? m.eneoct : null;
  });
  const idxAct = años.indexOf(ANIO_ACTUAL);
  const nombre = (B5_HIST.municipalidades[b5Muni] || "").toUpperCase();
  $("b5lineaMeta").textContent = `${CORTE_LABEL} ${años[0]}–${años[años.length - 1]} · ${nombre}`;

  if (b5ChartInstance) { b5ChartInstance.destroy(); b5ChartInstance = null; }
  b5ChartInstance = new Chart(canvas, {
    type: "line",
    data: {
      labels: años.map(a => a === ANIO_ACTUAL ? `${a} (a la fecha)` : `${CORTE_LABEL} ${a}`),
      datasets: [{
        label: "Recaudado " + CORTE_LABEL, data: valores,
        borderColor: "#1f5f92", borderWidth: 2.5, backgroundColor: "rgba(43,128,193,.08)",
        fill: true, tension: 0.3, spanGaps: true,
        pointBackgroundColor: años.map(a => a === ANIO_ACTUAL ? "#FFC526" : "#2B80C1"),
        pointBorderColor: "#1f5f92", pointBorderWidth: 2,
        pointRadius: años.map(a => a === ANIO_ACTUAL ? 8 : 5), pointHoverRadius: 9
      }]
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      layout: { padding: { top: 28, right: 24, left: 4 } },
      plugins: {
        legend: { display: false },
        tooltip: { callbacks: { label: c => c.raw == null ? " Sin datos" : " " + fmtS(c.raw) } }
      },
      scales: {
        x: { grid: { display: false },
             ticks: { font: { family: "'Barlow Condensed'", weight: "700", size: 12 },
                      color: c => c.index === idxAct ? "#92400e" : "#374151" } },
        y: { grace: "12%",
             ticks: { callback: v => "S/ " + Math.round(v).toLocaleString("es-PE"),
                      font: { family: "'Barlow'", size: 11 }, color: "#6b7280" },
             grid: { color: "#f3f4f6" } }
      },
      animation: { duration: 500 }
    },
    plugins: [{
      id: "b5Labels",
      afterDatasetsDraw(chart) {
        const { ctx } = chart, meta = chart.getDatasetMeta(0);
        ctx.save();
        chart.data.datasets[0].data.forEach((v, i) => {
          if (v == null) return;
          const pt = meta.data[i];
          ctx.font = "700 12px 'Barlow Condensed'";
          ctx.textAlign = "center"; ctx.textBaseline = "bottom";
          ctx.lineJoin = "round"; ctx.lineWidth = 4; ctx.strokeStyle = "#ffffff";
          ctx.strokeText(fmtS(v), pt.x, pt.y - 12);
          ctx.fillStyle = i === idxAct ? "#92400e" : "#1f5f92";
          ctx.fillText(fmtS(v), pt.x, pt.y - 12);
        });
        ctx.restore();
      }
    }]
  });
}

$("pick").addEventListener("click", () => $("file").click());
$("dropzone").addEventListener("click", () => $("file").click());
$("dropzone").addEventListener("keydown", e => {
  if (e.key === "Enter" || e.key === " ") { e.preventDefault(); $("file").click(); }
});
$("file").addEventListener("change", e => {
  procesarArchivos(e.target.files); e.target.value = "";
});
$("clear").addEventListener("click", () => {
  datos = {}; cargados = new Set(); actualizarFileList(); render();
});

const dz = $("dropzone");
["dragenter","dragover"].forEach(ev =>
  dz.addEventListener(ev, e => { e.preventDefault(); dz.classList.add("drag"); }));
["dragleave","drop"].forEach(ev =>
  dz.addEventListener(ev, e => { e.preventDefault(); dz.classList.remove("drag"); }));
dz.addEventListener("drop", e => {
  if (e.dataTransfer.files.length) procesarArchivos(e.dataTransfer.files);
});

let pdfLibsCargadas = false;

function exportarPDF() {
  if (!pdfLibsCargadas) {
    const s1 = document.createElement("script");
    s1.src = "https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js";
    s1.onload = () => {
      const s2 = document.createElement("script");
      s2.src = "https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js";
      s2.onload = () => { pdfLibsCargadas = true; generarPDF(); };
      document.head.appendChild(s2);
    };
    document.head.appendChild(s1);
  } else { generarPDF(); }
}

function generarPDF() {
  const { jsPDF } = window.jspdf;
  const btn = document.querySelector(".btn-pdf");
  const btnTxtOriginal = btn ? btn.innerHTML : null;

  const bloques = Array.from(document.querySelectorAll(".bloque"));
  if (!bloques.length) return;

  const pdf     = new jsPDF({ orientation: "landscape", unit: "mm", format: "a4" });
  const pw      = pdf.internal.pageSize.getWidth();
  const ph      = pdf.internal.pageSize.getHeight();
  const margen  = 8; // mm de margen alrededor de cada bloque

  if (btn) btn.innerHTML = "Generando PDF…";

  // Procesa los bloques uno por uno (en secuencia) para no saturar memoria
  bloques.reduce((promesa, bloque, idx) => {
    return promesa.then(() =>
      html2canvas(bloque, { scale: 2, useCORS: true, logging: false, backgroundColor: "#ffffff" })
        .then(canvas => {
          if (idx > 0) pdf.addPage("a4", "landscape");

          // Ajusta el bloque a la hoja horizontal manteniendo proporción
          const areaW = pw - margen * 2;
          const areaH = ph - margen * 2;
          let imgW = areaW;
          let imgH = canvas.height * imgW / canvas.width;
          if (imgH > areaH) {
            imgH = areaH;
            imgW = canvas.width * imgH / canvas.height;
          }
          const x = (pw - imgW) / 2;
          const y = (ph - imgH) / 2;

          pdf.addImage(canvas.toDataURL("image/jpeg", 0.92), "JPEG", x, y, imgW, imgH);
        })
    );
  }, Promise.resolve()).then(() => {
    const hoy = new Date();
    pdf.save(`MPC_Ingresos_${String(hoy.getDate()).padStart(2,"0")}${String(hoy.getMonth()+1).padStart(2,"0")}${hoy.getFullYear()}.pdf`);
    if (btn) btn.innerHTML = btnTxtOriginal;
  }).catch(err => {
    console.error("[MPC] Error generando PDF:", err);
    if (btn) btn.innerHTML = btnTxtOriginal;
  });
}

$("fechaHeader").textContent = fechaCorta();
["b1fecha","b2fecha","b3fecha","b4fecha"].forEach(id => {
  const el = $(id); if (el) el.textContent = fechaHoy();
});
autoCargar();

// ═══════════════════════════════════════════════════════════════
//  BLOQUE 7 — COMPARATIVO HISTÓRICO RUBRO 08: IMPUESTOS MUNICIPALES
// ═══════════════════════════════════════════════════════════════

let B7_HIST = {};
let b7ChartInstance = null;

fetch("data/historico_rubro08.json?" + Date.now())
  .then(r => r.json())
  .then(data => {
    B7_HIST = data;
    renderB7();
  })
  .catch(() => console.warn("[MPC] No se pudo cargar historico_rubro08.json"));

function getRubro08_2026() {
  if (!datos.rubro || !datos.rubro.registros) return 0;
  const r08 = datos.rubro.registros.find(r => {
    const cod = (r.descripcion.match(/^(\d+)/) || ["",""])[1];
    return cod === "08";
  });
  return r08 ? (r08.rec || 0) : 0;
}

function renderB7() {
  // Ver nota en renderB6: 2021-2025 = corte CORTE_LABEL congelado en JSON; el año en curso se
  // calcula siempre en vivo (nunca se guarda) mientras hoy esté dentro de la ventana.
  const hoyB7 = new Date();
  const añoActualB7 = hoyB7.getFullYear();
  const hayAñoVivoB7 = hoyB7.getMonth() <= CORTE_MES_MAX && !(String(añoActualB7) in B7_HIST);

  const añosHist   = Object.keys(B7_HIST).map(Number).sort();
  const años       = hayAñoVivoB7 ? [...añosHist, añoActualB7] : añosHist;
  const valores    = hayAñoVivoB7
    ? [...añosHist.map(a => B7_HIST[String(a)].total), getRubro08_2026()]
    : añosHist.map(a => B7_HIST[String(a)].total);
  const IDX_2026_B7 = hayAñoVivoB7 ? años.length - 1 : -1;

  const fmtM = n => {
    if (!n) return "S/ —";
    return "S/ " + Math.round(n).toLocaleString("es-PE");
  };

  // KPI cards
  const kpiContainer = $("b7kpis");
  if (kpiContainer) {
    kpiContainer.innerHTML = años.map((a, i) => {
      const v      = valores[i];
      const esVivo = hayAñoVivoB7 && a === añoActualB7;
      const prev   = i > 0 ? valores[i - 1] : null;
      let deltaHtml = "";
      if (prev && prev > 0 && v > 0) {
        const pct   = (v - prev) / prev * 100;
        const color = pct >= 0 ? "#2a7d46" : "#c0392b";
        const signo = pct >= 0 ? "▲" : "▼";
        deltaHtml = `<span style="font-size:10px;color:${color};font-weight:700">${signo} ${Math.abs(pct).toFixed(1)}%</span>`;
      }
      return `<div style="background:${esVivo ? "#fef3c7" : "#f9fafb"};border:1px solid ${esVivo ? "#fbbf24" : "#e5e7eb"};border-radius:10px;padding:10px 16px;min-width:110px;flex:1;text-align:center">
        <div style="font-family:'Barlow Condensed';font-size:13px;font-weight:700;color:#6b7280;margin-bottom:3px">${esVivo ? a + " (a la fecha) \u2605" : CORTE_LABEL + " " + a}</div>
        <div style="font-family:'Barlow Condensed';font-size:18px;font-weight:800;color:${esVivo ? "#92400e" : "#1f2937"}">${v ? fmtM(v) : "Cargando\u2026"}</div>
        <div style="margin-top:3px">${deltaHtml}</div>
      </div>`;
    }).join("");
  }

  const canvas = $("b7chart");
  if (!canvas) return;
  if (b7ChartInstance) { b7ChartInstance.destroy(); b7ChartInstance = null; }

  const colores       = años.map(a => hayAñoVivoB7 && a === añoActualB7 ? "#FFC526" : "#2B80C1");
  const borderColores = años.map(a => hayAñoVivoB7 && a === añoActualB7 ? "#d9a000" : "#1f5f92");

  b7ChartInstance = new Chart(canvas, {
    type: "bar",
    data: {
      labels: años.map(a => hayAñoVivoB7 && a === añoActualB7 ? `${a} (a la fecha) \u2605` : `${CORTE_LABEL} ${a}`),
      datasets: [{
        label: "Recaudado " + CORTE_LABEL + " Rubro 08",
        data: valores,
        backgroundColor: colores,
        borderColor: borderColores,
        borderWidth: 2,
        borderRadius: 6,
        borderSkipped: false,
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      layout: { padding: { top: 28 } },
      plugins: {
        legend: { display: false },
        tooltip: {
          callbacks: {
            label: ctx => {
              const v = ctx.raw;
              if (!v) return " Sin datos";
              return ` S/ ${Math.round(v).toLocaleString("es-PE")}`;
            }
          }
        }
      },
      scales: {
        x: {
          grid: { display: false },
          ticks: {
            font: { family: "'Barlow Condensed'", weight: "700", size: 12 },
            color: ctx => ctx.index === IDX_2026_B7 ? "#92400e" : "#374151"
          }
        },
        y: {
          beginAtZero: false,
          suggestedMax: Math.max(...valores.filter(Boolean)) * 1.12,
          ticks: {
            callback: v => {
              if (v >= 1e6) return "S/ " + (v / 1e6).toFixed(1) + " M";
              if (v >= 1e3) return "S/ " + (v / 1e3).toFixed(0) + " K";
              return "S/ " + v;
            },
            font: { family: "'Barlow'", size: 11 },
            color: "#6b7280"
          },
          grid: { color: "#f3f4f6" }
        }
      },
      animation: { duration: 600 }
    },
    plugins: [{
      id: "b7Labels",
      afterDatasetsDraw(chart) {
        const { ctx, data } = chart;
        ctx.save();
        data.datasets[0].data.forEach((value, i) => {
          if (!value) return;
          const meta = chart.getDatasetMeta(0);
          const bar  = meta.data[i];
          const txt  = fmtM(value);
          ctx.fillStyle    = i === IDX_2026_B7 ? "#92400e" : "#1f5f92";
          ctx.font         = "700 12px 'Barlow Condensed'";
          ctx.textAlign    = "center";
          ctx.textBaseline = "bottom";
          ctx.fillText(txt, bar.x, bar.y - 8);
        });
        ctx.restore();
      }
    }]
  });
}

// ═══════════════════════════════════════════════════════════════
//  BLOQUE 8 — COMPARATIVO HISTÓRICO RUBRO 09: RECURSOS DIRECTAMENTE RECAUDADOS
// ═══════════════════════════════════════════════════════════════

let B8_HIST = {};
let b8ChartInstance = null;

fetch("data/historico_rubro09.json?" + Date.now())
  .then(r => r.json())
  .then(data => {
    B8_HIST = data;
    renderB8();
  })
  .catch(() => console.warn("[MPC] No se pudo cargar historico_rubro09.json"));

function getRubro09_2026() {
  if (!datos.rubro || !datos.rubro.registros) return 0;
  const r09 = datos.rubro.registros.find(r => {
    const cod = (r.descripcion.match(/^(\d+)/) || ["",""])[1];
    return cod === "09";
  });
  return r09 ? (r09.rec || 0) : 0;
}

function renderB8() {
  // Ver nota en renderB6: 2021-2025 = corte CORTE_LABEL congelado en JSON; el año en curso se
  // calcula siempre en vivo (nunca se guarda) mientras hoy esté dentro de la ventana.
  const hoyB8 = new Date();
  const añoActualB8 = hoyB8.getFullYear();
  const hayAñoVivoB8 = hoyB8.getMonth() <= CORTE_MES_MAX && !(String(añoActualB8) in B8_HIST);

  const añosHist    = Object.keys(B8_HIST).map(Number).sort();
  const años        = hayAñoVivoB8 ? [...añosHist, añoActualB8] : añosHist;
  const valores     = hayAñoVivoB8
    ? [...añosHist.map(a => B8_HIST[String(a)].total), getRubro09_2026()]
    : añosHist.map(a => B8_HIST[String(a)].total);
  const IDX_2026_B8 = hayAñoVivoB8 ? años.length - 1 : -1;

  const fmtM = n => {
    if (!n) return "S/ —";
    return "S/ " + Math.round(n).toLocaleString("es-PE");
  };

  // KPI cards
  const kpiContainer = $("b8kpis");
  if (kpiContainer) {
    kpiContainer.innerHTML = años.map((a, i) => {
      const v      = valores[i];
      const esVivo = hayAñoVivoB8 && a === añoActualB8;
      const prev   = i > 0 ? valores[i - 1] : null;
      let deltaHtml = "";
      if (prev && prev > 0 && v > 0) {
        const pct   = (v - prev) / prev * 100;
        const color = pct >= 0 ? "#2a7d46" : "#c0392b";
        const signo = pct >= 0 ? "▲" : "▼";
        deltaHtml = `<span style="font-size:10px;color:${color};font-weight:700">${signo} ${Math.abs(pct).toFixed(1)}%</span>`;
      }
      return `<div style="background:${esVivo ? "#fef3c7" : "#f9fafb"};border:1px solid ${esVivo ? "#fbbf24" : "#e5e7eb"};border-radius:10px;padding:10px 16px;min-width:110px;flex:1;text-align:center">
        <div style="font-family:'Barlow Condensed';font-size:13px;font-weight:700;color:#6b7280;margin-bottom:3px">${esVivo ? a + " (a la fecha) \u2605" : CORTE_LABEL + " " + a}</div>
        <div style="font-family:'Barlow Condensed';font-size:18px;font-weight:800;color:${esVivo ? "#92400e" : "#1f2937"}">${v ? fmtM(v) : "Cargando\u2026"}</div>
        <div style="margin-top:3px">${deltaHtml}</div>
      </div>`;
    }).join("");
  }

  const canvas = $("b8chart");
  if (!canvas) return;
  if (b8ChartInstance) { b8ChartInstance.destroy(); b8ChartInstance = null; }

  const colores       = años.map(a => hayAñoVivoB8 && a === añoActualB8 ? "#FFC526" : "#2B80C1");
  const borderColores = años.map(a => hayAñoVivoB8 && a === añoActualB8 ? "#d9a000" : "#1f5f92");

  b8ChartInstance = new Chart(canvas, {
    type: "bar",
    data: {
      labels: años.map(a => hayAñoVivoB8 && a === añoActualB8 ? `${a} (a la fecha) \u2605` : `${CORTE_LABEL} ${a}`),
      datasets: [{
        label: "Recaudado " + CORTE_LABEL + " Rubro 09",
        data: valores,
        backgroundColor: colores,
        borderColor: borderColores,
        borderWidth: 2,
        borderRadius: 6,
        borderSkipped: false,
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      layout: { padding: { top: 28 } },
      plugins: {
        legend: { display: false },
        tooltip: {
          callbacks: {
            label: ctx => {
              const v = ctx.raw;
              if (!v) return " Sin datos";
              return ` S/ ${Math.round(v).toLocaleString("es-PE")}`;
            }
          }
        }
      },
      scales: {
        x: {
          grid: { display: false },
          ticks: {
            font: { family: "'Barlow Condensed'", weight: "700", size: 12 },
            color: ctx => ctx.index === IDX_2026_B8 ? "#92400e" : "#374151"
          }
        },
        y: {
          beginAtZero: false,
          suggestedMax: Math.max(...valores.filter(Boolean)) * 1.12,
          ticks: {
            callback: v => {
              if (v >= 1e6) return "S/ " + (v / 1e6).toFixed(1) + " M";
              if (v >= 1e3) return "S/ " + (v / 1e3).toFixed(0) + " K";
              return "S/ " + v;
            },
            font: { family: "'Barlow'", size: 11 },
            color: "#6b7280"
          },
          grid: { color: "#f3f4f6" }
        }
      },
      animation: { duration: 600 }
    },
    plugins: [{
      id: "b8Labels",
      afterDatasetsDraw(chart) {
        const { ctx, data } = chart;
        ctx.save();
        data.datasets[0].data.forEach((value, i) => {
          if (!value) return;
          const meta = chart.getDatasetMeta(0);
          const bar  = meta.data[i];
          const txt  = fmtM(value);
          ctx.fillStyle    = i === IDX_2026_B8 ? "#92400e" : "#1f5f92";
          ctx.font         = "700 12px 'Barlow Condensed'";
          ctx.textAlign    = "center";
          ctx.textBaseline = "bottom";
          ctx.fillText(txt, bar.x, bar.y - 8);
        });
        ctx.restore();
      }
    }]
  });
}

