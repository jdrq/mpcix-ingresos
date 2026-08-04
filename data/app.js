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

function barraHTML(pct) {
  if (pct === null) {
    return '<span style="font-family:\'Barlow Condensed\';font-weight:700;color:#888;font-size:13px">N/A</span>';
  }
  const col = colorSem(pct);
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

function renderB5() {
  const d     = datos.ranking;
  const tbody = $("b5tbody"), hl = $("b5highlight");
  if (!d || !d.registros.length) {
    tbody.innerHTML = '<tr><td colspan="6" class="vacio">Carga ranking.xls para ver los datos.</td></tr>';
    hl.textContent = "Carga ranking.xls para ver la posición de la Municipalidad Provincial de Chiclayo.";
    return;
  }

  const muns = d.registros.map(r => {
    const pct       = r.pim > 0 ? r.rec / r.pim * 100 : null;
    const nombre    = r.descripcion.replace(/^\d{6}-\d+:\s*/, "").trim().toUpperCase();
    const esMPC = r.descripcion.includes("140101") ||
                  (/CHICLAYO/i.test(r.descripcion) && /PROVINCIAL/i.test(r.descripcion));
    return { ...r, pct, nombre, esMPC };
  });

  const porPct   = [...muns].sort((a, b) => (b.pct ?? -1) - (a.pct ?? -1));
  const porMonto = [...muns].sort((a, b) => (b.rec || 0) - (a.rec || 0));

  const posPct   = porPct.findIndex(m => m.esMPC) + 1;
  const posMonto = porMonto.findIndex(m => m.esMPC) + 1;
  const mpc      = muns.find(m => m.esMPC);

  if (mpc) {
    const pctStr = mpc.pct !== null ? mpc.pct.toFixed(1) + "%" : "N/A";
    hl.innerHTML =
      `&#128269; La <strong>Municipalidad Provincial de Chiclayo</strong> ocupa el puesto
       <strong>${posPct}° de ${muns.length}</strong> en % de recaudación sobre PIM (${pctStr}).
       Por <strong>monto recaudado absoluto</strong>, se ubica en el puesto
       <strong>${posMonto}° de ${muns.length}</strong> con ${fmtS(mpc.rec)}.`;
  }

  tbody.innerHTML = porPct.map((m, i) => {
    const pos      = i + 1;
    const trC      = m.esMPC ? 'class="mpl-row"' : "";
    const posStyle = pos <= 3 ? 'style="color:var(--dorado-osc);font-weight:800"' : 'style="font-weight:700"';
    return `<tr ${trC}>
      <td style="text-align:center" ${posStyle}>${pos}°</td>
      <td style="font-weight:700">${m.nombre}</td>
      <td class="num">${fmtS(m.pia)}</td>
      <td class="num">${fmtS(m.pim)}</td>
      <td class="num">${fmtS(m.rec)}</td>
      <td>${barraHTML(m.pct)}</td>
    </tr>`;
  }).join("");
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

