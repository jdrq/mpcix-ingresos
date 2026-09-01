// historico-conceptos.js — Municipalidad Provincial de Chiclayo (MPC)
// Renderizador genérico de bloques históricos por concepto.
// Mismo motor que mpl-ingresos/data/historico-conceptos.js, adaptado
// a la paleta institucional de Chiclayo (azul #2B80C1 / dorado #FFC526).
// Cada bloque muestra 2 gráficos: Comparación Anual y Comparación de un periodo
// parcial configurable (Ene-Ago, Ene-Sep, etc). La ventana se define por página
// vía window.HIST_CONFIG.ventanaLabel / .ventanaCorta, para que actualizar un
// rubro (ej. 08 a Ene-Sep) no descuadre la etiqueta de otro rubro que aún no
// se ha actualizado (ej. 09, todavía en Ene-Ago).
// Se usa en historico-rubro08.html e historico-rubro09.html.
// Config esperada en window.HIST_CONFIG:
//   { jsonPath: "data/historico_conceptos_rubro08.json",
//     ventanaLabel: "Enero–Septiembre",  // texto largo del subtítulo
//     ventanaCorta: "Ene–Sep" }          // prefijo corto en el gráfico/etiquetas

// Números completos, sin abreviar a K/M — a pedido del jefe (mismo criterio que MPL).
const fmtM = n => {
  if (n === null || n === undefined) return "S/ —";
  return "S/ " + Math.round(n).toLocaleString("es-PE");
};

// Defaults de compatibilidad: si una página no define ventanaLabel/ventanaCorta
// en window.HIST_CONFIG, se comporta igual que antes (Ene-Ago).
function ventanaLabel() { return (window.HIST_CONFIG && window.HIST_CONFIG.ventanaLabel) || "Enero–Agosto"; }
function ventanaCorta() { return (window.HIST_CONFIG && window.HIST_CONFIG.ventanaCorta) || "Ene–Ago"; }

function crearBloqueHTML(numero, key, concepto) {
  const num = String(numero).padStart(2, "0");
  return `
  <div class="bloque bloque-doble" id="bloque-${key}">
    <div class="bloque-header">
      <div class="bloque-num">${num}</div>
      <div class="bloque-titulo">${concepto.label.toUpperCase()}</div>
    </div>
    <div class="doble-grid">
      <div class="doble-col">
        <div class="bloque-subtitle">
          <span class="bloque-subtitle-text">Comparación de Ingresos Anual</span>
          <span class="bloque-subtitle-date">${window.HIST_RANGO || "2021–2026"}</span>
        </div>
        <div class="bloque-body">
          <div style="position:relative;height:300px;background:#fff;border-radius:10px;padding:8px 0 0 0">
            <canvas id="chart-anual-${key}"></canvas>
          </div>
        </div>
      </div>
      <div class="doble-col doble-col-right">
        <div class="bloque-subtitle">
          <span class="bloque-subtitle-text">Comparación de Ingresos (${ventanaLabel()})</span>
          <span class="bloque-subtitle-date">${window.HIST_RANGO || "2021–2026"}</span>
        </div>
        <div class="bloque-body">
          <div style="position:relative;height:300px;background:#fff;border-radius:10px;padding:8px 0 0 0">
            <canvas id="chart-eneago-${key}"></canvas>
          </div>
        </div>
      </div>
    </div>
    <div class="bloque-footer">
      <span>Fuente: Consulta Amigable MEF — Ingresos</span>
      <span class="foot-label">MPC · SERIE ${window.HIST_RANGO || "2021–2026"}</span>
    </div>
  </div>`;
}

function pintarChart(canvasId, años, valores, IDX_ACTUAL, esAnual) {
  const canvas = document.getElementById(canvasId);
  if (!canvas) return;

  const maxVal = Math.max(0, ...valores.filter(v => v !== null));
  // Headroom del 22% arriba de la barra más alta para que la etiqueta nunca se corte
  const yMax = maxVal > 0 ? maxVal * 1.22 : 10;

  // Paleta MPC: azul institucional (#2B80C1 / #1f5f92) en vez del
  // rojo de MPL. El año destacado sigue en dorado (#FFC526), color
  // compartido por ambas instituciones.
  const colores       = años.map((a, i) => i === IDX_ACTUAL ? "#FFC526" : "#2B80C1");
  const borderColores = años.map((a, i) => i === IDX_ACTUAL ? "#d9a000" : "#1f5f92");

  new Chart(canvas, {
    type: "bar",
    data: {
      labels: años.map((a, i) => (esAnual ? `${a}` : `${ventanaCorta()} ${a}`) + (i === IDX_ACTUAL ? " ★" : "")),
      datasets: [{
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
      layout: { padding: { top: 26 } },
      plugins: {
        legend: { display: false },
        tooltip: {
          callbacks: {
            label: ctx => ctx.raw === null ? " Sin dato" : ` S/ ${Math.round(ctx.raw).toLocaleString("es-PE")}`
          }
        }
      },
      scales: {
        x: {
          grid: { display: false },
          ticks: {
            font: { family: "'Barlow Condensed'", weight: "700", size: 12 },
            color: ctx => ctx.index === IDX_ACTUAL ? "#92400e" : "#374151"
          }
        },
        y: {
          beginAtZero: true,
          suggestedMax: yMax,
          ticks: {
            callback: v => "S/ " + Math.round(v).toLocaleString("es-PE"),
            font: { family: "'Barlow'", size: 10.5 }, color: "#6b7280"
          },
          grid: { color: "#f3f4f6" }
        }
      },
      animation: { duration: 500 }
    },
    plugins: [{
      id: "labelsPlugin",
      afterDatasetsDraw(chart) {
        const { ctx, data } = chart;
        ctx.save();
        data.datasets[0].data.forEach((value, i) => {
          if (value === null) return;
          const meta = chart.getDatasetMeta(0);
          const bar  = meta.data[i];
          ctx.fillStyle    = i === IDX_ACTUAL ? "#92400e" : "#1f5f92";
          ctx.font         = "700 12px 'Barlow Condensed'";
          ctx.textAlign    = "center";
          ctx.textBaseline = "bottom";
          ctx.fillText(fmtM(value), bar.x, bar.y - 8);
        });
        ctx.restore();
      }
    }]
  });
}

function pintarBloque(key, concepto) {
  const años = Object.keys(concepto.anual).filter(k => /^\d{4}$/.test(k)).sort();
  const IDX_ACTUAL = años.length - 1;

  const valoresAnual  = años.map(a => concepto.anual[a] ?? null);
  const valoresEneAgo = años.map(a => concepto.eneago[a] ?? null);

  pintarChart("chart-anual-" + key, años, valoresAnual, IDX_ACTUAL, true);
  pintarChart("chart-eneago-" + key, años, valoresEneAgo, IDX_ACTUAL, false);
}

async function initHistoricoConceptos() {
  const cfg = window.HIST_CONFIG;
  const cont = document.getElementById("bloquesContainer");
  try {
    const r = await fetch(cfg.jsonPath + "?" + Date.now());
    const data = await r.json();
    const keys = Object.keys(data);

    // Rango de años dinámico para el subtítulo
    const añosMuestra = Object.keys(data[keys[0]].anual).filter(k => /^\d{4}$/.test(k)).sort();
    window.HIST_RANGO = añosMuestra.length ? `${añosMuestra[0]}–${añosMuestra[añosMuestra.length - 1]}` : "";

    cont.innerHTML = keys.map((k, i) => crearBloqueHTML(i + 1, k, data[k])).join("");
    keys.forEach(k => pintarBloque(k, data[k]));
  } catch (e) {
    cont.innerHTML = `<div class="bloque"><div class="bloque-body vacio">
      No se pudo cargar ${cfg.jsonPath}. Verifica la ruta del archivo.
    </div></div>`;
    console.error("[MPC] Error cargando histórico de conceptos:", e);
  }
}

document.addEventListener("DOMContentLoaded", initHistoricoConceptos);
