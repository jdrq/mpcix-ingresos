"""
descargar_xls_mef_ingresos.py
-------------------------------
Automatiza la descarga de los 4 archivos .xls de Consulta Amigable de
INGRESOS (MEF) para la Municipalidad Provincial de Chiclayo.

  fuente.xls    → Ejecución por Fuente de Financiamiento
  rubro.xls     → Ejecución por Rubro de Ingreso
  generica.xls  → Ejecución por Genérica de Ingreso
  ranking.xls   → Todas las municipalidades del Dpto. Lambayeque
                  (incluye a Chiclayo — el portal no permite exportar
                  el ranking filtrado solo por provincia)

MODO DE USO:
    pip install playwright
    playwright install chromium
    python descargar_xls_mef_ingresos.py

Los archivos se guardan en la carpeta xls/ del proyecto
(donde el index.html los espera con fetch("xls/...")).
"""

from playwright.sync_api import sync_playwright
import subprocess
import time
import shutil
from datetime import datetime
from pathlib import Path

# ─────────────────────────────────────────────────────────────────
# CONFIGURACIÓN
# ─────────────────────────────────────────────────────────────────
CARPETA_DESTINO = Path("xls")   # el index.html hace fetch("xls/archivo.xls")
ANIO            = "2026"
URL_BASE        = (
    f"https://apps5.mineco.gob.pe/transparenciaingresos/"
    f"Navegador/default.aspx?y={ANIO}"
)
FL = "#frame0"   # selector del iframe del portal


# ─────────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────────
def esperar(page, ms=5000):
    """Espera networkidle sin reventar si tarda."""
    try:
        page.wait_for_load_state("networkidle", timeout=ms)
    except Exception:
        pass


def fl(page):
    """Devuelve el content_frame del iframe principal."""
    return page.locator(FL).content_frame


def verificar_pivote(page, columna_esperada, nombre_boton, reintentos=3):
    """
    Confirma que la primera columna de la tabla realmente cambió al
    pivote esperado (ej. "Rubro") ANTES de exportar.

    Motivo: el portal a veces no re-renderiza a tiempo tras el clic
    (postback lento) y el sleep() fijo no alcanza a esperarlo — el
    resultado es exportar el archivo con el pivote anterior, sin que
    el script lo detecte. Esto causó que generica.xls y rubro.xls
    salieran con columna "Municipalidad" (el pivote de ranking.xls)
    en vez de su propio pivote.

    IMPORTANTE: el texto de la columna (ej. "Rubro") es el MISMO texto
    que el botón de pivote que acabamos de hacer clic — buscar ese
    texto en toda la página daría un falso positivo (el botón sigue
    visible aunque la tabla no haya cambiado). Por eso se busca
    específicamente dentro de celdas <td>/<th> de la tabla de datos,
    vía JS, no con un locator de texto genérico.
    """
    f = fl(page)
    for intento in range(1, reintentos + 1):
        esperar(page)
        time.sleep(1)

        encontrado = f.locator("td, th").evaluate_all(
            """(celdas, col) => {
                for (const c of celdas) {
                    if (c.children.length === 0 &&
                        c.textContent.trim().toLowerCase() === col.toLowerCase()) {
                        return true;
                    }
                }
                return false;
            }""",
            columna_esperada,
        )

        if encontrado:
            print(f"  ✓ Pivote confirmado: columna '{columna_esperada}' encontrada en la tabla")
            return True

        print(f"  ⚠ Pivote no confirmado (intento {intento}/{reintentos}): "
              f"no se encontró la columna '{columna_esperada}' en la tabla")
        if intento < reintentos:
            print(f"  → Reintentando clic en '{nombre_boton}'")
            f.get_by_role("button", name=nombre_boton).click()
            time.sleep(1.5)

    return False


def backup_y_guardar(descarga, nombre):
    """Guarda el archivo descargado en xls/ con backup previo."""
    destino = CARPETA_DESTINO / nombre
    if destino.exists():
        respaldo = CARPETA_DESTINO / "_respaldo_anterior" / nombre
        respaldo.parent.mkdir(exist_ok=True)
        shutil.copy(destino, respaldo)
        print(f"  → Backup guardado en {respaldo}")
    descarga.save_as(destino)
    print(f"  [OK] {nombre} → {destino.resolve()}")


# ─────────────────────────────────────────────────────────────────
# GIT: commit + push automático (solo si 4/4 exitosos)
# ─────────────────────────────────────────────────────────────────
def git_push_automatico():
    """
    Ejecuta git add xls/ → git commit → git push.
    Si cualquier paso falla, imprime el error y NO continúa.
    Nunca lanza excepción — el fallo es informativo, no fatal.
    """
    fecha_hoy = datetime.now().strftime("%d/%m/%Y")
    mensaje   = f"Actualización {fecha_hoy}"

    pasos = [
        ("git add",    ["git", "add", "xls/"]),
        ("git commit", ["git", "commit", "-m", mensaje]),
        ("git push",   ["git", "push"]),
    ]

    print("\n" + "─" * 60)
    print("  GIT — Publicando en GitHub Pages")
    print("─" * 60)

    for nombre_paso, cmd in pasos:
        print(f"  → {nombre_paso}...", end=" ", flush=True)
        resultado = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            encoding="utf-8",
        )
        if resultado.returncode == 0:
            print("OK")
            if nombre_paso == "git commit" and resultado.stdout.strip():
                print(f"     {resultado.stdout.strip().splitlines()[0]}")
        else:
            print("FALLÓ")
            detalle = (resultado.stderr or resultado.stdout or "sin detalle").strip()
            print(f"  [ERROR {nombre_paso}] {detalle}")
            print("  ⚠ El push fue cancelado. Revisa el error y ejecuta manualmente:")
            print(f'     git add xls/ && git commit -m "{mensaje}" && git push')
            return False

    print("\n  ✅ GitHub Pages actualizado correctamente.")
    print(f'     Commit: "{mensaje}"')
    return True


# ─────────────────────────────────────────────────────────────────
# FLUJO COMÚN: bajar hasta la MPC
# (Confirmado con codegen — nombres exactos del portal de Ingresos)
# ─────────────────────────────────────────────────────────────────
def navegar_hasta_mpc(page):
    """
    Navega desde TOTAL hasta la MPC (140101).
    Pasos confirmados con playwright codegen en el portal real:
      TOTAL → Nivel de Gobierno → M: GOBIERNOS LOCALES
            → Gob.Loc./Mancom. → M: MUNICIPALIDADES
            → Departamento     → : LAMBAYEQUE
            → Municipalidad    → 140101-301212: MUNICIPALIDAD...
    """
    print(f"  → Cargando {URL_BASE}")
    page.goto(URL_BASE)
    esperar(page)

    f = fl(page)

    # TOTAL
    f.get_by_role("cell", name="TOTAL", exact=True).click()
    esperar(page)
    time.sleep(1)

    # Nivel de Gobierno → M: GOBIERNOS LOCALES
    print("  → Nivel de Gobierno | M: GOBIERNOS LOCALES")
    f.get_by_role("button", name="Nivel de Gobierno").click()
    esperar(page); time.sleep(1)
    f.get_by_role("cell", name="M: GOBIERNOS LOCALES").click()
    esperar(page); time.sleep(1)

    # Gob.Loc./Mancom. → M: MUNICIPALIDADES
    print("  → Gob.Loc./Mancom. | M: MUNICIPALIDADES")
    f.get_by_role("button", name="Gob.Loc./Mancom.").click()
    esperar(page); time.sleep(1)
    f.get_by_role("cell", name="M: MUNICIPALIDADES").click()
    esperar(page); time.sleep(1)

    # Departamento → : LAMBAYEQUE
    print("  → Departamento | : LAMBAYEQUE")
    f.get_by_role("button", name="Departamento").click()
    esperar(page); time.sleep(1)
    f.get_by_role("cell", name=": LAMBAYEQUE").click()
    esperar(page); time.sleep(1)

    # Municipalidad → 140101-301212
    print("  → Municipalidad | 140101-301212: MUNICIPALIDAD...")
    f.get_by_role("button", name="Municipalidad").click()
    esperar(page); time.sleep(1)
    f.get_by_role("cell", name="140101-301212: MUNICIPALIDAD").click()
    esperar(page); time.sleep(1.5)


# ─────────────────────────────────────────────────────────────────
# DESCARGA DE CADA ARCHIVO
# ─────────────────────────────────────────────────────────────────
def descargar_fuente(page):
    """fuente.xls — pivote a 'Fuente'."""
    navegar_hasta_mpc(page)
    f = fl(page)
    print("  → Pivotando a 'Fuente'")
    f.get_by_role("button", name="Fuente").click()
    if not verificar_pivote(page, "Fuente de Financiamiento", "Fuente"):
        raise RuntimeError("No se pudo confirmar el pivote a 'Fuente' tras varios reintentos.")
    print("  → Exportando fuente.xls")
    with page.expect_download(timeout=30000) as dl:
        f.get_by_role("link", name="Exportar").click()
    backup_y_guardar(dl.value, "fuente.xls")


def descargar_rubro(page):
    """rubro.xls — pivote a 'Rubro'."""
    navegar_hasta_mpc(page)
    f = fl(page)
    print("  → Pivotando a 'Rubro'")
    f.get_by_role("button", name="Rubro").click()
    if not verificar_pivote(page, "Rubro", "Rubro"):
        raise RuntimeError("No se pudo confirmar el pivote a 'Rubro' tras varios reintentos.")
    print("  → Exportando rubro.xls")
    with page.expect_download(timeout=30000) as dl:
        f.get_by_role("link", name="Exportar").click()
    backup_y_guardar(dl.value, "rubro.xls")


def descargar_generica(page):
    """
    generica.xls — pivote a 'Genérica'.
    Nota codegen: en una sesión apareció un clic extra en
    'Nivel de Gobierno M:' antes del Departamento — es un
    artefacto de la sesión, no es necesario. La ruta estándar
    funciona igual.
    """
    navegar_hasta_mpc(page)
    f = fl(page)
    print("  → Pivotando a 'Genérica'")
    f.get_by_role("button", name="Genérica").click()
    if not verificar_pivote(page, "Genérica", "Genérica"):
        raise RuntimeError("No se pudo confirmar el pivote a 'Genérica' tras varios reintentos.")
    print("  → Exportando generica.xls")
    with page.expect_download(timeout=30000) as dl:
        f.get_by_role("link", name="Exportar").click()
    backup_y_guardar(dl.value, "generica.xls")


def descargar_ranking(page):
    """
    ranking.xls — baja hasta : LAMBAYEQUE, clic en 'Municipalidad'
    SIN seleccionar ninguna fila → lista las 38 municipalidades → Exportar.
    Confirmado con codegen: tras seleccionar LAMBAYEQUE se exporta
    directamente sin pivotar (el nivel queda en Municipalidad).
    """
    print(f"  → Cargando {URL_BASE}")
    page.goto(URL_BASE)
    esperar(page)

    f = fl(page)

    # TOTAL
    f.get_by_role("cell", name="TOTAL", exact=True).click()
    esperar(page); time.sleep(1)

    # Nivel de Gobierno → M: GOBIERNOS LOCALES
    print("  → Nivel de Gobierno | M: GOBIERNOS LOCALES")
    f.get_by_role("button", name="Nivel de Gobierno").click()
    esperar(page); time.sleep(1)
    f.get_by_role("cell", name="M: GOBIERNOS LOCALES").click()
    esperar(page); time.sleep(1)

    # Gob.Loc./Mancom. → M: MUNICIPALIDADES
    print("  → Gob.Loc./Mancom. | M: MUNICIPALIDADES")
    f.get_by_role("button", name="Gob.Loc./Mancom.").click()
    esperar(page); time.sleep(1)
    f.get_by_role("cell", name="M: MUNICIPALIDADES").click()
    esperar(page); time.sleep(1)

    # Departamento → : LAMBAYEQUE
    print("  → Departamento | : LAMBAYEQUE")
    f.get_by_role("button", name="Departamento").click()
    esperar(page); time.sleep(1)
    f.get_by_role("cell", name=": LAMBAYEQUE").click()
    esperar(page); time.sleep(1)

    # Municipalidad SIN seleccionar fila → queda el listado de 38
    print("  → Botón 'Municipalidad' (sin seleccionar fila)")
    f.get_by_role("button", name="Municipalidad").click()
    esperar(page); time.sleep(1.5)

    print("  → Exportando ranking.xls")
    with page.expect_download(timeout=30000) as dl:
        f.get_by_role("link", name="Exportar").click()
    backup_y_guardar(dl.value, "ranking.xls")


# ─────────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────────
TAREAS = [
    ("fuente.xls",   descargar_fuente),
    ("rubro.xls",    descargar_rubro),
    ("generica.xls", descargar_generica),
    ("ranking.xls",  descargar_ranking),
]

def main():
    print("=" * 60)
    print("  Descarga XLS — Consulta Amigable de INGRESOS")
    print(f"  Municipalidad Provincial de Chiclayo · {ANIO}")
    print(f"  Destino: {CARPETA_DESTINO.resolve()}")
    print("=" * 60)

    CARPETA_DESTINO.mkdir(exist_ok=True)

    exitosos = []
    fallidos  = []

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False, slow_mo=300)
        page    = browser.new_page()

        for i, (nombre, fn) in enumerate(TAREAS, 1):
            print(f"\n[{i}/{len(TAREAS)}] {nombre}")
            print("-" * 40)
            try:
                fn(page)
                exitosos.append(nombre)
            except Exception as e:
                print(f"  [ERROR] {e}")
                fallidos.append(nombre)

                # Captura para diagnóstico
                captura = Path(f"error_{nombre}.png")
                try:
                    page.screenshot(path=str(captura), full_page=True)
                    print(f"  [DIAGNÓSTICO] Captura: {captura.resolve()}")
                except Exception:
                    pass

                resp = input("\n  ¿Continuar con el siguiente archivo? [s/N]: ").strip().lower()
                if resp != "s":
                    print("  Deteniendo el script.")
                    break

        input("\nPresiona ENTER para cerrar el navegador...")
        browser.close()

    # ── Resumen ──────────────────────────────────────────────────
    print("\n" + "=" * 60)
    print("  RESUMEN DE DESCARGA")
    print("=" * 60)
    for n in exitosos:
        print(f"  ✓  {n}")
    for n in fallidos:
        print(f"  ✗  {n}  ← revisar manualmente")

    # ── Decisión Git: todo o nada ─────────────────────────────────
    total_esperado = len(TAREAS)
    if len(exitosos) == total_esperado and not fallidos:
        # 4/4 — publicar en GitHub Pages automáticamente
        git_push_automatico()
    else:
        # Falló al menos 1 — NO subir nada
        print("\n" + "=" * 60)
        print("  ⛔ GIT — NO SE SUBIÓ NADA AL REPOSITORIO")
        print("=" * 60)
        print(f"  {len(fallidos)} de {total_esperado} archivo(s) fallaron.")
        print("  El repositorio queda sin cambios para evitar publicar datos incompletos.")
        print(f"\n  Descarga manual: {URL_BASE}")
        print("  Una vez resuelto el problema, ejecuta el script nuevamente.")


if __name__ == "__main__":
    main()
