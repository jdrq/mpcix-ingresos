"""
descargar_xls_mef_ingresos_chiclayo.py
-------------------------------
Automatiza la descarga de los 4 archivos .xls de Consulta Amigable de
INGRESOS (MEF) para la Municipalidad Provincial de Chiclayo (MPC).

  fuente.xls    -> Ejecucion por Fuente de Financiamiento
  rubro.xls     -> Ejecucion por Rubro de Ingreso
  generica.xls  -> Ejecucion por Generica de Ingreso
  ranking.xls   -> Todas las municipalidades del Dpto. Lambayeque

MODO DE USO:
    pip install playwright
    playwright install chromium
    python descargar_xls_mef_ingresos_chiclayo.py

Los archivos se guardan en la carpeta xls/ del proyecto
(donde el index.html los espera con fetch("xls/...")).

Mismo motor robusto ya validado en descargar_xls_mef_ingresos_lamb.py
(MPL, 20-ago-2026) y en los scripts de rubro08/rubro09 de ambas
municipalidades. Los 3 mecanismos que evitan fallos intermitentes por
respuestas lentas del servidor MEF:

  1. clic_con_reintento(): polling (revisa cada 1s hasta 18s) en vez de
     un solo chequeo temprano, con hasta 3 intentos por clic y escalada
     a DOBLE clic desde el 2do intento (el UpdatePanel del portal a
     veces "traga" un clic simple si llega justo cuando termina el
     postback anterior).
  2. Checkpoint de fila: despues de clickear una fila del navegador de
     dimensiones, confirma que su radio (input:checked) quedo REALMENTE
     marcado antes de seguir -- hasta 3 intentos, escalando a doble clic.
  3. Pausa de 2s entre archivos: el servidor tiene un respiro entre una
     descarga y la siguiente, en vez de encadenarse sin descanso.

Se mantiene el criterio "todo o nada" en Git: solo se publica en GitHub
Pages si los 4 archivos se descargaron correctamente.

Municipalidad: 140101-301212 (MUNICIPALIDAD PROVINCIAL DE CHICLAYO)
Rutas de navegación confirmadas con playwright codegen real (24-ago-2026).
"""

from playwright.sync_api import sync_playwright
import subprocess
import time
import shutil
from datetime import datetime
from pathlib import Path

# --------------------------------------------------------------------
# CONFIGURACION
# --------------------------------------------------------------------
CARPETA_DESTINO = Path("xls")   # el index.html hace fetch("xls/archivo.xls")
ANIO            = "2026"
URL_BASE        = (
    f"https://apps5.mineco.gob.pe/transparenciaingresos/"
    f"Navegador/default.aspx?y={ANIO}"
)
FRAME_SELECTOR = "#frame0"

# Pasos comunes a los 4 archivos: TOTAL -> ... -> Municipalidad (MPC).
# (el clic en "TOTAL" se hace aparte, ver navegar_hasta_mpc)
PASOS_MPC = [
    ("Nivel de Gobierno", "M: GOBIERNOS LOCALES"),
    ("Gob.Loc./Mancom.",  "M: MUNICIPALIDADES"),
    ("Departamento",      ": LAMBAYEQUE"),
    ("Municipalidad",     "140101-301212: MUNICIPALIDAD"),
]


# --------------------------------------------------------------------
# CLIC CON REINTENTO (polling + escalada a doble clic)
# Mismo motor que descargar_xls_mef_ingresos_lamb.py / rubro08 / rubro09
# --------------------------------------------------------------------
def clic_con_reintento(fl, page, rol=None, nombre=None, intentos=3,
                        verificar=None, exacto=True, espera_verificacion=18):
    """
    Hace clic con reintentos automaticos, escalando a DOBLE CLIC desde el
    2do intento. Usa polling (revisa cada 1s hasta `espera_verificacion`
    segundos) en vez de un solo chequeo -- asi no confunde "el servidor
    esta tardando" con "el clic no sirvio".

    verificar: funcion sin argumentos que retorna True si el clic surtio
    efecto. Si no se pasa, solo se espera el postback y se asume exito.
    """
    locator = fl.get_by_role(rol, name=nombre, exact=exacto)
    for intento in range(intentos):
        try:
            if intento == 0:
                locator.click(timeout=10000)
            else:
                locator.click(timeout=10000)
                time.sleep(0.3)
                locator.click(timeout=10000)
        except Exception as e:
            print(f"  [AVISO] '{nombre}' no fue accionable a tiempo "
                  f"({e.__class__.__name__}), reintentando...")
            time.sleep(2)
            continue

        try:
            page.wait_for_load_state("networkidle", timeout=5000)
        except Exception:
            pass

        if verificar is None:
            time.sleep(0.8)
            return

        for _ in range(espera_verificacion):
            try:
                if verificar():
                    return
            except Exception:
                pass
            time.sleep(1)

        print(f"  [REINTENTO clic {intento + 1}/{intentos}] '{nombre}' "
              f"no surtio efecto tras {espera_verificacion}s, reintentando...")

    raise RuntimeError(
        f"El clic en '{nombre}' no surtio efecto tras {intentos} intentos "
        f"(ni con doble clic, ni con {espera_verificacion}s de espera cada uno). "
        f"Revisar manualmente."
    )


def clic_fila_con_checkpoint(fl, page, texto_fila, intentos=3,
                              espera_verificacion=18):
    """
    Clickea una fila del navegador de dimensiones y confirma que su
    radio (input:checked) quedo REALMENTE marcado antes de avanzar.
    Hasta `intentos` intentos, escalando a doble clic desde el 2do.
    """
    for intento in range(intentos):
        if intento == 0:
            fl.get_by_role("cell", name=texto_fila).click()
        else:
            fl.get_by_role("cell", name=texto_fila).click()
            time.sleep(0.3)
            fl.get_by_role("cell", name=texto_fila).click()

        try:
            page.wait_for_load_state("networkidle", timeout=5000)
        except Exception:
            pass
        time.sleep(1)

        for _ in range(espera_verificacion):
            try:
                fila = fl.locator(f"tr:has-text('{texto_fila}')").first
                if fila.locator("input:checked").count() > 0:
                    return
            except Exception:
                pass
            time.sleep(1)

        print(f"  [REINTENTO {intento + 1}/{intentos}] '{texto_fila}' no se "
              f"marco, volviendo a clickear (doble clic)...")

    raise RuntimeError(
        f"El nivel '{texto_fila}' no quedo marcado (radio) tras "
        f"{intentos} intentos. Revisar manualmente."
    )


def fl(page):
    """Devuelve el frame_locator del iframe principal."""
    return page.frame_locator(FRAME_SELECTOR)


def backup_y_guardar(descarga, nombre):
    """Guarda el archivo descargado en xls/ con backup previo."""
    destino = CARPETA_DESTINO / nombre
    if destino.exists():
        respaldo = CARPETA_DESTINO / "_respaldo_anterior" / nombre
        respaldo.parent.mkdir(exist_ok=True)
        shutil.copy(destino, respaldo)
        print(f"  -> Backup guardado en {respaldo}")
    descarga.save_as(destino)
    print(f"  [OK] {nombre} -> {destino.resolve()}")


# --------------------------------------------------------------------
# GIT: commit + push automatico (solo si 4/4 exitosos)
# --------------------------------------------------------------------
def git_push_automatico():
    """
    Ejecuta git add xls/ -> git commit -> git push.
    Si cualquier paso falla, imprime el error y NO continua.
    Nunca lanza excepcion -- el fallo es informativo, no fatal.
    """
    fecha_hoy = datetime.now().strftime("%d/%m/%Y")
    mensaje   = f"Actualizacion {fecha_hoy}"

    pasos = [
        ("git add",    ["git", "add", "xls/"]),
        ("git commit", ["git", "commit", "-m", mensaje]),
        ("git push",   ["git", "push"]),
    ]

    print("\n" + "-" * 60)
    print("  GIT -- Publicando en GitHub Pages")
    print("-" * 60)

    for nombre_paso, cmd in pasos:
        print(f"  -> {nombre_paso}...", end=" ", flush=True)
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
            salida = (resultado.stderr or resultado.stdout or "").strip()
            sin_cambios = (
                "nothing to commit" in salida
                or "nothing added to commit" in salida
                or "no changes added to commit" in salida
            )
            if nombre_paso == "git commit" and sin_cambios:
                print("sin cambios (los xls no cambiaron respecto al ultimo commit)")
                return True
            print("FALLO")
            print(f"  [ERROR {nombre_paso}] {salida or 'sin detalle'}")
            print("  Aviso: el push fue cancelado. Revisa el error y ejecuta manualmente:")
            print(f'     git add xls/ && git commit -m "{mensaje}" && git push')
            return False

    print("\n  OK -- GitHub Pages actualizado correctamente.")
    print(f'     Commit: "{mensaje}"')
    return True


# --------------------------------------------------------------------
# FLUJO COMUN: bajar hasta la MPC, con reintentos y checkpoint de fila
# --------------------------------------------------------------------
def navegar_hasta_mpc(page):
    """
    Navega desde TOTAL hasta la MPC (140101-301212), con
    clic_con_reintento() en cada boton y checkpoint de radio en cada
    fila seleccionada.
    """
    print(f"  -> Cargando {URL_BASE}")
    page.goto(URL_BASE)
    try:
        page.wait_for_load_state("networkidle", timeout=5000)
    except Exception:
        pass

    f = fl(page)

    # TOTAL
    clic_con_reintento(f, page, rol="cell", nombre="TOTAL", intentos=3)

    for etiqueta_boton, texto_fila in PASOS_MPC:
        print(f"  -> {etiqueta_boton} | {texto_fila}")

        # 1) Clic en el boton de la cadena, con reintento + verificacion
        #    de que la fila esperada se hizo visible.
        clic_con_reintento(
            f, page, rol="button", nombre=etiqueta_boton, intentos=3,
            verificar=lambda tf=texto_fila: f.get_by_role("cell", name=tf).first.is_visible()
        )
        time.sleep(1)  # margen extra: el servidor del MEF a veces es lento

        # 2) Clic en la fila, con checkpoint: confirma que el radio quedo
        #    REALMENTE marcado antes de avanzar.
        clic_fila_con_checkpoint(f, page, texto_fila, intentos=3)

    time.sleep(1.5)


def exportar(f, page, nombre_archivo):
    """Clic en 'Exportar' con reintento si el primer clic no dispara la descarga."""
    print(f"  -> Exportando {nombre_archivo}")
    descarga = None
    for intento in range(2):
        try:
            with page.expect_download(timeout=15000) as descarga_info:
                if intento == 0:
                    f.get_by_role("link", name="Exportar").click()
                else:
                    print("  [REINTENTO Exportar] no se disparo la "
                          "descarga, reintentando con doble clic...")
                    f.get_by_role("link", name="Exportar").click()
                    time.sleep(0.3)
                    f.get_by_role("link", name="Exportar").click()
            descarga = descarga_info.value
            break
        except Exception:
            if intento == 1:
                raise RuntimeError(
                    "El clic en 'Exportar' no disparo la descarga tras "
                    "2 intentos. Revisar manualmente."
                )
    backup_y_guardar(descarga, nombre_archivo)


# --------------------------------------------------------------------
# DESCARGA DE CADA ARCHIVO
# --------------------------------------------------------------------
def descargar_fuente(page):
    """fuente.xls -- pivote a 'Fuente'."""
    navegar_hasta_mpc(page)
    f = fl(page)
    print("  -> Pivotando a 'Fuente'")
    clic_con_reintento(f, page, rol="button", nombre="Fuente", intentos=3)
    time.sleep(1.5)
    exportar(f, page, "fuente.xls")


def descargar_rubro(page):
    """rubro.xls -- pivote a 'Rubro'."""
    navegar_hasta_mpc(page)
    f = fl(page)
    print("  -> Pivotando a 'Rubro'")
    clic_con_reintento(f, page, rol="button", nombre="Rubro", intentos=3)
    time.sleep(1.5)
    exportar(f, page, "rubro.xls")


def descargar_generica(page):
    """generica.xls -- pivote a 'Generica'."""
    navegar_hasta_mpc(page)
    f = fl(page)
    print("  -> Pivotando a 'Generica'")
    clic_con_reintento(f, page, rol="button", nombre="Genérica", intentos=3)
    time.sleep(1.5)
    exportar(f, page, "generica.xls")


def descargar_ranking(page):
    """
    ranking.xls -- baja hasta : LAMBAYEQUE, clic en 'Municipalidad'
    SIN seleccionar ninguna fila -> lista las 38 municipalidades -> Exportar.
    """
    print(f"  -> Cargando {URL_BASE}")
    page.goto(URL_BASE)
    try:
        page.wait_for_load_state("networkidle", timeout=5000)
    except Exception:
        pass

    f = fl(page)

    clic_con_reintento(f, page, rol="cell", nombre="TOTAL", intentos=3)

    pasos_hasta_departamento = PASOS_MPC[:3]  # Nivel Gob. -> Gob.Loc. -> Departamento
    for etiqueta_boton, texto_fila in pasos_hasta_departamento:
        print(f"  -> {etiqueta_boton} | {texto_fila}")
        clic_con_reintento(
            f, page, rol="button", nombre=etiqueta_boton, intentos=3,
            verificar=lambda tf=texto_fila: f.get_by_role("cell", name=tf).first.is_visible()
        )
        time.sleep(1)
        clic_fila_con_checkpoint(f, page, texto_fila, intentos=3)

    # Municipalidad SIN seleccionar fila -> queda el listado de 38
    print("  -> Boton 'Municipalidad' (sin seleccionar fila)")
    clic_con_reintento(f, page, rol="button", nombre="Municipalidad", intentos=3)
    time.sleep(1.5)

    exportar(f, page, "ranking.xls")


# --------------------------------------------------------------------
# MAIN
# --------------------------------------------------------------------
TAREAS = [
    ("fuente.xls",   descargar_fuente),
    ("rubro.xls",    descargar_rubro),
    ("generica.xls", descargar_generica),
    ("ranking.xls",  descargar_ranking),
]

def main():
    print("=" * 60)
    print("  Descarga XLS -- Consulta Amigable de INGRESOS")
    print(f"  Municipalidad Provincial de Chiclayo - {ANIO}")
    print(f"  Destino: {CARPETA_DESTINO.resolve()}")
    print("=" * 60)

    CARPETA_DESTINO.mkdir(exist_ok=True)

    exitosos = []
    fallidos  = []

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False, slow_mo=200)
        page    = browser.new_page()

        for i, (nombre, fn) in enumerate(TAREAS, 1):
            print(f"\n[{i}/{len(TAREAS)}] {nombre}")
            print("-" * 40)
            try:
                fn(page)
                exitosos.append(nombre)
                time.sleep(2)  # pausa entre archivos, respiro para el servidor MEF
            except Exception as e:
                print(f"  [ERROR] {e}")
                fallidos.append(nombre)

                # Captura para diagnostico
                captura = Path(f"error_{nombre}.png")
                try:
                    page.screenshot(path=str(captura), full_page=True)
                    print(f"  [DIAGNOSTICO] Captura: {captura.resolve()}")
                except Exception:
                    pass

                # Un archivo fallido no detiene el lote -- se sigue con
                # el siguiente (mismo criterio que rubro08.py / rubro09.py).

        input("\nPresiona ENTER para cerrar el navegador...")
        browser.close()

    # -- Resumen ---------------------------------------------------
    print("\n" + "=" * 60)
    print("  RESUMEN DE DESCARGA")
    print("=" * 60)
    for n in exitosos:
        print(f"  OK  {n}")
    for n in fallidos:
        print(f"  X   {n}  <- revisar manualmente")

    # -- Decision Git: todo o nada -----------------------------------
    total_esperado = len(TAREAS)
    if len(exitosos) == total_esperado and not fallidos:
        # 4/4 -- publicar en GitHub Pages automaticamente
        git_push_automatico()
    else:
        # Fallo al menos 1 -- NO subir nada
        print("\n" + "=" * 60)
        print("  GIT -- NO SE SUBIO NADA AL REPOSITORIO")
        print("=" * 60)
        print(f"  {len(fallidos)} de {total_esperado} archivo(s) fallaron.")
        print("  El repositorio queda sin cambios para evitar publicar datos incompletos.")
        print(f"\n  Descarga manual: {URL_BASE}")
        print("  Una vez resuelto el problema, ejecuta el script nuevamente.")


if __name__ == "__main__":
    main()
