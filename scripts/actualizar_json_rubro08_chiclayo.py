"""
actualizar_json_rubro08_chiclayo.py
-------------------------------
Lee los 5 archivos .xls descargados por
    scripts/descargar_rubro08_chiclayo.py
y actualiza automáticamente el corte "hasta la fecha" (2026) en
    data/historico_conceptos_rubro08.json

No usa pandas ni lxml -- parseo con regex sobre el HTML disfrazado de
XLS, mismo patrón que mpl-ingresos/scripts/actualizar_json_rubro08.py.

MODO DE USO (desde la raíz del repo orad-ingresos-chiclayo):
    python scripts/actualizar_json_rubro08_chiclayo.py

Qué hace:
  1. Abre cada .xls en xlsrubro08/
  2. Busca la fila exacta que corresponde a cada concepto (reglas abajo)
  3. Muestra en pantalla el valor VIEJO vs NUEVO de cada concepto
  4. Pide confirmación antes de escribir el JSON
  5. Actualiza 'anual' Y 'eneago' del año 2026 con el mismo valor
     (es el mismo "acumulado a la fecha" mientras el año no cierra)

Reglas de extracción (confirmadas contra los .xls de muestra reales
enviados por Juan el 24-ago-2026):
  predial.xls           -> fila "1: PREDIAL"                              -> Predial Corriente
  predial.xls           -> fila "2: PREDIAL - REGULARIZACIÓN TRIBUTARIA"  -> Regularización Predial
  alcabala.xls          -> fila "2: ALCABALA" (OJO: aquí es "2:", no "1:"
                            como en MPL -- en este archivo la fila "1:"
                            corresponde a PREDIAL, no a Alcabala, porque
                            el portal exporta el nivel "Específica"
                            completo con ambos conceptos juntos)
  patrivehicular.xls    -> fila "1: AL PATRIMONIO VEHICULAR"              -> Al Patrimonio Vehicular
  impuestoselectivo.xls -> fila que contiene "IMPUESTO SELECTIVO A PRODUCTOS ESPECIFICOS"
  intereses.xls         -> fila "1: INTERESES"                            -> Intereses
"""

import re
import json
from pathlib import Path
import subprocess
from datetime import datetime

CARPETA_RUBRO08 = Path("xlsrubro08")
JSON_PATH       = Path("data") / "historico_conceptos_rubro08.json"


def leer_filas(archivo):
    """Extrae todas las filas <tr> del XLS (HTML disfrazado) como listas de celdas."""
    with open(archivo, encoding="latin-1") as f:
        html = f.read()
    filas_html = re.findall(r"<tr[^>]*>(.*?)</tr>", html, re.S)
    filas = []
    for fh in filas_html:
        celdas = re.findall(r"<td[^>]*>(.*?)</td>", fh, re.S)
        celdas = [re.sub(r"<.*?>", "", c).strip() for c in celdas]
        if celdas:
            filas.append(celdas)
    return filas


def num(texto):
    """Convierte '2,034,212' -> 2034212"""
    try:
        return round(float(texto.replace(",", "").strip()))
    except (ValueError, AttributeError):
        return None


def buscar_fila(filas, patron_exacto=None, contiene=None):
    """
    Busca una fila cuya primera celda coincida.
    patron_exacto: la celda debe EMPEZAR con este texto (ej. '2: ALCABALA')
    contiene: la celda debe CONTENER este texto en cualquier parte
    Devuelve el valor de la 4ta columna (Recaudado) o None si no la encuentra.
    """
    for fila in filas:
        if len(fila) < 4:
            continue
        c0 = fila[0].upper()
        if patron_exacto and c0.startswith(patron_exacto.upper()):
            return num(fila[3])
        if contiene and contiene.upper() in c0:
            return num(fila[3])
    return None


# Reglas: (archivo, [(concepto_json, patron_exacto, contiene)])
REGLAS = [
    ("predial.xls", [
        ("PREDIAL_CORRIENTE",      "1: PREDIAL", None),
        ("REGULARIZACION_PREDIAL", "2: PREDIAL - REGULARIZACIÓN", None),
    ]),
    ("alcabala.xls", [
        # OJO: "2: ALCABALA", no "1:" -- en MPC el nivel "Específica"
        # exportado trae juntas "1: PREDIAL" y "2: ALCABALA". Se ignora
        # la fila 1 (PREDIAL) porque ese valor ya viene de predial.xls.
        ("ALCABALA", "2: ALCABALA", None),
    ]),
    ("patrivehicular.xls", [
        ("PATRIMONIO_VEHICULAR", "1: AL PATRIMONIO VEHICULAR", None),
    ]),
    ("impuestoselectivo.xls", [
        ("IMPUESTO_SELECTIVO", None, "IMPUESTO SELECTIVO A PRODUCTOS ESPECIFICOS"),
    ]),
    ("intereses.xls", [
        ("INTERESES", "1: INTERESES", None),
    ]),
]


def git_push_automatico():
    """
    Ejecuta git add data/historico_conceptos_rubro08.json -> commit -> push.
    Si cualquier paso falla, muestra el error y se detiene -- nunca deja
    el repo a medias ni lanza una excepción que rompa el script.
    """
    fecha_hoy = datetime.now().strftime("%d/%m/%Y")
    mensaje = f"Actualizar corte Rubro 08 - {fecha_hoy}"

    print("\n" + "=" * 60)
    print("  GIT — Publicando en GitHub Pages")
    print("=" * 60)

    pasos = [
        ("git add",    ["git", "add", str(JSON_PATH)]),
        ("git commit", ["git", "commit", "-m", mensaje]),
        ("git push",   ["git", "push"]),
    ]

    for nombre_paso, cmd in pasos:
        print(f"  → {nombre_paso}...", end=" ", flush=True)
        resultado = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8")

        if resultado.returncode == 0:
            print("OK")
            if nombre_paso == "git commit" and resultado.stdout.strip():
                print(f"     {resultado.stdout.strip().splitlines()[0]}")
        else:
            salida = (resultado.stdout + resultado.stderr)
            sin_cambios = (
                "nothing to commit" in salida
                or "nothing added to commit" in salida
                or "no changes added to commit" in salida
            )
            if nombre_paso == "git commit" and sin_cambios:
                print("sin cambios (el JSON ya estaba subido, no hay nada nuevo)")
                return
            print("FALLÓ")
            detalle = salida.strip() or "sin detalle"
            print(f"  [ERROR {nombre_paso}] {detalle}")
            print("  ⚠ Revisa el error y sube manualmente si hace falta:")
            print(f'     git add {JSON_PATH.as_posix()}')
            print(f'     git commit -m "{mensaje}"')
            print("     git push")
            return

    print(f"\n  ✅ GitHub Pages actualizado — commit: \"{mensaje}\"")


def main():
    if not JSON_PATH.exists():
        print(f"❌ No se encontró {JSON_PATH.resolve()}")
        print("   Ejecuta este script desde la raíz del repo orad-ingresos-chiclayo.")
        return

    data = json.loads(JSON_PATH.read_text(encoding="utf-8"))

    print("=" * 60)
    print("  Actualizar corte 2026 — Rubro 08 (MPC)")
    print(f"  Leyendo desde: {CARPETA_RUBRO08.resolve()}")
    print("=" * 60)

    cambios = {}
    faltantes = []

    for archivo, reglas in REGLAS:
        ruta = CARPETA_RUBRO08 / archivo
        if not ruta.exists():
            print(f"\n⚠️  {archivo} no encontrado en {CARPETA_RUBRO08} — se omite.")
            faltantes.append(archivo)
            continue

        filas = leer_filas(ruta)
        print(f"\n📄 {archivo}")
        for concepto, patron_exacto, contiene in reglas:
            valor_nuevo = buscar_fila(filas, patron_exacto, contiene)
            if valor_nuevo is None:
                print(f"   ⚠️  No se encontró la fila esperada para {concepto} — se omite.")
                continue
            valor_viejo = data.get(concepto, {}).get("eneago", {}).get("2026")
            label = data.get(concepto, {}).get("label", concepto)
            flecha = "→" if valor_viejo != valor_nuevo else "= (sin cambio)"
            print(f"   {label:45s} S/ {valor_viejo:>12,}  {flecha}  S/ {valor_nuevo:>12,}"
                  if valor_viejo is not None else
                  f"   {label:45s} (sin valor previo) {flecha} S/ {valor_nuevo:>12,}")
            cambios[concepto] = valor_nuevo

    if not cambios:
        print("\nNo hay cambios para aplicar.")
        return

    print("\n" + "=" * 60)
    if faltantes:
        print(f"  ⚠️  {len(faltantes)} archivo(s) no encontrados: {', '.join(faltantes)}")
        print("     Esos conceptos NO se actualizarán.")
    resp = input(f"\n¿Aplicar estos {len(cambios)} cambios a {JSON_PATH}? [s/N]: ").strip().lower()
    if resp != "s":
        print("Cancelado. No se modificó el JSON.")
        return

    for concepto, valor in cambios.items():
        data[concepto]["anual"]["2026"] = valor
        data[concepto]["eneago"]["2026"] = valor

    JSON_PATH.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"\n✅ {JSON_PATH} actualizado con {len(cambios)} concepto(s).")

    git_push_automatico()


if __name__ == "__main__":
    main()
