#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
generar_historico_ranking.py  —  MPC Ingresos (Municipalidad Provincial de Chiclayo)

Construye data/historico_ranking.json con el ranking de las 38 municipalidades
del departamento de Lambayeque, por año (2015-2025 cerrados + 2026 en curso):

    PIA, PIM (anuales), Recaudado Enero-Octubre y % de avance (Rec. Ene-Oct / PIM)

Entradas (todas exportadas de Consulta Amigable de Ingresos del MEF, formato HTML-.xls):
    historico/<AAAA>/anual_<AAAA>.xls     -> PIA, PIM y Recaudado anual por municipalidad
    historico/<AAAA>/1T_<AAAA>.xls        -> Recaudado 1er trimestre  (Ene-Mar)
    historico/<AAAA>/2T_<AAAA>.xls        -> Recaudado 2do trimestre  (Abr-Jun)
    historico/<AAAA>/3T_<AAAA>.xls        -> Recaudado 3er trimestre  (Jul-Sep)
    historico/<AAAA>/octubre_<AAAA>.xls   -> Recaudado de octubre
    xls/ranking.xls                       -> 2026 en curso (PIA, PIM, Recaudado acumulado)

Ene-Oct = 1T + 2T + 3T + octubre.

Validaciones (si alguna falla el script NO escribe el JSON):
    - cada archivo declara el año y el período esperados (evita archivos mal descargados)
    - las 38 municipalidades están en el anual; un archivo de período puede omitir una fila
      sin recaudación (el MEF no la exporta): se toma como 0 SOLO si el total del período cuadra
    - la suma de las 38 cuadra con el total departamental del período (tolerancia 20 soles por redondeo)
    - Ene-Oct <= Recaudado anual por municipalidad (solo aviso)

Uso (desde la raíz del repo):
    python scripts/generar_historico_ranking.py
    python scripts/generar_historico_ranking.py --raiz RUTA --sin-confirmar
"""
import argparse
import html
import json
import os
import re
import sys
from datetime import datetime

AÑOS = list(range(2015, 2026))
PERIODOS = [  # (archivo, texto esperado en la fila del período)
    ("1T_{y}.xls", r"^Trimestre I:"),
    ("2T_{y}.xls", r"^Trimestre II:"),
    ("3T_{y}.xls", r"^Trimestre III:"),
    ("octubre_{y}.xls", r"^Mes 10:"),
]
N_MUNIS = 38
TOLERANCIA = 20  # soles: el MEF redondea cada fila (máx. 0.5 sol x 38 municipalidades)
RE_MUNI = re.compile(r"^(\d{6})-(\d{6}):\s*(.+)$")


def n(s):
    s = (s or "").replace(",", "").strip()
    return int(s) if s else 0


def leer_filas(ruta):
    with open(ruta, encoding="latin-1") as f:
        t = f.read()
    filas = []
    for tr in re.findall(r"<tr.*?</tr>", t, re.S | re.I):
        celdas = [
            re.sub(r"\s+", " ", html.unescape(re.sub(r"<.*?>", "", c))).strip()
            for c in re.findall(r"<t[dh].*?</t[dh]>", tr, re.S | re.I)
        ]
        filas.append(celdas)
    return filas


def parsear(ruta, año, periodo_re=None):
    """Devuelve (munis, total_depto, fecha_consulta). munis: ubigeo -> (nombre, pia, pim, rec)."""
    filas = leer_filas(ruta)
    cab = filas[0][0] if filas and filas[0] else ""
    m = re.search(r"Año de Ejecución:\s*(\d{4})", cab)
    if not m or int(m.group(1)) != año:
        raise ValueError(f"{os.path.basename(ruta)}: declara año {m.group(1) if m else '?'} y se esperaba {año}")
    fecha = re.search(r"Fecha de la Consulta:\s*(\d{1,2}-\w+-\d{4})", cab)
    fecha = fecha.group(1) if fecha else ""
    if periodo_re:
        if not any(f and re.match(periodo_re, f[0]) for f in filas):
            raise ValueError(f"{os.path.basename(ruta)}: no contiene el período esperado {periodo_re}")
    munis, depto_total = {}, None
    for f in filas:
        if not f:
            continue
        if periodo_re and re.match(periodo_re, f[0]):
            depto_total = n(f[-1])
        if not periodo_re and f[0].startswith("Departamento 14"):
            depto_total = n(f[-1])
        mm = RE_MUNI.match(f[0])
        if mm and len(f) >= 4:
            munis[mm.group(1) + "-" + mm.group(2)] = (mm.group(3).strip(), n(f[1]), n(f[2]), n(f[3]))
    return munis, depto_total, fecha


def construir(raiz):
    errores, avisos = [], []
    data = {}
    nombres = {}
    for y in AÑOS:
        carpeta = os.path.join(raiz, "historico", str(y))
        try:
            anual, tot_anual, _ = parsear(os.path.join(carpeta, f"anual_{y}.xls"), y)
        except (OSError, ValueError) as e:
            errores.append(str(e))
            continue
        if len(anual) != N_MUNIS:
            errores.append(f"anual_{y}: {len(anual)} municipalidades (se esperaban {N_MUNIS})")
        if abs(sum(v[3] for v in anual.values()) - tot_anual) > TOLERANCIA:
            errores.append(f"anual_{y}: la suma de municipalidades no cuadra con el total departamental")
        eneoct = {k: 0 for k in anual}
        for plantilla, per_re in PERIODOS:
            nom = plantilla.format(y=y)
            try:
                p, tot, _ = parsear(os.path.join(carpeta, nom), y, per_re)
            except (OSError, ValueError) as e:
                errores.append(str(e))
                continue
            if set(p) - set(anual):
                errores.append(f"{nom}: municipalidades que no están en anual_{y}: {sorted(set(p) - set(anual))}")
                continue
            for k in sorted(set(anual) - set(p)):
                avisos.append(f"{y} {nom}: falta {anual[k][0]} -> se toma como 0 (sin recaudación ese período; "
                              f"se verifica con el total)")
            s = sum(v[3] for v in p.values())
            if tot is None or abs(s - tot) > TOLERANCIA:
                errores.append(f"{nom}: suma de las {len(p)} municipalidades = {s:,} vs total período = {tot}")
            for k, v in p.items():
                eneoct[k] += v[3]
        for k, v in anual.items():
            nombres[k] = v[0]
            if eneoct.get(k, 0) > v[3]:
                avisos.append(f"{y} {v[0]}: Ene-Oct ({eneoct[k]:,}) > anual ({v[3]:,})")
        data[y] = {
            "cerrado": True,
            "corte": f"Ene–Oct {y}",
            "munis": {
                k: {"pia": v[1], "pim": v[2], "eneoct": eneoct[k], "rec_anual": v[3]} for k, v in anual.items()
            },
        }
    # 2026 en curso
    ruta26 = os.path.join(raiz, "xls", "ranking.xls")
    try:
        m26, _t, fecha26 = parsear(ruta26, 2026)
        if len(m26) != N_MUNIS:
            errores.append(f"ranking.xls (2026): {len(m26)} municipalidades (se esperaban {N_MUNIS})")
        for k, v in m26.items():
            nombres.setdefault(k, v[0])
        data[2026] = {
            "cerrado": False,
            "corte": f"Ene–Oct 2026 (al {fecha26})",
            "fecha_consulta": fecha26,
            "munis": {k: {"pia": v[1], "pim": v[2], "eneoct": v[3], "rec_anual": None} for k, v in m26.items()},
        }
    except (OSError, ValueError) as e:
        errores.append(str(e))
    return data, nombres, errores, avisos


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--raiz", default=os.path.join(os.path.dirname(os.path.abspath(__file__)), ".."))
    ap.add_argument("--sin-confirmar", action="store_true")
    ap.add_argument("--salida", default=None)
    a = ap.parse_args()
    raiz = os.path.abspath(a.raiz)
    data, nombres, errores, avisos = construir(raiz)

    for av in avisos:
        print("AVISO:", av)
    if errores:
        print("\nERRORES — no se generó el JSON:")
        for e in errores:
            print("  -", e)
        sys.exit(1)

    años_orden = sorted(data)
    salida = {
        "generado": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "fuente": "MEF - Consulta Amigable de Ingresos (Gob.Loc./Mancom. M: Municipalidades, Dpto. 14 Lambayeque)",
        "nota": "eneoct = Recaudado acumulado Enero-Octubre. pia/pim son anuales (PIM final del año; en el año en curso, el vigente). "
                "avance = eneoct / pim.",
        "años": años_orden,
        "municipalidades": nombres,
        "datos": {str(y): data[y] for y in años_orden},
    }
    destino = a.salida or os.path.join(raiz, "data", "historico_ranking.json")
    print(f"\nAños: {años_orden[0]}–{años_orden[-1]} · municipalidades: {len(nombres)}")
    for y in años_orden:
        mm = data[y]["munis"]
        tot = sum(v["eneoct"] for v in mm.values())
        print(f"  {y}: Ene-Oct total = {tot:>15,}   MPC = {mm['140101-301212']['eneoct']:>13,}")
    if not a.sin_confirmar:
        if input(f"\nEscribir {destino}? [s/N] ").strip().lower() != "s":
            print("Cancelado.")
            return
    os.makedirs(os.path.dirname(destino), exist_ok=True)
    with open(destino, "w", encoding="utf-8") as f:
        json.dump(salida, f, ensure_ascii=False, indent=1)
    print("OK ->", destino)


if __name__ == "__main__":
    main()
