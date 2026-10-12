#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
generar_ranking_historico.py
============================
Lee los .xls de MEF Consulta Amigable guardados en  historico/<año>/  y genera
data/ranking_historico.json  (lo que consume el Bloque 10 del dashboard).

Para cada año necesita 4 archivos (Sólo Proyectos, agrupado por Municipalidad):
    1T_<año>.xls   -> Trimestre I
    2T_<año>.xls   -> Trimestre II
    3T_<año>.xls   -> Trimestre III
    octubre_<año>.xls -> Mes 10 (octubre)
Cada archivo trae SOLO el devengado de su periodo, por eso se SUMAN:
    Devengado Ene–Oct = 1T + 2T + 3T + octubre
(PIA y PIM vienen vacíos en estos archivos; no se usan.)

Solo usa la librería estándar de Python (no requiere pandas).
Uso:   python scripts/generar_ranking_historico.py
"""
import json, re, sys, html, datetime
from pathlib import Path

BASE = Path(__file__).resolve().parent.parent
HIST = BASE / "historico"
SALIDA = BASE / "data" / "ranking_historico.json"
CONTROL = BASE / "data" / "historico_egresos.json"   # totales MPCH ya conocidos (control cruzado)

PERIODOS = [("1T", "Trimestre I:"), ("2T", "Trimestre II:"),
            ("3T", "Trimestre III:"), ("octubre", "Mes 10")]
COD_MPCH = "140101"          # Municipalidad Provincial de Chiclayo
IDX_DEVENGADO = 6           # Municipalidad | PIA | PIM | Certif. | Compromiso | Atención | DEVENGADO | Girado | Avance%
MIN_MUNIS = 30              # un archivo con menos filas no viene desglosado por municipalidad
TOL_CONTROL = 10            # soles de diferencia aceptables contra historico_egresos.json (redondeos del MEF)

errores, avisos = [], []

def leer(path):
    raw = path.read_bytes()
    try:
        return raw.decode("utf-8")
    except UnicodeDecodeError:
        return raw.decode("latin-1")

def texto(celda):
    t = re.sub(r"<[^>]+>", " ", celda)
    return re.sub(r"\s+", " ", html.unescape(t)).strip()

def numero(s):
    s = s.replace(",", "").strip()
    if s in ("", "-"):
        return 0
    try:
        return int(round(float(s)))
    except ValueError:
        return 0

def leer_archivo(path, esperado_anio, etiqueta):
    """Devuelve ({cod: (nombre, devengado)}, [problemas])"""
    probs = []
    doc = leer(path)
    cab = texto(doc)
    m = re.search(r"Año de Ejecución:\s*(\d{4})", cab)
    if not m or int(m.group(1)) != esperado_anio:
        probs.append(f"el archivo es del año {m.group(1) if m else '?'}, no {esperado_anio}")
    if "Sólo Proyectos" not in cab and "Solo Proyectos" not in cab:
        probs.append("no está filtrado por 'Sólo Proyectos'")
    if etiqueta not in cab:
        probs.append(f"no corresponde al periodo esperado ('{etiqueta}')")
    filas = {}
    for tr in re.findall(r"<tr[^>]*>(.*?)</tr>", doc, flags=re.S | re.I):
        celdas = [texto(c) for c in re.findall(r"<td[^>]*>(.*?)</td>", tr, flags=re.S | re.I)]
        if len(celdas) < 9:
            continue
        mm = re.match(r"^(\d{6})-\d{6}:\s*(.+)$", celdas[0])
        if not mm:
            continue
        filas[mm.group(1)] = (mm.group(2), numero(celdas[IDX_DEVENGADO]))
    if len(filas) < MIN_MUNIS:
        probs.append(f"solo {len(filas)} municipalidades (¿exportaste agrupado por Municipalidad?)")
    return filas, probs

def procesar_anio(anio):
    carpeta = HIST / str(anio)
    acum, nombres = {}, {}
    for clave, etiqueta in PERIODOS:
        f = carpeta / f"{clave}_{anio}.xls"
        if not f.exists():
            avisos.append(f"{anio}: falta {f.name}; se omite el año completo")
            return None
        filas, probs = leer_archivo(f, anio, etiqueta)
        if probs:
            errores.append(f"{anio}/{f.name}: " + "; ".join(probs))
            return None
        for cod, (nom, dev) in filas.items():
            nombres[cod] = nom
            acum[cod] = acum.get(cod, 0) + dev
    return acum, nombres

def main():
    anios = sorted(int(p.name) for p in HIST.iterdir() if p.is_dir() and re.fullmatch(r"\d{4}", p.name))
    control = {}
    if CONTROL.exists():
        control = {int(k): v for k, v in json.loads(CONTROL.read_text(encoding="utf-8")).items()}

    # 1) leer todos los años; 2) completar con 0 las municipalidades que no gastaron
    #    nada en el periodo (el MEF no las incluye en el archivo) para que todos los
    #    años comparen exactamente las mismas municipalidades.
    leidos, maestro = {}, {}
    for a in anios:
        r = procesar_anio(a)
        if r is not None:
            leidos[a] = r
            maestro.update(r[1])

    ranking, muni, ok = {}, {}, []
    for a, (acum, _) in leidos.items():
        acum = {cod: acum.get(cod, 0) for cod in maestro}
        orden = sorted(acum.items(), key=lambda kv: (-kv[1], maestro[kv[0]]))
        lista, pos_prev, dev_prev = [], 0, None
        for i, (cod, dev) in enumerate(orden, 1):
            pos = pos_prev if dev == dev_prev else i      # empate -> misma posición
            pos_prev, dev_prev = pos, dev
            lista.append({"pos": pos, "cod": cod, "nombre": maestro[cod], "dev": dev})
        m = next((x for x in lista if x["cod"] == COD_MPCH), None)
        if m is None:
            errores.append(f"{a}: la MPCH ({COD_MPCH}) no aparece en ningún archivo")
            continue
        ranking[str(a)] = lista
        muni[str(a)] = {"dev": m["dev"], "pos": m["pos"], "total": sum(x["dev"] for x in lista),
                       "municipalidades": len(lista)}
        ok.append(a)
        if a in control and abs(control[a] - m["dev"]) > TOL_CONTROL:
            avisos.append(f"{a}: MPCH suma {m['dev']:,} pero historico_egresos.json dice {control[a]:,} "
                          f"(dif. {m['dev']-control[a]:+,}). Revisa los archivos de ese año.")

    salida = {
        "version": 1,
        "generado": datetime.datetime.now().isoformat(timespec="seconds"),
        "corte": "Ene–Oct",
        "criterio": "Devengado acumulado enero–octubre (1T+2T+3T+octubre) · Sólo Proyectos · Dpto. Lambayeque",
        "años": ok,
        "ranking": ranking,
        "muni": muni,
    }
    SALIDA.write_text(json.dumps(salida, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")

    print(f"Años generados: {ok}")
    for a in ok:
        m = muni[str(a)]
        print(f"  {a}: MPCH puesto {m['pos']:>2}/{m['municipalidades']}  devengado S/ {m['dev']:>12,}   total dpto S/ {m['total']:>13,}")
    for t in avisos:  print("AVISO :", t)
    for t in errores: print("ERROR :", t)
    print(f"-> {SALIDA}")
    return 1 if errores else 0

if __name__ == "__main__":
    sys.exit(main())
