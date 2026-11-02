#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
generar_ranking_historico.py
============================
Lee los .xls de MEF Consulta Amigable guardados en  historico/<año>/  y genera:
    data/ranking_historico.json   (Bloque de ranking: slide 2, y años 2015+ del slide 8)
    data/historico_egresos.json   (MPCH por año: slide 8)
Los dos JSON llevan el MES DE CORTE (campo "corte_mes"); index.html y app.js arman
sus textos ("Ene–Nov", "enero–noviembre"...) a partir de ese campo. Para pasar a otro
mes NO hay que tocar index.html ni app.js: solo cambiar CORTE_MES (o pasar --mes N),
bajar el archivo del mes nuevo a historico/<año>/ y correr este script.

Para cada año necesita (Sólo Proyectos, agrupado por Municipalidad):
    1T_<año>.xls, 2T_<año>.xls, 3T_<año>.xls   -> Trimestres I, II y III
    octubre_<año>.xls   -> Mes 10   (si CORTE_MES >= 10)
    noviembre_<año>.xls -> Mes 11   (si CORTE_MES >= 11)
    diciembre_<año>.xls -> Mes 12   (si CORTE_MES == 12)
Cada archivo trae SOLO el devengado de su periodo, por eso se SUMAN:
    Devengado Ene–<corte> = 1T + 2T + 3T + octubre [+ noviembre [+ diciembre]]
(PIA y PIM vienen vacíos en estos archivos; no se usan.)

PIM por año (opcional): si existe  anual_<año>.xls  (Ene–Dic, Sólo Proyectos, por
Municipalidad) se lee su PIA y PIM y se guardan junto a cada municipalidad. El dashboard
calcula  Avance % = Devengado acumulado / PIM anual.  El PIM es un valor "foto"
(stock): NUNCA se suma entre archivos. Si falta el anual de un año, esa columna
queda en "—" para ese año.

"Todo o nada": si cualquier archivo falla la validación (año, Sólo Proyectos, periodo,
nº de municipalidades) NO se escribe ningún JSON y el sitio sigue con el corte anterior.

Solo usa la librería estándar de Python (no requiere pandas).
Uso:   python scripts/generar_ranking_historico.py            (usa CORTE_MES)
       python scripts/generar_ranking_historico.py --mes 12   (otro corte, 10 a 12)
"""
import json, re, sys, html, datetime
from pathlib import Path

BASE = Path(__file__).resolve().parent.parent
HIST = BASE / "historico"
SALIDA = BASE / "data" / "ranking_historico.json"
CONTROL = BASE / "data" / "historico_egresos.json"   # MPCH por año: se LEE (corte anterior) y se REESCRIBE

# ── MES DE CORTE (lo único que se cambia cada mes) ───────────────────────────
CORTE_MES = 11               # 10 = octubre · 11 = noviembre · 12 = diciembre
# ─────────────────────────────────────────────────────────────────────────────
MESES = {10: ("octubre", "Oct", "Mes 10"),
         11: ("noviembre", "Nov", "Mes 11"),
         12: ("diciembre", "Dic", "Mes 12")}

def periodos(corte):
    """Archivos que se suman para llegar al corte (1T+2T+3T y los meses 10..corte)."""
    p = [("1T", "Trimestre I:"), ("2T", "Trimestre II:"), ("3T", "Trimestre III:")]
    for m in range(10, corte + 1):
        p.append((MESES[m][0], MESES[m][2]))
    return p

PERIODOS = periodos(CORTE_MES)
COD_MPCH = "140101"          # Municipalidad Provincial de Chiclayo
IDX_PIA = 1
IDX_PIM = 2
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

def leer_anual(path, anio):
    """Devuelve ({cod: (pim, devengado_anual, pia)}, [problemas]) de un anual_<año>.xls"""
    probs = []
    doc = leer(path)
    cab = texto(doc)
    m = re.search(r"Año de Ejecución:\s*(\d{4})", cab)
    if not m or int(m.group(1)) != anio:
        probs.append(f"el archivo es del año {m.group(1) if m else '?'}, no {anio}")
    if "Sólo Proyectos" not in cab and "Solo Proyectos" not in cab:
        probs.append("no está filtrado por 'Sólo Proyectos'")
    filas = {}
    for tr in re.findall(r"<tr[^>]*>(.*?)</tr>", doc, flags=re.S | re.I):
        celdas = [texto(c) for c in re.findall(r"<td[^>]*>(.*?)</td>", tr, flags=re.S | re.I)]
        if len(celdas) < 9:
            continue
        mm = re.match(r"^(\d{6})-\d{6}:\s*(.+)$", celdas[0])
        if mm:
            filas[mm.group(1)] = (numero(celdas[IDX_PIM]), numero(celdas[IDX_DEVENGADO]), numero(celdas[IDX_PIA]))
    if len(filas) < MIN_MUNIS:
        probs.append(f"solo {len(filas)} municipalidades (¿exportaste agrupado por Municipalidad?)")
    return filas, probs

def procesar_anio(anio):
    carpeta = HIST / str(anio)
    acum, nombres = {}, {}
    for clave, etiqueta in PERIODOS:
        f = carpeta / f"{clave}_{anio}.xls"
        if not f.exists():
            errores.append(f"{anio}: falta {f.name} (descárgalo en historico/{anio}/ y vuelve a correr)")
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
    global CORTE_MES, PERIODOS
    if "--mes" in sys.argv:
        try:
            CORTE_MES = int(sys.argv[sys.argv.index("--mes") + 1])
        except (IndexError, ValueError):
            print("ERROR : usa  --mes 10 | 11 | 12"); return 1
    if CORTE_MES not in MESES:
        print(f"ERROR : CORTE_MES={CORTE_MES} no soportado (10, 11 o 12)"); return 1
    PERIODOS = periodos(CORTE_MES)
    corte_corto = "Ene–" + MESES[CORTE_MES][1]            # p. ej. "Ene–Nov"
    corte_largo = "enero–" + MESES[CORTE_MES][0]          # p. ej. "enero–noviembre"
    print(f"Corte: {corte_corto}  (archivos: {', '.join(k for k, _ in PERIODOS)})")

    anios = sorted(int(p.name) for p in HIST.iterdir() if p.is_dir() and re.fullmatch(r"\d{4}", p.name))
    control = {}      # MPCH por año del corte ANTERIOR (lo que había en historico_egresos.json)
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

    ranking, muni, ok, pim_ok = {}, {}, [], []
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
        # PIM anual (opcional)
        fa = HIST / str(a) / f"anual_{a}.xls"
        if fa.exists():
            pims, probs = leer_anual(fa, a)
            if probs:
                errores.append(f"{a}/{fa.name}: " + "; ".join(probs))
                pims = {}
            else:
                pim_ok.append(a)
                for x in lista:
                    x["pim"] = pims.get(x["cod"], (0, 0, 0))[0]
                    x["pia"] = pims.get(x["cod"], (0, 0, 0))[2]
                    # control: lo devengado al corte no puede superar lo devengado Ene–Dic
                    dev_anual = pims.get(x["cod"], (0, 0, 0))[1]
                    if x["dev"] > dev_anual + TOL_CONTROL:
                        avisos.append(f"{a}: {x['nombre']} {corte_corto} S/ {x['dev']:,} > anual S/ {dev_anual:,}")
        else:
            avisos.append(f"{a}: sin anual_{a}.xls -> PIM y Avance % no disponibles ese año")
        ranking[str(a)] = lista
        muni[str(a)] = {"dev": m["dev"], "pos": m["pos"], "total": sum(x["dev"] for x in lista),
                       "municipalidades": len(lista)}
        ok.append(a)
        # control contra el corte anterior: un acumulado nunca baja al sumar meses
        if a in control and m["dev"] < control[a] - TOL_CONTROL:
            avisos.append(f"{a}: MPCH {corte_corto} S/ {m['dev']:,} es MENOR que el corte anterior "
                          f"S/ {control[a]:,} (dif. {m['dev']-control[a]:+,}). Revisa los archivos de ese año.")

    salida = {
        "version": 2,
        "generado": datetime.datetime.now().isoformat(timespec="seconds"),
        "corte": corte_corto,
        "corte_mes": CORTE_MES,
        "criterio": f"Devengado acumulado {corte_largo} ({'+'.join(k for k, _ in PERIODOS)}) · Sólo Proyectos · Dpto. Lambayeque",
        "años": ok,
        "años_con_pim": pim_ok,
        "ranking": ranking,
        "muni": muni,
    }

    print(f"Años generados: {ok}")
    for a in ok:
        m = muni[str(a)]
        prev = f"  (+{m['dev']-control[a]:,} vs corte anterior)" if a in control else ""
        print(f"  {a}: MPCH puesto {m['pos']:>2}/{m['municipalidades']}  devengado S/ {m['dev']:>12,}   total dpto S/ {m['total']:>13,}{prev}")
    for t in avisos:  print("AVISO :", t)
    for t in errores: print("ERROR :", t)

    if errores:      # todo o nada: no se toca ningún JSON
        print("NO SE ESCRIBIÓ NINGÚN ARCHIVO: corrige los ERROR de arriba y vuelve a correr.")
        return 1

    SALIDA.write_text(json.dumps(salida, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")
    # MPCH por año (slide 8): mismos números del ranking, un solo origen de verdad
    CONTROL.write_text(json.dumps({str(a): muni[str(a)]["dev"] for a in ok}, indent=2) + "\n", encoding="utf-8")
    print(f"-> {SALIDA}")
    print(f"-> {CONTROL}")
    return 0

if __name__ == "__main__":
    sys.exit(main())
