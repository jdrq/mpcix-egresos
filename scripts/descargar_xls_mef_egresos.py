"""
╔══════════════════════════════════════════════════════════════════╗
║  descargar_xls_mef_egresos.py                                    ║
║  Dashboard MPCH Egresos — Municipalidad Provincial de Chiclayo  ║
║  Autor: Juan David Reyes Quintana — ORPMI / GORE Lambayeque      ║
║  Versión: 1.0 — Julio 2026                                       ║
╠══════════════════════════════════════════════════════════════════╣
║  QUÉ HACE ESTE SCRIPT                                            ║
║  1. Abre el navegador (invisible)                                 ║
║  2. Navega a MEF Consulta Amigable de Gastos                     ║
║  3. Descarga los 6 archivos XLS necesarios:                      ║
║     rubro.xls · categoria.xls · proyecto.xls                     ║
║     ranking.xls · fuente.xls · funcion.xls                       ║
║  4. Los copia a la carpeta xls/ del repositorio local            ║
║  5. Hace git add + commit + push automáticamente                  ║
║                                                                  ║
║  REQUISITOS                                                       ║
║  pip install playwright                                           ║
║  playwright install chromium                                      ║
║                                                                  ║
║  USO                                                             ║
║  python descargar_xls_mef_egresos.py          ← modo automático  ║
║  python descargar_xls_mef_egresos.py          ← ventana visible  ║
║  python descargar_xls_mef_egresos.py --nogh   ← sin push GitHub  ║
╚══════════════════════════════════════════════════════════════════╝
"""

import sys
import os
import shutil
import subprocess
import time
from datetime import datetime
from pathlib import Path

# ── Instalación automática de Playwright si no está ──────────────
try:
    from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout
except ImportError:
    print("📦 Instalando Playwright...")
    subprocess.run([sys.executable, "-m", "pip", "install", "playwright"], check=True)
    subprocess.run([sys.executable, "-m", "playwright", "install", "chromium"], check=True)
    from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout

# ══════════════════════════════════════════════════════════════════
# CONFIGURACIÓN — Ajusta solo esta sección
# ══════════════════════════════════════════════════════════════════

# Ruta al repositorio: se detecta automáticamente como la carpeta
# padre de la carpeta donde vive este script (scripts/ → raíz del repo)
# Si falla, descomenta la línea REPO_DIR manual y ajusta la ruta.
REPO_DIR = Path(__file__).resolve().parent.parent
# REPO_DIR = Path(r"C:\Users\TU_USUARIO\Desktop\mlp-egresos")  # ← fallback manual

# Carpeta destino de los XLS dentro del repo
XLS_DIR = REPO_DIR / "xls"

# Carpeta temporal de descarga
TEMP_DIR = Path.home() / "Downloads" / "mpcix_egresos_temp"

# Año de ejecución a consultar
ANIO = datetime.now().year

# Mensaje del commit (la fecha se añade automáticamente)
COMMIT_MSG_BASE = "data: actualizacion diaria XLS egresos"  # Sin tildes — evita encoding corrupto en Windows

# URL base MEF Consulta Amigable de Gastos
URL_BASE = "https://apps5.mineco.gob.pe/transparencia/Navegador/default.aspx"

# Código de la Municipalidad Provincial de Chiclayo
# Pliego 140101 — Ejecutora 301212
CODIGO_PLIEGO    = "140101"
CODIGO_EJECUTORA = "301212"

# ══════════════════════════════════════════════════════════════════
# PARÁMETROS DE CONSULTA MEF
# Cada archivo = una dimensión de agrupación diferente
# ══════════════════════════════════════════════════════════════════

# URL params comunes a todos los archivos
# q = tipo de consulta (acumulado)
# Año, Pliego y Ejecutora se pasan por el formulario
ARCHIVOS = [
    {
        "nombre":       "rubro.xls",
        "dimension_btn": "Rubro",           # texto exacto del botón en el portal MEF
        "descripcion":  "Ejecución por Rubro de Financiamiento"
    },
    {
        "nombre":       "categoria.xls",
        "dimension_btn": "Categoría Presupuestal",
        "descripcion":  "Ejecución por Categoría Presupuestal"
    },
    {
        "nombre":        "proyecto.xls",
        "dimension_btn": "Producto/Proyecto",  # botón real en el portal MEF
        "descripcion":   "Detalle de Proyectos de Inversión"
        # Flujo idéntico al estándar: aplicar_filtros_mpl() + click dimensión + Exportar
    },
    {
        "nombre":       "fuente.xls",
        "dimension_btn": "Fuente"                  ,
        "descripcion":  "Ejecución por Fuente de Financiamiento"
    },
    {
        "nombre":       "funcion.xls",
        "dimension_btn": "Función",
        "descripcion":  "Ejecución por Función del Gasto"
    },
    {
        "nombre":      "ranking.xls",
        "descripcion": "Ranking de Municipalidades — Dpto. Lambayeque",
        "es_ranking":  True   # Flujo diferente: sin filtro de municipalidad específica
    },
]

# ══════════════════════════════════════════════════════════════════
# FUNCIONES AUXILIARES
# ══════════════════════════════════════════════════════════════════

def log(msg, tipo="INFO"):
    """Imprime un mensaje con timestamp y tipo."""
    hora = datetime.now().strftime("%H:%M:%S")
    iconos = {"INFO": "ℹ", "OK": "✅", "ERROR": "❌", "WARN": "⚠️", "GIT": "📤"}
    icono = iconos.get(tipo, "•")
    print(f"[{hora}] {icono}  {msg}")


def limpiar_temp():
    """Limpia la carpeta temporal de descargas."""
    if TEMP_DIR.exists():
        shutil.rmtree(TEMP_DIR)
    TEMP_DIR.mkdir(parents=True, exist_ok=True)
    log(f"Carpeta temporal lista: {TEMP_DIR}")


def verificar_repo():
    """Verifica que el repositorio local existe y tiene git."""
    if not REPO_DIR.exists():
        log(f"No se encontró el repositorio en: {REPO_DIR}", "ERROR")
        log("Ajusta la variable REPO_DIR en la configuración del script.", "ERROR")
        sys.exit(1)
    if not (REPO_DIR / ".git").exists():
        log(f"La carpeta {REPO_DIR} no es un repositorio git.", "ERROR")
        sys.exit(1)
    XLS_DIR.mkdir(parents=True, exist_ok=True)
    log(f"Repositorio OK: {REPO_DIR}", "OK")


def git_push(archivos_copiados):
    """Hace git add, commit y push al repositorio."""
    fecha_str = datetime.now().strftime("%d-%m-%Y")
    commit_msg = f"{COMMIT_MSG_BASE} — {fecha_str}"

    comandos = [
        ["git", "-C", str(REPO_DIR), "add", "xls/"],
        ["git", "-C", str(REPO_DIR), "commit", "-m", commit_msg],
        ["git", "-C", str(REPO_DIR), "push"],
    ]

    for cmd in comandos:
        log(f"Ejecutando: {' '.join(cmd)}", "GIT")
        result = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8", errors="replace")
        if result.returncode != 0:
            # "nothing to commit" no es error real
            if "nothing to commit" in result.stdout or "nothing to commit" in result.stderr:
                log("No hay cambios nuevos que commitear (archivos idénticos).", "WARN")
                return
            log(f"Error git: {result.stderr.strip()}", "ERROR")
            sys.exit(1)
        if result.stdout.strip():
            print(f"         {result.stdout.strip()}")

    log(f"Push exitoso — {len(archivos_copiados)} archivos actualizados", "OK")


# ══════════════════════════════════════════════════════════════════
# DESCARGA CON PLAYWRIGHT
# ══════════════════════════════════════════════════════════════════

def frame(page):
    """Devuelve el iframe principal del portal MEF."""
    return page.locator("#frame0").content_frame


def aplicar_filtros_mpl(page):
    """
    Aplica los filtros comunes a los 5 archivos de MPL:
    Solo Proyectos → TOTAL → Gobiernos Locales → Municipalidades
    → Lambayeque → Municipalidad Provincial de Chiclayo (140301-301238)

    Retorna True si todo fue OK, False si hubo error.
    """
    f = frame(page)
    try:
        # 1. Seleccionar "Proyecto" (Solo Proyectos) en el dropdown
        f.locator("select[name='ctl00$CPH1$DrpActProy']").select_option("Proyecto")
        time.sleep(2)

        # 2. Click en TOTAL para arrancar la navegación de filtros
        f.get_by_role("cell", name="TOTAL", exact=True).click()
        time.sleep(3)

        # 3. Nivel de Gobierno → M: GOBIERNOS LOCALES
        f.get_by_role("button", name="Nivel de Gobierno").click()
        f.get_by_role("cell", name="M: GOBIERNOS LOCALES").click()
        time.sleep(3)

        # 4. Gob.Loc./Mancom. → M: MUNICIPALIDADES
        f.get_by_role("button", name="Gob.Loc./Mancom.").click()
        f.get_by_role("cell", name="M: MUNICIPALIDADES").click()
        time.sleep(3)

        # 5. Departamento → LAMBAYEQUE
        f.locator("#ctl00_CPH1_BtnDepartamento").click()
        f.get_by_role("cell", name=": LAMBAYEQUE").click()
        time.sleep(3)

        # 6. Municipalidad → 140101-301212: MUNICIPALIDAD PROVINCIAL DE CHICLAYO
        f.get_by_role("button", name="Municipalidad").click()
        f.get_by_role("cell", name="140101-301212: MUNICIPALIDAD").click()
        time.sleep(3)

        return True

    except Exception as e:
        log(f"Error aplicando filtros MPCH: {e}", "ERROR")
        return False


def aplicar_filtros_ranking(page):
    """
    Aplica los filtros para el ranking departamental:
    Solo Proyectos → TOTAL → Gobiernos Locales → Municipalidades
    → Lambayeque → (sin filtro de municipalidad, exporta agrupado por Municipalidad)

    Retorna True si todo fue OK, False si hubo error.
    """
    f = frame(page)
    try:
        # 1. Seleccionar "Proyecto" (Solo Proyectos)
        f.locator("select[name='ctl00$CPH1$DrpActProy']").select_option("Proyecto")
        time.sleep(2)

        # 2. Click en TOTAL
        f.get_by_role("cell", name="TOTAL", exact=True).click()
        time.sleep(3)

        # 3. Nivel de Gobierno → M: GOBIERNOS LOCALES
        f.get_by_role("button", name="Nivel de Gobierno").click()
        f.get_by_role("cell", name="M: GOBIERNOS LOCALES").click()
        time.sleep(3)

        # 4. Gob.Loc./Mancom. → M: MUNICIPALIDADES
        f.get_by_role("button", name="Gob.Loc./Mancom.").click()
        f.get_by_role("cell", name="M: MUNICIPALIDADES").click()
        time.sleep(3)

        # 5. Departamento → LAMBAYEQUE
        f.locator("#ctl00_CPH1_BtnDepartamento").click()
        f.get_by_role("cell", name=": LAMBAYEQUE").click()
        time.sleep(3)

        # 6. Click en "Municipalidad" como DIMENSIÓN de agrupación
        #    (no seleccionamos una municipalidad específica — exportamos todas)
        f.get_by_role("button", name="Municipalidad").click()
        time.sleep(3)

        return True

    except Exception as e:
        log(f"Error aplicando filtros ranking: {e}", "ERROR")
        return False


def click_dimension_y_exportar(page, nombre_dimension):
    """
    Hace click en el botón de dimensión (Rubro, Fuente, etc.)
    y luego captura la descarga del link Exportar.

    Intenta primero con el texto exacto; si falla, intenta con
    coincidencia parcial (útil cuando el portal abrevia el nombre).
    Retorna el objeto download o None si falla.
    """
    f = frame(page)
    try:
        # Intento 1: texto exacto
        btn = f.get_by_role("button", name=nombre_dimension, exact=True)
        if btn.count() == 0:
            # Intento 2: coincidencia parcial (primera palabra)
            primera_palabra = nombre_dimension.split()[0]
            log(f"Botón '{nombre_dimension}' no encontrado, probando '{primera_palabra}'...", "WARN")
            btn = f.get_by_role("button", name=primera_palabra)

        btn.first.click()
        time.sleep(3)

        with page.expect_download(timeout=120_000) as dl_info:
            f.get_by_role("link", name="Exportar").click()

        return dl_info.value

    except PWTimeout:
        log(f"Timeout esperando descarga para dimensión '{nombre_dimension}'", "ERROR")
        return None
    except Exception as e:
        log(f"Error en click_dimension_y_exportar '{nombre_dimension}': {e}", "ERROR")
        return None


def exportar_directo(page):
    """
    Para ranking: después de los filtros ya se muestra la tabla.
    Solo hay que capturar la descarga del link Exportar.
    """
    f = frame(page)
    try:
        with page.expect_download(timeout=120_000) as dl_info:
            f.get_by_role("link", name="Exportar").click()
        return dl_info.value
    except PWTimeout:
        log("Timeout esperando descarga del ranking", "ERROR")
        return None
    except Exception as e:
        log(f"Error exportando ranking: {e}", "ERROR")
        return None


def descargar_archivos(headless=True):
    """
    Navega al portal MEF Consulta Amigable de Gastos y descarga
    los 6 archivos XLS usando clics reales (grabados con Playwright codegen).

    Estrategia:
    - Para los 5 archivos MPCH: aplica filtros comunes → click dimensión → Exportar
    - Para ranking: aplica filtros hasta Lambayeque → Exportar (agrupa por Municipalidad)

    Retorna lista de tuplas (nombre_archivo, ruta_temporal).
    """
    descargados = []

    URL_INICIO = f"{URL_BASE}?y={ANIO}&ap=Proyecto"

    with sync_playwright() as p:
        log("Iniciando navegador Chromium...")
        browser = p.chromium.launch(
            headless=headless,
            downloads_path=str(TEMP_DIR)
        )
        context = browser.new_context(
            accept_downloads=True,
            viewport={"width": 1366, "height": 768}
        )
        page = context.new_page()
        page.set_default_timeout(120_000)

        for archivo in ARCHIVOS:
            nombre      = archivo["nombre"]
            dimension   = archivo.get("dimension_btn")  # texto del botón en el portal
            descripcion = archivo["descripcion"]
            es_ranking  = archivo.get("es_ranking", False)

            log(f"[{nombre}] {descripcion}...")

            try:
                # Cada archivo empieza desde la URL inicial (portal fresco)
                page.goto(URL_INICIO, wait_until="networkidle", timeout=120_000)
                time.sleep(4)

                if es_ranking:
                    # Flujo ranking: filtros hasta Lambayeque → Exportar directo
                    ok = aplicar_filtros_ranking(page)
                    if not ok:
                        log(f"Saltando {nombre} por error en filtros", "WARN")
                        page.screenshot(path=str(TEMP_DIR / f"error_{nombre}.png"))
                        continue
                    download = exportar_directo(page)

                else:
                    # Flujo MPCH estándar (rubro, categoria, proyecto, fuente, funcion):
                    # aplicar_filtros_mpl() → click botón dimensión → Exportar
                    ok = aplicar_filtros_mpl(page)
                    if not ok:
                        log(f"Saltando {nombre} por error en filtros", "WARN")
                        page.screenshot(path=str(TEMP_DIR / f"error_{nombre}.png"))
                        continue
                    download = click_dimension_y_exportar(page, dimension)

                if download is None:
                    log(f"No se obtuvo descarga para {nombre}", "ERROR")
                    page.screenshot(path=str(TEMP_DIR / f"sin_descarga_{nombre}.png"))
                    continue

                # Guardar el archivo descargado con el nombre correcto
                ruta_temp = TEMP_DIR / nombre
                download.save_as(str(ruta_temp))
                kb = ruta_temp.stat().st_size // 1024
                log(f"{nombre} descargado OK ({kb} KB)", "OK")
                descargados.append((nombre, ruta_temp))

            except PWTimeout:
                log(f"Timeout general en {nombre}. El MEF puede estar lento.", "ERROR")
                page.screenshot(path=str(TEMP_DIR / f"timeout_{nombre}.png"))

            except Exception as e:
                log(f"Error inesperado en {nombre}: {e}", "ERROR")
                page.screenshot(path=str(TEMP_DIR / f"err_{nombre}.png"))

            time.sleep(4)  # Pausa entre archivos para no saturar el portal

        browser.close()

    return descargados


# ══════════════════════════════════════════════════════════════════
# FALLBACK: INSTRUCCIONES MANUALES SI EL PORTAL CAMBIA
# ══════════════════════════════════════════════════════════════════

def instrucciones_manuales():
    """
    Muestra instrucciones para descarga manual si el portal MEF no responde.
    Se llama si hay 0 archivos descargados.
    """
    print("""
╔══════════════════════════════════════════════════════════════╗
║  DESCARGA MANUAL — Pasos en Consulta Amigable MEF            ║
╠══════════════════════════════════════════════════════════════╣
║                                                              ║
║  URL: https://apps5.mineco.gob.pe/transparencia/             ║
║       Navegador/default.aspx                                 ║
║                                                              ║
║  Para CADA archivo, configura así:                           ║
║                                                              ║
║  ┌─────────────┬──────────────────────────────────────────┐  ║
║  │ Archivo     │ Agrupación a seleccionar                 │  ║
║  ├─────────────┼──────────────────────────────────────────┤  ║
║  │ rubro.xls   │ Rubro                                    │  ║
║  │ categoria.xls│ Categoría Presupuestal                  │  ║
║  │ proyecto.xls│ Proyecto (detalle)                       │  ║
║  │ ranking.xls │ Municipalidad (todo el departamento)     │  ║
║  │ fuente.xls  │ Fuente de Financiamiento                 │  ║
║  │ funcion.xls │ Función                                  │  ║
║  └─────────────┴──────────────────────────────────────────┘  ║
║                                                              ║
║  Filtros para rubro/categoria/proyecto/fuente/funcion:       ║
║  • Año: 2026                                                 ║
║  • Incluye: Solo Proyectos                                   ║
║  • Pliego: 140101 — MPCH                                     ║
║  • Ejecutora: 301212                                         ║
║                                                              ║
║  Filtros para ranking.xls:                                   ║
║  • Año: 2026                                                 ║
║  • Incluye: Solo Proyectos                                   ║
║  • Departamento: Lambayeque (14)                             ║
║  • (sin filtro de pliego/ejecutora)                          ║
║                                                              ║
║  Luego coloca los archivos en:                               ║
""")
    print(f"║  {str(XLS_DIR):<60}  ║")
    print("""║                                                              ║
║  Y ejecuta el script con:                                    ║
║  python descargar_xls_mef_egresos.py --nogh                  ║
║  (para copiar sin descargar) o manualmente haz git push      ║
╚══════════════════════════════════════════════════════════════╝
""")


def solo_git_push():
    """
    Modo --nogh-manual: no descarga, solo hace git push
    con lo que ya hay en xls/. Útil para descarga manual.
    """
    archivos_xls = list(XLS_DIR.glob("*.xls"))
    if not archivos_xls:
        log("No hay archivos .xls en la carpeta xls/. Colócalos primero.", "ERROR")
        sys.exit(1)
    log(f"Encontrados {len(archivos_xls)} archivos en xls/ — haciendo push...", "GIT")
    git_push(archivos_xls)


# ══════════════════════════════════════════════════════════════════
# MAIN
# ══════════════════════════════════════════════════════════════════

def main():
    # Leer argumentos
    args = sys.argv[1:]
    headless = "--ver" in args           # por defecto ventana VISIBLE; --ver no hace nada (compatibilidad)
    hacer_push = "--nogh" not in args    # --nogh → no hace git push
    solo_push  = "--push" in args        # --push → solo hace push sin descargar

    print()
    print("══════════════════════════════════════════════")
    print("  MPCH EGRESOS — Actualización diaria de datos")
    print(f"  {datetime.now().strftime('%A %d de %B de %Y — %H:%M')}")
    print("══════════════════════════════════════════════")
    print()

    # Verificar que el repo existe
    verificar_repo()

    # Modo --push: solo commitear lo que ya hay en xls/
    if solo_push:
        log("Modo --push: solo haciendo commit+push de xls/ existente")
        solo_git_push()
        return

    # Limpiar carpeta temporal
    limpiar_temp()

    # Descargar archivos del MEF
    log(f"Modo: {'SILENCIOSO' if headless else 'VISIBLE'} | Push GitHub: {'SÍ' if hacer_push else 'NO'}")
    log(f"Año consultado: {ANIO} | Solo Proyectos")
    print()

    descargados = descargar_archivos(headless=headless)

    print()

    # ── Verificar que se descargaron los 6 archivos ─────────────────
    total_esperado = len(ARCHIVOS)
    nombres_descargados = [d[0] for d in descargados]
    faltantes = [a["nombre"] for a in ARCHIVOS if a["nombre"] not in nombres_descargados]

    if not descargados or faltantes:
        print()
        print("══════════════════════════════════════════════")
        log(f"Descargados: {len(descargados)}/{total_esperado}", "WARN")
        if faltantes:
            log(f"Faltantes:  {', '.join(faltantes)}", "ERROR")
        log("Push cancelado — se necesitan los 6 archivos completos.", "ERROR")
        log("Vuelve a ejecutar el script para reintentar.", "ERROR")
        print("══════════════════════════════════════════════")
        print()
        instrucciones_manuales()
        sys.exit(1)

    # ── 6/6 descargados — copiar al repo ─────────────────────────
    log(f"6/6 archivos descargados correctamente.", "OK")
    print()
    log("Copiando archivos al repositorio...")
    archivos_copiados = []
    for nombre, ruta_temp in descargados:
        destino = XLS_DIR / nombre
        shutil.copy2(ruta_temp, destino)
        log(f"  {nombre} → xls/{nombre}", "OK")
        archivos_copiados.append(destino)

    print()

    # ── Push a GitHub (automático si no se usó --nogh) ───────────
    if hacer_push:
        log("6/6 archivos OK — subiendo a GitHub...", "GIT")
        git_push(archivos_copiados)
    else:
        log("Push omitido por --nogh. Para subir manualmente:", "WARN")
        log(f"  cd {REPO_DIR} && git add xls/ && git commit -m 'data: update' && git push")

    # ── Resumen final ─────────────────────────────────────────────
    print()
    print("══════════════════════════════════════════════")
    log(f"Proceso completado — {len(archivos_copiados)}/6 archivos actualizados", "OK")
    log(f"Dashboard: https://jdrq.github.io/mpcix-egresos/", "OK")
    print("══════════════════════════════════════════════")
    print()


if __name__ == "__main__":
    main()
