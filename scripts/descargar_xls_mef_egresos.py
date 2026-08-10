"""
╔══════════════════════════════════════════════════════════════════╗
║  descargar_xls_mef_egresos.py                                    ║
║  Dashboard MPCH Egresos — Municipalidad Provincial de Chiclayo   ║
║  Autor: Juan David Reyes Quintana — ORPMI / GORE Lambayeque      ║
║  Versión: 2.0 — Reconstruido desde cero con Playwright codegen   ║
╠══════════════════════════════════════════════════════════════════╣
║  QUÉ HACE ESTE SCRIPT                                            ║
║  1. Abre el navegador (visible, uno nuevo por cada archivo)      ║
║  2. Navega a MEF Consulta Amigable de Gastos                     ║
║  3. Descarga los 6 archivos XLS necesarios:                      ║
║     rubro.xls · categoria.xls · proyecto.xls                     ║
║     ranking.xls · fuente.xls · funcion.xls                       ║
║  4. Los copia a la carpeta xls/ del repositorio local            ║
║  5. Hace git add + commit + push (si no se pasa --nogh)          ║
║                                                                  ║
║  REQUISITOS                                                       ║
║  pip install playwright                                           ║
║  playwright install chromium                                      ║
║                                                                  ║
║  USO                                                             ║
║  python descargar_xls_mef_egresos.py          ← ventana visible  ║
║  python descargar_xls_mef_egresos.py --nogh   ← sin push GitHub  ║
║  python descargar_xls_mef_egresos.py --push   ← solo push de     ║
║                                                  lo que ya hay    ║
║                                                  en xls/          ║
╚══════════════════════════════════════════════════════════════════╝
"""

import sys
import shutil
import subprocess
import time
from datetime import datetime
from pathlib import Path

try:
    from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout
except ImportError:
    print("📦 Instalando Playwright...")
    subprocess.run([sys.executable, "-m", "pip", "install", "playwright"], check=True)
    subprocess.run([sys.executable, "-m", "playwright", "install", "chromium"], check=True)
    from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout

# ══════════════════════════════════════════════════════════════════
# CONFIGURACIÓN
# ══════════════════════════════════════════════════════════════════

REPO_DIR = Path(__file__).resolve().parent.parent
# REPO_DIR = Path(r"C:\Users\TU_USUARIO\Desktop\mpcix-egresos")  # ← fallback manual

XLS_DIR  = REPO_DIR / "xls"
TEMP_DIR = Path.home() / "Downloads" / "mpcix_egresos_temp"

ANIO = datetime.now().year
COMMIT_MSG_BASE = "data: actualizacion diaria XLS egresos"  # sin tildes

URL_INICIO = f"https://apps5.mineco.gob.pe/transparencia/Navegador/default.aspx?y={ANIO}&ap=Proyecto"

# Identificador exacto de Chiclayo en la tabla de municipalidades del portal
MUNICIPALIDAD_CELL = "140101-301212: MUNICIPALIDAD"

# Cada archivo = el botón de dimensión que se pulsa antes de "Exportar".
# "ranking.xls" no selecciona una municipalidad específica: exporta
# agrupado por Municipalidad para todo el departamento.
ARCHIVOS = [
    {"nombre": "rubro.xls",     "boton_dimension": "Rubro",                  "descripcion": "Ejecución por Rubro de Financiamiento"},
    {"nombre": "categoria.xls", "boton_dimension": "Categoría Presupuestal", "descripcion": "Ejecución por Categoría Presupuestal"},
    {"nombre": "proyecto.xls",  "boton_dimension": "Producto/Proyecto",      "descripcion": "Detalle de Proyectos de Inversión"},
    {"nombre": "fuente.xls",    "boton_dimension": "Fuente",                 "descripcion": "Ejecución por Fuente de Financiamiento"},
    {"nombre": "funcion.xls",   "boton_dimension": "Función",                "descripcion": "Ejecución por Función del Gasto"},
    {"nombre": "ranking.xls",   "boton_dimension": None,                     "descripcion": "Ranking de Municipalidades — Dpto. Lambayeque"},
]


def log(msg, tipo="INFO"):
    hora = datetime.now().strftime("%H:%M:%S")
    icono = {"INFO": "ℹ", "OK": "✅", "ERROR": "❌", "WARN": "⚠️", "GIT": "📤"}.get(tipo, "•")
    print(f"[{hora}] {icono}  {msg}")


def limpiar_temp():
    if TEMP_DIR.exists():
        shutil.rmtree(TEMP_DIR)
    TEMP_DIR.mkdir(parents=True, exist_ok=True)
    log(f"Carpeta temporal lista: {TEMP_DIR}")


def verificar_repo():
    if not REPO_DIR.exists():
        log(f"No se encontró el repositorio en: {REPO_DIR}", "ERROR")
        sys.exit(1)
    if not (REPO_DIR / ".git").exists():
        log(f"La carpeta {REPO_DIR} no es un repositorio git.", "ERROR")
        sys.exit(1)
    XLS_DIR.mkdir(parents=True, exist_ok=True)
    log(f"Repositorio OK: {REPO_DIR}", "OK")


def git_push(archivos_copiados):
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
            if "nothing to commit" in result.stdout or "nothing to commit" in result.stderr:
                log("No hay cambios nuevos que commitear (archivos idénticos).", "WARN")
                return
            log(f"Error git: {result.stderr.strip()}", "ERROR")
            sys.exit(1)
        if result.stdout.strip():
            print(f"         {result.stdout.strip()}")
    log(f"Push exitoso — {len(archivos_copiados)} archivos actualizados", "OK")


# ══════════════════════════════════════════════════════════════════
# DESCARGA — pasos exactos grabados con Playwright codegen
# ══════════════════════════════════════════════════════════════════

def descargar_un_archivo(playwright, nombre, boton_dimension):
    """
    Abre un navegador nuevo, ejecuta el flujo grabado con codegen y
    devuelve la ruta del archivo descargado, o None si falló.
    Un navegador por archivo: si Chromium crashea en uno, no afecta
    la descarga de los demás.
    """
    browser = playwright.chromium.launch(headless=False)
    context = browser.new_context(accept_downloads=True)
    page = context.new_page()

    try:
        page.goto(URL_INICIO)
        f = page.locator("#frame0").content_frame

        f.get_by_role("cell", name="TOTAL", exact=True).click()
        page.wait_for_timeout(2500)

        f.get_by_role("button", name="Nivel de Gobierno").click()
        page.wait_for_timeout(1500)
        f.get_by_role("cell", name="M: GOBIERNOS LOCALES").click()
        page.wait_for_timeout(2500)

        f.get_by_role("button", name="Gob.Loc./Mancom.").click()
        page.wait_for_timeout(1500)
        f.get_by_role("cell", name="M: MUNICIPALIDADES").click()
        page.wait_for_timeout(2500)

        f.locator("#ctl00_CPH1_BtnDepartamento").click()
        page.wait_for_timeout(1500)
        f.get_by_role("cell", name=": LAMBAYEQUE").click()
        page.wait_for_timeout(2500)

        f.get_by_role("button", name="Municipalidad").click()
        page.wait_for_timeout(1500)

        if boton_dimension is not None:
            # Los 5 archivos "MPCH": filtran a la municipalidad exacta
            # y luego agrupan por la dimensión (Rubro, Fuente, etc.)
            f.get_by_role("cell", name=MUNICIPALIDAD_CELL).click()
            page.wait_for_timeout(2500)
            f.get_by_role("button", name=boton_dimension).click()
            # Este es el clic más importante para el resultado final:
            # el postback que cambia la agrupación de la tabla. Le damos
            # más margen porque es el que estaba fallando (se exportaba
            # el estado anterior si no se esperaba lo suficiente).
            page.wait_for_timeout(3500)
        else:
            # ranking.xls: se exporta agrupado por "Municipalidad" para
            # todo el departamento — el postback del botón de arriba
            # también necesita su tiempo antes de exportar.
            page.wait_for_timeout(2000)
        # ranking.xls: no se filtra municipalidad — se exporta agrupado
        # por "Municipalidad" para todo el departamento (ya se hizo click
        # en el botón "Municipalidad" arriba, así que aquí no hace falta nada más)

        with page.expect_download(timeout=120_000) as download_info:
            f.get_by_role("link", name="Exportar").click()
        download = download_info.value

        ruta_temp = TEMP_DIR / nombre
        download.save_as(str(ruta_temp))
        kb = ruta_temp.stat().st_size // 1024
        log(f"{nombre} descargado OK ({kb} KB)", "OK")
        return ruta_temp

    except PWTimeout:
        log(f"Timeout esperando descarga de {nombre}. El MEF puede estar lento.", "ERROR")
        try:
            page.screenshot(path=str(TEMP_DIR / f"timeout_{nombre}.png"))
        except Exception:
            pass
        return None

    except Exception as e:
        log(f"Error descargando {nombre}: {e}", "ERROR")
        try:
            if not page.is_closed():
                page.screenshot(path=str(TEMP_DIR / f"err_{nombre}.png"))
        except Exception:
            pass
        return None

    finally:
        try:
            context.close()
            browser.close()
        except Exception:
            pass


def descargar_archivos():
    descargados = []
    with sync_playwright() as playwright:
        for archivo in ARCHIVOS:
            nombre = archivo["nombre"]
            log(f"[{nombre}] {archivo['descripcion']}...")
            ruta = descargar_un_archivo(playwright, nombre, archivo["boton_dimension"])
            if ruta:
                descargados.append((nombre, ruta))
            time.sleep(3)  # pausa entre archivos para no saturar el portal
    return descargados


def solo_git_push():
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
    args = sys.argv[1:]
    hacer_push = "--nogh" not in args
    solo_push  = "--push" in args

    print()
    print("══════════════════════════════════════════════")
    print("  MPCH EGRESOS — Actualización diaria de datos")
    print(f"  {datetime.now().strftime('%A %d de %B de %Y — %H:%M')}")
    print("══════════════════════════════════════════════")
    print()

    verificar_repo()

    if solo_push:
        log("Modo --push: solo haciendo commit+push de xls/ existente")
        solo_git_push()
        return

    limpiar_temp()
    log(f"Año consultado: {ANIO} | Solo Proyectos")
    print()

    descargados = descargar_archivos()
    print()

    total_esperado = len(ARCHIVOS)
    nombres_descargados = [d[0] for d in descargados]
    faltantes = [a["nombre"] for a in ARCHIVOS if a["nombre"] not in nombres_descargados]

    if not descargados or faltantes:
        print("══════════════════════════════════════════════")
        log(f"Descargados: {len(descargados)}/{total_esperado}", "WARN")
        if faltantes:
            log(f"Faltantes:  {', '.join(faltantes)}", "ERROR")
        log("Push cancelado — se necesitan los 6 archivos completos.", "ERROR")
        log("Revisa los screenshots err_*.png / timeout_*.png en la carpeta temporal.", "ERROR")
        log(f"  {TEMP_DIR}", "ERROR")
        print("══════════════════════════════════════════════")
        sys.exit(1)

    log("6/6 archivos descargados correctamente.", "OK")
    print()
    log("Copiando archivos al repositorio...")
    archivos_copiados = []
    for nombre, ruta_temp in descargados:
        destino = XLS_DIR / nombre
        shutil.copy2(ruta_temp, destino)
        log(f"  {nombre} → xls/{nombre}", "OK")
        archivos_copiados.append(destino)

    print()
    if hacer_push:
        log("6/6 archivos OK — subiendo a GitHub...", "GIT")
        git_push(archivos_copiados)
    else:
        log("Push omitido por --nogh. Para subir manualmente:", "WARN")
        log(f"  cd {REPO_DIR} && git add xls/ && git commit -m 'data: update' && git push")

    print()
    print("══════════════════════════════════════════════")
    log(f"Proceso completado — {len(archivos_copiados)}/6 archivos actualizados", "OK")
    log("Dashboard: https://jdrq.github.io/mpcix-egresos/", "OK")
    print("══════════════════════════════════════════════")
    print()


if __name__ == "__main__":
    main()
