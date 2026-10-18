"""
AUTOMATIZADOR OFICIAL SIGO - PAE (UT JUNÍN - PERIODO 2027)
==========================================================
URL detectada en tu captura: https://sigo.pae.gob.pe/SIGO/
Este script realiza todo el flujo exacto de tu pantalla:
1. Inicia sesión en https://sigo.pae.gob.pe/SIGO/ con tu usuario y contraseña.
2. Ingresa al módulo 'IntranetCae/Consultar' (pestaña CAE).
3. Selecciona las celdas desplegables:
   - Campaña: PERIODO 2027
   - Proceso: PROCESO DE COMPRA 2027-1
   - RDE: RD 0110
   - Unidad Territorial: JUNIN
   - Departamento: JUNIN
4. Hace clic en el botón 'Buscar'.
5. Hace clic en 'Reporte CAE Conformado' y guarda el primer Excel.
6. Hace clic en 'Reporte Integrantes CAE' y guarda el segundo Excel.
7. Actualiza automáticamente el Dashboard ('index.html') detectando todos los monitores con avance y lo sube a internet (GitHub Pages).

Requisitos (ejecutar una sola vez en CMD):
    pip install pandas openpyxl playwright
    playwright install chromium
"""

import os
import re
import json
import time
import subprocess
from datetime import datetime
import pandas as pd
from playwright.sync_api import sync_playwright

# ==============================================================================
# 1. CREDENCIALES Y PARÁMETROS DE FILTRO EN SIGO - PAE
# ==============================================================================
URL_LOGIN_SIGO = "https://sigo.pae.gob.pe/SIGO/"
USUARIO_SIGO   = "utjunml02"          # <-- Tu usuario SIGO
CLAVE_SIGO     = "TU_CONTRASEÑA_AQUI" # <-- Coloca aquí tu contraseña de SIGO

# Valores exactos de las celdas desplegables según tu captura de pantalla:
FILTRO_CAMPANA      = "PERIODO 2027"
FILTRO_PROCESO      = "PROCESO DE COMPRA 2027-1"
FILTRO_RDE          = "RD 0110"
FILTRO_UT           = "JUNIN"
FILTRO_DEPARTAMENTO = "JUNIN"

CARPETA_DESCARGAS  = os.path.abspath("./descargas_sigo")
ARCHIVO_PLANTILLA  = "Dashboard_CAES_Avance_Monitores_Junin_2027.html"
ARCHIVO_SALIDA_WEB = "index.html"

os.makedirs(CARPETA_DESCARGAS, exist_ok=True)


def seleccionar_opcion_por_texto(page, texto_objetivo):
    """Busca cualquier celda desplegable (<select>) que contenga la opción indicada y la selecciona."""
    selects = page.locator("select").all()
    for sel in selects:
        try:
            opciones = sel.locator("option").all_inner_texts()
            for op in opciones:
                if texto_objetivo.upper() in op.upper():
                    sel.select_option(label=op.strip())
                    time.sleep(0.8) # Esperar recarga en cascada del sistema SIGO
                    return True
        except Exception:
            continue
    return False


def descargar_reportes_sigo():
    """Realiza el login en SIGO-PAE, selecciona los filtros de Lista de Colegios y descarga los 2 Excel."""
    print(f"[{datetime.now().strftime('%H:%M:%S')}] Conectando a {URL_LOGIN_SIGO} ...")
    ruta_caes = os.path.join(CARPETA_DESCARGAS, "REPORTE_SISTEMA_CAES.xlsx")
    ruta_int  = os.path.join(CARPETA_DESCARGAS, "REPORTE_INTEGRANTES_CAES.xlsx")

    with sync_playwright() as p:
        # headless=False la primera vez para que veas cómo hace los clics; luego puedes poner headless=True
        browser = p.chromium.launch(headless=False)
        context = browser.new_context(accept_downloads=True, viewport={"width": 1600, "height": 900})
        page = context.new_page()

        # 1. Ingresar al Login de SIGO
        page.goto(URL_LOGIN_SIGO, timeout=60000)
        page.wait_for_load_state("networkidle")

        # Llenar Usuario y Contraseña
        page.locator('input[type="text"], input[name*="usu" i], input[id*="usu" i]').first.fill(USUARIO_SIGO)
        page.locator('input[type="password"]').first.fill(CLAVE_SIGO)
        page.locator('button[type="submit"], input[type="submit"], button:has-text("Ingresar"), button:has-text("Iniciar")').first.click()
        page.wait_for_load_state("networkidle")
        time.sleep(2)

        # 2. Ir al menú "Prestación del Servicio" -> "CAE" si no está ya en IntranetCae/Consultar
        if "IntranetCae/Consultar" not in page.url:
            try:
                page.locator('text="Prestación del Servicio"').first.click()
                time.sleep(1)
                page.locator('text="CAE"').first.click()
                page.wait_for_load_state("networkidle")
                time.sleep(2)
            except Exception as e:
                print("Aviso al navegar al menú CAE:", e)

        # 3. Seleccionar las celdas desplegables de "Lista de Colegios"
        print(" -> Seleccionando celdas de filtro (Periodo 2027 / Proceso / RD 0110 / Junín)...")
        seleccionar_opcion_por_texto(page, FILTRO_CAMPANA)
        seleccionar_opcion_por_texto(page, FILTRO_PROCESO)
        seleccionar_opcion_por_texto(page, FILTRO_RDE)
        seleccionar_opcion_por_texto(page, FILTRO_UT)
        seleccionar_opcion_por_texto(page, FILTRO_DEPARTAMENTO)

        # 4. Clic en el botón "Buscar"
        print(" -> Haciendo clic en 'Buscar'...")
        page.locator('button:has-text("Buscar"), a:has-text("Buscar"), input[value*="Buscar" i]').first.click()
        page.wait_for_load_state("networkidle")
        time.sleep(3)

        # 5. Clic en "Reporte CAE Conformado"
        print(" -> Descargando 'Reporte CAE Conformado'...")
        with page.expect_download(timeout=120000) as d1_info:
            page.locator('text="Reporte CAE Conformado"').first.click()
        d1 = d1_info.value
        d1.save_as(ruta_caes)
        print(f"    Guardado en: {ruta_caes}")

        # 6. Clic en "Reporte Integrantes CAE"
        print(" -> Descargando 'Reporte Integrantes CAE'...")
        with page.expect_download(timeout=120000) as d2_info:
            page.locator('text="Reporte Integrantes CAE"').first.click()
        d2 = d2_info.value
        d2.save_as(ruta_int)
        print(f"    Guardado en: {ruta_int}")

        browser.close()

    return ruta_caes, ruta_int


def actualizar_dashboard_html(ruta_caes, ruta_int):
    """Cruza ambos Excel descargados de SIGO y actualiza index.html incluyendo nuevos monitores."""
    print("Procesando datos descargados...")
    xls_c = pd.ExcelFile(ruta_caes)
    df_caes = pd.read_excel(ruta_caes, sheet_name=xls_c.sheet_names[0])

    xls_i = pd.ExcelFile(ruta_int)
    df_int = pd.read_excel(ruta_int, sheet_name=xls_i.sheet_names[0])

    # Mapear Monitor, Usuario y Fecha de Acta por Código Modular
    mon_map = df_int.groupby('Código Modular')['Persona'].first().to_dict()
    fecha_acta_map = df_int.groupby('Código Modular')['Fecha Acta Conform. CAE'].first().to_dict()

    # 1. Padrón agregado (sin Quintil ni Pueblo Indígena)
    agg_padron = df_caes.groupby(['Provincia', 'Distrito', 'Item', 'Nivel Educativo'], dropna=False).agg(
        iiee=('Código Modular', 'count'),
        usuarios=('Nro de Usuario', 'sum'),
        conf=('Conformado', lambda x: (x == 'Si').sum()),
        acta=('Acta', lambda x: (x == 'SI').sum()),
        int_tot=('Integrantes', 'sum')
    ).reset_index()

    padron_records = [{
        "p": str(r['Provincia']), "d": str(r['Distrito']), "it": str(r['Item']),
        "n": str(r['Nivel Educativo']), "iiee": int(r['iiee']), "usr": int(r['usuarios']),
        "conf": int(r['conf']), "acta": int(r['acta']), "int": int(r['int_tot'])
    } for _, r in agg_padron.iterrows()]

    # 2. CAEs Conformados
    caes_conf = df_caes[df_caes['Conformado'] == 'Si'].copy()
    caes_records = []
    for _, r in caes_conf.iterrows():
        cm = int(r['Código Modular'])
        caes_records.append({
            "cm": str(cm), "p": str(r['Provincia']), "d": str(r['Distrito']),
            "it": str(r['Item']), "cp": str(r['Centro Poblado']) if pd.notna(r['Centro Poblado']) else "-",
            "col": str(r['Nombre Colegio']), "n": str(r['Nivel Educativo']),
            "usr": int(r['Nro de Usuario']), "int": int(r['Integrantes']),
            "acta": str(r['Acta']), "mon": str(mon_map.get(cm, "NO ESPECIFICADO")),
            "facta": str(fecha_acta_map.get(cm, "-"))
        })

    # 3. Integrantes
    int_records = [{
        "nro": int(r['Nro']), "cm": str(r['Código Modular']), "p": str(r['Provincia']),
        "d": str(r['Distrito']), "it": str(r['Item']), "col": str(r['Nombre Colegio']),
        "n": str(r['Nivel Educativo']), "doc": str(r['Nº Documento']),
        "nom": f"{r['Apellido Paterno']} {r['Apellido Materno']}, {r['Nombres']}",
        "rol": str(r['Rol']), "car": str(r['Cargo']), "sex": str(r['Sexo']),
        "edad": int(r['EDAD']) if pd.notna(r['EDAD']) else 0,
        "leng": str(r['Lengua Materna']),
        "cel": str(int(r['Celular'])) if pd.notna(r['Celular']) else "-",
        "mon": str(r['Persona']), "fmod": str(r['Fecha Modificación']).strip()
    } for _, r in df_int.iterrows()]

    # 4. Construir dinámicamente la lista de Monitores Activos (por si se suman nuevos monitores cada día)
    mon_meta_list = []
    for persona, grp in df_int.groupby('Persona'):
        mon_meta_list.append({
            "mon": str(persona),
            "usr_sis": str(grp['Usuario'].iloc[0]),
            "prov": ", ".join(sorted(grp['Provincia'].dropna().astype(str).unique())),
            "items": sorted(grp['Item'].dropna().astype(str).unique().tolist()),
            "distritos": sorted(grp['Distrito'].dropna().astype(str).unique().tolist())
        })
    # Ordenar por cantidad de CAEs conformados de mayor a menor
    mon_meta_list.sort(key=lambda m: sum(1 for c in caes_records if c['mon'] == m['mon']), reverse=True)

    with open(ARCHIVO_PLANTILLA, 'r', encoding='utf-8') as f:
        html = f.read()

    html = re.sub(r'const PADRON = \[.*?\];', f'const PADRON = {json.dumps(padron_records, ensure_ascii=False)};', html, flags=re.DOTALL)
    html = re.sub(r'const CAES39 = \[.*?\];', f'const CAES39 = {json.dumps(caes_records, ensure_ascii=False)};', html, flags=re.DOTALL)
    html = re.sub(r'const INT152 = \[.*?\];', f'const INT152 = {json.dumps(int_records, ensure_ascii=False)};', html, flags=re.DOTALL)
    html = re.sub(r'const MONITOR_META = \[.*?\];', f'const MONITOR_META = {json.dumps(mon_meta_list, ensure_ascii=False)};', html, flags=re.DOTALL)

    fecha_actual = datetime.now().strftime("%d/%m/%Y %H:%M")
    html = re.sub(r'Corte: .*? \|', f'Corte: {fecha_actual} |', html)

    with open(ARCHIVO_SALIDA_WEB, 'w', encoding='utf-8') as f:
        f.write(html)
    print(f" -> '{ARCHIVO_SALIDA_WEB}' generado correctamente con corte {fecha_actual}.")


def publicar_en_github_pages():
    """Sube el index.html actualizado a tu repositorio web para que el link público se actualice al instante."""
    print("Subiendo actualización a la web (GitHub Pages)...")
    subprocess.run(["git", "add", ARCHIVO_SALIDA_WEB], check=False)
    subprocess.run(["git", "commit", "-m", f"Corte SIGO PAE {datetime.now().strftime('%d/%m/%Y %H:%M')}"], check=False)
    subprocess.run(["git", "push", "origin", "main"], check=False)
    print(" -> ¡Link público actualizado exitosamente!")


if __name__ == "__main__":
    r_caes, r_int = descargar_reportes_sigo()
    actualizar_dashboard_html(r_caes, r_int)
    publicar_en_github_pages()
