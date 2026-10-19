"""
AUTOMATIZADOR OFICIAL SIGO - PAE (UT JUNÍN - PERIODO 2027)
==========================================================
Selección en cascada por ID con espera AJAX para evitar el error "Seleccione proceso".
"""

import os
import re
import json
import time
import subprocess
from datetime import datetime
import pandas as pd
from playwright.sync_api import sync_playwright

URL_LOGIN_SIGO = "https://sigo.pae.gob.pe/SIGO/"
URL_CAE_MODULO = "https://sigo.pae.gob.pe/SIGO/IntranetCae/Consultar"
USUARIO_SIGO   = "UTJUNML02"
CLAVE_SIGO     = "UTJUNML02"

CARPETA_DESCARGAS  = os.path.abspath("./descargas_sigo")
ARCHIVO_PLANTILLA  = "index.html"
ARCHIVO_SALIDA_WEB = "index.html"

os.makedirs(CARPETA_DESCARGAS, exist_ok=True)

def descargar_reportes_sigo():
    print(f"[{datetime.now().strftime('%H:%M:%S')}] Conectando a {URL_LOGIN_SIGO} ...")
    ruta_caes = os.path.join(CARPETA_DESCARGAS, "REPORTE_SISTEMA_CAES.xlsx")
    ruta_int  = os.path.join(CARPETA_DESCARGAS, "REPORTE_INTEGRANTES_CAES.xlsx")

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        context = browser.new_context(accept_downloads=True, viewport={"width": 1600, "height": 900})
        page = context.new_page()

        # 1. Login
        page.goto(URL_LOGIN_SIGO, timeout=60000)
        page.wait_for_load_state("networkidle")

        page.locator('input[type="text"], input[name*="usu" i], input[id*="usu" i]').first.fill(USUARIO_SIGO)
        page.locator('input[type="password"]').first.fill(CLAVE_SIGO)
        page.locator('button[type="submit"], input[type="submit"], button:has-text("Ingresar"), button:has-text("Iniciar")').first.click()
        page.wait_for_load_state("networkidle")
        time.sleep(3)

        # 2. Navegar directo al módulo CAE
        print(" -> Navegando al módulo CAE...")
        page.goto(URL_CAE_MODULO, timeout=60000)
        page.wait_for_load_state("networkidle")
        time.sleep(3)

        # 3. Seleccionar Campaña en #ddlCampanaBuscar
        print(" -> Seleccionando Campaña: PERIODO 2027...")
        page.locator("#ddlCampanaBuscar").select_option(label="PERIODO 2027")
        time.sleep(2) # Esperar a que cargue la lista de procesos en cascada

        # 4. Seleccionar Proceso en #ddlProcesoBuscarCae
        print(" -> Seleccionando Proceso de Compra...")
        # Seleccionar la primera opción disponible que no sea vacía
        proceso_select = page.locator("#ddlProcesoBuscarCae")
        options = proceso_select.locator("option").all_inner_texts()
        for op in options:
            if "2027" in op or "PROCESO" in op.upper():
                proceso_select.select_option(label=op.strip())
                break
        time.sleep(2) # Esperar carga de RDE / Rige

        # 5. Seleccionar RDE si estuviera disponible
        try:
            rde_select = page.locator("#ddlRDECaeBuscar")
            rde_opts = rde_select.locator("option").all_inner_texts()
            if len(rde_opts) > 1:
                rde_select.select_option(index=1) # Seleccionar la primera opción válida
                time.sleep(1)
        except Exception:
            pass

        # 6. Hacer clic en el botón Buscar (#btnBuscar)
        print(" -> Haciendo clic en el botón 'Buscar' (#btnBuscar)...")
        page.locator('#btnBuscar').click()
        page.wait_for_load_state("networkidle")
        time.sleep(6) # Esperar carga completa de los 3,846 elementos

        # 7. Descargar Reporte CAE Conformado (#btnExportarColegioCae)
        print(" -> Descargando 'Reporte CAE Conformado' (#btnExportarColegioCae)...")
        with page.expect_download(timeout=120000) as d1_info:
            page.locator('#btnExportarColegioCae').click()
        d1 = d1_info.value
        d1.save_as(ruta_caes)
        print("    ¡Reporte CAE Conformado descargado con éxito!")

        # 8. Descargar Reporte Integrantes CAE (#btnExportarIntegrantesCae)
        print(" -> Descargando 'Reporte Integrantes CAE' (#btnExportarIntegrantesCae)...")
        with page.expect_download(timeout=120000) as d2_info:
            page.locator('#btnExportarIntegrantesCae').click()
        d2 = d2_info.value
        d2.save_as(ruta_int)
        print("    ¡Reporte Integrantes CAE descargado con éxito!")

        browser.close()
    return ruta_caes, ruta_int

def actualizar_dashboard_html(ruta_caes, ruta_int):
    print("Procesando datos y actualizando index.html...")
    xls_c = pd.ExcelFile(ruta_caes)
    df_caes = pd.read_excel(ruta_caes, sheet_name=xls_c.sheet_names[0])

    xls_i = pd.ExcelFile(ruta_int)
    df_int = pd.read_excel(ruta_int, sheet_name=xls_i.sheet_names[0])

    mon_map = df_int.groupby('Código Modular')['Persona'].first().to_dict()
    fecha_acta_map = df_int.groupby('Código Modular')['Fecha Acta Conform. CAE'].first().to_dict()

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

    mon_meta_list = []
    for persona, grp in df_int.groupby('Persona'):
        mon_meta_list.append({
            "mon": str(persona),
            "usr_sis": str(grp['Usuario'].iloc[0]),
            "prov": ", ".join(sorted(grp['Provincia'].dropna().astype(str).unique())),
            "items": sorted(grp['Item'].dropna().astype(str).unique().tolist()),
            "distritos": sorted(grp['Distrito'].dropna().astype(str).unique().tolist())
        })
    mon_meta_list.sort(key=lambda m: sum(1 for c in caes_records if c['mon'] == m['mon']), reverse=True)

    with open(ARCHIVO_PLANTILLA, 'r', encoding='utf-8') as f:
        html = f.read()

    html = re.sub(r'const PADRON = \[.*?\];', f'const PADRON = {json.dumps(padron_records, ensure_ascii=False)};', html, flags=re.DOTALL)
    html = re.sub(r'const CAES39 = \[.*?\];', f'const CAES39 = {json.dumps(caes_records, ensure_ascii=False)};', html, flags=re.DOTALL)
    html = re.sub(r'const INT152 = \[.*?\];', f'const INT152 = {json.dumps(int_records, ensure_ascii=False)};', html, flags=re.DOTALL)
    html = re.sub(r'const MONITOR_META = \[.*?\];', f'const MONITOR_META = {json.dumps(mon_meta_list, ensure_ascii=False)};', html, flags=re.DOTALL)

    now = datetime.now()
    fecha_actual = now.strftime("%d/%m/%Y")
    hora_actual = now.strftime("%I:%M %p")
    etiqueta_act = f"Actualizado Hoy a las {hora_actual}"
    html = re.sub(r'Corte: .*? \|', f'Corte: {fecha_actual} |', html)
    # Reemplazar la hora en la etiqueta verde
    html = re.sub(r'Actualizado Hoy a las .*?<\/div>', f'Actualizado Hoy a las {hora_actual}</div>', html)

    with open(ARCHIVO_SALIDA_WEB, 'w', encoding='utf-8') as f:
        f.write(html)
    print(f" -> '{ARCHIVO_SALIDA_WEB}' actualizado exitosamente ({fecha_actual}).")

def publicar_en_github_pages():
    print("Publicando cambios en GitHub Pages...")
    subprocess.run(["git", "add", ARCHIVO_SALIDA_WEB], check=False)
    subprocess.run(["git", "commit", -m, f"Actualización automática {datetime.now().strftime('%d/%m/%Y %H:%M')}"], check=False)
    subprocess.run(["git", "push", "origin", "main"], check=False)
    print(" -> ¡Publicado en línea con éxito!")

if __name__ == "__main__":
    r_c, r_i = descargar_reportes_sigo()
    actualizar_dashboard_html(r_c, r_i)
    publicar_en_github_pages()
