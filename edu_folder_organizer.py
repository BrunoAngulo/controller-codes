#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Organizador de Carpetas Educativas
Uso:
  - Arrastra carpetas sobre ejecutar_organizador.bat
  - Clic derecho → "Enviar a" → Organizar carpeta educativa
  - Clic derecho → "Organizar carpeta educativa"
"""

import os
import sys
import re
import shutil
import zipfile
from pathlib import Path
from datetime import datetime

import pandas as pd
from bs4 import BeautifulSoup

# Solo se procesan / copian estos formatos
ALLOWED_EXT = {".pdf", ".doc", ".docx"}


# ============================================================
#  UTILIDADES
# ============================================================

def log(msg: str, indent: int = 0):
    print("  " * indent + msg)


def get_unit_folder(name: str) -> str:
    """
    U01, U1, U02, U2, U10 … (case-insensitive)  →  'Unidad 01', 'Unidad 02' …
    Sin coincidencia                              →  'Documentos'
    """
    m = re.search(r"[Uu]0*(\d+)", name)
    return f"Unidad {int(m.group(1)):02d}" if m else "Documentos"


def safe_copy(src: str, dst: str) -> bool:
    try:
        os.makedirs(os.path.dirname(dst), exist_ok=True)
        shutil.copy2(src, dst)
        return True
    except PermissionError:
        log(f"[ERROR] Archivo en uso o sin permisos: {os.path.basename(src)}", 3)
    except FileNotFoundError:
        log(f"[ERROR] No encontrado: {src}", 3)
    except OSError as exc:
        log(f"[ERROR] Al copiar {os.path.basename(src)}: {exc}", 3)
    return False


def zip_folder(folder_path: str, zip_path: str) -> bool:
    try:
        os.makedirs(os.path.dirname(zip_path), exist_ok=True)
        base = os.path.dirname(folder_path)
        with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
            for root, _, files in os.walk(folder_path):
                for fname in files:
                    fp = os.path.join(root, fname)
                    zf.write(fp, os.path.relpath(fp, base))
        return True
    except PermissionError:
        log(f"[ERROR] Sin permisos para comprimir: {folder_path}", 3)
    except Exception as exc:
        log(f"[ERROR] Al comprimir: {exc}", 3)
    return False


def extract_html_names(html_path: str) -> dict:
    """
    Lee un index.html y devuelve {nombre_archivo: texto_visible}
    para enlaces del tipo <a href="recursos/archivo.pdf">Nombre</a>.
    """
    names: dict = {}
    if not os.path.isfile(html_path):
        return names
    try:
        with open(html_path, "r", encoding="utf-8", errors="ignore") as fh:
            soup = BeautifulSoup(fh.read(), "html.parser")
        for a in soup.find_all("a", href=True):
            href = a["href"]
            if re.search(r"recursos/", href, re.IGNORECASE):
                # Quita query strings y fragmentos
                fname = os.path.basename(re.split(r"[?#]", href)[0])
                vname = a.get_text(strip=True)
                if fname and vname:
                    names[fname] = vname
    except Exception as exc:
        log(f"[WARN] Al leer {os.path.basename(html_path)}: {exc}", 3)
    return names


def get_libromedia_title(html_path: str, fallback: str) -> str:
    """Extrae el título del Libromedia desde su index.html."""
    if not os.path.isfile(html_path):
        return fallback
    try:
        with open(html_path, "r", encoding="utf-8", errors="ignore") as fh:
            soup = BeautifulSoup(fh.read(), "html.parser")
        if soup.title and soup.title.string:
            t = soup.title.string.strip()
            if t:
                return t
        for tag in ("h1", "h2", "h3"):
            el = soup.find(tag)
            if el:
                t = el.get_text(strip=True)
                if t:
                    return t
    except Exception:
        pass
    return fallback


def is_libromedia(folder_path: str) -> bool:
    """Una carpeta es Libromedia si contiene index.html."""
    return os.path.isfile(os.path.join(folder_path, "index.html"))


def save_excel(records: list, folder: str, filename: str):
    if not records:
        log("[INFO] Sin registros — no se genera Excel.", 2)
        return
    try:
        os.makedirs(folder, exist_ok=True)
        cols = [
            "Nombre visible", "Nombre archivo", "Ruta destino",
            "Tipo", "Carpeta contenedora", "Carpeta fuente",
        ]
        out_path = os.path.join(folder, filename)
        pd.DataFrame(records, columns=cols).to_excel(out_path, index=False)
        log(f"[OK] Excel → {out_path}", 2)
    except Exception as exc:
        log(f"[ERROR] Guardando Excel: {exc}", 2)


# ============================================================
#  PROCESAMIENTO DE UNA CARPETA
# ============================================================

def process_folder(input_path: str, output_base: str) -> list:
    """
    Procesa input_path (o su subcarpeta 'recursos' si existe).
    Los archivos van directamente a output_base/Unidad XX/  o  output_base/Documentos/
    sin subcarpetas intermedias por nombre de fuente.
    """
    folder_label = os.path.basename(input_path.rstrip("\\/"))

    # Si la carpeta tiene subcarpeta 'recursos', trabajar desde ahí
    recursos = os.path.join(input_path, "recursos")
    work_path = recursos if os.path.isdir(recursos) else input_path

    records: list = []

    log(f"Ruta de trabajo : {work_path}", 1)

    try:
        entries = sorted(os.listdir(work_path))
    except PermissionError:
        log(f"[ERROR] Sin permisos para leer: {work_path}", 1)
        return records
    except FileNotFoundError:
        log(f"[ERROR] Carpeta no existe: {work_path}", 1)
        return records

    # Nombres visibles desde el index.html de la carpeta raíz (opcional)
    root_html_names = extract_html_names(os.path.join(work_path, "index.html"))

    for entry in entries:
        ep = os.path.join(work_path, entry)

        # ── Libromedia: subcarpeta que contiene index.html ────────────
        if os.path.isdir(ep) and is_libromedia(ep):
            unit       = get_unit_folder(entry)
            zip_name   = f"{entry}.zip"
            # Va directo a OUTPUT/Unidad XX/ o OUTPUT/Documentos/
            zip_path   = os.path.join(output_base, unit, zip_name)
            html_path  = os.path.join(ep, "index.html")
            vis_name   = get_libromedia_title(html_path, entry)
            html_names = extract_html_names(html_path)

            log(f"[ZIP ] {entry}  →  {unit}/{zip_name}", 2)
            if zip_folder(ep, zip_path):
                records.append({
                    "Nombre visible":      vis_name,
                    "Nombre archivo":      zip_name,
                    "Ruta destino":        zip_path,
                    "Tipo":                "CARPETA/ZIP",
                    "Carpeta contenedora": unit,
                    "Carpeta fuente":      folder_label,
                })
                # Registro individual de cada recurso dentro del ZIP
                for fname, vname in html_names.items():
                    records.append({
                        "Nombre visible":      vname,
                        "Nombre archivo":      fname,
                        "Ruta destino":        os.path.join(zip_path, entry, fname),
                        "Tipo":                "ARCHIVO (dentro de ZIP)",
                        "Carpeta contenedora": unit,
                        "Carpeta fuente":      folder_label,
                    })
            continue

        # ── Archivo PDF / Word ────────────────────────────────────────
        if not os.path.isfile(ep):
            continue
        if Path(entry).suffix.lower() not in ALLOWED_EXT:
            continue

        unit     = get_unit_folder(entry)
        # Va directo a OUTPUT/Unidad XX/ o OUTPUT/Documentos/
        dst      = os.path.join(output_base, unit, entry)
        vis_name = root_html_names.get(entry, Path(entry).stem)

        log(f"[DOC ] {entry}  →  {unit}", 2)
        if safe_copy(ep, dst):
            records.append({
                "Nombre visible":      vis_name,
                "Nombre archivo":      entry,
                "Ruta destino":        dst,
                "Tipo":                "ARCHIVO",
                "Carpeta contenedora": unit,
                "Carpeta fuente":      folder_label,
            })

    return records


# ============================================================
#  RESOLUCIÓN DE CARPETA OUTPUT
# ============================================================

def resolve_output_base(input_folders: list) -> str:
    """
    • Si todas las carpetas comparten el mismo padre  →  OUTPUT junto a ellas.
    • Si vienen de lugares distintos                  →  Desktop/OUTPUT_Educativo.
    """
    parents = {os.path.dirname(p.rstrip("\\/")) for p in input_folders}
    if len(parents) == 1:
        return os.path.join(parents.pop(), "OUTPUT")
    desktop = os.path.join(os.path.expanduser("~"), "Desktop", "OUTPUT_Educativo")
    return desktop


# ============================================================
#  MAIN
# ============================================================

def main():
    print("=" * 64)
    print("   ORGANIZADOR DE CARPETAS EDUCATIVAS")
    print(f"   {datetime.now().strftime('%Y-%m-%d  %H:%M:%S')}")
    print("=" * 64)

    # ── Leer carpetas desde argumentos (drag-drop / Send To / menú) ──
    raw_args      = sys.argv[1:]
    valid_folders = [os.path.normpath(a) for a in raw_args if os.path.isdir(a)]

    if not valid_folders:
        print()
        print("  No se recibieron carpetas válidas.")
        print()
        print("  Cómo usar este script:")
        print("   1. Arrastra una o varias carpetas sobre 'ejecutar_organizador.bat'")
        print("   2. Clic derecho en carpeta → 'Enviar a' → Organizar carpeta educativa")
        print("   3. Clic derecho en carpeta → 'Organizar carpeta educativa'")
        print("   4. Selecciona varias carpetas y arrástralas sobre el .bat")
        print()
        input("  Presiona Enter para cerrar...")
        sys.exit(1)

    output_base = resolve_output_base(valid_folders)
    os.makedirs(output_base, exist_ok=True)

    print(f"\n  Carpetas a procesar : {len(valid_folders)}")
    print(f"  Carpeta de salida   : {output_base}\n")

    all_records: list = []

    for folder_path in valid_folders:
        label = os.path.basename(folder_path)
        print(f"\n{'─' * 64}")
        print(f"  Procesando: {label}")
        print(f"  Ruta      : {folder_path}")
        print(f"{'─' * 64}")

        records = process_folder(folder_path, output_base)

        # Excel individual de esta fuente va a la raíz de OUTPUT
        save_excel(records, output_base, f"resumen_{label}.xlsx")

        print(f"  → {len(records)} registro(s) generado(s)")
        all_records.extend(records)

    print(f"\n{'─' * 64}")
    print("  Generando Excel global...")
    save_excel(all_records, output_base, "resumen_global.xlsx")

    print(f"\n{'=' * 64}")
    print(f"  COMPLETADO")
    print(f"  Total registros : {len(all_records)}")
    print(f"  Salida          : {output_base}")
    print(f"{'=' * 64}")
    print()
    input("  Presiona Enter para cerrar...")


if __name__ == "__main__":
    main()
