#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Organizador de Carpetas Educativas - CDCOMPRI1P
Clasifica archivos por unidad, comprime Libromedias en ZIP y genera Excel resumen.
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

# ============================================================
#  CONFIGURACIÓN
# ============================================================
BASE_PATH   = r"C:\Users\bangulo\Downloads\CDCOMPRI1P"
OUTPUT_PATH = os.path.join(BASE_PATH, "OUTPUT")

# Solo se copian/procesan estos tipos de archivo
ALLOWED_EXT = {".pdf", ".doc", ".docx"}

SOURCE_FOLDERS = [
    {
        "key":  "doc_curr",
        "name": "Documentos_Curriculares",
        "path": os.path.join(BASE_PATH, "doc_curr", "recursos"),
        "type": "files",
    },
    {
        "key":  "libmed",
        "name": "Libromedia",
        "path": os.path.join(BASE_PATH, "libmed", "recursos"),
        "type": "libromedia",
    },
    {
        "key":  "est_lect",
        "name": "Estrategias_de_Lectura",
        "path": os.path.join(BASE_PATH, "est_lect", "recursos"),
        "type": "libromedia",
    },
    {
        "key":  "dif_lect",
        "name": "Dificultades_Lectoescritura",
        "path": os.path.join(BASE_PATH, "dif_lect", "recursos"),
        "type": "files",
    },
    {
        "key":  "guia_met",
        "name": "Guia_Metodologica",
        "path": os.path.join(BASE_PATH, "guia_met", "recursos"),
        "type": "files",
    },
    {
        "key":  "mod_ie",
        "name": "Instrumentos_Evaluacion",
        "path": os.path.join(BASE_PATH, "mod_ie", "recursos"),
        "type": "files",
    },
    {
        "key":  "mat_imp",
        "name": "Material_Imprimible",
        "path": os.path.join(BASE_PATH, "mat_imp", "recursos"),
        "type": "files",
    },
    {
        "key":  "lam_didac",
        "name": "Laminas_Didacticas",
        "path": os.path.join(BASE_PATH, "lam_didac", "recursos"),
        "type": "files",
    },
    {
        "key":             "senpai",
        "name":            "Herramientas_Cooperativas",
        "path":            os.path.join(BASE_PATH, "senpai", "recursos"),
        "type":            "mixed",
        "libromedia_dirs": ["SENPAI_Asistente_cooperativo"],
    },
]

# ============================================================
#  UTILIDADES GENERALES
# ============================================================

def log(msg: str, indent: int = 0):
    print("  " * indent + msg)


def get_unit_folder(name: str) -> str:
    """
    Detecta número de unidad en el nombre del archivo o carpeta.
    Acepta: U01, U1, U02, U2, U10, U010 (case-insensitive).
    """
    match = re.search(r"[Uu]0*(\d+)", name)
    if match:
        return f"Unidad {int(match.group(1)):02d}"
    return "Documentos"


def safe_copy(src: str, dst: str) -> bool:
    try:
        os.makedirs(os.path.dirname(dst), exist_ok=True)
        shutil.copy2(src, dst)
        return True
    except PermissionError:
        log(f"[ERROR] Archivo en uso o sin permisos: {src}", 2)
    except FileNotFoundError:
        log(f"[ERROR] No encontrado: {src}", 2)
    except Exception as exc:
        log(f"[ERROR] Al copiar {src}: {exc}", 2)
    return False


def zip_folder(folder_path: str, zip_path: str) -> bool:
    try:
        os.makedirs(os.path.dirname(zip_path), exist_ok=True)
        base_dir = os.path.dirname(folder_path)
        with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
            for root, _dirs, files in os.walk(folder_path):
                for fname in files:
                    fpath   = os.path.join(root, fname)
                    arcname = os.path.relpath(fpath, base_dir)
                    zf.write(fpath, arcname)
        return True
    except PermissionError:
        log(f"[ERROR] Sin permisos para comprimir: {folder_path}", 2)
    except Exception as exc:
        log(f"[ERROR] Al comprimir {folder_path}: {exc}", 2)
    return False


# ============================================================
#  LECTURA DE HTML
# ============================================================

def extract_html_file_names(html_path: str) -> dict:
    """
    Extrae {nombre_archivo: texto_visible} de un index.html.
    Busca: <a href="recursos/archivo.pdf">Nombre visible</a>
    """
    names = {}
    if not os.path.isfile(html_path):
        return names
    try:
        with open(html_path, "r", encoding="utf-8", errors="ignore") as fh:
            soup = BeautifulSoup(fh.read(), "html.parser")
        for a_tag in soup.find_all("a", href=True):
            href = a_tag["href"]
            if re.search(r"recursos/", href, re.IGNORECASE):
                # Quita query strings y fragmentos
                filename = os.path.basename(re.split(r"[?#]", href)[0])
                visible  = a_tag.get_text(strip=True)
                if filename and visible:
                    names[filename] = visible
    except Exception as exc:
        log(f"[WARN] No se pudo leer HTML {html_path}: {exc}", 2)
    return names


def extract_libromedia_title(html_path: str, fallback: str) -> str:
    """Extrae el título del Libromedia desde su index.html."""
    if not os.path.isfile(html_path):
        return fallback
    try:
        with open(html_path, "r", encoding="utf-8", errors="ignore") as fh:
            soup = BeautifulSoup(fh.read(), "html.parser")
        if soup.title and soup.title.string:
            return soup.title.string.strip()
        for tag in ("h1", "h2", "h3"):
            el = soup.find(tag)
            if el:
                return el.get_text(strip=True)
    except Exception:
        pass
    return fallback


def is_libromedia(folder_path: str) -> bool:
    return os.path.isfile(os.path.join(folder_path, "index.html"))


# ============================================================
#  GENERACIÓN DE EXCEL
# ============================================================

def save_excel(records: list, output_folder: str, filename: str):
    if not records:
        log(f"[INFO] Sin registros para: {filename}", 2)
        return
    try:
        os.makedirs(output_folder, exist_ok=True)
        df   = pd.DataFrame(records, columns=[
            "Nombre visible",
            "Nombre archivo",
            "Ruta destino",
            "Tipo",
            "Carpeta contenedora",
            "Carpeta fuente",
        ])
        path = os.path.join(output_folder, filename)
        df.to_excel(path, index=False)
        log(f"[OK] Excel guardado: {path}", 2)
    except Exception as exc:
        log(f"[ERROR] No se pudo guardar {filename}: {exc}", 2)


# ============================================================
#  PROCESADORES POR TIPO
# ============================================================

def process_files(cfg: dict) -> list:
    """Copia archivos PDF/Word y los clasifica por unidad."""
    src_path = cfg["path"]
    out_base = os.path.join(OUTPUT_PATH, cfg["name"])
    records  = []

    if not os.path.isdir(src_path):
        log(f"[WARN] Carpeta no encontrada: {src_path}", 1)
        return records

    # Referencias de nombres visibles desde index.html raíz (si existe)
    html_names = extract_html_file_names(os.path.join(src_path, "index.html"))

    items = sorted(os.listdir(src_path))
    for entry in items:
        entry_path = os.path.join(src_path, entry)
        if not os.path.isfile(entry_path):
            continue
        if Path(entry).suffix.lower() not in ALLOWED_EXT:
            continue

        unit_folder  = get_unit_folder(entry)
        dest_path    = os.path.join(out_base, unit_folder, entry)
        visible_name = html_names.get(entry, Path(entry).stem)

        log(f"Copiando: {entry}  →  {unit_folder}", 2)
        if safe_copy(entry_path, dest_path):
            records.append({
                "Nombre visible":      visible_name,
                "Nombre archivo":      entry,
                "Ruta destino":        dest_path,
                "Tipo":                "ARCHIVO",
                "Carpeta contenedora": unit_folder,
                "Carpeta fuente":      cfg["name"],
            })

    return records


def process_libromedia(cfg: dict) -> list:
    """Comprime cada subcarpeta Libromedia (con index.html) en ZIP."""
    src_path = cfg["path"]
    out_base = os.path.join(OUTPUT_PATH, cfg["name"])
    records  = []

    if not os.path.isdir(src_path):
        log(f"[WARN] Carpeta no encontrada: {src_path}", 1)
        return records

    items = sorted(os.listdir(src_path))
    for entry in items:
        entry_path = os.path.join(src_path, entry)
        if not os.path.isdir(entry_path):
            continue
        if not is_libromedia(entry_path):
            log(f"[SKIP] Sin index.html: {entry}", 2)
            continue

        unit_folder  = get_unit_folder(entry)
        zip_name     = f"{entry}.zip"
        zip_path     = os.path.join(out_base, unit_folder, zip_name)
        html_path    = os.path.join(entry_path, "index.html")
        visible_name = extract_libromedia_title(html_path, entry)
        html_names   = extract_html_file_names(html_path)

        log(f"Comprimiendo: {entry}  →  {unit_folder}/{zip_name}", 2)
        if zip_folder(entry_path, zip_path):
            records.append({
                "Nombre visible":      visible_name,
                "Nombre archivo":      zip_name,
                "Ruta destino":        zip_path,
                "Tipo":                "CARPETA/ZIP",
                "Carpeta contenedora": unit_folder,
                "Carpeta fuente":      cfg["name"],
            })
            # Registrar cada recurso listado en el HTML
            for fname, vname in html_names.items():
                records.append({
                    "Nombre visible":      vname,
                    "Nombre archivo":      fname,
                    "Ruta destino":        f"{zip_path}\\{entry}\\{fname}",
                    "Tipo":                "ARCHIVO (dentro de ZIP)",
                    "Carpeta contenedora": unit_folder,
                    "Carpeta fuente":      cfg["name"],
                })

    return records


def process_mixed(cfg: dict) -> list:
    """
    Carpeta mixta: archivos PDF/Word normales + carpetas Libromedia designadas.
    (Caso: senpai/recursos)
    """
    src_path     = cfg["path"]
    out_base     = os.path.join(OUTPUT_PATH, cfg["name"])
    libmed_dirs  = set(cfg.get("libromedia_dirs", []))
    records      = []

    if not os.path.isdir(src_path):
        log(f"[WARN] Carpeta no encontrada: {src_path}", 1)
        return records

    html_names = extract_html_file_names(os.path.join(src_path, "index.html"))

    items = sorted(os.listdir(src_path))
    for entry in items:
        entry_path = os.path.join(src_path, entry)

        # ── Libromedia designadas explícitamente ─────────────────────
        if os.path.isdir(entry_path) and entry in libmed_dirs:
            unit_folder  = get_unit_folder(entry)
            zip_name     = f"{entry}.zip"
            zip_path     = os.path.join(out_base, unit_folder, zip_name)
            html_path    = os.path.join(entry_path, "index.html")
            visible_name = extract_libromedia_title(html_path, entry)
            sub_names    = extract_html_file_names(html_path)

            log(f"Comprimiendo Libromedia especial: {entry}  →  {unit_folder}/{zip_name}", 2)
            if zip_folder(entry_path, zip_path):
                records.append({
                    "Nombre visible":      visible_name,
                    "Nombre archivo":      zip_name,
                    "Ruta destino":        zip_path,
                    "Tipo":                "CARPETA/ZIP",
                    "Carpeta contenedora": unit_folder,
                    "Carpeta fuente":      cfg["name"],
                })
                for fname, vname in sub_names.items():
                    records.append({
                        "Nombre visible":      vname,
                        "Nombre archivo":      fname,
                        "Ruta destino":        f"{zip_path}\\{entry}\\{fname}",
                        "Tipo":                "ARCHIVO (dentro de ZIP)",
                        "Carpeta contenedora": unit_folder,
                        "Carpeta fuente":      cfg["name"],
                    })
            continue

        # ── Archivos PDF / Word normales ──────────────────────────────
        if not os.path.isfile(entry_path):
            continue
        if Path(entry).suffix.lower() not in ALLOWED_EXT:
            continue

        unit_folder  = get_unit_folder(entry)
        dest_path    = os.path.join(out_base, unit_folder, entry)
        visible_name = html_names.get(entry, Path(entry).stem)

        log(f"Copiando: {entry}  →  {unit_folder}", 2)
        if safe_copy(entry_path, dest_path):
            records.append({
                "Nombre visible":      visible_name,
                "Nombre archivo":      entry,
                "Ruta destino":        dest_path,
                "Tipo":                "ARCHIVO",
                "Carpeta contenedora": unit_folder,
                "Carpeta fuente":      cfg["name"],
            })

    return records


# ============================================================
#  PUNTO DE ENTRADA
# ============================================================

def main():
    print("=" * 62)
    print("   ORGANIZADOR DE CARPETAS EDUCATIVAS - CDCOMPRI1P")
    print(f"   {datetime.now().strftime('%Y-%m-%d  %H:%M:%S')}")
    print("=" * 62)
    print(f"\n  Base  : {BASE_PATH}")
    print(f"  Output: {OUTPUT_PATH}\n")

    os.makedirs(OUTPUT_PATH, exist_ok=True)

    all_records   = []
    total_ok      = 0
    total_errors  = 0

    PROCESSORS = {
        "files":      process_files,
        "libromedia": process_libromedia,
        "mixed":      process_mixed,
    }

    for cfg in SOURCE_FOLDERS:
        print(f"\n{'─' * 62}")
        print(f"  [{cfg['type'].upper()}]  {cfg['name']}")
        print(f"  Origen: {cfg['path']}")
        print(f"{'─' * 62}")

        processor = PROCESSORS.get(cfg["type"])
        if processor is None:
            log(f"[ERROR] Tipo desconocido: {cfg['type']}", 1)
            continue

        records = processor(cfg)

        # Excel individual
        out_dir    = os.path.join(OUTPUT_PATH, cfg["name"])
        excel_name = f"resumen_{cfg['key']}.xlsx"
        save_excel(records, out_dir, excel_name)

        ok  = sum(1 for r in records if "ZIP" in r["Tipo"] or r["Tipo"] == "ARCHIVO")
        log(f"→ {len(records)} registro(s) generado(s)", 1)

        all_records.extend(records)
        total_ok += ok

    # Excel global
    print(f"\n{'─' * 62}")
    print("  Generando Excel global...")
    save_excel(all_records, OUTPUT_PATH, "resumen_global.xlsx")

    print(f"\n{'=' * 62}")
    print(f"  COMPLETADO: {len(all_records)} registros totales")
    print(f"  Salida: {OUTPUT_PATH}")
    print(f"{'=' * 62}\n")


if __name__ == "__main__":
    main()
