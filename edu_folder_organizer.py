#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Organizador de Carpetas Educativas
Uso:
  - Arrastra una o varias carpetas sobre ejecutar_organizador.bat
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

ALLOWED_EXT = {".pdf", ".doc", ".docx"}


# ============================================================
#  UTILIDADES
# ============================================================

def log(msg: str, indent: int = 0):
    print("  " * indent + msg)


def get_unit_folder(name: str) -> str:
    """
    Detecta número de unidad en el nombre.
    Acepta (case-insensitive):
      U01, U1, U05, U5, U08, U8
      UNIDAD01, Unidad_5, unidad 05
    Retorna 'Unidad 01', 'Unidad 05' … o 'Documentos'.
    """
    m = re.search(r"(?:unidad|u)[\s_\-]*0*(\d+)", name, re.IGNORECASE)
    if m:
        num = int(m.group(1))
        return f"Unidad {num:02d}"
    return "Documentos"


def safe_copy(src: str, dst: str) -> bool:
    try:
        os.makedirs(os.path.dirname(dst), exist_ok=True)
        shutil.copy2(src, dst)
        return True
    except PermissionError:
        log(f"[ERROR] Archivo en uso: {os.path.basename(src)}", 3)
    except FileNotFoundError:
        log(f"[ERROR] No encontrado: {src}", 3)
    except OSError as exc:
        log(f"[ERROR] {os.path.basename(src)}: {exc}", 3)
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
    """Devuelve {nombre_archivo: texto_visible} para <a href="recursos/…">."""
    names: dict = {}
    if not os.path.isfile(html_path):
        return names
    try:
        with open(html_path, "r", encoding="utf-8", errors="ignore") as fh:
            soup = BeautifulSoup(fh.read(), "html.parser")
        for a in soup.find_all("a", href=True):
            href = a["href"]
            if re.search(r"recursos/", href, re.IGNORECASE):
                fname = os.path.basename(re.split(r"[?#]", href)[0])
                vname = a.get_text(strip=True)
                if fname and vname:
                    names[fname] = vname
    except Exception as exc:
        log(f"[WARN] HTML {os.path.basename(html_path)}: {exc}", 3)
    return names


def get_libromedia_title(html_path: str, fallback: str) -> str:
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
    return os.path.isfile(os.path.join(folder_path, "index.html"))


def save_excel(records: list, folder: str, filename: str):
    if not records:
        log("[INFO] Sin registros — Excel no generado.", 2)
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
        log(f"[ERROR] Excel: {exc}", 2)


# ============================================================
#  EXPANSIÓN DE RUTAS DE ENTRADA
# ============================================================

def expand_input_path(input_path: str) -> list:
    """
    Dado un path de entrada, devuelve la lista real de carpetas a procesar.

    Prioridad:
    1. La carpeta tiene subcarpeta 'recursos'        → [input_path]
    2. La carpeta contiene archivos PDF/Word o
       subcarpetas Libromedia directamente            → [input_path]
    3. Los hijos de la carpeta tienen 'recursos'     → todos esos hijos
       (caso: se arrastra la raíz CDCOMPRI1P)
    4. Fallback                                      → [input_path]
    """
    # Caso 1: tiene 'recursos'
    if os.path.isdir(os.path.join(input_path, "recursos")):
        return [input_path]

    try:
        entries = os.listdir(input_path)
    except Exception:
        return [input_path]

    # Caso 2: tiene archivos PDF/Word o Libromedias directamente
    for e in entries:
        ep = os.path.join(input_path, e)
        if os.path.isfile(ep) and Path(e).suffix.lower() in ALLOWED_EXT:
            return [input_path]
        if os.path.isdir(ep) and is_libromedia(ep):
            return [input_path]

    # Caso 3: hijos con 'recursos' (usuario arrastró la raíz)
    children = []
    for e in sorted(entries):
        child = os.path.join(input_path, e)
        if os.path.isdir(child) and os.path.isdir(os.path.join(child, "recursos")):
            children.append(child)
    if children:
        log(f"[AUTO] Detectadas {len(children)} subcarpetas con 'recursos' dentro de "
            f"{os.path.basename(input_path)}", 1)
        return children

    return [input_path]


# ============================================================
#  PROCESAMIENTO DE UNA CARPETA
# ============================================================

def process_folder(input_path: str, output_base: str) -> list:
    """
    Procesa una carpeta (usa su subcarpeta 'recursos' si existe).
    Archivos y ZIPs van directo a OUTPUT/Unidad XX/ o OUTPUT/Documentos/.
    """
    folder_label = os.path.basename(input_path.rstrip("\\/"))

    recursos  = os.path.join(input_path, "recursos")
    work_path = recursos if os.path.isdir(recursos) else input_path

    records: list = []
    log(f"Procesando : {work_path}", 1)

    try:
        entries = sorted(os.listdir(work_path))
    except PermissionError:
        log(f"[ERROR] Sin permisos: {work_path}", 1)
        return records
    except FileNotFoundError:
        log(f"[ERROR] No existe: {work_path}", 1)
        return records

    root_html_names = extract_html_names(os.path.join(work_path, "index.html"))

    for entry in entries:
        ep = os.path.join(work_path, entry)

        # ── Libromedia ────────────────────────────────────────────────
        if os.path.isdir(ep) and is_libromedia(ep):
            unit       = get_unit_folder(entry)
            zip_name   = f"{entry}.zip"
            zip_path   = os.path.join(output_base, unit, zip_name)
            html_path  = os.path.join(ep, "index.html")
            vis_name   = get_libromedia_title(html_path, entry)
            html_names = extract_html_names(html_path)

            log(f"[ZIP ] {entry}  →  {unit}/", 2)
            if zip_folder(ep, zip_path):
                records.append({
                    "Nombre visible":      vis_name,
                    "Nombre archivo":      zip_name,
                    "Ruta destino":        zip_path,
                    "Tipo":                "CARPETA/ZIP",
                    "Carpeta contenedora": unit,
                    "Carpeta fuente":      folder_label,
                })
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
        dst      = os.path.join(output_base, unit, entry)
        vis_name = root_html_names.get(entry, Path(entry).stem)

        log(f"[DOC ] {entry}  →  {unit}/", 2)
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
#  OUTPUT BASE
# ============================================================

def resolve_output_base(input_folders: list) -> str:
    """
    Todas las carpetas comparten padre  →  OUTPUT junto a ellas.
    Carpetas de distintos lugares       →  Desktop/OUTPUT_Educativo.
    """
    parents = {os.path.dirname(p.rstrip("\\/")) for p in input_folders}
    if len(parents) == 1:
        return os.path.join(parents.pop(), "OUTPUT")
    return os.path.join(os.path.expanduser("~"), "Desktop", "OUTPUT_Educativo")


# ============================================================
#  MAIN
# ============================================================

def main():
    print("=" * 64)
    print("   ORGANIZADOR DE CARPETAS EDUCATIVAS")
    print(f"   {datetime.now().strftime('%Y-%m-%d  %H:%M:%S')}")
    print("=" * 64)

    raw_args      = sys.argv[1:]
    valid_folders = [os.path.normpath(a) for a in raw_args if os.path.isdir(a)]

    if not valid_folders:
        print()
        print("  No se recibieron carpetas válidas.")
        print()
        print("  Formas de uso:")
        print("   1. Arrastra carpetas sobre 'ejecutar_organizador.bat'")
        print("   2. Clic derecho → 'Enviar a' → Organizar carpeta educativa")
        print("   3. Clic derecho → 'Organizar carpeta educativa'")
        print()
        input("  Presiona Enter para cerrar...")
        sys.exit(1)

    # Expandir carpetas padre (ej. CDCOMPRI1P → sus subcarpetas con recursos)
    folders_to_process: list = []
    for f in valid_folders:
        expanded = expand_input_path(f)
        folders_to_process.extend(expanded)

    output_base = resolve_output_base(folders_to_process)
    os.makedirs(output_base, exist_ok=True)

    print(f"\n  Carpetas a procesar : {len(folders_to_process)}")
    print(f"  Salida              : {output_base}\n")

    all_records: list = []

    for folder_path in folders_to_process:
        label = os.path.basename(folder_path)
        print(f"\n{'─' * 64}")
        print(f"  {label}")
        print(f"{'─' * 64}")

        records = process_folder(folder_path, output_base)
        print(f"  → {len(records)} registro(s)")
        all_records.extend(records)

    # Resumen de unidades detectadas
    if all_records:
        unidades = sorted({r["Carpeta contenedora"] for r in all_records})
        print(f"\n  Carpetas creadas en OUTPUT:")
        for u in unidades:
            count = sum(1 for r in all_records if r["Carpeta contenedora"] == u
                        and "dentro de ZIP" not in r["Tipo"])
            print(f"   • {u}  ({count} elemento(s))")

    print(f"\n{'─' * 64}")
    save_excel(all_records, output_base, "resumen_global.xlsx")

    print(f"\n{'=' * 64}")
    print(f"  COMPLETADO  —  {len(all_records)} registros")
    print(f"  Salida: {output_base}")
    print(f"{'=' * 64}")
    print()
    input("  Presiona Enter para cerrar...")


if __name__ == "__main__":
    main()
