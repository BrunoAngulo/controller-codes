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

# ============================================================
#  CONFIGURACIÓN
# ============================================================

# Solo se copian estos formatos
ALLOWED_EXT = {".pdf", ".doc", ".docx"}

# Carpetas cuyos SUBCONTENIDOS se comprimen como Libromedia.
# El script zipea cada subcarpeta que encuentre dentro de 'recursos/'.
# Agrega aquí el nombre de cualquier carpeta que contenga Libromedias.
LIBROMEDIA_SOURCES = {
    "libmed",
}

# Carpetas específicas que SIEMPRE se comprimen, sin importar dónde estén.
ALWAYS_ZIP = {
    "SENPAI_Asistente_cooperativo",
}


# ============================================================
#  UTILIDADES
# ============================================================

def log(msg: str, indent: int = 0):
    print("  " * indent + msg)


def get_unit_folder(name: str) -> str:
    """
    Detecta número de unidad en el nombre (case-insensitive).
    Acepta: U01, U1, U05, U5, UNIDAD01, Unidad_5, unidad 05 …
    Retorna 'Unidad 01', 'Unidad 05' … o 'Documentos'.
    """
    m = re.search(r"(?:unidad|u)[\s_\-]*0*(\d+)", name, re.IGNORECASE)
    if m:
        return f"Unidad {int(m.group(1)):02d}"
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
    """
    Lee TODOS los <a href="…"> del index.html y devuelve
    {nombre_clave: texto_visible} donde nombre_clave es:
      - el nombre del archivo  (RP25_PDF_EL_U1_CO1.pdf  → "Ficha 01")
      - el nombre de la carpeta (LMLACO1RAU01            → "Comprensión U01")
    Sirve tanto para nombrar archivos copiados como ZIPs de Libromedia.
    """
    names: dict = {}
    if not os.path.isfile(html_path):
        return names
    try:
        with open(html_path, "r", encoding="utf-8", errors="ignore") as fh:
            soup = BeautifulSoup(fh.read(), "html.parser")
        for a in soup.find_all("a", href=True):
            href  = a["href"]
            # Quita query strings y fragmentos
            clean = re.split(r"[?#]", href)[0].strip().rstrip("/")
            if not clean:
                continue

            # Si el href apunta a una carpeta via su index.html
            # ej. "recursos/LMLACO1RAU01/index.html" → clave = "LMLACO1RAU01"
            if re.search(r"/index\.html?$", clean, re.IGNORECASE):
                fname = os.path.basename(os.path.dirname(clean))
            else:
                fname = os.path.basename(clean)

            if not fname:
                continue
            vname = a.get_text(strip=True)
            if vname:
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
#  PROCESAMIENTO DE UNA CARPETA
# ============================================================

def process_folder(input_path: str, output_base: str) -> list:
    """
    Procesa input_path (o su subcarpeta 'recursos/' si existe).

    Regla de compresión:
      - Si el nombre de la carpeta fuente está en LIBROMEDIA_SOURCES
        → cada subcarpeta se comprime en ZIP.
      - Si el nombre de una subcarpeta está en ALWAYS_ZIP
        → esa subcarpeta específica se comprime.
      - En cualquier otro caso solo se copian archivos PDF/Word.
    """
    folder_label = os.path.basename(input_path.rstrip("\\/"))

    recursos  = os.path.join(input_path, "recursos")
    work_path = recursos if os.path.isdir(recursos) else input_path

    # ¿Esta carpeta fuente es una fuente de Libromedias?
    zip_all_subfolders = folder_label.lower() in {s.lower() for s in LIBROMEDIA_SOURCES}

    records: list = []
    log(f"Procesando : {work_path}", 1)
    if zip_all_subfolders:
        log(f"[Libromedia source] Todas las subcarpetas se zipearán", 2)

    try:
        entries = sorted(os.listdir(work_path))
    except PermissionError:
        log(f"[ERROR] Sin permisos: {work_path}", 1)
        return records
    except FileNotFoundError:
        log(f"[ERROR] No existe: {work_path}", 1)
        return records

    # Lee index.html desde work_path y también desde input_path (si son distintos)
    # y combina los resultados para máxima cobertura de nombres visibles.
    root_html_names = extract_html_names(os.path.join(work_path, "index.html"))
    if work_path != input_path:
        root_html_names.update(
            extract_html_names(os.path.join(input_path, "index.html"))
        )

    for entry in entries:
        ep = os.path.join(work_path, entry)

        # ── Subcarpeta ────────────────────────────────────────────────
        if os.path.isdir(ep):
            should_zip = zip_all_subfolders or (entry in ALWAYS_ZIP)
            if not should_zip:
                continue  # carpeta normal → ignorar

            unit       = get_unit_folder(entry)
            zip_name   = f"{entry}.zip"
            zip_path   = os.path.join(output_base, unit, zip_name)
            html_path  = os.path.join(ep, "index.html")
            html_names = extract_html_names(html_path)

            # Nombre visible: primero busca en el index del padre,
            # luego en el título interno del Libromedia, fallback = nombre carpeta
            vis_name = (
                root_html_names.get(entry)
                or get_libromedia_title(html_path, entry)
            )

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
#  EXPANSIÓN DE RUTAS DE ENTRADA
# ============================================================

def expand_input_path(input_path: str) -> list:
    """
    Si el path tiene subcarpeta 'recursos/'      → [input_path]
    Si tiene archivos PDF/Word directamente       → [input_path]
    Si sus hijos tienen 'recursos/' (raíz padre)  → lista de esos hijos
    """
    if os.path.isdir(os.path.join(input_path, "recursos")):
        return [input_path]

    try:
        entries = os.listdir(input_path)
    except Exception:
        return [input_path]

    for e in entries:
        ep = os.path.join(input_path, e)
        if os.path.isfile(ep) and Path(e).suffix.lower() in ALLOWED_EXT:
            return [input_path]

    children = [
        os.path.join(input_path, e)
        for e in sorted(entries)
        if os.path.isdir(os.path.join(input_path, e))
        and os.path.isdir(os.path.join(input_path, e, "recursos"))
    ]
    if children:
        log(f"[AUTO] {len(children)} subcarpetas detectadas en "
            f"'{os.path.basename(input_path)}'", 1)
        return children

    return [input_path]


# ============================================================
#  OUTPUT BASE
# ============================================================

def resolve_output_base(folders: list) -> str:
    parents = {os.path.dirname(p.rstrip("\\/")) for p in folders}
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

    folders_to_process: list = []
    for f in valid_folders:
        folders_to_process.extend(expand_input_path(f))

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

    if all_records:
        unidades = sorted({r["Carpeta contenedora"] for r in all_records})
        print(f"\n  Carpetas generadas en OUTPUT:")
        for u in unidades:
            n = sum(1 for r in all_records
                    if r["Carpeta contenedora"] == u
                    and "dentro de ZIP" not in r["Tipo"])
            print(f"   • {u}  ({n} elemento(s))")

    # ── Carpeta unificada con todo junto ─────────────────────────────
    completo_dir = os.path.join(output_base, "Completo")
    os.makedirs(completo_dir, exist_ok=True)
    print(f"\n{'─' * 64}")
    print("  Copiando todo a 'Completo/'...")
    copied = 0
    for r in all_records:
        if "dentro de ZIP" in r["Tipo"]:
            continue
        src = r["Ruta destino"]
        dst = os.path.join(completo_dir, r["Nombre archivo"])
        if os.path.isfile(src) and safe_copy(src, dst):
            copied += 1
    print(f"  → {copied} elemento(s) copiado(s) a Completo/")

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
