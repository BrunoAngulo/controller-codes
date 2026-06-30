#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Organizador de Carpetas Educativas
Primaria:
  - Arrastra carpetas sobre ejecutar_organizador_primaria.bat
Secundaria:
  - Arrastra carpetas sobre ejecutar_organizador_secundaria.bat
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


def find_visible_name(entry: str, html_names: dict) -> str:
    """
    Busca el nombre visible para entry en html_names.
    1. Match directo por nombre de archivo.
    2. Para .docx: intenta el equivalente PDF (_DOC_ → _PDF_).
    3. Fallback: nombre de la carpeta destino (Unidad 01, Documentos, etc.).
    """
    if entry in html_names:
        return html_names[entry]
    p = Path(entry)
    if p.suffix.lower() in {".doc", ".docx"}:
        pdf_equiv = re.sub(r"_DOC_", "_PDF_", p.stem, flags=re.IGNORECASE) + ".pdf"
        if pdf_equiv in html_names:
            return html_names[pdf_equiv]
    return get_unit_folder(entry)


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


def _infer_fname_from_vname(template_fname: str, vname: str) -> str | None:
    """
    Cuando el mismo href se repite con distintos textos visibles (bug copy-paste),
    intenta inferir el nombre real del archivo sustituyendo el número de unidad
    del texto visible en el patrón del href.

    Ejemplo:
      template_fname = "RA25_PDF_LAM_U01_CO2.pdf"
      vname          = "Unidad 02"
      → "RA25_PDF_LAM_U02_CO2.pdf"
    """
    m = re.search(r"(?:unidad|u)\s*0*(\d+)", vname, re.IGNORECASE)
    if not m:
        return None
    unit_num = int(m.group(1))
    p = Path(template_fname)
    new_stem = re.sub(r"_U\d+_", f"_U{unit_num:02d}_", p.stem, count=1, flags=re.IGNORECASE)
    if new_stem == p.stem:
        return None
    return new_stem + p.suffix


def extract_html_names(html_path: str) -> dict:
    """
    Lee TODOS los <a href="…"> del index.html y devuelve
    {nombre_clave: texto_visible} donde nombre_clave es:
      - el nombre del archivo  (RP25_PDF_EL_U1_CO1.pdf  → "Ficha 01")
      - el nombre de la carpeta (LMLACO1RAU01            → "Comprensión U01")

    Cuando el mismo href aparece varias veces con distintos textos visibles
    (bug de copia-pega en el HTML), infiere el nombre real del archivo a
    partir del número de unidad del texto visible.
    """
    names: dict = {}
    if not os.path.isfile(html_path):
        return names
    try:
        with open(html_path, "r", encoding="utf-8", errors="ignore") as fh:
            soup = BeautifulSoup(fh.read(), "html.parser")
        for a in soup.find_all("a", href=True):
            href  = a["href"]
            clean = re.split(r"[?#]", href)[0].strip().rstrip("/")
            if not clean:
                continue

            if re.search(r"/index\.html?$", clean, re.IGNORECASE):
                fname = os.path.basename(os.path.dirname(clean))
            else:
                fname = os.path.basename(clean)

            if not fname:
                continue
            vname = a.get_text(strip=True)
            if not vname:
                continue

            if fname not in names:
                names[fname] = vname
            else:
                # Href duplicado: intenta inferir el nombre real del archivo
                inferred = _infer_fname_from_vname(fname, vname)
                if inferred and inferred not in names:
                    names[inferred] = vname
    except Exception as exc:
        log(f"[WARN] HTML {os.path.basename(html_path)}: {exc}", 3)
    return names


def collect_all_html_names(root_path: str) -> dict:
    """
    Recorre todo el árbol bajo root_path, lee cada index.html encontrado
    y acumula {nombre_archivo: nombre_visible}.
    Los archivos más cercanos a la raíz tienen prioridad (no se sobreescriben).
    """
    names: dict = {}
    for dirpath, dirs, filenames in os.walk(root_path):
        dirs.sort()
        if "index.html" in filenames:
            partial = extract_html_names(os.path.join(dirpath, "index.html"))
            for k, v in partial.items():
                if k not in names:
                    names[k] = v
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


_DEFAULT_EXCEL_COLS = [
    "Nombre visible", "Nombre archivo", "Ruta destino",
    "Tipo", "Carpeta contenedora", "Carpeta fuente",
]

_SEC_EXCEL_COLS = [
    "Nombre visible", "Nombre archivo", "Ruta destino",
    "Tipo", "Carpeta contenedora", "Carpeta fuente", "Para",
]


def save_excel(records: list, folder: str, filename: str, cols: list | None = None):
    if not records:
        log("[INFO] Sin registros — Excel no generado.", 2)
        return
    try:
        os.makedirs(folder, exist_ok=True)
        out_cols = cols if cols is not None else _DEFAULT_EXCEL_COLS
        out_path = os.path.join(folder, filename)
        pd.DataFrame(records, columns=out_cols).to_excel(out_path, index=False)
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

    # Lee todos los index.html del árbol completo (input_path y work_path)
    # para máxima cobertura de nombres visibles.
    root_html_names = collect_all_html_names(input_path)
    if work_path != input_path:
        for k, v in collect_all_html_names(work_path).items():
            if k not in root_html_names:
                root_html_names[k] = v

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
        vis_name = find_visible_name(entry, root_html_names)

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
#  PROCESAMIENTO SECUNDARIA
# ============================================================

def process_secondary_folder(root_path: str, output_base: str) -> list:
    """
    Procesa una carpeta de secundaria con estructura numerada.

    Estructura esperada:
      root_path/
        01/                   ← Unidad 01
          CDMAT*A/recursos/   ← alumno
          CDMAT*P/recursos/   ← profesor
          LMLA*/              ← se zipea
          LMTE*/              ← se zipea
        02/ ...
        (archivos en la raíz) → Documentos
    """
    folder_label = os.path.basename(root_path.rstrip("\\/"))
    records: list = []

    log(f"Procesando (secundaria): {root_path}", 1)

    html_names = collect_all_html_names(root_path)

    try:
        root_entries = sorted(os.listdir(root_path))
    except Exception as exc:
        log(f"[ERROR] No se pudo listar: {root_path}: {exc}", 1)
        return records

    # ── Archivos en la raíz → Documentos ─────────────────────────────
    for entry in root_entries:
        ep = os.path.join(root_path, entry)
        if os.path.isfile(ep) and Path(entry).suffix.lower() in ALLOWED_EXT:
            dst      = os.path.join(output_base, "Documentos", entry)
            vis_name = find_visible_name(entry, html_names)
            log(f"[DOC ] {entry}  →  Documentos/", 2)
            if safe_copy(ep, dst):
                records.append({
                    "Para":                "SyD",
                    "Nombre visible":      vis_name,
                    "Nombre archivo":      entry,
                    "Ruta destino":        dst,
                    "Tipo":                "ARCHIVO",
                    "Carpeta contenedora": "Documentos",
                    "Carpeta fuente":      folder_label,
                })

    # ── Carpetas numeradas (unidades) ─────────────────────────────────
    unit_re    = re.compile(r"^\d+$")
    alumno_re  = re.compile(r"^CDMAT.*A$", re.IGNORECASE)
    profesor_re = re.compile(r"^CDMAT.*P$", re.IGNORECASE)

    for entry in root_entries:
        ep = os.path.join(root_path, entry)
        if not os.path.isdir(ep) or not unit_re.match(entry):
            continue

        unit_label = f"Unidad {int(entry):02d}"
        log(f"\n  {unit_label} ({entry}/)", 1)

        try:
            sub_entries = sorted(os.listdir(ep))
        except Exception:
            continue

        alumno_dir  = None
        profesor_dir = None
        zip_dirs    = []

        for sub in sub_entries:
            sp = os.path.join(ep, sub)
            if not os.path.isdir(sp):
                continue
            su = sub.upper()
            if su.startswith("LMLA") or su.startswith("LMTE"):
                zip_dirs.append(sp)
            elif alumno_re.match(sub):
                alumno_dir = sp
            elif profesor_re.match(sub):
                profesor_dir = sp

        # ZIP LMLA* y LMTE*
        for lm_dir in zip_dirs:
            lm_name  = os.path.basename(lm_dir)
            zip_name = f"{lm_name}.zip"
            zip_path = os.path.join(output_base, unit_label, zip_name)
            vis_name = html_names.get(lm_name) or lm_name
            log(f"[ZIP ] {lm_name}  →  {unit_label}/", 2)
            if zip_folder(lm_dir, zip_path):
                records.append({
                    "Para":                "SyD",
                    "Nombre visible":      vis_name,
                    "Nombre archivo":      zip_name,
                    "Ruta destino":        zip_path,
                    "Tipo":                "CARPETA/ZIP",
                    "Carpeta contenedora": unit_label,
                    "Carpeta fuente":      folder_label,
                })

        if not alumno_dir:
            log(f"[WARN] Sin carpeta alumno (CDMAT*A) en {entry}/", 2)
        if not profesor_dir:
            log(f"[WARN] Sin carpeta profesor (CDMAT*P) en {entry}/", 2)
        if not alumno_dir or not profesor_dir:
            continue

        # Intersección de recursos/
        alumno_rec   = os.path.join(alumno_dir,  "recursos")
        profesor_rec = os.path.join(profesor_dir, "recursos")

        if not os.path.isdir(alumno_rec) or not os.path.isdir(profesor_rec):
            log(f"[WARN] Falta carpeta 'recursos' en {entry}/", 2)
            continue

        try:
            alumno_set   = set(os.listdir(alumno_rec))
            profesor_set = set(os.listdir(profesor_rec))
        except Exception as exc:
            log(f"[ERROR] Listando recursos: {exc}", 2)
            continue

        all_res = sorted(alumno_set | profesor_set)
        log(f"  Recursos: {len(all_res)} elemento(s)", 2)

        for res_entry in all_res:
            in_a = res_entry in alumno_set
            in_p = res_entry in profesor_set

            if in_a and in_p:
                para = "SyD"
                rep  = os.path.join(alumno_rec, res_entry)
            elif in_a:
                para = "S"
                rep  = os.path.join(alumno_rec, res_entry)
            else:
                para = "D"
                rep  = os.path.join(profesor_rec, res_entry)

            if os.path.isdir(rep):
                zip_name = f"{res_entry}.zip"
                zip_path = os.path.join(output_base, unit_label, zip_name)
                vis_name = html_names.get(res_entry) or res_entry
                log(f"[ZIP ] {res_entry} [{para}]  →  {unit_label}/", 2)
                if zip_folder(rep, zip_path):
                    records.append({
                        "Para":                para,
                        "Nombre visible":      vis_name,
                        "Nombre archivo":      zip_name,
                        "Ruta destino":        zip_path,
                        "Tipo":                "CARPETA/ZIP",
                        "Carpeta contenedora": unit_label,
                        "Carpeta fuente":      folder_label,
                    })

            elif os.path.isfile(rep) and Path(res_entry).suffix.lower() in ALLOWED_EXT:
                dst      = os.path.join(output_base, unit_label, res_entry)
                vis_name = find_visible_name(res_entry, html_names)
                log(f"[DOC ] {res_entry} [{para}]  →  {unit_label}/", 2)
                if safe_copy(rep, dst):
                    records.append({
                        "Para":                para,
                        "Nombre visible":      vis_name,
                        "Nombre archivo":      res_entry,
                        "Ruta destino":        dst,
                        "Tipo":                "ARCHIVO",
                        "Carpeta contenedora": unit_label,
                        "Carpeta fuente":      folder_label,
                    })

    return records


# ============================================================
#  MAIN
# ============================================================

def main_primaria():
    print("=" * 64)
    print("   ORGANIZADOR DE CARPETAS EDUCATIVAS — PRIMARIA")
    print(f"   {datetime.now().strftime('%Y-%m-%d  %H:%M:%S')}")
    print("=" * 64)

    raw_args      = [a for a in sys.argv[1:] if a != "--primaria"]
    valid_folders = [os.path.normpath(a) for a in raw_args if os.path.isdir(a)]

    if not valid_folders:
        print()
        print("  No se recibieron carpetas válidas.")
        print()
        print("  Formas de uso:")
        print("   1. Arrastra carpetas sobre 'ejecutar_organizador_primaria.bat'")
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


def main_secundaria():
    print("=" * 64)
    print("   ORGANIZADOR DE CARPETAS EDUCATIVAS — SECUNDARIA")
    print(f"   {datetime.now().strftime('%Y-%m-%d  %H:%M:%S')}")
    print("=" * 64)

    raw_args      = [a for a in sys.argv[1:] if a != "--secundaria"]
    valid_folders = [os.path.normpath(a) for a in raw_args if os.path.isdir(a)]

    if not valid_folders:
        print()
        print("  No se recibieron carpetas válidas.")
        print()
        print("  Formas de uso:")
        print("   1. Arrastra carpetas sobre 'ejecutar_organizador_secundaria.bat'")
        print("   2. Clic derecho → 'Organizar carpeta educativa (secundaria)'")
        print()
        input("  Presiona Enter para cerrar...")
        sys.exit(1)

    print(f"\n  Carpetas a procesar : {len(valid_folders)}\n")

    all_records: list = []
    last_output_base: str = ""

    for folder_path in valid_folders:
        # OUTPUT se crea dentro de cada carpeta arrastrada
        output_base = os.path.join(folder_path, "OUTPUT")
        os.makedirs(output_base, exist_ok=True)
        last_output_base = output_base

        label = os.path.basename(folder_path)
        print(f"\n{'─' * 64}")
        print(f"  {label}")
        print(f"  Salida: {output_base}")
        print(f"{'─' * 64}")

        records = process_secondary_folder(folder_path, output_base)
        print(f"  → {len(records)} registro(s)")
        all_records.extend(records)

        if records:
            unidades = sorted({r["Carpeta contenedora"] for r in records})
            print(f"\n  Carpetas generadas:")
            for u in unidades:
                n = sum(1 for r in records
                        if r["Carpeta contenedora"] == u
                        and "dentro de ZIP" not in r["Tipo"])
                print(f"   • {u}  ({n} elemento(s))")

        completo_dir = os.path.join(output_base, "Completo")
        os.makedirs(completo_dir, exist_ok=True)
        print(f"\n{'─' * 64}")
        print("  Copiando todo a 'Completo/'...")
        copied = 0
        for r in records:
            if "dentro de ZIP" in r["Tipo"]:
                continue
            src = r["Ruta destino"]
            dst = os.path.join(completo_dir, r["Nombre archivo"])
            if os.path.isfile(src) and safe_copy(src, dst):
                copied += 1
        print(f"  → {copied} elemento(s) copiado(s) a Completo/")

        print(f"\n{'─' * 64}")
        save_excel(records, output_base, "resumen_global.xlsx", cols=_SEC_EXCEL_COLS)

    output_base = last_output_base

    print(f"\n{'=' * 64}")
    print(f"  COMPLETADO  —  {len(all_records)} registros")
    print(f"  Salida: {output_base}")
    print(f"{'=' * 64}")
    print()
    input("  Presiona Enter para cerrar...")


if __name__ == "__main__":
    if "--secundaria" in sys.argv:
        main_secundaria()
    else:
        main_primaria()
