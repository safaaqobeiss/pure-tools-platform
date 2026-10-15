
# --- Début : Fonctions de conversion XBRL/XML → CSV intégrées depuis script.py ---
import os
import shutil
import stat
import zipfile
import tarfile
import pandas as pd
from decimal import Decimal, InvalidOperation
from typing import Dict, List
from typing import Dict, List

def ensure_clean_dir(path: str):
    if os.path.exists(path):
        shutil.rmtree(path, onerror=lambda func, p, exc: (os.chmod(p, stat.S_IWRITE), func(p)))
    os.makedirs(path, exist_ok=True)

def safe_copy(src: str, dst_dir: str, force_name: str = "", overwrite: bool = False) -> str:
    os.makedirs(dst_dir, exist_ok=True)
    base_name = force_name if force_name else os.path.basename(src)
    dst = os.path.join(dst_dir, base_name)
    if overwrite and os.path.exists(dst):
        os.remove(dst)
    if not overwrite and os.path.exists(dst):
        stem, ext = os.path.splitext(base_name)
        i = 2
        while True:
            candidate = os.path.join(dst_dir, f"{stem}({i}){ext}")
            if not os.path.exists(candidate):
                dst = candidate
                break
            i += 1
    shutil.copy2(src, dst)
    return dst

def _preprocess_xbrl_for_conversion(xml_path: str) -> None:
    import re as _re
    with open(xml_path, "r", encoding="utf-8") as fh:
        content = fh.read()
    root_match = _re.search(r'(<xbrli:xbrl\\b[^>]*>)(.*?)(</xbrli:xbrl>)', content, _re.DOTALL)
    if not root_match:
        return
    root_open, inner, root_close = root_match.group(1), root_match.group(2), root_match.group(3)
    schema_ref_pat = _re.compile(r'<link:schemaRef\\b[^>]*/>\\s*|<link:schemaRef\\b[^>]*>.*?</link:schemaRef>\\s*', _re.DOTALL)
    unit_pat = _re.compile(r'<xbrli:unit\\s+id="([^"]+)".*?</xbrli:unit>\\s*', _re.DOTALL)
    context_pat = _re.compile(r'<xbrli:context\\s+id="([^"]+)".*?</xbrli:context>\\s*', _re.DOTALL)
    schema_refs = [m.group(0).strip() for m in schema_ref_pat.finditer(inner)]
    inner = schema_ref_pat.sub("", inner)
    unit_matches = list(unit_pat.finditer(inner))
    inner = unit_pat.sub("", inner)
    context_matches = list(context_pat.finditer(inner))
    inner = context_pat.sub("", inner)
    context_id_map: Dict[str, str] = {}
    rebuilt_contexts: List[str] = []
    next_ctx = 1
    for m in context_matches:
        old_id = m.group(1)
        block = m.group(0)
        if old_id in context_id_map:
            continue
        new_id = f"c{next_ctx:06d}"
        next_ctx += 1
        context_id_map[old_id] = new_id
        rebuilt_contexts.append(_re.sub(r'(<xbrli:context\\s+id=")[^"]+(\")', rf'\\1{new_id}\\2', block, count=1).strip())
    unit_id_map: Dict[str, str] = {}
    rebuilt_units: List[str] = []
    next_unit = 1
    for m in unit_matches:
        old_id = m.group(1)
        block = m.group(0)
        if old_id in unit_id_map:
            continue
        new_id = f"u{next_unit:06d}"
        next_unit += 1
        unit_id_map[old_id] = new_id
        rebuilt_units.append(_re.sub(r'(<xbrli:unit\\s+id=")[^"]+(\")', rf'\\1{new_id}\\2', block, count=1).strip())
    for old_id, new_id in sorted(context_id_map.items(), key=lambda kv: len(kv[0]), reverse=True):
        inner = inner.replace(f'contextRef="{old_id}"', f'contextRef="{new_id}"')
    for old_id, new_id in sorted(unit_id_map.items(), key=lambda kv: len(kv[0]), reverse=True):
        inner = inner.replace(f'unitRef="{old_id}"', f'unitRef="{new_id}"')
    sections: List[str] = []
    sections.extend(schema_refs)
    sections.extend(rebuilt_units)
    sections.extend(rebuilt_contexts)
    remaining = inner.strip()
    if remaining:
        sections.append(remaining)
    new_inner = "\n  " + "\n  ".join(s for s in sections if s) + "\n"
    new_content = content[: root_match.start()] + root_open + new_inner + root_close + content[root_match.end() :]
    with open(xml_path, "w", encoding="utf-8") as fh:
        fh.write(new_content)


# --- Conversion XBRL/XML → CSV ZIP en Python pur ---
def extract_reports_from_zip(zip_path: str, destination_directory: str) -> int:
    if not os.path.exists(zip_path):
        raise FileNotFoundError(f"CSV Zip '{zip_path}' not found.")
    os.makedirs(destination_directory, exist_ok=True)
    extracted = 0
    with zipfile.ZipFile(zip_path, "r") as zip_ref:
        for member in zip_ref.namelist():
            normalized_member = member.replace("\\", "/")
            if "/reports/" in normalized_member and normalized_member.endswith(".csv") and not normalized_member.endswith("/"):
                idx = normalized_member.find("reports/")
                relative_path = normalized_member[idx + len("reports/"):]
                destination_path = os.path.join(destination_directory, relative_path)
                os.makedirs(os.path.dirname(destination_path), exist_ok=True)
                with zip_ref.open(member) as src, open(destination_path, "wb") as dst:
                    shutil.copyfileobj(src, dst)
                extracted += 1
    return extracted

# Nouvelle version : conversion XBRL/XML → CSV ZIP en Python pur
def convert_xml_to_zip(xml_path: str, taxonomy_tar_path: str, work_dir: str, run_tag: str, output_dir: str) -> str:
    # Ici, on suppose que la conversion consiste à extraire les rapports CSV d’un ZIP fourni (ou à copier le XML dans un ZIP cible)
    # Si la logique doit être enrichie, ajouter ici le parsing XBRL/XML → CSV
    run_dir = os.path.join(work_dir, run_tag)
    ensure_clean_dir(run_dir)
    local_xml = safe_copy(xml_path, run_dir, overwrite=True)
    local_taxo = safe_copy(taxonomy_tar_path, run_dir, force_name="xbrl_package.tar.gz", overwrite=True)
    # Ici, on simule la conversion : on zippe simplement le XML (à adapter selon besoin réel)
    produced_zip = os.path.join(output_dir, os.path.basename(local_xml).replace(".xbrl", ".zip"))
    with zipfile.ZipFile(produced_zip, "w") as zipf:
        zipf.write(local_xml, arcname=os.path.basename(local_xml))
    return produced_zip
# --- Fin : Fonctions de conversion XBRL/XML → CSV ---
"""
validate_regulation_unified.py
==============================
Outil general de validation reglementaire par mapping taxonomie.
Fonctionne pour N'IMPORTE quel regulateur (RUBA, COREP, RESOL, etc.).

Enchainement:
  1. Convertir le fichier XBRL submission en CSV (via Xml2Csv.jar + taxonomie)
  2. Lire les JSONs taxonomie pour mapper datapoint -> {Table, Row, Column}
  3. Lire les fichiers CSV input (format long: tableCode;rowId;columnId;value)
  4. Comparer cellule par cellule (table, r, c) et generer le rapport

Usage:
  python validate_regulation_unified.py \\
    --input-dir     <dossier_csv_input> \\
    --xbrl-file     <fichier_xbrl_submission> \\
    --taxonomy      <xbrl_package.tar.gz> \\
    --converter-jar <chemin/vers/Xml2Csv.jar> \\
    --output-dir    <dossier_sortie_rapports>

Les correspondances de table (ex: RB.26.01.11 -> RB.26.01.01) sont detectees
automatiquement a partir des codes de table connus dans la taxonomie.
"""


import argparse
import csv
import json
import re
import shutil
import subprocess
import sys
import tarfile
import tempfile
import zipfile
from collections import defaultdict
from dataclasses import dataclass
from datetime import datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill


# ─────────────────────────────────────────────
#  Styles Excel
# ─────────────────────────────────────────────
FILL_OK = PatternFill("solid", fgColor="C6EFCE")
FILL_DIFF = PatternFill("solid", fgColor="FFC7CE")
FILL_MISSING = PatternFill("solid", fgColor="FFEB9C")
FILL_EXTRA = PatternFill("solid", fgColor="DDEBF7")
FILL_HEADER = PatternFill("solid", fgColor="4472C4")

STATUS_FILL = {
    "OK": FILL_OK,
    "DIFF": FILL_DIFF,
    "MISSING_IN_OUTPUT": FILL_MISSING,
    "CALCULATED_OR_EXTRA": FILL_EXTRA,
}


# ─────────────────────────────────────────────
#  Utilitaires
# ─────────────────────────────────────────────
def normalize_value(val) -> str:
    """Normalise une valeur pour comparaison (espaces, decimales, etc.)."""
    if val is None:
        return ""
    s = str(val).strip().replace("\u00a0", " ").strip()
    if not s:
        return ""
    # Supprimer les prefixes de namespace XBRL (ex: eba_GA:, fr_TA:, fr_RP:, ...)
    s = re.sub(r"^[a-zA-Z]{2,4}_[A-Za-z0-9]+:", "", s)
    s_num = s.replace(" ", "")
    if "," in s_num and "." in s_num:
        if s_num.rfind(",") > s_num.rfind("."):
            s_num = s_num.replace(".", "").replace(",", ".")
        else:
            s_num = s_num.replace(",", "")
    elif "," in s_num and "." not in s_num:
        s_num = s_num.replace(",", ".")
    try:
        d = Decimal(s_num)
        return format(d.normalize(), "f")
    except InvalidOperation:
        return s


def make_stamp() -> str:
    return datetime.now().strftime("%Y_%m_%d.%H.%M.%S")


# ─────────────────────────────────────────────
#  Dataclasses
# ─────────────────────────────────────────────
@dataclass
class InputRow:
    input_line: int
    input_file: str
    table_code: str
    row_id: str
    column_id: str
    value_raw: str


@dataclass
class OutputPoint:
    table_code: str
    row_id: str
    column_id: str
    datapoint: str
    value_raw: str
    source_file: str
    cell_code: str


# ─────────────────────────────────────────────
#  CellCode regex – general pour toutes taxonomies ACPR/EBA
#  Examples: {RB.26.01.01, R0020, C0070}
#            {C_01.00, R0010, C0010}
# ─────────────────────────────────────────────
CELLCODE_RE = re.compile(
    r"\{\s*([A-Za-z0-9_.\s]+?)\s*,\s*([RC][0-9]{3,4})\s*,\s*([RC][0-9]{3,4})\s*\}",
    re.IGNORECASE,
)


# ─────────────────────────────────────────────
#  1. Conversion XBRL -> CSV
# ─────────────────────────────────────────────
_SCHEMA_REF_RE = re.compile(
    r"(<link:schemaRef\s[^>]*?/>|<link:schemaRef\s[^>]*?>.*?</link:schemaRef>)",
    re.DOTALL,
)


def _fix_xbrl_schema_position(xbrl_file: Path, work_dir: Path) -> Path:
    """
    Certains fichiers XBRL (COREP, etc.) ont le <link:schemaRef> place
    en fin de fichier au lieu de juste apres l'ouverture <xbrli:xbrl>.
    Xml2Csv.jar ne le trouve pas dans ce cas.
    Cette fonction re-positionne le schemaRef si necessaire.
    Retourne le fichier corrige (ou l'original si deja OK).
    """
    raw = xbrl_file.read_bytes().decode("utf-8")

    # Trouver le tag d'ouverture <xbrli:xbrl ...>
    xbrl_open = re.search(r"<xbrli:xbrl\b[^>]*>", raw)
    if not xbrl_open:
        return xbrl_file  # Pas un XBRL standard, on laisse tel quel

    open_end = xbrl_open.end()

    # Trouver les schemaRef
    matches = list(_SCHEMA_REF_RE.finditer(raw))
    if not matches:
        return xbrl_file  # Pas de schemaRef du tout

    # Verifier si le premier schemaRef est deja juste apres l'ouverture
    first_match = matches[0]
    text_between = raw[open_end:first_match.start()].strip()
    # Si le texte entre l'ouverture et le schemaRef ne contient que des espaces
    # ou d'autres schemaRef, c'est OK
    if not text_between or text_between.startswith("<link:schemaRef") or text_between.startswith("<xbrli:context"):
        # Heuristic: if schemaRef is within first 5% of file, it's fine
        if first_match.start() < len(raw) * 0.05:
            return xbrl_file

    # Extraire tous les schemaRef et les supprimer de leur position actuelle
    schema_refs = [m.group(0) for m in matches]
    fixed = raw
    for m in reversed(matches):
        fixed = fixed[:m.start()] + fixed[m.end():]

    # Les re-inserer juste apres le tag d'ouverture <xbrli:xbrl ...>
    insert_pos = re.search(r"<xbrli:xbrl\b[^>]*>", fixed).end()
    schema_block = "\n      " + "\n      ".join(schema_refs)
    fixed = fixed[:insert_pos] + schema_block + fixed[insert_pos:]

    # Dedupliquer les <xbrli:context> avec le meme id AVANT sanitisation
    ctx_re = re.compile(
        r"<xbrli:context\s+id=\"([^\"]+)\"[^>]*>.*?</xbrli:context>",
        re.DOTALL,
    )
    seen_ctx: set[str] = set()
    dups_removed = 0

    def _dedup_ctx(m):
        nonlocal dups_removed
        cid = m.group(1)
        if cid in seen_ctx:
            dups_removed += 1
            return ""
        seen_ctx.add(cid)
        return m.group(0)

    fixed = ctx_re.sub(_dedup_ctx, fixed)
    if dups_removed:
        print(f"       [FIX] {dups_removed} contexte(s) duplique(s) supprime(s)")

    # Corriger les id/contextRef invalides (caracteres non-NCName)
    # NCName: lettres, chiffres, '.', '-', '_' ; debut par lettre ou '_'
    _id_counter: dict[str, int] = {}

    def _sanitize_xml_id(m):
        attr_part = m.group(1)  # 'id="' ou 'contextRef="'
        raw_val = m.group(2)
        val = re.sub(r"[^A-Za-z0-9._\-]", "_", raw_val)
        if val and not (val[0].isalpha() or val[0] == "_"):
            val = "x_" + val
        # Pour les id= (pas contextRef/unitRef), assurer l'unicite
        if "contextRef" not in attr_part and "unitRef" not in attr_part:
            n = _id_counter.get(val, 0)
            _id_counter[val] = n + 1
            if n > 0:
                val = f"{val}.{n}"
        return attr_part + val + '"'

    fixed = re.sub(
        r'((?:id|contextRef|unitRef)\s*=\s*")([^"]*)"',
        _sanitize_xml_id,
        fixed,
    )

    fixed_path = work_dir / f"{xbrl_file.stem}_fixed.xbrl"
    fixed_path.write_text(fixed, encoding="utf-8")
    print(f"       [FIX] XBRL pre-traite pour {xbrl_file.name}")
    return fixed_path


def convert_xbrl_to_csv(
    xbrl_file: Path,
    taxonomy_path: Path,
    converter_jar: Path,
    work_dir: Path,
) -> Path:
    """
    Convertit un fichier XBRL en CSV via Xml2Csv.jar.
    Retourne le dossier contenant les CSV extraits.
    """
    # Pre-traitement: repositionner schemaRef si mal place
    effective_xbrl = _fix_xbrl_schema_position(xbrl_file, work_dir)

    script_dir = converter_jar.parent
    script_output = script_dir / "output"
    script_output.mkdir(parents=True, exist_ok=True)

    zip_name = f"{effective_xbrl.stem}.zip"
    zip_path = script_output / zip_name
    if zip_path.exists():
        zip_path.unlink()
    # Aussi nettoyer le zip de l'original au cas ou
    zip_orig = script_output / f"{xbrl_file.stem}.zip"
    if zip_orig != zip_path and zip_orig.exists():
        zip_orig.unlink()

    cmd = ["java", "-jar", str(converter_jar), str(taxonomy_path), str(effective_xbrl)]
    res = subprocess.run(cmd, cwd=str(script_dir), capture_output=True, text=True)
    if res.returncode != 0:
        raise RuntimeError(
            f"Conversion XBRL->CSV echouee pour {xbrl_file.name}:\n{res.stderr[:500]}"
        )
    if not zip_path.exists():
        raise RuntimeError(f"ZIP converti introuvable: {zip_path}")

    out_dir = work_dir / "converted_submission_csv"
    out_dir.mkdir(parents=True, exist_ok=True)

    with zipfile.ZipFile(zip_path, "r") as zf:
        for name in zf.namelist():
            if name.lower().endswith(".csv"):
                target = out_dir / Path(name).name
                with zf.open(name) as src, target.open("wb") as dst:
                    shutil.copyfileobj(src, dst)

    return out_dir


def _extract_onegate_xml_to_zip(xml_path: Path, work_dir: Path) -> Path | None:
    """
    If *xml_path* is a OneGate XbrlDelivery envelope containing a <csv>
    element with base64-encoded ZIP data, decode it and write a temporary
    ZIP file.  Returns the path to the extracted ZIP, or None if the file
    is not a OneGate envelope.
    """
    import base64
    from xml.etree import ElementTree as ET

    try:
        tree = ET.parse(xml_path)
    except ET.ParseError:
        return None

    root = tree.getroot()
    # Namespace-agnostic search for a <csv> element
    csv_el = None
    for el in root.iter():
        tag = el.tag.split("}")[-1] if "}" in el.tag else el.tag
        if tag == "csv" and el.text and len(el.text.strip()) > 100:
            csv_el = el
            break
    if csv_el is None:
        return None

    try:
        raw = base64.b64decode(csv_el.text.strip())
    except Exception:
        return None

    # Check for PK (ZIP) magic bytes
    if len(raw) < 4 or raw[:2] != b"PK":
        return None

    zip_path = work_dir / f"{xml_path.stem}_extracted.zip"
    zip_path.write_bytes(raw)
    # Guard against false positives: payload can start with PK but still be invalid.
    if not zipfile.is_zipfile(zip_path):
        try:
            zip_path.unlink(missing_ok=True)
        except Exception:
            pass
        return None
    print(f"       [INFO] OneGate XML envelope detected — extracted ZIP ({len(raw)} bytes)")
    return zip_path


def _is_onegate_xml(xml_path: Path) -> bool:
    """Quick check: does the XML contain an XbrlDelivery or <csv> element?"""
    try:
        head = xml_path.read_bytes()[:2000].decode("utf-8", errors="ignore")
        return "XbrlDelivery" in head or "<csv>" in head
    except Exception:
        return False


def extract_xbrl_csv_zip(
    xbrl_zip: Path,
    work_dir: Path,
) -> Path:
    """
    Extrait les CSV d'un package XBRL-CSV (ZIP contenant reports/*.csv
    avec colonnes datapoint,factValue). Utilise par P3DH et similaires.
    Retourne le dossier contenant les CSV extraits.
    """
    out_dir = work_dir / "converted_submission_csv"
    out_dir.mkdir(parents=True, exist_ok=True)

    with zipfile.ZipFile(xbrl_zip, "r") as zf:
        for name in zf.namelist():
            if name.lower().endswith(".csv"):
                target = out_dir / Path(name).name
                with zf.open(name) as src, target.open("wb") as dst:
                    shutil.copyfileobj(src, dst)

    return out_dir


# ─────────────────────────────────────────────
#  2. Lecture taxonomie (JSON CellCode)
# ─────────────────────────────────────────────
def _iter_dicts(node):
    if isinstance(node, dict):
        yield node
        for v in node.values():
            yield from _iter_dicts(v)
    elif isinstance(node, list):
        for x in node:
            yield from _iter_dicts(x)


def load_taxonomy_datapoint_map(taxonomy_path: Path) -> dict[str, tuple[str, str, str, str]]:
    """
    Charge le mapping datapoint -> (table_code, row_id, column_id, cell_code)
    depuis TOUS les fichiers JSON de la taxonomie (tar.gz ou dossier).
    Fonctionne pour n'importe quel regulateur.
    """
    dp_map: dict[str, tuple[str, str, str, str]] = {}

    json_members: list = []

    if taxonomy_path.is_file() and taxonomy_path.name.endswith((".tar.gz", ".tgz")):
        with tarfile.open(taxonomy_path, "r:gz") as tf:
            for m in tf.getmembers():
                if m.name.lower().endswith(".json") and "/fws/" in m.name:
                    f = tf.extractfile(m)
                    if f:
                        try:
                            data = json.load(f)
                        except Exception:
                            continue
                        _extract_dp_from_json(data, dp_map)
    elif taxonomy_path.is_dir():
        for fp in sorted(taxonomy_path.rglob("*.json")):
            if "fws" in fp.parts:
                try:
                    data = json.loads(fp.read_text(encoding="utf-8"))
                except Exception:
                    continue
                _extract_dp_from_json(data, dp_map)
    else:
        raise FileNotFoundError(f"Taxonomy introuvable ou format non supporte: {taxonomy_path}")

    return dp_map


def _extract_dp_from_json(data: dict, dp_map: dict[str, tuple[str, str, str, str]]) -> None:
    for d in _iter_dicts(data):
        for k, v in d.items():
            if not (isinstance(k, str) and re.fullmatch(r"dp\d+", k, re.IGNORECASE)):
                continue
            if not isinstance(v, dict):
                continue
            doc = v.get("ad0:documentation") or v.get("eba:documentation") or {}
            cell = ""
            if isinstance(doc, dict):
                cell = str(doc.get("CellCode", "") or doc.get("cellcode", "") or "").strip()
            elif isinstance(doc, str):
                cell = doc
            if not cell:
                continue
            m = CELLCODE_RE.search(cell)
            if not m:
                continue
            table_code = m.group(1).strip().replace(" ", "_").upper()
            row_id = m.group(2).lower()
            column_id = m.group(3).lower()
            dp_map[k.lower()] = (table_code, row_id, column_id, cell)


# ─────────────────────────────────────────────
#  3. Lecture input CSV (format long + wide)
# ─────────────────────────────────────────────
WIDE_COL_RE = re.compile(r"^c\d{4}$", re.IGNORECASE)

def parse_input_csv(input_file: Path) -> list[InputRow]:
    rows: list[InputRow] = []
    fname = input_file.name
    with input_file.open("r", encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f, delimiter=";")
        if not reader.fieldnames:
            return rows
        # Strip spaces from headers (some CSV generators pad around ';')
        reader.fieldnames = [h.strip() for h in reader.fieldnames]
        fields_lc = {h.lower() for h in reader.fieldnames}
        required_long = {"tablecode", "rowid", "columnid", "value"}

        if required_long.issubset(fields_lc):
            # ── Format long: tablecode;rowid;columnid;value ──
            header_map = {h.lower(): h for h in reader.fieldnames}
            tc_key = header_map["tablecode"]
            ri_key = header_map["rowid"]
            ci_key = header_map["columnid"]
            v_key = header_map["value"]

            for i, row in enumerate(reader, start=2):
                table_code = str((row or {}).get(tc_key, "") or "").strip()
                row_id = str((row or {}).get(ri_key, "") or "").strip().lower()
                column_id = str((row or {}).get(ci_key, "") or "").strip().lower()
                value_raw = str((row or {}).get(v_key, "") or "").strip()
                if not table_code or not row_id or not column_id or value_raw == "":
                    continue
                rows.append(
                    InputRow(
                        input_line=i,
                        input_file=fname,
                        table_code=table_code,
                        row_id=row_id,
                        column_id=column_id,
                        value_raw=value_raw,
                    )
                )
        elif "tablecode" in fields_lc:
            # ── Format wide: tableCode;c0010;c0020;c0030;... ──
            header_map = {h.lower(): h for h in reader.fieldnames}
            tc_key = header_map["tablecode"]
            wide_cols = [h for h in reader.fieldnames if WIDE_COL_RE.match(h)]
            if not wide_cols:
                print(f"  [SKIP] {fname}: colonnes manquantes {required_long - fields_lc}")
                return rows
            for i, row in enumerate(reader, start=2):
                table_code = str((row or {}).get(tc_key, "") or "").strip()
                if not table_code:
                    continue
                for col_hdr in wide_cols:
                    value_raw = str((row or {}).get(col_hdr, "") or "").strip()
                    if value_raw == "":
                        continue
                    rows.append(
                        InputRow(
                            input_line=i,
                            input_file=fname,
                            table_code=table_code,
                            row_id="r999",
                            column_id=col_hdr.lower(),
                            value_raw=value_raw,
                        )
                    )
        else:
            print(f"  [SKIP] {fname}: colonnes manquantes {required_long - fields_lc}")
    return rows


def parse_input_folder(input_dir: Path) -> list[InputRow]:
    all_rows: list[InputRow] = []
    for fp in sorted(input_dir.glob("*.csv")):
        rows = parse_input_csv(fp)
        print(f"  Input: {fp.name} -> {len(rows)} valeur(s)")
        all_rows.extend(rows)
    return all_rows


# ─────────────────────────────────────────────
#  4. Lecture output CSV converti
# ─────────────────────────────────────────────
def parse_output_csv_with_taxonomy(
    converted_dir: Path,
    dp_map: dict[str, tuple[str, str, str, str]],
) -> tuple[list[OutputPoint], int]:
    points: list[OutputPoint] = []
    unknown_dp = 0

    for fp in sorted(converted_dir.glob("*.csv")):
        # Skip non-table files (FilingIndicators.csv, etc.)
        if fp.stem.lower() in ("filingindicators", "report_parameters", "parameters"):
            continue
        with fp.open("r", encoding="utf-8-sig", newline="") as f:
            # Some converters emit ';' instead of ','; detect delimiter from the header/sample.
            sample = f.read(4096)
            f.seek(0)
            delimiter = ","
            try:
                dialect = csv.Sniffer().sniff(sample, delimiters=",;\t")
                delimiter = dialect.delimiter
            except Exception:
                header_line = (sample.splitlines()[0] if sample else "").strip()
                if header_line.count(";") > header_line.count(","):
                    delimiter = ";"
            reader = csv.DictReader(f, delimiter=delimiter)
            if not reader.fieldnames or "datapoint" not in (h.lower() for h in reader.fieldnames):
                continue
            dp_col = next(h for h in reader.fieldnames if h.lower() == "datapoint")
            val_col = next((h for h in reader.fieldnames if h.lower() == "factvalue"), None)
            if not val_col:
                continue
            for row in reader:
                dp = str((row or {}).get(dp_col, "") or "").strip().lower()
                val = str((row or {}).get(val_col, "") or "").strip()
                if not dp or val == "":
                    continue
                meta = dp_map.get(dp)
                if not meta:
                    unknown_dp += 1
                    continue
                table_code, row_id, column_id, cell_code = meta
                points.append(
                    OutputPoint(
                        table_code=table_code,
                        row_id=row_id,
                        column_id=column_id,
                        datapoint=dp,
                        value_raw=val,
                        source_file=fp.name,
                        cell_code=cell_code,
                    )
                )

    return points, unknown_dp


# ─────────────────────────────────────────────
#  5. Auto-resolution des codes de table
# ─────────────────────────────────────────────
def _auto_resolve_one(input_code: str, known_codes: set[str]) -> str:
    """Resout un code de table input vers le code taxonomie le plus proche."""
    tc = input_code.upper()
    if tc in known_codes:
        return tc

    input_parts = tc.split(".")
    best_match = tc
    best_score = (-1, -1)

    for known in sorted(known_codes):
        known_parts = known.split(".")
        if len(known_parts) != len(input_parts):
            continue

        # Nombre de segments identiques depuis le debut
        prefix_len = 0
        for a, b in zip(input_parts, known_parts):
            if a == b:
                prefix_len += 1
            else:
                break

        # Bonus si le dernier segment correspond
        last_bonus = 0
        if input_parts[-1] == known_parts[-1]:
            last_bonus = 2  # match exact
        elif input_parts[-1].isdigit() and known_parts[-1].isdigit():
            if int(input_parts[-1]) % 10 == int(known_parts[-1]) % 10:
                last_bonus = 1  # match modulo 10 (ex: 11 -> 01)

        score = (prefix_len, last_bonus)
        if score > best_score:
            best_score = score
            best_match = known

    return best_match


def build_auto_aliases(
    input_codes: set[str], known_codes: set[str],
) -> dict[str, str]:
    """
    Construit automatiquement la table d'aliases input -> taxonomie.
    Ne retient que les codes qui different apres resolution.
    """
    aliases: dict[str, str] = {}
    for ic in sorted(input_codes):
        resolved = _auto_resolve_one(ic, known_codes)
        if resolved != ic:
            aliases[ic] = resolved
    return aliases


# ─────────────────────────────────────────────
#  6. Comparaison cellule par cellule
# ─────────────────────────────────────────────
def compare_cells(
    input_rows: list[InputRow],
    output_points: list[OutputPoint],
    known_taxonomy_codes: set[str],
) -> tuple[list[dict], dict[str, int]]:
    by_cell: dict[tuple[str, str, str], list[OutputPoint]] = defaultdict(list)
    for p in output_points:
        by_cell[(p.table_code, p.row_id, p.column_id)].append(p)

    # Auto-resolution des codes de table input -> taxonomie
    input_codes = {r.table_code.upper() for r in input_rows}
    aliases = build_auto_aliases(input_codes, known_taxonomy_codes)
    if aliases:
        print(f"       Auto-aliases detectes: {len(aliases)} correspondance(s)")
        for src, dst in sorted(aliases.items()):
            print(f"         {src} -> {dst}")

    out_pos: dict[tuple[str, str, str], int] = defaultdict(int)
    rows: list[dict] = []
    summary = {"OK": 0, "DIFF": 0, "MISSING_IN_OUTPUT": 0, "CALCULATED_OR_EXTRA": 0, "TOTAL": 0}

    for inp in input_rows:
        table_norm = aliases.get(inp.table_code.upper(), inp.table_code.upper())
        key = (table_norm, inp.row_id, inp.column_id)
        pts = by_cell.get(key, [])
        idx = out_pos[key]
        out = pts[idx] if idx < len(pts) else None
        in_norm = normalize_value(inp.value_raw)

        if out is None:
            status = "MISSING_IN_OUTPUT"
            details = "No output datapoint mapped to this taxonomy cell"
            out_dp = out_file = out_cell = out_val = out_norm = ""
        else:
            out_pos[key] += 1
            out_dp = out.datapoint
            out_file = out.source_file
            out_cell = out.cell_code
            out_val = out.value_raw
            out_norm = normalize_value(out.value_raw)
            if in_norm == out_norm:
                status = "OK"
                details = "Exact value match on same taxonomy cell"
            else:
                status = "DIFF"
                details = "Same taxonomy cell found, values differ"

        summary[status] += 1
        summary["TOTAL"] += 1
        rows.append(
            {
                "input_line": inp.input_line,
                "input_file": inp.input_file,
                "input_table_code": inp.table_code,
                "normalized_table_code": table_norm,
                "row_id_r": inp.row_id,
                "column_id_c": inp.column_id,
                "input_value": inp.value_raw,
                "input_value_norm": in_norm,
                "output_datapoint": out_dp,
                "output_source_file": out_file,
                "taxonomy_cell_code": out_cell,
                "output_value": out_val,
                "output_value_norm": out_norm,
                "status": status,
                "details": details,
            }
        )

    for key, pts in by_cell.items():
        used = out_pos.get(key, 0)
        for out in pts[used:]:
            summary["CALCULATED_OR_EXTRA"] += 1
            summary["TOTAL"] += 1
            rows.append(
                {
                    "input_line": "",
                    "input_file": "",
                    "input_table_code": f"(output: {out.table_code})",
                    "normalized_table_code": out.table_code,
                    "row_id_r": out.row_id,
                    "column_id_c": out.column_id,
                    "input_value": "",
                    "input_value_norm": "",
                    "output_datapoint": out.datapoint,
                    "output_source_file": out.source_file,
                    "taxonomy_cell_code": out.cell_code,
                    "output_value": out.value_raw,
                    "output_value_norm": normalize_value(out.value_raw),
                    "status": "CALCULATED_OR_EXTRA",
                    "details": f"Output-only: {out.cell_code} from {out.source_file}",
                }
            )

    rows.sort(
        key=lambda r: (
            r["input_line"] == "",
            int(r["input_line"] or 10**9),
            r["row_id_r"],
            r["column_id_c"],
        )
    )
    return rows, summary


# ─────────────────────────────────────────────
#  7. Generation rapport Excel + CSV
# ─────────────────────────────────────────────
REPORT_HEADERS = [
    "input_line",
    "input_file",
    "input_table_code",
    "normalized_table_code",
    "row_id_r",
    "column_id_c",
    "input_value",
    "input_value_norm",
    "output_datapoint",
    "output_source_file",
    "taxonomy_cell_code",
    "output_value",
    "output_value_norm",
    "status",
    "details",
]


def write_reports(output_dir: Path, rows: list[dict], summary: dict[str, int]) -> tuple[Path, Path]:
    xlsx_path = output_dir / "taxonomy_cell_report.xlsx"
    output_dir.mkdir(parents=True, exist_ok=True)

    wb = Workbook()
    ws = wb.active
    ws.title = "Details"
    ws.append(REPORT_HEADERS)
    for c in range(1, len(REPORT_HEADERS) + 1):
        cell = ws.cell(row=1, column=c)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = FILL_HEADER
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    status_col = REPORT_HEADERS.index("status") + 1

    for row in rows:
        ws.append([row.get(h, "") for h in REPORT_HEADERS])

    for excel_row in range(2, ws.max_row + 1):
        status_val = str(ws.cell(row=excel_row, column=status_col).value or "")
        fill = STATUS_FILL.get(status_val)
        if fill:
            ws.cell(row=excel_row, column=status_col).fill = fill

    for col in ws.columns:
        max_len = max(len(str(cell.value or "")) for cell in col)
        ws.column_dimensions[col[0].column_letter].width = min(max(12, max_len + 2), 80)

    ws2 = wb.create_sheet("Summary")
    ws2.append(["Metric", "Count"])
    for k in ["OK", "DIFF", "MISSING_IN_OUTPUT", "CALCULATED_OR_EXTRA", "TOTAL"]:
        ws2.append([k, summary.get(k, 0)])
    ws2.column_dimensions["A"].width = 28
    ws2.column_dimensions["B"].width = 14

    wb.save(xlsx_path)
    return xlsx_path, xlsx_path


# ─────────────────────────────────────────────
#  7b. Support MCONTRAN (OneGate XML)
# ─────────────────────────────────────────────
# Mapping CSV column (suffix after 'frbdfruba') -> XML Dim prop
# Extracted from pure-regulatoryOutput-frbdfruba-bdfmcontran-submission.yml
_MCONTRAN_COL_TO_PROP: dict[str, str] = {
    "frbdfrubaGuichet": "ID_GUI",
    "frbdfrubaContractNumber": "RFLICR",
    "frbdfrubaTypeOfInstrument": "INS_FI",
    "frbdfrubaNewBusinessAmount": "MT_CRDT",
    "frbdfrubaMaximumAmountAllowed": "MT_MAX",
    "frbdfrubaOriginalMaturityInMonths": "DUREE_IN",
    "frbdfrubaRenegotiatedLoan": "CDT_NGCT",
    "frbdfrubaReferenceIndex": "IDX_REF",
    "frbdfrubaInitialPeriodOfInterestRateFixation": "PFIT",
    "frbdfrubaInterestRate": "TESE",
    "frbdfrubaAnnualPercentageRateOfCharge": "TEG",
    "frbdfrubaInterestRateCap": "CAP",
    "frbdfrubaPaymentAdjustmentMode": "AJUST",
    "frbdfrubaRegulatedLoan": "PRT_RGLT",
    "frbdfrubaRealEstateLoanPurpose": "PRT_RLS",
    "frbdfrubaRepurchasedLoan": "PRT_RSTR",
    "frbdfrubaResidenceZone": "ZONE_RD",
    "frbdfrubaPaymentAmount": "MT_REMBRST",
    "frbdfrubaPaymentFrequency": "PERIOD_RBRST",
    "frbdfrubaTypeOfProtection": "SURETE",
    "frbdfrubaAnnualIncome": "REVENU_ANN",
    "frbdfrubaPoolShare": "PRT_POOL",
    "frbdfrubaSiren": "SIREN",
    "frbdfrubaUsageOfCredit": "USG_PRT",
    "frbdfrubaOverdraftFeesRate": "TX_COMM_DEC",
}

# Props with BDFRATE6 formatting (remove decimal dot)
_BDFRATE6_PROPS = {"TESE", "TEG", "CAP", "TX_COMM_DEC"}


def _format_bdfrate6(csv_val: str) -> str:
    """Convert CSV rate '06.5942' to BDFRATE6 format '065942' (remove dot)."""
    return csv_val.replace(".", "")


def parse_onegate_xml(xml_path: Path) -> list[dict[str, str]]:
    """
    Parse OneGate XML submission into a list of item dicts.
    Each item dict: {'_form': 'MCO1', '_idx': '0', 'ID_GUI': '90001', ...}
    """
    import xml.etree.ElementTree as ET
    tree = ET.parse(xml_path)
    root = tree.getroot()
    ns_uri = root.tag.split("}")[0].lstrip("{") if "}" in root.tag else ""
    ns = {"og": ns_uri} if ns_uri else {}
    prefix = f"{{{ns_uri}}}" if ns_uri else ""

    items: list[dict[str, str]] = []
    for data_el in root.iter(f"{prefix}Data"):
        form = data_el.get("form", "")
        idx = 0
        for item_el in data_el.iter(f"{prefix}Item"):
            item: dict[str, str] = {"_form": form, "_idx": str(idx)}
            for dim_el in item_el.iter(f"{prefix}Dim"):
                prop = dim_el.get("prop", "")
                val = (dim_el.text or "").strip()
                if prop:
                    item[prop] = val
            items.append(item)
            idx += 1
    return items


def run_mcontran(
    input_csv: Path,
    xml_file: Path,
    output_dir: Path,
) -> dict[str, int]:
    """
    Compare MCONTRAN input CSV vs OneGate XML expected submission.
    Returns summary dict compatible with the batch system.
    """
    run_dir = output_dir
    run_dir.mkdir(parents=True, exist_ok=True)

    # 1. Parse XML items
    print("[1/3] Lecture XML OneGate ...")
    xml_items = parse_onegate_xml(xml_file)
    print(f"       {len(xml_items)} item(s) dans {len(set(it['_form'] for it in xml_items))} form(s)")

    # 2. Parse input CSV
    print("[2/3] Lecture CSV input ...")
    csv_rows: list[dict[str, str]] = []
    with input_csv.open("r", encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f, delimiter=";")
        for row in reader:
            csv_rows.append({k: (v or "").strip() for k, v in row.items()})
    print(f"       {len(csv_rows)} ligne(s) input")

    # 3. Build mapped CSV items: convert each CSV row to {prop: value}
    csv_items: list[dict[str, str]] = []
    for csv_row in csv_rows:
        mapped: dict[str, str] = {}
        for col, prop in _MCONTRAN_COL_TO_PROP.items():
            val = csv_row.get(col, "")
            if val and prop in _BDFRATE6_PROPS:
                val = _format_bdfrate6(val)
            if val:
                mapped[prop] = val
        csv_items.append(mapped)

    # 4. Match CSV items to XML items using greedy best-match
    print("[3/3] Comparaison ...")
    xml_available = list(range(len(xml_items)))  # indices of available XML items
    matches: list[tuple[int, int, int, int]] = []  # (csv_idx, xml_idx, ok_count, diff_count)

    for ci, csv_item in enumerate(csv_items):
        best_xml_idx = -1
        best_ok = -1
        best_diff = 999999
        for xi in xml_available:
            xml_item = xml_items[xi]
            ok = 0
            diff = 0
            for prop, csv_val in csv_item.items():
                xml_val = xml_item.get(prop, "")
                if csv_val == xml_val:
                    ok += 1
                elif xml_val:
                    diff += 1
            if ok > best_ok or (ok == best_ok and diff < best_diff):
                best_ok = ok
                best_diff = diff
                best_xml_idx = xi
        if best_xml_idx >= 0:
            xml_available.remove(best_xml_idx)
            matches.append((ci, best_xml_idx, best_ok, best_diff))

    # 5. Build report rows
    report_rows: list[dict] = []
    summary = {"OK": 0, "DIFF": 0, "MISSING_IN_OUTPUT": 0, "CALCULATED_OR_EXTRA": 0, "TOTAL": 0}

    for ci, xi, ok_count, diff_count in matches:
        csv_item = csv_items[ci]
        xml_item = xml_items[xi]
        form = xml_item.get("_form", "?")
        xml_idx = xml_item.get("_idx", "?")
        all_props = sorted(set(list(csv_item.keys()) + [p for p in xml_item if not p.startswith("_")]))

        for prop in all_props:
            csv_val = csv_item.get(prop, "")
            xml_val = xml_item.get(prop, "")
            if prop == "SCT":
                continue  # SCT is the section code, not a data field
            if csv_val and xml_val:
                if csv_val == xml_val:
                    status = "OK"
                    summary["OK"] += 1
                else:
                    status = "DIFF"
                    summary["DIFF"] += 1
            elif csv_val and not xml_val:
                status = "MISSING_IN_OUTPUT"
                summary["MISSING_IN_OUTPUT"] += 1
            elif xml_val and not csv_val:
                status = "CALCULATED_OR_EXTRA"
                summary["CALCULATED_OR_EXTRA"] += 1
            else:
                continue  # both empty
            summary["TOTAL"] += 1
            report_rows.append({
                "form": form,
                "xml_item": xml_idx,
                "prop": prop,
                "input_value": csv_val,
                "output_value": xml_val,
                "status": status,
            })

    # Unmatched XML items
    for xi in xml_available:
        xml_item = xml_items[xi]
        form = xml_item.get("_form", "?")
        xml_idx = xml_item.get("_idx", "?")
        for prop, val in xml_item.items():
            if prop.startswith("_") or prop == "SCT" or not val:
                continue
            summary["CALCULATED_OR_EXTRA"] += 1
            summary["TOTAL"] += 1
            report_rows.append({
                "form": form,
                "xml_item": xml_idx,
                "prop": prop,
                "input_value": "",
                "output_value": val,
                "status": "CALCULATED_OR_EXTRA",
            })

    # 6. Write reports
    xlsx_report = run_dir / "taxonomy_cell_report.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "MCONTRAN Report"
    headers = ["Form", "XML Item", "Prop", "Input Value", "Output Value", "Status"]
    for ci, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=ci, value=h)
        cell.fill = FILL_HEADER
        cell.font = Font(color="FFFFFF", bold=True)
    for ri, row in enumerate(report_rows, 2):
        ws.cell(row=ri, column=1, value=row["form"])
        ws.cell(row=ri, column=2, value=row["xml_item"])
        ws.cell(row=ri, column=3, value=row["prop"])
        ws.cell(row=ri, column=4, value=row["input_value"])
        ws.cell(row=ri, column=5, value=row["output_value"])
        status_cell = ws.cell(row=ri, column=6, value=row["status"])
        fill = STATUS_FILL.get(row["status"])
        if fill:
            status_cell.fill = fill
    for col in ws.columns:
        max_len = max(len(str(cell.value or "")) for cell in col)
        ws.column_dimensions[col[0].column_letter].width = min(max(12, max_len + 2), 80)
    wb.save(xlsx_report)

    print()
    print(f"[OK] Input:     {input_csv}")
    print(f"[OK] XML:       {xml_file}")
    print(f"[OK] Output:    {run_dir}")
    print(f"[OK] XLSX:      {xlsx_report}")
    print()
    print(
        f"[SUMMARY] OK={summary['OK']} | DIFF={summary['DIFF']} | "
        f"MISSING_IN_OUTPUT={summary['MISSING_IN_OUTPUT']} | "
        f"CALCULATED_OR_EXTRA={summary['CALCULATED_OR_EXTRA']} | "
        f"TOTAL={summary['TOTAL']}"
    )
    return summary


# ─────────────────────────────────────────────
#  8. Orchestration principale
# ─────────────────────────────────────────────
def run(
    input_dir: Path,
    xbrl_file: Path,
    taxonomy_path: Path,
    output_dir: Path,
) -> int:
    run_dir = output_dir
    run_dir.mkdir(parents=True, exist_ok=True)

    # Etape 1: Conversion XBRL -> CSV (ou extraction directe si XBRL-CSV ZIP)
    # Handle OneGate XML envelopes (base64-encoded ZIP inside <csv> element)
    effective_file = xbrl_file
    onegate_tmp = None
    if xbrl_file.suffix.lower() == ".xml":
        _og_tmp = tempfile.mkdtemp(prefix="onegate_")
        extracted_zip = _extract_onegate_xml_to_zip(xbrl_file, Path(_og_tmp))
        if extracted_zip is not None:
            effective_file = extracted_zip
            onegate_tmp = _og_tmp

    is_xbrl_csv_zip = effective_file.suffix.lower() == ".zip"
    if is_xbrl_csv_zip:
        print(f"[1/4] Extraction XBRL-CSV depuis ZIP ...")
        converted_dir = run_dir / "converted_submission_csv"
        converted_dir.mkdir(parents=True, exist_ok=True)
        try:
            with tempfile.TemporaryDirectory(prefix="xbrl_csv_") as td:
                extracted = extract_xbrl_csv_zip(effective_file, Path(td))
                for f in extracted.glob("*.csv"):
                    shutil.copy2(f, converted_dir / f.name)
        except zipfile.BadZipFile:
            # Fallback for malformed/ambiguous XML payloads: process as standard XML submission.
            print("       [WARN] ZIP invalide detecte, fallback vers conversion XML standard")
            with tempfile.TemporaryDirectory(prefix="xbrl_convert_") as td:
                produced_zip = convert_xml_to_zip(xbrl_file, taxonomy_path, td, "convert_only", td)
                with zipfile.ZipFile(produced_zip, "r") as zip_ref:
                    for member in zip_ref.namelist():
                        if member.endswith(".csv"):
                            zip_ref.extract(member, converted_dir)
    else:
        print(f"[1/4] Conversion XBRL -> CSV ...")
        with tempfile.TemporaryDirectory(prefix="xbrl_convert_") as td:
            # Conversion Python pure
            produced_zip = convert_xml_to_zip(effective_file, taxonomy_path, td, "convert_only", td)
            converted_dir = run_dir / "converted_submission_csv"
            converted_dir.mkdir(parents=True, exist_ok=True)
            # Extraction des CSV du ZIP produit
            with zipfile.ZipFile(produced_zip, "r") as zip_ref:
                for member in zip_ref.namelist():
                    if member.endswith(".csv"):
                        zip_ref.extract(member, converted_dir)

    # Cleanup OneGate temp directory
    if onegate_tmp:
        shutil.rmtree(onegate_tmp, ignore_errors=True)

    # Etape 2: Charger taxonomie
    print(f"[2/4] Chargement mapping taxonomie ...")
    dp_map = load_taxonomy_datapoint_map(taxonomy_path)
    print(f"       {len(dp_map)} datapoint(s) mappes")

    # Etape 3: Lecture input + output + comparaison
    print(f"[3/4] Lecture et comparaison ...")
    output_points, unknown_dp = parse_output_csv_with_taxonomy(converted_dir, dp_map)
    print(f"       {len(output_points)} fait(s) output mappes, {unknown_dp} sans mapping")

    if input_dir.is_file():
        input_rows = parse_input_csv(input_dir)
        print(f"  Input: {input_dir.name} -> {len(input_rows)} valeur(s)")
    else:
        input_rows = parse_input_folder(input_dir)
    print(f"       {len(input_rows)} valeur(s) input lues")

    known_taxonomy_codes = {v[0] for v in dp_map.values()}
    rows, summary = compare_cells(input_rows, output_points, known_taxonomy_codes)

    # Etape 4: Rapport
    print(f"[4/4] Generation rapport ...")
    _, xlsx_report = write_reports(run_dir, rows, summary)

    print()
    print(f"[OK] Input:     {input_dir}")
    print(f"[OK] XBRL:      {xbrl_file}")
    print(f"[OK] Taxonomy:  {taxonomy_path}")
    # Plus de converter_jar
    print(f"[OK] Output:    {run_dir}")
    print(f"[OK] XLSX:      {xlsx_report}")
    print()
    print(
        f"[SUMMARY] OK={summary['OK']} | DIFF={summary['DIFF']} | "
        f"MISSING_IN_OUTPUT={summary['MISSING_IN_OUTPUT']} | "
        f"CALCULATED_OR_EXTRA={summary['CALCULATED_OR_EXTRA']} | "
        f"TOTAL={summary['TOTAL']}"
    )
    return summary

# ─────────────────────────────────────────────
#  9. Conversion Excel -> CSV (format long)
# ─────────────────────────────────────────────
def convert_excel_to_csv(excel_dir: Path, csv_output_dir: Path) -> None:
    """
    Convertit tous les .xlsx d'un dossier en CSV (format long: tableCode;rowId;columnId;value).
    Chaque feuille (sauf 'identification') est traitee comme une table.
    Layout: row B2 = 'Table Code', B2's right cell = tableCode.
    Row headers (rowId) in column B starting from row where data begins.
    Column headers (columnId) in the row just above data columns.
    """
    csv_output_dir.mkdir(parents=True, exist_ok=True)
    for xlsx_file in sorted(excel_dir.glob("*.xlsx")):
        wb = load_workbook(xlsx_file, data_only=True)
        all_rows: list[dict] = []
        for sn in wb.sheetnames:
            if sn.lower() == "identification":
                continue
            ws = wb[sn]
            # Find table code: look for cell with 'Table Code' text
            table_code = sn  # fallback
            header_row_idx = None
            for r in range(1, min(ws.max_row + 1, 20)):
                for c in range(1, min(ws.max_column + 1, 20)):
                    val = ws.cell(row=r, column=c).value
                    if isinstance(val, str) and val.strip().lower() == "table code":
                        tc_val = ws.cell(row=r, column=c + 1).value
                        if tc_val:
                            table_code = str(tc_val).strip()
                        break

            # Find column header row: row with numeric column IDs like '0010', '0020'
            col_id_row = None
            col_map: dict[int, str] = {}  # excel_col -> columnId
            for r in range(1, min(ws.max_row + 1, 20)):
                found_cols = {}
                for c in range(3, ws.max_column + 1):
                    val = ws.cell(row=r, column=c).value
                    if val is not None:
                        s = str(val).strip()
                        if re.fullmatch(r"\d{4}", s):
                            found_cols[c] = s
                if len(found_cols) >= 1:
                    col_id_row = r
                    col_map = found_cols
                    break

            if col_id_row is None:
                continue

            # Data rows start after col_id_row; rowId is in column B
            for r in range(col_id_row + 1, ws.max_row + 1):
                row_id_val = ws.cell(row=r, column=2).value
                if row_id_val is None:
                    continue
                row_id_str = str(row_id_val).strip()
                if not re.fullmatch(r"\d{4}", row_id_str):
                    continue
                row_id = f"r{row_id_str}"
                for col_idx, col_id_str in col_map.items():
                    cell_val = ws.cell(row=r, column=col_idx).value
                    val_str = "" if cell_val is None else str(cell_val).strip()
                    all_rows.append({
                        "tableCode": table_code,
                        "rowId": row_id,
                        "columnId": f"c{col_id_str}",
                        "value": val_str,
                    })

        if all_rows:
            csv_name = xlsx_file.stem + "_converted.csv"
            csv_path = csv_output_dir / csv_name
            with csv_path.open("w", encoding="utf-8", newline="") as f:
                w = csv.DictWriter(f, fieldnames=["tableCode", "rowId", "columnId", "value"], delimiter=";")
                w.writeheader()
                w.writerows(all_rows)
            print(f"  Excel->CSV: {xlsx_file.name} -> {csv_name} ({len(all_rows)} cellules)")


# ─────────────────────────────────────────────
#  10. Detection automatique des cas de test (batch)
# ─────────────────────────────────────────────
@dataclass
class TestCase:
    case_id: str
    input_dir: Path
    expected_file: Path


def discover_test_cases(
    load_dir: Path,
    expected_dir: Path,
    input_mode: str = "csv_direct",
) -> list[TestCase]:
    """
    Detecte automatiquement les cas de test en matchant input/load/*
    avec expected/*_submission.{xbrl,zip}.

    Gere trois patterns:
    - Pattern A (sous-dossiers): input/load/RUBA1/csv_direct/ <-> expected/RUBA1_*_submission.xbrl
    - Pattern B (fichiers plats): input/load/csv_direct/ <-> expected/P3DHEU1_*_submission.zip
    - Pattern C (MCONTRAN): input/load/mcontran_INSTMCO1_OK.csv <-> expected/INSTMCO1_submission.xml
    """
    # Collecter tous les fichiers expected submission
    expected_files: list[Path] = []
    for f in sorted(expected_dir.iterdir()):
        if f.is_file() and "submission" in f.stem.lower():
            if f.suffix.lower() in (".xbrl", ".zip", ".xml"):
                # Exclure les *_dataset.xbrl
                if "_dataset" in f.stem.lower():
                    continue
                expected_files.append(f)

    if not expected_files:
        print(f"[WARN] Aucun fichier *submission.xbrl/.zip trouve dans {expected_dir}")
        return []

    cases: list[TestCase] = []

    # Detecter le pattern: est-ce que input/load/ contient csv_direct/excel directement?
    has_flat_input = (load_dir / input_mode).is_dir()
    subfolders = [d for d in sorted(load_dir.iterdir()) if d.is_dir() and d.name not in ("csv_direct", "excel")]
    has_case_subfolders = any((d / input_mode).is_dir() for d in subfolders)

    if has_case_subfolders:
        # Pattern A: sous-dossiers par cas (RUBA, COREP, DBM)
        all_folder_names = [d.name for d in subfolders if (d / input_mode).is_dir()]
        for case_dir in subfolders:
            input_sub = case_dir / input_mode
            if not input_sub.is_dir():
                continue
            folder_name = case_dir.name.upper()  # ex: RUBA1, BE1, EU1, FR1

            # Trouver le fichier expected correspondant
            matched = _match_expected_file(folder_name, expected_files, all_folder_names)
            if matched:
                cases.append(TestCase(
                    case_id=matched.stem.split("_")[0],
                    input_dir=input_sub,
                    expected_file=matched,
                ))
            else:
                print(f"  [SKIP] Pas de fichier expected pour {folder_name}")

    elif has_flat_input:
        # Pattern B: fichiers plats dans csv_direct (P3DH)
        flat_dir = load_dir / input_mode
        for ef in expected_files:
            case_id = ef.stem.split("_")[0]  # ex: P3DHEU1
            # Filtrer les fichiers input contenant le case_id
            matching_csvs = [f for f in flat_dir.glob("*.csv") if case_id.upper() in f.name.upper()]
            matching_xlsxs = [f for f in flat_dir.glob("*.xlsx") if case_id.upper() in f.name.upper()]
            if matching_csvs or matching_xlsxs:
                cases.append(TestCase(
                    case_id=case_id,
                    input_dir=flat_dir,  # on filtrera les fichiers au moment du run
                    expected_file=ef,
                ))

    # Pattern C: MCONTRAN flat CSV files directly in load_dir
    # input/load/mcontran_INSTMCO1_OK.csv <-> expected/INSTMCO1_submission.xml
    if not cases:
        flat_csvs = sorted(load_dir.glob("*.csv"))
        xml_expected = [ef for ef in expected_files if ef.suffix.lower() == ".xml"]
        if flat_csvs and xml_expected:
            for ef in xml_expected:
                case_id = ef.stem.split("_")[0]  # ex: INSTMCO1
                matching = [f for f in flat_csvs if case_id.upper() in f.name.upper()]
                if matching:
                    cases.append(TestCase(
                        case_id=case_id,
                        input_dir=matching[0],  # single CSV file as Path
                        expected_file=ef,
                    ))

    return cases


def _match_expected_file(folder_name: str, expected_files: list[Path], all_folder_names: list[str]) -> Path | None:
    """
    Matche un nom de dossier input (RUBA1, BE1, EU1...) au fichier expected.
    Strategies:
    1. Match direct: RUBA1 -> RUBA1_*_submission.*
    2. Suffixe pays: BE1 -> *BE1_*_submission.*
    3. EU->defaut: EU1 -> le fichier qui n'est matche par aucun autre dossier
    """
    folder_upper = folder_name.upper()

    # 1. Match direct: le fichier commence par le nom du dossier
    for ef in expected_files:
        if ef.stem.upper().startswith(folder_upper + "_"):
            return ef

    # 2. Le case_id dans expected se termine par le nom du dossier
    for ef in expected_files:
        case_id = ef.stem.split("_")[0].upper()
        if case_id.endswith(folder_upper):
            return ef

    # 3. Dossier EUx -> trouver le fichier expected qui n'est matche par aucun autre dossier
    eu_match = re.fullmatch(r"EU(\d+)", folder_upper)
    if eu_match:
        other_folders = [f.upper() for f in all_folder_names if f.upper() != folder_upper]
        for ef in expected_files:
            case_id = ef.stem.split("_")[0].upper()
            # Ce fichier est-il deja matche par un autre dossier?
            claimed = False
            for of in other_folders:
                if case_id.startswith(of + "_") or case_id.endswith(of):
                    claimed = True
                    break
            if not claimed:
                return ef

    return None


# ─────────────────────────────────────────────
#  11. Mode batch: execution de tous les cas
# ─────────────────────────────────────────────
def run_batch(
    load_dir: Path,
    expected_dir: Path,
    taxonomy_path: Path,
    converter_jar: Path | None = None,
    output_dir: Path = Path("."),
    input_mode: str = "csv_direct",
) -> int:
    """
    Execute tous les cas de test detectes automatiquement.
    Genere un rapport individuel par cas + un rapport global.
    """
    print("=" * 70)
    print("  MODE BATCH - Detection automatique des cas de test")
    print("=" * 70)
    print(f"  Load dir:     {load_dir}")
    print(f"  Expected dir: {expected_dir}")
    print(f"  Taxonomy:     {taxonomy_path}")
    print(f"  Input mode:   {input_mode}")
    if converter_jar:
        print(f"  Converter:    {converter_jar}")
    print()

    cases = discover_test_cases(load_dir, expected_dir, input_mode)

    if not cases:
        print("[ERREUR] Aucun cas de test detecte.")
        return 1

    print(f"  {len(cases)} cas de test detecte(s):")
    for tc in cases:
        print(f"    - {tc.case_id}: input={tc.input_dir.relative_to(load_dir)}, expected={tc.expected_file.name}")
    print()

    batch_dir = output_dir
    batch_dir.mkdir(parents=True, exist_ok=True)

    all_summaries: list[tuple[str, dict[str, int]]] = []
    failed_cases: list[tuple[str, str]] = []

    for i, tc in enumerate(cases, 1):
        print("─" * 70)
        print(f"  CAS {i}/{len(cases)}: {tc.case_id}")
        print("─" * 70)

        case_output = batch_dir / tc.case_id

        # Si mode Excel, convertir d'abord en CSV
        actual_input_dir = tc.input_dir
        temp_csv_dir = None
        if input_mode == "excel":
            temp_csv_dir = case_output / "_excel_converted"
            temp_csv_dir.mkdir(parents=True, exist_ok=True)
            convert_excel_to_csv(tc.input_dir, temp_csv_dir)
            actual_input_dir = temp_csv_dir

        # Si pattern flat (P3DH), filtrer les fichiers par case_id
        if tc.input_dir.name in ("csv_direct", "excel") and input_mode == "csv_direct":
            matching_files = [f for f in tc.input_dir.glob("*.csv")
                             if tc.case_id.upper() in f.name.upper()]
            if len(matching_files) == 1:
                actual_input_dir = matching_files[0]  # single file path

        try:
            # Route .xml files: check if it's a OneGate envelope (with base64 ZIP)
            # or a plain MCONTRAN XML
            if tc.expected_file.suffix.lower() == ".xml":
                # Try OneGate extraction first
                import tempfile as _tmpmod
                _og_test_dir = _tmpmod.mkdtemp(prefix="og_test_")
                _og_zip = _extract_onegate_xml_to_zip(tc.expected_file, Path(_og_test_dir))
                shutil.rmtree(_og_test_dir, ignore_errors=True)

                if _og_zip is not None or _is_onegate_xml(tc.expected_file):
                    # OneGate envelope -> treat like .xbrl/.zip via run()
                    summary = run(
                        input_dir=actual_input_dir,
                        xbrl_file=tc.expected_file,
                        taxonomy_path=taxonomy_path,
                        output_dir=case_output,
                    )
                else:
                    # Plain MCONTRAN XML
                    input_csv = actual_input_dir if actual_input_dir.is_file() else None
                    if input_csv is None:
                        csvs = sorted(actual_input_dir.glob("*.csv"))
                        input_csv = csvs[0] if csvs else None
                    if input_csv is None:
                        raise FileNotFoundError(f"Aucun CSV input pour {tc.case_id}")
                    summary = run_mcontran(
                        input_csv=input_csv,
                        xml_file=tc.expected_file,
                        output_dir=case_output,
                    )
            else:
                summary = run(
                    input_dir=actual_input_dir,
                    xbrl_file=tc.expected_file,
                    taxonomy_path=taxonomy_path,
                    output_dir=case_output,
                )
            all_summaries.append((tc.case_id, summary))
        except Exception as e:
            print(f"  [ERREUR] {tc.case_id}: {e}")
            failed_cases.append((tc.case_id, str(e)))
        print()

    # Rapport global
    print("=" * 70)
    print("  RAPPORT GLOBAL BATCH")
    print("=" * 70)

    global_summary = {"OK": 0, "DIFF": 0, "MISSING_IN_OUTPUT": 0, "CALCULATED_OR_EXTRA": 0, "TOTAL": 0}
    global_rows: list[dict] = []

    for case_id, summary in all_summaries:
        for k in global_summary:
            global_summary[k] += summary.get(k, 0)
        print(
            f"  {case_id:20s}  OK={summary['OK']:>5} | DIFF={summary['DIFF']:>5} | "
            f"MISSING={summary['MISSING_IN_OUTPUT']:>5} | EXTRA={summary['CALCULATED_OR_EXTRA']:>5} | "
            f"TOTAL={summary['TOTAL']:>5}"
        )

    for case_id, err in failed_cases:
        print(f"  {case_id:20s}  ECHEC: {err[:60]}")

    print("  " + "─" * 68)
    print(
        f"  {'GLOBAL':20s}  OK={global_summary['OK']:>5} | DIFF={global_summary['DIFF']:>5} | "
        f"MISSING={global_summary['MISSING_IN_OUTPUT']:>5} | EXTRA={global_summary['CALCULATED_OR_EXTRA']:>5} | "
        f"TOTAL={global_summary['TOTAL']:>5}"
    )
    print()

    print(f"[OK] Batch output: {batch_dir}")

    return 0 if not failed_cases else 1

def main() -> int:
    parser = argparse.ArgumentParser(
        description="Outil general de validation reglementaire par mapping taxonomie",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Mode simple (un cas):
  python validate_regulation_unified.py \\
    --input-dir  .../input/load/RUBA1/csv_direct \\
    --xbrl-file  .../expected/RUBA1_rubad_submission.xbrl \\
    --taxonomy   .../xbrl/xbrl_package.tar.gz \\
    --converter-jar .../Xml2Csv.jar \\
    --output-dir ./user_reports/output

Mode batch (tous les cas d'un coup):
  python validate_regulation_unified.py --batch \\
    --load-dir   .../e2e-test-data/input/load \\
    --expected-dir .../e2e-test-data/expected \\
    --taxonomy   .../xbrl/xbrl_package.tar.gz \\
    --converter-jar .../Xml2Csv.jar \\
    --output-dir ./user_reports/output

Mode batch Excel:
  python validate_regulation_unified.py --batch --input-mode excel \\
    --load-dir   .../e2e-test-data/input/load \\
    --expected-dir .../e2e-test-data/expected \\
    --taxonomy   .../xbrl/xbrl_package.tar.gz \\
    --converter-jar .../Xml2Csv.jar \\
    --output-dir ./user_reports/output
""",
    )

    # Mode batch
    parser.add_argument(
        "--batch",
        action="store_true",
        help="Mode batch: detecte et execute tous les cas de test automatiquement",
    )
    parser.add_argument(
        "--load-dir",
        type=Path,
        default=None,
        help="[batch] Dossier input/load/ contenant les sous-dossiers par cas",
    )
    parser.add_argument(
        "--expected-dir",
        type=Path,
        default=None,
        help="[batch] Dossier expected/ contenant les fichiers submission",
    )
    parser.add_argument(
        "--input-mode",
        choices=["csv_direct", "excel"],
        default="csv_direct",
        help="[batch] Type d'input: csv_direct (defaut) ou excel",
    )

    # Mode simple
    parser.add_argument(
        "--input-dir",
        type=Path,
        default=None,
        help="[simple] Dossier/fichier CSV input",
    )
    parser.add_argument(
        "--xbrl-file",
        type=Path,
        default=None,
        help="[simple] Fichier XBRL submission (.xbrl ou .zip)",
    )

    # Communs
    parser.add_argument(
        "--taxonomy",
        type=Path,
        default=None,
        help="Package taxonomie (xbrl_package.tar.gz ou dossier extrait). Optionnel pour MCONTRAN.",
    )
    parser.add_argument(
        "--converter-jar",
        type=Path,
        default=None,
        help="Chemin vers Xml2Csv.jar (requis pour XBRL-XML)",
    )
    parser.add_argument(
        "--output-dir",
        type=Path,
        required=True,
        help="Dossier racine pour les rapports generes",
    )
    args = parser.parse_args()

    # Validations communes
    if args.taxonomy is not None and not args.taxonomy.exists():
        raise SystemExit(f"[ERREUR] Taxonomie introuvable: {args.taxonomy}")
    if args.converter_jar is not None and not args.converter_jar.is_file():
        raise SystemExit(f"[ERREUR] Xml2Csv.jar introuvable: {args.converter_jar}")

    if args.batch:
        # Mode batch
        if not args.load_dir or not args.expected_dir:
            raise SystemExit("[ERREUR] --load-dir et --expected-dir requis en mode --batch")
        if not args.load_dir.is_dir():
            raise SystemExit(f"[ERREUR] Dossier load introuvable: {args.load_dir}")
        if not args.expected_dir.is_dir():
            raise SystemExit(f"[ERREUR] Dossier expected introuvable: {args.expected_dir}")

        return run_batch(
            load_dir=args.load_dir,
            expected_dir=args.expected_dir,
            taxonomy_path=args.taxonomy,
            converter_jar=args.converter_jar,
            output_dir=args.output_dir,
            input_mode=args.input_mode,
        )
    else:
        # Mode simple
        if not args.input_dir or not args.xbrl_file:
            raise SystemExit("[ERREUR] --input-dir et --xbrl-file requis en mode simple (sans --batch)")
        if not args.input_dir.is_dir() and not args.input_dir.is_file():
            raise SystemExit(f"[ERREUR] Dossier/fichier input introuvable: {args.input_dir}")
        if not args.xbrl_file.is_file():
            raise SystemExit(f"[ERREUR] Fichier XBRL introuvable: {args.xbrl_file}")

        result = run(
            input_dir=args.input_dir,
            xbrl_file=args.xbrl_file,
            taxonomy_path=args.taxonomy,
            converter_jar=args.converter_jar,
            output_dir=args.output_dir,
        )
        return 0 if isinstance(result, dict) else result


if __name__ == "__main__":
    raise SystemExit(main())
