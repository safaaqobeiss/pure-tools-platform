#!/usr/bin/env python3
"""
Outil de récupération et renseignement des datatypes par colonne
================================================================

Ce script :
1. Lit les fichiers DPM (Annotated Table Layout .xlsx)
2. Pour chaque sheet DPM, extrait les colonnes avec leur unit/sign/Main Property
3. Cherche les datatypes correspondants dans le dictionnaire YML
4. Compare avec les pure-form XLSM et identifie les datatypes manquants
5. Mode REPORT : affiche les datatypes trouvés/manquants
6. Mode WRITE  : écrit les datatypes dans la ligne 7 des XLSM
7. Mode GUI    : interface interactive avec sélection des fichiers et sheets

Usage :
  python datatype_tool.py                       # Mode rapport
  python datatype_tool.py --write               # Écrire dans les XLSM
  python datatype_tool.py --gui                 # Interface graphique
  python datatype_tool.py --set ratio9Dec --sheet C_34.07.TOT --col 0020
"""

import argparse
import glob
import os
import re
import sys
import shutil
import yaml
import openpyxl
from openpyxl.utils import get_column_letter

# ── Configuration ──────────────────────────────────────────────────────────────

BASE_DIR = r'c:\Users\hnasserma\Documents\test'
DICT_FILE = os.path.join(BASE_DIR, 'pure-dictionary-euebacorep.yml')
SKIP_SHEETS = {'identification', 'cellSpecificities', 'labelTranslation', 'TOC'}

# ── Chargement du dictionnaire ─────────────────────────────────────────────────

def load_dictionary(dict_path):
    """Charge le dictionnaire YML et construit un index des définitions."""
    with open(dict_path, 'r', encoding='utf-8') as f:
        data = yaml.safe_load(f)

    def_index = {}   # {name: {datatypeId, type, formatType, ...}}

    for section_key in ('definitions', 'enumerations', 'pattern'):
        section = data.get(section_key, {})
        if not isinstance(section, dict):
            continue
        for name, entry in section.items():
            if not isinstance(entry, dict):
                continue
            if name not in def_index:
                def_index[name] = dict(entry)
                def_index[name]['_section'] = section_key

    return def_index


def get_dict_domain(dict_path):
    """Extrait le domaine du dictionnaire (ex: euebacorep) depuis le nom du fichier."""
    fname = os.path.basename(dict_path)
    m = re.search(r'pure-dictionary-(\w+)\.yml', fname)
    return m.group(1) if m else 'euebacorep'


# ── Parsing DPM ───────────────────────────────────────────────────────────────

def parse_dpm_cell(value):
    """
    Parse une cellule DPM contenant un data point.
    Format: 'dataPointId_x000D_\nunit_x000D_\nsign'
    ou    : 'dataPointId\r\nunit\r\nsign'
    Retourne: (data_point_id, unit, sign)
    """
    if not value:
        return None, None, None
    s = str(value)
    # Normaliser les séparateurs
    parts = re.split(r'_x000D_\s*|\r\n|\n', s)
    parts = [p.strip() for p in parts if p.strip()]
    if len(parts) < 2:
        return None, None, None
    dp_id = parts[0]
    unit = parts[1] if len(parts) > 1 else None
    sign = parts[2] if len(parts) > 2 else None
    return dp_id, unit, sign


def parse_dpm_property(value):
    """
    Parse une cellule DPM Main Property.
    Format: '(qHNJ) Exposure value' ou '(PDT) PD Range (Typed)'
    Retourne: (code, label)
    """
    if not value:
        return None, None
    s = str(value).strip()
    m = re.match(r'\(([^)]+)\)\s*(.*)', s)
    if m:
        return m.group(1), m.group(2).strip()
    return None, s


def read_dpm_columns(layout_path):
    """
    Lit un fichier DPM et extrait les métadonnées des colonnes pour chaque sheet.
    Retourne: {sheet_name: {col_code: {label, unit, sign, main_property_code, main_property_label}}}
    """
    if not os.path.exists(layout_path):
        print(f"  ERREUR: DPM non trouvé: {layout_path}", file=sys.stderr)
        return {}

    wb = openpyxl.load_workbook(layout_path, read_only=True, data_only=True)
    result = {}

    for sn in wb.sheetnames:
        if sn == 'TOC':
            continue
        ws = wb[sn]

        # Déterminer la structure du DPM sheet
        # La ligne des codes colonne est celle qui a "0010" ou "0005" etc.
        # Les labels sont une ligne au-dessus
        # Les data points commencent à la ligne suivante

        # Chercher la ligne des codes colonnes (row 7 typiquement)
        code_row = None
        first_data_col = None
        for row in range(5, 12):
            for col in range(4, 20):
                v = ws.cell(row=row, column=col).value
                if v and re.match(r'^0\d{3}$', str(v).strip()):
                    code_row = row
                    first_data_col = col
                    break
            if code_row:
                break

        if not code_row:
            continue

        # Extraire les codes colonnes et labels
        columns = {}
        label_row = code_row - 1
        # Il peut y avoir plusieurs lignes de labels (ex: row 5 + row 6)
        # On prend la ligne juste au-dessus des codes
        for col in range(first_data_col, first_data_col + 30):
            code_val = ws.cell(row=code_row, column=col).value
            if not code_val:
                continue
            code = str(code_val).strip()
            if not re.match(r'^0\d{2,4}$', code):
                continue

            # Label : chercher dans les lignes au-dessus
            label = None
            for lr in range(label_row, max(label_row - 3, 3), -1):
                lv = ws.cell(row=lr, column=col).value
                if lv:
                    label = str(lv).strip()
                    # Nettoyer le préfixe lettre (ex: "0a. Exposure value")
                    label = re.sub(r'^[a-z]\d*\.\s*', '', label, flags=re.IGNORECASE)
                    break

            columns[code] = {
                'label': label,
                'unit': None,
                'sign': None,
                'main_property_code': None,
                'main_property_label': None,
                'col_index': col,
            }

        # Extraire les data points pour obtenir unit/sign
        # Première ligne de données = code_row + 1 (ou +2 si "Rows" header)
        data_start = code_row + 1
        for row in range(data_start, data_start + 5):
            for code, info in columns.items():
                col = info['col_index']
                v = ws.cell(row=row, column=col).value
                if v and info['unit'] is None:
                    dp_id, unit, sign = parse_dpm_cell(v)
                    if unit:
                        info['unit'] = unit
                        info['sign'] = sign
                        info['data_point_id'] = dp_id

        # Chercher la ligne Main Property
        for row in range(data_start, data_start + 10):
            b_val = ws.cell(row=row, column=2).value
            if b_val and 'Main Property' in str(b_val):
                for code, info in columns.items():
                    col = info['col_index']
                    v = ws.cell(row=row, column=col).value
                    if v:
                        mp_code, mp_label = parse_dpm_property(v)
                        info['main_property_code'] = mp_code
                        info['main_property_label'] = mp_label
                break

        if columns:
            result[sn] = columns

    wb.close()
    return result


# ── Mapping DPM → Dictionnaire ────────────────────────────────────────────────

# Table de mapping connue: Main Property code → nom de définition dictionnaire
# Construite à partir de l'analyse des forms existants
KNOWN_PROPERTY_MAP = {
    # Montants (€£$)
    'qHNJ': 'positiveOrNegativeAmount',   # Exposure value
    'qINJ': 'positiveOrNegativeAmount',   # Risk weighted exposure amount (RWEA)

    # Pourcentages (%)
    'qILL': 'ratio9Dec',       # Probability of default (PD)
    'qIIJ': 'averwlgd',        # Loss Given Default (LGD)
    'qNNM': 'ratio9Dec',       # Density of RWEA
    'qNNL': 'ratio9Dec',       # Default rate

    # Entiers (#)
    'qFGE': 'nboblbygrd',      # Number of obligors
    'qGCF': 'nboblbygrd',      # Number of obligors defaulted
    'qGCA': 'mtrtyValueYears',  # Maturity (years)

    # Texte
    'PDT':  'alpha_255',       # PD Range (Typed)
    'qHIL': 'alpha_255',       # External Rating Equivalent
}

# Fallback par unit type quand pas de Main Property
UNIT_FALLBACK = {
    '€£$': 'positiveOrNegativeAmount',  # Montants
    '%':   'ratio9Dec',                 # Pourcentages → ratio
    '#':   'nboblbygrd',                # Entiers → nombre d'obligés par défaut
    'text': 'alpha_255',                # Texte
}

# Raffinement par sign DPM pour les montants €£$
SIGN_REFINEMENT = {
    'positive': 'positiveAmount',
    'negative': 'negativeOrNullAmount',
    None:       'positiveOrNegativeAmount',
}

# Affinement par mot-clé dans le label de colonne
# NB : LGD omis volontairement (ambigu : ratio9Dec ou averwlgd selon le formulaire)
LABEL_KEYWORDS = {
    'maturity': 'mtrtyValueYears',
    'maturité': 'mtrtyValueYears',
    'obligors': 'nboblbygrd',
    'obligés': 'nboblbygrd',
    'number of obligors': 'nboblbygrd',
    'probability of default': 'ratio9Dec',
    'density': 'ratio9Dec',
    'default rate': 'ratio9Dec',
    'average pd': 'ratio9Dec',
}


def determine_datatype(col_info, def_index, domain='euebacorep'):
    """
    Détermine le datatype PURE pour une colonne DPM.
    Retourne: (dict_name, confidence, source) ou (None, 0, reason)
    - dict_name : nom de la définition au dictionnaire (ex: 'ratio9Dec')
    - confidence : 0-3 (0=inconnu, 1=heuristique, 2=label, 3=Main Property)
    - source : explication
    """
    unit = col_info.get('unit')
    sign = col_info.get('sign')
    label = col_info.get('label', '') or ''
    mp_code = col_info.get('main_property_code')
    mp_label = col_info.get('main_property_label', '') or ''

    # 1. Main Property code → mapping connu
    if mp_code and mp_code in KNOWN_PROPERTY_MAP:
        name = KNOWN_PROPERTY_MAP[mp_code]
        if name and name in def_index:
            return name, 3, f"Main Property ({mp_code})"
        elif name is None:
            return None, 3, f"Main Property ({mp_code}) → défaut"

    # 2. Label keywords
    label_lower = label.lower()
    mp_label_lower = mp_label.lower()
    for keyword, dict_name in LABEL_KEYWORDS.items():
        if keyword in label_lower or keyword in mp_label_lower:
            if dict_name in def_index:
                return dict_name, 2, f"Label '{keyword}'"

    # 3. Fallback par unit type
    if unit in UNIT_FALLBACK:
        name = UNIT_FALLBACK[unit]
        if name and name in def_index:
            # Affiner pour # par le label
            if unit == '#' and ('matur' in label_lower or 'year' in label_lower or 'année' in label_lower):
                return 'mtrtyValueYears', 1, "Unit # + label maturity"
            return name, 1, f"Unit type '{unit}'"

    return None, 0, "Aucun mapping trouvé"


# ── Scan des formulaires XLSM ─────────────────────────────────────────────────

def scan_xlsm_datatypes(base_dir):
    """
    Scanne tous les pure-form XLSM et retourne les datatypes par colonne.
    Retourne: [{file, file_path, sheet, d2_datatype, col_datatypes: {code: name}, col_labels: {code: label}, col_codes: [codes]}]
    """
    results = []
    xlsm_files = sorted(glob.glob(os.path.join(base_dir, 'pure-form-*.xlsm')))

    for fpath in xlsm_files:
        fname = os.path.basename(fpath)
        if ' copy' in fname.lower() or '~$' in fname:
            continue
        try:
            wb = openpyxl.load_workbook(fpath, keep_vba=True)
        except Exception as e:
            print(f"  ERREUR chargement {fname}: {e}", file=sys.stderr)
            continue

        # Lire le datatype global depuis la sheet identification
        ident_datatype = None
        if 'identification' in wb.sheetnames:
            ws_id = wb['identification']
            ident_datatype = ws_id.cell(row=2, column=10).value  # J2

        for sn in wb.sheetnames:
            if sn in SKIP_SHEETS:
                continue
            ws = wb[sn]

            # Chercher la ligne d'en-tête des colonnes (où A='level')
            header_row = None
            for r in range(15, 22):
                if ws.cell(row=r, column=1).value == 'level':
                    header_row = r
                    break
            if header_row is None:
                continue

            # D2 = datatype au niveau sheet
            d2 = ws.cell(row=2, column=4).value

            # Row 7 = datatypes par colonne (colonne N+)
            col_datatypes = {}
            col_labels = {}
            col_codes = []
            for col in range(14, 60):  # N à BH
                code = ws.cell(row=header_row, column=col).value
                if not code:
                    continue
                code = str(code).strip()
                if not re.match(r'^0\d{2,4}$', code):
                    continue
                col_codes.append(code)

                # Datatype (row 7)
                dt = ws.cell(row=7, column=col).value
                if dt:
                    col_datatypes[code] = str(dt).strip()

                # Label : chercher dans les lignes entre 15 et header_row-1
                lbl = None
                for lr in range(15, header_row):
                    lv = ws.cell(row=lr, column=col).value
                    if lv:
                        lbl = str(lv).strip()
                        break
                if lbl:
                    col_labels[code] = lbl

            results.append({
                'file': fname,
                'file_path': fpath,
                'sheet': sn,
                'd2_datatype': d2,
                'ident_datatype': ident_datatype,
                'col_datatypes': col_datatypes,
                'col_labels': col_labels,
                'col_codes': col_codes,
            })

        wb.close()

    return results


# ── Matching DPM ↔ XLSM ──────────────────────────────────────────────────────

def match_dpm_to_form(dpm_columns, form_items):
    """
    Pour chaque form sheet, trouve le DPM sheet correspondant.
    Le matching se fait par nom de sheet.
    Retourne: [(form_item, dpm_sheet_name, dpm_cols)]
    """
    matches = []
    dpm_sheets = set(dpm_columns.keys())

    for item in form_items:
        sheet_name = item['sheet']

        # Match direct
        if sheet_name in dpm_sheets:
            matches.append((item, sheet_name, dpm_columns[sheet_name]))
            continue

        # Match par sous-tables (ex: C_34.07.TOT dans DPM mais pas le même nom)
        # Essayer les variantes
        found = False
        for dpm_sn in dpm_sheets:
            if sheet_name.replace('.', '_') == dpm_sn.replace('.', '_'):
                matches.append((item, dpm_sn, dpm_columns[dpm_sn]))
                found = True
                break
        if not found:
            matches.append((item, None, None))

    return matches


# ── Rapport ────────────────────────────────────────────────────────────────────

def print_report(matches, def_index, domain):
    """Affiche un rapport des datatypes à renseigner."""
    print("\n" + "=" * 90)
    print("RAPPORT DES DATATYPES PAR COLONNE")
    print("=" * 90)

    total_missing = 0
    total_filled = 0
    total_to_set = 0

    # Vérifier les identification datatypes manquants
    ident_checked = set()
    for form_item, dpm_sheet, dpm_cols in matches:
        fpath = form_item['file_path']
        if fpath in ident_checked:
            continue
        ident_checked.add(fpath)
        ident_dt = form_item.get('ident_datatype')
        if not ident_dt:
            print(f"\n⚠ {form_item['file']} : identification J2 (datatype) = (VIDE)")
            print(f"  → Suggéré : {domain}.positiveOrNegativeAmount")

    for form_item, dpm_sheet, dpm_cols in matches:
        ident_dt = form_item.get('ident_datatype')
        print(f"\n┌─ {form_item['file']} [{form_item['sheet']}]")
        print(f"│  identification datatype : {ident_dt or '(VIDE) ⚠'}")
        print(f"│  D2 (sheet datatype)     : {form_item['d2_datatype'] or '(vide)'}")

        if not dpm_cols:
            print(f"│  ⚠ Pas de DPM correspondant trouvé")
            # Afficher les datatypes existants
            for code in form_item['col_codes']:
                dt = form_item['col_datatypes'].get(code, '')
                lbl = form_item['col_labels'].get(code, '')
                if dt:
                    print(f"│  {code}: {dt:40s} [{lbl}]")
                    total_filled += 1
                else:
                    print(f"│  {code}: {'(vide)':40s} [{lbl}]")
                    total_missing += 1
            print(f"└{'─' * 80}")
            continue

        print(f"│  DPM sheet : {dpm_sheet}")
        print(f"│")

        for code in form_item['col_codes']:
            existing_dt = form_item['col_datatypes'].get(code)
            form_label = form_item['col_labels'].get(code, '')
            dpm_col = dpm_cols.get(code, {})
            dpm_unit = dpm_col.get('unit', '?')
            dpm_label = dpm_col.get('label', '')
            mp_code = dpm_col.get('main_property_code', '')

            if dpm_col:
                suggested, confidence, source = determine_datatype(dpm_col, def_index, domain)
            else:
                suggested, confidence, source = None, 0, "Colonne absente du DPM"

            # Formatter le résultat
            conf_str = ['?', '~', '✓', '★'][min(confidence, 3)]
            full_suggested = f"{domain}.{suggested}" if suggested else None

            if existing_dt:
                if suggested and existing_dt != full_suggested:
                    status = f"✎ existant={existing_dt}, suggéré={full_suggested}"
                    total_filled += 1
                else:
                    status = f"✓ {existing_dt}"
                    total_filled += 1
            elif suggested:
                status = f"→ {full_suggested}  [{conf_str} {source}]"
                total_to_set += 1
            else:
                status = f"  (aucun mapping)"
                total_missing += 1

            print(f"│  {code}: {status:60s} [{form_label or dpm_label}]")

        print(f"└{'─' * 80}")

    # Résumé identification
    ident_missing = set()
    for form_item, _, _ in matches:
        if not form_item.get('ident_datatype'):
            ident_missing.add(form_item['file'])

    print(f"\n{'=' * 90}")
    print("RÉSUMÉ")
    print(f"{'=' * 90}")
    if ident_missing:
        print(f"  Formulaires sans datatype identification : {len(ident_missing)}")
        for f in sorted(ident_missing):
            print(f"    → {f}")
    print(f"  Colonnes avec datatype renseigné : {total_filled}")
    print(f"  Colonnes à renseigner (suggéré)  : {total_to_set}")
    print(f"  Colonnes sans mapping            : {total_missing}")


# ── Écriture dans les XLSM ────────────────────────────────────────────────────

def write_datatypes(matches, def_index, domain, base_dir):
    """
    Écrit les datatypes dans les fichiers XLSM :
    - identification J2 : datatype par défaut si manquant
    - ligne 7 : datatypes par colonne si manquants
    """
    by_file = {}
    for form_item, dpm_sheet, dpm_cols in matches:
        by_file.setdefault(form_item['file_path'], []).append(
            (form_item, dpm_sheet, dpm_cols))

    modified_count = 0

    for fpath, items in by_file.items():
        fname = os.path.basename(fpath)

        # Backup
        backup_path = fpath + '.bak'
        if not os.path.exists(backup_path):
            shutil.copy2(fpath, backup_path)
            print(f"  Sauvegarde créée : {os.path.basename(backup_path)}")

        try:
            wb = openpyxl.load_workbook(fpath, keep_vba=True)
        except Exception as e:
            print(f"  ERREUR ouverture {fname}: {e}", file=sys.stderr)
            continue

        file_modified = False

        # ── Vérifier/renseigner identification J2 (datatype par défaut) ──
        ident_dt = items[0][0].get('ident_datatype')
        if not ident_dt and 'identification' in wb.sheetnames:
            ws_id = wb['identification']
            default_dt = f"{domain}.positiveOrNegativeAmount"
            old_val = ws_id.cell(row=2, column=10).value
            ws_id.cell(row=2, column=10, value=default_dt)
            print(f"  ✓ {fname}[identification] J2: '{old_val}' → '{default_dt}'")
            file_modified = True

        for form_item, dpm_sheet, dpm_cols in items:
            sn = form_item['sheet']
            if sn not in wb.sheetnames:
                print(f"  ⚠ Sheet {sn} non trouvée dans {fname}")
                continue
            if not dpm_cols:
                continue

            ws = wb[sn]
            cols_written = 0

            # Trouver la ligne d'en-tête (où A='level')
            header_row = None
            for r in range(15, 22):
                if ws.cell(row=r, column=1).value == 'level':
                    header_row = r
                    break
            if not header_row:
                continue

            for col in range(14, 60):
                code = ws.cell(row=header_row, column=col).value
                if not code:
                    continue
                code = str(code).strip()

                # Ne rien écrire si déjà renseigné
                existing = ws.cell(row=7, column=col).value
                if existing:
                    continue

                dpm_col = dpm_cols.get(code)
                if not dpm_col:
                    continue

                suggested, confidence, source = determine_datatype(
                    dpm_col, def_index, domain)
                if not suggested:
                    continue

                full_name = f"{domain}.{suggested}"
                ws.cell(row=7, column=col, value=full_name)
                cols_written += 1

            if cols_written > 0:
                print(f"  ✓ {fname}[{sn}] : {cols_written} datatypes écrits")
                file_modified = True
            else:
                print(f"  - {fname}[{sn}] : rien à écrire")

        if file_modified:
            try:
                wb.save(fpath)
                modified_count += 1
                print(f"  ✓ Fichier sauvegardé : {fname}")
            except PermissionError:
                print(f"  ✗ '{fname}' verrouillé (ouvert dans Excel ?)")

        wb.close()

    print(f"\n  Total fichiers modifiés : {modified_count}")


# ── Mode SET : renseigner une colonne spécifique ──────────────────────────────

def set_datatype_for_column(sheet_name, col_code, dict_name, def_index, domain, base_dir):
    """
    Renseigne le datatype pour une colonne spécifique d'une sheet.
    """
    if dict_name not in def_index:
        print(f"  ERREUR: '{dict_name}' non trouvé dans le dictionnaire!")
        # Suggérer des entrées proches
        candidates = [n for n in sorted(def_index.keys()) if dict_name.lower() in n.lower()]
        if candidates:
            print(f"  Entrées proches :")
            for c in candidates[:10]:
                info = def_index[c]
                print(f"    - {c} (datatypeId={info.get('datatypeId', '?')}, formatType={info.get('formatType', '?')})")
        return False

    full_name = f"{domain}.{dict_name}"

    # Trouver le fichier XLSM
    xlsm_files = sorted(glob.glob(os.path.join(base_dir, 'pure-form-*.xlsm')))
    target_file = None
    for f in xlsm_files:
        fname = os.path.basename(f)
        if ' copy' in fname.lower() or '~$' in fname:
            continue
        try:
            wb_check = openpyxl.load_workbook(f, keep_vba=True)
            if sheet_name in wb_check.sheetnames:
                target_file = f
            wb_check.close()
            if target_file:
                break
        except Exception:
            continue

    if not target_file:
        print(f"  ERREUR: Aucun fichier XLSM contenant la sheet '{sheet_name}'")
        return False

    fname = os.path.basename(target_file)

    # Backup
    backup_path = target_file + '.bak'
    if not os.path.exists(backup_path):
        shutil.copy2(target_file, backup_path)
        print(f"  Sauvegarde créée : {os.path.basename(backup_path)}")

    wb = openpyxl.load_workbook(target_file, keep_vba=True)
    ws = wb[sheet_name]

    # Trouver la colonne correspondant au code
    target_col = None
    # Trouver la ligne d'en-tête (où A='level')
    header_row = None
    for r in range(15, 22):
        if ws.cell(row=r, column=1).value == 'level':
            header_row = r
            break
    if not header_row:
        print(f"  ERREUR: Impossible de trouver la ligne d'en-tête dans {sheet_name}")
        wb.close()
        return False

    for col in range(14, 60):
        v = ws.cell(row=header_row, column=col).value
        if v and str(v).strip() == col_code:
            target_col = col
            break

    if target_col is None:
        print(f"  ERREUR: Code colonne '{col_code}' non trouvé dans la ligne 16 de {sheet_name}")
        wb.close()
        return False

    old_val = ws.cell(row=7, column=target_col).value
    ws.cell(row=7, column=target_col, value=full_name)
    print(f"  ✓ {fname}[{sheet_name}] col {col_code} ({get_column_letter(target_col)}7): '{old_val}' → '{full_name}'")

    wb.save(target_file)
    print(f"  ✓ Fichier sauvegardé : {fname}")
    wb.close()
    return True


# ── Mode GUI ──────────────────────────────────────────────────────────────────

def run_interactive(base_dir):
    """
    Mode interactif GUI :
      1. Sélection des fichiers DPM
      2. Sélection du dictionnaire
      3. Affichage des colonnes avec datatypes suggérés
      4. Sélection et écriture
    """
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk

    root = tk.Tk()
    root.withdraw()

    # ── Étape 1 : Sélection du/des fichier(s) DPM ───────────────────────────
    messagebox.showinfo(
        "Datatype Tool",
        "Étape 1/3 : Sélectionnez le(s) fichier(s) DPM\n"
        "(Annotated Table Layout .xlsx)\n\n"
        "Vous pouvez en sélectionner plusieurs (Ctrl+clic)."
    )

    layout_paths = filedialog.askopenfilenames(
        title="Sélectionner le(s) fichier(s) DPM (Annotated Table Layout)",
        initialdir=base_dir,
        filetypes=[("Fichiers Excel", "*.xlsx"), ("Tous", "*.*")]
    )

    if not layout_paths:
        print("Annulé : Aucun fichier DPM sélectionné.")
        root.destroy()
        return

    for lp in layout_paths:
        print(f"  DPM sélectionné : {os.path.basename(lp)}")

    # ── Étape 2 : Sélection du dictionnaire YML ─────────────────────────────
    messagebox.showinfo(
        "Datatype Tool",
        "Étape 2/3 : Sélectionnez le fichier dictionnaire\n"
        "(pure-dictionary .yml)"
    )

    dict_path = filedialog.askopenfilename(
        title="Sélectionner le fichier dictionnaire (.yml)",
        initialdir=base_dir,
        filetypes=[("Fichiers YAML", "*.yml *.yaml"), ("Tous", "*.*")]
    )

    if not dict_path:
        print("Annulé : Aucun dictionnaire sélectionné.")
        root.destroy()
        return

    domain = get_dict_domain(dict_path)
    print(f"  Dictionnaire sélectionné : {os.path.basename(dict_path)} (domaine: {domain})")

    # Chargement
    print("Chargement du dictionnaire...")
    def_index = load_dictionary(dict_path)
    print(f"  {len(def_index)} définitions trouvées")

    print("Lecture des DPM...")
    dpm_columns = {}
    for lp in layout_paths:
        partial = read_dpm_columns(lp)
        for sheet_name, cols in partial.items():
            if sheet_name not in dpm_columns:
                dpm_columns[sheet_name] = cols
    print(f"  {len(dpm_columns)} sheets trouvées dans les DPM")

    print("Scan des formulaires XLSM...")
    form_items = scan_xlsm_datatypes(base_dir)
    print(f"  {len(form_items)} sheets de données trouvées")

    matches = match_dpm_to_form(dpm_columns, form_items)

    # ── Préparer les données pour la table ───────────────────────────────────
    table_data = []  # [(file, sheet, col_code, label, existing, suggested, confidence, source, unit, form_item)]
    for form_item, dpm_sheet, dpm_cols in matches:
        for code in form_item['col_codes']:
            existing = form_item['col_datatypes'].get(code, '')
            form_label = form_item['col_labels'].get(code, '')

            if dpm_cols and code in dpm_cols:
                dpm_col = dpm_cols[code]
                suggested, confidence, source = determine_datatype(dpm_col, def_index, domain)
                unit = dpm_col.get('unit', '')
                dpm_label = dpm_col.get('label', '')
            else:
                suggested, confidence, source = None, 0, ""
                unit = ""
                dpm_label = ""

            label = form_label or dpm_label
            full_suggested = f"{domain}.{suggested}" if suggested else ""

            table_data.append({
                'file': form_item['file'],
                'file_path': form_item['file_path'],
                'sheet': form_item['sheet'],
                'col_code': code,
                'label': label,
                'existing': existing,
                'suggested': full_suggested,
                'suggested_name': suggested,
                'confidence': confidence,
                'source': source,
                'unit': unit,
                'form_item': form_item,
            })

    # ── Étape 3 : Fenêtre de sélection ───────────────────────────────────────
    sel_window = tk.Toplevel(root)
    sel_window.title("Datatype Tool — Datatypes par colonne")
    sel_window.geometry("1200x700")

    # Titre
    title_frame = tk.Frame(sel_window, padx=10, pady=5)
    title_frame.pack(fill=tk.X)
    tk.Label(title_frame,
             text="Étape 3/3 : Vérifiez et sélectionnez les datatypes à écrire",
             font=("Segoe UI", 12, "bold")).pack(anchor=tk.W)
    dpm_names = ', '.join(os.path.basename(lp) for lp in layout_paths)
    tk.Label(title_frame, text=f"DPM : {dpm_names}",
             font=("Segoe UI", 9), fg="gray").pack(anchor=tk.W)

    # Filtre
    filter_frame = tk.Frame(sel_window, padx=10, pady=3)
    filter_frame.pack(fill=tk.X)
    tk.Label(filter_frame, text="Filtre :").pack(side=tk.LEFT)
    filter_var = tk.StringVar(value="à renseigner")
    for txt, val in [("Tout", "all"), ("À renseigner", "à renseigner"), ("Déjà renseigné", "renseigné")]:
        tk.Radiobutton(filter_frame, text=txt, variable=filter_var, value=val).pack(side=tk.LEFT, padx=5)

    # Treeview avec scrollbar
    tree_frame = tk.Frame(sel_window, padx=10, pady=5)
    tree_frame.pack(fill=tk.BOTH, expand=True)

    columns = ("select", "file", "sheet", "colcode", "label", "unit", "existing", "suggested", "conf", "source")
    tree = ttk.Treeview(tree_frame, columns=columns, show='headings', selectmode='extended')

    tree.heading("select", text="✓")
    tree.heading("file", text="Fichier")
    tree.heading("sheet", text="Sheet")
    tree.heading("colcode", text="Col")
    tree.heading("label", text="Label")
    tree.heading("unit", text="Unit")
    tree.heading("existing", text="Existant")
    tree.heading("suggested", text="Suggéré")
    tree.heading("conf", text="Fiab.")
    tree.heading("source", text="Source")

    tree.column("select", width=30, anchor="center")
    tree.column("file", width=180)
    tree.column("sheet", width=100)
    tree.column("colcode", width=50, anchor="center")
    tree.column("label", width=200)
    tree.column("unit", width=40, anchor="center")
    tree.column("existing", width=150)
    tree.column("suggested", width=170)
    tree.column("conf", width=35, anchor="center")
    tree.column("source", width=140)

    scrollbar_y = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=tree.yview)
    scrollbar_x = ttk.Scrollbar(tree_frame, orient=tk.HORIZONTAL, command=tree.xview)
    tree.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)

    tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)
    scrollbar_x.pack(side=tk.BOTTOM, fill=tk.X)

    # État de sélection par item
    selected_set = set()

    def populate_tree(filter_mode="all"):
        tree.delete(*tree.get_children())
        for i, row in enumerate(table_data):
            if filter_mode == "à renseigner" and row['existing']:
                continue
            if filter_mode == "renseigné" and not row['existing']:
                continue
            if not row['suggested']:
                continue

            conf_str = ['?', '~', '✓', '★'][min(row['confidence'], 3)]
            sel_str = "☑" if i in selected_set else "☐"
            tag = 'selected' if i in selected_set else ''

            tree.insert("", tk.END, iid=str(i), values=(
                sel_str,
                row['file'],
                row['sheet'],
                row['col_code'],
                row['label'],
                row['unit'],
                row['existing'] or '',
                row['suggested'],
                conf_str,
                row['source'],
            ), tags=(tag,))

        tree.tag_configure('selected', background='#d4edda')

    def toggle_selection(event):
        item = tree.focus()
        if not item:
            return
        idx = int(item)
        if idx in selected_set:
            selected_set.discard(idx)
        else:
            selected_set.add(idx)
        populate_tree(filter_var.get())

    # Auto-sélectionner les colonnes sans datatype existant avec suggestion de confiance >= 1
    for i, row in enumerate(table_data):
        if not row['existing'] and row['suggested'] and row['confidence'] >= 1:
            selected_set.add(i)

    populate_tree("à renseigner")

    tree.bind('<Double-1>', toggle_selection)
    filter_var.trace_add('write', lambda *_: populate_tree(filter_var.get()))

    # Boutons
    btn_frame = tk.Frame(sel_window, padx=10, pady=10)
    btn_frame.pack(fill=tk.X)

    count_label = tk.Label(btn_frame, text="", font=("Segoe UI", 9))
    count_label.pack(side=tk.LEFT, padx=10)

    def update_count():
        n = len(selected_set)
        count_label.config(text=f"{n} colonne(s) sélectionnée(s)")
        sel_window.after(500, update_count)

    update_count()

    def select_all_visible():
        for item in tree.get_children():
            idx = int(item)
            if table_data[idx]['suggested'] and not table_data[idx]['existing']:
                selected_set.add(idx)
        populate_tree(filter_var.get())

    def deselect_all():
        selected_set.clear()
        populate_tree(filter_var.get())

    result_action = [None]

    def on_write():
        result_action[0] = 'write'
        sel_window.destroy()

    def on_cancel():
        sel_window.destroy()

    tk.Button(btn_frame, text="Tout sélectionner (visible)",
              command=select_all_visible).pack(side=tk.LEFT, padx=5)
    tk.Button(btn_frame, text="Tout désélectionner",
              command=deselect_all).pack(side=tk.LEFT, padx=5)
    tk.Button(btn_frame, text="Annuler",
              command=on_cancel).pack(side=tk.RIGHT, padx=5)
    tk.Button(btn_frame, text="Écrire les datatypes", command=on_write,
              bg="#4CAF50", fg="white",
              font=("Segoe UI", 10, "bold")).pack(side=tk.RIGHT, padx=5)

    sel_window.wait_window()

    if result_action[0] != 'write' or not selected_set:
        print("Annulé : aucune colonne sélectionnée.")
        root.destroy()
        return

    # ── Écriture ─────────────────────────────────────────────────────────────
    print(f"\nÉcriture de {len(selected_set)} datatypes...")

    # Grouper par fichier
    by_file = {}
    for idx in sorted(selected_set):
        row = table_data[idx]
        by_file.setdefault(row['file_path'], []).append(row)

    written = 0
    for fpath, rows in by_file.items():
        fname = os.path.basename(fpath)

        # Backup
        backup_path = fpath + '.bak'
        if not os.path.exists(backup_path):
            shutil.copy2(fpath, backup_path)
            print(f"  Sauvegarde créée : {os.path.basename(backup_path)}")

        try:
            wb = openpyxl.load_workbook(fpath, keep_vba=True)
        except PermissionError:
            print(f"  ✗ '{fname}' verrouillé (ouvert dans Excel ?)")
            continue
        except Exception as e:
            print(f"  ✗ Erreur ouverture {fname}: {e}")
            continue

        for row in rows:
            sn = row['sheet']
            if sn not in wb.sheetnames:
                continue
            ws = wb[sn]

            # Trouver la colonne pour ce code
            # D'abord trouver la ligne d'en-tête
            hdr = None
            for r in range(15, 22):
                if ws.cell(row=r, column=1).value == 'level':
                    hdr = r
                    break
            if not hdr:
                continue
            for col in range(14, 60):
                v = ws.cell(row=hdr, column=col).value
                if v and str(v).strip() == row['col_code']:
                    old_val = ws.cell(row=7, column=col).value
                    ws.cell(row=7, column=col, value=row['suggested'])
                    print(f"  ✓ {fname}[{sn}] col {row['col_code']} ({get_column_letter(col)}7): "
                          f"'{old_val}' → '{row['suggested']}'")
                    written += 1
                    break

        try:
            wb.save(fpath)
            print(f"  ✓ Fichier sauvegardé : {fname}")
        except PermissionError:
            print(f"  ✗ '{fname}' verrouillé, impossible de sauvegarder")
        wb.close()

    print(f"\nTerminé ! {written} datatypes écrits.")
    messagebox.showinfo("Datatype Tool", f"Terminé !\n{written} datatypes écrits.")
    root.destroy()


# ── Main ───────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description="Outil de récupération et renseignement des datatypes par colonne"
    )
    parser.add_argument(
        '--write', action='store_true',
        help="Écrire les datatypes dans les fichiers XLSM (sinon : rapport seul)"
    )
    parser.add_argument(
        '--set', metavar='TYPE',
        help="Renseigner un datatype pour une colonne. Ex: --set ratio9Dec"
    )
    parser.add_argument(
        '--sheet', metavar='SHEET',
        help="Sheet cible pour --set. Ex: --sheet C_34.07.TOT"
    )
    parser.add_argument(
        '--col', metavar='CODE',
        help="Code colonne pour --set. Ex: --col 0020"
    )
    parser.add_argument(
        '--dir', default=BASE_DIR,
        help=f"Répertoire de travail (défaut: {BASE_DIR})"
    )
    parser.add_argument(
        '--gui', action='store_true',
        help="Mode interactif GUI"
    )
    args = parser.parse_args()

    base_dir = args.dir

    # ── Mode GUI ──
    if args.gui:
        run_interactive(base_dir)
        return

    # Trouver le dictionnaire
    dict_path = DICT_FILE
    dict_files = glob.glob(os.path.join(base_dir, 'pure-dictionary-*.yml'))
    if dict_files:
        dict_path = dict_files[0]

    if not os.path.exists(dict_path):
        print(f"ERREUR: Dictionnaire non trouvé: {dict_path}", file=sys.stderr)
        sys.exit(1)

    domain = get_dict_domain(dict_path)

    # Charger le dictionnaire
    print("Chargement du dictionnaire...")
    def_index = load_dictionary(dict_path)
    print(f"  {len(def_index)} définitions trouvées (domaine: {domain})")

    # ── Mode SET : renseigner une colonne spécifique ──
    if args.set:
        if not args.sheet or not args.col:
            print("ERREUR: --set nécessite --sheet et --col.",
                  file=sys.stderr)
            print("  Ex: python datatype_tool.py --set ratio9Dec --sheet C_34.07.TOT --col 0020",
                  file=sys.stderr)
            sys.exit(1)
        ok = set_datatype_for_column(args.sheet, args.col, args.set, def_index, domain, base_dir)
        sys.exit(0 if ok else 1)

    # ── Lire les DPM ──
    print("Lecture des DPM...")
    dpm_files = sorted(glob.glob(os.path.join(base_dir, '*Annotated*Table*Layout*.xlsx')))
    if not dpm_files:
        print("  ⚠ Aucun fichier DPM trouvé. Rapport basé sur les XLSM uniquement.")
        dpm_columns = {}
    else:
        dpm_columns = {}
        for lp in dpm_files:
            print(f"  Lecture de {os.path.basename(lp)}...")
            partial = read_dpm_columns(lp)
            for sheet_name, cols in partial.items():
                if sheet_name not in dpm_columns:
                    dpm_columns[sheet_name] = cols
        print(f"  {len(dpm_columns)} sheets trouvées dans les DPM")

    # ── Scanner les XLSM ──
    print("Scan des formulaires XLSM...")
    form_items = scan_xlsm_datatypes(base_dir)
    print(f"  {len(form_items)} sheets de données trouvées")

    if not form_items:
        print("Aucun formulaire pure-form trouvé.")
        sys.exit(0)

    # Matching
    matches = match_dpm_to_form(dpm_columns, form_items)

    # Rapport
    print_report(matches, def_index, domain)

    # ── Mode WRITE ──
    if args.write:
        print("\n" + "=" * 90)
        print("ÉCRITURE DES DATATYPES")
        print("=" * 90)
        write_datatypes(matches, def_index, domain, base_dir)


if __name__ == '__main__':
    main()
