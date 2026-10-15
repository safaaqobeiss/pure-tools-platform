#!/usr/bin/env python3
"""
Outil unifié d'enrichissement des pure-form XLSM
=================================================

Ce script fusionne les fonctionnalités de zaxis_tool et datatype_tool :

  Z-AXIS :
    1. Lit les fichiers DPM (Annotated Table Layout) pour trouver les Z-axis
    2. Mappe les variants DPM → types du dictionnaire YML
    3. Écrit J2 (code), J4 (associatedAttribute), enum values dans les XLSM

  DATATYPE :
    1. Lit les colonnes DPM avec unit/sign/Main Property
    2. Détermine le datatype PURE via mapping connu + heuristiques
    3. Écrit les datatypes dans identification J2 et row 7 des XLSM

Modes CLI :
  python enrichment_tool.py                          # Rapport Z-axis + datatypes
  python enrichment_tool.py --write                  # Écrire tout
  python enrichment_tool.py --zaxis-only             # Rapport Z-axis seulement
  python enrichment_tool.py --datatype-only          # Rapport datatypes seulement
  python enrichment_tool.py --write --zaxis-only     # Écrire Z-axis seulement
  python enrichment_tool.py --write --datatype-only  # Écrire datatypes seulement
  python enrichment_tool.py --gui                    # Interface graphique unifiée

  # Renseigner le Z-axis pour une sheet :
  python enrichment_tool.py --set-zaxis typspelendingexp --sheet K_29.02.a

  # Renseigner un datatype pour une colonne :
  python enrichment_tool.py --set-datatype ratio9Dec --sheet C_34.07.TOT --col 0020
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

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SKIP_SHEETS = {'identification', 'cellSpecificities', 'labelTranslation', 'TOC'}


# ══════════════════════════════════════════════════════════════════════════════
#  PARTIE COMMUNE : dictionnaire, scan XLSM, utilitaires
# ══════════════════════════════════════════════════════════════════════════════

def load_dictionary(dict_path):
    """Charge le dictionnaire YML et construit deux index :
    - type_index : types avec enum/submissionValue (pour Z-axis)
    - def_index  : toutes les définitions avec datatypeId/formatType (pour datatypes)
    """
    with open(dict_path, 'r', encoding='utf-8') as f:
        data = yaml.safe_load(f)

    type_index = {}   # {type_name: {enum, submissionValue, section}}
    def_index = {}    # {name: {datatypeId, type, formatType, ...}}

    for section_key in ('enumerations', 'definitions', 'pattern'):
        section = data.get(section_key, {})
        if not isinstance(section, dict):
            continue
        for name, entry in section.items():
            if not isinstance(entry, dict):
                continue

            # Index des définitions (pour datatypes)
            if name not in def_index:
                def_index[name] = dict(entry)
                def_index[name]['_section'] = section_key

            # Index des types avec enum (pour Z-axis)
            enum_vals = entry.get('enum', [])
            if enum_vals and name not in type_index:
                sub_vals = entry.get('submissionValue', {})
                if not isinstance(sub_vals, dict):
                    sub_vals = {}
                type_index[name] = {
                    'enum': enum_vals,
                    'submissionValue': sub_vals,
                    'section': section_key,
                }

    return type_index, def_index


def get_dict_domain(dict_path):
    """Extrait le domaine du dictionnaire (ex: euebacorep) depuis le nom du fichier."""
    fname = os.path.basename(dict_path)
    m = re.search(r'pure-dictionary-(\w+)\.yml', fname)
    if not m:
        raise ValueError(f"Impossible d'extraire le domaine depuis le nom du fichier : {fname}\n"
                         f"Le fichier doit suivre le pattern : pure-dictionary-<domaine>.yml")
    return m.group(1)


def find_xlsm_files(base_dir):
    """Trouve tous les fichiers pure-form XLSM valides."""
    xlsm_files = sorted(glob.glob(os.path.join(base_dir, 'pure-form-*.xlsm')))
    return [f for f in xlsm_files
            if '~$' not in os.path.basename(f)
            and ' copy' not in os.path.basename(f).lower()]


def create_backup(fpath):
    """Crée une sauvegarde .bak si elle n'existe pas déjà."""
    backup_path = fpath + '.bak'
    if not os.path.exists(backup_path):
        shutil.copy2(fpath, backup_path)
        print(f"  Sauvegarde créée : {os.path.basename(backup_path)}")
    return backup_path


def find_header_row(ws):
    """Cherche la ligne d'en-tête (où A='level') entre les lignes 15 et 30."""
    for r in range(15, 31):
        if ws.cell(row=r, column=1).value == 'level':
            return r
    return None


def find_sheet_in_xlsm(sheet_name, base_dir):
    """Trouve le fichier XLSM contenant une sheet donnée."""
    for fpath in find_xlsm_files(base_dir):
        try:
            wb = openpyxl.load_workbook(fpath, keep_vba=True)
            found = sheet_name in wb.sheetnames
            wb.close()
            if found:
                return fpath
        except Exception:
            continue
    return None


# ══════════════════════════════════════════════════════════════════════════════
#  PARTIE Z-AXIS : lecture DPM, mapping, scan, rapport, écriture
# ══════════════════════════════════════════════════════════════════════════════

def read_annotated_layout(layout_path):
    """
    Lit un fichier DPM (Annotated Table Layout .xlsx) et retourne les sheets
    avec un Z-axis.
    Retourne: {sheet_name: {'hierarchy': str, 'variant': str}}
    """
    if not os.path.exists(layout_path):
        print(f"  ERREUR: Fichier DPM non trouvé: {layout_path}", file=sys.stderr)
        return {}

    wb = openpyxl.load_workbook(layout_path, read_only=True, data_only=True)
    result = {}

    for sn in wb.sheetnames:
        ws = wb[sn]
        b3 = ws.cell(row=3, column=2).value
        if not b3 or 'Sheet per' not in str(b3):
            continue

        d3 = ws.cell(row=3, column=4).value
        if not d3:
            continue

        m = re.search(r'\((\w+):(\w+)\(', str(d3))
        if not m:
            continue

        hierarchy = m.group(1)
        variant = m.group(2)
        result[sn] = {'hierarchy': hierarchy, 'variant': variant}

    wb.close()
    return result


def map_variant_to_type(hierarchy, variant, type_index):
    """
    Mappe un code variant DPM (ex: qAE5, new_TA9, EN2) à un type dictionnaire.
    Utilise les submissionValue avec préfixe eba_<hierarchy>: pour trouver les candidats.
    """
    prefix = f"eba_{hierarchy}:"

    candidates = []
    for type_name, info in type_index.items():
        sub_vals = info.get('submissionValue', {})
        if not sub_vals:
            continue
        has_prefix = any(str(v).startswith(prefix) for v in sub_vals.values() if v)
        if not has_prefix:
            continue
        all_prefix = all(str(v).startswith(prefix) for v in sub_vals.values() if v)
        candidates.append((type_name, all_prefix, len(info.get('enum', []))))

    if not candidates:
        return None

    pure = [(name, count) for name, all_match, count in candidates if all_match]
    if not pure:
        pure = [(name, count) for name, _, count in candidates]

    m = re.search(r'(\d+)$', variant.replace('new_', ''))
    if m:
        n = m.group(1)
        sub_n_pattern = f"sub{n}"
        sub_n_matches = [name for name, _ in pure if sub_n_pattern in name.lower()]
        if sub_n_matches:
            # Préférer les noms venant de 'definitions' (ex: expcladsub5)
            # plutôt que 'enumerations' (ex: exposureClassSub5)
            def_matches = [name for name in sub_n_matches
                           if type_index[name].get('section') == 'definitions']
            if def_matches:
                return def_matches[0]
            return sub_n_matches[0]

    if len(pure) == 1:
        return pure[0][0]

    if pure:
        return pure[0][0]

    return None


def scan_xlsm_zaxis(base_dir):
    """
    Scanne tous les pure-form XLSM et retourne les infos Z-axis.
    Retourne: [{file, file_path, sheet, j2, j4, comp_rule_row, existing_values}]
    """
    results = []

    for fpath in find_xlsm_files(base_dir):
        fname = os.path.basename(fpath)
        try:
            wb = openpyxl.load_workbook(fpath, keep_vba=True)
        except Exception as e:
            print(f"  ERREUR chargement {fname}: {e}", file=sys.stderr)
            continue

        for sn in wb.sheetnames:
            if sn in SKIP_SHEETS:
                continue
            ws = wb[sn]
            i1 = ws.cell(row=1, column=9).value
            if i1 != 'zAxis':
                continue

            j2 = ws.cell(row=2, column=10).value
            j4 = ws.cell(row=4, column=10).value

            comp_rule_row = None
            for row in range(5, ws.max_row + 1):
                if ws.cell(row=row, column=9).value == 'computationRuleOrder':
                    comp_rule_row = row
                    break

            existing_values = []
            if comp_rule_row:
                for row in range(5, comp_rule_row):
                    val = ws.cell(row=row, column=10).value
                    if val:
                        existing_values.append(val)

            results.append({
                'file': fname,
                'file_path': fpath,
                'sheet': sn,
                'j2': j2,
                'j4': j4,
                'comp_rule_row': comp_rule_row,
                'existing_values': existing_values,
            })

        wb.close()

    return results


def print_zaxis_report(zaxis_items, type_index):
    """Affiche un rapport détaillé des classes Z-axis par formulaire."""
    has_zaxis = [item for item in zaxis_items if item['j2']]
    no_zaxis = [item for item in zaxis_items if not item['j2']]

    print("\n" + "=" * 80)
    print("RAPPORT DES CLASSES Z-AXIS PAR FORMULAIRE")
    print("=" * 80)

    for item in has_zaxis:
        type_name = item['j2']
        print(f"\n┌─ Fichier : {item['file']}")
        print(f"│  Sheet   : {item['sheet']}")
        print(f"│  Type    : {type_name}")
        print(f"│  Attribut: {item['j4']}")

        if type_name in type_index:
            ti = type_index[type_name]
            enum_vals = ti['enum']
            sub_vals = ti['submissionValue']
            print(f"│  Nombre de classes : {len(enum_vals)}")
            print(f"│")
            for i, val in enumerate(enum_vals, 1):
                sv = sub_vals.get(val, '—')
                print(f"│  {i:3d}. {val:20s} → {sv}")
        else:
            print(f"│  ⚠ Type '{type_name}' NON TROUVÉ dans le dictionnaire!")

        if item['existing_values']:
            print(f"│")
            print(f"│  Valeurs déjà présentes dans le XLSM : {item['existing_values']}")

        print(f"└{'─' * 70}")

    if no_zaxis:
        print(f"\n{'─' * 80}")
        print(f"Sheets SANS Z-axis (J2 vide) : {len(no_zaxis)}")
        for item in no_zaxis:
            print(f"  - {item['file']} → [{item['sheet']}]")

    print(f"\n{'=' * 80}")
    print("RÉSUMÉ Z-AXIS")
    print(f"{'=' * 80}")
    print(f"  Total sheets avec Z-axis header : {len(zaxis_items)}")
    print(f"  Sheets avec Z-axis renseigné    : {len(has_zaxis)}")
    print(f"  Sheets sans Z-axis (normal)     : {len(no_zaxis)}")
    total_classes = sum(
        len(type_index[item['j2']]['enum'])
        for item in has_zaxis
        if item['j2'] in type_index
    )
    print(f"  Total classes à renseigner      : {total_classes}")


def write_zaxis_values(zaxis_items, type_index, base_dir):
    """Écrit les valeurs enum Z-axis dans les fichiers XLSM."""
    by_file = {}
    for item in zaxis_items:
        if not item['j2']:
            continue
        if item['j2'] not in type_index:
            continue
        by_file.setdefault(item['file_path'], []).append(item)

    modified_count = 0

    for fpath, items in by_file.items():
        fname = os.path.basename(fpath)
        create_backup(fpath)

        try:
            wb = openpyxl.load_workbook(fpath, keep_vba=True)
        except Exception as e:
            print(f"  ERREUR ouverture {fname}: {e}", file=sys.stderr)
            continue

        file_modified = False

        for item in items:
            sn = item['sheet']
            type_name = item['j2']
            ti = type_index[type_name]
            enum_vals = ti['enum']
            sub_vals = ti['submissionValue']

            if sn not in wb.sheetnames:
                print(f"  ⚠ Sheet {sn} non trouvée dans {fname}")
                continue

            ws = wb[sn]

            comp_rule_row = item['comp_rule_row']
            start_row = 5
            needed_rows = len(enum_vals) + 1

            if comp_rule_row and (start_row + needed_rows) > comp_rule_row:
                shift = (start_row + needed_rows) - comp_rule_row
                print(f"  ⚠ {fname}[{sn}] : décalage de {shift} lignes pour faire de la place")
                ws.insert_rows(comp_rule_row, amount=shift)

            ws.cell(row=start_row, column=9, value='enumCode')
            ws.cell(row=start_row, column=10, value='submissionValue')

            for i, val in enumerate(enum_vals):
                row = start_row + 1 + i
                sv = sub_vals.get(val, '')
                ws.cell(row=row, column=9, value=val)
                ws.cell(row=row, column=10, value=sv)

            print(f"  ✓ {fname}[{sn}] : {len(enum_vals)} classes écrites (lignes {start_row+1}-{start_row+len(enum_vals)})")
            file_modified = True

        if file_modified:
            try:
                wb.save(fpath)
                modified_count += 1
                print(f"  ✓ Fichier sauvegardé : {fname}")
            except PermissionError:
                print(f"  ✗ ERREUR: '{fname}' est verrouillé (ouvert dans Excel ?). Fichier ignoré.")

        wb.close()

    print(f"\n  Total fichiers modifiés : {modified_count}")


def set_zaxis_for_sheet(sheet_name, type_name, type_index, base_dir, domain):
    """
    Renseigne le Z-axis complet pour une sheet donnée :
      - J2 = code (type_name)
      - J4 = associatedAttribute (<domain>.<type_name>)
      - Lignes 5+ = enum values + submissionValue
    """
    if type_name not in type_index:
        print(f"  ERREUR: Type '{type_name}' non trouvé dans le dictionnaire!")
        print(f"  Types disponibles contenant 'expcla' ou 'exposure' :")
        for k in sorted(type_index.keys()):
            if 'expcla' in k.lower() or 'exposure' in k.lower() or 'typspe' in k.lower():
                print(f"    - {k} ({len(type_index[k]['enum'])} valeurs)")
        return False

    ti = type_index[type_name]
    enum_vals = ti['enum']
    sub_vals = ti['submissionValue']
    attribute = f"{domain}.{type_name}"

    target_file = find_sheet_in_xlsm(sheet_name, base_dir)
    if not target_file:
        print(f"  ERREUR: Aucun fichier XLSM trouvé contenant la sheet '{sheet_name}'")
        return False

    fname = os.path.basename(target_file)
    print(f"\n  Fichier cible : {fname}")
    print(f"  Sheet cible   : {sheet_name}")
    print(f"  Type Z-axis   : {type_name}")
    print(f"  Attribut      : {attribute}")
    print(f"  Nb classes    : {len(enum_vals)}")

    create_backup(target_file)

    wb = openpyxl.load_workbook(target_file, keep_vba=True)
    ws = wb[sheet_name]

    i1 = ws.cell(row=1, column=9).value
    if i1 != 'zAxis':
        print(f"  ERREUR: I1 = '{i1}', attendu 'zAxis'. Cette sheet n'a pas de zone Z-axis.")
        wb.close()
        return False

    # J2 (code)
    old_j2 = ws.cell(row=2, column=10).value
    ws.cell(row=2, column=10, value=type_name)
    print(f"\n  J2 (code)               : '{old_j2}' → '{type_name}'")

    # J4 (associatedAttribute)
    old_j4 = ws.cell(row=4, column=10).value
    ws.cell(row=4, column=10, value=attribute)
    print(f"  J4 (associatedAttribute): '{old_j4}' → '{attribute}'")

    # Trouver computationRuleOrder et nettoyer ancienne écriture
    comp_rule_row = None
    for row in range(5, ws.max_row + 1):
        if ws.cell(row=row, column=9).value == 'computationRuleOrder':
            comp_rule_row = row
            break
        if ws.cell(row=row, column=9).value == 'enumCode':
            r = row
            while r < ws.max_row and ws.cell(row=r, column=9).value:
                if ws.cell(row=r, column=9).value == 'computationRuleOrder':
                    break
                ws.cell(row=r, column=9, value=None)
                ws.cell(row=r, column=10, value=None)
                r += 1

    comp_rule_row = None
    for row in range(5, ws.max_row + 1):
        if ws.cell(row=row, column=9).value == 'computationRuleOrder':
            comp_rule_row = row
            break

    start_row = 5
    needed_rows = len(enum_vals) + 1

    if comp_rule_row and (start_row + needed_rows) > comp_rule_row:
        shift = (start_row + needed_rows) - comp_rule_row
        print(f"  Décalage de {shift} lignes pour faire de la place")
        ws.insert_rows(comp_rule_row, amount=shift)

    ws.cell(row=start_row, column=9, value='enumCode')
    ws.cell(row=start_row, column=10, value='submissionValue')

    for i, val in enumerate(enum_vals):
        row = start_row + 1 + i
        sv = sub_vals.get(val, '')
        ws.cell(row=row, column=9, value=val)
        ws.cell(row=row, column=10, value=sv)

    print(f"\n  ✓ {len(enum_vals)} classes écrites (lignes {start_row+1}-{start_row+len(enum_vals)}) :")
    for i, val in enumerate(enum_vals, 1):
        sv = sub_vals.get(val, '—')
        print(f"    {i:3d}. {val:20s} → {sv}")

    wb.save(target_file)
    print(f"\n  ✓ Fichier sauvegardé : {fname}")
    wb.close()
    return True


# ══════════════════════════════════════════════════════════════════════════════
#  PARTIE DATATYPE : parsing DPM, mapping, scan, rapport, écriture
# ══════════════════════════════════════════════════════════════════════════════

def parse_dpm_cell(value):
    """
    Parse une cellule DPM contenant un data point.
    Format: 'dataPointId_x000D_\\nunit_x000D_\\nsign'
    Retourne: (data_point_id, unit, sign)
    """
    if not value:
        return None, None, None
    s = str(value)
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
    Format: '(qHNJ) Exposure value'
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

        columns = {}
        label_row = code_row - 1
        for col in range(first_data_col, first_data_col + 30):
            code_val = ws.cell(row=code_row, column=col).value
            if not code_val:
                continue
            code = str(code_val).strip()
            if not re.match(r'^0\d{2,4}$', code):
                continue

            label = None
            for lr in range(label_row, max(label_row - 3, 3), -1):
                lv = ws.cell(row=lr, column=col).value
                if lv:
                    label = str(lv).strip()
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


# ── Mapping DPM → Dictionnaire (datatypes) ────────────────────────────────────

KNOWN_PROPERTY_MAP = {
    # Montants (€£$)
    'qHNJ': 'positiveOrNegativeAmount',   # Exposure value
    'qINJ': 'positiveOrNegativeAmount',   # Risk weighted exposure amount (RWEA)

    # Pourcentages pondérés (%) — exposure weighted averages → averwlgd (rate_2_4Dec)
    'qILL': 'averwlgd',        # Exposure weighted average PD (%)
    'qIIJ': 'averwlgd',        # Exposure weighted average LGD (%)
    'qNNM': 'averwlgd',        # Density of risk weighted exposure amounts (%)

    # Pourcentages simples (%) — ratios non-pondérés → ratio9Dec
    'qNNL': 'ratio9Dec',       # Default rate

    # Entiers (#)
    'qFGE': 'nboblbytyp',      # Number of obligors (by type)
    'qGCF': 'nboblbytyp',      # Number of obligors defaulted (by type)
    'qGCA': 'mtrtyValueYears',  # Maturity (years)

    # Texte
    'PDT':  'alpha_255',       # PD Range (Typed)
    'qHIL': 'alpha_255',       # External Rating Equivalent
}

UNIT_FALLBACK = {
    '€£$': 'positiveOrNegativeAmount',
    '%':   'averwlgd',          # Par défaut les % EBA sont des moyennes pondérées
    '#':   'nboblbytyp',
    'text': 'alpha_255',
}

SIGN_REFINEMENT = {
    'positive': 'positiveAmount',
    'negative': 'negativeOrNullAmount',
    None:       'positiveOrNegativeAmount',
}

LABEL_KEYWORDS = {
    'maturity': 'mtrtyValueYears',
    'maturité': 'mtrtyValueYears',
    'obligors': 'nboblbytyp',
    'obligés': 'nboblbytyp',
    'number of obligors': 'nboblbytyp',
    'probability of default': 'averwlgd',
    'density': 'averwlgd',
    'default rate': 'ratio9Dec',
    'average pd': 'averwlgd',
    'average lgd': 'averwlgd',
    'weighted average': 'averwlgd',
}


def determine_datatype(col_info, def_index, domain):
    """
    Détermine le datatype PURE pour une colonne DPM.
    Retourne: (dict_name, confidence, source) ou (None, 0, reason)
    """
    unit = col_info.get('unit')
    sign = col_info.get('sign')
    label = col_info.get('label', '') or ''
    mp_code = col_info.get('main_property_code')
    mp_label = col_info.get('main_property_label', '') or ''

    if mp_code and mp_code in KNOWN_PROPERTY_MAP:
        name = KNOWN_PROPERTY_MAP[mp_code]
        if name and name in def_index:
            return name, 3, f"Main Property ({mp_code})"
        elif name is None:
            return None, 3, f"Main Property ({mp_code}) → défaut"

    label_lower = label.lower()
    mp_label_lower = mp_label.lower()
    for keyword, dict_name in LABEL_KEYWORDS.items():
        if keyword in label_lower or keyword in mp_label_lower:
            if dict_name in def_index:
                return dict_name, 2, f"Label '{keyword}'"

    if unit in UNIT_FALLBACK:
        name = UNIT_FALLBACK[unit]
        if name and name in def_index:
            if unit == '#' and ('matur' in label_lower or 'year' in label_lower or 'année' in label_lower):
                return 'mtrtyValueYears', 1, "Unit # + label maturity"
            return name, 1, f"Unit type '{unit}'"

    return None, 0, "Aucun mapping trouvé"


def scan_xlsm_datatypes(base_dir):
    """
    Scanne tous les pure-form XLSM et retourne les datatypes par colonne.
    Retourne: [{file, file_path, sheet, d2_datatype, ident_datatype, col_datatypes, col_labels, col_codes}]
    """
    results = []

    for fpath in find_xlsm_files(base_dir):
        fname = os.path.basename(fpath)
        try:
            wb = openpyxl.load_workbook(fpath, keep_vba=True)
        except Exception as e:
            print(f"  ERREUR chargement {fname}: {e}", file=sys.stderr)
            continue

        ident_datatype = None
        if 'identification' in wb.sheetnames:
            ws_id = wb['identification']
            ident_datatype = ws_id.cell(row=2, column=10).value

        for sn in wb.sheetnames:
            if sn in SKIP_SHEETS:
                continue
            ws = wb[sn]

            header_row = find_header_row(ws)
            if header_row is None:
                continue

            d2 = ws.cell(row=2, column=4).value

            col_datatypes = {}
            col_labels = {}
            col_codes = []
            for col in range(14, 60):
                code = ws.cell(row=header_row, column=col).value
                if not code:
                    continue
                code = str(code).strip()
                if not re.match(r'^0\d{2,4}$', code):
                    continue
                col_codes.append(code)

                dt = ws.cell(row=7, column=col).value
                if dt:
                    col_datatypes[code] = str(dt).strip()

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


def match_dpm_to_form(dpm_columns, form_items):
    """
    Pour chaque form sheet, trouve le DPM sheet correspondant.
    Retourne: [(form_item, dpm_sheet_name, dpm_cols)]
    """
    matches = []
    dpm_sheets = set(dpm_columns.keys())

    for item in form_items:
        sheet_name = item['sheet']

        if sheet_name in dpm_sheets:
            matches.append((item, sheet_name, dpm_columns[sheet_name]))
            continue

        found = False
        for dpm_sn in dpm_sheets:
            if sheet_name.replace('.', '_') == dpm_sn.replace('.', '_'):
                matches.append((item, dpm_sn, dpm_columns[dpm_sn]))
                found = True
                break
        if not found:
            matches.append((item, None, None))

    return matches


def print_datatype_report(matches, def_index, domain):
    """Affiche un rapport des datatypes à renseigner."""
    print("\n" + "=" * 90)
    print("RAPPORT DES DATATYPES PAR COLONNE")
    print("=" * 90)

    total_missing = 0
    total_filled = 0
    total_to_set = 0

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

    ident_missing = set()
    for form_item, _, _ in matches:
        if not form_item.get('ident_datatype'):
            ident_missing.add(form_item['file'])

    print(f"\n{'=' * 90}")
    print("RÉSUMÉ DATATYPES")
    print(f"{'=' * 90}")
    if ident_missing:
        print(f"  Formulaires sans datatype identification : {len(ident_missing)}")
        for f in sorted(ident_missing):
            print(f"    → {f}")
    print(f"  Colonnes avec datatype renseigné : {total_filled}")
    print(f"  Colonnes à renseigner (suggéré)  : {total_to_set}")
    print(f"  Colonnes sans mapping            : {total_missing}")


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
        create_backup(fpath)

        try:
            wb = openpyxl.load_workbook(fpath, keep_vba=True)
        except Exception as e:
            print(f"  ERREUR ouverture {fname}: {e}", file=sys.stderr)
            continue

        file_modified = False

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

            header_row = find_header_row(ws)
            if not header_row:
                continue

            for col in range(14, 60):
                code = ws.cell(row=header_row, column=col).value
                if not code:
                    continue
                code = str(code).strip()

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


def set_datatype_for_column(sheet_name, col_code, dict_name, def_index, domain, base_dir):
    """Renseigne le datatype pour une colonne spécifique d'une sheet."""
    if dict_name not in def_index:
        print(f"  ERREUR: '{dict_name}' non trouvé dans le dictionnaire!")
        candidates = [n for n in sorted(def_index.keys()) if dict_name.lower() in n.lower()]
        if candidates:
            print(f"  Entrées proches :")
            for c in candidates[:10]:
                info = def_index[c]
                print(f"    - {c} (datatypeId={info.get('datatypeId', '?')}, formatType={info.get('formatType', '?')})")
        return False

    full_name = f"{domain}.{dict_name}"

    target_file = find_sheet_in_xlsm(sheet_name, base_dir)
    if not target_file:
        print(f"  ERREUR: Aucun fichier XLSM contenant la sheet '{sheet_name}'")
        return False

    fname = os.path.basename(target_file)
    create_backup(target_file)

    wb = openpyxl.load_workbook(target_file, keep_vba=True)
    ws = wb[sheet_name]

    header_row = find_header_row(ws)
    if not header_row:
        print(f"  ERREUR: Impossible de trouver la ligne d'en-tête dans {sheet_name}")
        wb.close()
        return False

    target_col = None
    for col in range(14, 60):
        v = ws.cell(row=header_row, column=col).value
        if v and str(v).strip() == col_code:
            target_col = col
            break

    if target_col is None:
        print(f"  ERREUR: Code colonne '{col_code}' non trouvé dans {sheet_name}")
        wb.close()
        return False

    old_val = ws.cell(row=7, column=target_col).value
    ws.cell(row=7, column=target_col, value=full_name)
    print(f"  ✓ {fname}[{sheet_name}] col {col_code} ({get_column_letter(target_col)}7): '{old_val}' → '{full_name}'")

    wb.save(target_file)
    print(f"  ✓ Fichier sauvegardé : {fname}")
    wb.close()
    return True


# ══════════════════════════════════════════════════════════════════════════════
#  GUI UNIFIÉE
# ══════════════════════════════════════════════════════════════════════════════

def run_interactive(base_dir):
    """
    Mode interactif GUI unifié :
      1. Sélection du/des fichier(s) DPM
      2. Sélection du dictionnaire
      3. Onglet Z-axis + onglet Datatypes
      4. Écriture combinée
    """
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk

    root = tk.Tk()
    root.withdraw()

    # ── Étape 1 : Sélection du/des fichier(s) DPM ───────────────────────────
    messagebox.showinfo(
        "Enrichment Tool",
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
        "Enrichment Tool",
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

    # ── Chargement des données ───────────────────────────────────────────────
    print("Chargement du dictionnaire...")
    type_index, def_index = load_dictionary(dict_path)
    print(f"  {len(type_index)} types Z-axis, {len(def_index)} définitions datatypes")

    # Z-axis : lecture DPM
    print("Lecture des Z-axis depuis les DPM...")
    ann_zaxis = {}
    for lp in layout_paths:
        partial = read_annotated_layout(lp)
        for sheet_name, info in partial.items():
            if sheet_name not in ann_zaxis:
                ann_zaxis[sheet_name] = info
    print(f"  {len(ann_zaxis)} sheets avec Z-axis trouvées dans les DPM")

    # Datatypes : lecture DPM
    print("Lecture des colonnes DPM pour datatypes...")
    dpm_columns = {}
    for lp in layout_paths:
        partial = read_dpm_columns(lp)
        for sheet_name, cols in partial.items():
            if sheet_name not in dpm_columns:
                dpm_columns[sheet_name] = cols
    print(f"  {len(dpm_columns)} sheets avec colonnes trouvées")

    # Scan XLSM (une seule fois !)
    print("Scan des formulaires XLSM...")
    xlsm_files = find_xlsm_files(base_dir)

    # Z-axis scan
    zaxis_items = scan_xlsm_zaxis(base_dir)
    print(f"  {len(zaxis_items)} sheets avec header zAxis")

    # Datatypes scan
    form_items = scan_xlsm_datatypes(base_dir)
    print(f"  {len(form_items)} sheets de données")

    dt_matches = match_dpm_to_form(dpm_columns, form_items)

    # ── Index sheet → fichier XLSM ──────────────────────────────────────────
    sheet_to_file = {}
    for fpath in xlsm_files:
        try:
            wb = openpyxl.load_workbook(fpath, keep_vba=True)
            for sn in wb.sheetnames:
                sheet_to_file[sn] = fpath
            wb.close()
        except Exception:
            continue

    # ── Préparer les données Z-axis ─────────────────────────────────────────
    zaxis_selectable = []

    for ann_sheet, info in ann_zaxis.items():
        hierarchy = info['hierarchy']
        variant = info['variant']
        type_name = map_variant_to_type(hierarchy, variant, type_index)

        fpath = sheet_to_file.get(ann_sheet)
        current_j2 = None
        h2_value = None
        j3_value = None

        if fpath:
            try:
                wb_tmp = openpyxl.load_workbook(fpath, keep_vba=True)
                if ann_sheet in wb_tmp.sheetnames:
                    ws_tmp = wb_tmp[ann_sheet]
                    current_j2 = ws_tmp.cell(row=2, column=10).value
                    h2_value = ws_tmp.cell(row=2, column=8).value
                    if h2_value:
                        h2_matches = re.findall(
                            r'(' + re.escape(domain) + r'\w*\.' + re.escape(domain) + r'(\w+))', str(h2_value))
                        for full_path, tok in reversed(h2_matches):
                            for tn in type_index:
                                if tn.lower() == tok.lower():
                                    type_name = tn
                                    j3_value = full_path
                                    break
                            else:
                                continue
                            break
                wb_tmp.close()
            except Exception:
                pass

        if not current_j2:
            status = "à renseigner"
            needs_write = True
        elif type_name and current_j2.lower() != type_name.lower():
            status = f"⚠ à corriger ({current_j2} → {type_name})"
            needs_write = True
        else:
            status = "✓ OK"
            needs_write = False

        zaxis_selectable.append({
            'sheet': ann_sheet,
            'file': fpath,
            'type_name': type_name,
            'source': 'DPM',
            'current_j2': current_j2,
            'status': status,
            'needs_write': needs_write,
            'h2': h2_value,
            'j3': j3_value,
        })

    # Sheets XLSM avec zAxis NON dans le DPM
    dpm_zaxis_sheets = set(ann_zaxis.keys())
    for fpath in xlsm_files:
        try:
            wb = openpyxl.load_workbook(fpath, keep_vba=True)
            for sn in wb.sheetnames:
                if sn in SKIP_SHEETS or sn in dpm_zaxis_sheets:
                    continue
                ws = wb[sn]
                if ws.cell(row=1, column=9).value != 'zAxis':
                    continue
                j2 = ws.cell(row=2, column=10).value
                j4 = ws.cell(row=4, column=10).value
                tn = None
                if j2:
                    if j2 in type_index:
                        tn = j2
                    else:
                        for t in type_index:
                            if t.lower() == j2.lower():
                                tn = t
                                break
                if not tn and j4:
                    m = re.search(re.escape(domain) + r'\.(\w+)', str(j4))
                    if m and m.group(1) in type_index:
                        tn = m.group(1)
                if not j2:
                    status = "à renseigner"
                    needs_w = True
                elif tn and j2.lower() != tn.lower():
                    status = f"⚠ à corriger ({j2} → {tn})"
                    needs_w = True
                else:
                    status = "✓ OK"
                    needs_w = False
                zaxis_selectable.append({
                    'sheet': sn, 'file': fpath, 'type_name': tn,
                    'source': 'XLSM', 'current_j2': j2, 'status': status,
                    'needs_write': needs_w,
                })
            wb.close()
        except Exception:
            continue

    # ── Préparer les données datatypes ──────────────────────────────────────
    dt_table_data = []
    for form_item, dpm_sheet, dpm_cols in dt_matches:
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
            dt_table_data.append({
                'file': form_item['file'], 'file_path': form_item['file_path'],
                'sheet': form_item['sheet'], 'col_code': code, 'label': label,
                'existing': existing, 'suggested': full_suggested,
                'suggested_name': suggested, 'confidence': confidence,
                'source': source, 'unit': unit, 'form_item': form_item,
            })

    # ── Étape 3 : Fenêtre à onglets ─────────────────────────────────────────
    sel_window = tk.Toplevel(root)
    sel_window.title("Enrichment Tool — Z-axis & Datatypes")
    sel_window.geometry("1200x700")

    title_frame = tk.Frame(sel_window, padx=10, pady=5)
    title_frame.pack(fill=tk.X)
    tk.Label(title_frame,
             text="Étape 3/3 : Sélectionnez les éléments à écrire",
             font=("Segoe UI", 12, "bold")).pack(anchor=tk.W)
    dpm_names = ', '.join(os.path.basename(lp) for lp in layout_paths)
    tk.Label(title_frame, text=f"DPM : {dpm_names}",
             font=("Segoe UI", 9), fg="gray").pack(anchor=tk.W)

    notebook = ttk.Notebook(sel_window)
    notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

    # ── Onglet Z-axis ───────────────────────────────────────────────────────
    zaxis_frame = tk.Frame(notebook)
    notebook.add(zaxis_frame, text=f" Z-axis ({len(zaxis_selectable)} sheets) ")

    z_list_frame = tk.Frame(zaxis_frame, padx=5, pady=5)
    z_list_frame.pack(fill=tk.BOTH, expand=True)

    z_canvas = tk.Canvas(z_list_frame)
    z_scrollbar = tk.Scrollbar(z_list_frame, orient=tk.VERTICAL, command=z_canvas.yview)
    z_inner = tk.Frame(z_canvas)
    z_inner.bind("<Configure>", lambda e: z_canvas.configure(scrollregion=z_canvas.bbox("all")))
    z_canvas.create_window((0, 0), window=z_inner, anchor=tk.NW)
    z_canvas.configure(yscrollcommand=z_scrollbar.set)
    z_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    z_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    z_headers = ["", "Sheet", "Type", "Source", "Statut", "Fichier"]
    for col_idx, h in enumerate(z_headers):
        tk.Label(z_inner, text=h, font=("Segoe UI", 9, "bold"),
                 borderwidth=1, relief="solid", padx=5, pady=2
                 ).grid(row=0, column=col_idx, sticky="ew")

    z_check_vars = []
    for i, item in enumerate(zaxis_selectable):
        can_select = item['file'] is not None and item['type_name'] is not None
        var = tk.BooleanVar(value=can_select and item.get('needs_write', True))
        z_check_vars.append(var)

        cb = tk.Checkbutton(z_inner, variable=var,
                            state=tk.NORMAL if can_select else tk.DISABLED)
        cb.grid(row=i + 1, column=0)

        fname = os.path.basename(item['file']) if item['file'] else "—"
        type_str = item['type_name'] or "?"
        fg = "black" if can_select else "gray"

        tk.Label(z_inner, text=item['sheet'], fg=fg, anchor="w"
                 ).grid(row=i + 1, column=1, sticky="w", padx=5)
        tk.Label(z_inner, text=type_str, fg=fg, anchor="w"
                 ).grid(row=i + 1, column=2, sticky="w", padx=5)
        tk.Label(z_inner, text=item['source'], fg=fg, anchor="w"
                 ).grid(row=i + 1, column=3, sticky="w", padx=5)
        tk.Label(z_inner, text=item['status'], fg=fg, anchor="w"
                 ).grid(row=i + 1, column=4, sticky="w", padx=5)
        tk.Label(z_inner, text=fname, fg=fg, anchor="w"
                 ).grid(row=i + 1, column=5, sticky="w", padx=5)

    # ── Onglet Datatypes ─────────────────────────────────────────────────────
    dt_frame = tk.Frame(notebook)
    notebook.add(dt_frame, text=f" Datatypes ({len(dt_table_data)} colonnes) ")

    # Filtre
    dt_filter_frame = tk.Frame(dt_frame, padx=10, pady=3)
    dt_filter_frame.pack(fill=tk.X)
    tk.Label(dt_filter_frame, text="Filtre :").pack(side=tk.LEFT)
    dt_filter_var = tk.StringVar(value="à renseigner")
    for txt, val in [("Tout", "all"), ("À renseigner", "à renseigner"), ("Déjà renseigné", "renseigné")]:
        tk.Radiobutton(dt_filter_frame, text=txt, variable=dt_filter_var, value=val).pack(side=tk.LEFT, padx=5)

    # Treeview
    dt_tree_frame = tk.Frame(dt_frame, padx=10, pady=5)
    dt_tree_frame.pack(fill=tk.BOTH, expand=True)

    dt_columns = ("select", "file", "sheet", "colcode", "label", "unit", "existing", "suggested", "conf", "source")
    dt_tree = ttk.Treeview(dt_tree_frame, columns=dt_columns, show='headings', selectmode='extended')

    dt_tree.heading("select", text="✓")
    dt_tree.heading("file", text="Fichier")
    dt_tree.heading("sheet", text="Sheet")
    dt_tree.heading("colcode", text="Col")
    dt_tree.heading("label", text="Label")
    dt_tree.heading("unit", text="Unit")
    dt_tree.heading("existing", text="Existant")
    dt_tree.heading("suggested", text="Suggéré")
    dt_tree.heading("conf", text="Fiab.")
    dt_tree.heading("source", text="Source")

    dt_tree.column("select", width=30, anchor="center")
    dt_tree.column("file", width=180)
    dt_tree.column("sheet", width=100)
    dt_tree.column("colcode", width=50, anchor="center")
    dt_tree.column("label", width=200)
    dt_tree.column("unit", width=40, anchor="center")
    dt_tree.column("existing", width=150)
    dt_tree.column("suggested", width=170)
    dt_tree.column("conf", width=35, anchor="center")
    dt_tree.column("source", width=140)

    dt_scroll_y = ttk.Scrollbar(dt_tree_frame, orient=tk.VERTICAL, command=dt_tree.yview)
    dt_scroll_x = ttk.Scrollbar(dt_tree_frame, orient=tk.HORIZONTAL, command=dt_tree.xview)
    dt_tree.configure(yscrollcommand=dt_scroll_y.set, xscrollcommand=dt_scroll_x.set)

    dt_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    dt_scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
    dt_scroll_x.pack(side=tk.BOTTOM, fill=tk.X)

    dt_selected_set = set()

    def dt_populate_tree(filter_mode="all"):
        dt_tree.delete(*dt_tree.get_children())
        for i, row in enumerate(dt_table_data):
            if filter_mode == "à renseigner" and row['existing']:
                continue
            if filter_mode == "renseigné" and not row['existing']:
                continue
            if not row['suggested']:
                continue
            conf_str = ['?', '~', '✓', '★'][min(row['confidence'], 3)]
            sel_str = "☑" if i in dt_selected_set else "☐"
            tag = 'selected' if i in dt_selected_set else ''
            dt_tree.insert("", tk.END, iid=str(i), values=(
                sel_str, row['file'], row['sheet'], row['col_code'],
                row['label'], row['unit'], row['existing'] or '',
                row['suggested'], conf_str, row['source'],
            ), tags=(tag,))
        dt_tree.tag_configure('selected', background='#d4edda')

    def dt_toggle(event):
        item = dt_tree.focus()
        if not item:
            return
        idx = int(item)
        if idx in dt_selected_set:
            dt_selected_set.discard(idx)
        else:
            dt_selected_set.add(idx)
        dt_populate_tree(dt_filter_var.get())

    # Auto-sélection
    for i, row in enumerate(dt_table_data):
        if not row['existing'] and row['suggested'] and row['confidence'] >= 1:
            dt_selected_set.add(i)

    dt_populate_tree("à renseigner")
    dt_tree.bind('<Double-1>', dt_toggle)
    dt_filter_var.trace_add('write', lambda *_: dt_populate_tree(dt_filter_var.get()))

    # ── Boutons communs ──────────────────────────────────────────────────────
    btn_frame = tk.Frame(sel_window, padx=10, pady=10)
    btn_frame.pack(fill=tk.X)

    count_label = tk.Label(btn_frame, text="", font=("Segoe UI", 9))
    count_label.pack(side=tk.LEFT, padx=10)

    def update_count():
        nz = sum(1 for v in z_check_vars if v.get())
        nd = len(dt_selected_set)
        count_label.config(text=f"Z-axis: {nz} sheet(s) | Datatypes: {nd} colonne(s)")
        sel_window.after(500, update_count)

    update_count()

    def z_select_all():
        for i, var in enumerate(z_check_vars):
            item = zaxis_selectable[i]
            if item['file'] and item['type_name']:
                var.set(True)

    def z_deselect_all():
        for var in z_check_vars:
            var.set(False)

    def dt_select_all_visible():
        for item in dt_tree.get_children():
            idx = int(item)
            if dt_table_data[idx]['suggested'] and not dt_table_data[idx]['existing']:
                dt_selected_set.add(idx)
        dt_populate_tree(dt_filter_var.get())

    def dt_deselect_all():
        dt_selected_set.clear()
        dt_populate_tree(dt_filter_var.get())

    result_action = [None]

    def on_write():
        result_action[0] = 'write'
        sel_window.destroy()

    def on_cancel():
        sel_window.destroy()

    tk.Button(btn_frame, text="Z: Tout sélectionner",
              command=z_select_all).pack(side=tk.LEFT, padx=3)
    tk.Button(btn_frame, text="Z: Désélectionner",
              command=z_deselect_all).pack(side=tk.LEFT, padx=3)
    tk.Button(btn_frame, text="DT: Tout sélectionner",
              command=dt_select_all_visible).pack(side=tk.LEFT, padx=3)
    tk.Button(btn_frame, text="DT: Désélectionner",
              command=dt_deselect_all).pack(side=tk.LEFT, padx=3)

    tk.Button(btn_frame, text="Annuler",
              command=on_cancel).pack(side=tk.RIGHT, padx=5)
    tk.Button(btn_frame, text="Écrire tout", command=on_write,
              bg="#4CAF50", fg="white",
              font=("Segoe UI", 10, "bold")).pack(side=tk.RIGHT, padx=5)

    sel_window.wait_window()

    if result_action[0] != 'write':
        print("Annulé.")
        root.destroy()
        return

    # ── Écriture Z-axis ──────────────────────────────────────────────────────
    z_selected = [zaxis_selectable[i] for i, v in enumerate(z_check_vars) if v.get()]
    z_written = 0
    if z_selected:
        print(f"\n{'=' * 80}")
        print(f"ÉCRITURE Z-AXIS ({len(z_selected)} sheets)")
        print(f"{'=' * 80}")

        z_by_file = {}
        for item in z_selected:
            z_by_file.setdefault(item['file'], []).append(item)

        for fpath, items in z_by_file.items():
            fname = os.path.basename(fpath)
            create_backup(fpath)

            try:
                wb = openpyxl.load_workbook(fpath, keep_vba=True)
            except PermissionError:
                print(f"  ✗ '{fname}' verrouillé (ouvert dans Excel ?)")
                continue
            except Exception as e:
                print(f"  ✗ Erreur ouverture {fname}: {e}")
                continue

            for item in items:
                sn = item['sheet']
                tn = item['type_name']
                if sn not in wb.sheetnames:
                    print(f"  ⚠ Sheet '{sn}' non trouvée dans {fname}")
                    continue

                ws = wb[sn]
                attribute = f"{domain}.{tn}"
                has_h2 = item.get('h2') is not None
                j3_val = item.get('j3')

                j2_val = tn.lower() if has_h2 else tn

                old_j2 = ws.cell(row=2, column=10).value
                old_j3 = ws.cell(row=3, column=10).value
                old_j4 = ws.cell(row=4, column=10).value

                ws.cell(row=2, column=10, value=j2_val)
                ws.cell(row=4, column=10, value=attribute)

                if j3_val:
                    ws.cell(row=3, column=10, value=j3_val)

                log = f"  ✓ {fname}[{sn}] : J2='{old_j2}'→'{j2_val}' | J4='{old_j4}'→'{attribute}'"
                if j3_val:
                    log += f" | J3='{old_j3}'→'{j3_val}'"
                print(log)
                z_written += 1

            try:
                wb.save(fpath)
                print(f"  ✓ Fichier sauvegardé : {fname}")
            except PermissionError:
                print(f"  ✗ '{fname}' verrouillé, impossible de sauvegarder")
            wb.close()

    # ── Écriture Datatypes ───────────────────────────────────────────────────
    dt_written = 0
    if dt_selected_set:
        print(f"\n{'=' * 80}")
        print(f"ÉCRITURE DATATYPES ({len(dt_selected_set)} colonnes)")
        print(f"{'=' * 80}")

        dt_by_file = {}
        for idx in sorted(dt_selected_set):
            row = dt_table_data[idx]
            dt_by_file.setdefault(row['file_path'], []).append(row)

        for fpath, rows in dt_by_file.items():
            fname = os.path.basename(fpath)
            create_backup(fpath)

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
                hdr = find_header_row(ws)
                if not hdr:
                    continue
                for col in range(14, 60):
                    v = ws.cell(row=hdr, column=col).value
                    if v and str(v).strip() == row['col_code']:
                        old_val = ws.cell(row=7, column=col).value
                        ws.cell(row=7, column=col, value=row['suggested'])
                        print(f"  ✓ {fname}[{sn}] col {row['col_code']} ({get_column_letter(col)}7): "
                              f"'{old_val}' → '{row['suggested']}'")
                        dt_written += 1
                        break

            try:
                wb.save(fpath)
                print(f"  ✓ Fichier sauvegardé : {fname}")
            except PermissionError:
                print(f"  ✗ '{fname}' verrouillé, impossible de sauvegarder")
            wb.close()

    # ── Résumé ───────────────────────────────────────────────────────────────
    summary = f"Terminé !\nZ-axis : {z_written} sheets\nDatatypes : {dt_written} colonnes"
    print(f"\n{summary}")
    messagebox.showinfo("Enrichment Tool", summary)
    root.destroy()


# ══════════════════════════════════════════════════════════════════════════════
#  MAIN
# ══════════════════════════════════════════════════════════════════════════════

def main():
    parser = argparse.ArgumentParser(
        description="Outil unifié d'enrichissement Z-axis + Datatypes pour pure-form XLSM"
    )
    parser.add_argument('--write', action='store_true',
                        help="Écrire les valeurs dans les fichiers XLSM")
    parser.add_argument('--zaxis-only', action='store_true',
                        help="Traiter uniquement les Z-axis")
    parser.add_argument('--datatype-only', action='store_true',
                        help="Traiter uniquement les datatypes")
    parser.add_argument('--set-zaxis', metavar='TYPE',
                        help="Renseigner J2/J4/enum pour une sheet. Ex: --set-zaxis typspelendingexp")
    parser.add_argument('--set-datatype', metavar='TYPE',
                        help="Renseigner un datatype pour une colonne. Ex: --set-datatype ratio9Dec")
    parser.add_argument('--sheet', metavar='SHEET',
                        help="Sheet cible pour --set-zaxis ou --set-datatype")
    parser.add_argument('--col', metavar='CODE',
                        help="Code colonne pour --set-datatype. Ex: --col 0020")
    parser.add_argument('--dir', default=BASE_DIR,
                        help=f"Répertoire de travail (défaut: {BASE_DIR})")
    parser.add_argument('--gui', action='store_true',
                        help="Mode interactif GUI unifié")
    args = parser.parse_args()

    base_dir = args.dir

    # ── Mode GUI ──
    if args.gui:
        run_interactive(base_dir)
        return

    # ── Trouver et charger le dictionnaire ──
    dict_files = glob.glob(os.path.join(base_dir, 'pure-dictionary-*.yml'))
    if not dict_files:
        print(f"ERREUR: Aucun fichier pure-dictionary-*.yml trouvé dans {base_dir}", file=sys.stderr)
        sys.exit(1)
    dict_path = dict_files[0]

    domain = get_dict_domain(dict_path)

    print("Chargement du dictionnaire...")
    type_index, def_index = load_dictionary(dict_path)
    print(f"  {len(type_index)} types Z-axis, {len(def_index)} définitions datatypes (domaine: {domain})")

    # ── Mode SET Z-axis ──
    if args.set_zaxis:
        if not args.sheet:
            print("ERREUR: --set-zaxis nécessite --sheet.", file=sys.stderr)
            sys.exit(1)
        print(f"\n{'=' * 80}")
        print("MODE SET : Z-axis")
        print(f"{'=' * 80}")
        ok = set_zaxis_for_sheet(args.sheet, args.set_zaxis, type_index, base_dir, domain)
        sys.exit(0 if ok else 1)

    # ── Mode SET Datatype ──
    if args.set_datatype:
        if not args.sheet or not args.col:
            print("ERREUR: --set-datatype nécessite --sheet et --col.", file=sys.stderr)
            sys.exit(1)
        ok = set_datatype_for_column(args.sheet, args.col, args.set_datatype,
                                     def_index, domain, base_dir)
        sys.exit(0 if ok else 1)

    # ── Mode REPORT / WRITE ──
    do_zaxis = not args.datatype_only
    do_datatype = not args.zaxis_only

    if do_zaxis:
        print("\nScan des fichiers pure-form XLSM (Z-axis)...")
        zaxis_items = scan_xlsm_zaxis(base_dir)
        print(f"  {len(zaxis_items)} sheets avec header zAxis trouvées")
        print_zaxis_report(zaxis_items, type_index)

        if args.write:
            print(f"\n{'=' * 80}")
            print("ÉCRITURE DES CLASSES Z-AXIS")
            print(f"{'=' * 80}")
            write_zaxis_values(zaxis_items, type_index, base_dir)

    if do_datatype:
        print("\nLecture des DPM (datatypes)...")
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

        print("Scan des formulaires XLSM (datatypes)...")
        form_items = scan_xlsm_datatypes(base_dir)
        print(f"  {len(form_items)} sheets de données trouvées")

        if form_items:
            matches = match_dpm_to_form(dpm_columns, form_items)
            print_datatype_report(matches, def_index, domain)

            if args.write:
                print(f"\n{'=' * 90}")
                print("ÉCRITURE DES DATATYPES")
                print(f"{'=' * 90}")
                write_datatypes(matches, def_index, domain, base_dir)

    if not args.write:
        print(f"\n💡 Pour écrire, relancer avec : python enrichment_tool.py --write")
        if not args.zaxis_only and not args.datatype_only:
            print(f"   Ou : --write --zaxis-only  /  --write --datatype-only")


if __name__ == '__main__':
    main()
