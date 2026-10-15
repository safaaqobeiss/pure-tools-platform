#!/usr/bin/env python3
"""
PURE Unified Pipeline: Generation + Enrichment + RowKey
=======================================================

All-in-one script combining:
  - Z-axis + Datatypes enrichment
  - RowKey filling for Open Rows
  - Orchestration pipeline with forms_from_dpm

Modes:
  python pure_form_generator.py                                    # Full pipeline GUI
  python pure_form_generator.py enrichment                         # Enrichment report
  python pure_form_generator.py enrichment --gui                   # Enrichment GUI
  python pure_form_generator.py enrichment --write                 # Write Z-axis + datatypes
  python pure_form_generator.py enrichment --set-zaxis TYPE --sheet SHEET
  python pure_form_generator.py enrichment --set-datatype TYPE --sheet SHEET --col CODE
  python pure_form_generator.py rowkey                             # RowKey report
  python pure_form_generator.py rowkey --gui                       # RowKey GUI
  python pure_form_generator.py rowkey --write                     # Write rowKey
  python pure_form_generator.py rowkey --set-rowkey "formula" --sheet SHEET
"""

import argparse
import glob
import logging
import os
import re
import shutil
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, Label, Entry, Button, Checkbutton, BooleanVar
import yaml
import openpyxl
from openpyxl.utils import get_column_letter

import forms_from_dpm

# -- Configuration --------------------------------------------------------------

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SKIP_SHEETS = {'identification', 'cellSpecificities', 'labelTranslation', 'TOC'}


######################################################################################################################################################
#                                                                                                                                                    #
#       Part 1: Common utilities — dictionary loading, file discovery, backup, header detection                              #
#                                                                                                                                                    #
#   load_dictionary: loads YML and builds type_index (enumerations) + def_index (definitions + datatypeId)                                    #
#   get_dict_domain: extracts domain from dictionary file path                                                                          #
#   find_xlsm_files: finds all pure-form-*.xlsm files in the directory                                                                   #
#   create_backup: creates a .bak copy before any modification                                                                                     #
#   find_header_row: finds the header row (row 1, col 1 not empty) in a sheet                                                               #
#   find_sheet_in_xlsm: locates a given sheet across all XLSM files                                                                   #
#                                                                                                                                                    #
######################################################################################################################################################

def load_dictionary(dict_path):
    """Loads the YML dictionary and builds two indexes:
    - type_index : types with enum/submissionValue (for Z-axis)
    - def_index  : all definitions with datatypeId/formatType (for datatypes)
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

            # Definitions index (for datatypes)
            if name not in def_index:
                def_index[name] = dict(entry)
                def_index[name]['_section'] = section_key

            # Index of types with enum (for Z-axis)
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
    """Extracts the dictionary domain (e.g. euebacorep) from the file name."""
    fname = os.path.basename(dict_path)
    m = re.search(r'pure-dictionary-(\w+)', fname)
    if not m:
        raise ValueError(f"Cannot extract domain from filename: {fname}\n"
                         f"File must follow the pattern: pure-dictionary-<domain>.yml")
    return m.group(1)


def find_xlsm_files(base_dir):
    """Finds all valid pure-form XLSM files."""
    xlsm_files = sorted(glob.glob(os.path.join(base_dir, 'pure-form-*.xlsm')))
    return [f for f in xlsm_files
            if '~$' not in os.path.basename(f)
            and ' copy' not in os.path.basename(f).lower()]


def create_backup(fpath):
    """Creates a .bak backup if it does not already exist."""
    backup_path = fpath + '.bak'
    if not os.path.exists(backup_path):
        shutil.copy2(fpath, backup_path)
        print(f"  Backup created: {os.path.basename(backup_path)}")
    return backup_path


def find_header_row(ws):
    """Finds the header row (where A='level') between rows 15 and 30."""
    for r in range(15, 31):
        if ws.cell(row=r, column=1).value == 'level':
            return r
    return None


def find_sheet_in_xlsm(sheet_name, base_dir):
    """Finds the XLSM file containing a given sheet."""
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


######################################################################################################################################################
#                                                                                                                                                    #
#       Part 2: Z-Axis enrichment — DPM reading, variant mapping, scan, report, writing                                                  #
#                                                                                                                                                    #
#   read_annotated_layout: reads DPM and extracts Z-axis variants per sheet                                                                     #
#   map_variant_to_type: maps a variant (hierarchy + code) to a dictionary type                                                        #
#   scan_xlsm_zaxis: scans XLSM files to find empty or fillable zAxis cells                                                             #
#   print_zaxis_report: displays Z-axis report with statuses (resolved, NOT_RESOLVED, etc.)                                                         #
#   write_zaxis_values: writes resolved Z-axis values to XLSM files                                                                    #
#   set_zaxis_for_sheet: writes Z-axis for a specific sheet (single CLI mode)                                                               #
#                                                                                                                                                    #
######################################################################################################################################################

def read_annotated_layout(layout_path):
    """
    Reads a DPM file (Annotated Table Layout .xlsx) and returns sheets
    with Z-axis.
    Returns: {sheet_name: {'hierarchy': str, 'variant': str}}
    """
    if not os.path.exists(layout_path):
        print(f"  ERROR: DPM file not found: {layout_path}", file=sys.stderr)
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
    Maps a DPM variant code (e.g. qAE5, new_TA9, EN2) to a dictionary type.
    Uses submissionValue entries with eba_<hierarchy>: prefix to find candidates.
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

    # Extract variant identifier (e.g. qAE5->"5", qEC3->"3", new_TA9->"9")
    clean_variant = variant.replace('new_', '')
    m = re.search(r'(\d+)$', clean_variant)
    if m:
        n = m.group(1)
        # Search by multiple patterns : "sub5", "qEC3", "EC3", etc.
        patterns = [f"sub{n}", f"{hierarchy}{n}", clean_variant.lower()]
        # Deduplicate while preserving order
        seen = set()
        unique_patterns = []
        for p in patterns:
            p_lower = p.lower()
            if p_lower not in seen:
                seen.add(p_lower)
                unique_patterns.append(p_lower)

        for pattern in unique_patterns:
            matches = [name for name, _ in pure if pattern in name.lower()]
            if matches:
                # Prefer names from 'definitions' (e.g. expcladsub5)
                # rather than 'enumerations' (e.g. exposureClassSub5)
                def_matches = [name for name in matches
                               if type_index[name].get('section') == 'definitions']
                if def_matches:
                    return def_matches[0]
                return matches[0]

    if len(pure) == 1:
        return pure[0][0]

    if pure:
        return pure[0][0]

    return None


def scan_xlsm_zaxis(base_dir):
    """
    Scans all pure-form XLSM files and returns Z-axis info.
    Returns: [{file, file_path, sheet, j2, j4, comp_rule_row, existing_values}]
    """
    results = []

    for fpath in find_xlsm_files(base_dir):
        fname = os.path.basename(fpath)
        try:
            wb = openpyxl.load_workbook(fpath, keep_vba=True)
        except Exception as e:
            print(f"  ERROR loading {fname}: {e}", file=sys.stderr)
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
    """Displays a detailed Z-axis classes report per form."""
    has_zaxis = [item for item in zaxis_items if item['j2']]
    no_zaxis = [item for item in zaxis_items if not item['j2']]

    print("\n" + "=" * 80)
    print("Z-AXIS CLASSES REPORT BY FORM")
    print("=" * 80)

    for item in has_zaxis:
        type_name = item['j2']
        print(f"\n+- File : {item['file']}")
        print(f"|  Sheet   : {item['sheet']}")
        print(f"|  Type    : {type_name}")
        print(f"|  Attribute: {item['j4']}")

        if type_name in type_index:
            ti = type_index[type_name]
            enum_vals = ti['enum']
            sub_vals = ti['submissionValue']
            print(f"|  Number of classes: {len(enum_vals)}")
            print(f"|")
            for i, val in enumerate(enum_vals, 1):
                sv = sub_vals.get(val, '—')
                print(f"|  {i:3d}. {val:20s} -> {sv}")
        else:
            print(f"|  [!] Type '{type_name}' NOT FOUND in dictionary!")

        if item['existing_values']:
            print(f"|")
            print(f"|  Values already present in XLSM: {item['existing_values']}")

        print(f"+{'-' * 70}")

    if no_zaxis:
        print(f"\n{'-' * 80}")
        print(f"Sheets WITHOUT Z-axis (J2 empty): {len(no_zaxis)}")
        for item in no_zaxis:
            print(f"  - {item['file']} -> [{item['sheet']}]")

    print(f"\n{'=' * 80}")
    print("Z-AXIS SUMMARY")
    print(f"{'=' * 80}")
    print(f"  Total sheets with Z-axis header: {len(zaxis_items)}")
    print(f"  Sheets with Z-axis set         : {len(has_zaxis)}")
    print(f"  Sheets without Z-axis (normal)  : {len(no_zaxis)}")
    total_classes = sum(
        len(type_index[item['j2']]['enum'])
        for item in has_zaxis
        if item['j2'] in type_index
    )
    print(f"  Total classes to fill           : {total_classes}")


def write_zaxis_values(zaxis_items, type_index, base_dir):
    """Writes Z-axis enum values to XLSM files."""
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
            print(f"  ERROR opening {fname}: {e}", file=sys.stderr)
            continue

        file_modified = False

        for item in items:
            sn = item['sheet']
            type_name = item['j2']
            ti = type_index[type_name]
            enum_vals = ti['enum']
            sub_vals = ti['submissionValue']

            if sn not in wb.sheetnames:
                print(f"  [!] Sheet {sn} not found in {fname}")
                continue

            ws = wb[sn]

            comp_rule_row = item['comp_rule_row']
            start_row = 5
            needed_rows = len(enum_vals) + 1

            if comp_rule_row and (start_row + needed_rows) > comp_rule_row:
                shift = (start_row + needed_rows) - comp_rule_row
                print(f"  [!] {fname}[{sn}] : shifting {shift} rows to make space")
                ws.insert_rows(comp_rule_row, amount=shift)

            ws.cell(row=start_row, column=9, value='enumCode')
            ws.cell(row=start_row, column=10, value='submissionValue')

            for i, val in enumerate(enum_vals):
                row = start_row + 1 + i
                sv = sub_vals.get(val, '')
                ws.cell(row=row, column=9, value=val)
                ws.cell(row=row, column=10, value=sv)

            print(f"  [OK] {fname}[{sn}] : {len(enum_vals)} classes written (rows {start_row+1}-{start_row+len(enum_vals)})")
            file_modified = True

        if file_modified:
            try:
                wb.save(fpath)
                modified_count += 1
                print(f"  [OK] File saved: {fname}")
            except PermissionError:
                print(f"  [X] ERROR: '{fname}' is locked (open in Excel?). File skipped.")

        wb.close()

    print(f"\n  Total files modified: {modified_count}")


def set_zaxis_for_sheet(sheet_name, type_name, type_index, base_dir, domain):
    """
    Sets the complete Z-axis for a given sheet:
      - J2 = code (type_name)
      - J4 = associatedAttribute (<domain>.<type_name>)
      - Rows 5+ = enum values + submissionValue
    """
    if type_name not in type_index:
        print(f"  ERROR: Type '{type_name}' not found in dictionary!")
        print(f"  Available types containing 'expcla' or 'exposure':")
        for k in sorted(type_index.keys()):
            if 'expcla' in k.lower() or 'exposure' in k.lower() or 'typspe' in k.lower():
                print(f"    - {k} ({len(type_index[k]['enum'])} values)")
        return False

    ti = type_index[type_name]
    enum_vals = ti['enum']
    sub_vals = ti['submissionValue']
    attribute = f"{domain}.{type_name}"

    target_file = find_sheet_in_xlsm(sheet_name, base_dir)
    if not target_file:
        print(f"  ERROR: No XLSM file containing sheet '{sheet_name}'")
        return False

    fname = os.path.basename(target_file)
    print(f"\n  Target file : {fname}")
    print(f"  Target sheet  : {sheet_name}")
    print(f"  Z-axis type   : {type_name}")
    print(f"  Attribute     : {attribute}")
    print(f"  Num classes   : {len(enum_vals)}")

    create_backup(target_file)

    wb = openpyxl.load_workbook(target_file, keep_vba=True)
    ws = wb[sheet_name]

    i1 = ws.cell(row=1, column=9).value
    if i1 != 'zAxis':
        print(f"  ERROR: I1 = '{i1}', expected 'zAxis'. This sheet has no Z-axis zone.")
        wb.close()
        return False

    # J2 (code)
    old_j2 = ws.cell(row=2, column=10).value
    ws.cell(row=2, column=10, value=type_name)
    print(f"\n  J2 (code)               : '{old_j2}' -> '{type_name}'")

    # J4 (associatedAttribute)
    old_j4 = ws.cell(row=4, column=10).value
    ws.cell(row=4, column=10, value=attribute)
    print(f"  J4 (associatedAttribute): '{old_j4}' -> '{attribute}'")

    # Find computationRuleOrder and clean up previous writing
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
        print(f"  Shifting {shift} rows to make space")
        ws.insert_rows(comp_rule_row, amount=shift)

    ws.cell(row=start_row, column=9, value='enumCode')
    ws.cell(row=start_row, column=10, value='submissionValue')

    for i, val in enumerate(enum_vals):
        row = start_row + 1 + i
        sv = sub_vals.get(val, '')
        ws.cell(row=row, column=9, value=val)
        ws.cell(row=row, column=10, value=sv)

    print(f"\n  [OK] {len(enum_vals)} classes written (rows {start_row+1}-{start_row+len(enum_vals)}):")
    for i, val in enumerate(enum_vals, 1):
        sv = sub_vals.get(val, '—')
        print(f"    {i:3d}. {val:20s} -> {sv}")

    wb.save(target_file)
    print(f"\n  [OK] File saved: {fname}")
    wb.close()
    return True


######################################################################################################################################################
#                                                                                                                                                    #
#       Part 3: Datatypes enrichment — DPM parsing, generic determination, scan, matching, report, writing                               #
#                                                                                                                                                    #
#   parse_dpm_cell / parse_dpm_property: parsing DPM cells (dataPointId, unit, sign, Main Property)                                          #
#   read_dpm_columns: reads DPM columns with their unit, sign and Main Property per sheet                                                    #
#   _build_dtid_index / determine_datatype: reverse index datatypeId -> name, generic 3-step determination                                     #
#   scan_xlsm_datatypes: scans XLSM files to find columns with empty datatype                                                                 #
#   match_dpm_to_form: cross-references DPM <-> form to match columns                                                                       #
#   print_datatype_report / write_datatypes: detailed report and datatype writing                                                             #
#   set_datatype_for_column: writes datatype for a specific column (single CLI mode)                                                   #
#                                                                                                                                                    #
######################################################################################################################################################

def parse_dpm_cell(value):
    """
    Parses a DPM cell containing a data point.
    Format: 'dataPointId_x000D_\\nunit_x000D_\\nsign'
    Returns: (data_point_id, unit, sign)
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
    Parses a DPM Main Property cell.
    Format: '(qHNJ) Exposure value'
    Returns: (code, label)
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
    Reads a DPM file and extracts column metadata for each sheet.
    Returns: {sheet_name: {col_code: {label, unit, sign, main_property_code, main_property_label}}}
    """
    if not os.path.exists(layout_path):
        print(f"  ERROR: DPM not found: {layout_path}", file=sys.stderr)
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


# -- Mapping DPM -> Dictionary (datatypes) ------------------------------------

def _build_dtid_index(def_index):
    """
    Builds a reverse index datatypeId -> definition name.
    Prefers: 1) name == datatypeId  2) name starting with 'generic'  3) first found.
    """
    dtid_to_def = {}
    for name, info in def_index.items():
        dtid = info.get('datatypeId')
        if not dtid:
            continue
        if dtid not in dtid_to_def:
            dtid_to_def[dtid] = name
        elif dtid == name:
            dtid_to_def[dtid] = name
        elif name.startswith('generic') and not dtid_to_def[dtid].startswith('generic') and dtid_to_def[dtid] != dtid:
            dtid_to_def[dtid] = name
    return dtid_to_def


def determine_datatype(col_info, def_index, domain):
    """
    Determines the PURE datatype for a DPM column.
    Generic approach: unit/sign -> datatypeId -> definition name.
    Works for any EBA project (COREP, FICOD, FINREP, etc.)
    Returns: (dict_name, confidence, source) or (None, 0, reason)
    """
    unit = col_info.get('unit')
    sign = col_info.get('sign')
    mp_label = col_info.get('main_property_label', '') or ''

    dtid_to_def = _build_dtid_index(def_index)

    # 1. Enum columns: Main Property contains [hierarchy:code]
    m = re.search(r'\[(\w+):(\w+)\]', mp_label)
    if m:
        hierarchy = m.group(1)
        enum_ref = f"{hierarchy}:{m.group(2)}"
        prefix = f"eba_{hierarchy}:"
        candidates = []
        for name, info in def_index.items():
            sv = info.get('submissionValue', {})
            if sv and any(str(v).startswith(prefix) for v in sv.values() if v):
                candidates.append(name)
        if len(candidates) == 1:
            return candidates[0], 3, f"Enum [{enum_ref}]"
        elif candidates:
            defs_only = [n for n in candidates if def_index[n].get('section') == 'definitions']
            pool = defs_only or candidates

            # Disambiguate using F1 similarity: col_label + MP label vs definition name
            col_label = col_info.get('label', '') or ''
            mp_clean = re.sub(r'\(ei\d+\)', '', mp_label)  # remove (eiNNNN) only
            mp_clean = re.sub(r'\[.*?\]', '', mp_clean)    # remove [hierarchy:code]

            stop_words = {'of', 'the', 'and', 'for', 'in', 'by', 'per', 'with', 'from', 'to'}
            col_words = set(re.findall(r'[a-zA-Z]{2,}', col_label.lower())) - stop_words
            mp_words = set(re.findall(r'[a-zA-Z]{2,}', mp_clean.lower())) - stop_words

            def _split_name(n):
                clean = re.sub(r'[^a-zA-Z]', '', n)
                parts = re.sub(r'([a-z])([A-Z])', r'\1 \2', clean)
                parts = re.sub(r'([A-Z]+)([A-Z][a-z])', r'\1 \2', parts)
                return set(w.lower() for w in parts.split() if len(w) >= 2) - stop_words

            def _fuzzy_count(name_ws, label_ws):
                count = 0
                for nw in name_ws:
                    for lw in label_ws:
                        if nw == lw or (len(nw) >= 3 and len(lw) >= 3
                                        and (nw.startswith(lw) or lw.startswith(nw))):
                            count += 1
                            break
                return count

            def _f1(common, size1, size2):
                if size1 == 0 or size2 == 0 or common == 0:
                    return 0.0
                r, p = common / size1, common / size2
                return 2 * r * p / (r + p)

            scored = []
            for name in pool:
                nw = _split_name(name)
                if not nw:
                    continue
                col_f1 = _f1(_fuzzy_count(nw, col_words), len(col_words), len(nw))
                mp_f1 = _f1(_fuzzy_count(nw, mp_words), len(mp_words), len(nw))
                score = 0.4 * col_f1 + 0.6 * mp_f1
                scored.append((score, name))

            if scored:
                scored.sort(key=lambda x: x[0], reverse=True)
                if scored[0][0] > 0:
                    return scored[0][1], 2, f"Enum [{enum_ref}] (label match)"

            choice = pool[0]
            return choice, 1, f"Enum [{enum_ref}] ({len(candidates)} candidates)"

    # 2. Metric columns: unit + sign -> datatypeId -> definition name
    target_dtids = []
    unit_lower = (unit or '').lower().strip()

    if unit == '€£$':
        if sign == 'positive':
            target_dtids = ['positiveAmount', 'positiveOrNegativeAmount']
        elif sign == 'negative':
            target_dtids = ['negativeAmount', 'positiveOrNegativeAmount']
        else:
            target_dtids = ['positiveOrNegativeAmount']
    elif unit == '%':
        target_dtids = ['positiveRate_1_4Dec', 'rate_2_4Dec', 'percentage',
                        'positiveRate_1_9Dec']
    elif unit == '#':
        target_dtids = ['positiveInteger_10', 'positiveorNegativeInteger_10',
                        'positiveInteger_5', 'positiveInteger_6',
                        'positiveOrNullInteger']
    elif unit_lower and re.match(r'^#[.#]+$', unit_lower):
        # Decimal format like #.### -> decimal number
        target_dtids = ['positiveDecimalNumber_11_2Dec', 'positiveDecimalNumber_2_4Dec',
                        'rate_2_4Dec', 'positiveRate_1_4Dec', 'percentage']
    elif unit == 'text':
        target_dtids = ['alpha_255', 'alpha_1_120', 'identifier_120']
    elif unit_lower in ('date', 'yyyy-mm-dd', 'dd/mm/yyyy', 'mm/dd/yyyy'):
        target_dtids = ['date', 'date_YYYYMMDD', 'metricDate']
    elif unit == 'bool':
        target_dtids = ['metricBool', 'boolean_0_1']
    elif unit_lower.startswith('[') and unit_lower.endswith(']'):
        # Bracket unit like [Accounting standard] -> search dtid_to_def
        bracket_text = unit_lower[1:-1].strip().replace(' ', '').lower()
        for dtid in dtid_to_def:
            if dtid.lower().replace('_', '') == bracket_text:
                return dtid_to_def[dtid], 2, f"Bracket unit '{unit}' -> {dtid}"
        # Also try matching against definition names
        for name in def_index:
            if name.lower().replace('_', '') == bracket_text:
                return name, 2, f"Bracket unit '{unit}' -> {name}"
        # Prefix/substring match (e.g. "accountstd" vs "accountingstandard")
        for dtid in dtid_to_def:
            dtid_clean = dtid.lower().replace('_', '')
            # Check shared prefix of at least 7 chars
            min_len = min(len(dtid_clean), len(bracket_text))
            if min_len >= 7 and dtid_clean[:7] == bracket_text[:7]:
                return dtid_to_def[dtid], 1, f"Bracket unit '{unit}' ~-> {dtid}"
        for name in def_index:
            name_clean = name.lower().replace('_', '')
            min_len = min(len(name_clean), len(bracket_text))
            if min_len >= 7 and name_clean[:7] == bracket_text[:7]:
                return name, 1, f"Bracket unit '{unit}' ~-> {name}"

    for dtid in target_dtids:
        if dtid in dtid_to_def:
            return dtid_to_def[dtid], 2, f"Unit '{unit}' sign '{sign}' -> {dtid}"

    # 3. Fallback: infer from label when unit is absent
    if not unit or (unit_lower.startswith('[') and unit_lower.endswith(']')):
        label = col_info.get('label', '') or ''
        label_lower = label.lower()
        mp_lower = mp_label.lower()
        combined = label_lower + ' ' + mp_lower

        if 'date' in combined or 'maturity' in combined or 'expiry' in combined:
            for dtid in ['date', 'date_YYYYMMDD', 'metricDate']:
                if dtid in dtid_to_def:
                    return dtid_to_def[dtid], 1, f"Label -> {dtid}"
        if '%' in label or 'percentage' in combined or 'ratio' in combined:
            for dtid in ['positiveRate_1_4Dec', 'rate_2_4Dec', 'percentage']:
                if dtid in dtid_to_def:
                    return dtid_to_def[dtid], 1, f"Label -> {dtid}"
        if 'comment' in combined or 'description' in combined or 'name' in combined:
            for dtid in ['alpha_255', 'alpha_1_120', 'identifier_120']:
                if dtid in dtid_to_def:
                    return dtid_to_def[dtid], 1, f"Label -> {dtid}"
        if any(kw in combined for kw in [
            'amount', 'carrying', 'fair value', 'allowance', 'impairment',
            'exposure', 'servicing', 'forborne', 'forbearance',
            'performing', 'non-performing', 'defaulted', 'past due',
        ]):
            for dtid in ['positiveOrNegativeAmount', 'positiveOrNegativeAmount_2Dec',
                         'positiveorNullAmount_18_2Dec', 'negativeorNullAmount_18_2Dec']:
                if dtid in dtid_to_def:
                    return dtid_to_def[dtid], 1, f"Label (amount) -> {dtid}"

    return None, 0, "No mapping found"


def scan_xlsm_datatypes(base_dir):
    """
    Scans all pure-form XLSM files and returns datatypes per column.
    Returns: [{file, file_path, sheet, d2_datatype, ident_datatype, col_datatypes, col_labels, col_codes}]
    """
    results = []

    for fpath in find_xlsm_files(base_dir):
        fname = os.path.basename(fpath)
        try:
            wb = openpyxl.load_workbook(fpath, keep_vba=True)
        except Exception as e:
            print(f"  ERROR loading {fname}: {e}", file=sys.stderr)
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
    For each form sheet, finds the matching DPM sheet.
    Returns: [(form_item, dpm_sheet_name, dpm_cols)]
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
    """Displays a report of datatypes to set."""
    print("\n" + "=" * 90)
    print("DATATYPES REPORT BY COLUMN")
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
            print(f"\n[!] {form_item['file']} : identification J2 (datatype) = (EMPTY)")
            print(f"  -> Suggested : {domain}.positiveOrNegativeAmount")

    for form_item, dpm_sheet, dpm_cols in matches:
        ident_dt = form_item.get('ident_datatype')
        print(f"\n+- {form_item['file']} [{form_item['sheet']}]")
        print(f"|  identification datatype: {ident_dt or '(EMPTY) [!]'}")
        print(f"|  D2 (sheet datatype)    : {form_item['d2_datatype'] or '(empty)'}")

        if not dpm_cols:
            print(f"|  [!] No matching DPM found")
            for code in form_item['col_codes']:
                dt = form_item['col_datatypes'].get(code, '')
                lbl = form_item['col_labels'].get(code, '')
                if dt:
                    print(f"|  {code}: {dt:40s} [{lbl}]")
                    total_filled += 1
                else:
                    print(f"|  {code}: {'(empty)':40s} [{lbl}]")
                    total_missing += 1
            print(f"+{'-' * 80}")
            continue

        print(f"|  DPM sheet : {dpm_sheet}")
        print(f"|")

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
                suggested, confidence, source = None, 0, "Column not in DPM"

            conf_str = ['?', '~', '[OK]', '[*]'][min(confidence, 3)]
            full_suggested = f"{domain}.{suggested}" if suggested else None

            if existing_dt:
                if suggested and existing_dt != full_suggested:
                    status = f"[~] existing={existing_dt}, suggested={full_suggested}"
                    total_filled += 1
                else:
                    status = f"[OK] {existing_dt}"
                    total_filled += 1
            elif suggested:
                status = f"-> {full_suggested}  [{conf_str} {source}]"
                total_to_set += 1
            else:
                status = f"  (no mapping)"
                total_missing += 1

            print(f"|  {code}: {status:60s} [{form_label or dpm_label}]")

        print(f"+{'-' * 80}")

    ident_missing = set()
    for form_item, _, _ in matches:
        if not form_item.get('ident_datatype'):
            ident_missing.add(form_item['file'])

    print(f"\n{'=' * 90}")
    print("DATATYPES SUMMARY")
    print(f"{'=' * 90}")
    if ident_missing:
        print(f"  Forms without identification datatype : {len(ident_missing)}")
        for f in sorted(ident_missing):
            print(f"    -> {f}")
    print(f"  Columns with datatype set : {total_filled}")
    print(f"  Columns to set (suggested)  : {total_to_set}")
    print(f"  Columns without mapping            : {total_missing}")


def write_datatypes(matches, def_index, domain, base_dir):
    """
    Writes datatypes to XLSM files:
    - identification J2: default datatype if missing
    - row 7: datatypes per column if missing
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
            print(f"  ERROR opening {fname}: {e}", file=sys.stderr)
            continue

        file_modified = False

        ident_dt = items[0][0].get('ident_datatype')
        if not ident_dt and 'identification' in wb.sheetnames:
            ws_id = wb['identification']
            default_dt = f"{domain}.positiveOrNegativeAmount"
            old_val = ws_id.cell(row=2, column=10).value
            ws_id.cell(row=2, column=10, value=default_dt)
            print(f"  [OK] {fname}[identification] J2: '{old_val}' -> '{default_dt}'")
            file_modified = True

        for form_item, dpm_sheet, dpm_cols in items:
            sn = form_item['sheet']
            if sn not in wb.sheetnames:
                print(f"  [!] Sheet {sn} not found in {fname}")
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

                dpm_col = dpm_cols.get(code) if dpm_cols else None

                if dpm_col:
                    # DPM column has data but is empty (no unit, no label, no property)
                    has_data = dpm_col.get('unit') or dpm_col.get('label') or dpm_col.get('main_property_label')
                    if not has_data:
                        continue
                    suggested, confidence, source = determine_datatype(
                        dpm_col, def_index, domain)
                    if suggested:
                        full_name = f"{domain}.{suggested}"
                    else:
                        full_name = "NOT_RESOLVED"
                else:
                    full_name = "NOT_RESOLVED"

                ws.cell(row=7, column=col, value=full_name)
                cols_written += 1

            if cols_written > 0:
                print(f"  [OK] {fname}[{sn}] : {cols_written} datatypes written")
                file_modified = True
            else:
                print(f"  - {fname}[{sn}] : nothing to write")

        if file_modified:
            try:
                wb.save(fpath)
                modified_count += 1
                print(f"  [OK] File saved: {fname}")
            except PermissionError:
                print(f"  [X] '{fname}' locked (open in Excel?)")

        wb.close()

    print(f"\n  Total files modified: {modified_count}")


def set_datatype_for_column(sheet_name, col_code, dict_name, def_index, domain, base_dir):
    """Sets the datatype for a specific column of a sheet."""
    if dict_name not in def_index:
        print(f"  ERROR: '{dict_name}' not found in dictionary!")
        candidates = [n for n in sorted(def_index.keys()) if dict_name.lower() in n.lower()]
        if candidates:
            print(f"  Close matches:")
            for c in candidates[:10]:
                info = def_index[c]
                print(f"    - {c} (datatypeId={info.get('datatypeId', '?')}, formatType={info.get('formatType', '?')})")
        return False

    full_name = f"{domain}.{dict_name}"

    target_file = find_sheet_in_xlsm(sheet_name, base_dir)
    if not target_file:
        print(f"  ERROR: No XLSM file containing sheet '{sheet_name}'")
        return False

    fname = os.path.basename(target_file)
    create_backup(target_file)

    wb = openpyxl.load_workbook(target_file, keep_vba=True)
    ws = wb[sheet_name]

    header_row = find_header_row(ws)
    if not header_row:
        print(f"  ERROR: Cannot find header row in {sheet_name}")
        wb.close()
        return False

    target_col = None
    for col in range(14, 60):
        v = ws.cell(row=header_row, column=col).value
        if v and str(v).strip() == col_code:
            target_col = col
            break

    if target_col is None:
        print(f"  ERROR: Column code '{col_code}' not found in {sheet_name}")
        wb.close()
        return False

    old_val = ws.cell(row=7, column=target_col).value
    ws.cell(row=7, column=target_col, value=full_name)
    print(f"  [OK] {fname}[{sheet_name}] col {col_code} ({get_column_letter(target_col)}7): '{old_val}' -> '{full_name}'")

    wb.save(target_file)
    print(f"  [OK] File saved: {fname}")
    wb.close()
    return True


######################################################################################################################################################
#                                                                                                                                                    #
#       Part 4: Enrichment GUI — unified tkinter interface for Z-axis + Datatypes                                                           #
#                                                                                                                                                    #
#   run_enrichment_gui: interactive GUI with DPM/dictionary selection and per-sheet actions                                                      #
#   Displays Z-axis and Datatypes results in a table with statuses                                                                           #
#   Allows selective writing (per sheet or global) of resolved values                                                                          #
#                                                                                                                                                    #
######################################################################################################################################################

def run_enrichment_gui(base_dir):
    """
    Unified interactive GUI mode:
      1. DPM file(s) selection
      2. Dictionary selection
      3. Z-axis tab + Datatypes tab
      4. Combined writing
    """
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk

    root = tk.Tk()
    root.withdraw()

    # -- Step 1: DPM file(s) selection ---------------------------
    messagebox.showinfo(
        "Enrichment Tool",
        "Step 1/3: Select DPM file(s)\n"
        "(Annotated Table Layout .xlsx)\n\n"
        "You can select multiple files (Ctrl+click)."
    )

    layout_paths = filedialog.askopenfilenames(
        title="Select DPM file(s) (Annotated Table Layout)",
        initialdir=base_dir,
        filetypes=[("Excel files", "*.xlsx"), ("All", "*.*")]
    )

    if not layout_paths:
        print("Cancelled: No DPM file selected.")
        root.destroy()
        return

    for lp in layout_paths:
        print(f"  DPM selected : {os.path.basename(lp)}")

    # -- Step 2: YML dictionary selection -----------------------------
    messagebox.showinfo(
        "Enrichment Tool",
        "Step 2/3: Select dictionary file\n"
        "(pure-dictionary .yml)"
    )

    dict_path = filedialog.askopenfilename(
        title="Select dictionary file (.yml)",
        initialdir=base_dir,
        filetypes=[("YAML files", "*.yml *.yaml"), ("All", "*.*")]
    )

    if not dict_path:
        print("Cancelled: No dictionary selected.")
        root.destroy()
        return

    domain = get_dict_domain(dict_path)
    print(f"  Dictionary selected : {os.path.basename(dict_path)} (domaine: {domain})")

    # -- Data loading -----------------------------------------------
    print("Loading dictionary...")
    type_index, def_index = load_dictionary(dict_path)
    print(f"  {len(type_index)} Z-axis types, {len(def_index)} datatype definitions")

    # Z-axis : lecture DPM
    print("Reading Z-axis from DPM files...")
    ann_zaxis = {}
    for lp in layout_paths:
        partial = read_annotated_layout(lp)
        for sheet_name, info in partial.items():
            if sheet_name not in ann_zaxis:
                ann_zaxis[sheet_name] = info
    print(f"  {len(ann_zaxis)} sheets with Z-axis found in DPM files")

    # Datatypes : lecture DPM
    print("Reading DPM columns for datatypes...")
    dpm_columns = {}
    for lp in layout_paths:
        partial = read_dpm_columns(lp)
        for sheet_name, cols in partial.items():
            if sheet_name not in dpm_columns:
                dpm_columns[sheet_name] = cols
    print(f"  {len(dpm_columns)} sheets with columns found")

    # Scan XLSM (once only!)
    print("Scanning XLSM forms...")
    xlsm_files = find_xlsm_files(base_dir)

    # Z-axis scan
    zaxis_items = scan_xlsm_zaxis(base_dir)
    print(f"  {len(zaxis_items)} sheets with zAxis header")

    # Datatypes scan
    form_items = scan_xlsm_datatypes(base_dir)
    print(f"  {len(form_items)} data sheets")

    dt_matches = match_dpm_to_form(dpm_columns, form_items)

    # -- Sheet -> XLSM file index ------------------------------------------
    sheet_to_file = {}
    for fpath in xlsm_files:
        try:
            wb = openpyxl.load_workbook(fpath, keep_vba=True)
            for sn in wb.sheetnames:
                sheet_to_file[sn] = fpath
            wb.close()
        except Exception:
            continue

    # -- Prepare Z-axis data -----------------------------------------
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

        if not type_name:
            status = "[!] NOT_RESOLVED (no match found)"
            needs_write = True
        elif type_name and not current_j2:
            status = "to set"
            needs_write = True
        elif current_j2.lower() != type_name.lower():
            status = f"[!] to fix ({current_j2} -> {type_name})"
            needs_write = True
        else:
            status = "[OK] OK"
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

    # XLSM sheets with zAxis NOT in DPM
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
                    status = "to set"
                    needs_w = True
                elif tn and j2.lower() != tn.lower():
                    status = f"[!] to fix ({j2} -> {tn})"
                    needs_w = True
                else:
                    status = "[OK] OK"
                    needs_w = False
                zaxis_selectable.append({
                    'sheet': sn, 'file': fpath, 'type_name': tn,
                    'source': 'XLSM', 'current_j2': j2, 'status': status,
                    'needs_write': needs_w,
                })
            wb.close()
        except Exception:
            continue

    # -- Prepare datatypes data --------------------------------------
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
            if suggested:
                full_suggested = f"{domain}.{suggested}"
            elif not existing:
                full_suggested = "NOT_RESOLVED"
            else:
                full_suggested = ""
            dt_table_data.append({
                'file': form_item['file'], 'file_path': form_item['file_path'],
                'sheet': form_item['sheet'], 'col_code': code, 'label': label,
                'existing': existing, 'suggested': full_suggested,
                'suggested_name': suggested, 'confidence': confidence,
                'source': source, 'unit': unit, 'form_item': form_item,
            })

    # -- Step 3: Tabbed window -----------------------------------------
    sel_window = tk.Toplevel(root)
    sel_window.title("Enrichment Tool — Z-axis & Datatypes")
    sel_window.geometry("1200x700")

    title_frame = tk.Frame(sel_window, padx=10, pady=5)
    title_frame.pack(fill=tk.X)
    tk.Label(title_frame,
             text="Step 3/3: Select items to write",
             font=("Segoe UI", 12, "bold")).pack(anchor=tk.W)
    dpm_names = ', '.join(os.path.basename(lp) for lp in layout_paths)
    tk.Label(title_frame, text=f"DPM : {dpm_names}",
             font=("Segoe UI", 9), fg="gray").pack(anchor=tk.W)

    notebook = ttk.Notebook(sel_window)
    notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

    # -- Z-axis tab -------------------------------------------------------
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

    z_headers = ["", "Sheet", "Type", "Source", "Status", "File"]
    for col_idx, h in enumerate(z_headers):
        tk.Label(z_inner, text=h, font=("Segoe UI", 9, "bold"),
                 borderwidth=1, relief="solid", padx=5, pady=2
                 ).grid(row=0, column=col_idx, sticky="ew")

    z_check_vars = []
    for i, item in enumerate(zaxis_selectable):
        can_select = item['file'] is not None and (item['type_name'] is not None or 'NOT_RESOLVED' in item.get('status', ''))
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

    # -- Datatypes tab -----------------------------------------------------
    dt_frame = tk.Frame(notebook)
    notebook.add(dt_frame, text=f" Datatypes ({len(dt_table_data)} columns) ")

    # Filter
    dt_filter_frame = tk.Frame(dt_frame, padx=10, pady=3)
    dt_filter_frame.pack(fill=tk.X)
    tk.Label(dt_filter_frame, text="Filter:").pack(side=tk.LEFT)
    dt_filter_var = tk.StringVar(value="to set")
    for txt, val in [("All", "all"), ("To set", "to set"), ("Already set", "already set")]:
        tk.Radiobutton(dt_filter_frame, text=txt, variable=dt_filter_var, value=val).pack(side=tk.LEFT, padx=5)

    # Treeview
    dt_tree_frame = tk.Frame(dt_frame, padx=10, pady=5)
    dt_tree_frame.pack(fill=tk.BOTH, expand=True)

    dt_columns = ("select", "file", "sheet", "colcode", "label", "unit", "existing", "suggested", "conf", "source")
    dt_tree = ttk.Treeview(dt_tree_frame, columns=dt_columns, show='headings', selectmode='extended')

    dt_tree.heading("select", text="[OK]")
    dt_tree.heading("file", text="File")
    dt_tree.heading("sheet", text="Sheet")
    dt_tree.heading("colcode", text="Col")
    dt_tree.heading("label", text="Label")
    dt_tree.heading("unit", text="Unit")
    dt_tree.heading("existing", text="Existing")
    dt_tree.heading("suggested", text="Suggested")
    dt_tree.heading("conf", text="Conf.")
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
            if filter_mode == "to set" and row['existing']:
                continue
            if filter_mode == "already set" and not row['existing']:
                continue
            conf_str = ['?', '~', '[OK]', '[*]'][min(row['confidence'], 3)]
            sel_str = "[x]" if i in dt_selected_set else "[ ]"
            is_unresolved = row['suggested'] == 'NOT_RESOLVED'
            if is_unresolved:
                tag = 'unresolved'
            elif i in dt_selected_set:
                tag = 'selected'
            else:
                tag = ''
            dt_tree.insert("", tk.END, iid=str(i), values=(
                sel_str, row['file'], row['sheet'], row['col_code'],
                row['label'], row['unit'], row['existing'] or '',
                row['suggested'] or '(no mapping)', conf_str, row['source'],
            ), tags=(tag,))
        dt_tree.tag_configure('selected', background='#d4edda')
        dt_tree.tag_configure('unresolved', background='#f8d7da')

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

    # Auto-selection (exclude NOT_RESOLVED and empty suggestions)
    for i, row in enumerate(dt_table_data):
        if not row['existing'] and row['suggested'] and row['suggested'] != 'NOT_RESOLVED':
            dt_selected_set.add(i)

    dt_populate_tree("to set")
    dt_tree.bind('<Double-1>', dt_toggle)
    dt_filter_var.trace_add('write', lambda *_: dt_populate_tree(dt_filter_var.get()))

    # -- Common buttons ----------------------------------------------------
    btn_frame = tk.Frame(sel_window, padx=10, pady=10)
    btn_frame.pack(fill=tk.X)

    count_label = tk.Label(btn_frame, text="", font=("Segoe UI", 9))
    count_label.pack(side=tk.LEFT, padx=10)

    def update_count():
        nz = sum(1 for v in z_check_vars if v.get())
        nd = len(dt_selected_set)
        count_label.config(text=f"Z-axis: {nz} sheet(s) | Datatypes: {nd} column(s)")
        sel_window.after(500, update_count)

    update_count()

    def z_select_all():
        for i, var in enumerate(z_check_vars):
            item = zaxis_selectable[i]
            if item['file'] and (item['type_name'] or 'NOT_RESOLVED' in item.get('status', '')):
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

    tk.Button(btn_frame, text="Z: Select All",
              command=z_select_all).pack(side=tk.LEFT, padx=3)
    tk.Button(btn_frame, text="Z: Deselect All",
              command=z_deselect_all).pack(side=tk.LEFT, padx=3)
    tk.Button(btn_frame, text="DT: Select All",
              command=dt_select_all_visible).pack(side=tk.LEFT, padx=3)
    tk.Button(btn_frame, text="DT: Deselect All",
              command=dt_deselect_all).pack(side=tk.LEFT, padx=3)

    tk.Button(btn_frame, text="Cancel",
              command=on_cancel).pack(side=tk.RIGHT, padx=5)
    tk.Button(btn_frame, text="Write All", command=on_write,
              bg="#4CAF50", fg="white",
              font=("Segoe UI", 10, "bold")).pack(side=tk.RIGHT, padx=5)

    sel_window.wait_window()

    if result_action[0] != 'write':
        print("Cancelled.")
        root.destroy()
        return

    # -- Write Z-axis ------------------------------------------------------
    z_selected = [zaxis_selectable[i] for i, v in enumerate(z_check_vars) if v.get()]
    z_written = 0
    if z_selected:
        print(f"\n{'=' * 80}")
        print(f"WRITING Z-AXIS ({len(z_selected)} sheets)")
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
                print(f"  [X] '{fname}' locked (open in Excel?)")
                continue
            except Exception as e:
                print(f"  [X] Error opening {fname}: {e}")
                continue

            for item in items:
                sn = item['sheet']
                tn = item['type_name']
                if sn not in wb.sheetnames:
                    print(f"  [!] Sheet '{sn}' not found in {fname}")
                    continue

                ws = wb[sn]

                # Z-axis present but no match -> write NOT_RESOLVED
                # (only if status indicates NOT_RESOLVED, i.e. DPM says it should exist)
                if not tn:
                    if 'NOT_RESOLVED' in item.get('status', ''):
                        old_j2 = ws.cell(row=2, column=10).value
                        ws.cell(row=2, column=10, value='NOT_RESOLVED')
                        ws.cell(row=4, column=10, value='NOT_RESOLVED')
                        print(f"  [!] {fname}[{sn}] : Z-axis present, no match found -> NOT_RESOLVED")
                        z_written += 1
                    continue

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

                log = f"  [OK] {fname}[{sn}] : J2='{old_j2}'->'{j2_val}' | J4='{old_j4}'->'{attribute}'"
                if j3_val:
                    log += f" | J3='{old_j3}'->'{j3_val}'"
                print(log)
                z_written += 1

            try:
                wb.save(fpath)
                print(f"  [OK] File saved: {fname}")
            except PermissionError:
                print(f"  [X] '{fname}' locked, unable to save")
            wb.close()

    # -- Write Datatypes ---------------------------------------------------
    dt_written = 0
    if dt_selected_set:
        print(f"\n{'=' * 80}")
        print(f"WRITING DATATYPES ({len(dt_selected_set)} columns)")
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
                print(f"  [X] '{fname}' locked (open in Excel?)")
                continue
            except Exception as e:
                print(f"  [X] Error opening {fname}: {e}")
                continue

            for row in rows:
                sn = row['sheet']
                if sn not in wb.sheetnames:
                    continue
                if not row['suggested']:
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
                        print(f"  [OK] {fname}[{sn}] col {row['col_code']} ({get_column_letter(col)}7): "
                              f"'{old_val}' -> '{row['suggested']}'")
                        dt_written += 1
                        break

            try:
                wb.save(fpath)
                print(f"  [OK] File saved: {fname}")
            except PermissionError:
                print(f"  [X] '{fname}' locked, unable to save")
            wb.close()

    # -- Summary ---------------------------------------------------------------
    summary = f"Done!\nZ-axis: {z_written} sheets\nDatatypes: {dt_written} columns"
    print(f"\n{summary}")
    messagebox.showinfo("Enrichment Tool", summary)
    root.destroy()


######################################################################################################################################################
#                                                                                                                                                    #
#       Part 5: RowKey — DPM Open Rows reading, XLSM scan, building, matching, writing, GUI                                                 #
#                                                                                                                                                    #
#   get_domain_from_identification: domain extraction from identification sheet                                                            #
#   read_dpm_open_rows: reads DPM sheets with Open Rows and <Key value> (multi-column)                                                       #
#   scan_xlsm_rowkeys: scans XLSM files for sheets with rowNumbering=999                                                                          #
#   build_rowkey / match_and_prepare: builds rowKey formulas and DPM -> XLSM matching                                                       #
#   write_rowkey / print_report: writes rowKey and CLI report                                                                                 #
#   run_rowkey_gui: interactive GUI for analyzing and writing rowKey                                                                              #
#                                                                                                                                                    #
######################################################################################################################################################

def get_domain_from_identification(wb):
    """Extracts the domain (linkedDomain) from the identification sheet."""
    if 'identification' not in wb.sheetnames:
        return None
    ws = wb['identification']
    # Row 1 = headers, Row 2 = values
    # Find linkedDomain column
    for c in range(1, 20):
        if ws.cell(1, c).value == 'linkedDomain':
            return ws.cell(2, c).value
    return None


# -- DPM reading: Open Rows and <Key value> -------------------------------------

def read_dpm_open_rows(layout_path):
    """
    Reads a DPM file and returns sheets with Open Rows and <Key value>.

    Returns: {sheet_name: {
        'key_col_codes': [str],     # ex: ['0010'] ou ['0020','0030','0090','0100']
        'is_sheet_per': bool,       # True si B3 = "Sheet per..."
        'key_property': str,        # ex: 'qID', 'qAE', 'EN'
        'key_hierarchy': str,       # ex: 'qRAC', 'qAE', 'EN'
        'key_text': str,            # Full text of first <Key value> cell
        'key_col_code': str,        # first key column (compatibility)
    }}
    """
    if not os.path.exists(layout_path):
        print(f"  ERROR: DPM file not found: {layout_path}", file=sys.stderr)
        return {}

    wb = openpyxl.load_workbook(layout_path, read_only=True, data_only=True)
    result = {}

    for sn in wb.sheetnames:
        if sn == 'TOC':
            continue
        ws = wb[sn]

        # Check for "Open Rows" in column B
        has_open_rows = False
        for r in range(1, 25):
            b = ws.cell(r, 2).value
            if b and "Open Rows" in str(b):
                has_open_rows = True
                break

        if not has_open_rows:
            continue

        # Check if "Sheet per" pattern
        b3 = ws.cell(3, 2).value
        is_sheet_per = b3 is not None and "Sheet per" in str(b3)

        # Find ALL <Key value> occurrences
        key_cells = []  # [(row, col, text)]
        for r in range(1, 30):
            for c in range(2, 25):
                v = ws.cell(r, c).value
                if v and "<Key value>" in str(v):
                    key_cells.append((r, c, str(v)))

        if not key_cells:
            # Open Rows without <Key value> -> no rowKey
            result[sn] = {
                'key_col_codes': [],
                'key_col_code': None,
                'is_sheet_per': is_sheet_per,
                'key_property': None,
                'key_hierarchy': None,
                'key_text': None,
            }
            continue

        # Find the code row (row with column codes like 0005, 0010, etc.)
        code_row = None
        for r in range(5, 20):
            for c in range(4, 25):
                v = ws.cell(r, c).value
                if v and re.match(r'^0\d{3}$', str(v).strip()):
                    code_row = r
                    break
            if code_row:
                break

        # When "Sheet per" is active, <Key value> cells above the code row
        # are z-axis dimensions (not row keys). Keep only those below.
        if is_sheet_per and code_row:
            row_key_cells = [(r, c, t) for r, c, t in key_cells if r > code_row]
            if row_key_cells:
                key_cells = row_key_cells

        # Resolve column code for each <Key value>, deduplicate by column
        key_col_codes = []
        seen_cols = set()
        if code_row:
            for _, kc, _ in key_cells:
                if kc in seen_cols:
                    continue
                seen_cols.add(kc)
                v = ws.cell(code_row, kc).value
                if v:
                    key_col_codes.append(str(v).strip())

        # Extract property/hierarchy from first key text
        key_text = key_cells[0][2]
        key_property = None
        key_hierarchy = None
        if key_text:
            # Try (hierarchy:variant(hierarchy)) pattern
            m = re.match(r'\((\w+):(\w*)\(', key_text)
            if m:
                key_hierarchy = m.group(1)
                key_property = m.group(2) if m.group(2) else m.group(1)
            else:
                # Try (property:) pattern
                m = re.match(r'\((\w+):\)', key_text)
                if m:
                    key_hierarchy = m.group(1)
                    key_property = m.group(1)

        result[sn] = {
            'key_col_codes': key_col_codes,
            'key_col_code': key_col_codes[0] if key_col_codes else None,
            'is_sheet_per': is_sheet_per,
            'key_property': key_property,
            'key_hierarchy': key_hierarchy,
            'key_text': key_text,
        }

    wb.close()
    return result


# -- Scan XLSM: sheets with Open Rows (rowNumbering=999) ----------------------

def scan_xlsm_rowkeys(base_dir):
    """
    Scans all pure-form XLSM files and returns sheets with rowNumbering=999.

    Returns: [{
        'file': str,          # file name
        'file_path': str,     # full path
        'sheet': str,         # sheet name
        'header_row': int,    # header row (level/rowKey/...)
        'rowkey_col': int,    # rowKey column (usually B=2)
        'rownumbering_col': int,
        'open_rows': [(row, rowkey_value)],   # rows with 999
        'h2_value': str,      # allocationRuleCondition (H2)
        'domain': str,        # linkedDomain du fichier
    }]
    """
    result = []
    for fpath in find_xlsm_files(base_dir):
        try:
            wb = openpyxl.load_workbook(fpath, keep_vba=True)
        except Exception:
            continue

        domain = get_domain_from_identification(wb)
        fname = os.path.basename(fpath)

        for sn in wb.sheetnames:
            if sn in SKIP_SHEETS:
                continue
            ws = wb[sn]

            header_row = find_header_row(ws)
            if not header_row:
                continue

            # Find rowKey and rowNumbering columns
            rowkey_col = None
            rownumbering_col = None
            for c in range(1, 15):
                hv = ws.cell(header_row, c).value
                if hv:
                    hv_s = str(hv).strip()
                    if hv_s == 'rowKey':
                        rowkey_col = c
                    elif hv_s == 'rowNumbering':
                        rownumbering_col = c

            if not rownumbering_col:
                continue

            # Find rows with rowNumbering=999
            open_rows = []
            for r in range(header_row + 1, ws.max_row + 1):
                v = ws.cell(r, rownumbering_col).value
                if v is not None and str(v).strip() == '999':
                    rk_val = ws.cell(r, rowkey_col).value if rowkey_col else None
                    open_rows.append((r, rk_val))

            if not open_rows:
                continue

            # H2 = allocationRuleCondition
            h2 = ws.cell(2, 8).value

            result.append({
                'file': fname,
                'file_path': fpath,
                'sheet': sn,
                'header_row': header_row,
                'rowkey_col': rowkey_col or 2,  # default to B
                'rownumbering_col': rownumbering_col,
                'open_rows': open_rows,
                'h2_value': h2,
                'domain': domain,
            })

        wb.close()

    return result


# -- Building and matching: DPM -> XLSM -> rowKey ----------------------------

def extract_dataset_field_from_h2(h2_value):
    """
    Extracts the dataset.field from an allocationRuleCondition H2.
    Exemple: "euebacorep_prudentvaluation.euebacorepRankava <> ('NULL')"
             -> "euebacorep_prudentvaluation.euebacorepRankava"
    """
    if not h2_value:
        return None
    # Pattern: n"ndd  x .domainField <> ('NULL')
    m = re.search(r'(\w+\.\w+)\s*<>\s*\(\s*[\'"]NULL[\'"]\s*\)', str(h2_value))
    if m:
        return m.group(1)
    return None


def build_rowkey(sheet_name, key_col_code, dataset_field=None, key_col_codes=None):
    """
    Construit la formule rowKey.
      - Simple  : ${sheet|r999|cXXXX}
      - Multi   : ${sheet|r999|cXXXX}|${sheet|r999|cYYYY}
      - Avec ref: ${sheet|r999|cXXXX};dataset.field
    """
    codes = key_col_codes if key_col_codes and len(key_col_codes) > 0 else [key_col_code]
    parts = [f"${{{sheet_name}|r999|c{c}}}" for c in codes]
    base = "|".join(parts)
    if dataset_field:
        return f"{base};{dataset_field}"
    return base


def match_and_prepare(xlsm_items, dpm_data):
    """
    Cross-references XLSM data (sheets with 999) and DPM data (open rows / key values).
    Returns a list of actions to perform.

    Chaque action: {
        'file_path': str,
        'file': str,
        'sheet': str,
        'rowkey_col': int,
        'rows': [(row_num, existing_value)],
        'expected': str,         # formule rowKey attendue
        'status': str,           # 'to set', '[OK] OK', '[!] to fix'
        'needs_write': bool,
        'h2_value': str,
        'h2_expected': str,      # allocationRuleCondition attendue (ou None)
        'h2_needs_write': bool,
        'key_col_code': str,
        'is_sheet_per': bool,
        'dataset_field': str,
        'dpm_source': str,       # info de diagnostic
    }
    """
    actions = []

    for item in xlsm_items:
        sheet_name = item['sheet']
        h2 = item['h2_value']

        # Match with DPM
        dpm_info = dpm_data.get(sheet_name)

        key_col_code = None
        key_col_codes = []
        is_sheet_per = False
        dataset_field = None
        dpm_source = ""
        h2_expected = None
        h2_needs_write = False

        if dpm_info:
            key_col_code = dpm_info['key_col_code']
            key_col_codes = dpm_info.get('key_col_codes', [])
            is_sheet_per = dpm_info['is_sheet_per']
            dpm_source = f"DPM: {dpm_info['key_text'] or 'Open Rows (no key)'}".replace('\n', ' ')[:80]
            if len(key_col_codes) > 1:
                dpm_source += f" (+{len(key_col_codes)-1} keys)"

            if not is_sheet_per and key_col_code:
                # Inline key -> need dataset.field from H2 or flag
                dataset_field = extract_dataset_field_from_h2(h2)
                if dataset_field:
                    h2_expected = f"{dataset_field} <> ('NULL')"
                else:
                    # H2 empty -> flag as incomplete inline key
                    dataset_field = "NOT_RESOLVED"
                    print(f"  [!] {sheet_name}: inline key (pas 'Sheet per') mais H2 vide -> dataset.field=NOT_RESOLVED")
        else:
            # Not found in DPM, try to detect from XLSM existing values
            for _, rk_val in item['open_rows']:
                if rk_val:
                    # Extract ALL key codes from existing rowKey formula
                    found = re.findall(r'\$\{[^}]+\|r999\|c(\d+)\}', str(rk_val).strip())
                    if found:
                        key_col_codes = found
                        key_col_code = found[0]
                        break
            dpm_source = "XLSM (no DPM)"

        if not key_col_code:
            # Cannot determine key column
            expected = None
            status = "[!] unknown key column"
            needs_write = False
        else:
            expected = build_rowkey(sheet_name, key_col_code, dataset_field,
                                    key_col_codes=key_col_codes)

            # Compare with existing
            existing_values = [rk for _, rk in item['open_rows'] if rk]
            if existing_values:
                existing_clean = str(existing_values[0]).strip()
                if existing_clean == expected:
                    status = "[OK] OK"
                    needs_write = False
                else:
                    status = f"[!] to fix ({existing_clean} -> {expected})"
                    needs_write = True
            else:
                status = "to set"
                needs_write = True

            # Flag incomplete inline keys
            if not is_sheet_per and dataset_field == "NOT_RESOLVED":
                status = f"[!] inline key, H2 vide (dataset.field manquant)"
                needs_write = True

            # Check if H2 needs writing for inline key
            if not is_sheet_per and dataset_field and h2_expected:
                if h2 and h2.strip() == h2_expected.strip():
                    h2_needs_write = False
                elif not h2:
                    h2_needs_write = False

        actions.append({
            'file_path': item['file_path'],
            'file': item['file'],
            'sheet': sheet_name,
            'rowkey_col': item['rowkey_col'],
            'rows': item['open_rows'],
            'expected': expected,
            'status': status,
            'needs_write': needs_write,
            'h2_value': h2,
            'h2_expected': h2_expected,
            'h2_needs_write': h2_needs_write,
            'key_col_code': key_col_code,
            'key_col_codes': key_col_codes,
            'is_sheet_per': is_sheet_per,
            'dataset_field': dataset_field,
            'dpm_source': dpm_source,
            'header_row': item['header_row'],
        })

    return actions


# -- Writing -------------------------------------------------------------------

def write_rowkey(action):
    """Writes the rowKey to the XLSM file for a given action."""
    fpath = action['file_path']
    sheet_name = action['sheet']
    expected = action['expected']

    if not expected:
        print(f"  [!] {sheet_name}: no rowKey to write (key column unknown)")
        return False

    create_backup(fpath)

    wb = openpyxl.load_workbook(fpath, keep_vba=True)
    ws = wb[sheet_name]
    changed = False

    for row_num, existing in action['rows']:
        existing_clean = str(existing).strip() if existing else ''
        if existing_clean != expected:
            ws.cell(row_num, action['rowkey_col']).value = expected
            print(f"  [OK] {sheet_name} B{row_num}: {expected}")
            changed = True

    # Also write H2 if needed
    if action['h2_needs_write'] and action['h2_expected']:
        ws.cell(2, 8).value = action['h2_expected']
        print(f"  [OK] {sheet_name} H2: {action['h2_expected']}")
        changed = True

    if changed:
        wb.save(fpath)
        print(f"  -> Saved: {os.path.basename(fpath)}")
    else:
        print(f"  — {sheet_name}: nothing to change")

    wb.close()
    return changed


# -- Report (CLI mode) ---------------------------------------------------------

def print_report(actions):
    """Displays a report of detected and expected rowKey."""
    print("\n" + "=" * 90)
    print("  ROWKEY REPORT")
    print("=" * 90)

    to_fill = [a for a in actions if a['needs_write']]
    ok = [a for a in actions if a['status'].startswith('[OK]')]
    unknown = [a for a in actions if 'unknown' in a['status']]
    incomplete = [a for a in actions if a.get('dataset_field') == 'NOT_RESOLVED']

    print(f"\n  Total sheets with Open Rows (999): {len(actions)}")
    print(f"  [OK] OK:                              {len(ok)}")
    print(f"  To set/fix:             {len(to_fill)}")
    if unknown:
        print(f"  [!] Key column unknown:           {len(unknown)}")
    if incomplete:
        print(f"  [!] Inline key, H2 vide:          {len(incomplete)}")
        print(f"    -> La formule rowKey sera écrite avec ;NOT_RESOLVED")
        print(f"    -> Il faut remplir H2 (allocationRuleCondition) manuellement")

    for a in actions:
        pattern = "Sheet per" if a['is_sheet_per'] else "Inline"
        existing_display = ""
        if a['rows']:
            ex_val = a['rows'][0][1]
            if ex_val:
                existing_display = str(ex_val).strip()

        print(f"\n  [{a['status']}]")
        print(f"    File    : {a['file']}")
        print(f"    Sheet   : {a['sheet']}")
        print(f"    Pattern : {pattern} (key col: {a['key_col_code'] or '?'})")
        if existing_display:
            print(f"    Current : {existing_display}")
        if a['expected']:
            print(f"    Expected: {a['expected']}")
        if a['dataset_field']:
            print(f"    Dataset : {a['dataset_field']}")
        if a['h2_value']:
            print(f"    H2      : {a['h2_value']}")
        print(f"    Source  : {a['dpm_source']}")


# -- GUI mode -------------------------------------------------------------------

def run_rowkey_gui(base_dir):
    """Interactive GUI mode for analyzing and writing rowKey."""
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk

    root = tk.Tk()
    root.withdraw()

    # -- Step 1: Select DPM file(s) --------------------------------------------
    messagebox.showinfo(
        "RowKey Tool",
        "Step 1/2: Select DPM file(s)\n"
        "(Annotated Table Layout .xlsx)\n\n"
        "You can select multiple files (Ctrl+click)."
    )

    layout_paths = filedialog.askopenfilenames(
        title="Select DPM file(s) (Annotated Table Layout)",
        initialdir=base_dir,
        filetypes=[("Excel files", "*.xlsx"), ("All", "*.*")]
    )

    if not layout_paths:
        print("Cancelled: No DPM file selected.")
        root.destroy()
        return

    for lp in layout_paths:
        print(f"  DPM selected : {os.path.basename(lp)}")

    # -- Loading data ---------------------------------------------------------
    print("Reading Open Rows from DPM files...")
    dpm_data = {}
    for lp in layout_paths:
        partial = read_dpm_open_rows(lp)
        for sheet_name, info in partial.items():
            if sheet_name not in dpm_data:
                dpm_data[sheet_name] = info
    print(f"  {len(dpm_data)} sheets with Open Rows found in DPM files")

    print("Scanning XLSM forms...")
    xlsm_items = scan_xlsm_rowkeys(base_dir)
    print(f"  {len(xlsm_items)} sheets with rowNumbering=999 in XLSM files")

    actions = match_and_prepare(xlsm_items, dpm_data)

    # -- Step 2: Selection window --------------------------------------
    sel_window = tk.Toplevel(root)
    sel_window.title("RowKey Tool — Fill rowKey")
    sel_window.geometry("1200x600")

    title_frame = tk.Frame(sel_window, padx=10, pady=5)
    title_frame.pack(fill=tk.X)
    tk.Label(title_frame,
             text="Step 2/2: Select rowKey to write",
             font=("Segoe UI", 12, "bold")).pack(anchor=tk.W)
    dpm_names = ', '.join(os.path.basename(lp) for lp in layout_paths)
    tk.Label(title_frame, text=f"DPM : {dpm_names}",
             font=("Segoe UI", 9), fg="gray").pack(anchor=tk.W)

    # -- Treeview -------------------------------------------------------------
    tree_frame = tk.Frame(sel_window, padx=10, pady=5)
    tree_frame.pack(fill=tk.BOTH, expand=True)

    columns = ("select", "file", "sheet", "pattern", "col_code", "existing",
               "expected", "dataset_field", "status")
    tree = ttk.Treeview(tree_frame, columns=columns, show='headings', selectmode='extended')

    tree.heading("select", text="[OK]")
    tree.heading("file", text="File")
    tree.heading("sheet", text="Sheet")
    tree.heading("pattern", text="Pattern")
    tree.heading("col_code", text="Key col")
    tree.heading("existing", text="Current")
    tree.heading("expected", text="Expected")
    tree.heading("dataset_field", text="Dataset.Field")
    tree.heading("status", text="Status")

    tree.column("select", width=30, anchor="center")
    tree.column("file", width=220)
    tree.column("sheet", width=100)
    tree.column("pattern", width=75, anchor="center")
    tree.column("col_code", width=60, anchor="center")
    tree.column("existing", width=200)
    tree.column("expected", width=250)
    tree.column("dataset_field", width=200)
    tree.column("status", width=110)

    scroll_y = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=tree.yview)
    scroll_x = ttk.Scrollbar(tree_frame, orient=tk.HORIZONTAL, command=tree.xview)
    tree.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)

    tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
    scroll_x.pack(side=tk.BOTTOM, fill=tk.X)

    selected_set = set()

    def populate_tree():
        tree.delete(*tree.get_children())
        for i, a in enumerate(actions):
            pattern = "Sheet per" if a['is_sheet_per'] else "Inline"
            existing = ""
            if a['rows']:
                ex_val = a['rows'][0][1]
                if ex_val:
                    existing = str(ex_val).strip()

            sel_str = "[x]" if i in selected_set else "[ ]"
            tag = 'selected' if i in selected_set else ''

            tree.insert("", tk.END, iid=str(i), values=(
                sel_str, a['file'], a['sheet'], pattern,
                a['key_col_code'] or '?', existing,
                a['expected'] or '—', a['dataset_field'] or '',
                a['status'],
            ), tags=(tag,))

        tree.tag_configure('selected', background='#d4edda')

    def toggle_select(event):
        item = tree.focus()
        if not item:
            return
        idx = int(item)
        if idx in selected_set:
            selected_set.discard(idx)
        else:
            if actions[idx]['expected']:
                selected_set.add(idx)
        populate_tree()
        update_count()

    # Auto-select items that need writing
    for i, a in enumerate(actions):
        if a['needs_write'] and a['expected']:
            selected_set.add(i)

    tree.bind('<Double-1>', toggle_select)
    populate_tree()

    # -- Manual dataset.field editing button -----------------------------
    edit_frame = tk.Frame(sel_window, padx=10, pady=3)
    edit_frame.pack(fill=tk.X)

    tk.Label(edit_frame, text="Dataset.Field :").pack(side=tk.LEFT)
    field_entry = tk.Entry(edit_frame, width=50)
    field_entry.pack(side=tk.LEFT, padx=5)

    def apply_field():
        item = tree.focus()
        if not item:
            messagebox.showwarning("RowKey Tool", "Select a row first.")
            return
        idx = int(item)
        a = actions[idx]
        field_val = field_entry.get().strip()

        if field_val:
            a['dataset_field'] = field_val
            a['expected'] = build_rowkey(a['sheet'], a['key_col_code'], field_val)
            a['h2_expected'] = f"{field_val} <> ('NULL')"
            a['h2_needs_write'] = not a['h2_value']
        else:
            a['dataset_field'] = None
            a['expected'] = build_rowkey(a['sheet'], a['key_col_code'])
            a['h2_expected'] = None
            a['h2_needs_write'] = False

        # Update status
        existing_values = [rk for _, rk in a['rows'] if rk]
        if existing_values:
            existing_clean = str(existing_values[0]).strip()
            if existing_clean == a['expected']:
                a['status'] = "[OK] OK"
                a['needs_write'] = False
            else:
                a['status'] = f"[!] to fix"
                a['needs_write'] = True
        else:
            a['status'] = "to set"
            a['needs_write'] = True

        if a['needs_write'] and a['expected']:
            selected_set.add(idx)

        populate_tree()
        update_count()

    tk.Button(edit_frame, text="Apply to selected row",
              command=apply_field).pack(side=tk.LEFT, padx=5)

    # -- Action buttons -----------------------------------------------------
    btn_frame = tk.Frame(sel_window, padx=10, pady=10)
    btn_frame.pack(fill=tk.X)

    count_label = tk.Label(btn_frame, text="", font=("Segoe UI", 9))
    count_label.pack(side=tk.LEFT, padx=10)

    def update_count():
        n = len(selected_set)
        count_label.config(text=f"{n} rowKey selected for writing")

    update_count()

    # Select All / Deselect All
    def select_all():
        for i, a in enumerate(actions):
            if a['expected']:
                selected_set.add(i)
        populate_tree()
        update_count()

    def deselect_all():
        selected_set.clear()
        populate_tree()
        update_count()

    tk.Button(btn_frame, text="Select All", command=select_all).pack(side=tk.LEFT, padx=2)
    tk.Button(btn_frame, text="Deselect All", command=deselect_all).pack(side=tk.LEFT, padx=2)

    def do_write():
        if not selected_set:
            messagebox.showwarning("RowKey Tool", "No rowKey selected.")
            return

        to_write = [actions[i] for i in sorted(selected_set)]
        msg = f"Write {len(to_write)} rowKey?\n\n"
        for a in to_write[:10]:
            msg += f"  {a['sheet']} -> {a['expected']}\n"
        if len(to_write) > 10:
            msg += f"  ... and {len(to_write) - 10} more\n"

        if not messagebox.askyesno("Confirm write", msg):
            return

        # Group by file
        by_file = {}
        for a in to_write:
            by_file.setdefault(a['file_path'], []).append(a)

        total_written = 0
        for fpath, file_actions in by_file.items():
            create_backup(fpath)
            wb = openpyxl.load_workbook(fpath, keep_vba=True)

            for a in file_actions:
                ws = wb[a['sheet']]
                for row_num, _ in a['rows']:
                    ws.cell(row_num, a['rowkey_col']).value = a['expected']
                    print(f"  [OK] {a['sheet']} [{row_num},{a['rowkey_col']}]: {a['expected']}")
                    total_written += 1

                if a['h2_needs_write'] and a['h2_expected']:
                    ws.cell(2, 8).value = a['h2_expected']
                    print(f"  [OK] {a['sheet']} H2: {a['h2_expected']}")

            wb.save(fpath)
            print(f"  -> Saved: {os.path.basename(fpath)}")
            wb.close()

        messagebox.showinfo("RowKey Tool",
                            f"Writing complete!\n{total_written} rowKey written.")

        # Refresh
        nonlocal xlsm_items, actions
        xlsm_items = scan_xlsm_rowkeys(base_dir)
        actions = match_and_prepare(xlsm_items, dpm_data)
        selected_set.clear()
        populate_tree()
        update_count()

    tk.Button(btn_frame, text="[~] Write selected rowKey",
              command=do_write, bg="#28a745", fg="white",
              font=("Segoe UI", 10, "bold")).pack(side=tk.RIGHT, padx=10)

    tk.Button(btn_frame, text="Cancel",
              command=lambda: (sel_window.destroy(), root.destroy())
              ).pack(side=tk.RIGHT, padx=5)

    sel_window.protocol("WM_DELETE_WINDOW",
                        lambda: (sel_window.destroy(), root.destroy()))
    root.mainloop()


######################################################################################################################################################
#                                                                                                                                                    #
#       Part 6 (final): Pipeline and CLI — full orchestration and entry points                                                           #
#                                                                                                                                                    #
#   collect_all_inputs: GUI for collecting all parameters (DPM, dictionary, form)                                                      #
#   run_generation / run_enrichment_step / run_rowkey_step: sequential execution of all 3 steps                                                     #
#   enrichment_main / rowkey_main: CLI entry points for enrichment and rowkey modes                                                      #
#   pipeline_main: full pipeline GUI orchestration with final report                                                                         #
#   main: dispatches to the appropriate mode (pipeline / enrichment / rowkey)                                                                         #
#                                                                                                                                                    #
######################################################################################################################################################

def collect_all_inputs():
    """
    Single GUI collecting all required inputs:
      - DPM source file (for forms_from_dpm + enrichment_tool)
      - YML dictionary (for enrichment_tool)
      - Form parameters (formCode, formGroupCode, etc.)
      - Options (debug mode, auto-write enrichment)
    """
    root = tk.Tk()
    root.withdraw()

    # -- 1. DPM source file selection ----------------------------------
    source_path = filedialog.askopenfilename(
        title="Select DPM source (Annotated Table Layout)",
        filetypes=[("Excel files", "*.xlsx *.xlsm *.xls")]
    )
    if not source_path:
        messagebox.showerror("Error", "No DPM file selected. Cancelled.")
        root.destroy()
        return None

    # -- 2. YML dictionary selection ------------------------------------
    source_dir = os.path.dirname(source_path)
    # Automatically search for dictionary in working directory
    dict_candidates = glob.glob(os.path.join(source_dir, 'pure-dictionary-*.yml'))
    default_dict = dict_candidates[0] if dict_candidates else ""

    dict_path = filedialog.askopenfilename(
        title="Select dictionary file (pure-dictionary .yml)",
        initialdir=source_dir,
        initialfile=os.path.basename(default_dict) if default_dict else "",
        filetypes=[("YAML files", "*.yml *.yaml"), ("All", "*.*")]
    )
    if not dict_path:
        messagebox.showerror("Error", "No dictionary selected. Cancelled.")
        root.destroy()
        return None

    # -- 2b. Template XLSM selection --------------------------------------
    # Automatically search for template in current directories
    template_search_dirs = [
        os.path.join(source_dir, '..', 'template'),
        os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'template'),
        source_dir,
    ]
    default_template = ""
    for d in template_search_dirs:
        candidate = os.path.join(d, 'pure-form-regulatoryOutputCode-formCode-(shortName).xlsm')
        if os.path.isfile(candidate):
            default_template = os.path.abspath(candidate)
            break

    template_path = filedialog.askopenfilename(
        title="Select template XLSM (pure-form-regulatoryOutputCode-...)",
        initialdir=os.path.dirname(default_template) if default_template else source_dir,
        initialfile=os.path.basename(default_template) if default_template else "",
        filetypes=[("Excel XLSM", "*.xlsm"), ("All", "*.*")]
    )
    if not template_path:
        messagebox.showerror("Error", "No template selected. Cancelled.")
        root.destroy()
        return None

    # -- 2c. Output directory selection --------------------------------
    # Default to generated_output folder in DPM's grandparent directory
    default_output = os.path.abspath(os.path.join(source_dir, '..', '..', 'generated_output'))
    if not os.path.isdir(default_output):
        default_output = source_dir

    output_dir = filedialog.askdirectory(
        title="Select output directory",
        initialdir=default_output,
    )
    if not output_dir:
        messagebox.showerror("Error", "No output directory selected. Cancelled.")
        root.destroy()
        return None

    # -- 3. Multi-field form ------------------------------------------
    input_window = tk.Toplevel(root)
    input_window.title("Pure Form Generator — Parameters")
    input_window.geometry("600x450")

    tk.Label(input_window, text="Generation + Enrichment",
             font=("Segoe UI", 12, "bold")).grid(row=0, column=0, columnspan=2, pady=10)

    # Auto-deduce formCode from dictionary domain (strip 'eu' prefix if present)
    domain = get_dict_domain(dict_path)
    default_form_code = domain[2:] if domain.startswith('eu') else domain

    fields = [
        ("formCode", default_form_code),
        ("formGroupCode", ""),
        ("version", ""),
        ("validityStartingDate", ""),
        ("linkedDomain", ""),
        ("datatype (identification J2)", ""),
        ("roundingUnit", ""),
    ]

    entries = {}
    for idx, (label_text, default) in enumerate(fields, start=1):
        Label(input_window, text=label_text, anchor="e").grid(row=idx, column=0, sticky="e", padx=5, pady=2)
        entry = Entry(input_window, width=50)
        entry.insert(0, default)
        entry.grid(row=idx, column=1, padx=5, pady=2)
        entries[label_text] = entry

    # Options
    debug_var = BooleanVar(value=False)
    auto_enrich_var = BooleanVar(value=True)

    row_opts = len(fields) + 1
    Checkbutton(input_window, text="Debug mode", variable=debug_var
                ).grid(row=row_opts, column=0, columnspan=2, sticky="w", padx=10, pady=5)
    Checkbutton(input_window, text="Auto-enrich (Z-axis + Datatypes)",
                variable=auto_enrich_var
                ).grid(row=row_opts + 1, column=0, columnspan=2, sticky="w", padx=10)

    auto_rowkey_var = BooleanVar(value=True)
    Checkbutton(input_window, text="Auto-fill rowKey (Open Rows)",
                variable=auto_rowkey_var
                ).grid(row=row_opts + 2, column=0, columnspan=2, sticky="w", padx=10)

    # Info
    tk.Label(input_window,
             text=f"DPM : {os.path.basename(source_path)}\n"
                  f"Dict : {os.path.basename(dict_path)}\n"
                  f"Template : {os.path.basename(template_path)}\n"
                  f"Output : {output_dir}",
             font=("Segoe UI", 9), fg="gray", justify="left"
             ).grid(row=row_opts + 3, column=0, columnspan=2, pady=5)

    result = [None]

    def submit():
        result[0] = {
            'source_path': source_path,
            'dict_path': dict_path,
            'formCode': entries["formCode"].get().strip(),
            'formGroupCode': entries["formGroupCode"].get().strip(),
            'version': entries["version"].get().strip(),
            'validityStartingDate': entries["validityStartingDate"].get().strip(),
            'linkedDomain': entries["linkedDomain"].get().strip(),
            'datatype': entries["datatype (identification J2)"].get().strip(),
            'roundingUnit': entries["roundingUnit"].get().strip(),
            'debug': debug_var.get(),
            'auto_enrich': auto_enrich_var.get(),
            'auto_rowkey': auto_rowkey_var.get(),
            'template_path': template_path,
            'output_dir': output_dir,
        }
        input_window.quit()
        input_window.destroy()

    Button(input_window, text="Start", command=submit,
           bg="#4CAF50", fg="white", font=("Segoe UI", 10, "bold")
           ).grid(row=row_opts + 4, column=0, columnspan=2, pady=15)

    input_window.protocol("WM_DELETE_WINDOW", lambda: (input_window.quit(), input_window.destroy()))
    input_window.mainloop()
    root.destroy()

    return result[0]


def run_generation(params):
    """
    Step 1: Generate pure-form XLSM files via forms_from_dpm.
    Returns the list of generated files and the output_dir.
    """
    import openpyxl
    import shutil

    source_path = params['source_path']
    form_code = params['formCode']

    # Configure forms_from_dpm debug mode
    forms_from_dpm.DEBUG_MODE = params['debug']
    forms_from_dpm.configure_logging()

    logging.info("")
    logging.info("=" * 80)
    logging.info("  STEP 1: GENERATING PURE-FORM FILES")
    logging.info("=" * 80)

    output_dir = params['output_dir']
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    logging.info(f"  Output directory: {output_dir}")

    # -- Monkey-patch copy_template_file to use the correct template --------
    _original_copy_template = forms_from_dpm.copy_template_file
    _template_path = os.path.abspath(params['template_path'])

    def _patched_copy_template(root_name, output_dir_arg, form_code_arg):
        target_filename = f"pure-form-{form_code_arg}-{root_name}-({root_name.replace('_', ' ')}).xlsm"
        target_path = os.path.join(output_dir_arg, target_filename)
        shutil.copy2(_template_path, target_path)
        logging.debug(f"   Copied template from: {_template_path}")
        logging.debug(f"   Copied template to: {target_path}")
        dedicated_logger = None
        if forms_from_dpm.EXCEL_LOGGER_AVAILABLE:
            try:
                dedicated_logger = forms_from_dpm.excel_logger.setup_excel_file_logging(
                    target_path, "forms_from_dpm")
                forms_from_dpm.excel_logger.log_excel_step(
                    target_path, "Template copied",
                    f"Template: {_template_path} -> Target: {target_path}")
            except Exception as e:
                logging.warning(f"Error initializing Excel logging: {e}")
        return target_path, dedicated_logger

    forms_from_dpm.copy_template_file = _patched_copy_template

    # Load DPM source
    logging.info(f"  Loading DPM: {os.path.basename(source_path)}")
    source_wb = openpyxl.load_workbook(source_path, data_only=True)

    relevant_sheets = forms_from_dpm.get_relevant_sheets(source_wb)
    grouped_sheets = forms_from_dpm.group_sheets_by_root_name(relevant_sheets)

    logging.info(f"  {len(relevant_sheets)} sheets retained, {len(grouped_sheets)} groups")

    generated_files = []

    for root_name, sheet_list in grouped_sheets.items():
        logging.info("")
        logging.info(f"  [>>] Groupe '{root_name}' : {sheet_list}")

        try:
            target_path, dedicated_logger = forms_from_dpm.copy_template_file(
                root_name, output_dir, form_code)
            if not target_path:
                logging.error(f"  Cannot create file for '{root_name}'")
                continue

            target_wb = openpyxl.load_workbook(target_path, keep_vba=True)
            forms_from_dpm.create_empty_target_sheets(target_wb, sheet_list)

            for idx, sheet_name in enumerate(sheet_list):
                # Only fill identification on the first sheet to avoid
                # overwriting D2/E2 with the last sheet's name
                if idx == 0:
                    forms_from_dpm.process_sheet_data(
                        source_wb, target_wb, sheet_name,
                        params['formGroupCode'], params['version'],
                        params['validityStartingDate'], params['linkedDomain'],
                        params['datatype'], params['roundingUnit'],
                        target_path, dedicated_logger
                    )
                else:
                    forms_from_dpm.process_sheet_data(
                        source_wb, target_wb, sheet_name,
                        '', '', '', '', '', '',
                        target_path, dedicated_logger
                    )

            forms_from_dpm.move_specific_sheets_to_end(target_wb)

            target_wb.save(target_path)
            target_wb.close()

            # Feeding rules (if available)
            if forms_from_dpm.FEEDING_RULES_AVAILABLE and forms_from_dpm.VALIDATION_FILE_PATH:
                try:
                    forms_from_dpm.feeding_rules.process_excel_with_validation_rules(
                        target_path,
                        validation_file_path=forms_from_dpm.VALIDATION_FILE_PATH,
                        debug_mode=params['debug'],
                        dedicated_logger=dedicated_logger
                    )
                except Exception as e:
                    logging.warning(f"  [!] Feeding rules error: {e}")

            # Save with Excel for border compatibility
            try:
                forms_from_dpm.autosave_with_excel(target_path)
            except Exception as e:
                logging.warning(f"  [!] Excel autosave failed for '{root_name}': {e} (file saved by openpyxl)")

            generated_files.append(target_path)
            logging.info(f"  [OK] Generated: {os.path.basename(target_path)}")

        except Exception as e:
            logging.error(f"  [ERR] Error group '{root_name}': {e}")
            continue

    source_wb.close()

    # Restore original function
    forms_from_dpm.copy_template_file = _original_copy_template

    logging.info("")
    logging.info(f"  Generation complete: {len(generated_files)} file(s)")

    return generated_files, output_dir


def run_enrichment_step(params, output_dir):
    """
    Step 2: Z-axis + Datatypes enrichment via 
    Returns: {
        'zaxis_not_resolved': [(file, sheet, hierarchy, variant), ...],
        'zaxis_resolved': [(file, sheet, type_name), ...],
        'dt_missing': [(file, sheet, col_code, label), ...],
        'dt_written': [(file, sheet, col_code, datatype), ...],
    }
    """
    report = {
        'zaxis_not_resolved': [],
        'zaxis_resolved': [],
        'dt_missing': [],
        'dt_written': [],
    }

    logging.info("")
    logging.info("=" * 80)
    logging.info("  STEP 2: Z-AXIS + DATATYPES ENRICHMENT")
    logging.info("=" * 80)

    dict_path = params['dict_path']
    domain = get_dict_domain(dict_path)

    logging.info(f"  Dictionary: {os.path.basename(dict_path)} (domain: {domain})")
    type_index, def_index = load_dictionary(dict_path)
    logging.info(f"  {len(type_index)} Z-axis types, {len(def_index)} datatype definitions")

    # -- Z-axis ---------------------------------------------------------------
    logging.info("")
    logging.info("  --- Z-AXIS ---")

    # Read Z-axis from DPM source
    ann_zaxis = read_annotated_layout(params['source_path'])
    logging.info(f"  {len(ann_zaxis)} sheets with Z-axis in DPM")

    # Scan generated XLSM files
    zaxis_items = scan_xlsm_zaxis(output_dir)
    logging.info(f"  {len(zaxis_items)} sheets with zAxis header in XLSM files")

    # Write Z-axis
    if zaxis_items:
        # For each Z-axis sheet, find the correct type
        z_written = 0
        by_file = {}
        for item in zaxis_items:
            by_file.setdefault(item['file_path'], []).append(item)

        for fpath, items in by_file.items():
            fname = os.path.basename(fpath)
            create_backup(fpath)

            try:
                wb = openpyxl.load_workbook(fpath, keep_vba=True)
            except Exception as e:
                logging.error(f"  [ERR] Error opening {fname}: {e}")
                continue

            for item in items:
                sn = item['sheet']
                if sn not in wb.sheetnames:
                    continue

                # Look up type via DPM
                type_name = None
                if sn in ann_zaxis:
                    info = ann_zaxis[sn]
                    type_name = map_variant_to_type(
                        info['hierarchy'], info['variant'], type_index)

                # Fallback: try via H2
                if not type_name:
                    ws = wb[sn]
                    h2 = ws.cell(row=2, column=8).value
                    if h2:
                        import re
                        h2_matches = re.findall(
                            r'(' + re.escape(domain) + r'\w*\.' + re.escape(domain) + r'(\w+))',
                            str(h2))
                        for full_path, tok in reversed(h2_matches):
                            for tn in type_index:
                                if tn.lower() == tok.lower():
                                    type_name = tn
                                    break
                            if type_name:
                                break

                if not type_name:
                    # Only mark NOT_RESOLVED if DPM says there should be a Z-axis
                    if sn in ann_zaxis:
                        ws = wb[sn]
                        old_j2 = ws.cell(row=2, column=10).value
                        dpm_info = ann_zaxis[sn]
                        report['zaxis_not_resolved'].append((
                            fname, sn,
                            dpm_info.get('hierarchy', '?'),
                            dpm_info.get('variant', '?'),
                        ))
                        if old_j2 and old_j2 != 'NOT_RESOLVED':
                            logging.info(f"  - {fname}[{sn}] : no Z-axis type found (existing J2: '{old_j2}')")
                        else:
                            ws.cell(row=2, column=10, value='NOT_RESOLVED')
                            ws.cell(row=4, column=10, value='NOT_RESOLVED')
                            logging.warning(f"  [!] {fname}[{sn}] : Z-axis present but no match found -> NOT_RESOLVED")
                            z_written += 1
                    else:
                        logging.info(f"  - {fname}[{sn}] : no Z-axis in DPM, skipped")
                    continue

                ws = wb[sn]
                attribute = f"{domain}.{type_name}"

                old_j2 = ws.cell(row=2, column=10).value
                ws.cell(row=2, column=10, value=type_name)
                ws.cell(row=4, column=10, value=attribute)

                logging.info(f"  [OK] {fname}[{sn}] : J2='{old_j2}'->'{type_name}' | J4->'{attribute}'")
                report['zaxis_resolved'].append((fname, sn, type_name))
                z_written += 1

            try:
                wb.save(fpath)
            except PermissionError:
                logging.error(f"  [ERR] '{fname}' locked")
            wb.close()

        logging.info(f"  Z-axis: {z_written} sheets updated")

    # -- Datatypes ------------------------------------------------------------
    logging.info("")
    logging.info("  --- DATATYPES ---")

    # Read DPM columns
    dpm_columns = read_dpm_columns(params['source_path'])
    logging.info(f"  {len(dpm_columns)} sheets with columns in DPM")

    # Scan XLSM files
    form_items = scan_xlsm_datatypes(output_dir)
    logging.info(f"  {len(form_items)} data sheets in XLSM files")

    if form_items:
        matches = match_dpm_to_form(dpm_columns, form_items)

        # Collect info on missing datatypes before writing
        for form_item, dpm_sheet, dpm_cols in matches:
            for code in form_item['col_codes']:
                existing = form_item['col_datatypes'].get(code)
                if existing:
                    continue
                if dpm_cols and code in dpm_cols:
                    suggested, confidence, source = determine_datatype(
                        dpm_cols[code], def_index, domain)
                else:
                    suggested = None
                if not suggested:
                    label = form_item['col_labels'].get(code, '')
                    report['dt_missing'].append((
                        form_item['file'], form_item['sheet'], code, label))

        # Write
        write_datatypes(matches, def_index, domain, output_dir)

    logging.info("")
    logging.info("  Enrichment complete")

    return report


def run_rowkey_step(params, output_dir):
    """
    Step 3: RowKey filling for Open Rows via 
    """
    logging.info("")
    logging.info("=" * 80)
    logging.info("  STEP 3: FILLING ROWKEY (OPEN ROWS)")
    logging.info("=" * 80)

    # Read Open Rows from DPM
    dpm_data = read_dpm_open_rows(params['source_path'])
    logging.info(f"  {len(dpm_data)} sheets with Open Rows in DPM")

    if not dpm_data:
        logging.info("  No sheets with Open Rows -> nothing to do")
        return 0

    # Scan generated XLSM files
    xlsm_items = scan_xlsm_rowkeys(output_dir)
    logging.info(f"  {len(xlsm_items)} sheets with rowNumbering=999 in XLSM files")

    if not xlsm_items:
        logging.info("  No sheets with rowNumbering=999 -> nothing to do")
        return 0

    # Cross-reference DPM and XLSM
    actions = match_and_prepare(xlsm_items, dpm_data)
    to_write = [a for a in actions if a['needs_write'] and a['expected']]

    logging.info(f"  {len(actions)} sheets analyzed, {len(to_write)} rowKey to write")

    if not to_write:
        logging.info("  All rowKey are already correct")
        return 0

    # Group by file and write
    import openpyxl
    by_file = {}
    for a in to_write:
        by_file.setdefault(a['file_path'], []).append(a)

    total_written = 0
    for fpath, file_actions in by_file.items():
        fname = os.path.basename(fpath)
        create_backup(fpath)

        try:
            wb = openpyxl.load_workbook(fpath, keep_vba=True)
        except Exception as e:
            logging.error(f"  [ERR] Error opening {fname}: {e}")
            continue

        for a in file_actions:
            ws = wb[a['sheet']]
            for row_num, _ in a['rows']:
                ws.cell(row_num, a['rowkey_col']).value = a['expected']
                logging.info(f"  [OK] {fname}[{a['sheet']}] B{row_num}: {a['expected']}")
                total_written += 1

            if a['h2_needs_write'] and a['h2_expected']:
                ws.cell(2, 8).value = a['h2_expected']
                logging.info(f"  [OK] {fname}[{a['sheet']}] H2: {a['h2_expected']}")

        try:
            wb.save(fpath)
            logging.info(f"  -> Saved: {fname}")
        except PermissionError:
            logging.error(f"  [ERR] '{fname}' locked")
        wb.close()

    logging.info(f"  RowKey: {total_written} cells updated")
    return total_written


# -- CLI: enrichment mode --------------------------------------------------

def enrichment_main():
    parser = argparse.ArgumentParser(
        description="Unified Z-axis + Datatypes enrichment tool for pure-form XLSM"
    )
    parser.add_argument('--write', action='store_true',
                        help="Write values to XLSM files")
    parser.add_argument('--zaxis-only', action='store_true',
                        help="Process Z-axis only")
    parser.add_argument('--datatype-only', action='store_true',
                        help="Process datatypes only")
    parser.add_argument('--set-zaxis', metavar='TYPE',
                        help="Set J2/J4/enum for a sheet. Ex: --set-zaxis typspelendingexp")
    parser.add_argument('--set-datatype', metavar='TYPE',
                        help="Set datatype for a column. Ex: --set-datatype ratio9Dec")
    parser.add_argument('--sheet', metavar='SHEET',
                        help="Target sheet for --set-zaxis or --set-datatype")
    parser.add_argument('--col', metavar='CODE',
                        help="Column code for --set-datatype. Ex: --col 0020")
    parser.add_argument('--dir', default=BASE_DIR,
                        help=f"Working directory (default: {BASE_DIR})")
    parser.add_argument('--gui', action='store_true',
                        help="Interactive GUI mode")
    args = parser.parse_args()

    base_dir = args.dir

    # -- GUI mode --
    if args.gui:
        run_enrichment_gui(base_dir)
        return

    # -- Find and load dictionary --
    dict_files = glob.glob(os.path.join(base_dir, 'pure-dictionary-*.yml'))
    if not dict_files:
        print(f"ERROR: No pure-dictionary-*.yml file found in {base_dir}", file=sys.stderr)
        sys.exit(1)
    dict_path = dict_files[0]

    domain = get_dict_domain(dict_path)

    print("Loading dictionary...")
    type_index, def_index = load_dictionary(dict_path)
    print(f"  {len(type_index)} Z-axis types, {len(def_index)} datatype definitions (domain: {domain})")

    # -- SET Z-axis mode --
    if args.set_zaxis:
        if not args.sheet:
            print("ERROR: --set-zaxis requires --sheet.", file=sys.stderr)
            sys.exit(1)
        print(f"\n{'=' * 80}")
        print("SET MODE: Z-axis")
        print(f"{'=' * 80}")
        ok = set_zaxis_for_sheet(args.sheet, args.set_zaxis, type_index, base_dir, domain)
        sys.exit(0 if ok else 1)

    # -- SET Datatype mode --
    if args.set_datatype:
        if not args.sheet or not args.col:
            print("ERROR: --set-datatype requires --sheet and --col.", file=sys.stderr)
            sys.exit(1)
        ok = set_datatype_for_column(args.sheet, args.col, args.set_datatype,
                                     def_index, domain, base_dir)
        sys.exit(0 if ok else 1)

    # -- REPORT / WRITE mode --
    do_zaxis = not args.datatype_only
    do_datatype = not args.zaxis_only

    if do_zaxis:
        print("\nScanning pure-form XLSM files (Z-axis)...")
        zaxis_items = scan_xlsm_zaxis(base_dir)
        print(f"  {len(zaxis_items)} sheets with zAxis header found")
        print_zaxis_report(zaxis_items, type_index)

        if args.write:
            print(f"\n{'=' * 80}")
            print("WRITING Z-AXIS CLASSES")
            print(f"{'=' * 80}")
            write_zaxis_values(zaxis_items, type_index, base_dir)

    if do_datatype:
        print("\nReading DPM (datatypes)...")
        dpm_files = sorted(glob.glob(os.path.join(base_dir, '*Annotated*Table*Layout*.xlsx')))
        if not dpm_files:
            print("  [!] No DPM files found. Report based on XLSM only.")
            dpm_columns = {}
        else:
            dpm_columns = {}
            for lp in dpm_files:
                print(f"  Reading {os.path.basename(lp)}...")
                partial = read_dpm_columns(lp)
                for sheet_name, cols in partial.items():
                    if sheet_name not in dpm_columns:
                        dpm_columns[sheet_name] = cols
            print(f"  {len(dpm_columns)} sheets found in DPM files")

        print("Scanning XLSM forms (datatypes)...")
        form_items = scan_xlsm_datatypes(base_dir)
        print(f"  {len(form_items)} data sheets found")

        if form_items:
            matches = match_dpm_to_form(dpm_columns, form_items)
            print_datatype_report(matches, def_index, domain)

            if args.write:
                print(f"\n{'=' * 90}")
                print("WRITING DATATYPES")
                print(f"{'=' * 90}")
                write_datatypes(matches, def_index, domain, base_dir)

    if not args.write:
        print(f"\n[i] To write, re-run with: python pure_form_generator.py enrichment --write")
        if not args.zaxis_only and not args.datatype_only:
            print(f"   Or: --write --zaxis-only  /  --write --datatype-only")


# -- CLI: rowkey mode ----------------------------------------------------------

def rowkey_main():
    parser = argparse.ArgumentParser(
        description="Fill rowKey in pure-form XLSM"
    )
    parser.add_argument('--gui', action='store_true',
                        help="Graphical interface")
    parser.add_argument('--write', action='store_true',
                        help="Write rowKey (batch mode)")
    parser.add_argument('--set-rowkey', type=str,
                        help="Manually write a rowKey")
    parser.add_argument('--sheet', type=str,
                        help="Target sheet (with --set-rowkey)")
    parser.add_argument('--dpm', type=str, nargs='+',
                        help="Path(s) to DPM file(s)")
    parser.add_argument('--dir', type=str, default=BASE_DIR,
                        help="Working directory")
    args = parser.parse_args()

    base_dir = args.dir

    if args.gui:
        run_rowkey_gui(base_dir)
        return

    if args.set_rowkey:
        # Manual write mode
        if not args.sheet:
            print("ERROR: --sheet required with --set-rowkey", file=sys.stderr)
            sys.exit(1)

        # Find the sheet
        fpath = None
        for fp in find_xlsm_files(base_dir):
            try:
                wb = openpyxl.load_workbook(fp, keep_vba=True)
                if args.sheet in wb.sheetnames:
                    fpath = fp
                wb.close()
                if fpath:
                    break
            except Exception:
                continue

        if not fpath:
            print(f"ERROR: Sheet '{args.sheet}' not found", file=sys.stderr)
            sys.exit(1)

        create_backup(fpath)
        wb = openpyxl.load_workbook(fpath, keep_vba=True)
        ws = wb[args.sheet]

        header_row = find_header_row(ws)
        if not header_row:
            print(f"ERROR: Header not found in {args.sheet}", file=sys.stderr)
            wb.close()
            sys.exit(1)

        # Find rowKey column
        rowkey_col = None
        rownumbering_col = None
        for c in range(1, 15):
            hv = ws.cell(header_row, c).value
            if hv:
                hv_s = str(hv).strip()
                if hv_s == 'rowKey':
                    rowkey_col = c
                elif hv_s == 'rowNumbering':
                    rownumbering_col = c

        if not rowkey_col:
            print(f"ERROR: rowKey column not found in {args.sheet}", file=sys.stderr)
            wb.close()
            sys.exit(1)

        # Find 999 rows and write
        written = 0
        for r in range(header_row + 1, ws.max_row + 1):
            if rownumbering_col:
                v = ws.cell(r, rownumbering_col).value
                if v is not None and str(v).strip() == '999':
                    ws.cell(r, rowkey_col).value = args.set_rowkey
                    print(f"  [OK] {args.sheet} [{r},{rowkey_col}]: {args.set_rowkey}")
                    written += 1

        if written:
            wb.save(fpath)
            print(f"  -> Saved: {os.path.basename(fpath)} ({written} rowKey)")
        else:
            print(f"  [!] No row 999 found in {args.sheet}")

        wb.close()
        return

    # Mode rapport / batch
    # Find DPM files
    dpm_paths = args.dpm
    if not dpm_paths:
        # Open a dialog to let the user choose DPM file(s)
        import tkinter as tk
        from tkinter import filedialog
        _root = tk.Tk()
        _root.withdraw()
        selected = filedialog.askopenfilenames(
            title="Select DPM file(s) (Annotated Table Layout)",
            initialdir=base_dir,
            filetypes=[("Excel files", "*.xlsx"), ("All", "*.*")]
        )
        _root.destroy()
        if selected:
            dpm_paths = list(selected)
        else:
            print("Cancelled: No DPM file selected.", file=sys.stderr)
            sys.exit(1)

    print("Reading Open Rows from DPM files...")
    dpm_data = {}
    for lp in dpm_paths:
        partial = read_dpm_open_rows(lp)
        for sheet_name, info in partial.items():
            if sheet_name not in dpm_data:
                dpm_data[sheet_name] = info
    print(f"  {len(dpm_data)} sheets with Open Rows found in DPM files")

    print("Scanning XLSM forms...")
    xlsm_items = scan_xlsm_rowkeys(base_dir)
    print(f"  {len(xlsm_items)} sheets with rowNumbering=999 in XLSM files")

    actions = match_and_prepare(xlsm_items, dpm_data)

    if args.write:
        to_write = [a for a in actions if a['needs_write'] and a['expected']]
        if not to_write:
            print("\n  Nothing to write: all rowKey are already correct.")
            return

        print(f"\n  Writing {len(to_write)} rowKey...")
        by_file = {}
        for a in to_write:
            by_file.setdefault(a['file_path'], []).append(a)

        for fpath, file_actions in by_file.items():
            create_backup(fpath)
            wb = openpyxl.load_workbook(fpath, keep_vba=True)

            for a in file_actions:
                ws = wb[a['sheet']]
                for row_num, _ in a['rows']:
                    ws.cell(row_num, a['rowkey_col']).value = a['expected']
                    print(f"  [OK] {a['sheet']} [{row_num},{a['rowkey_col']}]: {a['expected']}")

                if a['h2_needs_write'] and a['h2_expected']:
                    ws.cell(2, 8).value = a['h2_expected']
                    print(f"  [OK] {a['sheet']} H2: {a['h2_expected']}")

            wb.save(fpath)
            print(f"  -> Saved: {os.path.basename(fpath)}")
            wb.close()
    else:
        print_report(actions)


# -- Pipeline: main ------------------------------------------------------------

def pipeline_main():
    # -- Collect inputs ------------------------------------------------
    params = collect_all_inputs()
    if not params:
        return

    # Configure global logging
    level = logging.DEBUG if params['debug'] else logging.INFO
    for handler in logging.root.handlers[:]:
        logging.root.removeHandler(handler)
    logging.basicConfig(
        level=level,
        format='[%(levelname)s] %(message)s',
        handlers=[logging.StreamHandler()]
    )

    logging.info("+==============================================================+")
    logging.info("|   PURE FORM GENERATOR: Generation + Enrichment + RowKey   |")
    logging.info("+==============================================================+")
    logging.info(f"  DPM        : {os.path.basename(params['source_path'])}")
    logging.info(f"  Dictionary : {os.path.basename(params['dict_path'])}")
    logging.info(f"  Template   : {params['template_path']}")
    logging.info(f"  Output     : {params['output_dir']}")
    logging.info(f"  formCode   : {params['formCode'] or '(empty)'}")

    # -- Step 1: Generation -------------------------------------------------
    generated_files, output_dir = run_generation(params)

    if not generated_files:
        logging.error("No files generated. Process stopped.")
        return

    # -- Step 2: Enrichment (if requested) --------------------------------
    enrichment_report = None
    if params['auto_enrich']:
        enrichment_report = run_enrichment_step(params, output_dir)
    else:
        logging.info("")
        logging.info("  Enrichment disabled (checkbox unchecked)")

    # -- Step 3: RowKey (if requested) ----------------------------------------
    rowkey_count = 0
    if params['auto_rowkey']:
        rowkey_count = run_rowkey_step(params, output_dir)
    else:
        logging.info("")
        logging.info("  RowKey disabled (checkbox unchecked)")

    # -- Final summary ---------------------------------------------------------
    logging.info("")
    logging.info("=" * 80)
    logging.info("  FINAL SUMMARY")
    logging.info("=" * 80)
    logging.info(f"  Files generated: {len(generated_files)}")
    for i, f in enumerate(generated_files, 1):
        logging.info(f"    {i}. {os.path.basename(f)}")
    logging.info(f"  Directory: {output_dir}")
    if params['auto_enrich']:
        logging.info(f"  Enrichment: Z-axis + Datatypes applied")
    if params['auto_rowkey']:
        logging.info(f"  RowKey: {rowkey_count} cell(s) written")

    # -- Unresolved items report --------------------------------
    has_issues = False
    issue_summary_lines = []

    if enrichment_report:
        zaxis_nr = enrichment_report.get('zaxis_not_resolved', [])
        dt_miss = enrichment_report.get('dt_missing', [])

        if zaxis_nr:
            has_issues = True
            logging.info("")
            logging.info("  [!] UNRESOLVED Z-AXIS:")
            issue_summary_lines.append(f"Unresolved Z-axis: {len(zaxis_nr)}")
            for fname, sn, hierarchy, variant in zaxis_nr:
                logging.info(f"    -> {fname} [{sn}] : hierarchy={hierarchy}, variant={variant}")
                issue_summary_lines.append(f"  - [{sn}] ({hierarchy}:{variant})")

        if dt_miss:
            has_issues = True
            logging.info("")
            logging.info(f"  [!] UNRESOLVED DATATYPES ({len(dt_miss)} columns):")
            issue_summary_lines.append(f"Unresolved datatypes: {len(dt_miss)} columns")
            # Group by file/sheet for readability
            dt_by_sheet = {}
            for fname, sn, col_code, label in dt_miss:
                key = f"{fname} [{sn}]"
                dt_by_sheet.setdefault(key, []).append((col_code, label))
            for key, cols in dt_by_sheet.items():
                col_list = ', '.join(c for c, _ in cols)
                logging.info(f"    -> {key} : columns {col_list}")
                issue_summary_lines.append(f"  - {key} : {col_list}")

        if not has_issues:
            logging.info("")
            logging.info("  [OK] All Z-axis and datatypes resolved successfully")

    logging.info("=" * 80)
    logging.info("Generation complete!")

    # Final message
    try:
        import tkinter as tk2
        from tkinter import messagebox as mb2
        r = tk2.Tk()
        r.withdraw()
        summary = f"Generation complete!\n\n{len(generated_files)} file(s) generated\n"
        if params['auto_enrich']:
            summary += "Z-axis + Datatypes enriched\n"
        if params['auto_rowkey']:
            summary += f"RowKey: {rowkey_count} cell(s)\n"
        summary += f"\nDirectory: {output_dir}"
        if has_issues:
            summary += "\n\n[!] UNRESOLVED ITEMS:\n"
            summary += '\n'.join(issue_summary_lines)
        mb2.showinfo("Pure Form Generator", summary)
        r.destroy()
    except Exception:
        pass


# -- Unified entry point ------------------------------------------------------

def main():
    """Dispatches to the appropriate mode based on the first argument."""
    if len(sys.argv) > 1 and sys.argv[1] == 'enrichment':
        sys.argv.pop(1)
        enrichment_main()
    elif len(sys.argv) > 1 and sys.argv[1] == 'rowkey':
        sys.argv.pop(1)
        rowkey_main()
    else:
        pipeline_main()


if __name__ == '__main__':
    main()
