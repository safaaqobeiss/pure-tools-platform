#!/usr/bin/env python3
"""
Headless CLI runner for the PURE Form Generator pipeline.
Accepts all inputs via environment variables (no tkinter).

Environment variables:
  PURE_NO_DIALOG=1           Required — disables all tkinter dialogs
  PURE_DPM_SOURCE            Path to DPM source file (.xlsx)
  PURE_DICT_PATH             Path to dictionary file (.yml)
  PURE_TEMPLATE_PATH         Path to template XLSM
  PURE_OUTPUT_DIR            Path to output directory
  PURE_FORM_CODE             formCode (e.g. "corep")
  PURE_FORM_GROUP_CODE       formGroupCode (optional)
  PURE_VERSION               version (optional)
  PURE_VALIDITY_DATE         validityStartingDate (optional)
  PURE_LINKED_DOMAIN         linkedDomain (optional)
  PURE_DATATYPE              datatype for identification J2 (optional)
  PURE_ROUNDING_UNIT         roundingUnit (optional)
  PURE_AUTO_ENRICH           "1" to enable enrichment (default "1")
  PURE_AUTO_ROWKEY           "1" to enable rowkey filling (default "1")
  PURE_DEBUG                 "1" for debug mode (default "0")
"""

import logging
import os
import sys

# Ensure the script's directory is in the path so that forms_from_dpm can be imported
script_dir = os.path.dirname(os.path.abspath(__file__))
if script_dir not in sys.path:
    sys.path.insert(0, script_dir)

from pure_form_generator import (
    run_generation,
    run_enrichment_step,
    run_rowkey_step,
    create_backup,
)


def main():
    if os.environ.get("PURE_NO_DIALOG") != "1":
        print("ERROR: PURE_NO_DIALOG must be set to 1 for headless mode.", file=sys.stderr)
        sys.exit(1)

    source_path = os.environ.get("PURE_DPM_SOURCE", "")
    dict_path = os.environ.get("PURE_DICT_PATH", "")
    template_path = os.environ.get("PURE_TEMPLATE_PATH", "")
    output_dir = os.environ.get("PURE_OUTPUT_DIR", "")
    form_code = os.environ.get("PURE_FORM_CODE", "")
    form_group_code = os.environ.get("PURE_FORM_GROUP_CODE", "")
    version = os.environ.get("PURE_VERSION", "")
    validity_date = os.environ.get("PURE_VALIDITY_DATE", "")
    linked_domain = os.environ.get("PURE_LINKED_DOMAIN", "")
    datatype = os.environ.get("PURE_DATATYPE", "")
    rounding_unit = os.environ.get("PURE_ROUNDING_UNIT", "")
    auto_enrich = os.environ.get("PURE_AUTO_ENRICH", "1") == "1"
    auto_rowkey = os.environ.get("PURE_AUTO_ROWKEY", "1") == "1"
    debug = os.environ.get("PURE_DEBUG", "0") == "1"

    # Validate required inputs
    if not source_path or not os.path.isfile(source_path):
        print(f"ERROR: PURE_DPM_SOURCE is missing or invalid: '{source_path}'", file=sys.stderr)
        sys.exit(1)
    if not dict_path or not os.path.isfile(dict_path):
        print(f"ERROR: PURE_DICT_PATH is missing or invalid: '{dict_path}'", file=sys.stderr)
        sys.exit(1)
    if not template_path or not os.path.isfile(template_path):
        print(f"ERROR: PURE_TEMPLATE_PATH is missing or invalid: '{template_path}'", file=sys.stderr)
        sys.exit(1)
    if not output_dir:
        print("ERROR: PURE_OUTPUT_DIR is required.", file=sys.stderr)
        sys.exit(1)

    os.makedirs(output_dir, exist_ok=True)

    # Configure logging
    level = logging.DEBUG if debug else logging.INFO
    for handler in logging.root.handlers[:]:
        logging.root.removeHandler(handler)
    logging.basicConfig(
        level=level,
        format='[%(levelname)s] %(message)s',
        handlers=[logging.StreamHandler()],
    )

    params = {
        'source_path': source_path,
        'dict_path': dict_path,
        'template_path': template_path,
        'output_dir': output_dir,
        'formCode': form_code,
        'formGroupCode': form_group_code,
        'version': version,
        'validityStartingDate': validity_date,
        'linkedDomain': linked_domain,
        'datatype': datatype,
        'roundingUnit': rounding_unit,
        'auto_enrich': auto_enrich,
        'auto_rowkey': auto_rowkey,
        'debug': debug,
    }

    logging.info("[OK] PURE FORM GENERATOR: Generation + Enrichment + RowKey")
    logging.info(f"  DPM        : {os.path.basename(source_path)}")
    logging.info(f"  Dictionary : {os.path.basename(dict_path)}")
    logging.info(f"  Template   : {os.path.basename(template_path)}")
    logging.info(f"  Output     : {output_dir}")
    logging.info(f"  formCode   : {form_code or '(empty)'}")

    # Step 1: Generation
    generated_files, output_dir = run_generation(params)
    if not generated_files:
        logging.error("No files generated. Process stopped.")
        sys.exit(1)

    # Step 2: Enrichment
    enrichment_report = None
    if auto_enrich:
        enrichment_report = run_enrichment_step(params, output_dir)
    else:
        logging.info("  Enrichment disabled")

    # Step 3: RowKey
    rowkey_count = 0
    if auto_rowkey:
        rowkey_count = run_rowkey_step(params, output_dir)
    else:
        logging.info("  RowKey disabled")

    # Summary
    logging.info("")
    logging.info("=" * 80)
    logging.info("  FINAL SUMMARY")
    logging.info("=" * 80)
    logging.info(f"  Files generated: {len(generated_files)}")
    for i, f in enumerate(generated_files, 1):
        logging.info(f"    {i}. {os.path.basename(f)}")
    logging.info(f"  Directory: {output_dir}")
    if auto_enrich:
        logging.info("  Enrichment: Z-axis + Datatypes applied")
    if auto_rowkey:
        logging.info(f"  RowKey: {rowkey_count} cell(s) written")
    logging.info("=" * 80)
    logging.info("Generation complete!")


if __name__ == '__main__':
    main()
