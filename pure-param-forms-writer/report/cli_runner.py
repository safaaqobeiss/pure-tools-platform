"""
Headless CLI runner for Report forms_from_report.
Bypasses tkinter dialogs by monkey-patching get_user_input().
Reads parameters from environment variables.
"""
import os
import re
import sys
import shutil
from pathlib import Path


def main():
    source_file = os.environ.get("PURE_REPORT_FORM", "")
    output_dir = os.environ.get("PURE_OUTPUT_DIR", "")
    form_group_code = os.environ.get("PURE_FORM_GROUP_CODE", "schemaaXYZ")
    version = os.environ.get("PURE_VERSION", "2026.1.0")
    validity_date = os.environ.get("PURE_VALIDITY_DATE", "01/01/2026")
    linked_domain = os.environ.get("PURE_LINKED_DOMAIN", "bebnbschemaa")
    datatype = os.environ.get("PURE_DATATYPE", "bebnbschemaa.genericPositiveAmount")
    rounding_unit = os.environ.get("PURE_ROUNDING_UNIT", "thousands")
    convert_rows = os.environ.get("PURE_CONVERT_ROWS", "0")

    if not source_file or not Path(source_file).exists():
        print(f"ERROR: Report form file not found: {source_file}", file=sys.stderr)
        sys.exit(1)

    # Extract formCode from filename
    input_filename = os.path.basename(source_file)
    base_name = os.path.splitext(input_filename)[0]
    match = re.search(r'FRM_(\d{3,4})', base_name)
    form_code = match.group(1) if match else ""
    form_code_length = len(form_code) if form_code else 0

    # Determine conversion flag
    if form_code_length == 4:
        conversion_6_4_digits = "N"
    else:
        conversion_6_4_digits = "Y" if convert_rows == "1" else "N"

    # Add the report directory to path so imports work
    report_dir = Path(__file__).resolve().parent
    sys.path.insert(0, str(report_dir))

    # Import the module
    import forms_from_report

    # Override get_user_input to return our values without GUI
    def headless_get_user_input():
        forms_from_report.DEBUG_MODE = False
        forms_from_report.configure_logging()
        return (
            source_file,
            form_code,
            form_code_length,
            conversion_6_4_digits,
            form_group_code,
            version,
            validity_date,
            linked_domain,
            datatype,
            rounding_unit,
        )

    forms_from_report.get_user_input = headless_get_user_input

    # Override create_output_directory to use our output dir
    def headless_create_output_directory(source_path):
        Path(output_dir).mkdir(parents=True, exist_ok=True)
        return output_dir

    forms_from_report.create_output_directory = headless_create_output_directory

    # Run the main logic
    forms_from_report.main()

    print(f"\nGeneration complete in: {output_dir}")


if __name__ == "__main__":
    main()
