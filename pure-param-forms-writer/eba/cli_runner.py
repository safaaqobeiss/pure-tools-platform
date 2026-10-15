"""
Headless CLI runner for EBA forms_from_dpm.
Bypasses tkinter dialogs by monkey-patching get_user_input().
Reads parameters from environment variables.
"""
import os
import sys
import shutil
from pathlib import Path


def main():
    source_file = os.environ.get("PURE_DPM_SOURCE", "")
    output_dir = os.environ.get("PURE_OUTPUT_DIR", "")
    form_code = os.environ.get("PURE_FORM_CODE", "")
    form_group_code = os.environ.get("PURE_FORM_GROUP_CODE", "")
    version = os.environ.get("PURE_VERSION", "")
    validity_date = os.environ.get("PURE_VALIDITY_DATE", "")
    linked_domain = os.environ.get("PURE_LINKED_DOMAIN", "")
    datatype = os.environ.get("PURE_DATATYPE", "euebacorep.positiveOrNullAmount")
    rounding_unit = os.environ.get("PURE_ROUNDING_UNIT", "")
    validation_rules = os.environ.get("PURE_VALIDATION_RULES", "")

    if not source_file or not Path(source_file).exists():
        print(f"ERROR: DPM source file not found: {source_file}", file=sys.stderr)
        sys.exit(1)

    # Add the eba directory to path so imports work
    eba_dir = Path(__file__).resolve().parent
    sys.path.insert(0, str(eba_dir))

    # Import the module but prevent tkinter from being used
    import forms_from_dpm

    # Override get_user_input to return our values without GUI
    def headless_get_user_input():
        forms_from_dpm.DEBUG_MODE = False
        forms_from_dpm.configure_logging()
        return (
            source_file,
            form_code,
            form_group_code,
            version,
            validity_date,
            linked_domain,
            datatype,
            rounding_unit,
        )

    forms_from_dpm.get_user_input = headless_get_user_input

    # Override create_output_directory to use our output dir
    def headless_create_output_directory(source_path):
        Path(output_dir).mkdir(parents=True, exist_ok=True)
        return output_dir

    forms_from_dpm.create_output_directory = headless_create_output_directory

    # Override validation rules file selection
    if validation_rules and Path(validation_rules).exists():
        forms_from_dpm.VALIDATION_FILE_PATH = validation_rules

    # Run the main logic
    forms_from_dpm.main()

    print(f"\nGeneration complete in: {output_dir}")


if __name__ == "__main__":
    main()
