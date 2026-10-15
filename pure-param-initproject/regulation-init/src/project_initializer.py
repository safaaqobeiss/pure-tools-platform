"""
Orchestrates the full project initialization workflow:
    1. Build directory trees for Regulation & Output projects
    2. Copy exact template-shaped files from local reference folders
"""
import logging
import shutil
from copy import copy
from pathlib import Path
from openpyxl import load_workbook

from config import AppConfig
from models import ProjectInput

logger = logging.getLogger(__name__)


class ProjectInitializer:
    """Creates and populates the Regulation and Output project trees."""

    def __init__(self, config: AppConfig, project_input: ProjectInput):
        self.config = config
        self.input = project_input
        self.reg = project_input.regulation
        src_root = Path(__file__).resolve().parent
        self.regulation_reference_dir = src_root / "regulation"
        self.output_reference_dir = src_root / "Output"
        self.dictionary_source_file = src_root / "pure-dictionary-euebaliquidity.xlsx"
        self.regulation_source_file = src_root / "pure-regulation-euebaliquidity.xlsx"

    # ------------------------------------------------------------------ #
    # Public entry point
    # ------------------------------------------------------------------ #

    def run(self) -> Path:
        """
        Execute the full initialization.
        Returns the root directory (output_base_dir).
        """
        base_name = (self.config.project_base_name or self.reg.code).strip()
        root = self.config.output_base_dir
        root.mkdir(parents=True, exist_ok=True)

        reg_folder = self.config.regulation_project_name or f"pure-regulation-{base_name}"
        out_folder = self.config.output_project_name or f"pure-regulatoryoutput-{base_name}"
        regulation_dir = root / reg_folder
        output_dir = root / out_folder

        # 1 – Create selected project structures
        if self.config.project_type in ("both", "regulation"):
            self._create_dirs(regulation_dir, [])
            self._init_regulation_project(regulation_dir, self.config.regulation_folder_name)
            self._create_e2e_test_data(regulation_dir)

        if self.config.project_type in ("both", "output"):
            self._create_dirs(output_dir, [])
            self._init_output_project(output_dir, self.config.output_folder_name)
            self._create_e2e_test_data(output_dir)

        logger.info("Project initialization complete → %s", root)
        return root

    @staticmethod
    def _create_e2e_test_data(project_dir: Path):
        """Create the e2e-test-data directory structure with .gitkeep files."""
        e2e_dir = project_dir / "e2e-test-data"
        expected_dir = e2e_dir / "expected"
        load_dir = e2e_dir / "input" / "load"
        expected_dir.mkdir(parents=True, exist_ok=True)
        load_dir.mkdir(parents=True, exist_ok=True)
        (expected_dir / ".gitkeep").touch(exist_ok=True)
        (load_dir / ".gitkeep").touch(exist_ok=True)
        logger.info("Created e2e-test-data in %s", project_dir)

    # ------------------------------------------------------------------ #
    # Regulation project
    # ------------------------------------------------------------------ #

    def _init_regulation_project(self, base: Path, folder_name: str):
        target = base if not folder_name else (base / folder_name)
        self._create_empty_files_from_reference(
            self.regulation_reference_dir,
            target,
            self.config.regulation_file_map,
        )
        self._copy_gitlab_folder(self.regulation_reference_dir, base)
        self._fill_regulation_sheets(target)
        self._fill_dictionary_standard_sheets(target)
        self._fill_dictionary_identification(target)
        self._fill_transformation_rules_identification(target)
        self._fill_transformed_data_identification(target)
        self._fill_ui_dashboard_identification(target)

    # ------------------------------------------------------------------ #
    # Output project
    # ------------------------------------------------------------------ #

    def _init_output_project(self, base: Path, folder_name: str):
        target = base if not folder_name else (base / folder_name)
        self._create_empty_files_from_reference(
            self.output_reference_dir,
            target,
            self.config.output_file_map,
        )
        self._copy_gitlab_folder(self.output_reference_dir, base)
        self._fill_output_sheets(target)
        self._fill_output_structure_identification(target)

    # ------------------------------------------------------------------ #
    # Helpers
    # ------------------------------------------------------------------ #

    @staticmethod
    def _copy_gitlab_folder(reference_dir: Path, project_dir: Path):
        """Copy the .gitlab folder from the reference directory into the project."""
        gitlab_src = reference_dir / ".gitlab"
        gitlab_dest = project_dir / ".gitlab"
        if gitlab_src.exists() and gitlab_src.is_dir():
            shutil.copytree(gitlab_src, gitlab_dest, dirs_exist_ok=True)
            logger.info("Copied .gitlab folder to %s", gitlab_dest)
        else:
            logger.warning(".gitlab folder not found in %s", reference_dir)

    @staticmethod
    def _create_dirs(base: Path, subdirs: list[str]):
        base.mkdir(parents=True, exist_ok=True)
        for d in subdirs:
            (base / d).mkdir(parents=True, exist_ok=True)
        logger.info("Created directory tree: %s", base)

    @staticmethod
    def _create_empty_files_from_reference(
        reference_dir: Path,
        dest_dir: Path,
        file_map: dict[str, str] | None = None,
    ):
        dest_dir.mkdir(parents=True, exist_ok=True)
        if not reference_dir.exists():
            logger.warning("Reference directory not found: %s", reference_dir)
            return

        created = 0
        if file_map:
            for src_name, dest_name in file_map.items():
                ref_file = reference_dir / src_name
                if not ref_file.exists() or not ref_file.is_file():
                    logger.warning("Reference file not found: %s", ref_file)
                    continue
                target = dest_dir / dest_name
                try:
                    shutil.copy2(ref_file, target)
                    created += 1
                except PermissionError:
                    logger.error(
                        "Impossible de copier %s → %s. "
                        "Fermez le fichier s'il est ouvert dans Excel.",
                        ref_file.name, target,
                    )
                    raise SystemExit(
                        f"\nErreur: le fichier '{dest_name}' est verrouillé.\n"
                        f"Fermez-le dans Excel puis relancez l'outil."
                    )
        else:
            for ref_file in sorted(reference_dir.iterdir()):
                if not ref_file.is_file():
                    continue
                target = dest_dir / ref_file.name
                try:
                    shutil.copy2(ref_file, target)
                    created += 1
                except PermissionError:
                    logger.error(
                        "Impossible de copier %s → %s. "
                        "Fermez le fichier s'il est ouvert dans Excel.",
                        ref_file.name, target,
                    )
                    raise SystemExit(
                        f"\nErreur: le fichier '{ref_file.name}' est verrouillé.\n"
                        f"Fermez-le dans Excel puis relancez l'outil."
                    )

        logger.info("Created %d empty file(s) in %s", created, dest_dir)

    def _fill_dictionary_identification(self, regulation_target_dir: Path):
        data = self.config.dictionary_identification or {}
        if not data:
            return

        file_map = self.config.regulation_file_map or {}
        dictionary_target_name = file_map.get("pure-dictionary-regulatorCode.xlsx")
        if not dictionary_target_name:
            return

        dictionary_path = regulation_target_dir / dictionary_target_name
        if not dictionary_path.exists():
            logger.warning("Dictionary target file not found for identification fill: %s", dictionary_path)
            return

        try:
            wb = load_workbook(dictionary_path)
        except Exception as exc:
            logger.warning("Unable to open dictionary workbook for identification fill (%s): %s", dictionary_path, exc)
            return

        ws = None
        for name in wb.sheetnames:
            if name.strip().lower() == "identification":
                ws = wb[name]
                break

        if ws is None:
            logger.warning("Sheet 'identification' not found in %s", dictionary_path)
            return

        header_positions: dict[str, tuple[int, int]] = {}
        targets = ["version", "validityStartingDate", "description (EN)", "description (FR)", "regulatorName (EN)", "regulatorName (FR)"]
        for row in range(1, min(40, ws.max_row + 10)):
            for col in range(1, min(80, ws.max_column + 10)):
                value = ws.cell(row=row, column=col).value
                if value is None:
                    continue
                text = str(value).strip()
                for key in targets:
                    if text == key and key not in header_positions:
                        header_positions[key] = (row, col)

        missing = [k for k in targets if k not in header_positions]
        if missing:
            logger.warning("Missing identification headers in dictionary workbook: %s", ", ".join(missing))

        for key, val in data.items():
            if key not in header_positions:
                continue
            row, col = header_positions[key]
            ws.cell(row=row + 1, column=col, value=val)

        wb.save(dictionary_path)
        logger.info("Dictionary identification sheet populated: %s", dictionary_path)

    def _fill_transformation_rules_identification(self, regulation_target_dir: Path):
        data = self.config.transformation_rules_identification or {}
        if not data:
            return

        file_map = self.config.regulation_file_map or {}
        target_name = file_map.get("pure-transformationRules-regulatorCode.xlsx")
        if not target_name:
            return

        path = regulation_target_dir / target_name
        if not path.exists():
            logger.warning("TransformationRules target file not found for identification fill: %s", path)
            return

        try:
            wb = load_workbook(path)
        except Exception as exc:
            logger.warning("Unable to open transformationRules workbook for identification fill (%s): %s", path, exc)
            return

        self._fill_identification_sheet_generic(
            wb=wb,
            sheet_name="identification",
            values=data,
            required_headers=["version", "validityStartingDate", "description (EN)", "description (FR)"],
            workbook_label=str(path),
        )

        wb.save(path)
        logger.info("TransformationRules identification sheet populated: %s", path)

    def _fill_transformed_data_identification(self, regulation_target_dir: Path):
        data = self.config.transformed_data_identification or {}
        if not data:
            return

        file_map = self.config.regulation_file_map or {}
        target_name = file_map.get("pure-transformedDataStructure-regulatoryDomainCode.xlsm")
        if not target_name:
            return

        path = regulation_target_dir / target_name
        if not path.exists():
            logger.warning("TransformedDataStructure target file not found for identification fill: %s", path)
            return

        try:
            wb = load_workbook(path, keep_vba=True)
        except Exception as exc:
            logger.warning("Unable to open transformedDataStructure workbook for identification fill (%s): %s", path, exc)
            return

        self._fill_identification_sheet_generic(
            wb=wb,
            sheet_name="identification",
            values=data,
            required_headers=["version", "validityStartingDate", "description (EN)", "description (FR)"],
            workbook_label=str(path),
        )

        wb.save(path)
        logger.info("TransformedDataStructure identification sheet populated: %s", path)

    def _fill_ui_dashboard_identification(self, regulation_target_dir: Path):
        data = self.config.ui_dashboard_identification or {}
        if not data:
            return

        file_map = self.config.regulation_file_map or {}
        target_name = file_map.get("pure-uiDashboard-regulatoryDomainCode.xlsx")
        if not target_name:
            return

        path = regulation_target_dir / target_name
        if not path.exists():
            logger.warning("UiDashboard target file not found for identification fill: %s", path)
            return

        try:
            wb = load_workbook(path)
        except Exception as exc:
            logger.warning("Unable to open uiDashboard workbook for identification fill (%s): %s", path, exc)
            return

        self._fill_identification_sheet_generic(
            wb=wb,
            sheet_name="identification",
            values=data,
            required_headers=["version", "validityStartingDate", "description (EN)", "description (FR)"],
            workbook_label=str(path),
        )

        wb.save(path)
        logger.info("UiDashboard identification sheet populated: %s", path)

    def _fill_dictionary_standard_sheets(self, regulation_target_dir: Path):
        """Copy standard dictionary content sheets from source dictionary workbook."""
        file_map = self.config.regulation_file_map or {}
        dictionary_target_name = file_map.get("pure-dictionary-regulatorCode.xlsx")
        if not dictionary_target_name:
            return

        target_path = regulation_target_dir / dictionary_target_name
        source_path = self.dictionary_source_file
        if not target_path.exists():
            logger.warning("Dictionary target file not found for standard sheet copy: %s", target_path)
            return
        if not source_path.exists():
            logger.warning("Dictionary source file not found: %s", source_path)
            return

        try:
            src_wb = load_workbook(source_path)
            dst_wb = load_workbook(target_path)
        except Exception as exc:
            logger.warning("Unable to open dictionary workbook(s) for standard copy: %s", exc)
            return

        for sheet_name in ("patterns", "attributes", "enumerations"):
            src_ws = self._find_sheet_case_insensitive(src_wb, sheet_name)
            dst_ws = self._find_sheet_case_insensitive(dst_wb, sheet_name)
            if src_ws is None or dst_ws is None:
                logger.warning("Sheet '%s' missing in source or target dictionary workbook", sheet_name)
                continue
            self._copy_worksheet_content(src_ws, dst_ws)

        dst_wb.save(target_path)
        logger.info("Dictionary sheets patterns/attributes/enumerations populated from %s", source_path)

    def _fill_regulation_sheets(self, regulation_target_dir: Path):
        data = self.config.regulation_identification or {}
        if not data:
            return

        file_map = self.config.regulation_file_map or {}
        target_name = file_map.get("pure-regulation-regulationCode.xlsx")
        if not target_name:
            return

        path = regulation_target_dir / target_name
        if not path.exists():
            logger.warning("Regulation target file not found for sheet fill: %s", path)
            return

        try:
            wb = load_workbook(path)
        except Exception as exc:
            logger.warning("Unable to open regulation workbook for fill (%s): %s", path, exc)
            return

        self._fill_identification_sheet_generic(
            wb=wb,
            sheet_name="identification",
            values=data,
            required_headers=[
                "regulatorCode",
                "version",
                "validityStartingDate",
                "name (EN)",
                "name (FR)",
                "description (EN)",
                "description (FR)",
            ],
            workbook_label=str(path),
        )

        self._fill_reference_document_sheet(wb=wb, values=data, workbook_label=str(path))
        self._fill_institution_characteristics_sheet(wb=wb, workbook_label=str(path))
        self._fill_precision_sheet(wb=wb, workbook_label=str(path))
        self._clear_validationruleset_sheet(wb=wb, workbook_label=str(path))
        self._fill_formgroup_sheet(wb=wb, values=data, workbook_label=str(path))

        wb.save(path)
        logger.info("Regulation workbook populated (identification/referenceDocument/institutionCharacteristics/precision/validationRuleSet/formgroup): %s", path)

    def _fill_output_sheets(self, output_target_dir: Path):
        data = self.config.output_identification or {}
        if not data:
            return

        file_map = self.config.output_file_map or {}
        target_name = file_map.get("pure-regulatoryOutput-regulationCode-regulatoryOutputCode.xlsx")
        if not target_name:
            return

        path = output_target_dir / target_name
        if not path.exists():
            logger.warning("Output target file not found for sheet fill: %s", path)
            return

        try:
            wb = load_workbook(path)
        except Exception as exc:
            logger.warning("Unable to open output workbook for fill (%s): %s", path, exc)
            return

        self._fill_identification_sheet_generic(
            wb=wb,
            sheet_name="identification",
            values=data,
            required_headers=[
                "code",
                "version",
                "validityStartingDate",
                "periodicity",
                "commonCollection",
                "outputType",
                "name (EN)",
                "name (FR)",
                "description (EN)",
                "description (FR)",
            ],
            workbook_label=str(path),
        )

        self._fill_submission_fileset_sheet(
            wb=wb,
            submission_format=data.get("submissionFormat", ""),
            container_type=data.get("containerType", ""),
            workbook_label=str(path),
        )

        wb.save(path)
        logger.info("Output workbook populated (identification/submissionFileset): %s", path)

    def _fill_output_structure_identification(self, output_target_dir: Path):
        data = self.config.output_structure_identification or {}
        if not data:
            return

        file_map = self.config.output_file_map or {}
        target_name = file_map.get("pure-outputStructure-regulatoryOutputCode-outputStructureCode.xlsx")
        if not target_name:
            return

        path = output_target_dir / target_name
        if not path.exists():
            logger.warning("OutputStructure target file not found for identification fill: %s", path)
            return

        try:
            wb = load_workbook(path)
        except Exception as exc:
            logger.warning("Unable to open outputStructure workbook for identification fill (%s): %s", path, exc)
            return

        self._fill_identification_sheet_generic(
            wb=wb,
            sheet_name="identification",
            values=data,
            required_headers=["version", "validityStartingDate", "name (EN)", "name (FR)", "description (EN)", "description (FR)"],
            workbook_label=str(path),
        )

        wb.save(path)
        logger.info("OutputStructure identification sheet populated: %s", path)

    @staticmethod
    def _fill_identification_sheet_generic(
        wb,
        sheet_name: str,
        values: dict[str, str],
        required_headers: list[str],
        workbook_label: str,
    ):
        ws = ProjectInitializer._find_sheet_case_insensitive(wb, sheet_name)
        if ws is None:
            logger.warning("Sheet '%s' not found in %s", sheet_name, workbook_label)
            return

        header_positions: dict[str, tuple[int, int]] = {}
        for row in range(1, min(50, ws.max_row + 10)):
            for col in range(1, min(120, ws.max_column + 10)):
                value = ws.cell(row=row, column=col).value
                if value is None:
                    continue
                text = str(value).strip()
                for key in required_headers:
                    if text == key and key not in header_positions:
                        header_positions[key] = (row, col)

        missing = [k for k in required_headers if k not in header_positions]
        if missing:
            logger.warning("Missing headers in '%s' sheet (%s): %s", sheet_name, workbook_label, ", ".join(missing))

        for key, val in values.items():
            if key not in header_positions:
                continue
            row, col = header_positions[key]
            ws.cell(row=row + 1, column=col, value=val)

    @staticmethod
    def _fill_reference_document_sheet(wb, values: dict[str, str], workbook_label: str):
        ws = ProjectInitializer._find_sheet_case_insensitive(wb, "referenceDocument")
        if ws is None:
            logger.warning("Sheet 'referenceDocument' not found in %s", workbook_label)
            return

        normalized_values = {
            str(key).strip().lower(): "" if val is None else str(val)
            for key, val in values.items()
        }

        header_row = None
        for row in range(1, min(40, ws.max_row + 10)):
            for col in range(1, min(120, ws.max_column + 10)):
                value = ws.cell(row=row, column=col).value
                if value is None:
                    continue
                if str(value).strip().lower() == "documentsource":
                    header_row = row
                    break
            if header_row is not None:
                break

        if header_row is None:
            logger.warning("Header row for 'referenceDocument' not found in %s", workbook_label)
            return

        for col in range(1, min(160, ws.max_column + 20)):
            header = ws.cell(row=header_row, column=col).value
            if header is None:
                continue
            header_text = str(header).strip()
            if not header_text:
                continue

            target_cell = ws.cell(row=header_row + 1, column=col)
            header_norm = header_text.lower()
            if header_norm == "documentsource":
                target_cell.value = ""
            elif header_norm in normalized_values:
                target_cell.value = normalized_values.get(header_norm, "")
            else:
                target_cell.value = normalized_values.get("name (en)", "")

    def _fill_institution_characteristics_sheet(self, wb, workbook_label: str):
        source_path = self.regulation_source_file
        if not source_path.exists():
            logger.warning("Regulation source file not found: %s", source_path)
            return

        try:
            src_wb = load_workbook(source_path)
        except Exception as exc:
            logger.warning("Unable to open regulation source workbook (%s): %s", source_path, exc)
            return

        src_ws = self._find_sheet_case_insensitive(src_wb, "institutionCharacteristics")
        dst_ws = self._find_sheet_case_insensitive(wb, "institutionCharacteristics")
        if src_ws is None or dst_ws is None:
            logger.warning("Sheet 'institutionCharacteristics' missing in source or target workbook (%s)", workbook_label)
            return

        self._copy_worksheet_content(src_ws, dst_ws)

    def _fill_precision_sheet(self, wb, workbook_label: str):
        source_path = self.regulation_source_file
        if not source_path.exists():
            logger.warning("Regulation source file not found: %s", source_path)
            return

        try:
            src_wb = load_workbook(source_path)
        except Exception as exc:
            logger.warning("Unable to open regulation source workbook (%s): %s", source_path, exc)
            return

        src_ws = self._find_sheet_case_insensitive(src_wb, "precision")
        dst_ws = self._find_sheet_case_insensitive(wb, "precision")
        if src_ws is None or dst_ws is None:
            logger.warning("Sheet 'precision' missing in source or target workbook (%s)", workbook_label)
            return

        self._copy_worksheet_content(src_ws, dst_ws)

    def _clear_validationruleset_sheet(self, wb, workbook_label: str):
        ws = self._find_sheet_case_insensitive(wb, "validationRuleSet")
        if ws is None:
            logger.warning("Sheet 'validationRuleSet' not found in %s", workbook_label)
            return

        # Keep template headers/layout and clear only data rows.
        for row in range(2, ws.max_row + 1):
            for col in range(1, ws.max_column + 1):
                ws.cell(row=row, column=col).value = None

    def _fill_formgroup_sheet(self, wb, values: dict[str, str], workbook_label: str):
        ws = self._find_sheet_case_insensitive(wb, "formgroup")
        if ws is None:
            logger.warning("Sheet 'formgroup' not found in %s", workbook_label)
            return

        ref_values = self._read_reference_document_values(wb)
        if not ref_values:
            logger.warning("Unable to read referenceDocument values for formgroup fill in %s", workbook_label)

        required_headers = ["code", "name (EN)", "name (FR)", "description (EN)", "description (FR)"]
        header_positions: dict[str, tuple[int, int]] = {}

        for row in range(1, min(60, ws.max_row + 10)):
            for col in range(1, min(160, ws.max_column + 20)):
                value = ws.cell(row=row, column=col).value
                if value is None:
                    continue
                text = str(value).strip()
                for key in required_headers:
                    if text.lower() == key.lower() and key not in header_positions:
                        header_positions[key] = (row, col)

        missing = [k for k in required_headers if k not in header_positions]
        if missing:
            logger.warning("Missing headers in 'formgroup' sheet (%s): %s", workbook_label, ", ".join(missing))

        output_values = {
            "code": values.get("regulatorCode", ""),
            "name (EN)": ref_values.get("name (en)", ""),
            "name (FR)": ref_values.get("name (fr)", ""),
            "description (EN)": ref_values.get("description (en)", ""),
            "description (FR)": ref_values.get("description (fr)", ""),
        }

        for key, value in output_values.items():
            if key not in header_positions:
                continue
            row, col = header_positions[key]
            ws.cell(row=row + 1, column=col, value=value)

    @staticmethod
    def _fill_submission_fileset_sheet(wb, submission_format: str, container_type: str = "", workbook_label: str = ""):
        ws = ProjectInitializer._find_sheet_case_insensitive(wb, "submissionFileset")
        if ws is None:
            logger.warning("Sheet 'submissionFileset' not found in %s", workbook_label)
            return

        targets = {"submissionformat": submission_format, "containertype": container_type}
        header_positions: dict[str, tuple[int, int]] = {}

        for row in range(1, min(60, ws.max_row + 10)):
            for col in range(1, min(160, ws.max_column + 20)):
                value = ws.cell(row=row, column=col).value
                if value is None:
                    continue
                text = str(value).strip().lower()
                if text in targets and text not in header_positions:
                    header_positions[text] = (row, col)

        for key, val in targets.items():
            if not val or key not in header_positions:
                continue
            row, col = header_positions[key]
            ws.cell(row=row + 1, column=col, value=val)

    @staticmethod
    def _read_reference_document_values(wb) -> dict[str, str]:
        ws = ProjectInitializer._find_sheet_case_insensitive(wb, "referenceDocument")
        if ws is None:
            return {}

        header_row = None
        for row in range(1, min(40, ws.max_row + 10)):
            for col in range(1, min(120, ws.max_column + 10)):
                value = ws.cell(row=row, column=col).value
                if value is None:
                    continue
                text = str(value).strip().lower()
                if text == "documentsource":
                    header_row = row
                    break
            if header_row is not None:
                break

        if header_row is None:
            return {}

        collected: dict[str, str] = {}
        for col in range(1, min(160, ws.max_column + 20)):
            header = ws.cell(row=header_row, column=col).value
            if header is None:
                continue
            header_text = str(header).strip()
            if not header_text:
                continue

            value = ws.cell(row=header_row + 1, column=col).value
            collected[header_text.lower()] = "" if value is None else str(value)

        return collected

    @staticmethod
    def _find_sheet_case_insensitive(workbook, target_name: str):
        for name in workbook.sheetnames:
            if name.strip().lower() == target_name.lower():
                return workbook[name]
        return None

    @staticmethod
    def _copy_worksheet_content(src_ws, dst_ws):
        # Reset merged cells in destination then copy source merges.
        existing_merges = list(dst_ws.merged_cells.ranges)
        for rng in existing_merges:
            dst_ws.unmerge_cells(str(rng))

        # Clear existing content and styles in the currently used destination range.
        max_row = max(dst_ws.max_row, 1)
        max_col = max(dst_ws.max_column, 1)
        for row in dst_ws.iter_rows(min_row=1, max_row=max_row, min_col=1, max_col=max_col):
            for cell in row:
                cell.value = None
                cell._style = copy(cell._style)

        # Copy cell values and styles.
        for row in src_ws.iter_rows(min_row=1, max_row=src_ws.max_row, min_col=1, max_col=src_ws.max_column):
            for src_cell in row:
                dst_cell = dst_ws.cell(row=src_cell.row, column=src_cell.column)
                dst_cell.value = src_cell.value
                dst_cell._style = copy(src_cell._style)
                dst_cell.number_format = src_cell.number_format
                dst_cell.protection = copy(src_cell.protection)
                dst_cell.alignment = copy(src_cell.alignment)

        # Copy row/column dimensions.
        for key, dim in src_ws.column_dimensions.items():
            dst_dim = dst_ws.column_dimensions[key]
            dst_dim.width = dim.width
            dst_dim.hidden = dim.hidden

        for key, dim in src_ws.row_dimensions.items():
            dst_dim = dst_ws.row_dimensions[key]
            dst_dim.height = dim.height
            dst_dim.hidden = dim.hidden

        # Copy merged ranges.
        for merged_range in src_ws.merged_cells.ranges:
            dst_ws.merge_cells(str(merged_range))
