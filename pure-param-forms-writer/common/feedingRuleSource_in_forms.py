"""

### 🔧 **Prompt for feeding rules**

**Process Excel files to apply feeding rules from pure validation rules**

1. Create Python code to fill the feeding rules columns, rows and cells of a PURE forms from a pure-validationRules....xlsx file
   * Ask the user the path of the PURE form source workbook "pure-form....xlsm" (with file dialog) 
   * Ask the user the path of the PURE validationRules source workbook "pure-validationRules-xxx.xlsx" (with file dialog)

2. Identify the rootSheet
    Copies the pure-form file in the folder "generated_output" (will be created if not exist at the grand-parent level of the source workbook (../))
    Get the rootSheet name of this workbook
      The workbook contains multiple sheet : identification, cellSpecificities and labelTranslation that don't need to be taken into account
      The workbook contains at least one more tab, the rootSheet name can be inferred from them, such as
          eg. K_42.00 for K_42_00.a and K_42.00.b
          or eg. C_18.00 for C_18.00(0001)

3. Identify the applicable rules from "pure-validationRules-xxx.xlsx"
    Filter the rootSheet on the "T1" column. Be careful, this filter need to be done on the rootSheet name and possibly the "_" can be replaced by " "
    Filter on the type of formula, in the "Formula" column : we only need exact equation
      For eg., we need to remove all rows that have in their formula a ">" or "<" or a "if"
      
4. Pre-processed the formulas to get the complete description which can lead to more formula
    We need to combine the "formula" with the "rows", "columns" and "sheets" other columns of the "pure-validationRules-xxx.xlsx"
    We need also to take into account "pure-form....xlsm" to know the numbering of rows, columns and sheets.
    A complete description of the formula should include within "{}" the complete descriptions of cells it impacts : 
        {tab, rXXXX, cXXXX} where : tab is the concatenation of "T1" and "sheets" column
                                    rXXXX is the row number either directly provided or inferred from the "rows" column
                                    cXXXX is the column number either directly provided or inferred from the "columns" column

    Sometimes, there is already the complete description of the cells such as {C 18.00(0001), r0010, c0010} and nothing need to be added as nothing appear in the other columns 'rows/columns/sheets'.
    Sometimes, there is no complete description and we need to build it from the other columns.
    For eg : 
       T1   | rows | columns | sheets | Formula
	C18.00  |      |         |  (002) | {r0011, c0060} = {r0020, c0060}               
    ==> complete description : {C18.00(002), r0011, c0060} = {C18.00(002), r0020, c0060} 

    Other eg : 
       T1   | rows   | columns | sheets | Formula
	C18.00  | (0010) |         |  (002) | {c0060} = {c0070}  
    ==> complete description : {C18.00(002), r0010, c0060} = {C18.00(002), r0010, c0070} 

    Columns : "rows/columns/sheets" can include also ";" or "-" that separate multiple entries.
    In that case, ";" need to be interpreted as a logical AND, while "-" should be interpreted as a logical suite.
    For eg., 
       T1   | rows        | columns | sheets | Formula
	C18.00  | (0010;0011) |         |  (002) | {c0060} = {c0070}  
    ==> lead to exactly 2 rules with complete description : 
        {C18.00(002), r0010, c0060} = {C18.00(002), r0010, c0070} 
        {C18.00(002), r0011, c0060} = {C18.00(002), r0011, c0070}

    The "-" symbol is the most complex to infer, as we need to have a look into the "pure-form....xlsm" sheets.
    More precisely, if the "-" symbol appears to define rows :
        we should find all the rows, starting from the first row mentioned before the "-" symbol and the last rows mentioned after the "-" symbol and that are located under the "rowNumbering" cell of the "L" column in "pure-form....xlsm" for the sheet mentionned.
    Else, if the "-" symbol appears to define columns :
        we should find all the columns, starting from the first column mentioned before the "-" symbol and the last columns mentioned after the "-" symbol and that are located on the right of the "rowLabelColumnNumbering" cell of the "M" column in "pure-form....xlsm" for the sheet mentionned.

    Else if the "-" symbol appears to define sheets :
        we should find all the sheets, starting from the first sheet mentioned before the "-" symbol and the last sheets mentioned after the "-" symbol in "pure-form....xlsm".

    For eg., 
       T1   | rows        | columns | sheets | Formula
	C18.00  | (0010-0030) |         |  (002) | {c0060} = {c0070}  
    ==> lead to X rules with complete description, where X here will be the number of rows in the "pure-form....xlsm" sheet labelled C_18.00(0002) that match the rows from 0010 to 0030
        let's say here there are the r0010, r0020 and r0030, we will have thus 3 rules :
        {C18.00(002), r0010, c0060} = {C18.00(002), r0010, c0070} 
        {C18.00(002), r0020, c0060} = {C18.00(002), r0020, c0070}
        {C18.00(002), r0030, c0060} = {C18.00(002), r0030, c0070}

    Also, we need to map the "sum" operator to the "+" operator in Excel.
    For eg., 
       T1   | rows        | columns | sheets | Formula
	C35.02  | (0010-0020) |         |        | {c0110} = sum({C 35.02, (c0030-0100)})
    ==> lead to X rules with complete description, where X here will be the number of rows in the "pure-form....xlsm" sheet labelled C_35.02(0001) that match the rows from 0010 to 0020 
        and the formula will be : 
         {C35.02, r0010, c0110} =  {C35.02, r0010, c0030} + {C35.02, r0010, c0040} + ... + {C35.02, r0010, c0100}
         where the columns from c0030 to c0100 are all the columns found on the right of the "rowLabelColumnNumbering" cell of the "M" column in "pure-form....xlsm" for the sheet C_35.02
    

5. Basic position of the rule
    For each applicable complete description of rule: 
        Identify the target position cell of the equation rule : it should be the part of the equation that is unitary
            For eg., in the formula "{r0010, c0060} * 12.5 = {r0010, c0070}", the right part here of the equation "{r0010, c0070}" is the target position cell.
            For eg., in the formula "{r0010, c0060} = {r0010, c0070} + {r0010, c0080}", the left part here of the equation "{r0010, c0060}" is the target position cell.
            For eg., in the formula  "{r0010, c0070} =  {r0010, c0060}", the left part here of the equation "{r0010, c0070}" is the target position cell (always left by default if no other indications).

        Copy only the part of the equation that does not represent the coordinate in the correct cell in the "pure-form....xlsm" by looking at the coordinates of the target position cell.
            For eg., in the formula "{r0010, c0060} * 12.5 = {r0010, c0070}", "= {r0010, c0060} * 12.5" should be written in the correct cell.
            For eg., in the formula "{r0010, c0060} = {r0010, c0070} + {r0010, c0080}", "= {r0010, c0070} + {r0010, c0080}" should be written in the correct cell.

        Change the color of the cell filled, to a pink color (Red : 255, Green 153; Blue 204 / #FF99CC)
        The formula should start with an apostrophe "'" to make sure Excel can still open the file without any security warning.
        
        NB : if 2 rules are present in the same cell, a Warning log should be generated to inform the user. We should not concatenate the 2 rules but only keep the first one.


6. Postprocessing in the effective valued tables :
        Find the effective valued tables :
            from cell with coordinate (row = rowLabelColumnNumbering cell + 1 ; column = rowLabelColumnNumbering cell + 1 ) 
            till end of table finding a large border (both on right and bottom)to simplify the formula

        a. Postprocessing to simplify the formula
           For each cell formula, remove the unnecessary information.
           For eg., if in the cell with the coordinate  {C_18.00(0001), r0020, c0020}, we have the formula :
               '= {C_18.00(0001), r0030, c0020} + {C_18.00(0001), r0080, c0020} + {C_18.00(0001), r0120, c0020}
                we can simplify it to:
                '= {r0030} + {r0080} + {r0120}
                since the column information is identical here to the coordinate as well as the sheet name.

        b. Postprocessing to mutualize majority of rules on same column and respectively on same row
            Parse each rows found (but take into account the "pink" cells FF99CC) :
                - If one of the formula is contained in more than 50% of pink cells of the row and more than 2 cells
                    ==> copy it in the column "G" known as feedingRuleSource column and remove this formula from the other cells

            And do the same for columns, by centralized identical formulas within the row "10" known as feedingRuleSource row.

            Exception: Don't mutualize formulas containing the operator "/"

        c. Postprocessing to keep in cells (not in feedingRuleSource centralized formula) the complete coordinates of the formula with rows and columns
            For eg., if in the cell with the coordinate  {C_18.00(0001), r0020, c0020}, we have the formula :
                '= {r0030} + {r0080} + {r0120}
                we need to complete it such as :
                '= {r0030,c0020} + {r0080,c0020} + {r0120,c0020}


"""

# Standard library imports
import os
import sys
import logging
import argparse
import shutil
import re
from pathlib import Path

# Third-party imports
import openpyxl
from openpyxl.styles import PatternFill, Font, Border, Side
from openpyxl.utils import get_column_letter, column_index_from_string
from tkinter import Tk, filedialog, messagebox


# Global debug mode variable
DEBUG_MODE = False

# Global cache for form structure markers per worksheet
FORM_STRUCTURE_CACHE = {}


def configure_logging(debug_mode=False, preserve_handlers=False):
    """Configure logging based on debug mode
    
    Args:
        debug_mode: Enable debug level logging
        preserve_handlers: If True, don't clear existing handlers (useful when called from another script)
    """
    global DEBUG_MODE
    DEBUG_MODE = debug_mode
    level = logging.DEBUG if DEBUG_MODE else logging.INFO
    
    # Clear any existing handlers only if not preserving them
    if not preserve_handlers:
        for handler in logging.root.handlers[:]:
            logging.root.removeHandler(handler)
    
        logging.basicConfig(
            level=level,
            format='[%(levelname)s] %(message)s',
            handlers=[logging.StreamHandler()]
        )
    else:
        # When preserving handlers, always set root logger to DEBUG
        # Individual handlers will filter based on their own level settings
        logging.root.setLevel(logging.DEBUG)


def get_validation_rules_file():
    """
    Get the validation rules file from user input (optional step)
    Returns: Path to the validation rules Excel file, or None if user cancels
    """
    root = Tk()
    root.withdraw()  # Hide the main window

    # Select validation rules file
    validation_file = filedialog.askopenfilename(
        title="[OPTIONAL] Select pure-validationRules Excel file (or Cancel to skip)",
        filetypes=[("Excel files", "*.xlsx *.xlsm")],
        initialdir=os.getcwd()
    )
    
    # If user cancels, return None silently (no error message)
    if not validation_file:
        return None
    
    # Validate file name pattern
    filename = os.path.basename(validation_file)
    if not filename.startswith("pure-validationRules-"):
        messagebox.showwarning("Warning", 
                             f"File name '{filename}' doesn't match expected pattern 'pure-validationRules-xxx.xlsx'")
    
    return validation_file


def get_pure_form_file():
    """
    Get the pure-form file from user input
    Returns: Path to the pure-form Excel file
    """
    root = Tk()
    root.withdraw()  # Hide the main window

    # Select pure-form file
    form_file = filedialog.askopenfilename(
        title="Select pure-form Excel file to modify",
        filetypes=[("Excel files", "*.xlsx *.xlsm")],
        initialdir=os.getcwd()
    )
    
    if not form_file:
        messagebox.showerror("Error", "No pure-form file selected. Exiting.")
        return None
    
    # Validate file name pattern
    filename = os.path.basename(form_file)
    if not filename.startswith("pure-form"):
        messagebox.showwarning("Warning", 
                             f"File name '{filename}' doesn't match expected pattern 'pure-form....xlsm'")
    
    return form_file


def create_output_directory(source_path):
    """
    Create generated_output directory at grand-parent level of source workbook
    Args:
        source_path: Path to the source file
    Returns:
        Path to the created output directory
    """
    # Get the directory containing the source file
    source_dir = os.path.dirname(source_path)
    # Go up to the grand-parent level of the source workbook
    grandparent_dir = os.path.dirname(os.path.dirname(source_path))
    output_dir = os.path.join(grandparent_dir, "generated_output")
    
    os.makedirs(output_dir, exist_ok=True)
    logging.debug(f"   Created/verified output directory: {output_dir}")
    
    return output_dir


def copy_form_to_output(source_path, output_dir):
    """
    Copy the pure-form file to the output directory
    Args:
        source_path: Path to the source pure-form file
        output_dir: Path to the output directory
    Returns:
        Path to the copied file
    """
    source_name = os.path.basename(source_path)
    
    # Keep the original filename without modification
    modified_path = os.path.join(output_dir, source_name)
    
    try:
        shutil.copy2(source_path, modified_path)
        logging.info(f"   File copied to: {modified_path}")
        return modified_path
    except Exception as e:
        logging.error(f"Error copying file: {str(e)}")
        return None


def identify_root_sheet(workbook):
    """
    Identify the root sheet name from the workbook
    Args:
        workbook: openpyxl workbook object
    Returns:
        Root sheet name
    """
    ignore_sheets = ["identification", "cellSpecificities", "labelTranslation"]
    
    # Get all sheet names except ignored ones
    data_sheets = [sheet for sheet in workbook.sheetnames if sheet not in ignore_sheets]
    
    if not data_sheets:
        logging.warning("No data sheets found")
        return None
    
    # For sheets like K_42.00.a, K_42.00.b -> root is K_42.00
    # For sheets like C_18.00(0001) -> root is C_18.00
    # For sheets like C_35.02(0001) -> root is C_35.02
    first_sheet = data_sheets[0]
    
    # Remove suffixes like .a, .b (single lowercase letter after dot)
    # or (0001), (0002) etc. but keep version numbers like .02, .00
    root_name = re.sub(r'(\.[a-z]|\(\d+\)).*$', '', first_sheet)
    
    logging.debug(f"   Identified root sheet: {root_name}")
    return root_name


def load_validation_rules(validation_file_path, root_sheet):
    """
    Load and filter validation rules from Excel file
    Args:
        validation_file_path: Path to validation rules file
        root_sheet: Root sheet name to filter on
    Returns:
        List of applicable validation rules
    """
    logging.info(f"🔧 Loading validation rules from: {validation_file_path}")
    
    try:
        wb = openpyxl.load_workbook(validation_file_path, data_only=True)
        
        # Look for the main rules sheet (usually first sheet or one with data)
        rules_sheet = None
        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            # Check if sheet has headers like T1, Formula, etc.
            if ws['A1'].value and 'T1' in str(ws['A1'].value):
                rules_sheet = ws
                break
        
        if not rules_sheet:
            # Try first sheet
            rules_sheet = wb.active
        
        logging.debug(f"   Using sheet: {rules_sheet.title}")
        
        # Find column indices
        headers = {}
        for col in range(1, 30):  # Check first 30 columns
            cell_value = rules_sheet.cell(row=1, column=col).value
            if cell_value:
                headers[str(cell_value).strip()] = col
        
        #logging.debug(f"   Found headers: {list(headers.keys())}")
        
        if 'T1' not in headers or 'Formula' not in headers:
            logging.error("Required columns T1 and Formula not found")
            return []
        
        # Load rules
        applicable_rules = []
        for row in range(2, rules_sheet.max_row + 1):
            t1_value = rules_sheet.cell(row=row, column=headers['T1']).value
            formula_value = rules_sheet.cell(row=row, column=headers['Formula']).value
            
            if not t1_value or not formula_value:
                continue
            
            # Filter by root sheet (handle underscore vs space and dots)
            t1_normalized = str(t1_value).replace('_', '').replace(' ', '').replace('.', '')
            root_normalized = root_sheet.replace('_', '').replace(' ', '').replace('.', '')
            
            if not t1_normalized.lower().startswith(root_normalized.lower()):
                continue
            
            # Filter out non-equation formulas
            formula_str = str(formula_value)
            if any(op in formula_str for op in ['>', '<', 'if', 'IF']):
                continue
            
            # Filter out formulas ending with "= empty"
            if formula_str.strip().lower().endswith('= empty'):
                continue
            
            # Filter out formulas containing brackets [] (e.g., [CEG=eba_GA:x1])
            if '[' in formula_str or ']' in formula_str:
                logging.debug(f"   ⛔️ Square brackets rule filtered out: {formula_str}")
                continue
            
            rule = {
                'T1': t1_value,
                'Formula': formula_value,
                'rows': rules_sheet.cell(row=row, column=headers.get('rows', 0)).value if 'rows' in headers else None,
                'columns': rules_sheet.cell(row=row, column=headers.get('columns', 0)).value if 'columns' in headers else None,
                'sheets': rules_sheet.cell(row=row, column=headers.get('sheets', 0)).value if 'sheets' in headers else None
            }
            
            applicable_rules.append(rule)
            #logging.debug(f"   Added rule: {rule}")
        
        wb.close()
        logging.info(f"   ✅ Loaded {len(applicable_rules)} applicable rules")
        return applicable_rules
        
    except Exception as e:
        logging.error(f"Error loading validation rules: {str(e)}")
        return []


def expand_range_values(value_str, form_wb, sheet_name, value_type='sheet'):
    """
    Expand range values like (0010-0030) or (0010;0011) or (All) using actual form structure
    Args:
        value_str: String containing ranges/lists
        form_wb: Form workbook to check actual rows/columns
        sheet_name: Sheet name to check
        value_type: 'sheet', 'row', or 'column'
    Returns:
        List of expanded values
    """
    if not value_str:
        return []
    
    # Remove parentheses
    clean_str = str(value_str).strip('()')
    
    # Handle "(All)" - get all actual rows/columns/sheets from the form
    if clean_str.lower() == 'all':
        if value_type == 'sheet':
            # Get all sheet suffixes from the form
            ignore_sheets = ["identification", "cellSpecificities", "labelTranslation"]
            all_sheets = [sheet for sheet in form_wb.sheetnames if sheet not in ignore_sheets]
            sheet_suffixes = []
            for sheet in all_sheets:
                match = re.search(r'\((\d+)\)', sheet)
                if match:
                    sheet_suffixes.append(f"({match.group(1)})")
            return sheet_suffixes if sheet_suffixes else ['']
        elif value_type == 'row':
            return get_actual_rows_from_form(form_wb, sheet_name)
        else:  # column
            return get_actual_columns_from_form(form_wb, sheet_name)
    
    # Handle semicolon separation (AND logic)
    if ';' in clean_str:
        parts = clean_str.split(';')
        result = []
        for part in parts:
            result.extend(expand_range_values(f"({part})", form_wb, sheet_name, value_type))
        return result
    
    # Handle dash separation (range logic)
    if '-' in clean_str:
        parts = clean_str.split('-')
        if len(parts) == 2:
            start_val, end_val = parts[0].strip(), parts[1].strip()
            
            if value_type == 'sheet':
                # For sheets, check actual sheet numbering in form
                try:
                    start_num = int(start_val)
                    end_num = int(end_val)
                    # Get actual sheets that fall in this range
                    ignore_sheets = ["identification", "cellSpecificities", "labelTranslation"]
                    all_sheets = [sheet for sheet in form_wb.sheetnames if sheet not in ignore_sheets]
                    sheet_numbers = []
                    for sheet in all_sheets:
                        match = re.search(r'\((\d+)\)', sheet)
                        if match:
                            sheet_num = int(match.group(1))
                            if start_num <= sheet_num <= end_num:
                                sheet_numbers.append(f"({match.group(1)})")
                    return sheet_numbers if sheet_numbers else [f"({start_val})", f"({end_val})"]
                except:
                    return [f"({start_val})", f"({end_val})"]
            elif value_type == 'row':
                # Use actual form structure to find rows in range
                return get_rows_in_range(form_wb, sheet_name, start_val, end_val)
            else:  # column
                # Use actual form structure to find columns in range
                return get_columns_in_range(form_wb, sheet_name, start_val, end_val)
    
    # Single value
    return [clean_str]


def get_actual_rows_from_form(form_wb, sheet_name):
    """Get all actual row codes from the form structure"""
    try:
        # Find the target worksheet
        target_ws = None
        for ws_name in form_wb.sheetnames:
            sheet_base = re.sub(r'[.(].*$', '', ws_name)
            target_base = re.sub(r'[.(].*$', '', sheet_name)
            if (sheet_base.replace('_', '').replace(' ', '').lower() == 
                target_base.replace('_', '').replace(' ', '').lower()):
                target_ws = form_wb[ws_name]
                break
        
        if not target_ws:
            return [f"{i:04d}" for i in range(10, 400, 10)]  # Fallback
        
        # Find rowNumbering marker
        markers = find_form_structure_markers(target_ws)
        if not markers['rowNumbering_pos']:
            return [f"{i:04d}" for i in range(10, 400, 10)]  # Fallback
        
        row_marker_row, row_marker_col = markers['rowNumbering_pos']
        row_codes = []
        
        # Extract all row codes below the marker
        for search_row in range(row_marker_row + 1, min(row_marker_row + 200, target_ws.max_row + 1)):
            cell_value = target_ws.cell(row=search_row, column=row_marker_col).value
            if cell_value:
                # Look for 4-digit row codes
                matches = re.findall(r'\b(\d{4})\b', str(cell_value))
                for match in matches:
                    if match not in row_codes:
                        row_codes.append(match)
        
        return row_codes if row_codes else [f"{i:04d}" for i in range(10, 400, 10)]
        
    except Exception as e:
        logging.warning(f"Error getting rows from form: {str(e)}")
        return [f"{i:04d}" for i in range(10, 400, 10)]


def get_actual_columns_from_form(form_wb, sheet_name):
    """Get all actual column codes from the form structure"""
    try:
        # Find the target worksheet
        target_ws = None
        for ws_name in form_wb.sheetnames:
            sheet_base = re.sub(r'[.(].*$', '', ws_name)
            target_base = re.sub(r'[.(].*$', '', sheet_name)
            if (sheet_base.replace('_', '').replace(' ', '').lower() == 
                target_base.replace('_', '').replace(' ', '').lower()):
                target_ws = form_wb[ws_name]
                break
        
        if not target_ws:
            return [f"{i:04d}" for i in range(10, 100, 10)]  # Fallback
        
        # Find rowLabelColumnNumbering marker
        markers = find_form_structure_markers(target_ws)
        if not markers['rowLabelColumnNumbering_pos']:
            return [f"{i:04d}" for i in range(10, 100, 10)]  # Fallback
        
        col_marker_row, col_marker_col = markers['rowLabelColumnNumbering_pos']
        col_codes = []
        
        # Extract all column codes to the right of the marker
        for search_col in range(col_marker_col + 1, min(col_marker_col + 100, target_ws.max_column + 1)):
            cell_value = target_ws.cell(row=col_marker_row, column=search_col).value
            if cell_value:
                # Look for 4-digit column codes
                matches = re.findall(r'\b(\d{4})\b', str(cell_value))
                for match in matches:
                    if match not in col_codes:
                        col_codes.append(match)
        
        return col_codes if col_codes else [f"{i:04d}" for i in range(10, 100, 10)]
        
    except Exception as e:
        logging.warning(f"Error getting columns from form: {str(e)}")
        return [f"{i:04d}" for i in range(10, 100, 10)]


def get_rows_in_range(form_wb, sheet_name, start_val, end_val):
    """Get actual row codes in the specified range from form structure"""
    try:
        all_rows = get_actual_rows_from_form(form_wb, sheet_name)
        start_num = int(start_val)
        end_num = int(end_val)
        
        # Filter rows that fall within the range
        range_rows = []
        for row_code in all_rows:
            row_num = int(row_code)
            if start_num <= row_num <= end_num:
                range_rows.append(row_code)
        
        return range_rows if range_rows else [start_val, end_val]
        
    except Exception as e:
        logging.warning(f"Error getting rows in range: {str(e)}")
        return [start_val, end_val]


def get_columns_in_range(form_wb, sheet_name, start_val, end_val):
    """Get actual column codes in the specified range from form structure"""
    try:
        all_columns = get_actual_columns_from_form(form_wb, sheet_name)
        start_num = int(start_val)
        end_num = int(end_val)
        
        # Filter columns that fall within the range
        range_columns = []
        for col_code in all_columns:
            col_num = int(col_code)
            if start_num <= col_num <= end_num:
                range_columns.append(col_code)
        
        return range_columns if range_columns else [start_val, end_val]
        
    except Exception as e:
        logging.warning(f"Error getting columns in range: {str(e)}")
        return [start_val, end_val]


def build_complete_formulas(rules, form_wb, root_sheet):
    """
    Build complete formula descriptions from validation rules
    Args:
        rules: List of validation rules
        form_wb: Form workbook
        root_sheet: Root sheet name
    Returns:
        List of complete formulas
    """
    complete_formulas = []
    filtered_inter_table_count = 0
    
    def extract_sheets_from_formula(formula):
        """Extract all unique sheet names from a formula"""
        refs = re.findall(r'\{([^}]+)\}', formula)
        sheets = set()
        for ref in refs:
            parts = [p.strip() for p in ref.split(',')]
            for part in parts:
                # Sheet name = part that doesn't start with 'r' or 'c' or '('
                # Ignore ranges like (c0010-0100) or (r0010-0050)
                if not part.startswith('r') and not part.startswith('c') and not part.startswith('('):
                    sheets.add(part)
        return sheets

    for rule in rules:
        logging.debug(f"   From original rule '{rule}'  ")  
        try:
            formula = rule['Formula']
            t1 = rule['T1']
            rows_val = rule['rows']
            cols_val = rule['columns'] 
            sheets_val = rule['sheets']

            # Expand ranges
            row_values = expand_range_values(rows_val, form_wb, root_sheet, 'row') if rows_val else ['']
            col_values = expand_range_values(cols_val, form_wb, root_sheet, 'column') if cols_val else ['']
            sheet_values = expand_range_values(sheets_val, form_wb, root_sheet, 'sheet') if sheets_val else ['']

            # If no expansions, use single empty value
            if not row_values:
                row_values = ['']
            if not col_values:
                col_values = ['']
            if not sheet_values:
                sheet_values = ['']

            # Generate all combinations
            for sheet_val in sheet_values:
                for row_val in row_values:
                    for col_val in col_values:
                        # Build sheet name
                        if sheet_val and sheet_val != "(All)":
                            full_sheet = f"{t1.replace(' ', '_')}{sheet_val}"
                        else:
                            full_sheet = t1.replace(' ', '_')

                        # Process formula to add missing components
                        processed_formula = process_formula(formula, full_sheet, row_val, col_val)
                        
                        # Expand sum() operator to individual additions
                        processed_formula = expand_sum_operator(processed_formula, form_wb)

                        if processed_formula:
                            # Filter inter-table rules: exclude if multiple different sheets in the formula
                            sheets_in_formula = extract_sheets_from_formula(processed_formula)
                            if len(sheets_in_formula) > 1:
                                filtered_inter_table_count += 1
                                logging.debug(f"      ⛔️ Inter-table rule filtered out: {processed_formula}")
                                continue
                            complete_formulas.append({
                                'formula': processed_formula,
                                'sheet': full_sheet,
                                'original_rule': rule
                            })
                            logging.debug(f"    ---> Generated formula: {processed_formula}")
        except Exception as e:
            logging.warning(f"Error processing rule {rule}: {str(e)}")
            continue
    
    if filtered_inter_table_count > 0:
        logging.info(f"   🔍 Filtered out {filtered_inter_table_count} inter-table rules (referencing multiple sheets)")
    
    return complete_formulas


def process_formula(formula, sheet_name, row_val, col_val):
    """
    Process a formula to add missing sheet, row, column information
    Args:
        formula: Original formula string
        sheet_name: Sheet name to add
        row_val: Row value to add
        col_val: Column value to add
    Returns:
        Processed formula string
    """
    # If formula already has complete descriptions, return as-is
    if sheet_name in formula and 'r' in formula and 'c' in formula:
        return formula
    
    # Pattern to find cell references like {r0010, c0060} or {c0060}
    pattern = r'\{([^}]+)\}'
    
    def replace_cell_ref(match):
        content = match.group(1)
        parts = [p.strip() for p in content.split(',')]
        
        # Separate parts into sheet, row, column
        existing_sheet = None
        existing_rows = []
        existing_cols = []
        
        for part in parts:
            if part.startswith('r'):
                existing_rows.append(part)
            elif part.startswith('c'):
                existing_cols.append(part)
            else:
                # This is a sheet name
                existing_sheet = part.replace(' ', '_')  # Normalize immediately
        
        # Build new parts in correct order: sheet, rows, columns
        new_parts = []
        
        # Add sheet: use existing sheet if present, otherwise use sheet_name
        if existing_sheet:
            new_parts.append(existing_sheet)
        else:
            new_parts.append(sheet_name)
        
        # Add existing rows
        if existing_rows:
            new_parts.extend(existing_rows)
        # Add row if missing and available
        elif row_val:
            new_parts.append(f"r{row_val}")
        
        # Add existing columns
        if existing_cols:
            new_parts.extend(existing_cols)
        # Add column if missing and available  
        elif col_val:
            new_parts.append(f"c{col_val}")
        
        return '{' + ', '.join(new_parts) + '}'
    
    processed = re.sub(pattern, replace_cell_ref, formula)
    
    # Final pass: normalize all sheet names in the formula (replace spaces with underscores)
    # This handles sheet names that were already in the original formula
    def normalize_sheet_in_cell_ref(match):
        content = match.group(1)
        parts = [p.strip() for p in content.split(',')]
        normalized_parts = []
        for part in parts:
            if not part.startswith('r') and not part.startswith('c'):
                # This is a sheet name, normalize it
                normalized_parts.append(part.replace(' ', '_'))
            else:
                normalized_parts.append(part)
        return '{' + ', '.join(normalized_parts) + '}'
    
    return re.sub(pattern, normalize_sheet_in_cell_ref, processed)


def expand_sum_operator(formula, form_wb):
    """
    Expand sum() operator to individual additions with actual columns from form structure
    Args:
        formula: Formula string potentially containing sum()
        form_wb: Form workbook to get actual column structure
    Returns:
        Expanded formula string
    """
    # Check if formula contains sum()
    if 'sum(' not in formula.lower():
        return formula
    
    try:
        # First, extract sheet name from the entire formula (it might be outside the sum())
        # Look for sheet names in cell references like {C_35.02, r0010, c0110}
        all_refs = re.findall(r'\{([^}]+)\}', formula)
        formula_sheet_name = None
        for ref in all_refs:
            parts = [p.strip() for p in ref.split(',')]
            for part in parts:
                if not part.startswith('r') and not part.startswith('c') and not part.startswith('('):
                    formula_sheet_name = part.replace(' ', '_')
                    break
            if formula_sheet_name:
                break
        
        # Pattern to find sum({...})
        sum_pattern = r'sum\(\s*\{([^}]+)\}\s*\)'
        
        def expand_sum(match):
            content = match.group(1)
            parts = [p.strip() for p in content.split(',')]
            
            # Extract sheet, row, column, and ranges from inside sum()
            sheet_name = None
            row_code = None
            col_code = None
            row_range = None
            col_range = None
            
            for part in parts:
                if part.startswith('r') and not part.startswith('('):
                    # Single row like r0010
                    row_code = part
                elif part.startswith('c') and not part.startswith('('):
                    # Single column like c0040
                    col_code = part
                elif part.startswith('(') and 'r' in part:
                    # This is a row range like (r0020-0040)
                    row_range = part.strip('()')
                elif part.startswith('(') and 'c' in part:
                    # This is a column range like (c0030-0100)
                    col_range = part.strip('()')
                elif not part.startswith('c') and not part.startswith('r') and not part.startswith('('):
                    # This is a sheet name
                    sheet_name = part.replace(' ', '_')
            
            # If no sheet name found inside sum(), use the one from the formula
            if not sheet_name and formula_sheet_name:
                sheet_name = formula_sheet_name
            
            # Check if we have either a row range or a column range
            if not row_range and not col_range:
                return match.group(0)
            
            if not sheet_name:
                logging.warning(f"No sheet name found in sum() expansion for: {match.group(0)}")
                return match.group(0)
            
            cell_refs = []
            
            # Case 1: Column range sum - sum over columns
            if col_range and '-' in col_range:
                # Parse column range
                col_parts = col_range.replace('c', '').split('-')
                if len(col_parts) != 2:
                    return match.group(0)
                
                start_col = col_parts[0].strip()
                end_col = col_parts[1].strip()
                
                # Get actual columns from form structure
                actual_columns = get_columns_in_range(form_wb, sheet_name, start_col, end_col)
                
                # Build individual cell references
                for col_code_val in actual_columns:
                    cell_parts = []
                    if sheet_name:
                        cell_parts.append(sheet_name)
                    if row_code:
                        cell_parts.append(row_code)
                    cell_parts.append(f"c{col_code_val}")
                    
                    cell_refs.append('{' + ', '.join(cell_parts) + '}')
            
            # Case 2: Row range sum - sum over rows
            elif row_range and '-' in row_range:
                # Parse row range
                row_parts = row_range.replace('r', '').split('-')
                if len(row_parts) != 2:
                    return match.group(0)
                
                start_row = row_parts[0].strip()
                end_row = row_parts[1].strip()
                
                # Get actual rows from form structure
                actual_rows = get_rows_in_range(form_wb, sheet_name, start_row, end_row)
                
                # Build individual cell references
                for row_code_val in actual_rows:
                    cell_parts = []
                    if sheet_name:
                        cell_parts.append(sheet_name)
                    cell_parts.append(f"r{row_code_val}")
                    if col_code:
                        cell_parts.append(col_code)
                    
                    cell_refs.append('{' + ', '.join(cell_parts) + '}')
            
            # Join with + operator
            if cell_refs:
                return ' + '.join(cell_refs)
            else:
                return match.group(0)
        
        # Apply expansion
        expanded = re.sub(sum_pattern, expand_sum, formula, flags=re.IGNORECASE)
        
        if expanded != formula:
            logging.debug(f"      ✨ Expanded sum() operator")
            #logging.debug(f"         From: {formula}")
            #logging.debug(f"         To:   {expanded[:500]}{'...' if len(expanded) > 500 else ''}")
        
        return expanded
        
    except Exception as e:
        logging.warning(f"Error expanding sum() operator: {str(e)}")
        return formula


def find_target_cell(formula):
    """
    Find the target cell (unitary part) from a formula according to prompt specification
    Args:
        formula: Complete formula string
    Returns:
        Target cell reference or None
    """
    # Split on = to get left and right sides
    parts = formula.split('=', 1)
    if len(parts) != 2:
        return None
    
    left_side = parts[0].strip()
    right_side = parts[1].strip()
    
    # Find cell references on each side
    left_refs = re.findall(r'\{[^}]+\}', left_side)
    right_refs = re.findall(r'\{[^}]+\}', right_side)
    
    # Check which side is "unitary" (contains only one cell reference without operators)
    def is_unitary_side(side, refs):
        """Check if a side contains only one cell reference without operators"""
        if len(refs) != 1:
            return False
        
        # Remove the cell reference and check if anything significant remains
        side_without_ref = side
        for ref in refs:
            side_without_ref = side_without_ref.replace(ref, '').strip()
        
        # If only whitespace remains, it's unitary
        return len(side_without_ref) == 0
    
    left_is_unitary = is_unitary_side(left_side, left_refs)
    right_is_unitary = is_unitary_side(right_side, right_refs)
    
    # Apply the logic from the prompt:
    # - If right side is unitary and left has operators/multiple refs -> target is right
    # - If left side is unitary and right has operators/multiple refs -> target is left  
    # - If both or neither are unitary -> default to left
    
    if right_is_unitary and not left_is_unitary:
        #logging.debug(f"      Target is RIGHT (unitary): {right_refs[0]}")
        return right_refs[0]
    elif left_is_unitary and not right_is_unitary:
        #logging.debug(f"      Target is LEFT (unitary): {left_refs[0]}")
        return left_refs[0]
    else:
        # Default to left side as specified in prompt
        target = left_refs[0] if left_refs else (right_refs[0] if right_refs else None)
        #logging.debug(f"      Target is LEFT (default): {target}")
        return target


def parse_cell_reference(cell_ref):
    """
    Parse a cell reference like {C18.00(002), r0010, c0060}
    Args:
        cell_ref: Cell reference string
    Returns:
        Dictionary with sheet, row, column
    """
    # Remove braces
    content = cell_ref.strip('{}')
    parts = [p.strip() for p in content.split(',')]
    
    result = {'sheet': None, 'row': None, 'column': None}
    
    for part in parts:
        if part.startswith('r'):
            result['row'] = part[1:]  # Remove 'r' prefix
        elif part.startswith('c'):
            result['column'] = part[1:]  # Remove 'c' prefix
        else:
            result['sheet'] = part
    
    return result


def find_form_structure_markers(worksheet):
    """
    Find the rowNumbering and rowLabelColumnNumbering markers in the form with caching
    Args:
        worksheet: openpyxl worksheet
    Returns:
        Dictionary with marker positions
    """
    global FORM_STRUCTURE_CACHE
    
    # Create a unique key for this worksheet
    cache_key = f"{worksheet.parent.path}#{worksheet.title}" if hasattr(worksheet.parent, 'path') else f"workbook#{worksheet.title}"
    
    # Check if already cached
    if cache_key in FORM_STRUCTURE_CACHE:
        #logging.debug(f"   Using cached markers for {worksheet.title}")
        return FORM_STRUCTURE_CACHE[cache_key]
    
    markers = {
        'rowNumbering_pos': None,
        'rowLabelColumnNumbering_pos': None
    }
    
    try:
        # Search for "rowNumbering" in column L (column 12)
        for row in range(1, min(100, worksheet.max_row + 1)):  # Search first 100 rows
            cell_value = worksheet.cell(row=row, column=12).value  # Column L
            if cell_value and 'rowNumbering' in str(cell_value):
                markers['rowNumbering_pos'] = (row, 12)
                logging.debug(f"      Found rowNumbering at L{row} in {worksheet.title}")
                break
        
        # Search for "rowLabelColumnNumbering" in column M (column 13)
        for row in range(1, min(100, worksheet.max_row + 1)):  # Search first 100 rows
            cell_value = worksheet.cell(row=row, column=13).value  # Column M
            if cell_value and 'rowLabelColumnNumbering' in str(cell_value):
                markers['rowLabelColumnNumbering_pos'] = (row, 13)
                logging.debug(f"      Found rowLabelColumnNumbering at M{row} in {worksheet.title}")
                break
                
    except Exception as e:
        logging.warning(f"      Error searching for form markers in {worksheet.title}: {str(e)}")
    
    # Cache the result
    FORM_STRUCTURE_CACHE[cache_key] = markers
    logging.debug(f"      Cached markers for {worksheet.title}")
    
    return markers


def convert_to_excel_coords(row_str, col_str, form_wb, sheet_name):
    """
    Convert r0010, c0060 format to actual Excel row/column numbers using form structure
    Args:
        row_str: Row string like "0010"
        col_str: Column string like "0060"
        form_wb: Form workbook
        sheet_name: Sheet name with specific sheet identifier (e.g., "C_18.00(0001)")
    Returns:
        Tuple of (excel_row, excel_col, actual_worksheet) or None
    """
    try:
        # Find the target worksheet - must match the EXACT sheet name including sheet number
        target_ws = None
        
        # First try exact match
        if sheet_name in form_wb.sheetnames:
            target_ws = form_wb[sheet_name]
            #logging.debug(f"   Exact match found: {sheet_name}")
        else:
            # Try to find matching sheet with sheet number
            for ws_name in form_wb.sheetnames:
                # For sheet_name like "C_18.00(0001)", match exactly
                if sheet_name == ws_name:
                    target_ws = form_wb[ws_name]
                    #logging.debug(f"   Exact sheet match: {ws_name}")
                    break
                # For partial matches, check base name and sheet number
                elif sheet_name.replace('_', ' ') == ws_name.replace('_', ' '):
                    target_ws = form_wb[ws_name]
                    #logging.debug(f"   Sheet match with space/underscore: {ws_name}")
                    break
        
        if not target_ws:
            logging.warning(f"Could not find worksheet for {sheet_name}")
            logging.debug(f"   Available sheets: {form_wb.sheetnames}")
            return None
        
        # Find form structure markers
        markers = find_form_structure_markers(target_ws)
        
        if not markers['rowNumbering_pos'] or not markers['rowLabelColumnNumbering_pos']:
            logging.warning(f"Could not find form structure markers in {target_ws.title}")
            # Fallback to simple mapping
            target_row_num = int(row_str)
            target_col_num = int(col_str)
            excel_row = target_row_num // 10 + 15  # More conservative offset
            excel_col = target_col_num // 10 + 2   # More conservative offset
            return excel_row, excel_col, target_ws
        
        # Convert validation rule coordinates to actual Excel positions
        target_row_code = f"{int(row_str):04d}"  # Ensure 4-digit format
        target_col_code = f"{int(col_str):04d}"  # Ensure 4-digit format
        
        # Search for actual row position below rowNumbering marker
        excel_row = None
        row_marker_row, row_marker_col = markers['rowNumbering_pos']
        
        # Look for the target row code below the rowNumbering marker
        for search_row in range(row_marker_row + 1, min(row_marker_row + 200, target_ws.max_row + 1)):
            cell_value = target_ws.cell(row=search_row, column=row_marker_col).value
            if cell_value and target_row_code in str(cell_value):
                excel_row = search_row
                #logging.debug(f"   Found row {target_row_code} at Excel row {excel_row} in {target_ws.title}")
                break
        
        # Search for actual column position to the right of rowLabelColumnNumbering marker
        excel_col = None
        col_marker_row, col_marker_col = markers['rowLabelColumnNumbering_pos']
        
        # Look for the target column code to the right of the rowLabelColumnNumbering marker
        found_columns = []  # Track what columns we find for debugging
        for search_col in range(col_marker_col + 1, min(col_marker_col + 100, target_ws.max_column + 1)):
            cell_value = target_ws.cell(row=col_marker_row, column=search_col).value
            if cell_value:
                found_columns.append(str(cell_value))
            if cell_value and target_col_code in str(cell_value):
                excel_col = search_col
                excel_col_letter = get_column_letter(excel_col)
                #logging.debug(f"   Found column {target_col_code} at Excel column {excel_col_letter} in {target_ws.title}")
                break
        
        # If we couldn't find exact matches, log warning and return None
        if excel_row is None:
            logging.warning(f"    Row {target_row_code} not found in form structure for {target_ws.title}")
            return None
        
        if excel_col is None:
            logging.warning(f"    Column {target_col_code} not found in form structure for {target_ws.title}")
            return None
        
        # Validate coordinates are reasonable
        if excel_row > target_ws.max_row + 20 or excel_col > target_ws.max_column + 20:
            excel_col_letter = get_column_letter(excel_col)
            logging.warning(f"Calculated coordinates seem too large: row {excel_row}, col {excel_col_letter} in {target_ws.title}")
            return None
        
        return excel_row, excel_col, target_ws
        
    except Exception as e:
        logging.warning(f"Error converting coordinates {row_str}, {col_str}: {str(e)}")
        return None


def apply_feeding_rules_to_form(form_file_path, validation_rules_path):
    """
    Apply feeding rules to the pure-form file
    Args:
        form_file_path: Path to the form file
        validation_rules_path: Path to validation rules file
    Returns:
        Boolean indicating success
    """
    global FORM_STRUCTURE_CACHE
    
    logging.info(f"   Applying feeding rules to: {form_file_path}")
    
    # Clear cache at the beginning of each processing
    FORM_STRUCTURE_CACHE.clear()
    #logging.debug("   Cleared form structure cache")
    
    try:
        # Load the form workbook with macro preservation
        form_wb = openpyxl.load_workbook(form_file_path, keep_vba=True)
        
        # Identify root sheet
        root_sheet = identify_root_sheet(form_wb)
        if not root_sheet:
            logging.error("Could not identify root sheet")
            return False
        
        # Load validation rules
        rules = load_validation_rules(validation_rules_path, root_sheet)
        if not rules:
            logging.warning("No applicable validation rules found")
            return True  # Not an error, just no rules to apply
        
        # Build complete formulas
        complete_formulas = build_complete_formulas(rules, form_wb, root_sheet)
        
        # Apply each formula - only add values to existing cells
        applied_count = 0
        logging.debug("   ✅ Writing formulas in generated Excel :")
        for formula_data in complete_formulas:
            if apply_single_formula(form_wb, formula_data):
                applied_count += 1
        
        logging.info(f"    🟢 Applied {applied_count} feeding rules successfully")
        # Save the workbook carefully
        form_wb.save(form_file_path)
        
        # Apply post-processing to simplify and mutualize identical formulas
        logging.info("   🔧 Starting post-processing: simplifying formulas, mutualizing identical ones, and completing coordinates...")
        postprocess_count = postprocess_mutualize_formulas(form_wb)
        
        # Save again after post-processing
        form_wb.save(form_file_path)
        form_wb.close()
        
        logging.info(f"    🔄 Post-processed {postprocess_count} operations (simplifications + mutualizations + coordinate completions)")
        #logging.debug(f"   Cache was used for {len(FORM_STRUCTURE_CACHE)} worksheets")
        
        return True
        
    except Exception as e:
        logging.error(f"Error applying feeding rules: {str(e)}")
        return False


def simplify_formula_in_cell(cell, sheet_name, row_num, col_num):
    """
    Simplify a formula by removing unnecessary information (step 6.a)
    Args:
        cell: openpyxl cell object
        sheet_name: Current sheet name
        row_num: Current row number 
        col_num: Current column number
    Returns:
        Simplified formula string or None if no changes needed
    """
    if not cell.value or not isinstance(cell.value, str):
        return None
    
    formula_text = str(cell.value)
    
    # Only process formulas (text starting with '=)
    if not formula_text.startswith("'="):
        return None
    
    try:
        # Extract the current cell's coordinate information
        current_sheet_info = sheet_name.replace('_', ' ')  # Handle sheet name variants
        current_row_code = f"r{row_num:04d}"  # Format as r0020, etc.
        current_col_code = f"c{col_num:04d}"  # Format as c0020, etc.
        
        logging.debug(f"      Simplifying formula in {sheet_name}[{row_num},{col_num}]: {formula_text}")
        
        # Pattern to find cell references like {C_18.00(0001), r0030, c0020}
        pattern = r'\{([^}]+)\}'
        
        def simplify_cell_ref(match):
            content = match.group(1)
            parts = [p.strip() for p in content.split(',')]
            
            # Check what can be simplified
            simplified_parts = []
            
            for part in parts:
                # Remove sheet name if it matches current sheet
                if (part.replace('_', ' ') == current_sheet_info or 
                    part.replace(' ', '_') == sheet_name):
                    continue  # Skip identical sheet info
                
                # Remove row info if it matches current row
                if part == current_row_code:
                    continue  # Skip identical row info
                
                # Remove column info if it matches current column  
                if part == current_col_code:
                    continue  # Skip identical column info
                
                # Keep different information
                simplified_parts.append(part)
            
            # If all parts were removed, keep at least the row/column that's different
            if not simplified_parts:
                # This shouldn't happen in normal cases, but as fallback
                return match.group(0)  # Return original
            
            return '{' + ', '.join(simplified_parts) + '}'
        
        # Apply simplification to all cell references in the formula
        simplified_formula = re.sub(pattern, simplify_cell_ref, formula_text)
        
        # Only return if actually simplified
        if simplified_formula != formula_text:
            logging.debug(f"                                    Simplified to: {simplified_formula}")
            return simplified_formula
        
        return None
        
    except Exception as e:
        logging.warning(f"Error simplifying formula in {sheet_name}[{row_num},{col_num}]: {str(e)}")
        return None


def simplify_formulas_in_table(worksheet, start_row, start_col, end_row, end_col):
    """
    Simplify formulas in the data table (step 6.a)
    Args:
        worksheet: openpyxl worksheet
        start_row, start_col, end_row, end_col: Table boundaries
    Returns:
        Number of formulas simplified
    """
    simplify_count = 0
    
    try:
        # Get the sheet name for coordinate calculations
        sheet_name = worksheet.title
        
        # Find form structure markers to get proper row/column mapping
        markers = find_form_structure_markers(worksheet)
        if not markers['rowNumbering_pos'] or not markers['rowLabelColumnNumbering_pos']:
            logging.warning(f"Could not find form structure markers in {sheet_name}")
            return 0
        
        row_marker_row, row_marker_col = markers['rowNumbering_pos']
        col_marker_row, col_marker_col = markers['rowLabelColumnNumbering_pos']
        
        for row in range(start_row, end_row + 1):
            for col in range(start_col, end_col + 1):
                cell = worksheet.cell(row=row, column=col)
                
                # Skip empty cells
                if not cell.value:
                    continue
                
                # Calculate the logical row/column codes from Excel coordinates
                logical_row = (row - row_marker_row - 1) * 10 + 10  # Approximate mapping
                logical_col = (col - col_marker_col - 1) * 10 + 10  # Approximate mapping
                
                # Try to get more precise mapping from actual form structure
                try:
                    # Look for actual row code in the rowNumbering column
                    actual_row_cell = worksheet.cell(row=row, column=row_marker_col)
                    if actual_row_cell.value:
                        row_matches = re.findall(r'\b(\d{4})\b', str(actual_row_cell.value))
                        if row_matches:
                            logical_row = int(row_matches[0])
                    
                    # Look for actual column code in the rowLabelColumnNumbering row
                    actual_col_cell = worksheet.cell(row=col_marker_row, column=col)
                    if actual_col_cell.value:
                        col_matches = re.findall(r'\b(\d{4})\b', str(actual_col_cell.value))
                        if col_matches:
                            logical_col = int(col_matches[0])
                except:
                    pass  # Use calculated values as fallback
                
                # Simplify the formula in this cell
                simplified = simplify_formula_in_cell(cell, sheet_name, logical_row, logical_col)
                
                if simplified:
                    cell.value = simplified
                    simplify_count += 1
        
        if simplify_count > 0:
            logging.debug(f"     Simplified {simplify_count} formulas in {sheet_name}")
        
    except Exception as e:
        logging.warning(f"Error simplifying formulas in table: {str(e)}")
    
    return simplify_count


def postprocess_mutualize_formulas(workbook):
    """
    Post-process to simplify and mutualize identical formulas in rows and columns
    Args:
        workbook: openpyxl workbook
    Returns:
        Number of operations performed (simplifications + mutualizations)
    """
    total_operations = 0
    ignore_sheets = ["identification", "cellSpecificities", "labelTranslation"]
    
    for sheet_name in workbook.sheetnames:
        if sheet_name in ignore_sheets:
            continue
            
        worksheet = workbook[sheet_name]
        logging.debug(f"     Post-processing sheet: {sheet_name}")
        
        # Find the data table boundaries
        table_bounds = find_data_table_bounds(worksheet)
        if not table_bounds:
            logging.debug(f"     No data table found in {sheet_name}")
            continue
            
        start_row, start_col, end_row, end_col = table_bounds
        
        # Step 6.a: Simplify formulas by removing unnecessary information
        simplify_count = simplify_formulas_in_table(worksheet, start_row, start_col, end_row, end_col)
        
        # Step 6.b: Mutualize formulas in rows (move common formulas to column G)
        row_mutualizations = mutualize_row_formulas(worksheet, start_row, start_col, end_row, end_col)
        
        # Step 6.b: Mutualize formulas in columns (move common formulas to row 10)
        col_mutualizations = mutualize_column_formulas(worksheet, start_row, start_col, end_row, end_col)
        
        # Step 6.c: Complete coordinates in remaining formulas (add missing row/column info)
        complete_coordinates = complete_formula_coordinates(worksheet, start_row, start_col, end_row, end_col)
        
        sheet_total = simplify_count + row_mutualizations + col_mutualizations + complete_coordinates
        total_operations += sheet_total
        
        if sheet_total > 0:
            logging.debug(f"     Sheet {sheet_name}: {simplify_count} simplifications, {row_mutualizations} row mutualizations, {col_mutualizations} column mutualizations, {complete_coordinates} coordinate completions")
    
    return total_operations


def find_data_table_bounds(worksheet):
    """
    Find the bounds of the data table in the worksheet
    Args:
        worksheet: openpyxl worksheet
    Returns:
        Tuple of (start_row, start_col, end_row, end_col) or None
    """
    try:
        # Find form structure markers
        markers = find_form_structure_markers(worksheet)
        
        if not markers['rowNumbering_pos'] or not markers['rowLabelColumnNumbering_pos']:
            return None
            
        row_marker_row, row_marker_col = markers['rowNumbering_pos']
        col_marker_row, col_marker_col = markers['rowLabelColumnNumbering_pos']
        
        # Data table starts at (rowLabelColumnNumbering cell + 1, rowLabelColumnNumbering cell + 1)
        start_row = col_marker_row + 1
        start_col = col_marker_col + 1
        
        # Find end of table by looking for the actual data boundaries
        end_row = start_row
        end_col = start_col
        
        # Find end row by looking for thick bottom border
        max_search_row = min(start_row + 200, worksheet.max_row + 1)
        for row in range(start_row, max_search_row):
            # Check if this row has a thick bottom border across the table width
            has_thick_bottom_border = False
            for col in range(start_col, min(start_col + 100, worksheet.max_column + 1)):
                cell = worksheet.cell(row=row, column=col)
                if (cell.border and cell.border.bottom and 
                    cell.border.bottom.style in ['thick', 'medium']):
                    has_thick_bottom_border = True
                    break
            
            if has_thick_bottom_border:
                end_row = row
                break
            else:
                # If no thick border found, use data presence as fallback
                has_data = False
                for col in range(start_col, min(start_col + 100, worksheet.max_column + 1)):
                    cell = worksheet.cell(row=row, column=col)
                    if cell.value is not None and str(cell.value).strip():
                        has_data = True
                        break
                if has_data:
                    end_row = row
        
        # Find end column by looking for thick right border (as specified in prompt)
        max_search_col = min(start_col + 100, worksheet.max_column + 1)
        for col in range(start_col, max_search_col):
            # Check if this column has a thick right border across the table height
            has_thick_right_border = False
            for row in range(start_row, end_row + 1):
                cell = worksheet.cell(row=row, column=col)
                if (cell.border and cell.border.right and 
                    cell.border.right.style in ['thick', 'medium']):
                    has_thick_right_border = True
                    break
            
            if has_thick_right_border:
                end_col = col
                break
            else:
                # If no thick border found, use data presence as fallback
                has_data = False
                for row in range(start_row, end_row + 1):
                    cell = worksheet.cell(row=row, column=col)
                    if cell.value is not None and str(cell.value).strip():
                        has_data = True
                        break
                if has_data:
                    end_col = col
        
        logging.debug(f"     Found data table bounds: ({start_row},{get_column_letter(start_col)}) to ({end_row},{get_column_letter(end_col)})")
        return start_row, start_col, end_row, end_col
        
    except Exception as e:
        logging.warning(f"Error finding data table bounds: {str(e)}")
        return None


def mutualize_row_formulas(worksheet, start_row, start_col, end_row, end_col):
    """
    Mutualize formulas that appear in more than 50% of pink cells within rows (move to column G)
    Args:
        worksheet: openpyxl worksheet
        start_row, start_col, end_row, end_col: Table boundaries
    Returns:
        Number of row mutualizations performed
    """
    mutualize_count = 0
    pink_color = 'FFFF99CC'  # Pink cells to consider (FF99CC)
    
    try:
        for row in range(start_row, end_row + 1):
            # First, count all pink cells in this row and extract individual formulas
            pink_cells = []
            row_formulas = {}  # formula -> list of (row, col, all_formulas_in_cell)
            
            for col in range(start_col, end_col + 1):
                cell = worksheet.cell(row=row, column=col)
                
                # Only consider pink cells
                if not (cell.fill and cell.fill.start_color and cell.fill.start_color.rgb == pink_color):
                    continue
                
                # This is a pink cell
                pink_cells.append((row, col))
                    
                # Extract individual formulas from cells (handle multiple formulas per cell)
                if cell.value and isinstance(cell.value, str):
                    cell_content = str(cell.value)
                    # Split by newlines to handle multiple formulas in one cell
                    individual_formulas = [f.strip() for f in cell_content.split('\n') if f.strip().startswith("'=")]
                    
                    for formula in individual_formulas:
                        if formula not in row_formulas:
                            row_formulas[formula] = []
                        row_formulas[formula].append((row, col, cell_content))
            
            # Only proceed if there are at least one formula and some pink cells
            if row_formulas and len(pink_cells) > 0:
                # Find formulas that can be mutualized
                for formula, cells_with_formula in row_formulas.items():
                    # Exception: Don't mutualize formulas containing the operator "/"
                    if "/" in formula:
                        logging.debug(f"     * Skipping mutualization of formula containing '/': {formula}")
                        continue
                    
                    # Mutualize formula if it appears in more than 50% of pink cells and more than 2 cells
                    formula_percentage = len(cells_with_formula) / len(pink_cells)
                    should_mutualize = formula_percentage > 0.5 and len(cells_with_formula) > 2
                    
                    if should_mutualize:
                        # Move formula to column G (column 7)
                        g_cell = worksheet.cell(row=row, column=7)
                        if g_cell.value:
                            g_cell.value = str(g_cell.value) + "; " + formula
                        else:
                            g_cell.value = formula
                        # Force Excel to format as text to make apostrophe invisible
                        g_cell.number_format = '@'  # @ is Excel's text format
                        g_cell.data_type = 's'  # 's' for string/text
                        
                        # Remove formula from other cells
                        for cell_row, cell_col, original_content in cells_with_formula:
                            cell = worksheet.cell(row=cell_row, column=cell_col)
                            
                            # Handle removal of specific formula from multi-formula cells
                            if original_content == formula:
                                # Cell contains only this formula
                                cell.value = None
                            elif formula in original_content:
                                # Cell contains multiple formulas, remove only this one
                                lines = original_content.split('\n')
                                remaining_lines = [line for line in lines if line.strip() != formula]
                                new_content = '\n'.join(remaining_lines).strip()
                                cell.value = new_content if new_content else None
                        
                        mutualize_count += 1
                        logging.debug(f"   Mutualized row formula at G{row}: {formula} (found in {len(cells_with_formula)}/{len(pink_cells)} pink cells, {formula_percentage:.1%})")
        
    except Exception as e:
        logging.warning(f"Error mutualizing row formulas: {str(e)}")
    
    return mutualize_count


def mutualize_column_formulas(worksheet, start_row, start_col, end_row, end_col):
    """
    Mutualize formulas that appear in more than 50% of pink cells within columns (move to row 10)
    Args:
        worksheet: openpyxl worksheet
        start_row, start_col, end_row, end_col: Table boundaries
    Returns:
        Number of column mutualizations performed
    """
    mutualize_count = 0
    pink_color = 'FFFF99CC'  # Pink cells to consider (FF99CC)
    
    try:
        for col in range(start_col, end_col + 1):
            # First, count all pink cells in this column and extract individual formulas
            pink_cells = []
            col_formulas = {}  # formula -> list of (row, col, all_formulas_in_cell)
            
            for row in range(start_row, end_row + 1):
                cell = worksheet.cell(row=row, column=col)
                
                # Only consider pink cells
                if not (cell.fill and cell.fill.start_color and cell.fill.start_color.rgb == pink_color):
                    continue
                
                # This is a pink cell
                pink_cells.append((row, col))
                    
                # Extract individual formulas from cells (handle multiple formulas per cell)
                if cell.value and isinstance(cell.value, str):
                    cell_content = str(cell.value)
                    # Split by newlines to handle multiple formulas in one cell
                    individual_formulas = [f.strip() for f in cell_content.split('\n') if f.strip().startswith("'=")]
                    
                    for formula in individual_formulas:
                        if formula not in col_formulas:
                            col_formulas[formula] = []
                        col_formulas[formula].append((row, col, cell_content))
            
            # Only proceed if there are at least one formula and some pink cells
            if col_formulas and len(pink_cells) > 0:
                # Find formulas that can be mutualized
                for formula, cells_with_formula in col_formulas.items():
                    # Exception: Don't mutualize formulas containing the operator "/"
                    if "/" in formula:
                        logging.debug(f"     * Skipping mutualization of formula containing '/': {formula}")
                        continue
                    
                    # Mutualize formula if it appears in more than 50% of pink cells and more than 2 cells
                    formula_percentage = len(cells_with_formula) / len(pink_cells)
                    should_mutualize = formula_percentage > 0.5 and len(cells_with_formula) > 2
                    
                    if should_mutualize:
                        # Move formula to row 10
                        row_10_cell = worksheet.cell(row=10, column=col)
                        if row_10_cell.value:
                            row_10_cell.value = str(row_10_cell.value) + "; " + formula
                        else:
                            row_10_cell.value = formula
                        # Force Excel to format as text to make apostrophe invisible
                        row_10_cell.number_format = '@'  # @ is Excel's text format
                        row_10_cell.data_type = 's'  # 's' for string/text
                        
                        
                        # Remove formula from other cells
                        for cell_row, cell_col, original_content in cells_with_formula:
                            cell = worksheet.cell(row=cell_row, column=cell_col)
                            
                            # Handle removal of specific formula from multi-formula cells
                            if original_content == formula:
                                # Cell contains only this formula
                                cell.value = None
                            elif formula in original_content:
                                # Cell contains multiple formulas, remove only this one
                                lines = original_content.split('\n')
                                remaining_lines = [line for line in lines if line.strip() != formula]
                                new_content = '\n'.join(remaining_lines).strip()
                                cell.value = new_content if new_content else None
                        
                        mutualize_count += 1
                        col_letter = get_column_letter(col)
                        logging.debug(f"   Mutualized column formula at {col_letter}10: {formula} (found in {len(cells_with_formula)}/{len(pink_cells)} pink cells, {formula_percentage:.1%})")
        
    except Exception as e:
        logging.warning(f"Error mutualizing column formulas: {str(e)}")
    
    return mutualize_count


def complete_formula_coordinates(worksheet, start_row, start_col, end_row, end_col):
    """
    Complete partial coordinates in formulas with missing row/column information (step 6.c)
    Args:
        worksheet: openpyxl worksheet
        start_row, start_col, end_row, end_col: Table boundaries
    Returns:
        Number of formulas with completed coordinates
    """
    complete_count = 0
    grey_color = 'FF808080'  # Grey cells to ignore
    
    try:
        # Find form structure markers to get proper row/column mapping
        markers = find_form_structure_markers(worksheet)
        if not markers['rowNumbering_pos'] or not markers['rowLabelColumnNumbering_pos']:
            logging.warning(f"Could not find form structure markers in {worksheet.title}")
            return 0
        
        row_marker_row, row_marker_col = markers['rowNumbering_pos']
        col_marker_row, col_marker_col = markers['rowLabelColumnNumbering_pos']
        
        for row in range(start_row, end_row + 1):
            for col in range(start_col, end_col + 1):
                cell = worksheet.cell(row=row, column=col)
                
                # Skip empty cells
                if not cell.value:
                    continue
                
                # Skip grey cells
                if cell.fill and cell.fill.start_color and cell.fill.start_color.rgb == grey_color:
                    continue
                
                # Only process cells that contain formulas (text starting with '=)
                cell_content = str(cell.value)
                if not cell_content.startswith("'="):
                    continue
                
                # Calculate the logical row/column codes from Excel coordinates
                logical_row = (row - row_marker_row - 1) * 10 + 10  # Approximate mapping
                logical_col = (col - col_marker_col - 1) * 10 + 10  # Approximate mapping
                
                # Try to get more precise mapping from actual form structure
                try:
                    # Look for actual row code in the rowNumbering column
                    actual_row_cell = worksheet.cell(row=row, column=row_marker_col)
                    if actual_row_cell.value:
                        row_matches = re.findall(r'\b(\d{4})\b', str(actual_row_cell.value))
                        if row_matches:
                            logical_row = int(row_matches[0])
                    
                    # Look for actual column code in the rowLabelColumnNumbering row
                    actual_col_cell = worksheet.cell(row=col_marker_row, column=col)
                    if actual_col_cell.value:
                        col_matches = re.findall(r'\b(\d{4})\b', str(actual_col_cell.value))
                        if col_matches:
                            logical_col = int(col_matches[0])
                except:
                    pass  # Use calculated values as fallback
                
                # Complete the coordinates in this cell's formula
                completed_formula = complete_coordinates_in_formula(
                    cell_content, 
                    logical_row, 
                    logical_col
                )
                
                if completed_formula and completed_formula != cell_content:
                    cell.value = completed_formula
                    complete_count += 1
                    excel_col_letter = get_column_letter(col)
                    logging.debug(f"     Completed coordinates in {worksheet.title}[{row},{excel_col_letter}]: {cell_content} -> {completed_formula}")
        
        if complete_count > 0:
            logging.debug(f"     Completed coordinates in {complete_count} formulas in {worksheet.title}")
        
    except Exception as e:
        logging.warning(f"Error completing formula coordinates: {str(e)}")
    
    return complete_count


def complete_coordinates_in_formula(formula_text, current_row, current_col):
    """
    Complete partial coordinates in a formula by adding missing row/column information
    Args:
        formula_text: Formula text like '= {r0030} + {r0080} + {r0120}'
        current_row: Current logical row number (e.g., 20 for r0020)
        current_col: Current logical column number (e.g., 20 for c0020)
    Returns:
        Formula with completed coordinates or None if no changes needed
    """
    try:
        # Pattern to find cell references like {r0030}, {c0060}, {r0030, c0060}
        pattern = r'\{([^}]+)\}'
        
        def complete_cell_ref(match):
            content = match.group(1)
            parts = [p.strip() for p in content.split(',')]
            
            # Check what's present and what's missing
            has_row = any(p.startswith('r') for p in parts)
            has_col = any(p.startswith('c') for p in parts)
            
            # Only add missing coordinates if at least one coordinate is already present
            # (don't modify sheet-only references)
            if not has_row and not has_col:
                # This is likely a sheet reference only, don't modify
                return match.group(0)
            
            new_parts = []
            
            # Add existing parts
            for part in parts:
                new_parts.append(part)
            
            # Add missing row if not present but column is present
            if not has_row and has_col:
                new_parts.append(f"r{current_row:04d}")
            
            # Add missing column if not present but row is present
            if has_row and not has_col:
                new_parts.append(f"c{current_col:04d}")
            
            return '{' + ','.join(new_parts) + '}'
        
        # Apply completion to all cell references in the formula
        completed_formula = re.sub(pattern, complete_cell_ref, formula_text)
        
        # Only return if actually completed
        if completed_formula != formula_text:
            return completed_formula
        
        return None
        
    except Exception as e:
        logging.warning(f"Error completing coordinates in formula: {str(e)}")
        return None


def apply_single_formula(workbook, formula_data):
    """
    Apply a single formula to the workbook - safely add only text values to the correct sheet
    Args:
        workbook: openpyxl workbook
        formula_data: Dictionary with formula and metadata
    Returns:
        Boolean indicating success
    """
    try:
        formula = formula_data['formula']
        target_sheet_name = formula_data['sheet']
        
        # Find target cell
        target_cell_ref = find_target_cell(formula)
        if not target_cell_ref:
            logging.warning(f"Could not find target cell in formula: {formula}")
            return False
        
        # Parse target cell
        target_info = parse_cell_reference(target_cell_ref)
        
        # Convert to Excel coordinates - now returns the actual worksheet too
        coords_result = convert_to_excel_coords(target_info['row'], target_info['column'], 
                                              workbook, target_sheet_name)
        if not coords_result:
            logging.warning(f"⚠️  Could not apply formula: {formula} - Row/Column not found in form structure")
            return False
        
        excel_row, excel_col, target_ws = coords_result
        
        # Check if coordinates are within reasonable bounds to avoid corruption
        if excel_row > 1000 or excel_col > 100:
            excel_col_letter = get_column_letter(excel_col)
            logging.warning(f"Coordinates too large: row {excel_row}, col {excel_col_letter} in {target_ws.title}")
            return False
        
        # Apply the formula as text (safe approach) - concatenate with existing content
        target_cell = target_ws.cell(row=excel_row, column=excel_col)
        formula_text = convert_to_excel_formula(formula, target_cell_ref)
        
        # Apply pink background color (#FF99CC - Red: 255, Green: 153, Blue: 204)
        pink_fill = PatternFill(start_color='FFFF99CC', end_color='FFFF99CC', fill_type='solid')
        target_cell.fill = pink_fill
        
        # Check if cell already has content (NB: if 2 rules are present in the same cell)
        if target_cell.value is not None:
            # Generate warning and keep only the first rule (do not concatenate)
            existing_content = str(target_cell.value).strip()
            if existing_content:  # Only warn if there's actual content, not just whitespace
                excel_col_letter = get_column_letter(excel_col)
                logging.warning(f"⚠️  Multiple rules for same cell {target_ws.title}[{excel_row},{excel_col_letter}]: "
                              f"Keeping first rule '{existing_content}', ignoring new rule '{formula_text}'")
                return True  # Return success but don't modify the cell
        
        # Cell is empty or only contains whitespace, safe to add formula text
        target_cell.value = formula_text
        excel_col_letter = get_column_letter(excel_col)
        logging.debug(f"   Applied formula to {target_ws.title}[{excel_row},{excel_col_letter}]: {formula_text}")
        
        # Multiple approaches to force text formatting and hide apostrophe
        target_cell.number_format = '@'  # @ is Excel's text format
        target_cell.data_type = 's'  # 's' for string/text
        
        return True
        
    except Exception as e:
        logging.warning(f"Error applying formula {formula_data}: {str(e)}")
        return False


def convert_to_excel_formula(validation_formula, target_cell_ref):
    """
    Convert validation rule formula to Excel formula format, extracting only the non-target part
    Args:
        validation_formula: Formula like {C18.00(002), r0010, c0060} = {C18.00(002), r0010, c0070}
        target_cell_ref: Target cell reference like {C18.00(002), r0010, c0060}
    Returns:
        Excel formula string with only the calculation part
    """
    try:
        # Split the formula on = sign
        parts = validation_formula.split('=', 1)
        if len(parts) != 2:
            return f"'{validation_formula}"  # Fallback to original if not a proper equation
        
        left_side = parts[0].strip()
        right_side = parts[1].strip()
        
        # Determine which side is the target and which is the formula
        if target_cell_ref.strip() == left_side:
            # Target is on left, formula is on right
            formula_part = f"= {right_side}"
            #logging.debug(f"  Target is on left, formula is on right : {formula_part}")
        elif target_cell_ref.strip() == right_side:
            # Target is on right, formula is on left
            formula_part = f"= {left_side}"
            #logging.debug(f"  Target is on right, formula is on left : {formula_part}")
        else:
            # Target not found exactly, default to right side
            formula_part = f"= {right_side}"
            #logging.debug(f"  Target not found exactly, default to right side : {formula_part}")
        
        # Add apostrophe prefix to treat as text, preventing Excel corruption
        return f"'{formula_part}"
        
    except Exception as e:
        logging.warning(f"Error converting formula {validation_formula}: {str(e)}")
        return f"'{validation_formula}"  # Fallback to original


def main():
    """Main function for standalone execution"""
    parser = argparse.ArgumentParser(description='Process feeding rules from pure validation rules')
    parser.add_argument('--validation-file', '-v', 
                       help='Path to pure-validationRules-xxx.xlsx file')
    parser.add_argument('--form-file', '-f',
                       help='Path to pure-form....xlsm file to modify')
    parser.add_argument('--debug', '-d', action='store_true',
                       help='Enable debug mode')
    
    args = parser.parse_args()
    
    # Configure basic logging for console first
    configure_logging(args.debug)
    
    logging.info("= ℹ️  Starting Feeding Rules Processor =")
    
    # Get validation rules file
    validation_file = args.validation_file
    if not validation_file:
        validation_file = get_validation_rules_file()
        if not validation_file:
            return 1
    
    # Get pure-form file
    form_file = args.form_file
    if not form_file:
        form_file = get_pure_form_file()
        if not form_file:
            return 1

    # If debug mode not specified in command line, ask user via dialog
    debug_mode = args.debug
    if not debug_mode:
        # Create root window for dialog
        root = Tk()
        root.withdraw()  # Hide the main window
        
        # Ask user for debug mode
        debug_mode = messagebox.askyesno("Debug Mode", "Enable debug mode?")
        
        root.destroy()  # Clean up the root window
    
    # Reconfigure logging with user's choice
    configure_logging(debug_mode)
    
    # Create output directory and copy file
    output_dir = create_output_directory(form_file)
    copied_file = copy_form_to_output(form_file, output_dir)
    if not copied_file:
        logging.error("Failed to copy form file")
        return 1
    
    # Setup dedicated file logger for this Excel file
    dedicated_logger = None
    try:
        import excel_logger
        dedicated_logger = excel_logger.setup_excel_file_logging(copied_file, "feedingRuleSource")
        
        # Log initial information
        if dedicated_logger:
            dedicated_logger.info("=" * 70)
            dedicated_logger.info("FEEDING RULES PROCESSING - STANDALONE MODE")
            dedicated_logger.info("=" * 70)
            dedicated_logger.info(f"Source form file: {form_file}")
            dedicated_logger.info(f"Validation rules file: {validation_file}")
            dedicated_logger.info(f"Output file: {copied_file}")
            dedicated_logger.info(f"Debug mode: {debug_mode}")
            dedicated_logger.info("=" * 70)
            
            # Redirect logging to include the dedicated logger
            for handler in dedicated_logger.handlers:
                if handler not in logging.root.handlers:
                    logging.root.addHandler(handler)
                    if debug_mode:
                        handler.setLevel(logging.DEBUG)
    except ImportError:
        logging.warning("Excel logger not available - logging to console only")
    
    # Apply feeding rules
    success = apply_feeding_rules_to_form(copied_file, validation_file)
    
    # Finalize logging
    if dedicated_logger:
        try:
            import excel_logger
            if success:
                dedicated_logger.info("=" * 70)
                dedicated_logger.info("🟢 Processing completed successfully")
                dedicated_logger.info(f"📁 Modified file saved as: {copied_file}")
                dedicated_logger.info("=" * 70)
                excel_logger.complete_excel_logging(
                    copied_file,
                    success=True,
                    total_sheets=1,
                    processed_sheets=["Feeding Rules Applied"],
                    form_code="N/A",
                    feeding_rules_applied=True
                )
            else:
                dedicated_logger.error("=" * 70)
                dedicated_logger.error("❌ Processing failed")
                dedicated_logger.error("=" * 70)
                excel_logger.complete_excel_logging(
                    copied_file,
                    success=False,
                    error="Feeding rules processing failed"
                )
        except Exception as e:
            logging.warning(f"Error finalizing Excel logging: {e}")
    
    if success:
        logging.info("🟢 Processing completed successfully")
        logging.info(f"📁 Modified file saved as: {copied_file}")
        
        # Display log file location
        if dedicated_logger:
            try:
                import excel_logger
                log_file = excel_logger.excel_file_logger.log_files.get(copied_file)
                if log_file:
                    logging.info(f"📋 Log file: {log_file}")
            except:
                pass
        
        return 0
    else:
        logging.error("❌ Processing failed")
        return 1


# Function to be called from forms_from_dpm.py
def process_excel_with_validation_rules(target_excel_path, validation_file_path=None, debug_mode=False, dedicated_logger=None):
    """
    Function to be called from other scripts (like forms_from_dmp.py)
    Args:
        target_excel_path: Path to the Excel file to process
        validation_file_path: Optional path to validation rules file
        debug_mode: Enable debug logging
        dedicated_logger: Optional logger to use (from calling script)
    Returns:
        Boolean indicating success
    """
    # If a dedicated logger is provided, preserve existing handlers and use it
    if dedicated_logger:
        configure_logging(debug_mode, preserve_handlers=True)
        # Redirect root logger to use the dedicated logger's handlers
        for handler in dedicated_logger.handlers:
            if handler not in logging.root.handlers:
                logging.root.addHandler(handler)
                # Ensure the handler level allows DEBUG messages if debug_mode is enabled
                if debug_mode:
                    handler.setLevel(logging.DEBUG)
    else:
        configure_logging(debug_mode, preserve_handlers=False)
    
    # Get validation rules file if not provided
    if not validation_file_path:
        logging.info("🔧 Requesting validation rules file from user...")
        validation_file_path = get_validation_rules_file()
        if not validation_file_path:
            logging.warning("No validation rules file selected by user")
            return False
    
    # Apply feeding rules directly to the target file (no copy needed when called from forms_from_dpm)
    return apply_feeding_rules_to_form(target_excel_path, validation_file_path)


if __name__ == "__main__":
    sys.exit(main())
