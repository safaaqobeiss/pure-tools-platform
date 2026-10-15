"""
### 🔧 **Prompt 

**Create one output Excel file from one input Excel file using an Excel template**

1. Create Python code to generate a target Excel file based on all sheets in a single input Excel workbook.
   * Ask the user the path of the input Excel file in format xls, xlsx or xlsm (with file dialog)
   * The formCode is automatically derived from the input filename (FRM_XXX or FRM_YYYY where XXX/YYYY are 3-4 digits)
   * Determine formCode length (3 or 4 digits) - this affects data mapping and conversion logic
   * Ask the user about 5/6-digit to 4-digit rowID conversion (only if formCode length ≠ 4):
     - Show dialog: "Would you like to convert 5/6-digit rowID's into 4-digit rowID's?"
     - Dropdown options: "Yes" or "No"
     - If "Yes": conversion_6_4_digits = "Y" (uses Mapping_Rub_6_4_digits.xlsx)
     - If "No": conversion_6_4_digits = "N"
     - If formCode length = 4: conversion_6_4_digits automatically = "N" (no dialog shown)
   * Ask the user for optional fields with default values:
     - formGroupCode (default: "schemaaXYZ")
     - version (default: "2026.1.0")
     - validityStartingDate (default: "01/01/2026")
     - linkedDomain (default: "bebnb")
     - datatype (default: "bebnb.genericPositiveAmount")
     - roundingUnit (default: "thousands")
   * Inform the user that the folder "generated_output" for generated files will be created if not exist as a subfolder of the current directory
   * The template file to be used is here "pure-form-regulatoryOutputCode-formCode-(shortName).xlsm" in the current directory
   * One input Excel file will generate exactly one output Excel file

2. For each sheet in the input workbook:
   * Filter sheets: Only process sheets whose names start with "Sheet" (skip Presentation, Properties, Cell particularities, etc.)
   * Create SA.*.* sheets in the output file:
     - Extract formCode from input filename (FRM_XXX or FRM_YYYY → XXX or YYYY)
     - Determine formCode length (3 or 4 digits) - this will determine data mapping strategy:
       * Length 3: formCode like 010, 020, 030, etc.
       * Length 4: formCode like 2152, 1234, etc.
     - Apply naming conventions:
       * "Sheet T1" → "SA.XXX.TEC" (technical sheet)
       * "Sheet 99" (where 99 is any 2-digit number) → "SA.XXX.YY"
         Examples: "Sheet 01" → "SA.010.01", "Sheet 12" → "SA.2152.12"
   * Fill A2 and B2 cells in each SA.*.* sheet:
     - Base text: Same value as D2/E2 from "identification" sheet (form description)
     - Conditional suffixes:
       a) "Sheet T1" → always add " - Feuillet technique"
       b) "Sheet 99" pattern:
          - If only 1 sheet OR (2 sheets with one being T1) → no suffix
          - If 2+ sheets matching "Sheet 99" → add " - {sheet_label_value}"
            where sheet_label_value is from the SECOND cell to the right of "Sheet label" 
            text in the input sheet (typically AJ42 when "Sheet label" is at AH42)
   * Map relevant data from input "Sheet .." sheets to output SA.*.* sheets:
     - Data mapping logic differs based on formCode length (3 vs 4 digits)
     - Extract and transform data according to specific mapping rules for each form type
   * All sheets from the input file will be included in the single output Excel file
   * Copy the template file and rename using pattern: pure-form-bnb{formGroupCode}-SA.XXX-(SA.XXX).xlsm
   * In the new file, copy (with data validation) and rename the "code" sheet for each input sheet
        ==> each initial "code" sheet will be called table
               
3. Each table in the generated workbook : 

   * Get the header of the columns (color-based header extraction strategy):
        * Extract all header rows in the source sheet that have the same fill color as the "Columns" cell.
        * Capture them exactly as they appear in the Excel file—merged/unmerged cells, blank cells, border formatting, etc.
        * Prepare them to be pasted later into the target sheet.
   * In the table sheet, we need to have the same number of cells starting with "columnLabel" as we have #rowsForColumns
        * Ensure the target sheet’s header block (where "columnLabel 1", "columnLabel 2", etc. are) matches the number of header rows extracted from the source sheet.
        * Unmerge first the region L15:L17.
        * Insert or remove rows as needed to achieve this.
        * Re-merge L15 downwards to include all consecutive empty cells.
        * Replace the columnLabel cell values so they are sequentially numbered (columnLabel 1, 2, 3, etc.).
        * Copy the fill color and styling (font, border, alignment and data validation) so the target sheet visually matches the source.
   * Paste the selection of cells representing the header of the columns from cell N15 point of the table without copying the format.
  

   * Get the header of the rows:
        * Extract all header in the source sheet that have the same fill color as the "Rows" cell and that are in column B.
        * Capture indent level for each row header in column B.
        * Paste the selection of cells in column M of the table just below the cell "rowNumbering" without copying the format.
        * Paste indent in column A below "level", same row as header, but with specific logic :
             * If indent is 0, level → 1
             * If indent is 2, level → 2
             * If indent is 4, level → 3
             * If indent is 6, level → 4
             * If indent is 8, level → 5   
             * If indent is 10, level → 6   
        * Extract all header in the source sheet that have the same fill color as the "Rows" cell and that are in column C.
        * Paste the selection of cells in column L of the table just below the cell "rowLabelColumnNumbering" without copying the format.
        * By the way, copy the entire row’s border style and alignments from the first pasted row and apply it to all cells in each pasted row (covering the whole row).


4. Post-processing
        * Border
             * Starts from the top-left cell (L15).
             * Determines the height and width of the filled area (the table region).
             * Draws bottom and right borders for the region to complete the rectangle. 

        * Yellow, grey (equivalent to adds-in "Cell Type" = "Manual" and "Not Required") and format of cells 

                * In the target worksheet, apply by default a yellow background color with reference: FFFFFFCC 
                     * to all cells located in the rectangle from the top-left cell starting "rowLabelColumnNumbering" (and +1 for row and column) till a heavy border of the rectangle found is found on a cell on the right and bottom corner.
                * In the target worksheet, apply a grey background color with reference: 808080 
                     * wherever from the source worksheet, the cell value are equals to # or % or €£$ 
                     * find position within the internal table thanks to rowLabelColumnNumbering value should be the same for the column coordinate and rowNumbering for the row coordinate.
                * Apply same process, to change the format cells in the target worksheet if from the source worksheet the cell contains these symbols : 
                     * "€£$" or "mi" ==> Number, with 2 decimals and with thousand separator
                     * "ri" ==> Number with 4 decimals and with thousand separator
                     * "ii" or "#" ==> Number without decimal with thousand separator
                     * "pi" or "%" ==> personalized and type # ##0,00%
                     * "di" or "date" ==> date and type "yyyy-mm-dd"

        * Handle open rows : Look for 'Open Rows' entries and add '999' under rowNumbering
                // TODO in the future : handle as well the <key value> so that it puts mandatory at yes in the column
                          
        * Handle asbtract rows
            * If a row has no value in its "rowNumbering" column,
            * it marks the "abstractRow" column as "yes",
            * extracts a numeric prefix from the "rowLabelColumnNumbering" cell (if present) and uses it as the new "rowNumbering" value,
            * and fills all cells to the right of "rowNumbering" (excluding the next column!) with a grey fill (808080).

        * Merged cells within the rows of the worksheet (consecutive partial rows from coordinate N15) till there is at one cell in the row starting with "columnLabel" in column M
            * Do all vertical Merging first (Top to Bottom)
                * If a cell is empty AND the respective cell above has no bottom border ==> Merge the empty cels with the one above it (vertical span)
                * Apply this logic for all the consecutives vertical cells that could be empty
            * Then, Horizontal Merging (Left to Right)
                * Case 1: Empty cell, and the cell to the left is not empty and not already merged ==> Merge both cells horizontally
                * Case 2: Empty cell, and the cell to the left is already merged ==> Copy the value of the left cell into the current cell (no merge)
                * Apply this logic for all the consecutives horizontal cells that could be empty, but stop when a large border is found on the right side of the cell

        * Remove the extra right columns after the end of the table borders

        * Removes empty rows (both the cell under "rowNumbering" is empty and the cell under "rowLabelColumnNumbering" is empty)
            * Before deleting, it copies the bottom border of that row up to the row above (to preserve the table edge visually).

        * Just before finalizing on this Excel, reopen the file for border compatibility with win32com on Windows as openpyxl known issue with border        

 
5. Fill the identification sheet in the generated workbook : 
   * Find the cell containing "Description (form label)" in the "Properties" sheet of the input workbook
   * Get the value from the first cell to the right of it (typically C5)
   * Copy this value to D2 cell and E2 cell of "identification" sheet from the generated workbook
   * In the G2 cell, write "direct"

6. The labelTranslation sheet in the output file remains unchanged from the template (no transformations)

7. The final code must:
   * Include robust error handling
   * Provide detailed debugging information : there should be a debug mode that can be activated and asked to the user to provide maximum of information at each step
   * Properly handle file paths with spaces
   * Not alter any original files



"""
# Standard library imports
import os
import re
import shutil
import logging
import sys
import datetime
from copy import copy

# Third-party imports
import openpyxl
from openpyxl.styles import Border, Side, PatternFill, Alignment
from openpyxl.styles.borders import BORDER_THIN
from openpyxl.utils import get_column_letter, range_boundaries
from openpyxl.cell.cell import MergedCell
from openpyxl.worksheet.cell_range import CellRange
from tkinter import Tk, filedialog, messagebox, simpledialog, Label, Entry, Button, ttk
import tkinter as tk
import win32com.client as win32
import win32com.client.gencache

# Import the Excel file logger
sys.path.append(os.path.join(os.path.dirname(__file__), '..', 'common'))
try:
    import excel_logger
    EXCEL_LOGGER_AVAILABLE = True
except ImportError:
    EXCEL_LOGGER_AVAILABLE = False

###########################################################################################################################
#                                                                                                                         #  
#      1st part of this task: the user input handling, directory preparation, and general setup (like debug mode)         #  
#                                                                                                                         #  
# User input for file path and debug mode                                                                                 #  
# Output directory creation                                                                                               #  
# Template file copying function                                                                                          #  
# Root name extraction function                                                                                           #  
###########################################################################################################################


# Global debug mode variable
DEBUG_MODE = False

# Logging configuration function
def configure_logging():
    global DEBUG_MODE
    level = logging.DEBUG if DEBUG_MODE else logging.INFO

    # Clear any existing handlers
    for handler in logging.root.handlers[:]:
        logging.root.removeHandler(handler)
    
    logging.basicConfig(
        level=level,
        format='[%(levelname)s] %(message)s',
        handlers=[logging.StreamHandler()]
    )
      

def get_user_input():
    root = Tk()
    root.withdraw()  # Hide the main window

    # 🟩 Create input subfolder if it doesn't exist
    script_dir = os.path.dirname(os.path.abspath(__file__))
    input_dir = os.path.join(script_dir, "input")
    if not os.path.exists(input_dir):
        os.makedirs(input_dir)
        logging.debug(f"   Created input directory: {input_dir}")

    # 🟩 Select input file starting from the input subfolder
    source_path = filedialog.askopenfilename(
        title="Select input Excel file to process",
        initialdir=input_dir,
        filetypes=[("Excel files", "*.xlsx *.xlsm *.xls")]
    )
    if not source_path:
        messagebox.showerror("Error", "No file selected. Exiting.")
        exit(1)

    global DEBUG_MODE
    DEBUG_MODE = messagebox.askyesno("Debug Mode", "Enable debug mode?")
    configure_logging()
    
    # 🟩 Extract formCode from input filename (needed for conversion dialog)
    input_filename = os.path.basename(source_path)
    base_name = os.path.splitext(input_filename)[0]
    match = re.search(r'FRM_(\d{3,4})', base_name)
    if match:
        form_code = match.group(1)
        logging.info(f"   Extracted formCode: {form_code} from filename: {input_filename}")
    else:
        logging.warning(f"   Could not extract formCode from filename '{input_filename}'. Using empty string.")
        form_code = ""
    
    # 🟩 Determine formCode length for data mapping
    # The length determines different data mapping strategies:
    # - length 3: formCode like 010, 020, etc. (3 digits)
    # - length 4: formCode like 2152, 1234, etc. (4 digits)
    form_code_length = len(form_code) if form_code else 0
    logging.debug(f"   formCode length: {form_code_length}")
    
    # 🟩 Determine if 5/6-digit rowID conversion is needed
    # - If form_code_length = 4, conversion is always "N"
    # - Otherwise, ask the user via dialog
    if form_code_length == 4:
        conversion_6_4_digits = "N"
        logging.debug(f"   formCode length is 4, conversion_6_4_digits automatically set to 'N'")
    else:
        # Ask user if they want to convert 5/6-digit rowIDs to 4-digit rowIDs
        from tkinter import ttk
        conversion_window = Tk()
        conversion_window.title("RowID Conversion")
        
        Label(conversion_window, text="Would you like to convert 5/6-digit rowID's into 4-digit rowID's?").grid(row=0, column=0, columnspan=2, padx=10, pady=10)
        
        conversion_var = tk.StringVar(conversion_window)
        conversion_var.set("No")  # Default value
        
        conversion_dropdown = ttk.Combobox(conversion_window, textvariable=conversion_var, values=["Yes", "No"], state="readonly", width=10)
        conversion_dropdown.grid(row=1, column=0, columnspan=2, padx=10, pady=10)
        
        def submit_conversion():
            conversion_window.quit()
            conversion_window.destroy()
        
        Button(conversion_window, text="OK", command=submit_conversion).grid(row=2, column=0, columnspan=2, pady=10)
        
        conversion_window.mainloop()
        
        conversion_6_4_digits = "Y" if conversion_var.get() == "Yes" else "N"
        logging.debug(f"   User selected conversion: {conversion_var.get()} -> conversion_6_4_digits = '{conversion_6_4_digits}'")

    # 🟩 Create a pop-up to collect all inputs together
    def ask_multiple_fields():
        # Create a new dialog window
        input_window = Tk()
        input_window.title("Form Details")     

        entries = {}

        labels = [
            ("formGroupCode (optional)", False),
            ("version (optional)", False),
            ("validityStartingDate (optional)", False),
            ("linkedDomain (optional)", False),
            ("datatype (optional)", False),
            ("roundingUnit (optional)", False)
        ]

        for idx, (label_text, required) in enumerate(labels):
            label = Label(input_window, text=label_text)
            label.grid(row=idx, column=0, sticky="e")
            entry = Entry(input_window, width=50)
            entry.grid(row=idx, column=1)
            entries[label_text] = entry

        # 🟩 Button to submit
        def submit():
            # Collect all fields
            result = {
                "formGroupCode": entries["formGroupCode (optional)"].get().strip() or "schemaaXYZ",
                "version": entries["version (optional)"].get().strip() or "2026.1.0",
                "validityStartingDate": entries["validityStartingDate (optional)"].get().strip() or "01/01/2026",
                "linkedDomain": entries["linkedDomain (optional)"].get().strip() or "bebnb",
                "datatype": entries["datatype (optional)"].get().strip() or "bebnb.genericPositiveAmount",
                "roundingUnit": entries["roundingUnit (optional)"].get().strip() or "thousands"
            }
            input_window.result = result
            input_window.quit()  # <--- This stops the mainloop!
            input_window.destroy()

        submit_btn = Button(input_window, text="Submit", command=submit)
        submit_btn.grid(row=len(labels), column=0, columnspan=2, pady=10)

        input_window.mainloop()
        return input_window.result

    # 🟩 Call the function to show the window
    user_inputs = ask_multiple_fields()
    logging.debug(f"   user_inputs captured: {user_inputs}")

    # 🟩 Return everything including formCode length and conversion flag
    return source_path, form_code, form_code_length, conversion_6_4_digits, user_inputs["formGroupCode"], user_inputs["version"], user_inputs["validityStartingDate"], user_inputs["linkedDomain"], user_inputs["datatype"], user_inputs["roundingUnit"]



def create_output_directory(source_path):
    script_dir = os.path.dirname(os.path.abspath(__file__))
    output_dir = os.path.join(script_dir, "generated_output")
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
        logging.info(f"Created output directory: {output_dir}")
    else:
        logging.debug(f"   Output directory already exists: {output_dir}")
    return output_dir


def copy_template_file(input_filename, output_dir, form_code, form_group_code):
    # Get the directory where this script is located
    script_dir = os.path.dirname(os.path.abspath(__file__))
    template_path = os.path.join(script_dir, "pure-form-regulatoryOutputCode-formCode-(shortName).xlsm")
    template_path = os.path.abspath(template_path)  # Resolve any .. in the path
    
    # Extract numeric pattern from input filename (FRM_XXX.xlsm or FRM_YYYY.xlsm)
    # Expected format: FRM_XXX or FRM_YYYY where XXX is 3 digits, YYYY is 4 digits
    base_name = os.path.splitext(input_filename)[0]
    match = re.search(r'FRM_(\d{3,4})', base_name)
    if match:
        numeric_pattern = match.group(1)
        sa_code = f"SA.{numeric_pattern}"
        # Use concatenation of "bnb" and formGroupCode instead of form_code
        target_filename = f"pure-form-bnb{form_group_code}-{sa_code}-({sa_code}).xlsm"
    else:
        # Fallback if pattern doesn't match expected format
        logging.warning(f"Input filename '{input_filename}' doesn't match expected pattern FRM_XXX or FRM_YYYY. Using fallback naming.")
        target_filename = f"pure-form-{form_code}-{base_name}-({base_name.replace('_', ' ')}).xlsm"
    
    target_path = os.path.join(output_dir, target_filename)
    shutil.copy2(template_path, target_path)

    logging.debug(f"   Copied template from: {template_path}")
    logging.debug(f"   Copied template to: {target_path}")
    
    # Initialize Excel file logging if available
    dedicated_logger = None
    if EXCEL_LOGGER_AVAILABLE:
        try:
            dedicated_logger = excel_logger.setup_excel_file_logging(target_path, "forms_from_input")
            excel_logger.log_excel_step(target_path, "Template copied", 
                                      f"Template: {template_path} -> Target: {target_path}")
        except Exception as e:
            logging.warning(f"Error initializing Excel logging: {e}")
    
    # Return sa_code along with target_path and dedicated_logger
    return target_path, dedicated_logger, sa_code if match else None


def extract_root_name(sheet_name):
    if sheet_name.count('.') == 1:
        return sheet_name
    elif sheet_name.count('.') == 2:
        return sheet_name.split('.')[0] + '.' + sheet_name.split('.')[1]
    else:
        return sheet_name  # fallback


###############################################################################################################################################################
#                                                                                                                                                             #  
#     2nd part: handling the source workbook to identify sheets, skip unwanted ones, group them by root name, and create the target workbooks                 #  
#                                                                                                                                                             #    
# Gets relevant sheets, skipping TOC, K_00.03, and Entry points.                                                                                              #  
# Groups sheets by root name for those with two '.'                                                                                                           #  
# Copies template files and duplicates/renames the "code" sheet for each entry in a group.                                                                    #     
#                                                                                                                                                             #  
###############################################################################################################################################################

def get_relevant_sheets(workbook):
    sheets = workbook.sheetnames
    filtered_sheets = []
    for sheet in sheets:
        # Only keep sheets that start with "Sheet"
        if sheet.startswith("Sheet"):
            filtered_sheets.append(sheet)
            logging.debug(f"   Retained sheet: {sheet}")
        else:
            logging.debug(f"   Skipped sheet: {sheet}")
    return filtered_sheets


def extract_internal_tables(source_wb, sheet_list):
    """
    Extract internal tables from sheets (except "Sheet T1").
    
    For each sheet, extracts all data between rows where column A contains:
    - "HNRB" (start marker - included)
    - "HNRE" (end marker - included)
    
    The internal table structure:
    - First column "rowNumbering": extracted from column AG (column 33) - contains Row ID
      Note: Row ID location is variable and not always labeled, so we use column AG consistently
    - Second column "rowLabelColumnNumbering": extracted from column AH (column 34)
    - Third column "feedingRuleSource": extracted from column AE (column 31) - Line allocation formula technical
      Only included if first non-blank character is "=", otherwise left empty
    - Fourth column "abstractRow": set to "yes" if column A contains "I", otherwise empty
    - Fifth column "listBreakdown": set to "yes" if column S (List Breakdown) contains non-whitespace value
      Used to identify list tables (e.g., 39XX, 8091)
    - Remaining columns contain the full row data
    
    Args:
        source_wb: Source workbook
        sheet_list: List of sheet names to process
    
    Returns:
        dict: Dictionary mapping sheet names to their internal table data
              {sheet_name: [{'rowNumbering': value, 'rowLabelColumnNumbering': value, 
                             'feedingRuleSource': value, 'abstractRow': value, 'listBreakdown': value,
                             'row_num': row_num, 'row_data': [...]}, ...]}
    """
    internal_tables = {}
    
    for sheet_name in sheet_list:
        # Skip "Sheet T1" - it doesn't need an internal table
        if sheet_name == "Sheet T1":
            logging.debug(f"   Skipping internal table extraction for '{sheet_name}' (technical sheet)")
            continue
        
        source_sheet = source_wb[sheet_name]
        start_row = None
        end_row = None
        
        # Find HNRB (start marker) and HNRE (end marker) in column A
        for row in source_sheet.iter_rows(min_col=1, max_col=1):
            cell = row[0]
            if cell.value and "HNRB" in str(cell.value):
                start_row = cell.row
                logging.debug(f"   Found HNRB at row {start_row} in '{sheet_name}'")
            if cell.value and "HNRE" in str(cell.value):
                end_row = cell.row
                logging.debug(f"   Found HNRE at row {end_row} in '{sheet_name}'")
                break  # Stop after finding HNRE
        
        # Extract data if both markers were found
        if start_row and end_row:
            sheet_data = []
            for row_num in range(start_row, end_row + 1):
                # Extract rowNumbering from column AG (column 33)
                row_numbering = source_sheet.cell(row=row_num, column=33).value
                
                # Extract rowLabelColumnNumbering from column AH (column 34)
                row_label_col_numbering = source_sheet.cell(row=row_num, column=34).value
                
                # Extract feedingRuleSource from column AE (column 31)
                # Only include if first non-blank character is "="
                ae_value = source_sheet.cell(row=row_num, column=31).value
                feeding_rule_source = ""
                if ae_value:
                    # Convert to string and strip leading/trailing whitespace
                    ae_str = str(ae_value).strip()
                    # Check if first non-blank character is "="
                    if ae_str and ae_str[0] == "=":
                        feeding_rule_source = ae_value
                
                # Extract abstractRow from column A (column 1)
                # Set to "yes" if column A contains "I"
                a_value = source_sheet.cell(row=row_num, column=1).value
                abstract_row = ""
                if a_value and "I" in str(a_value):
                    abstract_row = "yes"
                
                # Extract listBreakdown from column S (column 19) - List Breakdown
                # Set to "yes" if column S contains non-whitespace value
                s_value = source_sheet.cell(row=row_num, column=19).value
                list_breakdown = ""
                if s_value and str(s_value).strip():
                    list_breakdown = "yes"
                
                # Extract full row data
                row_data = []
                for cell in source_sheet[row_num]:
                    row_data.append(cell.value)
                
                # Store structured data
                sheet_data.append({
                    'rowNumbering': row_numbering,
                    'rowLabelColumnNumbering': row_label_col_numbering,
                    'feedingRuleSource': feeding_rule_source,
                    'abstractRow': abstract_row,
                    'listBreakdown': list_breakdown,
                    'row_num': row_num,
                    'row_data': row_data
                })
            
            internal_tables[sheet_name] = sheet_data
            logging.debug(f"   Extracted {len(sheet_data)} rows from '{sheet_name}' (rows {start_row}-{end_row})")
            
            # Count rows with feeding rules, abstract rows, and list breakdown
            feeding_rule_count = sum(1 for row in sheet_data if row['feedingRuleSource'])
            abstract_row_count = sum(1 for row in sheet_data if row['abstractRow'] == "yes")
            list_breakdown_count = sum(1 for row in sheet_data if row['listBreakdown'] == "yes")
            logging.debug(f"   First row: rowNumbering={sheet_data[0]['rowNumbering']}, rowLabelColumnNumbering={sheet_data[0]['rowLabelColumnNumbering']}" if sheet_data else "   No data extracted")
            logging.debug(f"   Rows with feeding rules: {feeding_rule_count}/{len(sheet_data)}")
            logging.debug(f"   Abstract rows: {abstract_row_count}/{len(sheet_data)}")
            logging.debug(f"   List breakdown rows: {list_breakdown_count}/{len(sheet_data)}")
        else:
            logging.warning(f"   Could not find HNRB/HNRE markers in '{sheet_name}'. Start: {start_row}, End: {end_row}")
            internal_tables[sheet_name] = []
    
    return internal_tables


def extract_column_internal_tables(source_wb, relevant_sheets):
    """
    Extract column-based internal tables from the source workbook.
    
    For each sheet starting with "Sheet" (except "Sheet T1"), create an internal table
    that will be used to fill in the columns part of sheet SA.*.* in the output file.
    
    The internal table contains data for all columns between start and end (inclusive):
    - start: In the line where column A contains "BC" (Column id line), find the first
             column whose value matches the pattern of 2, 3, or 4 consecutive digits
             (excludes alphabetical values like "aa", "aaa", or dummy values like "1")
    - end: In the same line, find the last column matching the same pattern
    
    Args:
        source_wb: The source workbook (openpyxl)
        relevant_sheets: List of sheet names to process
    
    Returns:
        dict: Column internal tables {sheet_name: column_data_dict}
              where column_data_dict = {
                  'bc_row': row number containing "BC",
                  'start_col': first column with 2-4 digit pattern,
                  'end_col': last column with 2-4 digit pattern,
                  'column_range': list of column letters/numbers in range
              }
    """
    import re
    column_internal_tables = {}
    
    logging.info("🔧  Extracting column-based internal tables from sheets")
    
    for sheet_name in relevant_sheets:
        # Skip "Sheet T1" (technical sheet)
        if sheet_name == "Sheet T1":
            logging.debug(f"   Skipping column extraction for technical sheet: '{sheet_name}'")
            continue
        
        source_sheet = source_wb[sheet_name]
        
        # Find the row where column A contains "BC"
        bc_row = None
        for row in source_sheet.iter_rows(min_row=1, max_row=200, min_col=1, max_col=1):
            cell_value = str(row[0].value).strip() if row[0].value else ""
            if "BC" in cell_value:
                bc_row = row[0].row
                logging.debug(f"   Found 'BC' at row {bc_row} in column A of '{sheet_name}'")
                break
        
        if not bc_row:
            logging.warning(f"   Could not find 'BC' in column A of '{sheet_name}'. Skipping column extraction.")
            column_internal_tables[sheet_name] = None
            continue
        
        # Find the row where column A contains "CE" (Column allocation formula technical)
        ce_row = None
        for row in source_sheet.iter_rows(min_row=1, max_row=200, min_col=1, max_col=1):
            cell_value = str(row[0].value).strip() if row[0].value else ""
            if "CE" in cell_value:
                ce_row = row[0].row
                logging.debug(f"   Found 'CE' at row {ce_row} in column A of '{sheet_name}'")
                break
        
        if not ce_row:
            logging.warning(f"   Could not find 'CE' in column A of '{sheet_name}'. Skipping column extraction.")
            column_internal_tables[sheet_name] = None
            continue
        
        # In the BC row, find columns with 2-4 consecutive digits pattern
        digit_pattern = re.compile(r'^\d{2,4}$')  # Matches exactly 2, 3, or 4 digits
        start_col = None
        end_col = None
        
        # Search through columns (up to column ZZ, which is reasonable)
        for col_idx in range(1, 500):  # Check up to column 500
            cell_value = source_sheet.cell(row=bc_row, column=col_idx).value
            cell_value_str = str(cell_value).strip() if cell_value else ""
            
            # Check if the value matches the pattern
            if digit_pattern.match(cell_value_str):
                if start_col is None:
                    start_col = col_idx
                    logging.debug(f"   Found start column at {openpyxl.utils.get_column_letter(col_idx)} (column {col_idx}), value: '{cell_value_str}'")
                end_col = col_idx  # Keep updating end_col as we find more matches
        
        if start_col and end_col:
            logging.debug(f"   Column range: {openpyxl.utils.get_column_letter(start_col)} to {openpyxl.utils.get_column_letter(end_col)} (columns {start_col}-{end_col})")
            
            # Create column range list
            column_range = list(range(start_col, end_col + 1))
            
            # Extract the Column IDs (rowLabelColumnNumbering) from bc_row
            row_label_column_numbering = []
            for col_idx in column_range:
                cell_value = source_sheet.cell(row=bc_row, column=col_idx).value
                cell_value_str = str(cell_value).strip() if cell_value else ""
                row_label_column_numbering.append(cell_value_str)
            
            # Extract the Column allocation formula technical (feedingRuleSource) from ce_row
            # Only include if first non-blank character is "="
            feeding_rule_source = []
            for col_idx in column_range:
                cell_value = source_sheet.cell(row=ce_row, column=col_idx).value
                cell_value_str = ""
                
                if cell_value:
                    # Convert to string and strip leading/trailing whitespace
                    value_str = str(cell_value).strip()
                    # Check if first non-blank character is "="
                    if value_str and value_str[0] == "=":
                        cell_value_str = cell_value  # Keep original value (could be formula)
                
                feeding_rule_source.append(cell_value_str)
            
            column_internal_tables[sheet_name] = {
                'bc_row': bc_row,
                'ce_row': ce_row,
                'start_col': start_col,
                'end_col': end_col,
                'column_range': column_range,
                'column_count': len(column_range),
                'rowLabelColumnNumbering': row_label_column_numbering,
                'feedingRuleSource': feeding_rule_source
            }
            
            logging.debug(f"   Extracted column internal table for '{sheet_name}': {len(column_range)} columns")
            logging.debug(f"   Column IDs (rowLabelColumnNumbering): {row_label_column_numbering}")
            logging.debug(f"   Column allocation formulas (feedingRuleSource): {feeding_rule_source}")
        else:
            logging.warning(f"   Could not find valid column range in '{sheet_name}' (BC row: {bc_row})")
            column_internal_tables[sheet_name] = None
    
    logging.info(f"   Extracted column internal tables from {len([v for v in column_internal_tables.values() if v])} sheets")
    return column_internal_tables


def create_cross_internal_tables(internal_tables, column_internal_tables, form_code, form_code_length):
    """
    Create cross_internal_tables combining row and column data to map all cells.
    
    Creates a 2-dimensional internal table for each sheet:
    - First dimension: Row numbers from internal_tables (row_num between start_row and end_row)
    - Second dimension: Column letters from column_internal_tables (column_range)
    
    For form_code_length = 3: Skip rows where trimmed rowNumbering is empty (cleanup rule)
    
    For formCode 010, 020, 040, 050:
    - Copy "Sheet 01" data to "Sheet T1"
    - For formCode 020: Remove row_num 228 (rowNumbering 286T) from "Sheet T1"
    
    Args:
        internal_tables: Dictionary of row-based internal tables {sheet_name: [...]}
        column_internal_tables: Dictionary of column-based internal tables {sheet_name: {...}}
        form_code: The form code (e.g., "010", "2142")
        form_code_length: Length of form code (3 or 4)
    
    Returns:
        dict: Cross internal tables {sheet_name: {cell_ref: {...}, ...}}
    """
    from openpyxl.utils import get_column_letter
    
    logging.info("🔧  Creating cross internal tables")
    
    cross_internal_tables = {}
    
    # Process each sheet (except "Sheet T1" initially)
    for sheet_name in internal_tables.keys():
        # Skip "Sheet T1" for now (will be handled later for specific formCodes)
        if sheet_name == "Sheet T1":
            continue
        
        # Get row and column data
        row_data = internal_tables.get(sheet_name, [])
        column_data = column_internal_tables.get(sheet_name)
        
        if not row_data or not column_data:
            logging.debug(f"   Skipping '{sheet_name}': missing row or column data")
            continue
        
        # Get column range
        column_range = column_data.get('column_range', [])
        if not column_range:
            logging.debug(f"   Skipping '{sheet_name}': empty column range")
            continue
        
        sheet_cross_table = {}
        valid_row_count = 0
        
        # Process each row
        for row_entry in row_data:
            row_num = row_entry.get('row_num')
            row_numbering = row_entry.get('rowNumbering')
            
            # For form_code_length = 3: Skip rows where trimmed rowNumbering is empty
            if form_code_length == 3:
                row_numbering_str = str(row_numbering).strip() if row_numbering else ""
                if not row_numbering_str:
                    logging.debug(f"   Sheet '{sheet_name}', row {row_num}: Skipping (empty rowNumbering for form_code_length=3)")
                    continue
            
            valid_row_count += 1
            
            # Process each column in the range
            for col_idx in column_range:
                col_letter = get_column_letter(col_idx)
                cell_ref = f"{col_letter}{row_num}"
                
                # Store cell information
                sheet_cross_table[cell_ref] = {
                    'row_num': row_num,
                    'col_num': col_idx,
                    'col_letter': col_letter,
                    'cell_ref': cell_ref,
                    'rowNumbering': row_numbering,
                    'rowLabelColumnNumbering': row_entry.get('rowLabelColumnNumbering'),
                    'row_data': row_entry
                }
        
        cross_internal_tables[sheet_name] = sheet_cross_table
        logging.debug(f"   Created cross internal table for '{sheet_name}': {valid_row_count} rows × {len(column_range)} columns = {len(sheet_cross_table)} cells")
    
    # Special handling for formCode 010, 020, 040, 050
    # Copy "Sheet 01" to "Sheet T1"
    if form_code in ["010", "020", "040", "050"]:
        if "Sheet 01" in cross_internal_tables:
            logging.info(f"   FormCode {form_code}: Copying cross internal table from 'Sheet 01' to 'Sheet T1'")
            
            # Copy Sheet 01 data
            sheet_01_data = cross_internal_tables["Sheet 01"]
            sheet_t1_data = {}
            
            # For formCode 020: Exclude row_num 228 (rowNumbering 286T)
            if form_code == "020":
                excluded_count = 0
                for cell_ref, cell_data in sheet_01_data.items():
                    if cell_data['row_num'] == 228:
                        # Skip this row for Sheet T1
                        excluded_count += 1
                        logging.debug(f"      Excluding row 228 (rowNumbering 286T) for 'Sheet T1'")
                    else:
                        sheet_t1_data[cell_ref] = cell_data.copy()
                
                logging.debug(f"      Sheet T1 has {len(sheet_t1_data)} cells ({excluded_count} excluded from Sheet 01)")
            else:
                # For other formCodes (010, 040, 050): Copy all data
                for cell_ref, cell_data in sheet_01_data.items():
                    sheet_t1_data[cell_ref] = cell_data.copy()
                
                logging.debug(f"      Sheet T1 has {len(sheet_t1_data)} cells (copied from Sheet 01)")
            
            cross_internal_tables["Sheet T1"] = sheet_t1_data
        else:
            logging.warning(f"   FormCode {form_code}: 'Sheet 01' not found, cannot create 'Sheet T1' cross table")
    
    logging.info(f"   Created cross internal tables for {len(cross_internal_tables)} sheets")
    return cross_internal_tables


def process_cell_colors_and_formulas(cross_internal_tables, source_wb, form_code):
    """
    Process cell colors and formulas for all cells in cross_internal_tables.
    
    For each cell:
    - Determine cell_color based on input file cell color:
      * White (#FFFFFF or no fill) → Grey (#808080)
      * Red (#FF0000) → Pink (#FF99CC)
      * Pink (#FF99CC) for "Sheet 01" of formCode 010 only → Pink (#FF99CC)
      * All other cases → Yellow (#FFFFCC)
    
    - Determine cell_formula (only if cell value starts with "="):
      * Remove all spaces
      * Replace L999999 (1-6 digits) with {r999999}
      * Convert C99 to C099
      * Replace C9999 (1-4 digits) with {c9999}
      * Add ' at beginning if not present
    
    Args:
        cross_internal_tables: Dictionary of cross internal tables {sheet_name: {cell_ref: {...}, ...}}
        source_wb: Source workbook
        form_code: The form code (e.g., "010", "2142")
    
    Returns:
        dict: Updated cross internal tables with cell_color and cell_formula attributes
    """
    import re
    
    logging.info("🔧  Processing cell colors and formulas")
    
    # Pattern to match L followed by 1 to 6 digits
    l_pattern = re.compile(r'L(\d{1,6})')
    # Pattern to match C followed by exactly 2 digits (not part of 3+ digits)
    c2_pattern = re.compile(r'C(\d{2})(?!\d)')
    # Pattern to match C followed by 1 to 4 digits
    c_pattern = re.compile(r'C(\d{1,4})')
    
    for sheet_name, sheet_cross_table in cross_internal_tables.items():
        # Determine source sheet name
        # For Sheet T1 in formCodes 010, 020, 040, 050, read from Sheet T1 (not Sheet 01)
        source_sheet_name = sheet_name
        
        # Check if source sheet exists
        if source_sheet_name not in source_wb.sheetnames:
            logging.warning(f"   Source sheet '{source_sheet_name}' not found in workbook, skipping color/formula processing")
            continue
        
        source_sheet = source_wb[source_sheet_name]
        
        logging.debug(f"   Processing {len(sheet_cross_table)} cells in '{sheet_name}' (reading from '{source_sheet_name}')")
        
        cells_with_color = 0
        cells_with_formula = 0
        
        for cell_ref, cell_data in sheet_cross_table.items():
            row_num = cell_data['row_num']
            col_num = cell_data['col_num']
            
            # Get the source cell
            source_cell = source_sheet.cell(row=row_num, column=col_num)
            
            # ===== DETERMINE CELL COLOR =====
            cell_color = "#FFFFCC"  # Default: Yellow
            
            # Get cell fill color
            fill = source_cell.fill
            hex_color = None
            
            # Indexed color mapping (Excel's default color palette)
            # Reference: https://docs.microsoft.com/en-us/office/troubleshoot/excel/determine-rgb-value-of-color
            indexed_colors = {
                0: "#000000",   # Black
                1: "#FFFFFF",   # White
                2: "#FF0000",   # Red
                3: "#00FF00",   # Green
                4: "#0000FF",   # Blue
                5: "#FFFF00",   # Yellow
                6: "#FF00FF",   # Magenta
                7: "#00FFFF",   # Cyan
                8: "#000000",   # Black
                9: "#FFFFFF",   # White
                10: "#FF0000",  # Red
                11: "#00FF00",  # Bright Green
                12: "#0000FF",  # Blue
                13: "#FFFF00",  # Yellow
                14: "#FF00FF",  # Magenta
                15: "#00FFFF",  # Cyan
                16: "#800000",  # Dark Red
                17: "#008000",  # Green
                18: "#000080",  # Dark Blue
                19: "#808000",  # Dark Yellow/Olive
                20: "#800080",  # Violet/Purple
                21: "#008080",  # Teal
                22: "#C0C0C0",  # Gray-25%
                23: "#808080",  # Gray-50%
                24: "#9999FF",  # Periwinkle
                25: "#993366",  # Plum
                26: "#FFFFCC",  # Ivory
                27: "#CCFFFF",  # Light Turquoise
                28: "#660066",  # Dark Purple
                29: "#FF8080",  # Coral
                30: "#0066CC",  # Ocean Blue
                31: "#CCCCFF",  # Ice Blue
                32: "#000080",  # Dark Blue
                33: "#FF00FF",  # Pink
                34: "#FFFF00",  # Yellow
                35: "#00FFFF",  # Turquoise
                36: "#800080",  # Violet
                37: "#800000",  # Dark Red
                38: "#008080",  # Teal
                39: "#0000FF",  # Blue
                40: "#00CCFF",  # Sky Blue
                41: "#CCFFFF",  # Light Turquoise
                42: "#CCFFCC",  # Light Green
                43: "#FFFF99",  # Light Yellow
                44: "#99CCFF",  # Pale Blue
                45: "#FF99CC",  # Rose/Pink
                46: "#CC99FF",  # Lavender
                47: "#FFCC99",  # Peach
                48: "#3366FF",  # Bright Blue
                49: "#33CCCC",  # Aqua
                50: "#99CC00",  # Lime
                51: "#FFCC00",  # Gold
                52: "#FF9900",  # Light Orange
                53: "#FF6600",  # Orange
                54: "#666699",  # Blue Gray
                55: "#969696",  # Gray-40%
                56: "#003366",  # Dark Teal
                57: "#339966",  # Sea Green
                58: "#003300",  # Dark Green
                59: "#333300",  # Olive Green
                60: "#993300",  # Brown
                61: "#993366",  # Plum
                62: "#333399",  # Indigo
                63: "#333333",  # Gray-80%
            }
            
            if fill and fill.start_color:
                # Check if this is an indexed color
                try:
                    if fill.start_color.indexed is not None and isinstance(fill.start_color.indexed, int):
                        indexed = fill.start_color.indexed
                        hex_color = indexed_colors.get(indexed)
                        if hex_color is None:
                            logging.debug(f"   Unknown indexed color {indexed} at {cell_ref}, defaulting to yellow")
                except (TypeError, ValueError, AttributeError):
                    pass  # Not an indexed color or error accessing it
                
                # Otherwise, try to get RGB value
                if hex_color is None and fill.start_color.rgb:
                    rgb = fill.start_color.rgb
                    # Handle RGB strings
                    if isinstance(rgb, str):
                        # Remove alpha channel (first 2 chars if 8 chars long)
                        if len(rgb) == 8:
                            rgb = rgb[2:]
                        hex_color = f"#{rgb.upper()}"
            
            # Apply color mapping based on determined hex_color
            if hex_color == "#FF0000":
                # Red → Pink (checked FIRST)
                cell_color = "#FF99CC"
                cells_with_color += 1
            elif hex_color == "#FF99CC" and sheet_name == "Sheet 01" and form_code == "010":
                # Special case: Pink in Sheet 01 of formCode 010
                cell_color = "#FF99CC"
                cells_with_color += 1
            elif hex_color in ["#FFFFFF", "#000000"]:
                # White or black → Grey
                cell_color = "#808080"
                cells_with_color += 1
            elif fill and fill.fill_type is None:
                # No fill type → Grey
                cell_color = "#808080"
                cells_with_color += 1
            elif hex_color is None and fill:
                # Has fill but couldn't determine color → Grey
                cell_color = "#808080"
                cells_with_color += 1
            # else: keep default yellow
            
            # ===== DETERMINE CELL FORMULA =====
            cell_formula = ""
            
            # Get cell value
            cell_value = source_cell.value
            if cell_value:
                value_str = str(cell_value).strip()
                
                # Check if value contains = (possibly after apostrophes)
                if value_str and '=' in value_str:
                    # Step 1: Remove apostrophes before first = sign
                    equals_pos = value_str.find('=')
                    if equals_pos > 0:
                        # Remove all ' characters before the =
                        before_equals = value_str[:equals_pos].replace("'", "")
                        value_str = before_equals + value_str[equals_pos:]
                    
                    # Now check if it starts with = after removing apostrophes
                    if value_str and value_str[0] == "=":
                        # Step 2: Remove all spaces
                        formula_str = value_str.replace(" ", "")
                        
                        # Step 3: Replace L999999 (1-6 digits) with {r999999}
                        formula_str = l_pattern.sub(r'{r\1}', formula_str)
                        
                        # Step 4: Convert C99 to C099 (add 0 between C and 2 digits)
                        formula_str = c2_pattern.sub(r'C0\1', formula_str)
                        
                        # Step 5: Replace C9999 (1-4 digits) with {c9999}
                        formula_str = c_pattern.sub(r'{c\1}', formula_str)
                        
                        # Step 6: Add apostrophe at beginning for data area formulas (to display as text)
                        if not formula_str.startswith("'"):
                            formula_str = "'" + formula_str
                        
                        cell_formula = formula_str
                        cells_with_formula += 1
            
            # Update cell data with new attributes
            cell_data['cell_color'] = cell_color
            cell_data['cell_formula'] = cell_formula
        
        logging.debug(f"   Sheet '{sheet_name}': {cells_with_color} cells with non-yellow color, {cells_with_formula} cells with formulas")
    
    logging.info(f"   Processed cell colors and formulas for {len(cross_internal_tables)} sheets")
    return cross_internal_tables


def cleanup_column_internal_tables(column_internal_tables):
    """
    Clean up column internal tables to ensure ColumnNumbering contains at least 3 digits.
    
    Rules:
    1. If rowLabelColumnNumbering has 2 digits: add a leading 0 to make it 3 digits
    2. If feedingRuleSource is filled in: convert patterns C99 to C099 (add 0 between C and 2 digits)
    
    Args:
        column_internal_tables: Dictionary of column tables {sheet_name: {...}}
    
    Returns:
        dict: Cleaned column internal tables
    """
    import re
    
    cleaned_tables = {}
    
    for sheet_name, sheet_data in column_internal_tables.items():
        if not sheet_data:
            cleaned_tables[sheet_name] = None
            continue
        
        # Get the lists from sheet_data
        row_label_col_numbering = sheet_data['rowLabelColumnNumbering']
        feeding_rule_source = sheet_data['feedingRuleSource']
        
        # Clean up rowLabelColumnNumbering: add leading 0 if only 2 digits
        cleaned_row_labels = []
        for col_id in row_label_col_numbering:
            col_id_str = str(col_id).strip() if col_id else ""
            if len(col_id_str) == 2:
                # Add leading 0
                col_id_str = "0" + col_id_str
                logging.debug(f"   Sheet '{sheet_name}': Added leading 0 to column ID '{col_id}' -> '{col_id_str}'")
            cleaned_row_labels.append(col_id_str)
        
        # Clean up feedingRuleSource: convert C99 patterns to C099
        cleaned_feeding_rules = []
        pattern = re.compile(r'C(\d{2})(?!\d)')  # Match C followed by exactly 2 digits (not part of 3+ digits)
        
        for formula in feeding_rule_source:
            formula_str = str(formula) if formula else ""
            if formula_str:
                # Replace C + 2 digits with C0 + 2 digits
                original = formula_str
                formula_str = pattern.sub(r'C0\1', formula_str)
                if original != formula_str:
                    logging.debug(f"   Sheet '{sheet_name}': Converted formula '{original}' -> '{formula_str}'")
            cleaned_feeding_rules.append(formula_str)
        
        # Create cleaned sheet data
        cleaned_tables[sheet_name] = {
            'bc_row': sheet_data['bc_row'],
            'ce_row': sheet_data['ce_row'],
            'start_col': sheet_data['start_col'],
            'end_col': sheet_data['end_col'],
            'column_range': sheet_data['column_range'],
            'column_count': sheet_data['column_count'],
            'rowLabelColumnNumbering': cleaned_row_labels,
            'feedingRuleSource': cleaned_feeding_rules
        }
    
    return cleaned_tables


def process_column_feeding_rules(column_internal_tables):
    """
    Process column feedingRuleSource values according to transformation rules.
    
    For each column where feedingRuleSource is filled in:
    1. Remove all spaces from feedingRuleSource
    2. Replace all patterns C9999 (where 9999 is 1-4 digits) with {c9999} 
       (lowercase c, wrapped in braces)
    3. Remove all apostrophes before the first = sign
    
    Args:
        column_internal_tables: Dictionary of column tables {sheet_name: {...}}
    
    Returns:
        dict: Processed column internal tables
    """
    import re
    
    processed_tables = {}
    
    # Pattern to match C followed by 1 to 4 digits
    c_pattern = re.compile(r'C(\d{1,4})')
    
    for sheet_name, sheet_data in column_internal_tables.items():
        if not sheet_data:
            processed_tables[sheet_name] = None
            continue
        
        # Get the feedingRuleSource list
        feeding_rule_source = sheet_data['feedingRuleSource']
        
        # Process each formula
        processed_feeding_rules = []
        for formula in feeding_rule_source:
            formula_str = str(formula) if formula else ""
            
            if formula_str:  # Only process if not empty
                original = formula_str
                
                # Step 1: Remove apostrophes before first = sign
                if '=' in formula_str:
                    equals_pos = formula_str.find('=')
                    if equals_pos > 0:
                        # Remove all ' characters before the =
                        before_equals = formula_str[:equals_pos].replace("'", "")
                        formula_str = before_equals + formula_str[equals_pos:]
                
                # Step 2: Remove all spaces
                formula_str = formula_str.replace(" ", "")
                
                # Step 3: Replace C9999 patterns with {c9999}
                # Replace C with lowercase c and wrap in braces
                formula_str = c_pattern.sub(r'{c\1}', formula_str)
                
                if original != formula_str:
                    logging.debug(f"   Sheet '{sheet_name}': Transformed column formula '{original}' -> '{formula_str}'")
            
            processed_feeding_rules.append(formula_str)
        
        # Create processed sheet data
        processed_tables[sheet_name] = {
            'bc_row': sheet_data['bc_row'],
            'ce_row': sheet_data['ce_row'],
            'start_col': sheet_data['start_col'],
            'end_col': sheet_data['end_col'],
            'column_range': sheet_data['column_range'],
            'column_count': sheet_data['column_count'],
            'rowLabelColumnNumbering': sheet_data['rowLabelColumnNumbering'],
            'feedingRuleSource': processed_feeding_rules
        }
    
    return processed_tables


def calculate_column_computation_rule_order_for_sheet(column_table_data, output_sheet, sheet_name):
    """
    Calculate computationRuleOrder for columns in a single sheet based on feeding rule dependencies.
    
    For each column with a feedingRuleSource:
    - Initialize with max(column I values from row 19+) + 1 (or 0 if all empty)
    - Extract all c{1-4 digits} patterns from feedingRuleSource
    - For each pattern, find the corresponding column and check if it has dependencies
    - Recursively calculate the depth and take the maximum across all branches
    
    Args:
        column_table_data: Dictionary with column data {'rowLabelColumnNumbering': [...], 'feedingRuleSource': [...], ...}
        output_sheet: The output worksheet to read max computationRuleOrder from column I
        sheet_name: Name of the sheet (for logging)
    
    Returns:
        dict: Column table data with computationRuleOrder added
    """
    import re
    
    def calculate_max_column_depth(feeding_rule, column_lookup, base_order, visited):
        """
        Recursively calculate the maximum depth of the column dependency tree.
        
        Args:
            feeding_rule: The feeding rule string containing {c...} patterns
            column_lookup: Dictionary mapping column ID to column data
            base_order: The base computationRuleOrder value
            visited: Set of visited column IDs to avoid cycles
        
        Returns:
            Maximum computationRuleOrder across all branches
        """
        # Extract all patterns c{1-4 digits} from feedingRuleSource
        pattern = r'c(\d{1,4})'
        matches = re.findall(pattern, feeding_rule)
        
        if not matches:
            # No dependencies, return base order
            return base_order
        
        max_child_order = base_order
        has_dependency = False
        
        for col_id in matches:
            # Avoid cycles
            if col_id in visited:
                continue
            
            # Find the corresponding column
            if col_id in column_lookup:
                child_feeding_rule = column_lookup[col_id]
                
                if child_feeding_rule:
                    child_feeding_rule_str = str(child_feeding_rule).strip()
                    
                    if child_feeding_rule_str:
                        has_dependency = True
                        # Mark as visited to avoid cycles
                        new_visited = visited.copy()
                        new_visited.add(col_id)
                        
                        # Recursively calculate order with +1 for this level
                        child_order = calculate_max_column_depth(
                            child_feeding_rule_str, 
                            column_lookup, 
                            base_order + 1, 
                            new_visited
                        )
                        max_child_order = max(max_child_order, child_order)
        
        # If we had dependencies, return the max child order
        # Otherwise return base order
        return max_child_order if has_dependency else base_order
    
    # Get the max computationRuleOrder from column I (row 19+)
    max_row_order = None
    for row in output_sheet.iter_rows(min_row=19, min_col=9, max_col=9):  # Column I is column 9
        cell_value = row[0].value
        if cell_value is not None:
            try:
                order_value = int(cell_value)
                if max_row_order is None or order_value > max_row_order:
                    max_row_order = order_value
            except (ValueError, TypeError):
                pass
    
    # Initialize base order: max + 1, or 0 if no values found
    base_order = (max_row_order + 1) if max_row_order is not None else 0
    
    logging.debug(f"   Sheet '{sheet_name}': Base column computationRuleOrder = {base_order} (max row order = {max_row_order})")
    
    # Build a lookup dictionary: column ID -> feedingRuleSource
    column_lookup = {}
    row_label_col_numbering = column_table_data['rowLabelColumnNumbering']
    feeding_rule_source = column_table_data['feedingRuleSource']
    
    for i, col_id in enumerate(row_label_col_numbering):
        if i < len(feeding_rule_source):
            column_lookup[col_id] = feeding_rule_source[i]
    
    # Calculate computationRuleOrder for each column
    computation_orders = []
    
    for i, formula in enumerate(feeding_rule_source):
        formula_str = str(formula) if formula else ""
        
        if formula_str:
            # Calculate order recursively
            col_id = row_label_col_numbering[i]
            visited = set([col_id])  # Start with current column as visited
            final_order = calculate_max_column_depth(formula_str, column_lookup, base_order, visited)
            computation_orders.append(final_order)
            logging.debug(f"   Sheet '{sheet_name}', column {col_id}: computationRuleOrder = {final_order}")
        else:
            # Empty formula, no order needed
            computation_orders.append("")
    
    # Create result with computationRuleOrder added
    result = {
        'bc_row': column_table_data['bc_row'],
        'ce_row': column_table_data['ce_row'],
        'start_col': column_table_data['start_col'],
        'end_col': column_table_data['end_col'],
        'column_range': column_table_data['column_range'],
        'column_count': column_table_data['column_count'],
        'rowLabelColumnNumbering': column_table_data['rowLabelColumnNumbering'],
        'feedingRuleSource': column_table_data['feedingRuleSource'],
        'computationRuleOrder': computation_orders
    }
    
    return result


def copy_column_data_to_output(table_sheet, column_table_data, sheet_name):
    """
    Copy column internal table data to the output SA.*.* sheet.
    
    For each column:
    - Copy feedingRuleSource to row 10, starting at column N (14)
    - Copy computationRuleOrder to row 12
    - Copy rowLabelColumnNumbering to row 18 (with Text format)
    
    Args:
        table_sheet: The output worksheet
        column_table_data: Dictionary with column data including computationRuleOrder
        sheet_name: Name of the sheet (for logging)
    """
    from openpyxl.styles import Alignment
    
    if not column_table_data:
        return
    
    # Get data from column internal table
    feeding_rules = column_table_data.get('feedingRuleSource', [])
    computation_orders = column_table_data.get('computationRuleOrder', [])
    column_ids = column_table_data.get('rowLabelColumnNumbering', [])
    
    # Starting column is N (column 14)
    start_col = 14
    
    logging.debug(f"   Copying column data for '{sheet_name}': {len(column_ids)} columns")
    
    for i, col_id in enumerate(column_ids):
        output_col = start_col + i
        
        # Row 10: feedingRuleSource (with wrap text and Text format)
        if i < len(feeding_rules) and feeding_rules[i]:
            cell = table_sheet.cell(row=10, column=output_col)
            cell.value = feeding_rules[i]
            cell.alignment = Alignment(wrap_text=True)
            # Set cell format to Text and data type to string
            cell.number_format = '@'
            cell.data_type = 's'  # Force string type to prevent formula interpretation
            logging.debug(f"      Column {output_col}: Row 10 = {feeding_rules[i]}")
        
        # Row 12: computationRuleOrder
        if i < len(computation_orders) and computation_orders[i] != "":
            cell = table_sheet.cell(row=12, column=output_col)
            cell.value = computation_orders[i]
            logging.debug(f"      Column {output_col}: Row 12 = {computation_orders[i]}")
        
        # Row 18: rowLabelColumnNumbering (with Text format)
        if col_id:
            cell = table_sheet.cell(row=18, column=output_col)
            # Convert to string to ensure Text format
            cell.value = str(col_id)
            # Set cell format to Text
            cell.number_format = '@'
            logging.debug(f"      Column {output_col}: Row 18 = {col_id} (Text format)")
    
    logging.info(f"   Copied column data for {len(column_ids)} columns to rows 10, 12, 18")


def fill_column_labels(table_sheet, column_count, form_code_length, sheet_name, source_sheet=None, column_start_col=None, column_table_data=None):
    """
    Fill columnLabel rows (15, 16, 17) based on form_code_length.
    
    For form_code_length = 3:
    - columnLabel 1 (row 15): Merge cells from column N to last column, 
      fill with "Schéma sur base territoriale et sociale - Etablissements de crédit Belge et Succursales d'EC de l'UE"
    - columnLabel 2 (row 16): Extract text from input sheet based on "Start reference" and "HT" markers
    
    For form_code_length = 4:
    - Determine row and column ranges based on "HT" marker and rowLabelColumnNumbering match
    - Fill columnLabels according to match/no-match logic
    
    Args:
        table_sheet: The output sheet to fill
        column_count: Number of columns in the output table
        form_code_length: Length of the form code (3 or 4)
        sheet_name: Name of the sheet (to check if it's .01 sheet)
        source_sheet: The source worksheet from input file (needed for columnLabel 2)
        column_start_col: The starting column index in source sheet for column data
        column_table_data: Dictionary with column internal table data (needed for form_code_length=4)
    
    Returns:
        Dict with border fix information, or None if no borders need fixing
        Format: {'row': 15, 'last_column': 22}
    """
    from openpyxl.styles import Alignment
    import re
    
    # Only process for form_code_length = 3 or 4
    if form_code_length not in [3, 4]:
        logging.debug(f"   Skipping columnLabel filling: form_code_length={form_code_length} (not 3 or 4)")
        return None
    
    # For form_code_length = 3: only process .01 sheets (base sheets, not TEC)
    # For form_code_length = 4: process all sheets
    if form_code_length == 3 and not sheet_name.endswith('.01'):
        logging.debug(f"   Skipping columnLabel filling: sheet {sheet_name} is not a .01 sheet (form_code_length=3)")
        return None
    
    if column_count == 0:
        logging.debug(f"   Skipping columnLabel filling: no columns to process")
        return None
    
    from openpyxl.utils import get_column_letter
    start_col = 14  # Column N
    last_col = start_col + column_count - 1  # Last column with data
    
    # ========================================================================
    # FORM_CODE_LENGTH = 4 LOGIC
    # ========================================================================
    if form_code_length == 4:
        logging.debug(f"   Filling columnLabel rows for form_code_length=4")
        logging.debug(f"   Column range: {start_col} to {last_col} ({column_count} columns)")
        
        if not source_sheet or column_start_col is None or column_table_data is None:
            logging.warning(f"   Missing required data for form_code_length=4, skipping columnLabel filling")
            return None
        
        # Step 1: Find HT row
        ht_row = None
        for row in source_sheet.iter_rows(min_col=1, max_col=1):
            cell = row[0]
            if cell.value and isinstance(cell.value, str) and "HT" in cell.value:
                ht_row = cell.row
                logging.debug(f"   Found 'HT' at row {ht_row}")
                break
        
        if ht_row is None:
            logging.warning(f"   Could not find 'HT' marker, skipping columnLabel filling")
            return None
        
        # Step 2: Get cell at (HT_row, last_column_of_range) and rowLabelColumnNumbering
        last_source_col = column_start_col + column_count - 1
        ht_cell_value = source_sheet.cell(row=ht_row, column=last_source_col).value
        ht_cell_str = str(ht_cell_value) if ht_cell_value else ""
        
        # Get rowLabelColumnNumbering from column_table_data (it's a list at the top level)
        # We need the LAST element in the list (index -1) which represents the last column's ID
        row_label_col_numbering = None
        if column_table_data and 'rowLabelColumnNumbering' in column_table_data:
            row_label_list = column_table_data['rowLabelColumnNumbering']
            if row_label_list and len(row_label_list) > 0:
                row_label_col_numbering = str(row_label_list[-1])  # Get last column's ID
        
        if row_label_col_numbering is None:
            logging.warning(f"   Could not find rowLabelColumnNumbering, skipping columnLabel filling")
            return None
        
        logging.debug(f"   HT cell value: '{ht_cell_str}'")
        logging.debug(f"   rowLabelColumnNumbering (last column): '{row_label_col_numbering}'")
        
        # Step 3: Check if last 2 characters match
        ht_last_2 = ht_cell_str[-2:] if len(ht_cell_str) >= 2 else ht_cell_str
        label_last_2 = row_label_col_numbering[-2:] if len(row_label_col_numbering) >= 2 else row_label_col_numbering
        
        has_match = (ht_last_2 == label_last_2) and ht_last_2 != ""
        logging.debug(f"   Last 2 chars match: {has_match} (HT: '{ht_last_2}', Label: '{label_last_2}')")
        
        # Step 4: Process based on match/no-match
        if not has_match:
            # NO MATCH CASE
            logging.info(f"   Processing NO MATCH case for form_code_length=4")
            
            # columnLabel 1 (row 15)
            if ht_cell_value is not None and str(ht_cell_value).strip() != "" and str(ht_cell_value).strip() != ".":
                # Case 1: HT cell is filled - copy from each column
                logging.debug(f"   Case 1: HT cell is filled, copying from each source column")
                output_col = start_col
                for col_offset in range(column_count):
                    source_col = column_start_col + col_offset
                    cell_value = source_sheet.cell(row=ht_row, column=source_col).value
                    if cell_value:
                        table_sheet.cell(row=15, column=output_col).value = str(cell_value)
                        table_sheet.cell(row=15, column=output_col).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                        logging.debug(f"   Row 15, col {get_column_letter(output_col)}: '{cell_value}' from {get_column_letter(source_col)}{ht_row}")
                    output_col += 1
            else:
                # Case 2: HT cell is empty or equals "." - find first non-empty cell to the left
                logging.debug(f"   Case 2: HT cell is empty, searching left for non-empty cell")
                label_value = None
                for col_offset in range(1, column_start_col):
                    check_col = column_start_col - col_offset
                    check_cell = source_sheet.cell(row=ht_row, column=check_col)
                    if check_cell.value and str(check_cell.value).strip() != "":
                        label_value = str(check_cell.value)
                        logging.debug(f"   Found label at {get_column_letter(check_col)}{ht_row}: '{label_value}'")
                        break
                
                if label_value:
                    # Copy to all columns
                    for output_col in range(start_col, last_col + 1):
                        table_sheet.cell(row=15, column=output_col).value = label_value
                        table_sheet.cell(row=15, column=output_col).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                    logging.debug(f"   Filled all columns with: '{label_value}'")
                else:
                    logging.warning(f"   Could not find label value to the left of column range")
            
            logging.info(f"   Filled columnLabel 1 (row 15)")
            
            # No border fix needed for form_code_length=4
            return None
        
        else:
            # MATCH CASE
            logging.info(f"   Processing MATCH case for form_code_length=4")
            
            # Step 1: Find the label start row from cell AJ24
            label_start_row = None
            aj24_cell = source_sheet['AJ24']
            if aj24_cell.value is not None:
                aj24_value = str(aj24_cell.value).strip()
                try:
                    # Try to parse as integer first
                    label_start_row = int(aj24_value)
                    logging.debug(f"   Label start row from AJ24: {label_start_row}")
                except (ValueError, TypeError):
                    # If not a number, try to extract row number from cell reference (e.g., "AH89" -> 89)
                    import re
                    match = re.search(r'(\d+)$', aj24_value)
                    if match:
                        label_start_row = int(match.group(1))
                        logging.debug(f"   Label start row extracted from cell reference '{aj24_value}': {label_start_row}")
                    else:
                        logging.warning(f"   Could not parse label start row from AJ24: '{aj24_value}'")
                        return None
            else:
                logging.warning(f"   Cell AJ24 is empty, cannot determine label start row")
                return None
            
            # Step 2: Calculate label end row (HT row - 1)
            label_end_row = ht_row - 1
            logging.debug(f"   Label end row (HT - 1): {label_end_row}")
            
            # Step 3: Calculate number of label rows
            label_row_count = label_end_row - label_start_row + 1
            logging.debug(f"   Label row count: {label_row_count}")
            
            if label_row_count <= 0:
                logging.warning(f"   Invalid label row count: {label_row_count}, skipping columnLabel filling")
                return None
            
            # Step 4: Fill columnLabel rows based on label_row_count
            if label_row_count == 1:
                # Single label row: only fill columnLabel 1 (row 15)
                logging.debug(f"   Single label row: filling only columnLabel 1")
                output_col = start_col
                for col_offset in range(column_count):
                    source_col = column_start_col + col_offset
                    label_value = source_sheet.cell(row=label_start_row, column=source_col).value
                    if label_value:
                        table_sheet.cell(row=15, column=output_col).value = str(label_value)
                        table_sheet.cell(row=15, column=output_col).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                        logging.debug(f"   Row 15, col {get_column_letter(output_col)}: '{label_value}' from {get_column_letter(source_col)}{label_start_row}")
                    output_col += 1
                
                # Rows 16 and 17 remain unchanged
                logging.debug(f"   Rows 16 and 17 (columnLabel 2 and 3) remain unchanged")
                
                # Return border fix information for Excel COM finalization
                border_fix = {'row': 15, 'last_column': last_col}
                logging.debug(f"   Returning border fix info: {border_fix}")
                return border_fix
            
            elif label_row_count == 3:
                # Special case: 3 label rows that reduce to 2 rows with concatenation
                # Row 1 -> output row 15 (with horizontal merges)
                # Rows 2+3 -> output row 16 (concatenate text)
                # Row after labels -> output row 17 (column IDs with zero-padding)
                logging.debug(f"   3 label rows: filling rows 15-17 with concatenation pattern")
                
                # Calculate the end column in source sheet
                column_end_col = column_start_col + column_count - 1
                
                # Get merged cells from row 1 (label_start_row) only - for horizontal merges
                row1_merged_cells = []
                for merged_range in source_sheet.merged_cells.ranges:
                    # Check if merged range is in row 1 and within column range
                    if (merged_range.min_row == label_start_row and 
                        merged_range.max_row == label_start_row and
                        merged_range.min_col >= column_start_col and 
                        merged_range.max_col <= column_end_col):
                        row1_merged_cells.append({
                            'min_col': merged_range.min_col,
                            'max_col': merged_range.max_col
                        })
                        logging.debug(f"   Row 1 horizontal merge: {get_column_letter(merged_range.min_col)}{merged_range.min_row}:{get_column_letter(merged_range.max_col)}{merged_range.max_row}")
                
                # Step 1: Fill row 15 with row 1 values
                output_col = start_col
                for col_offset in range(column_count):
                    source_col = column_start_col + col_offset
                    label_value = source_sheet.cell(row=label_start_row, column=source_col).value
                    if label_value:
                        table_sheet.cell(row=15, column=output_col).value = str(label_value)
                        table_sheet.cell(row=15, column=output_col).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                        logging.debug(f"   Row 15, col {get_column_letter(output_col)}: '{label_value}' from {get_column_letter(source_col)}{label_start_row}")
                    output_col += 1
                
                # Apply horizontal merges from row 1 to row 15
                for merge_info in row1_merged_cells:
                    col_offset_start = merge_info['min_col'] - column_start_col
                    col_offset_end = merge_info['max_col'] - column_start_col
                    output_min_col = start_col + col_offset_start
                    output_max_col = start_col + col_offset_end
                    
                    if output_min_col != output_max_col:
                        merge_range = f"{get_column_letter(output_min_col)}15:{get_column_letter(output_max_col)}15"
                        try:
                            table_sheet.merge_cells(merge_range)
                            logging.debug(f"   Merged row 15: {merge_range}")
                        except ValueError as e:
                            logging.warning(f"   Could not merge {merge_range}: {e}")
                
                # Step 2: Fill row 16 with concatenated text from rows 2 and 3
                row2 = label_start_row + 1
                row3 = label_start_row + 2
                output_col = start_col
                for col_offset in range(column_count):
                    source_col = column_start_col + col_offset
                    value2 = source_sheet.cell(row=row2, column=source_col).value
                    value3 = source_sheet.cell(row=row3, column=source_col).value
                    
                    # Concatenate with space
                    if value2 and value3:
                        concatenated = f"{value2} {value3}"
                    elif value2:
                        concatenated = str(value2)
                    elif value3:
                        concatenated = str(value3)
                    else:
                        concatenated = None
                    
                    if concatenated:
                        table_sheet.cell(row=16, column=output_col).value = concatenated
                        table_sheet.cell(row=16, column=output_col).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                        logging.debug(f"   Row 16, col {get_column_letter(output_col)}: '{concatenated}' (concatenated from {get_column_letter(source_col)}{row2} and {row3})")
                    output_col += 1
                
                # Row 17 (columnLabel 3) remains unchanged
                logging.debug(f"   Row 17 (columnLabel 3) remains unchanged")
                
                # Return border fix information
                border_fix = {'row': 15, 'last_column': last_col}
                logging.debug(f"   Returning border fix info: {border_fix}")
                return border_fix
            
            elif label_row_count == 4:
                # Four label rows: rows 103-106 map to output rows 15-17
                # Rows 103-104 → Rows 15-16
                # Rows 105-106 → Row 17 (vertically merged)
                logging.debug(f"   Four label rows: mapping to 3 output rows")
                
                # Get source merged cells
                column_end_col = column_start_col + column_count - 1
                source_merged_cells = []
                for merged_range in source_sheet.merged_cells.ranges:
                    if (merged_range.min_row >= label_start_row and 
                        merged_range.max_row <= label_end_row and
                        merged_range.min_col >= column_start_col and 
                        merged_range.max_col <= column_end_col):
                        source_merged_cells.append({
                            'min_row': merged_range.min_row,
                            'max_row': merged_range.max_row,
                            'min_col': merged_range.min_col,
                            'max_col': merged_range.max_col
                        })
                
                # Copy rows 103-104 to rows 15-16
                for output_row_offset in range(2):
                    source_row = label_start_row + output_row_offset
                    output_row = 15 + output_row_offset
                    output_col = start_col
                    for col_offset in range(column_count):
                        source_col = column_start_col + col_offset
                        label_value = source_sheet.cell(row=source_row, column=source_col).value
                        if label_value:
                            table_sheet.cell(row=output_row, column=output_col).value = str(label_value)
                            table_sheet.cell(row=output_row, column=output_col).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                        output_col += 1
                
                # Merge rows 105-106 to row 17
                row3 = label_start_row + 2
                row4 = label_start_row + 3
                output_col = start_col
                for col_offset in range(column_count):
                    source_col = column_start_col + col_offset
                    value3 = source_sheet.cell(row=row3, column=source_col).value
                    value4 = source_sheet.cell(row=row4, column=source_col).value
                    merged_value = str(value3) if value3 else (str(value4) if value4 else None)
                    if merged_value:
                        table_sheet.cell(row=17, column=output_col).value = merged_value
                        table_sheet.cell(row=17, column=output_col).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                    output_col += 1
                
                # Apply source merged cells to output
                for merge_info in source_merged_cells:
                    source_min_row = merge_info['min_row']
                    source_max_row = merge_info['max_row']
                    col_offset_start = merge_info['min_col'] - column_start_col
                    col_offset_end = merge_info['max_col'] - column_start_col
                    output_min_col = start_col + col_offset_start
                    output_max_col = start_col + col_offset_end
                    
                    # Map source rows to output rows
                    if source_min_row == label_start_row and source_max_row == label_end_row:
                        output_min_row, output_max_row = 15, 17
                    elif source_min_row == label_start_row and source_max_row == label_start_row + 1:
                        output_min_row, output_max_row = 15, 16
                    elif source_min_row == label_start_row + 2 and source_max_row == label_end_row:
                        output_min_row, output_max_row = 17, 17
                    elif source_min_row == label_start_row:
                        output_min_row = 15
                        output_max_row = 15 if source_max_row == label_start_row else (17 if source_max_row >= label_start_row + 2 else 16)
                    elif source_min_row == label_start_row + 1:
                        output_min_row = 16
                        output_max_row = 16 if source_max_row == label_start_row + 1 else 17
                    else:
                        output_min_row, output_max_row = 17, 17
                    
                    if output_min_row != output_max_row or output_min_col != output_max_col:
                        merge_range = f"{get_column_letter(output_min_col)}{output_min_row}:{get_column_letter(output_max_col)}{output_max_row}"
                        try:
                            table_sheet.merge_cells(merge_range)
                            logging.debug(f"   Merged: {merge_range}")
                        except ValueError as e:
                            logging.warning(f"   Could not merge {merge_range}: {e}")
                
                border_fix = {'row': 15, 'last_column': last_col}
                return border_fix
            
            else:
                # Multiple label rows: need to check if all columns are fully vertically merged
                logging.debug(f"   Multiple label rows ({label_row_count}): checking merge pattern")
                
                # Calculate the end column in source sheet
                column_end_col = column_start_col + column_count - 1
                
                # Get merged cells from the label range in source sheet
                source_merged_cells = []
                for merged_range in source_sheet.merged_cells.ranges:
                    # Check if merged range is within label rows and column range
                    if (merged_range.min_row >= label_start_row and 
                        merged_range.max_row <= label_end_row and
                        merged_range.min_col >= column_start_col and 
                        merged_range.max_col <= column_end_col):
                        source_merged_cells.append({
                            'min_row': merged_range.min_row,
                            'max_row': merged_range.max_row,
                            'min_col': merged_range.min_col,
                            'max_col': merged_range.max_col
                        })
                        logging.debug(f"   Source merged: {get_column_letter(merged_range.min_col)}{merged_range.min_row}:{get_column_letter(merged_range.max_col)}{merged_range.max_row}")
                
                # Check if ALL columns are fully vertically merged
                all_columns_fully_merged = True
                for col_offset in range(column_count):
                    source_col = column_start_col + col_offset
                    # Check if this column has a vertical merge spanning all label rows
                    col_fully_merged = False
                    for merge_info in source_merged_cells:
                        if (merge_info['min_col'] == source_col and 
                            merge_info['max_col'] == source_col and
                            merge_info['min_row'] == label_start_row and 
                            merge_info['max_row'] == label_end_row):
                            col_fully_merged = True
                            logging.debug(f"   Column {get_column_letter(source_col)} is fully merged vertically")
                            break
                    
                    if not col_fully_merged:
                        all_columns_fully_merged = False
                        logging.debug(f"   Column {get_column_letter(source_col)} is NOT fully merged vertically")
                        break
                
                if all_columns_fully_merged:
                    # Special case: All columns are vertically merged -> only fill row 15
                    logging.debug(f"   All columns vertically merged: filling only columnLabel 1 (row 15)")
                    
                    output_col = start_col
                    for col_offset in range(column_count):
                        source_col = column_start_col + col_offset
                        # Get value from first row of merged cell (label_start_row)
                        label_value = source_sheet.cell(row=label_start_row, column=source_col).value
                        if label_value:
                            table_sheet.cell(row=15, column=output_col).value = str(label_value)
                            table_sheet.cell(row=15, column=output_col).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                            logging.debug(f"   Row 15, col {get_column_letter(output_col)}: '{label_value}' from {get_column_letter(source_col)}{label_start_row}")
                        output_col += 1
                    
                    # Rows 16 and 17 remain unchanged
                    logging.debug(f"   Rows 16 and 17 (columnLabel 2 and 3) remain unchanged")
                    
                    # Return border fix information
                    border_fix = {'row': 15, 'last_column': last_col}
                    logging.debug(f"   Returning border fix info: {border_fix}")
                    return border_fix
                
                else:
                    # Normal case: Not all columns are fully merged -> fill rows 15 and 16 with merges
                    logging.debug(f"   Multiple label rows with partial merges: filling columnLabel 1 and 2")
                    
                    # First pass: copy all cell values
                    for label_row_offset in range(label_row_count):
                        source_row = label_start_row + label_row_offset
                        output_row = 15 + label_row_offset  # Row 15 for first label row, 16 for second
                        
                        output_col = start_col
                        for col_offset in range(column_count):
                            source_col = column_start_col + col_offset
                            label_value = source_sheet.cell(row=source_row, column=source_col).value
                            
                            if label_value:
                                table_sheet.cell(row=output_row, column=output_col).value = str(label_value)
                                table_sheet.cell(row=output_row, column=output_col).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                                logging.debug(f"   Row {output_row}, col {get_column_letter(output_col)}: '{label_value}' from {get_column_letter(source_col)}{source_row}")
                            
                            output_col += 1
                    
                    # Second pass: apply merged cells
                    for merge_info in source_merged_cells:
                        # Calculate output position
                        # Offset from column_start_col to merge min_col
                        col_offset_start = merge_info['min_col'] - column_start_col
                        col_offset_end = merge_info['max_col'] - column_start_col
                        output_min_col = start_col + col_offset_start
                        output_max_col = start_col + col_offset_end
                        
                        # Offset from label_start_row to merge min_row
                        row_offset_start = merge_info['min_row'] - label_start_row
                        row_offset_end = merge_info['max_row'] - label_start_row
                        output_min_row = 15 + row_offset_start
                        output_max_row = 15 + row_offset_end
                        
                        # Only merge if range spans multiple cells
                        if output_min_row != output_max_row or output_min_col != output_max_col:
                            merge_range = f"{get_column_letter(output_min_col)}{output_min_row}:{get_column_letter(output_max_col)}{output_max_row}"
                            try:
                                table_sheet.merge_cells(merge_range)
                                logging.debug(f"   Merged output: {merge_range}")
                            except ValueError as e:
                                logging.warning(f"   Could not merge {merge_range}: {e}")
                    
                    # Row 17 (columnLabel 3) remains unchanged
                    logging.debug(f"   Row 17 (columnLabel 3) remains unchanged")
                    
                    # Return border fix information for Excel COM finalization
                    border_fix = {'row': 15, 'last_column': last_col}
                    logging.debug(f"   Returning border fix info: {border_fix}")
                    return border_fix
    
    # ========================================================================
    # FORM_CODE_LENGTH = 3 LOGIC (existing code)
    # ========================================================================
    logging.debug(f"   Filling columnLabel rows for form_code_length=3")
    logging.debug(f"   Column range: {start_col} to {last_col} ({column_count} columns)")
    
    # Row 15 (columnLabel 1): Merge cells and fill with text
    # First, unmerge any existing merged cells in this range to avoid conflicts
    merged_ranges_to_remove = []
    for merged_range in table_sheet.merged_cells.ranges:
        if merged_range.min_row == 15 and merged_range.max_row == 15:
            if merged_range.min_col >= start_col and merged_range.max_col <= last_col:
                merged_ranges_to_remove.append(merged_range)
    
    for merged_range in merged_ranges_to_remove:
        table_sheet.unmerge_cells(str(merged_range))
        logging.debug(f"   Unmerged existing range: {merged_range}")
    
    # Save the border styles of the last column from a reference row (row 18 or 19)
    # Row 15 might not have the correct border yet, so we use row 18 as reference
    from openpyxl.utils import get_column_letter
    from copy import copy
    reference_cell = table_sheet.cell(row=18, column=last_col)
    saved_right_border = copy(reference_cell.border.right) if reference_cell.border and reference_cell.border.right else None
    saved_border_color = copy(reference_cell.border.right.color) if reference_cell.border and reference_cell.border.right and reference_cell.border.right.color else None
    
    # Merge cells from column N to last column in row 15
    if column_count > 1:
        start_cell = f"{get_column_letter(start_col)}15"
        end_cell = f"{get_column_letter(last_col)}15"
        merge_range = f"{start_cell}:{end_cell}"
        table_sheet.merge_cells(merge_range)
        logging.debug(f"   Merged cells: {merge_range}")
    
    # Fill the merged cell with text
    cell = table_sheet.cell(row=15, column=start_col)
    cell.value = "Schéma sur base territoriale et sociale - Etablissements de crédit Belge et Succursales d'EC de l'UE"
    cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    
    # Restore the thick right border on the last column by creating a new Border with explicit Side
    from openpyxl.styles import Border, Side
    last_col_cell = table_sheet.cell(row=15, column=last_col)
    current_border = last_col_cell.border
    
    # Create an explicit Side for the right border with medium style
    right_side = Side(style='medium', color=saved_border_color) if saved_border_color else Side(style='medium')
    
    last_col_cell.border = Border(
        left=current_border.left if current_border else None,
        right=right_side,
        top=current_border.top if current_border else None,
        bottom=current_border.bottom if current_border else None
    )
    logging.debug(f"   Restored right border (style=medium) on column {get_column_letter(last_col)}")
    
    logging.info(f"   Filled columnLabel 1 (row 15) for {column_count} columns")
    
    # ========== columnLabel 2 (row 16) ==========
    if source_sheet and column_start_col is not None:
        logging.info(f"   Filling columnLabel 2 (row 16)")
        
        # Step 1: Find "Start reference" cell
        start_ref_cell = None
        start_row_num = None
        
        for row in source_sheet.iter_rows():
            for cell in row:
                if cell.value and isinstance(cell.value, str) and "Start reference" in cell.value:
                    start_ref_cell = cell
                    logging.debug(f"   Found 'Start reference' at {cell.coordinate}")
                    break
            if start_ref_cell:
                break
        
        if start_ref_cell:
            # Step 2: Navigate right to find first non-empty cell
            start_ref_row = start_ref_cell.row
            start_ref_col = start_ref_cell.column
            
            # Look for non-empty cell to the right
            ref_value_cell = None
            for col_offset in range(1, 10):  # Check up to 10 columns to the right
                check_cell = source_sheet.cell(row=start_ref_row, column=start_ref_col + col_offset)
                if check_cell.value:
                    ref_value_cell = check_cell
                    logging.debug(f"   Found reference value at {check_cell.coordinate}: {check_cell.value}")
                    break
            
            if ref_value_cell:
                # Step 3: Extract cell reference from value (could be direct or formula)
                ref_text = str(ref_value_cell.value)
                
                # Try to find cell reference pattern (e.g., AH97, $A$10, etc.)
                # Pattern matches: letter(s) followed by number(s)
                cell_ref_pattern = r'([A-Z]+)(\d+)'
                matches = re.findall(cell_ref_pattern, ref_text)
                
                if matches:
                    # Use the last match (in case of complex formulas)
                    col_letters, row_str = matches[-1]
                    start_row_num = int(row_str)
                    logging.debug(f"   Extracted start row: {start_row_num} from reference '{ref_text}'")
        
        if start_row_num is None:
            logging.warning(f"   Could not find start row from 'Start reference', skipping columnLabel 2")
        else:
            # Step 4: Find "HT" in column A to determine end row
            end_row_num = None
            for row in source_sheet.iter_rows(min_col=1, max_col=1):
                cell = row[0]
                if cell.value and isinstance(cell.value, str) and "HT" in cell.value:
                    end_row_num = cell.row - 1
                    logging.debug(f"   Found 'HT' at A{cell.row}, end row: {end_row_num}")
                    break
            
            if end_row_num is None:
                logging.warning(f"   Could not find 'HT' marker, skipping columnLabel 2")
            else:
                # Step 5: Extract text from identified rows
                output_col = 14  # Start at column N (14)
                
                for col_offset in range(column_count):
                    source_col = column_start_col + col_offset
                    
                    if end_row_num == start_row_num:
                        # Single row case: copy text directly
                        text = source_sheet.cell(row=start_row_num, column=source_col).value
                        if text:
                            table_sheet.cell(row=16, column=output_col).value = str(text)
                            logging.debug(f"   Row 16, col {get_column_letter(output_col)}: '{text}' from {get_column_letter(source_col)}{start_row_num}")
                    
                    elif end_row_num == start_row_num + 1:
                        # Two row case: combine with space
                        text1 = source_sheet.cell(row=start_row_num, column=source_col).value
                        text2 = source_sheet.cell(row=end_row_num, column=source_col).value
                        combined = ""
                        if text1:
                            combined = str(text1)
                        if text2:
                            combined = combined + " " + str(text2) if combined else str(text2)
                        
                        if combined:
                            table_sheet.cell(row=16, column=output_col).value = combined
                            logging.debug(f"   Row 16, col {get_column_letter(output_col)}: '{combined}' (combined)")
                    
                    else:
                        logging.warning(f"   Unexpected row range: start={start_row_num}, end={end_row_num}")
                    
                    # Apply center alignment to the cell
                    table_sheet.cell(row=16, column=output_col).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                    output_col += 1
                
                logging.info(f"   Filled columnLabel 2 (row 16) with data from rows {start_row_num}-{end_row_num}")
    
    # Return border fix information for Excel COM finalization
    border_fix = {'row': 15, 'last_column': last_col}
    logging.debug(f"   Returning border fix info: {border_fix}")
    return border_fix


def cleanup_internal_tables(internal_tables, form_code_length, form_code):
    """
    Clean up internal tables according to specific rules.
    
    Cleanup rule 1 (only for form_code_length = 3):
    - If rowNumbering (trimmed) is empty:
      1. Take rowLabelColumnNumbering from current line
      2. Remove all leading spaces (LTRIM)
      3. Add 1 single leading space
      4. Append to rowLabelColumnNumbering of previous line
      5. Delete the current line
    
    Cleanup rule 2:
    - For lines where listBreakdown = "yes":
      Replace rowNumbering value with "999"
    
    Cleanup rule 3 (only for formCode 5144):
    - Remove all occurrences of capital letter "M" from rowNumbering and feedingRuleSource columns
    - Reason: alphanumeric row numbers are not allowed in SBPRR
    
    Cleanup rule 4 (only for formCode 9312):
    - For sheets "Sheet 03", "Sheet 04", "Sheet 05":
      * If rowNumbering = "T1" or "0T1", replace with "777"
      * If rowNumbering = "T2" or "0T2", replace with "888"
    - For sheets "Sheet 04", "Sheet 05":
      * Overwrite rowLabelColumnNumbering with content from column AI (instead of column AH)
      * Reason: content of column AH is the same as column AG
    
    Args:
        internal_tables: Dictionary of internal tables {sheet_name: [row_data, ...]}
        form_code_length: Length of the form code (3 or 4 digits)
        form_code: The form code string (e.g., "010", "5144", "9312")
    
    Returns:
        dict: Cleaned internal tables
    """
    cleaned_tables = {}
    
    for sheet_name, sheet_data in internal_tables.items():
        if not sheet_data:
            cleaned_tables[sheet_name] = []
            continue
        
        # Rule 1: Merge rows with empty rowNumbering (only for form_code_length = 3)
        if form_code_length == 3:
            cleaned_data = []
            rows_merged = 0
            
            for i, row in enumerate(sheet_data):
                # Check if rowNumbering is empty (after trimming)
                row_num_str = str(row['rowNumbering']).strip() if row['rowNumbering'] else ""
                
                if not row_num_str and i > 0:
                    # This row has empty rowNumbering - merge with previous
                    current_label = row['rowLabelColumnNumbering']
                    if current_label:
                        # Remove all leading spaces and add 1 single leading space
                        trimmed_label = str(current_label).lstrip()
                        appended_text = " " + trimmed_label
                        
                        # Append to previous row's rowLabelColumnNumbering
                        if cleaned_data:
                            prev_label = cleaned_data[-1]['rowLabelColumnNumbering'] or ""
                            cleaned_data[-1]['rowLabelColumnNumbering'] = str(prev_label) + appended_text
                            rows_merged += 1
                            logging.debug(f"   Merged row {row['row_num']} into previous row (empty rowNumbering)")
                    # Skip this row (delete it by not adding to cleaned_data)
                else:
                    # Keep this row
                    cleaned_data.append(row)
            
            cleaned_tables[sheet_name] = cleaned_data
            
            if rows_merged > 0:
                logging.debug(f"   Cleanup rule 1: Merged {rows_merged} rows in '{sheet_name}' (form_code_length=3)")
                logging.debug(f"   Rows after cleanup: {len(cleaned_data)} (was {len(sheet_data)})")
        else:
            # form_code_length != 3: no cleanup needed for rule 1
            cleaned_tables[sheet_name] = sheet_data
    
    # Rule 2: Replace rowNumbering with "999" for rows where listBreakdown = "yes"
    for sheet_name in cleaned_tables:
        rows_updated = 0
        for row in cleaned_tables[sheet_name]:
            if row.get('listBreakdown') == "yes":
                old_value = row['rowNumbering']
                row['rowNumbering'] = "999"
                rows_updated += 1
                logging.debug(f"   Cleanup rule 2: Changed rowNumbering from '{old_value}' to '999' for row {row['row_num']} (listBreakdown=yes)")
        
        if rows_updated > 0:
            logging.debug(f"   Cleanup rule 2: Updated {rows_updated} rows in '{sheet_name}' (listBreakdown=yes)")
    
    # Rule 3: Remove "M" from rowNumbering and feedingRuleSource (only for formCode 5144)
    if form_code == "5144":
        for sheet_name in cleaned_tables:
            rows_updated = 0
            for row in cleaned_tables[sheet_name]:
                # Remove "M" from rowNumbering
                if row['rowNumbering']:
                    old_row_num = row['rowNumbering']
                    new_row_num = str(old_row_num).replace("M", "")
                    if old_row_num != new_row_num:
                        row['rowNumbering'] = new_row_num
                        rows_updated += 1
                        logging.debug(f"   Cleanup rule 3: Changed rowNumbering from '{old_row_num}' to '{new_row_num}' for row {row['row_num']} (removed M)")
                
                # Remove "M" from feedingRuleSource
                if row['feedingRuleSource']:
                    old_feeding = row['feedingRuleSource']
                    new_feeding = str(old_feeding).replace("M", "")
                    if old_feeding != new_feeding:
                        row['feedingRuleSource'] = new_feeding
                        logging.debug(f"   Cleanup rule 3: Changed feedingRuleSource from '{old_feeding}' to '{new_feeding}' for row {row['row_num']} (removed M)")
            
            if rows_updated > 0:
                logging.debug(f"   Cleanup rule 3: Removed 'M' from {rows_updated} rows in '{sheet_name}' (formCode=5144)")
    
    # Rule 4: Special handling for formCode 9312
    if form_code == "9312":
        # Part 1: Replace T1/0T1 with "777" and T2/0T2 with "888" for specific sheets
        target_sheets_part1 = ["Sheet 03", "Sheet 04", "Sheet 05"]
        for sheet_name in cleaned_tables:
            if sheet_name in target_sheets_part1:
                rows_updated = 0
                for row in cleaned_tables[sheet_name]:
                    if row['rowNumbering']:
                        old_row_num = str(row['rowNumbering']).strip()
                        new_row_num = old_row_num
                        
                        if old_row_num in ["T1", "0T1"]:
                            new_row_num = "777"
                        elif old_row_num in ["T2", "0T2"]:
                            new_row_num = "888"
                        
                        if old_row_num != new_row_num:
                            row['rowNumbering'] = new_row_num
                            rows_updated += 1
                            logging.debug(f"   Cleanup rule 4: Changed rowNumbering from '{old_row_num}' to '{new_row_num}' for row {row['row_num']}")
                
                if rows_updated > 0:
                    logging.debug(f"   Cleanup rule 4 (part 1): Updated {rows_updated} rows in '{sheet_name}' (formCode=9312)")
        
        # Part 2: Overwrite rowLabelColumnNumbering with column AI for Sheet 04 and Sheet 05
        target_sheets_part2 = ["Sheet 04", "Sheet 05"]
        for sheet_name in cleaned_tables:
            if sheet_name in target_sheets_part2:
                rows_updated = 0
                for row in cleaned_tables[sheet_name]:
                    # Column AI is column 35 (0-indexed: 34 in row_data)
                    if row.get('row_data') and len(row['row_data']) > 34:
                        ai_value = row['row_data'][34]  # Column AI (35th column, 0-indexed = 34)
                        if ai_value is not None:
                            old_label = row['rowLabelColumnNumbering']
                            row['rowLabelColumnNumbering'] = ai_value
                            rows_updated += 1
                            logging.debug(f"   Cleanup rule 4: Overwrote rowLabelColumnNumbering from '{old_label}' to '{ai_value}' for row {row['row_num']} (using column AI)")
                
                if rows_updated > 0:
                    logging.debug(f"   Cleanup rule 4 (part 2): Overwrote {rows_updated} rowLabelColumnNumbering values in '{sheet_name}' with column AI (formCode=9312)")
    
    return cleaned_tables


def calculate_level_column(internal_tables, form_code_length):
    """
    Add a 'level' column to each row in the internal tables based on form_code_length.
    
    For form_code_length = 3:
    - level = length(rowNumbering) - 1
    - Exception: if (length(rowNumbering) = 3) AND ((last digit = 0) OR (last 2 digits = 99)),
      then level = length(rowNumbering) - 2
    
    For form_code_length = 4:
    - Parse all rowLabelColumnNumbering to count leading spaces
    - Create unique sorted list of leading space counts
    - Assign level starting from 1 for each unique count
    - Fill level for each row based on its leading space count
    
    Args:
        internal_tables: Dictionary of internal tables {sheet_name: [row_data, ...]}
        form_code_length: Length of the form code (3 or 4 digits)
    
    Returns:
        dict: Internal tables with 'level' column added to each row
    """
    logging.info("🔧  Calculating level column for internal tables")
    
    for sheet_name, sheet_data in internal_tables.items():
        if not sheet_data:
            continue
        
        if form_code_length == 3:
            # Form code length 3: level based on rowNumbering length
            for row in sheet_data:
                row_numbering = str(row.get('rowNumbering', '')).strip()
                
                if not row_numbering:
                    row['level'] = 0
                    continue
                
                row_num_len = len(row_numbering)
                
                # Check exception: length = 3 AND (last digit = 0 OR last 2 digits = 99)
                if row_num_len == 3:
                    last_digit = row_numbering[-1]
                    last_two_digits = row_numbering[-2:] if len(row_numbering) >= 2 else ""
                    
                    if last_digit == '0' or last_two_digits == '99':
                        level = row_num_len - 2
                    else:
                        level = row_num_len - 1
                else:
                    level = row_num_len - 1
                
                row['level'] = level
                logging.debug(f"   Sheet '{sheet_name}', row {row.get('row_num')}: rowNumbering='{row_numbering}', level={level}")
        
        elif form_code_length == 4:
            # Form code length 4: level based on leading spaces in rowLabelColumnNumbering
            
            # Step 1: Parse all rowLabelColumnNumbering to count leading spaces
            leading_spaces_map = {}
            for row in sheet_data:
                label = row.get('rowLabelColumnNumbering', '')
                if label is None:
                    label = ''
                label_str = str(label)
                
                # Count leading spaces
                leading_spaces = len(label_str) - len(label_str.lstrip(' '))
                leading_spaces_map[id(row)] = leading_spaces
            
            # Step 2: Create unique sorted list of leading space counts
            unique_leading_spaces = sorted(set(leading_spaces_map.values()))
            
            # Step 3: Create mapping from leading space count to level
            space_to_level = {}
            for idx, space_count in enumerate(unique_leading_spaces, start=1):
                space_to_level[space_count] = idx
            
            logging.debug(f"   Sheet '{sheet_name}': Leading spaces to level mapping: {space_to_level}")
            
            # Step 4: Assign level to each row
            for row in sheet_data:
                leading_spaces = leading_spaces_map[id(row)]
                level = space_to_level[leading_spaces]
                row['level'] = level
                logging.debug(f"   Sheet '{sheet_name}', row {row.get('row_num')}: leading_spaces={leading_spaces}, level={level}")
        
        logging.info(f"   Calculated level column for sheet '{sheet_name}' ({len(sheet_data)} rows)")
    
    return internal_tables


def remove_leading_spaces_from_labels(internal_tables):
    """
    Remove all leading spaces from rowLabelColumnNumbering for all rows in the internal tables.
    This is done after level calculation since form_code_length=4 uses leading spaces to determine level.
    
    Args:
        internal_tables: Dictionary of internal tables {sheet_name: [row_data, ...]}
    
    Returns:
        dict: Internal tables with leading spaces removed from rowLabelColumnNumbering
    """
    logging.info("🔧  Removing leading spaces from rowLabelColumnNumbering")
    
    for sheet_name, sheet_data in internal_tables.items():
        if not sheet_data:
            continue
        
        rows_updated = 0
        for row in sheet_data:
            label = row.get('rowLabelColumnNumbering', '')
            if label is not None and label != '':
                label_str = str(label)
                stripped_label = label_str.lstrip(' ')
                
                if label_str != stripped_label:
                    row['rowLabelColumnNumbering'] = stripped_label
                    rows_updated += 1
                    logging.debug(f"   Sheet '{sheet_name}', row {row.get('row_num')}: removed {len(label_str) - len(stripped_label)} leading space(s)")
        
        if rows_updated > 0:
            logging.info(f"   Removed leading spaces from {rows_updated} rows in sheet '{sheet_name}'")
    
    return internal_tables


def convert_6_4_digit_row_numbers(internal_tables, mapping_file_path):
    """
    Convert 6-digit (or 5-digit) row numbers to 4-digit equivalents using a mapping file.
    Only applicable when conversion_6_4_digits = 'Y'.
    
    For rows where length(rowNumbering) > 4:
    - Look up rowNumbering in mapping file (sheet RUBORI_RUBNEW, column A)
    - Update rowLabelColumnNumbering by prepending "(original_rowNumbering) "
    - Replace rowNumbering with value from column B (RUBNEW)
    
    For rows where feedingRuleSource is filled:
    - Find patterns L99999 or L999999 (L followed by 5 or 6 digits)
    - Replace with corresponding values from mapping file
    
    Args:
        internal_tables: Dictionary of internal tables {sheet_name: [row_data, ...]}
        mapping_file_path: Path to Mapping_Rub_6_4_digits.xlsx file
    
    Returns:
        dict: Internal tables with converted row numbers
    """
    import re
    
    logging.info("🔧  Converting 6/5-digit row numbers to 4-digit using mapping file")
    
    # Load mapping file
    try:
        mapping_wb = openpyxl.load_workbook(mapping_file_path, data_only=True)
        mapping_sheet = mapping_wb['RUBORI_RUBNEW']
        logging.debug(f"   Loaded mapping file: {mapping_file_path}")
    except Exception as e:
        logging.error(f"   Failed to load mapping file: {e}")
        return internal_tables
    
    # Build lookup dictionary: RUBORI (column A) -> RUBNEW (column B)
    mapping_dict = {}
    for row in mapping_sheet.iter_rows(min_row=2, values_only=True):  # Skip header row
        rubori = row[0]  # Column A
        rubnew = row[1]  # Column B
        
        if rubori is not None and rubnew is not None:
            rubori_str = str(rubori).strip()
            rubnew_str = str(rubnew).strip()
            if rubori_str and rubnew_str:
                mapping_dict[rubori_str] = rubnew_str
    
    mapping_wb.close()
    logging.debug(f"   Built mapping dictionary with {len(mapping_dict)} entries")
    
    # Process each sheet
    for sheet_name, sheet_data in internal_tables.items():
        if not sheet_data:
            continue
        
        rows_converted = 0
        feeding_rules_updated = 0
        
        for row in sheet_data:
            row_numbering = str(row.get('rowNumbering', '')).strip()
            
            # Convert rowNumbering if length > 4
            if row_numbering and len(row_numbering) > 4:
                if row_numbering in mapping_dict:
                    new_row_numbering = mapping_dict[row_numbering]
                    
                    # Update rowLabelColumnNumbering by prepending "(original) "
                    current_label = row.get('rowLabelColumnNumbering', '')
                    if current_label:
                        updated_label = f"({row_numbering}) {current_label}"
                        row['rowLabelColumnNumbering'] = updated_label
                        logging.debug(f"   Sheet '{sheet_name}', row {row.get('row_num')}: rowLabelColumnNumbering updated from '{current_label}' to '{updated_label}'")
                    
                    # Update rowNumbering with new value
                    row['rowNumbering'] = new_row_numbering
                    rows_converted += 1
                    logging.debug(f"   Sheet '{sheet_name}', row {row.get('row_num')}: converted '{row_numbering}' -> '{new_row_numbering}'")
                else:
                    logging.warning(f"   Sheet '{sheet_name}', row {row.get('row_num')}: no mapping found for '{row_numbering}'")
            
            # Convert patterns in feedingRuleSource
            feeding_rule = row.get('feedingRuleSource', '')
            if feeding_rule:
                feeding_rule_str = str(feeding_rule)
                
                # Find all patterns L99999 (5 digits) or L999999 (6 digits)
                pattern = r'L(\d{5,6})'
                matches = re.findall(pattern, feeding_rule_str)
                
                if matches:
                    updated_feeding_rule = feeding_rule_str
                    for match in matches:
                        if match in mapping_dict:
                            # Replace Lxxxxx with Lyyyy where yyyyy is the mapped value
                            old_pattern = f"L{match}"
                            new_pattern = f"L{mapping_dict[match]}"
                            updated_feeding_rule = updated_feeding_rule.replace(old_pattern, new_pattern)
                            logging.debug(f"   Sheet '{sheet_name}', row {row.get('row_num')}: replaced '{old_pattern}' with '{new_pattern}' in feedingRuleSource")
                        else:
                            logging.warning(f"   Sheet '{sheet_name}', row {row.get('row_num')}: no mapping found for '{match}' in feedingRuleSource")
                    
                    if updated_feeding_rule != feeding_rule_str:
                        row['feedingRuleSource'] = updated_feeding_rule
                        feeding_rules_updated += 1
        
        if rows_converted > 0:
            logging.info(f"   Converted {rows_converted} row numbers in sheet '{sheet_name}'")
        if feeding_rules_updated > 0:
            logging.info(f"   Updated {feeding_rules_updated} feeding rules in sheet '{sheet_name}'")
    
    return internal_tables


def transform_feeding_rule_patterns(internal_tables):
    """
    Transform feeding rule patterns:
    - Remove all spaces from feedingRuleSource
    - Replace patterns L{1-6 digits} with {r{digits}}
    
    Example: "L123 + L456789" -> "{r123}+{r456789}"
    """
    logging.info("🔧  Transforming feeding rule patterns")
    
    for sheet_name, sheet_data in internal_tables.items():
        if not sheet_data:
            continue
        
        rules_transformed = 0
        
        for row in sheet_data:
            feeding_rule = row.get('feedingRuleSource', '')
            
            # Only process if feedingRuleSource is filled in
            if feeding_rule:
                feeding_rule_str = str(feeding_rule).strip()
                
                if feeding_rule_str:
                    # Step 1: Remove apostrophes before first = sign
                    if '=' in feeding_rule_str:
                        equals_pos = feeding_rule_str.find('=')
                        if equals_pos > 0:
                            # Remove all ' characters before the =
                            before_equals = feeding_rule_str[:equals_pos].replace("'", "")
                            feeding_rule_str = before_equals + feeding_rule_str[equals_pos:]
                    
                    # Step 2: Remove all spaces
                    feeding_rule_no_spaces = feeding_rule_str.replace(' ', '')
                    
                    # Step 3: Replace L{1-6 digits} with {r{digits}}
                    # Pattern: L followed by 1 to 6 digits
                    pattern = r'L(\d{1,6})'
                    transformed_rule = re.sub(pattern, r'{r\1}', feeding_rule_no_spaces)
                    
                    # Update the row if changed
                    if transformed_rule != feeding_rule:
                        row['feedingRuleSource'] = transformed_rule
                        rules_transformed += 1
                        logging.debug(f"   Sheet '{sheet_name}', row {row.get('row_num')}: transformed feeding rule from '{feeding_rule}' to '{transformed_rule}'")
        
        if rules_transformed > 0:
            logging.info(f"   Transformed {rules_transformed} feeding rules in sheet '{sheet_name}'")
    
    return internal_tables


def calculate_max_depth(feeding_rule, row_lookup, visited):
    """
    Recursively calculate the maximum depth of the dependency tree for feeding rules.
    
    Args:
        feeding_rule: The feeding rule string containing {r...} patterns
        row_lookup: Dictionary mapping rowNumbering to row data
        visited: Set of visited row numbers to avoid cycles
    
    Returns:
        Maximum depth of the dependency tree (0 if no dependencies)
    """
    # Extract all patterns r{1-6 digits} from feedingRuleSource
    pattern = r'r(\d{1,6})'
    matches = re.findall(pattern, feeding_rule)
    
    if not matches:
        return 0
    
    max_child_depth = 0
    has_dependency = False
    
    for row_numbering in matches:
        # Avoid cycles
        if row_numbering in visited:
            continue
        
        # Find the corresponding row
        if row_numbering in row_lookup:
            child_row = row_lookup[row_numbering]
            child_feeding_rule = child_row.get('feedingRuleSource', '')
            
            if child_feeding_rule:
                child_feeding_rule_str = str(child_feeding_rule).strip()
                
                if child_feeding_rule_str:
                    has_dependency = True
                    # Mark as visited to avoid cycles
                    new_visited = visited.copy()
                    new_visited.add(row_numbering)
                    
                    # Recursively calculate depth
                    child_depth = 1 + calculate_max_depth(child_feeding_rule_str, row_lookup, new_visited)
                    max_child_depth = max(max_child_depth, child_depth)
    
    return max_child_depth


def calculate_computation_rule_order(internal_tables):
    """
    Calculate computationRuleOrder for each row based on feeding rule dependencies.
    This represents the maximum depth of the dependency tree.
    
    For each row with a feedingRuleSource:
    - Initialize computationRuleOrder to 0
    - Extract all r{digits} patterns from feedingRuleSource
    - For each pattern, find the corresponding row and check if it has dependencies
    - Recursively calculate the depth and take the maximum across all branches
    """
    logging.info("🔧  Calculating computationRuleOrder for internal tables")
    
    for sheet_name, sheet_data in internal_tables.items():
        if not sheet_data:
            continue
        
        # Build a lookup dictionary: rowNumbering -> row data
        row_lookup = {}
        for row in sheet_data:
            row_numbering = str(row.get('rowNumbering', '')).strip()
            if row_numbering:
                row_lookup[row_numbering] = row
        
        # Calculate computationRuleOrder for each row
        rows_calculated = 0
        
        for row in sheet_data:
            feeding_rule = row.get('feedingRuleSource', '')
            
            if feeding_rule:
                feeding_rule_str = str(feeding_rule).strip()
                
                if feeding_rule_str:
                    # Calculate the maximum depth of the dependency tree
                    computation_order = calculate_max_depth(feeding_rule_str, row_lookup, set())
                    row['computationRuleOrder'] = computation_order
                    rows_calculated += 1
                    logging.debug(f"   Sheet '{sheet_name}', row {row.get('row_num')}: computationRuleOrder = {computation_order}")
                else:
                    row['computationRuleOrder'] = None
            else:
                # Empty feedingRuleSource: no computationRuleOrder
                row['computationRuleOrder'] = None
        
        if rows_calculated > 0:
            logging.info(f"   Calculated computationRuleOrder for {rows_calculated} rows in sheet '{sheet_name}'")
    
    return internal_tables


def copy_internal_table_to_output(output_sheet, internal_table_data):
    """
    Copy data from internal table to the SA.*.* output sheet.
    
    Copies data from the internal table dictionary to the corresponding cells
    in the output sheet, starting at row 19.
    
    Mapping:
    - internal_table row 0 → output row 19
    - internal_table row 1 → output row 20
    - etc.
    
    Column mapping:
    - level → Column A
    - abstractRow → Column C
    - feedingRuleSource → Column G
    - computationRuleOrder → Column I
    - rowNumbering → Column L
    - rowLabelColumnNumbering → Column M
    
    Args:
        output_sheet: The target worksheet (SA.*.* sheet)
        internal_table_data: List of dictionaries containing row data
    """
    logging.info(f"🔧  Copying internal table data to output sheet")
    
    if not internal_table_data:
        logging.info(f"   No data to copy (internal table is empty)")
        return
    
    # Starting row in output sheet
    output_start_row = 19
    
    # Column mapping: internal_table_key → output_column_index
    column_mapping = {
        'level': 1,                      # Column A
        'abstractRow': 3,                 # Column C
        'feedingRuleSource': 7,          # Column G
        'computationRuleOrder': 9,       # Column I
        'rowNumbering': 12,              # Column L
        'rowLabelColumnNumbering': 13    # Column M
    }
    
    # Get default value for empty rowLabelColumnNumbering from cell A2
    default_row_label = None
    a2_value = output_sheet.cell(row=2, column=1).value
    if a2_value:
        # Remove all text from beginning until first ":" (including the ":")
        if ":" in str(a2_value):
            temp_value = str(a2_value).split(":", 1)[1]  # Take everything after first ":"
            default_row_label = temp_value.lstrip()  # Left trim to remove leading spaces
            logging.debug(f"   Default rowLabelColumnNumbering from A2: '{default_row_label}'")
        else:
            # No ":" found, use the value as-is with left trim
            default_row_label = str(a2_value).lstrip()
            logging.debug(f"   Default rowLabelColumnNumbering from A2 (no colon): '{default_row_label}'")
    
    rows_copied = 0
    
    # Loop through all rows in the internal table
    for idx, row_data in enumerate(internal_table_data):
        output_row = output_start_row + idx
        
        # Copy each column value to the corresponding cell
        for internal_key, output_col in column_mapping.items():
            value = row_data.get(internal_key)
            
            # Special handling for rowLabelColumnNumbering (Column M)
            if internal_key == 'rowLabelColumnNumbering':
                # If value is empty/None, use default from A2
                if value is None or (isinstance(value, str) and value.strip() == ""):
                    if default_row_label:
                        value = default_row_label
                        logging.debug(f"   Row {output_row}, Column M: Using default from A2: '{value}'")
                
                # Only write if we have a value
                if value is not None:
                    cell = output_sheet.cell(row=output_row, column=output_col)
                    cell.value = value
                    # Enable wrap text for this cell
                    cell.alignment = Alignment(wrap_text=True)
            
            # Only write non-None values for other columns
            elif value is not None:
                cell = output_sheet.cell(row=output_row, column=output_col)
                
                # Special handling for feedingRuleSource (Column G)
                if internal_key == 'feedingRuleSource':
                    cell.value = value
                    # Enable wrap text for this cell
                    cell.alignment = Alignment(wrap_text=True)
                    # Set cell format to Text and data type to string
                    cell.number_format = '@'
                    cell.data_type = 's'  # Force string type to prevent formula interpretation
                
                # Special handling for rowNumbering (Column L) - ensure text format
                elif internal_key == 'rowNumbering':
                    # Convert to string to ensure it's treated as text
                    cell.value = str(value)
                    # Set cell number format to text
                    cell.number_format = '@'
                
                # Default handling for other columns
                else:
                    cell.value = value
        
        rows_copied += 1
        
        # Log details for debugging (only for first few and last few rows to avoid spam)
        if idx < 3 or idx >= len(internal_table_data) - 3:
            logging.debug(f"   Row {idx} → output row {output_row}: rowNumbering={row_data.get('rowNumbering')}, level={row_data.get('level')}, abstractRow={row_data.get('abstractRow')}")
    
    logging.info(f"   Copied {rows_copied} rows from internal table to output sheet")


def group_sheets_by_root_name(sheets):
    grouped = {}
    for sheet in sheets:
        root_name = extract_root_name(sheet)
        if root_name not in grouped:
            grouped[root_name] = []
        grouped[root_name].append(sheet)
        logging.debug(f"   Grouping: {sheet} -> {root_name}")
    return grouped


def create_empty_target_sheets(target_wb, sheet_list, sa_code):
    """
    Create SA.*.* sheets in the target workbook based on input sheet names.
    
    This function applies specific naming conventions to transform input sheet names
    to output sheet names in the format SA.XXX.YY or SA.XXX.TEC:
    
    Naming conventions:
    - "Sheet T1" → "SA.XXX.TEC" (technical sheet)
    - "Sheet 99" (where 99 is any 2-digit number) → "SA.XXX.YY"
      Examples: "Sheet 01" → "SA.010.01", "Sheet 12" → "SA.010.12"
    
    Args:
        target_wb: Target workbook where sheets will be created
        sheet_list: List of sheet names from the input file to process
        sa_code: The SA code (e.g., "SA.010", "SA.2152") extracted from formCode
    
    Returns:
        dict: Mapping from input sheet names to output sheet names
              Example: {"Sheet 01": "SA.010.01", "Sheet T1": "SA.010.TEC"}
    """
    logging.debug("   === PROCESS WORKBOOKS ===")

    # Reference to the template sheet ('code')
    template_sheet = target_wb["code"]

    # Create a mapping from input sheet names to output sheet names
    sheet_name_mapping = {}

    # Create copies for each target sheet
    for input_sheet_name in sheet_list:
        # Determine the output sheet name based on naming conventions
        if input_sheet_name == "Sheet T1":
            # Special case: Sheet T1 -> SA.XXX.TEC
            output_sheet_name = f"{sa_code}.TEC"
        elif input_sheet_name.startswith("Sheet "):
            # Extract the 2-digit number from "Sheet 99"
            match = re.search(r'Sheet (\d{2})', input_sheet_name)
            if match:
                digits = match.group(1)
                output_sheet_name = f"{sa_code}.{digits}"
            else:
                # Fallback: use original name if pattern doesn't match
                output_sheet_name = input_sheet_name
                logging.warning(f"   Sheet name '{input_sheet_name}' doesn't match expected pattern. Using original name.")
        else:
            # Fallback for unexpected format
            output_sheet_name = input_sheet_name

        # Create the new sheet with the output name
        new_sheet = target_wb.copy_worksheet(template_sheet)
        new_sheet.title = output_sheet_name
        sheet_name_mapping[input_sheet_name] = output_sheet_name

        logging.debug(f"   Created sheet: '{input_sheet_name}' -> '{output_sheet_name}'")

        # 🟢 Reapply data validations manually (for openpyxl < 3.1)
        for dv in template_sheet.data_validations.dataValidation:
            # Clone data validation and reassign its range to the new sheet
            new_dv = copy(dv)  # Python's copy.copy
            new_sheet.add_data_validation(new_dv)

    # Remove the original 'code' sheet after duplication
    target_wb.remove(template_sheet)

    logging.debug(f"   Target workbook sheets after processing: {target_wb.sheetnames}")
    
    return sheet_name_mapping


###############################################################################################################################################################
#                                                                                                                                                             #  
#       3rd part: handling the extraction and pasting of headers (both column and row headers) in the generated workbooks.                                    #  
#                                                                                                                                                             #  
#   Extracts column headers and row headers from the source sheet.                                                                                            #  
#   Adjusts the "columnLabel" rows in the target table sheet to match the number of extracted column headers.                                                 #  
#                                                                                                                                                             #                                                                                                                                               #  
###############################################################################################################################################################

def get_fill_rgb(cell):
    if not cell or not cell.fill:
        return None
        
    fill = cell.fill
    if fill.fgColor and fill.fgColor.type == 'rgb':
        return fill.fgColor.rgb
    return None

def find_cell_by_value(sheet, target_value):
    #  Find the first cell containing a specific value.
    if not sheet or not target_value:
        return None
        
    for row in sheet.iter_rows(min_row=1, max_row=sheet.max_row, max_col=sheet.max_column):
        for cell in row:
            if cell.value == target_value:
                return cell
    return None

def get_reference_color_info(sheet, reference_value):
    #  Find a reference cell and return its color and position information.
    reference_cell = find_cell_by_value(sheet, reference_value)
    
    if not reference_cell:
        logging.debug(f"   Cell '{reference_value}' not found.")
        return None, None, None
    
    color = get_fill_rgb(reference_cell)
    
    if not color:
        logging.debug(f"   Cell '{reference_value}' found at {reference_cell.coordinate} but no fill color.")
        return None, None, None
    
    logging.debug(f"   Cell '{reference_value}' found at {reference_cell.coordinate} with RGB color {color}")
    return color, reference_cell.row + 1, reference_cell.column

def get_columns_fill_color(sheet):
    #  Find the 'Columns' cell and return its fill color as RGB.
    return get_reference_color_info(sheet, "Columns")

def get_rows_fill_color(sheet):
    #  Find the 'Rows' cell and return its fill color as RGB.
    return get_reference_color_info(sheet, "Rows")

def extract_header_rows_with_color(sheet, reference_color, start_row, start_col):
    #  Extract all consecutive rows that have the same fill color as the reference.
    headers = []
    current_row = start_row

    while current_row <= sheet.max_row:
        first_cell = sheet.cell(current_row, start_col)
        row_fill_color = get_fill_rgb(first_cell)

        if row_fill_color != reference_color:
            break

        row_cells = []
        for col_idx in range(start_col, sheet.max_column + 1):
            cell = sheet.cell(row=current_row, column=col_idx)
            cell_data = {
                'value': cell.value,
                'coordinate': cell.coordinate,
                'border': cell.border
            }
            row_cells.append(cell_data)
            logging.debug(f"   Added cell '{cell.value}' at {cell.coordinate} ")

        headers.append(row_cells)
        current_row += 1

    return headers

def get_column_headers(sheet):
    #  Extract column headers based on the 'Columns' cell fill color.
    columns_fill_color, start_row, start_col = get_columns_fill_color(sheet)

    if not start_row or not columns_fill_color:
        logging.debug("   No 'Columns' cell or no fill color.")
        return [], 0

    headers = extract_header_rows_with_color(sheet, columns_fill_color, start_row, start_col)
    return headers, len(headers)


def adjust_column_labels(table_sheet, num_rows):
    # 🟢 1️⃣ Unmerge the initial region L15:L17
    l15_l17_range = "L15:L17"
    for merged_range in table_sheet.merged_cells.ranges:
        if merged_range.coord == l15_l17_range:
            table_sheet.unmerge_cells(l15_l17_range)
            logging.debug(f"   Unmerged {l15_l17_range} at start of method.")
            break

    label_row_index = None
    label_col_index = None  # Save this too!
    for row in table_sheet.iter_rows():
        for cell in row:
            if str(cell.value).startswith("columnLabel"):
                label_row_index = cell.row
                label_col_index = cell.column  # Save column index
                break
        if label_row_index:
            break

    if not label_row_index or not label_col_index:
        logging.debug("   Could not find columnLabel row in table sheet.")
        return

    logging.debug(f"   Found columnLabel starting at row {label_row_index}, column {label_col_index}")

    # 🟩 Save data validations of row 19 only
    saved_validations_row_19 = []
    for dv in table_sheet.data_validations.dataValidation:
        for cr in dv.cells.ranges:
            if cr.min_row <= 19 <= cr.max_row:
                saved_validations_row_19.append({
                    'dv': dv,
                    'min_col': cr.min_col,
                    'max_col': cr.max_col,
                    'type': dv.type,
                    'formula1': dv.formula1,
                    'formula2': dv.formula2,
                })
    logging.debug(f"   Saved {len(saved_validations_row_19)} data validations from row 19.")

    # Count current label rows
    current_label_rows = 0
    for row in table_sheet.iter_rows(min_row=label_row_index):
        label_row_found = any(str(cell.value).startswith("columnLabel") for cell in row if cell.value)
        if label_row_found:
            current_label_rows += 1
        else:
            break

    diff = num_rows - current_label_rows - 1  # Remove one extra row

    logging.debug(f"   Adjusting label rows: current={current_label_rows}, needed={num_rows}, diff={diff}")

    if diff < 0:
        rows_to_remove = abs(diff)
        for i in range(rows_to_remove):
            table_sheet.delete_rows(label_row_index - 1 + num_rows)
            logging.debug(f"   Deleted extra label row at index {label_row_index - 1 + num_rows}")
    elif diff > 0:
        for i in range(diff):
            template_row = table_sheet[label_row_index + current_label_rows - 1]
            insert_at = label_row_index + current_label_rows
            table_sheet.insert_rows(insert_at)

            for col_idx, cell in enumerate(template_row, start=1):
                new_cell = table_sheet.cell(row=insert_at, column=col_idx)
                new_cell.value = cell.value
                new_cell.fill = copy(cell.fill)
                new_cell.font = copy(cell.font)
                new_cell.alignment = copy(cell.alignment)
                new_cell.border = copy(cell.border)

            # ✅ Replace first cell with incremented label in the true column!
            new_label_num = current_label_rows + 1
            new_cell_label = table_sheet.cell(row=insert_at, column=label_col_index)
            new_cell_label.value = f"columnLabel {new_label_num}"

            logging.debug(f"   Inserted new label row at index {insert_at} "
                          f"with updated label '{new_cell_label.value}' in column {label_col_index}")

            current_label_rows += 1

    # 🟩 Reapply saved data validations for row 19 but shifted!
    if saved_validations_row_19:
        new_row = 19 + diff
        for saved in saved_validations_row_19:
            new_sqref = f"{openpyxl.utils.get_column_letter(saved['min_col'])}{new_row}:{openpyxl.utils.get_column_letter(saved['max_col'])}{new_row}"
            new_dv = copy(saved['dv'])
            new_dv.sqref = new_sqref
            table_sheet.add_data_validation(new_dv)
            logging.debug(f"   Re-applied data validation (type: {saved['type']}) "
                          f"to {new_sqref} (original row 19 -> new row {new_row})")

    # 🟩 Final label row values print-out
    final_row_index = label_row_index + current_label_rows - 1
    for cell in table_sheet[final_row_index]:
        logging.debug(f"   Final label row cell: {cell.coordinate} value: {cell.value}")

    # 🟢 2️⃣ Re-merge L15 down to include all empty cells below
    merge_start_row = 15
    merge_col = 12  # L column
    merge_end_row = merge_start_row

    max_row = table_sheet.max_row
    for row_idx in range(merge_start_row + 1, max_row + 1):
        cell = table_sheet.cell(row=row_idx, column=merge_col)
        if cell.value is None:
            merge_end_row = row_idx
        else:
            break

    if merge_end_row > merge_start_row:
        merge_range = f"L{merge_start_row}:L{merge_end_row}"
        table_sheet.merge_cells(merge_range)
        logging.debug(f"   Re-merged {merge_range} at end of method.")


###############################################################################################################################################################
#                                                                                                                                                             #  
#       4th part: pasting headers                                                                                                                             #  
#                                                                                                                                                             #  
#   Pastes the extracted headers into the correct cells in the target sheets.                                                                                 #  
#                                                                                                                                                             #                                                                                                                                               #  
###############################################################################################################################################################

def paste_column_headers(table_sheet, header_rows, start_row=15, start_col=14):

    if not header_rows:
        logging.debug("   No header rows to paste.")
        return

    # 1️⃣ Determine the maximum number of columns to paste
    nb_cols = max(len(row) for row in header_rows)
    logging.debug(f"   Number of columns to paste: {nb_cols}")

    # 2️⃣ Copy the border, style, and width from entire column N to target columns
    source_col_letter = get_column_letter(start_col)
    source_width = table_sheet.column_dimensions[source_col_letter].width
    max_row = table_sheet.max_row

    # 🟩 Identify data validations that include the source column anywhere in their sqref
    dvs_to_copy = []
    for dv in table_sheet.data_validations.dataValidation:
        for cr in dv.cells.ranges:
            if cr.min_col <= start_col <= cr.max_col:
                dvs_to_copy.append(dv)
                break

    logging.debug(f"   Found {len(dvs_to_copy)} data validations covering source column {source_col_letter}")

    for col_offset in range(nb_cols):
        target_col = start_col + col_offset
        target_col_letter = get_column_letter(target_col)

        # Copy width
        table_sheet.column_dimensions[target_col_letter].width = source_width

        # Copy styles row-by-row
        for row in range(1, max_row + 1):
            source_cell = table_sheet.cell(row=row, column=start_col)
            target_cell = table_sheet.cell(row=row, column=target_col)
            target_cell.border = copy(source_cell.border)
            target_cell.fill = copy(source_cell.fill)
            target_cell.font = copy(source_cell.font)
            target_cell.alignment = copy(source_cell.alignment)

        # 🟩 Apply data validations exactly matching row ranges
        for dv in dvs_to_copy:
            for cr in dv.cells.ranges:
                if cr.min_col <= start_col <= cr.max_col:
                    # 🟩 Create a new range for the target column, keeping the row range
                    new_sqref_list = []
                    for row_num in range(cr.min_row, cr.max_row + 1):
                        new_sqref_list.append(f"{target_col_letter}{row_num}")
                    new_sqref_str = " ".join(new_sqref_list)

                    new_dv = copy(dv)
                    new_dv.sqref = new_sqref_str
                    table_sheet.add_data_validation(new_dv)
                            
        logging.debug(f"   Completed column {target_col_letter}: width, style, and precise data validations copied from {source_col_letter}")

    # 3️⃣ Paste header values and override border style for those cells
    for r_idx, row_cells in enumerate(header_rows):
        for c_idx, cell_info in enumerate(row_cells):
            target_cell = table_sheet.cell(row=start_row + r_idx, column=start_col + c_idx)
            target_cell.value = cell_info['value']

            # Override border with the one from the source (pasted) cell
            if 'border' in cell_info and cell_info['border'] is not None:
                target_cell.border = copy(cell_info['border'])
                #  logging.debug(f"   Overrode border at {target_cell.coordinate} with source border.")

            logging.debug(f"   Pasted header style for'{cell_info['value']}' at {target_cell.coordinate}")



def extract_row_headers_with_indents(source_sheet, reference_color, start_row, column):
    #  Extract row headers and indent levels from a specific column.
    headers = []
    indents = []
    current_row = start_row
    
    while current_row <= source_sheet.max_row:
        cell = source_sheet.cell(row=current_row, column=column)
        cell_fill = get_fill_rgb(cell)

        # Stop if color no longer matches
        if cell_fill != reference_color:
            # Check if the cell is part of a merged range
            is_merged = any(cell.coordinate in merged_range for merged_range in source_sheet.merged_cells.ranges)
            if not is_merged:
                logging.debug(f"   Stopping at row {current_row} (fill mismatch: {cell_fill}, is_merged={is_merged}).")
                break

        header_value = cell.value if cell.value is not None else ""
        headers.append(header_value)
        
        if column == 2:  # Column B - extract indents
            indent_level = cell.alignment.indent if cell.alignment and cell.alignment.indent is not None else 0
            indents.append(indent_level)
            logging.debug(f"   Added header '{header_value}' with indent {indent_level} from {cell.coordinate}")
        else:  # Column C - no indents
            logging.debug(f"   Added row label header '{header_value}' from {cell.coordinate}")
        
        current_row += 1

    return headers, indents


def find_target_paste_rows(table_sheet, target_labels):
    #  Find the starting rows for pasting based on target labels.
    result = {}
    for row in table_sheet.iter_rows():
        for cell in row:
            if cell.value in target_labels:
                result[cell.value] = cell.row + 1
            if len(result) == len(target_labels):
                return result
    return result

def paste_row_headers(table_sheet, source_sheet):
    # Extract and paste row headers from source to target sheet
    # Find "Rows" cell and fill color using existing function
    rows_fill_color, start_row, _ = get_rows_fill_color(source_sheet)

    if not rows_fill_color or not start_row:
        logging.debug("   Could not find 'Rows' cell or no fill color in source.")
        return

    # Adjust start_row since get_rows_fill_color returns row+1
    start_row = start_row - 1

    # Extract headers and indent levels from columns B and C
    headers_B, indents_B = extract_row_headers_with_indents(source_sheet, rows_fill_color, start_row, 2)
    headers_C, _ = extract_row_headers_with_indents(source_sheet, rows_fill_color, start_row, 3)

    # Find target paste locations
    paste_locations = find_target_paste_rows(table_sheet, ["rowNumbering", "level"])
    start_row_numbering = paste_locations.get("rowNumbering")
    start_row_level = paste_locations.get("level")

    if not start_row_numbering or not start_row_level:
        logging.debug("   Could not find 'rowNumbering' or 'level' cells in target sheet.")
        return

    # Paste headers_B and levels below "rowNumbering" and "level"
    first_paste_row = start_row_numbering
    if start_row_numbering and start_row_level:
        for idx, (header, indent) in enumerate(zip(headers_B, indents_B)):
            # Map indent to level according to specifications
            if indent == 0:
                level = 1
            elif indent == 2:
                level = 2
            elif indent == 4:
                level = 3
            elif indent == 6:
                level = 4
            elif indent == 8:
                level = 5
            elif indent == 10:
                level = 6
            else:
                level = indent  # Default fallback for unexpected indent values

            # Paste header in column M
            header_cell = table_sheet.cell(row=start_row_numbering + idx, column=13)
            header_cell.value = header

            # Paste level in column A
            level_cell = table_sheet.cell(row=start_row_level + idx, column=1)
            level_cell.value = level

            logging.debug(f"   Pasted row header '{header}' with indent {indent} -> level '{level}' at rows {start_row_numbering + idx}")
 
    # Copy border and alignment style from the first pasted row (entire row)
    if first_paste_row:
        max_col = table_sheet.max_column

        # Collect border and alignment from first pasted row
        first_row_styles = [
            (copy(table_sheet.cell(row=first_paste_row, column=col).border),
            copy(table_sheet.cell(row=first_paste_row, column=col).alignment))
            for col in range(1, max_col + 1)
        ]

        # 🟩 Collect data validations that directly cover the first pasted row
        dvs_to_copy = []
        for dv in table_sheet.data_validations.dataValidation:
            for cr in dv.cells.ranges:
                if any(cell[0] == first_paste_row for cell in cr.cells):
                    dvs_to_copy.append(dv)
                    break

        logging.debug(f"   Found {len(dvs_to_copy)} data validations in first pasted row {first_paste_row}")

        # 🟩 Apply to each target row
        for idx in range(len(headers_B)):
            target_row = start_row_numbering + idx
            for col in range(1, max_col + 1):
                target_cell = table_sheet.cell(row=target_row, column=col)
                border_style, alignment_style = first_row_styles[col - 1]
                target_cell.border = copy(border_style)
                target_cell.alignment = copy(alignment_style)

            # 🟩 Reapply data validations for this row
            for dv in dvs_to_copy:
                new_dv = copy(dv)
                new_sqref_list = []
                for cr in dv.cells.ranges:
                    for col_num in range(cr.min_col, cr.max_col + 1):
                        col_letter = openpyxl.utils.get_column_letter(col_num)
                        new_sqref_list.append(f"{col_letter}{target_row}")
                new_sqref_str = " ".join(new_sqref_list)
                new_dv.sqref = new_sqref_str
                table_sheet.add_data_validation(new_dv)

                logging.debug(f"   Copied data validation to {new_sqref_str}")

            logging.debug(f"   Applied full row border, alignment, and data validations to row {target_row}")

    # Paste headers_C below "rowLabelColumnNumbering" in target column L
    start_row_label = None
    for row in table_sheet.iter_rows():
        for cell in row:
            if cell.value == "rowLabelColumnNumbering":
                start_row_label = cell.row + 1
                break
        if start_row_label:
            break

    if start_row_label:
        for idx, header in enumerate(headers_C):
            target_cell = table_sheet.cell(row=start_row_label + idx, column=12)
            target_cell.value = header
            logging.debug(f"   Pasted row label header '{header}' at {target_cell.coordinate} for rowLabelColumnNumbering")


def adjust_output_table_rows(table_sheet, internal_table_row_count):
    """
    Adjust the number of rows in the output table based on the internal table size.
    
    The output table starts at row 19 and initially has 3 rows (19, 20, 21).
    This function adjusts the table to match the number of rows in the internal table:
    - 1 row: delete rows 19 and 20 (leaving only row 21, which becomes row 19)
    - 2 rows: delete row 20 (leaving rows 19 and 21, which become rows 19-20)
    - 3 rows: do nothing (keep rows 19, 20, 21)
    - 4+ rows: insert rows above row 20, copying format from row 20
    
    Args:
        table_sheet: Target worksheet to adjust
        internal_table_row_count: Number of rows in the internal table
    """
    if internal_table_row_count == 0:
        logging.debug("   No rows in internal table, skipping row adjustment")
        return
    
    current_rows = 3  # Output table currently has rows 19, 20, 21
    start_row = 19
    
    if internal_table_row_count == 1:
        # Delete rows 19 and 20
        table_sheet.delete_rows(start_row, 2)
        logging.debug(f"   Deleted 2 rows (19-20) for 1-row internal table")
    
    elif internal_table_row_count == 2:
        # Delete row 20
        table_sheet.delete_rows(start_row + 1, 1)
        logging.debug(f"   Deleted 1 row (20) for 2-row internal table")
    
    elif internal_table_row_count == 3:
        # No adjustment needed
        logging.debug(f"   No row adjustment needed for 3-row internal table")
    
    elif internal_table_row_count >= 4:
        # Insert rows above row 20, copying format from row 20
        rows_to_insert = internal_table_row_count - current_rows
        insert_position = start_row + 1  # Insert above row 20 (which is at position start_row + 1)
        template_row = start_row + 1  # Row 20 is the template
        
        # Get the template row cells before insertion
        max_col = table_sheet.max_column
        template_cells = []
        for col in range(1, max_col + 1):
            cell = table_sheet.cell(row=template_row, column=col)
            template_cells.append({
                'value': cell.value,
                'fill': copy(cell.fill),
                'font': copy(cell.font),
                'alignment': copy(cell.alignment),
                'border': copy(cell.border),
                'number_format': cell.number_format
            })
        
        # Collect ALL data validations in the table area (any row, columns A-P)
        # We need to find DVs for columns E, F, etc. (but NOT column C, which will be copied from E later)
        dvs_by_column = {}
        dvs_to_remove = []
        table_max_col = 16  # Column P
        
        for dv in table_sheet.data_validations.dataValidation:
            should_remove_dv = False
            for cr in dv.cells.ranges:
                # Check if this DV intersects with table columns (1-16, i.e., A-P)
                if cr.min_col <= table_max_col and cr.max_col >= 1:
                    # Check if it intersects with or is near the table rows (within a reasonable range)
                    # Include DVs that cover rows around our table area (e.g., rows 1-1000)
                    if cr.min_row <= 1000 and cr.max_row >= start_row:
                        should_remove_dv = True
                        # Extract which columns this DV covers (excluding column C which will be handled separately)
                        for col in range(max(1, cr.min_col), min(table_max_col + 1, cr.max_col + 1)):
                            if col != 3 and col not in dvs_by_column:  # Skip column C
                                dvs_by_column[col] = dv
                                logging.debug(f"   Found data validation for column {openpyxl.utils.get_column_letter(col)} (DV range: {cr})")
            if should_remove_dv and dv not in dvs_to_remove:
                dvs_to_remove.append(dv)
        
        logging.debug(f"   Found data validations for columns: {[openpyxl.utils.get_column_letter(c) for c in sorted(dvs_by_column.keys())]}")
        logging.debug(f"   Will remove {len(dvs_to_remove)} old data validations before creating new ones")
        
        # Remove old data validations that covered the original table rows
        for dv in dvs_to_remove:
            table_sheet.data_validations.dataValidation.remove(dv)
        
        # Insert all required rows at once
        table_sheet.insert_rows(insert_position, rows_to_insert)
        logging.debug(f"   Inserted {rows_to_insert} rows starting at row {insert_position}")
        
        # Apply template formatting to all inserted rows
        for offset in range(rows_to_insert):
            current_row = insert_position + offset
            
            # Apply template formatting to the inserted row
            for col_idx, template_cell in enumerate(template_cells, start=1):
                new_cell = table_sheet.cell(row=current_row, column=col_idx)
                new_cell.value = template_cell['value']
                new_cell.fill = copy(template_cell['fill'])
                new_cell.font = copy(template_cell['font'])
                new_cell.alignment = copy(template_cell['alignment'])
                new_cell.border = copy(template_cell['border'])
                new_cell.number_format = template_cell['number_format']
            
            logging.debug(f"   Applied template format to row {current_row}")
        
        # Now create new data validations covering ALL table rows (19 through end)
        # Column C will be handled separately by copying from E
        final_row = start_row + internal_table_row_count - 1
        for col, original_dv in dvs_by_column.items():
            new_dv = copy(original_dv)
            col_letter = openpyxl.utils.get_column_letter(col)
            
            # Parse the original sqref and preserve non-table ranges
            # Original format might be like: "C19:C21 E2 E19:E21 N6:P6 N8:P8"
            # We need to replace E19:E21 with E19:E170, remove C19:C21 (will be added separately), keep E2, N6:P6, N8:P8
            original_ranges = str(original_dv.sqref).split()
            new_ranges = []
            
            for range_str in original_ranges:
                # Check if this range is in our target column and covers table rows (19+)
                if ':' in range_str:
                    start_cell, end_cell = range_str.split(':')
                    range_col = openpyxl.utils.coordinate_to_tuple(start_cell)[1]
                    range_start_row = openpyxl.utils.coordinate_to_tuple(start_cell)[0]
                    range_end_row = openpyxl.utils.coordinate_to_tuple(end_cell)[0]
                    
                    # Skip ranges that are in column C (will be handled separately)
                    if range_col == 3 and range_start_row >= start_row and range_end_row >= start_row:
                        continue
                    
                    # Skip ranges that are in our current column and cover table rows (we'll replace with extended range)
                    if range_col == col and range_start_row >= start_row and range_end_row >= start_row:
                        continue
                else:
                    # Single cell reference
                    range_col = openpyxl.utils.coordinate_to_tuple(range_str)[1]
                    range_row = openpyxl.utils.coordinate_to_tuple(range_str)[0]
                    
                    # Skip if it's in column C and in table rows
                    if range_col == 3 and range_row >= start_row:
                        continue
                    
                    # Skip if it's in our column and in table rows
                    if range_col == col and range_row >= start_row:
                        continue
                
                new_ranges.append(range_str)
            
            # Add the new extended range for the table rows
            new_table_range = f"{col_letter}{start_row}:{col_letter}{final_row}"
            new_ranges.append(new_table_range)
            
            # Set the new sqref
            new_dv.sqref = ' '.join(new_ranges)
            table_sheet.add_data_validation(new_dv)
            logging.debug(f"   Added data validation to column {col_letter}: {new_dv.sqref}")
        
        # Copy data validation from column E to column C (they have the same dropdown lists)
        # At this point, column E's validation has been fully processed with extended ranges
        if 5 in dvs_by_column:  # Column E exists
            # Find the newly created E validation (it's the last one added for column E)
            column_e_validation = None
            for dv in table_sheet.data_validations.dataValidation:
                # Check if this DV covers column E table rows
                dv_ranges = str(dv.sqref).split()
                for range_str in dv_ranges:
                    if f"E{start_row}:E{final_row}" in range_str or range_str.startswith("E") and ":" in range_str:
                        column_e_validation = dv
                        break
                if column_e_validation:
                    break
            
            if column_e_validation:
                new_dv_for_c = copy(column_e_validation)
                
                # Parse the E validation's sqref and replace E ranges with C ranges
                original_ranges = str(column_e_validation.sqref).split()
                new_ranges_for_c = []
                
                for range_str in original_ranges:
                    # Check if this range is in column E and covers table rows (19+)
                    if ':' in range_str:
                        start_cell, end_cell = range_str.split(':')
                        range_col = openpyxl.utils.coordinate_to_tuple(start_cell)[1]
                        range_start_row = openpyxl.utils.coordinate_to_tuple(start_cell)[0]
                        range_end_row = openpyxl.utils.coordinate_to_tuple(end_cell)[0]
                        
                        # Replace E table range with C table range
                        if range_col == 5 and range_start_row >= start_row and range_end_row >= start_row:
                            new_ranges_for_c.append(f"C{range_start_row}:C{range_end_row}")
                            continue
                    else:
                        # Single cell reference
                        range_col = openpyxl.utils.coordinate_to_tuple(range_str)[1]
                        range_row = openpyxl.utils.coordinate_to_tuple(range_str)[0]
                        
                        # Replace E single cell with C single cell (like E2 → C2)
                        if range_col == 5:
                            new_ranges_for_c.append(f"C{range_row}")
                            continue
                    
                    # Keep all other ranges as-is (N6:V6, N8:V8, etc.)
                    new_ranges_for_c.append(range_str)
                
                # Set the new sqref
                new_dv_for_c.sqref = ' '.join(new_ranges_for_c)
                table_sheet.add_data_validation(new_dv_for_c)
                logging.debug(f"   Copied data validation from column E to column C: {new_dv_for_c.sqref}")
            else:
                logging.warning(f"   Column E validation not found after processing, cannot copy to column C")
        else:
            logging.warning(f"   Column E data validation not found, cannot copy to column C")
        
        logging.debug(f"   Completed insertion of {rows_to_insert} rows for {internal_table_row_count}-row internal table")


def adjust_output_table_columns(table_sheet, internal_table_column_count):
    """
    Adjust the number of columns in the output table based on the internal table column count.
    
    The output table initially has 3 columns: N, O, P (columns 14, 15, 16).
    This function adjusts the table to match the number of columns in the internal table:
    - 1 column: delete columns N and O (leaving only P, which becomes N)
    - 2 columns: delete column O (leaving N and P, which become N-O)
    - 3 columns: do nothing (keep N, O, P)
    - 4+ columns: insert columns to the left of O, copying format from O and data validations from rows 6, 8, 9
    
    Args:
        table_sheet: Target worksheet to adjust
        internal_table_column_count: Number of columns in the internal table
    """
    if internal_table_column_count == 0:
        logging.debug("   No columns in internal table, skipping column adjustment")
        return
    
    current_cols = 3  # Output table currently has columns N, O, P (14, 15, 16)
    col_n = 14  # Column N
    col_o = 15  # Column O
    col_p = 16  # Column P
    
    if internal_table_column_count == 1:
        # Delete columns N and O
        table_sheet.delete_cols(col_n, 2)
        logging.debug(f"   Deleted 2 columns (N, O) for 1-column internal table")
    
    elif internal_table_column_count == 2:
        # Delete column O
        table_sheet.delete_cols(col_o, 1)
        logging.debug(f"   Deleted 1 column (O) for 2-column internal table")
    
    elif internal_table_column_count == 3:
        # No adjustment needed
        logging.debug(f"   No column adjustment needed for 3-column internal table")
    
    elif internal_table_column_count >= 4:
        # Insert columns to the left of column O, copying format from column O
        cols_to_insert = internal_table_column_count - current_cols
        insert_position = col_o  # Insert before column O (column 15)
        template_col = col_o  # Column O is the template
        
        # Store the original column width from column N (all columns N, O, P should have same width)
        original_width = table_sheet.column_dimensions[openpyxl.utils.get_column_letter(col_n)].width
        logging.debug(f"   Original column width: {original_width}")
        
        # Get the template column cells before insertion
        max_row = table_sheet.max_row
        template_cells = []
        for row in range(1, max_row + 1):
            cell = table_sheet.cell(row=row, column=template_col)
            template_cells.append({
                'value': cell.value,
                'fill': copy(cell.fill),
                'font': copy(cell.font),
                'alignment': copy(cell.alignment),
                'border': copy(cell.border),
                'number_format': cell.number_format
            })
        
        # Collect data validations for rows 6, 8, and 9 from columns N, O, or P
        dvs_to_copy = {}  # Map from DV object id to the DV itself
        dvs_to_remove = []
        target_rows = [6, 8, 9]
        
        logging.debug(f"   Looking for data validations in target rows {target_rows}, columns N-P ({col_n}-{col_p})")
        logging.debug(f"   Total data validations in sheet: {len(list(table_sheet.data_validations.dataValidation))}")
        
        for dv in table_sheet.data_validations.dataValidation:
            logging.debug(f"   Checking DV with sqref: {dv.sqref}")
            found_target_row = False
            for cr in dv.cells.ranges:
                logging.debug(f"     Range: {cr}, min_col={cr.min_col}, max_col={cr.max_col}, min_row={cr.min_row}, max_row={cr.max_row}")
                # Check if this DV covers any of the columns N, O, or P (before insertion)
                if cr.min_col <= col_p and cr.max_col >= col_n:
                    logging.debug(f"       Range overlaps columns N-P")
                    # Check if it covers any of our target rows
                    for target_row in target_rows:
                        if cr.min_row <= target_row and cr.max_row >= target_row:
                            logging.debug(f"         Range covers target row {target_row}")
                            found_target_row = True
                            break
                    if found_target_row:
                        break
            
            if found_target_row:
                # Use the object id as key to avoid duplicates
                dv_id = id(dv)
                if dv_id not in dvs_to_copy:
                    dvs_to_copy[dv_id] = dv
                    logging.debug(f"   Found data validation (DV sqref: {dv.sqref})")
                # Mark this entire DV for removal since we'll recreate it with updated ranges
                if dv not in dvs_to_remove:
                    dvs_to_remove.append(dv)
        
        logging.debug(f"   Found {len(dvs_to_copy)} unique data validations to process")
        logging.debug(f"   Will remove {len(dvs_to_remove)} old data validations before creating new ones")
        
        # Remove old data validations that covered the original table columns
        for dv in dvs_to_remove:
            table_sheet.data_validations.dataValidation.remove(dv)
        
        # Insert all required columns at once
        table_sheet.insert_cols(insert_position, cols_to_insert)
        logging.debug(f"   Inserted {cols_to_insert} columns starting at column {openpyxl.utils.get_column_letter(insert_position)}")
        
        # Apply template formatting to all inserted columns
        for offset in range(cols_to_insert):
            current_col = insert_position + offset
            col_letter = openpyxl.utils.get_column_letter(current_col)
            
            # Set the column width to match the original width
            table_sheet.column_dimensions[col_letter].width = original_width
            
            # Apply template formatting to the inserted column
            for row_idx, template_cell in enumerate(template_cells, start=1):
                new_cell = table_sheet.cell(row=row_idx, column=current_col)
                new_cell.value = template_cell['value']
                new_cell.fill = copy(template_cell['fill'])
                new_cell.font = copy(template_cell['font'])
                new_cell.alignment = copy(template_cell['alignment'])
                new_cell.border = copy(template_cell['border'])
                new_cell.number_format = template_cell['number_format']
            
            logging.debug(f"   Applied template format and width to column {col_letter}")
        
        # Set width for the original O and P columns which got shifted after insertion
        # After insertion, original O is now at (insert_position + cols_to_insert)
        # and original P is at (insert_position + cols_to_insert + 1)
        shifted_o_col = insert_position + cols_to_insert
        shifted_p_col = insert_position + cols_to_insert + 1
        table_sheet.column_dimensions[openpyxl.utils.get_column_letter(shifted_o_col)].width = original_width
        table_sheet.column_dimensions[openpyxl.utils.get_column_letter(shifted_p_col)].width = original_width
        logging.debug(f"   Set width for shifted columns {openpyxl.utils.get_column_letter(shifted_o_col)} and {openpyxl.utils.get_column_letter(shifted_p_col)}")
        
        # Now create new data validations covering ALL table columns (N through final column)
        # Process each unique DV once and extend ALL target row ranges (6, 8, 9) in columns N-P
        final_col = col_n + internal_table_column_count - 1
        for dv_id, original_dv in dvs_to_copy.items():
            new_dv = copy(original_dv)
            start_col_letter = openpyxl.utils.get_column_letter(col_n)
            end_col_letter = openpyxl.utils.get_column_letter(final_col)
            
            # Parse the original sqref and update ALL ranges in columns N-P for target rows
            # Original format might be like: "C2 C19:C170 E2 E19:E170 N6:P6 N8:P8 N9:P9"
            # We need to replace ALL N6:P6, N8:P8, N9:P9 with N6:V6, N8:V8, N9:V9
            original_ranges = str(original_dv.sqref).split()
            new_ranges = []
            extended_rows = set()  # Track which target rows we've extended
            
            for range_str in original_ranges:
                # Check if this range covers any target row in columns N-P
                if ':' in range_str:
                    start_cell, end_cell = range_str.split(':')
                    start_col = openpyxl.utils.column_index_from_string(openpyxl.utils.get_column_letter(openpyxl.utils.coordinate_to_tuple(start_cell)[1]))
                    end_col = openpyxl.utils.column_index_from_string(openpyxl.utils.get_column_letter(openpyxl.utils.coordinate_to_tuple(end_cell)[1]))
                    start_row = openpyxl.utils.coordinate_to_tuple(start_cell)[0]
                    end_row = openpyxl.utils.coordinate_to_tuple(end_cell)[0]
                    
                    # Check if this is a target row range in columns N-P
                    is_target_range = False
                    for target_row in target_rows:
                        if start_row <= target_row <= end_row and start_col >= col_n and end_col <= col_p:
                            # Skip this range, we'll add the extended version later
                            is_target_range = True
                            extended_rows.add(target_row)
                            break
                    
                    if not is_target_range:
                        new_ranges.append(range_str)
                else:
                    # Single cell reference
                    col_idx = openpyxl.utils.coordinate_to_tuple(range_str)[1]
                    row_idx = openpyxl.utils.coordinate_to_tuple(range_str)[0]
                    
                    # Skip if it's a target row in columns N-P
                    if row_idx in target_rows and col_n <= col_idx <= col_p:
                        extended_rows.add(row_idx)
                        continue
                    
                    new_ranges.append(range_str)
            
            # Add the extended ranges for ALL target rows found in this DV
            for target_row in sorted(extended_rows):
                new_table_range = f"{start_col_letter}{target_row}:{end_col_letter}{target_row}"
                new_ranges.append(new_table_range)
            
            # Set the new sqref
            new_dv.sqref = ' '.join(new_ranges)
            table_sheet.add_data_validation(new_dv)
            logging.debug(f"   Added data validation: {new_dv.sqref}")
        
        logging.debug(f"   Completed insertion of {cols_to_insert} columns for {internal_table_column_count}-column internal table")



###############################################################################################################################################################
#                                                                                                                                                             #  
#       5th part: post-processings and filling in the identification and labelTranslation sheets.                                                             #  
#                                                                                                                                                             #  
#   Handles the external borders and merged cells                                                                                                             #  
#   Add colors to cells (manual/not required)                                                                                                                 #  
#   Fills in the identification and labelTranslation sheets
#   Move 'cellSpecificities' and 'labelTranslation' sheets to the end of the workbook.                                                                                                   # 
#   Reopen Excel for border compatibility with win32com on Windows as openpyxl known issue with border                                                        # 
#                                                                                                                                                             #  
###############################################################################################################################################################

def create_standard_fill_styles():
    #  Create standard fill styles used throughout the application.
    return {
        'yellow': PatternFill(start_color="FFFFFFCC", end_color="FFFFFFCC", fill_type="solid"),
        'grey': PatternFill(start_color="FF808080", end_color="FF808080", fill_type="solid")
    }

def has_heavy_border(cell, direction):
    #  Check if a cell has a heavy border in the specified direction.
    border = cell.border
    if direction == "right":
        return border.right and border.right.style in ["medium", "thick", "double"]
    elif direction == "bottom":
        return border.bottom and border.bottom.style in ["medium", "thick", "double"]
    elif direction == "left":
        return border.left and border.left.style in ["medium", "thick", "double"]
    elif direction == "top":
        return border.top and border.top.style in ["medium", "thick", "double"]
    return False

def find_table_boundaries(sheet, start_row, start_col):
    #  Find the boundaries of a table by following borders.
    # Determine width by following top borders
    width = 1  # include starting cell
    current_col = start_col
    while True:
        next_col = current_col + 1
        cell = sheet.cell(row=start_row, column=next_col)
        if cell.border.top and cell.border.top.style:
            width += 1
            current_col = next_col
        else:
            break

    # Determine height by following left borders
    height = 1  # include starting cell
    current_row = start_row
    while True:
        next_row = current_row + 1
        cell = sheet.cell(row=next_row, column=start_col)
        if cell.border.left and cell.border.left.style:
            height += 1
            current_row = next_row
        else:
            break
    
    return width, height

def apply_border_to_rectangle(sheet, start_row, start_col, width, height, border_style='medium'):
    # Apply borders to complete a rectangular table
    border_side = Side(style=border_style)
    
    # Draw bottom border on the bottom row
    bottom_row = start_row + height - 1
    for col in range(start_col, start_col + width):
        cell = sheet.cell(row=bottom_row, column=col)
        cell.border = Border(
            left=cell.border.left,
            right=cell.border.right,
            top=cell.border.top,
            bottom=border_side
        )

    # Draw right border on the rightmost column
    right_col = start_col + width - 1
    for row in range(start_row, start_row + height):
        cell = sheet.cell(row=row, column=right_col)
        cell.border = Border(
            left=cell.border.left,
            right=border_side,
            top=cell.border.top,
            bottom=cell.border.bottom
        )

def finalize_table_rectangle(table_sheet):
    # Finalize table rectangle by adding bottom and right borders
    start_row = 15
    start_col = 12  # L column
    
    # Find table boundaries using helper function
    width, height = find_table_boundaries(table_sheet, start_row, start_col)
    
    logging.debug(f"   Final rectangle: top-left={table_sheet.cell(row=start_row, column=start_col).coordinate}, width={width}, height={height}")
    
    # Apply borders using helper function
    apply_border_to_rectangle(table_sheet, start_row, start_col, width, height)



def find_rectangle_boundaries_by_borders(sheet, start_row, start_col):
    # Find the end row and column of a rectangle by detecting heavy borders
    end_row = start_row
    end_col = start_col
    max_row = sheet.max_row
    max_col = sheet.max_column

    found_corner = False
    for r in range(start_row, max_row + 1):
        for c in range(start_col, max_col + 1):
            cell = sheet.cell(row=r, column=c)
            if has_heavy_border(cell, "right") and has_heavy_border(cell, "bottom"):
                end_row = r
                end_col = c
                found_corner = True
                break
        if found_corner:
            break
    
    return end_row, end_col

def build_cell_mappings(sheet, start_row, start_col, end_row, end_col):
    # Build column and row mappings for a table region
    # Build column mapping from header row above the rectangle
    col_mapping = {}
    for col in sheet.iter_cols(min_row=start_row - 1, max_row=start_row - 1, 
                              min_col=start_col, max_col=end_col):
        for cell in col:
            if cell.value:
                col_mapping[str(cell.value)] = cell.column

    # Find the column where 'rowNumbering' is located
    row_numbering_col = None
    for row in sheet.iter_rows():
        for cell in row:
            if cell.value == "rowNumbering":
                row_numbering_col = cell.column
                break
        if row_numbering_col:
            break

    # Build row mapping by scanning down the 'rowNumbering' column
    row_mapping = {}
    if row_numbering_col:
        r = start_row
        while r <= sheet.max_row:
            cell = sheet.cell(row=r, column=row_numbering_col)
            if cell.value:
                row_mapping[str(cell.value)] = r
            # Stop if heavy bottom border
            if has_heavy_border(cell, "bottom"):
                break
            r += 1

    return col_mapping, row_mapping

def get_number_format_for_symbol(symbol_value):
    # Get the appropriate number format based on a symbol value
    format_type = symbol_value.lower()
    
    if any(sym in format_type for sym in ["€", "£", "$", "mi"]):
        return '#,##0.00'
    elif "ri" in format_type:
        return '#,##0.0000'
    elif any(sym in format_type for sym in ["ii", "#"]):
        return '#,##0'
    elif any(sym in format_type for sym in ["pi", "%"]):
        return '# ##0.00%'
    elif "date" in format_type or "di" in format_type:
        return 'yyyy-mm-dd'
    
    return None

def should_apply_grey_fill(cell_value, cell_color):
    # Determine if a cell should have grey fill applied
    special_values = ["#", "%", "€£$", "text", "TRUE/FALSE"]
    has_grey_color = cell_color == 55 or cell_color == "55"
    
    return cell_value in special_values or has_grey_color

def apply_cell_type(source_sheet, target_sheet):
    # Apply cell types (colors and number formats) from source to target sheet
    # Create standard fill styles
    fill_styles = create_standard_fill_styles()

    # Locate the start of the rectangle: below & right of 'rowLabelColumnNumbering'
    start_row = None
    start_col = None
    for row in target_sheet.iter_rows():
        for cell in row:
            if cell.value == "rowLabelColumnNumbering":
                start_row = cell.row + 1
                start_col = cell.column + 1
                break
        if start_row:
            break

    if not start_row or not start_col:
        return

    # Find rectangle boundaries and build mappings
    end_row, end_col = find_rectangle_boundaries_by_borders(target_sheet, start_row, start_col)
    col_mapping, row_mapping = build_cell_mappings(target_sheet, start_row, start_col, end_row, end_col)

    # Apply yellow fill to entire rectangle
    for r in range(start_row, end_row + 1):
        for c in range(start_col, end_col + 1):
            cell = target_sheet.cell(row=r, column=c)
            cell.fill = copy(fill_styles['yellow'])

    # Get reference colors for dynamic positioning
    columns_fill_color, _, _ = get_columns_fill_color(source_sheet)
    rows_fill_color, _, _ = get_rows_fill_color(source_sheet)

    logging.debug(f"   col_mapping: {col_mapping}")
    logging.debug(f"   row_mapping: {row_mapping}")

    # Process source sheet cells for formatting
    for row in source_sheet.iter_rows():
        for cell in row:
            val = str(cell.value).strip() if cell.value is not None else ""
            if not val:
                continue

            # Get dynamic row_id and col_id
            row_id = find_cell_id_by_color(source_sheet, cell, rows_fill_color, "left")
            col_id = find_cell_id_by_color(source_sheet, cell, columns_fill_color, "up")

            if not (row_id and col_id and row_id in row_mapping and col_id in col_mapping):
                continue

            target_r = row_mapping[row_id]
            target_c = col_mapping[col_id]
            target_cell = target_sheet.cell(row=target_r, column=target_c)

            # Apply grey fill for special symbols
            cell_color = cell.fill.fgColor.indexed if cell.fill and cell.fill.fgColor.type == "indexed" else None
            if should_apply_grey_fill(val, cell_color):
                target_cell.fill = copy(fill_styles['grey'])
                logging.debug(f"   Applied grey fill at {target_cell.coordinate} for '{val}'")

            # Apply number formatting
            number_format = get_number_format_for_symbol(val)
            if number_format:
                target_cell.number_format = number_format
                logging.debug(f"   Set number_format '{number_format}' at {target_cell.coordinate} for '{val}'")

    logging.debug("   Finished applying cell types.")

def find_cell_id_by_color(sheet, cell, reference_color, direction):
    # Find a cell ID by searching in a direction for a matching color
    if direction == "left":
        for c in range(cell.column, 0, -1):
            search_cell = sheet.cell(row=cell.row, column=c)
            cell_color = search_cell.fill.fgColor.rgb if search_cell.fill and search_cell.fill.fgColor.type == "rgb" else None
            if cell_color == reference_color and search_cell.value:
                return str(search_cell.value)
    
    elif direction == "up":
        for r in range(cell.row, 0, -1):
            search_cell = sheet.cell(row=r, column=cell.column)
            cell_color = search_cell.fill.fgColor.rgb if search_cell.fill and search_cell.fill.fgColor.type == "rgb" else None
            if cell_color == reference_color and search_cell.value:
                return str(search_cell.value)
    
    return None




def handle_open_rows(source_sheet, table_sheet):
    # Handle 'Open Rows' entries by setting rowNumbering to '999'
    # Find required columns
    columns = find_table_columns(table_sheet, "rowLabelColumnNumbering", "rowNumbering")
    
    if "rowLabelColumnNumbering" not in columns or "rowNumbering" not in columns:
        logging.debug("   Could not find required columns")
        return
    
    label_row, row_label_col = columns["rowLabelColumnNumbering"]
    _, row_numbering_col = columns["rowNumbering"]

    # Look for 'Open Rows' entries and add '999' under rowNumbering
    current_row = label_row + 1
    while current_row <= table_sheet.max_row:
        label_cell = table_sheet.cell(row=current_row, column=row_label_col)
        if label_cell.value and str(label_cell.value).strip().lower() == "open rows":
            num_cell = table_sheet.cell(row=current_row, column=row_numbering_col)
            num_cell.value = "999"
            logging.debug(f"   Found 'Open Rows' at row {current_row}, set rowNumbering=999")
        current_row += 1

def find_table_columns(sheet, *column_names):
    # Find columns for specific header names in a table
    result = {}
    for row in sheet.iter_rows():
        for cell in row:
            if cell.value in column_names:
                result[cell.value] = (cell.row, cell.column)
                if len(result) == len(column_names):
                    return result
    return result

def extract_numeric_prefix(text):
    # Extract numeric prefix from text
    if not text:
        return None, text
    
    prefix_match = re.match(r"^\d+", str(text).strip())
    if prefix_match:
        prefix = prefix_match.group(0)
        remaining = str(text)[len(prefix):].lstrip(" ")
        return prefix, remaining
    
    return None, text

   

def mark_abstract_rows(table_sheet):
    # 🟢 Find "rowNumbering" column index and row
    row_numbering_row = None
    row_numbering_col = None
    for row in table_sheet.iter_rows():
        for cell in row:
            if cell.value == "rowNumbering":
                row_numbering_row = cell.row
                row_numbering_col = cell.column
                logging.debug(f"   Found 'rowNumbering' at {cell.coordinate}")
                break
        if row_numbering_row:
            break

    if not row_numbering_row or not row_numbering_col:
        logging.debug("   Could not find 'rowNumbering' cell.")
        return

    # 🟢 Find "rowLabelColumnNumbering" column
    row_label_col = None
    for row in table_sheet.iter_rows(min_row=row_numbering_row, max_row=row_numbering_row):
        for cell in row:
            if cell.value == "rowLabelColumnNumbering":
                row_label_col = cell.column
                logging.debug(f"   Found 'rowLabelColumnNumbering' at {cell.coordinate}")
                break
        if row_label_col:
            break

    if not row_label_col:
        logging.debug("   Could not find 'rowLabelColumnNumbering' column.")
        return

    # Create standard fill styles using helper function
    fill_styles = create_standard_fill_styles()

    # Find required columns using helper function
    columns = find_table_columns(table_sheet, "rowNumbering", "rowLabelColumnNumbering", "abstractRow")
    
    if len(columns) < 3:
        logging.debug("   Could not find all required columns for abstract rows.")
        return

    row_numbering_row, row_numbering_col = columns["rowNumbering"]
    _, row_label_col = columns["rowLabelColumnNumbering"]
    _, abstract_col = columns["abstractRow"]

    # Process rows below header
    current_row = row_numbering_row + 1
    while current_row < table_sheet.max_row:
        row_num_cell = table_sheet.cell(row=current_row, column=row_numbering_col)
        
        # Check if rowNumbering cell is empty
        if not row_num_cell.value or str(row_num_cell.value).strip() == "":
            # Mark as abstract row
            abstract_cell = table_sheet.cell(row=current_row, column=abstract_col)
            abstract_cell.value = "yes"

            # Extract numeric prefix from rowLabelColumnNumbering
            row_label_cell = table_sheet.cell(row=current_row, column=row_label_col)
            if row_label_cell.value:
                prefix, remaining_text = extract_numeric_prefix(row_label_cell.value)
                if prefix:
                    row_num_cell.value = prefix
                    row_label_cell.value = remaining_text
                    logging.debug(f"   Extracted prefix '{prefix}' from row {current_row}")

            # Apply grey fill to cells to the right of rowNumbering (excluding next column)
            for col in range(row_numbering_col + 2, table_sheet.max_column + 1):
                target_cell = table_sheet.cell(row=current_row, column=col)
                target_cell.fill = copy(fill_styles['grey'])

            logging.debug(f"   Marked row {current_row} as abstract")
        
        current_row += 1



def remove_extra_right_columns(target_sheet):
    # Remove columns to the right of the table's bottom-right corner (determined by heavy right and bottom borders)
    max_row = target_sheet.max_row
    max_col = target_sheet.max_column

    # 🟢 Locate top-left of the table (after rowLabelColumnNumbering)
    start_row = None
    start_col = None
    for row in target_sheet.iter_rows():
        for cell in row:
            if cell.value == "rowLabelColumnNumbering":
                start_row = cell.row + 1
                start_col = cell.column + 1
                logging.debug(f"   Found 'rowLabelColumnNumbering' at {cell.coordinate}, start at {start_row},{start_col}")
                break
        if start_row:
            break

    if not start_row or not start_col:
        logging.debug("   Could not find 'rowLabelColumnNumbering' in target.")
        return

    # 🟢 Find bottom-right cell of the table using heavy borders
    def has_heavy_border(cell, direction):
        border = cell.border
        if direction == "right":
            return border.right and border.right.style in ["medium", "thick", "double"]
        if direction == "bottom":
            return border.bottom and border.bottom.style in ["medium", "thick", "double"]
        return False

    end_col = start_col
    end_row = start_row
    found_corner = False
    for r in range(start_row, max_row + 1):
        for c in range(start_col, max_col + 1):
            cell = target_sheet.cell(row=r, column=c)
            if has_heavy_border(cell, "right") and has_heavy_border(cell, "bottom"):
                end_row = r
                end_col = c
                found_corner = True
                logging.debug(f"   Found bottom-right corner at {cell.coordinate}")
                break
        if found_corner:
            break

    # 🟢 Remove columns to the right of end_col
    if end_col < max_col:
        cols_to_delete = max_col - end_col
        for i in range(cols_to_delete):
            target_sheet.delete_cols(end_col + 1)
    logging.debug(f"   Deleted {cols_to_delete} extra columns after column {end_col}")

    logging.debug("   Extra right columns removed.")



def delete_empty_rows(table_sheet):
    # 🔍 Locate the row and column of key headers
    row_numbering_col = None
    row_label_col = None
    header_row = None

    for row in table_sheet.iter_rows():
        for cell in row:
            if cell.value == "rowNumbering":
                row_numbering_col = cell.column
                header_row = cell.row
            elif cell.value == "rowLabelColumnNumbering":
                row_label_col = cell.column
        if row_numbering_col and row_label_col:
            break

    if not row_numbering_col or not row_label_col:
        logging.debug("   Could not find key header columns.")
        return

    # 🧹 Loop from bottom up to safely delete rows
    for row in range(table_sheet.max_row, header_row, -1):
        row_numbering_val = table_sheet.cell(row=row, column=row_numbering_col).value
        row_label_val = table_sheet.cell(row=row, column=row_label_col).value

        if not row_numbering_val and not row_label_val:
            # ✅ Copy bottom border of deleted row to the row above
            for col in range(1, table_sheet.max_column + 1):
                cell_to_delete = table_sheet.cell(row=row, column=col)
                cell_above = table_sheet.cell(row=row - 1, column=col)

                bottom_border = cell_to_delete.border.bottom
                if bottom_border and bottom_border.style:
                    existing_border = cell_above.border
                    updated_border = Border(
                        left=existing_border.left,
                        right=existing_border.right,
                        top=existing_border.top,
                        bottom=bottom_border
                    )
                    cell_above.border = updated_border

            table_sheet.delete_rows(row)
            logging.debug(f"   Deleted row {row} and copied its bottom borders to row {row-1}")




def merge_header_cells(table_sheet):
    start_row = 15
    start_col = 14  # Column N
    label_col = 13  # Column M
    max_row = table_sheet.max_row
    max_col = table_sheet.max_column

    def has_bottom_border(cell):
        return cell.border and cell.border.bottom and cell.border.bottom.style == BORDER_THIN

    # Find end_row before first 'rowLabelColumnNumbering' in column M
    end_row = start_row
    for row in range(start_row, max_row + 1):
        val = table_sheet.cell(row=row, column=label_col).value
        if isinstance(val, str) and val.strip().startswith("rowLabelColumnNumbering"):
            break
        end_row = row

    logging.debug(f"   Processing rows {start_row} to {end_row}")

    # --- VERTICAL MERGING ---
    for col in range(start_col, max_col + 1):
        merge_start = None
        for row in range(start_row, end_row + 1):
            cell = table_sheet.cell(row=row, column=col)
            above = table_sheet.cell(row=row - 1, column=col)
            if cell.value in (None, ''):
                if row > start_row and above.value not in (None, '') and not has_bottom_border(above):
                    if merge_start is None:
                        merge_start = row - 1
            else:
                if merge_start is not None and row - 1 > merge_start:
                    table_sheet.merge_cells(start_row=merge_start, start_column=col,
                                            end_row=row - 1, end_column=col)
                    logging.debug(f"   Vertically merged: {get_column_letter(col)}{merge_start}:{get_column_letter(col)}{row - 1}")
                merge_start = None
        if merge_start is not None and end_row > merge_start:
            table_sheet.merge_cells(start_row=merge_start, start_column=col,
                                    end_row=end_row, end_column=col)
            logging.debug(f"   Vertically merged: {get_column_letter(col)}{merge_start}:{get_column_letter(col)}{end_row}")

    # --- HORIZONTAL MERGING ---
    for row in range(start_row, end_row + 1):
        merge_start = None
        for col in range(start_col, max_col + 2):  # +2 for sentinel
            cell = table_sheet.cell(row=row, column=col) if col <= max_col else None
            left_cell = table_sheet.cell(row=row, column=col - 1) if col > start_col else None

            if col > max_col or (cell and cell.value not in (None, '') and not isinstance(cell, MergedCell)):
                if merge_start is not None and col - 1 > merge_start:
                    # Prevent merging a single cell
                   if col - 1 != merge_start:
                        if not is_overlapping_merge(merge_start, col - 1, row, table_sheet.merged_cells.ranges):
                            table_sheet.merge_cells(start_row=row, start_column=merge_start,
                                                    end_row=row, end_column=col - 1)
                            logging.debug(f"   Horizontally merged: {get_column_letter(merge_start)}{row}:{get_column_letter(col - 1)}{row}")
                        else:
                            logging.debug(f"   Skipped merge due to overlap at: {get_column_letter(merge_start)}{row}:{get_column_letter(col - 1)}{row}")

                merge_start = None
            elif cell and cell.value in (None, '') and not any(f"{get_column_letter(col)}{row}" in str(rng) for rng in table_sheet.merged_cells.ranges):

                if left_cell and left_cell.value not in (None, ''):
                    coord_left = f"{get_column_letter(col - 1)}{row}"
                    is_merged = any(coord_left in str(rng) for rng in table_sheet.merged_cells.ranges)
                    if is_merged:
                        if not isinstance(cell, MergedCell):
                            cell.value = left_cell.value
                            logging.debug(f"   Copied value from {coord_left} to {get_column_letter(col)}{row}")
                    else:
                        if merge_start is None:
                            merge_start = col - 1
            else:
                if merge_start is not None and col - 1 > merge_start:
                    table_sheet.merge_cells(start_row=row, start_column=merge_start,
                                            end_row=row, end_column=col - 1)
                    logging.debug(f"   Horizontally merged 2: {get_column_letter(merge_start)}{row}:{get_column_letter(col - 1)}{row}")
                merge_start = None

    # Apply vertical center + left horizontal alignment
    header_alignment = Alignment(horizontal='left', vertical='center')

    for row in range(start_row, end_row + 1):
        for col in range(start_col, max_col + 1):
            cell = table_sheet.cell(row=row, column=col)
            if not isinstance(cell, MergedCell):
                cell.alignment = header_alignment


def is_overlapping_merge(start_col, end_col, row, merged_ranges):
    for rng in merged_ranges:
        min_col, min_row, max_col, max_row = range_boundaries(str(rng))
        if row >= min_row and row <= max_row:
            if not (end_col < min_col or start_col > max_col):
                return True
    return False


def fill_identification_and_labelTranslation(source_wb, target_wb,
                                             form_group_code, version, validity_starting_date,
                                             linked_domain, datatype, roundingUnit):
    # 🟢 Find "Description (form label)" in the Properties sheet
    extracted_text = None
    
    if "Properties" in source_wb.sheetnames:
        properties_sheet = source_wb["Properties"]
        
        # Search for "Description (form label)" cell
        for row in properties_sheet.iter_rows():
            for cell in row:
                if cell.value and "Description (form label)" in str(cell.value):
                    # Get the value from the cell to the right
                    right_cell = properties_sheet.cell(row=cell.row, column=cell.column + 1)
                    extracted_text = right_cell.value
                    logging.debug(f"   Found 'Description (form label)' at {cell.coordinate}")
                    logging.debug(f"   Extracted text from {right_cell.coordinate}: '{extracted_text}'")
                    break
            if extracted_text:
                break
    
    if not extracted_text:
        logging.debug("   Could not find 'Description (form label)' in Properties sheet or no value to the right.")
        return None

    # 🟢 Fill identification sheet in the target workbook
    ident_sheet = target_wb["identification"]
    ident_sheet["D2"] = extracted_text
    ident_sheet["E2"] = extracted_text
    ident_sheet["G2"] = "direct"

    ident_sheet["A2"] = form_group_code
    ident_sheet["B2"] = version
    ident_sheet["C2"] = validity_starting_date
    ident_sheet["F2"] = linked_domain
    ident_sheet["J2"] = datatype
    ident_sheet["L2"] = roundingUnit

    logging.debug("   identification sheet updated.")
    
    # 🟢 Fill labelTranslation sheet in the target workbook
    label_translation_sheet = target_wb["labelTranslation"]
    label_translation_sheet["B1"] = "FR"
    label_translation_sheet["A2"] = extracted_text
    label_translation_sheet["B2"] = extracted_text
    
    # Apply wrap text formatting to A2 and B2
    wrap_alignment = Alignment(wrap_text=True)
    label_translation_sheet["A2"].alignment = wrap_alignment
    label_translation_sheet["B2"].alignment = wrap_alignment
    
    logging.debug("   labelTranslation sheet updated.")
    
    return extracted_text


def fill_sheet_labels(source_wb, target_wb, sheet_list, sheet_name_mapping, base_text):
    """
    Fill A2 and B2 cells in SA.*.* sheets with conditional text based on sheet patterns.
    
    This function populates cells A2 and B2 in each output SA.*.* sheet with text that
    starts with the base description from the identification sheet (D2/E2 value), and
    adds conditional suffixes based on the input sheet configuration:
    
    Conditional logic for suffixes:
    1. "Sheet T1" → always adds " - Feuillet technique"
    2. "Sheet 99" (2-digit pattern):
       a) If only 1 sheet total → no suffix
       b) If 2 sheets total with one being "Sheet T1" → no suffix
       c) If 2+ sheets matching "Sheet 99" pattern → adds " - {sheet_label_value}"
          where sheet_label_value is extracted from the SECOND cell to the right of
          the cell containing "Sheet label" text in the input sheet (typically AJ42
          when "Sheet label" is at AH42)
    
    Examples:
    - Input: "Sheet T1" 
      Output A2/B2: "00.10 : Etat comptable périodique : Section 1. Actif - Feuillet technique"
    
    - Input: "Sheet 01" (only sheet or paired with Sheet T1)
      Output A2/B2: "00.10 : Etat comptable périodique : Section 1. Actif"
    
    - Input: "Sheet 01" (with Sheet 02 also present)
      Output A2/B2: "02.15 : Transferts nets... - Toutes devises - Pays de l'UE (y compris la Belgique)"
    
    Args:
        source_wb: Source workbook containing the input sheets
        target_wb: Target workbook where SA.*.* sheets have been created
        sheet_list: List of sheet names from the input file
        sheet_name_mapping: Dictionary mapping input sheet names to output sheet names
        base_text: Base text from identification sheet (D2/E2 value)
    
    Returns:
        None (modifies target_wb in place)
    """
    import re
    
    if not base_text:
        logging.debug("   No base text available, skipping A2/B2 filling")
        return
    
    # Count sheets matching "Sheet 99" pattern (2 digits)
    sheet_99_pattern_count = sum(1 for s in sheet_list if re.match(r'Sheet \d{2}$', s))
    
    logging.debug(f"   Found {sheet_99_pattern_count} sheets with 'Sheet 99' pattern")
    
    for input_sheet_name in sheet_list:
        output_sheet_name = sheet_name_mapping.get(input_sheet_name, input_sheet_name)
        
        if output_sheet_name not in target_wb.sheetnames:
            logging.debug(f"   Skipping {output_sheet_name} - not found in target workbook")
            continue
            
        additional_text = ""
        
        # Rule 1: If sheet is "Sheet T1", add " - Feuillet technique"
        if input_sheet_name == "Sheet T1":
            additional_text = " - Feuillet technique"
            logging.debug(f"   Sheet '{input_sheet_name}' is T1 sheet, adding technical sheet suffix")
        
        # Rule 2: If sheet matches "Sheet 99" pattern
        elif re.match(r'Sheet \d{2}$', input_sheet_name):
            # Check if we should skip adding text
            if len(sheet_list) == 1:
                # Only 1 sheet total
                additional_text = ""
                logging.debug(f"   Only 1 sheet in input file, no suffix for '{input_sheet_name}'")
            elif len(sheet_list) == 2 and "Sheet T1" in sheet_list:
                # 2 sheets, one is T1
                additional_text = ""
                logging.debug(f"   2 sheets with one being T1, no suffix for '{input_sheet_name}'")
            elif sheet_99_pattern_count >= 2:
                # Multiple sheets with "Sheet 99" pattern - look for "Sheet label" value
                source_sheet = source_wb[input_sheet_name]
                sheet_label_value = None
                
                # Search for "Sheet label" cell
                for row in source_sheet.iter_rows():
                    for cell in row:
                        if cell.value and "Sheet label" in str(cell.value):
                            # Get the value from the SECOND cell to the right (column + 2)
                            label_cell = source_sheet.cell(row=cell.row, column=cell.column + 2)
                            sheet_label_value = label_cell.value
                            logging.debug(f"   Found 'Sheet label' at {cell.coordinate} in '{input_sheet_name}'")
                            logging.debug(f"   Sheet label value from {label_cell.coordinate}: '{sheet_label_value}'")
                            break
                    if sheet_label_value is not None:
                        break
                
                if sheet_label_value and sheet_label_value != "None":
                    additional_text = f" - {sheet_label_value}"
                    logging.debug(f"   Adding sheet label suffix: '{additional_text}'")
                else:
                    logging.debug(f"   Sheet label value is None or not found, no suffix added")
        
        # Fill A2 and B2 cells in output sheet
        final_text = base_text + additional_text
        target_sheet = target_wb[output_sheet_name]
        target_sheet["A2"] = final_text
        target_sheet["B2"] = final_text
        logging.debug(f"   Filled A2/B2 in '{output_sheet_name}': '{final_text}'")


def move_specific_sheets_to_end(workbook):
    sheets_to_move = ["cellSpecificities", "labelTranslation"]
    
    for sheet_name in sheets_to_move:
        if sheet_name in workbook.sheetnames:
            # Get the sheet object
            sheet = workbook[sheet_name]
            
            # Remove the sheet from its current position
            workbook.remove(sheet)
            
            # Add it back at the end
            workbook._add_sheet(sheet)
            
            logging.debug(f"   Moved sheet '{sheet_name}' to the end of the workbook")
        else:
            logging.debug(f"   Sheet '{sheet_name}' not found in workbook")
    
    logging.debug(f"   Final sheet order: {workbook.sheetnames}")


def process_tec_sheets(target_wb, form_code, sa_code, cross_internal_tables=None):
    """
    Process SA.*.TEC sheets for form codes 010, 020, 040, 050.
    Creates TEC sheet by copying SA.*.01 sheet and removing formula columns.
    Then fills it with Sheet T1 data from cross_internal_tables.
    
    Args:
        target_wb: openpyxl Workbook object
        form_code: Form code string (e.g., '010')
        sa_code: SA code string (e.g., 'SA.010')
        cross_internal_tables: Dict of cross internal tables data (optional)
    """
    logging.info(f"🔧  Processing TEC sheet for form code {form_code}")
    
    # Define sheet names
    tec_sheet_name = f"{sa_code}.TEC"
    base_sheet_name = f"{sa_code}.01"
    
    # Check if both sheets exist
    if tec_sheet_name not in target_wb.sheetnames:
        logging.warning(f"   TEC sheet '{tec_sheet_name}' not found, skipping TEC processing")
        return
    
    if base_sheet_name not in target_wb.sheetnames:
        logging.warning(f"   Base sheet '{base_sheet_name}' not found, cannot create TEC sheet")
        return
    
    logging.debug(f"   Found TEC sheet: {tec_sheet_name}")
    logging.debug(f"   Found base sheet: {base_sheet_name}")
    
    # Step 1: Store TEC label from A2
    tec_sheet = target_wb[tec_sheet_name]
    tec_label = tec_sheet['A2'].value
    logging.debug(f"   Stored TEC label from A2: '{tec_label}'")
    
    # Get the index of the TEC sheet for insertion location
    tec_index = target_wb.sheetnames.index(tec_sheet_name)
    logging.debug(f"   TEC sheet index: {tec_index}")
    
    # Step 2: Delete the SA.*.TEC sheet
    target_wb.remove(tec_sheet)
    logging.debug(f"   Deleted original TEC sheet")
    
    # Step 3 & 4: Copy SA.*.01 and insert at TEC location
    base_sheet = target_wb[base_sheet_name]
    new_tec_sheet = target_wb.copy_worksheet(base_sheet)
    
    # Move the new sheet to the correct position
    # After copying, it's at the end, so we need to move it
    current_sheets = target_wb.sheetnames
    current_index = current_sheets.index(new_tec_sheet.title)
    
    # If we need to move it to an earlier position
    if tec_index < current_index:
        # Move to the position where TEC was
        target_wb._sheets.insert(tec_index, target_wb._sheets.pop(current_index))
    
    # Step 5: Rename to SA.*.TEC
    new_tec_sheet.title = tec_sheet_name
    logging.debug(f"   Created new TEC sheet by copying base sheet")
    logging.debug(f"   Renamed to: {tec_sheet_name}")
    
    # Step 6: Restore A2 and B2 with TEC label
    new_tec_sheet['A2'].value = tec_label
    new_tec_sheet['B2'].value = tec_label
    logging.debug(f"   Restored A2 and B2 with TEC label: '{tec_label}'")
    
    # Step 7: Clear columns G and I from row 19 downward
    max_row = new_tec_sheet.max_row
    max_col = new_tec_sheet.max_column
    cleared_g_count = 0
    cleared_i_count = 0
    
    for row_idx in range(19, max_row + 1):
        # Clear column G (feedingRuleSource)
        if new_tec_sheet.cell(row=row_idx, column=7).value is not None:
            new_tec_sheet.cell(row=row_idx, column=7).value = None
            cleared_g_count += 1
        
        # Clear column I (computationRuleOrder)
        if new_tec_sheet.cell(row=row_idx, column=9).value is not None:
            new_tec_sheet.cell(row=row_idx, column=9).value = None
            cleared_i_count += 1
    
    logging.debug(f"   Cleared column G (feedingRuleSource) from row 19 to {max_row}: {cleared_g_count} cells cleared")
    logging.debug(f"   Cleared column I (computationRuleOrder) from row 19 to {max_row}: {cleared_i_count} cells cleared")
    
    # Step 7b: Clear rows 10 and 12 from column N (14) onwards (column data)
    cleared_row10_count = 0
    cleared_row12_count = 0
    
    for col_idx in range(14, max_col + 1):  # Column N (14) onwards
        # Clear row 10 (feedingRuleSource for columns)
        if new_tec_sheet.cell(row=10, column=col_idx).value is not None:
            new_tec_sheet.cell(row=10, column=col_idx).value = None
            cleared_row10_count += 1
        
        # Clear row 12 (computationRuleOrder for columns)
        if new_tec_sheet.cell(row=12, column=col_idx).value is not None:
            new_tec_sheet.cell(row=12, column=col_idx).value = None
            cleared_row12_count += 1
    
    logging.debug(f"   Cleared row 10 (column feedingRuleSource) from column N onwards: {cleared_row10_count} cells cleared")
    logging.debug(f"   Cleared row 12 (column computationRuleOrder) from column N onwards: {cleared_row12_count} cells cleared")
    
    # Step 8: Copy data validations from columns C, E, F and rows 6, 8, 9 of base sheet to TEC sheet
    from openpyxl.utils import get_column_letter
    from copy import copy
    
    # Part 1: Copy data validations for columns C, E, F (row 19 and below)
    columns_to_copy_validations = [3, 5, 6]  # C, E, F
    validations_copied = 0
    processed_dvs = set()  # Track which data validations we've already processed
    
    for dv in base_sheet.data_validations.dataValidation:
        # Check if this validation covers any of columns C, E, F in row 19+
        relevant_ranges = []
        for cr in dv.cells.ranges:
            # Check if this range intersects with columns C, E, F and row 19+
            for col_idx in columns_to_copy_validations:
                if cr.min_col <= col_idx <= cr.max_col and cr.max_row >= 19:
                    # This range covers one of our target columns in row 19+
                    col_letter = get_column_letter(col_idx)
                    start_row = max(cr.min_row, 19)
                    end_row = min(cr.max_row, max_row)
                    
                    # Add cells for this column/row range
                    for row_num in range(start_row, end_row + 1):
                        relevant_ranges.append(f"{col_letter}{row_num}")
        
        if relevant_ranges:
            # Create a new data validation with only the relevant ranges
            new_sqref_str = " ".join(relevant_ranges)
            new_dv = copy(dv)
            new_dv.sqref = new_sqref_str
            new_tec_sheet.add_data_validation(new_dv)
            validations_copied += 1
            logging.debug(f"   Copied data validation for columns C/E/F rows 19+")
    
    logging.debug(f"   Copied {validations_copied} data validations to columns C, E, F")
    
    # Part 2: Copy data validations for rows 6, 8, 9 (all columns)
    rows_to_copy_validations = [6, 8, 9]
    row_validations_copied = 0
    
    for dv in base_sheet.data_validations.dataValidation:
        # Check if this validation covers any of rows 6, 8, 9
        relevant_ranges = []
        for cr in dv.cells.ranges:
            # Check if this range intersects with rows 6, 8, 9
            for row_idx in rows_to_copy_validations:
                if cr.min_row <= row_idx <= cr.max_row:
                    # This range covers one of our target rows
                    # Add cells for this row, same column range as original
                    for col_num in range(cr.min_col, cr.max_col + 1):
                        col_letter = get_column_letter(col_num)
                        relevant_ranges.append(f"{col_letter}{row_idx}")
        
        if relevant_ranges:
            # Create a new data validation with only the relevant ranges
            new_sqref_str = " ".join(relevant_ranges)
            new_dv = copy(dv)
            new_dv.sqref = new_sqref_str
            new_tec_sheet.add_data_validation(new_dv)
            row_validations_copied += 1
            logging.debug(f"   Copied data validation for rows 6/8/9")
    
    logging.debug(f"   Copied {row_validations_copied} data validations to rows 6, 8, 9")
    
    # Step 9: Fill TEC sheet with Sheet T1 data from cross_internal_tables
    # This must be done AFTER copying from Sheet 01 to preserve Sheet T1 colors/formulas
    if cross_internal_tables and 'Sheet T1' in cross_internal_tables:
        logging.info("   Filling TEC sheet with Sheet T1 data (colors, formulas)")
        sheet_t1_cross_table = cross_internal_tables['Sheet T1']
        fill_output_table_from_cross_internal_tables(new_tec_sheet, sheet_t1_cross_table, tec_sheet_name, form_code)
        logging.debug(f"   Filled {len(sheet_t1_cross_table)} cells in TEC sheet from Sheet T1")
    else:
        logging.debug("   No Sheet T1 cross_internal_tables data available, skipping TEC sheet fill")
    
    logging.info(f"   ✅ TEC sheet processed successfully")


def autosave_with_excel(path, fix_merged_column_borders=None):
    """
    Open and save an Excel file using Excel COM for border compatibility.
    
    Args:
        path: Path to the Excel file
        fix_merged_column_borders: Optional dict with format:
            {
                'sheet_name': {'row': 15, 'last_column': 22}  # Apply medium border to right of last column
            }
    """
    import time
    
    # Retry logic for opening the file (in case of file locking issues)
    max_retries = 3
    retry_delay = 2  # seconds
    
    excel = None
    wb = None
    
    for attempt in range(max_retries):
        try:
            # Wait a bit before each attempt to ensure file is released
            if attempt > 0:
                logging.debug(f"   Retry attempt {attempt + 1}/{max_retries} to open Excel file")
                time.sleep(retry_delay)
            
            # Initialize Excel application
            excel = win32.gencache.EnsureDispatch('Excel.Application')
            excel.DisplayAlerts = False
            excel.Visible = False
            
            # Try to open the workbook
            wb = excel.Workbooks.Open(os.path.abspath(path))
            logging.debug(f"   Successfully opened Excel file: {path}")
            break  # Success, exit retry loop
            
        except Exception as e:
            logging.warning(f"   Attempt {attempt + 1}/{max_retries} failed to open Excel file: {e}")
            
            # Clean up on failure
            if wb:
                try:
                    wb.Close(SaveChanges=False)
                except:
                    pass
            if excel:
                try:
                    excel.Quit()
                except:
                    pass
            
            # If this was the last attempt, re-raise the error
            if attempt == max_retries - 1:
                raise
    
    try:
        # Fix borders for merged cells if specified
        if fix_merged_column_borders:
            from openpyxl.utils import get_column_letter
            for sheet_name, border_info in fix_merged_column_borders.items():
                try:
                    ws = wb.Worksheets(sheet_name)
                    row = border_info['row']
                    last_col = border_info['last_column']
                    col_letter = get_column_letter(last_col)
                    cell_address = f"{col_letter}{row}"
                    
                    # Set the right border using Excel constants
                    # xlEdgeRight = 10, xlContinuous = 1, xlMedium = -4138
                    ws.Range(cell_address).Borders(10).LineStyle = 1     # 10 = xlEdgeRight, 1 = xlContinuous
                    ws.Range(cell_address).Borders(10).Weight = -4138    # -4138 = xlMedium
                    logging.debug(f"   Excel COM: Set medium right border on {sheet_name}!{cell_address}")
                except Exception as e:
                    logging.warning(f"   Excel COM: Failed to set border on {sheet_name}: {e}")
        
        # Save the workbook
        wb.Save()
        logging.debug(f"   Saved Excel file successfully")
        
    finally:
        # Ensure cleanup happens even if there's an error
        if wb:
            try:
                wb.Close(SaveChanges=False)  # Already saved above
            except Exception as e:
                logging.debug(f"   Error closing workbook: {e}")
        
        if excel:
            try:
                excel.Quit()
            except Exception as e:
                logging.debug(f"   Error quitting Excel: {e}")



###############################################################################################################################################################
#                                                                                                                                                             #  
#       6th and last part: Main                                                                                                                               #  
#                                                                                                                                                             #  
#   Completes everything in a main() function!                                                                                                                #  
#                                                                                                                                                             #                                                                                                                                               #  
###############################################################################################################################################################

def fill_output_table_from_cross_internal_tables(table_sheet, cross_internal_tables, sheet_name, form_code):
    """
    Fill the output table cells using cross_internal_tables data.
    
    For each cell in cross_internal_tables:
    - Map to corresponding cell in output table starting at N19
    - If cell_formula is filled in, write it to the cell and apply wrap text
    - Apply cell_color as fill color
    - If cell_color is #FF99CC (pink) or #FFFFCC (yellow), format as Number with 2 decimals and thousand separator
    - After filling a complete row, if all cells are #808080 (grey), set abstractRow (column C) to "yes"
    
    Special case for formCode 020 and SA.020.TEC sheet:
    - Row 147 (cells N147-V147) should all be grey (#808080) and C147 should be "yes"
    
    Args:
        table_sheet: The output worksheet
        cross_internal_tables: Dictionary of cross internal tables data for this sheet
        sheet_name: Name of the sheet (for logging and special case handling)
        form_code: The form code (for special case handling)
    """
    from openpyxl.styles import PatternFill, Alignment
    from openpyxl.utils import get_column_letter
    
    if not cross_internal_tables:
        logging.debug(f"   No cross internal tables data for '{sheet_name}', skipping table fill")
        return
    
    logging.info(f"   Filling output table from cross internal tables: {len(cross_internal_tables)} cells")
    
    # Starting position for output table
    start_row = 19
    start_col = 14  # Column N
    
    # Group cells by row number for processing
    cells_by_row = {}
    for cell_ref, cell_data in cross_internal_tables.items():
        row_num = cell_data['row_num']
        if row_num not in cells_by_row:
            cells_by_row[row_num] = []
        cells_by_row[row_num].append(cell_data)
    
    # Sort rows by row number
    sorted_rows = sorted(cells_by_row.keys())
    
    # Track output row mapping: source_row_num -> output_row_num
    output_row_mapping = {}
    current_output_row = start_row
    
    # Special handling for formCode 020 and SA.020.TEC sheet
    # Row 228 (rowNumbering 286T) is excluded from Sheet T1 cross_internal_tables
    # This means row 147 in the output needs to be skipped and filled with grey manually
    skip_output_row_147 = (form_code == '020' and sheet_name == 'SA.020.TEC')
    
    # Fill cells row by row
    for source_row_num in sorted_rows:
        row_cells = cells_by_row[source_row_num]
        
        # For formCode 020 SA.020.TEC: Skip output row 147 when we reach source row 229
        # (because source row 228 is excluded from Sheet T1)
        if skip_output_row_147 and source_row_num >= 229 and current_output_row == 147:
            logging.debug(f"   Skipping output row 147 for special grey row handling (source row {source_row_num})")
            current_output_row = 148  # Jump to row 148
        
        # Sort cells by column number
        row_cells.sort(key=lambda c: c['col_num'])
        
        # Track if all cells in this row are grey
        all_grey = True
        
        # Fill each cell in the row
        for i, cell_data in enumerate(row_cells):
            output_col = start_col + i
            output_cell = table_sheet.cell(row=current_output_row, column=output_col)
            
            # Apply cell formula if present, or clear existing formula/value if no formula in source
            cell_formula = cell_data.get('cell_formula', '')
            if cell_formula:
                output_cell.value = cell_formula
                # Apply wrap text
                output_cell.alignment = Alignment(wrap_text=True)
                logging.debug(f"   Cell {get_column_letter(output_col)}{current_output_row}: Applied formula '{cell_formula}' with wrap text")
            else:
                # Clear any existing formula or value (TEC sheet may have been copied from Sheet 01)
                output_cell.value = None
            
            # Apply cell color
            cell_color = cell_data.get('cell_color', '#FFFFCC')
            fill = PatternFill(start_color=cell_color.lstrip('#'), end_color=cell_color.lstrip('#'), fill_type='solid')
            output_cell.fill = fill
            
            # Check if cell is not grey
            if cell_color.upper() != '#808080':
                all_grey = False
            
            # Apply number formatting based on cell color
            if cell_color.upper() in ['#FF99CC', '#FFFFCC']:
                # Pink or yellow cells: number format with 2 decimals
                output_cell.number_format = '#,##0.00'
                logging.debug(f"   Cell {get_column_letter(output_col)}{current_output_row}: Applied number format with 2 decimals and thousand separator")
            elif cell_color.upper() == '#808080':
                # Grey cells: General format (clear any inherited number formatting)
                output_cell.number_format = 'General'
        
        # After filling all cells in the row, check if all were grey
        if all_grey:
            # Set abstractRow (column C) to "yes"
            abstract_cell = table_sheet.cell(row=current_output_row, column=3)
            abstract_cell.value = "yes"
            logging.debug(f"   Row {current_output_row}: All cells are grey, set abstractRow to 'yes'")
        
        # Store mapping for this source row
        output_row_mapping[source_row_num] = current_output_row
        current_output_row += 1
    
    # Special case for formCode 020 and SA.020.TEC sheet
    # Row 147 needs to be filled with grey even though it's not in cross_internal_tables
    if form_code == '020' and sheet_name == 'SA.020.TEC':
        special_row = 147
        # Determine the number of columns to fill (from column N onwards)
        # Based on the cross_internal_tables, we can determine the column count
        if cross_internal_tables:
            # Get the first cell to determine column count
            first_cell = next(iter(cross_internal_tables.values()))
            # Count unique column numbers in the first row
            first_row_cells = [c for c in cross_internal_tables.values() if c['row_num'] == first_cell['row_num']]
            column_count = len(first_row_cells)
            
            # Fill cells N147 through N147+column_count with grey
            grey_fill = PatternFill(start_color='808080', end_color='808080', fill_type='solid')
            for i in range(column_count):
                output_col = start_col + i
                output_cell = table_sheet.cell(row=special_row, column=output_col)
                output_cell.fill = grey_fill
                # Clear any inherited number format (set to General)
                output_cell.number_format = 'General'
                # Clear any formula or value
                output_cell.value = None
            
            # Set C147 to "yes"
            abstract_cell = table_sheet.cell(row=special_row, column=3)
            abstract_cell.value = "yes"
            
            logging.info(f"   Special case: Filled row 147 with grey for formCode 020, SA.020.TEC sheet")
    
    logging.info(f"   Filled output table: {len(sorted_rows)} rows processed")


def process_sheet_data(source_wb, target_wb, input_sheet_name, output_sheet_name, internal_tables, column_internal_tables, cross_internal_tables, form_code, form_group_code, version, validity_starting_date, linked_domain, datatype, roundingUnit, form_code_length, target_path=None, dedicated_logger=None):
    
    # Helper function to log to the appropriate logger
    def log_info(msg):
        if dedicated_logger:
            dedicated_logger.info(msg)
        else:
            logging.info(msg)
    
    def log_debug(msg):
        if dedicated_logger:
            dedicated_logger.debug(msg)
        else:
            logging.debug(msg)
    
    # Temporarily redirect logging calls to our dedicated logger if available
    original_debug = logging.debug
    original_info = logging.info
    
    if dedicated_logger:
        # Replace logging functions temporarily within this function scope
        logging.debug = dedicated_logger.debug
        logging.info = dedicated_logger.info
    
    try:
        # Use dedicated logger if available, otherwise fallback to standard logging
        log_info("")
        log_info(f"== 📄 PROCESSING SHEET DATA : {input_sheet_name} -> {output_sheet_name} ==")
        
        # Log Excel step if logger is available
        if EXCEL_LOGGER_AVAILABLE and target_path:
            excel_logger.log_excel_step(target_path, f"Processing sheet: {input_sheet_name} -> {output_sheet_name}")
        
        source_sheet = source_wb[input_sheet_name]
        table_sheet = target_wb[output_sheet_name]
        
        # Get internal table data for this sheet
        internal_table_data = internal_tables.get(input_sheet_name, [])
        
        log_info("✅  Adjusting output table rows")
        if EXCEL_LOGGER_AVAILABLE and target_path:
            excel_logger.log_excel_step(target_path, "Adjusting output table rows", f"Internal table rows: {len(internal_table_data)}")
        # Also log to console for visibility
        logging.info(f"   Adjusting table rows for '{input_sheet_name}': {len(internal_table_data)} rows in internal table")
        adjust_output_table_rows(table_sheet, len(internal_table_data))
        
        # Get column internal table data for this sheet
        column_table_data = column_internal_tables.get(input_sheet_name)
        
        # Adjust output table columns based on column internal table
        if column_table_data:
            column_count = column_table_data['column_count']
            print(f"   Adjusting output table columns: {column_count} columns")
            if EXCEL_LOGGER_AVAILABLE and target_path:
                excel_logger.log_excel_step(target_path, "Adjusting output table columns", f"Internal table columns: {column_count}")
            adjust_output_table_columns(table_sheet, column_count)
        else:
            logging.debug(f"   No column internal table found for '{input_sheet_name}', skipping column adjustment")
        
        #Copy data from internal table to output sheet
        log_info("✅  Copying internal table data to output sheet")
        if EXCEL_LOGGER_AVAILABLE and target_path:
            excel_logger.log_excel_step(target_path, "Copying internal table to output", f"Rows to copy: {len(internal_table_data)}")
        # Also log to console for visibility
        print(f"   Copying {len(internal_table_data)} rows from internal table to output sheet '{output_sheet_name}'")
        copy_internal_table_to_output(table_sheet, internal_table_data)
        
        # Calculate column computationRuleOrder AFTER copying row data (so we can read max from column I)
        if column_table_data:
            print("   Calculating column computationRuleOrder")
            column_internal_tables[input_sheet_name] = calculate_column_computation_rule_order_for_sheet(
                column_table_data, table_sheet, input_sheet_name
            )
            
            # Copy column data to output sheet (rows 10, 12, 18)
            print("   Copying column data to output sheet")
            copy_column_data_to_output(table_sheet, column_internal_tables[input_sheet_name], input_sheet_name)
            
            # Fill columnLabel rows (15, 16, 17) based on form_code_length
            print("   Filling columnLabel rows")
            border_fix_info = fill_column_labels(
                table_sheet, 
                column_table_data['column_count'], 
                form_code_length, 
                output_sheet_name,
                source_sheet=source_sheet,
                column_start_col=column_table_data.get('start_col'),
                column_table_data=column_table_data
            )
            logging.debug(f"   border_fix_info from fill_column_labels: {border_fix_info}")
        else:
            border_fix_info = None
        
        # Fill output table cells from cross_internal_tables
        # Get cross internal tables data for this sheet
        sheet_cross_table = cross_internal_tables.get(input_sheet_name, {})
        
        # Skip filling for Sheet T1 when formCode is 010, 020, 040, 050
        # because TEC sheet will be recreated later by copying Sheet 01
        # and we need to fill it AFTER recreation to preserve Sheet T1 colors/formulas
        skip_filling = (input_sheet_name == "Sheet T1" and form_code in ['010', '020', '040', '050'])
        
        if sheet_cross_table and not skip_filling:
            print("   Filling output table from cross internal tables")
            fill_output_table_from_cross_internal_tables(table_sheet, sheet_cross_table, output_sheet_name, form_code)
        elif skip_filling:
            logging.debug(f"   Skipping output table fill for '{input_sheet_name}' (will be filled after TEC sheet recreation)")
        
        log_info("✅  Sheet processing complete")
        
        # Return border fix information for Excel COM finalization
        logging.debug(f"   Returning from process_sheet_data: border_fix_info={border_fix_info}")
        return border_fix_info
            
    finally:
        # Always restore original logging functions
        if dedicated_logger:
            logging.debug = original_debug
            logging.info = original_info
    

def main():
    # Capture initial logs to replay them in each file's log
    initial_logs = []
    
    # Get user input (input file path + debug mode)
    msg = "🔧  Getting user input"
    logging.info(msg)
    initial_logs.append(("info", msg))
    
    source_path, form_code, form_code_length, conversion_6_4_digits, form_group_code, version, validity_starting_date, linked_domain, datatype, roundingUnit = get_user_input()
    output_dir = create_output_directory(source_path)
    
    # Get input filename for output naming
    input_filename = os.path.basename(source_path)
    logging.info(f"Input file: {input_filename}")
    logging.info(f"FormCode: {form_code} (length: {form_code_length})")
    logging.info(f"Conversion 5/6-digit to 4-digit rowIDs: {conversion_6_4_digits}")

    # Load the input workbook (read-only for data)
    msg = f"🔧  Loading input workbook: {source_path}"
    logging.info(msg)
    initial_logs.append(("info", msg))
    source_wb = openpyxl.load_workbook(source_path, data_only=True)
    logging.debug(f"   Loaded input workbook: {source_path}")

    # Get the relevant sheets from the input file
    msg = "🔧  Getting relevant sheets from input file"
    logging.info(msg)
    initial_logs.append(("info", msg))
    relevant_sheets = get_relevant_sheets(source_wb)
    
    # Log details about retained/skipped sheets
    for sheet in source_wb.sheetnames:
        if sheet in relevant_sheets:
            msg = f"   Retained sheet: {sheet}"
            logging.debug(msg)
            initial_logs.append(("debug", msg))
        else:
            msg = f"   Skipped sheet: {sheet}"
            logging.debug(msg)
            initial_logs.append(("debug", msg))
    
    # Extract internal tables from sheets (between HNRB and HNRE markers)
    msg = "🔧  Extracting internal tables from sheets"
    logging.info(msg)
    initial_logs.append(("info", msg))
    internal_tables = extract_internal_tables(source_wb, relevant_sheets)
    logging.info(f"   Extracted internal tables from {len(internal_tables)} sheets")
    
    # Extract column-based internal tables (for filling column part of SA.*.* sheets)
    msg = "🔧  Extracting column-based internal tables"
    logging.info(msg)
    initial_logs.append(("info", msg))
    column_internal_tables = extract_column_internal_tables(source_wb, relevant_sheets)
    logging.info(f"   Extracted column internal tables from {sum(1 for v in column_internal_tables.values() if v)} sheets")
    
    # Clean up column internal tables (ensure 3-digit column numbering)
    msg = "🔧  Cleaning up column internal tables"
    logging.info(msg)
    initial_logs.append(("info", msg))
    column_internal_tables = cleanup_column_internal_tables(column_internal_tables)
    logging.info(f"   Cleaned column internal tables from {sum(1 for v in column_internal_tables.values() if v)} sheets")
    
    # Process column feeding rules (transform formulas)
    msg = "🔧  Processing column feeding rules"
    logging.info(msg)
    initial_logs.append(("info", msg))
    column_internal_tables = process_column_feeding_rules(column_internal_tables)
    logging.info(f"   Processed column feeding rules from {sum(1 for v in column_internal_tables.values() if v)} sheets")
    
    # Clean up internal tables according to specific rules
    msg = "🔧  Cleaning up internal tables"
    logging.info(msg)
    initial_logs.append(("info", msg))
    internal_tables = cleanup_internal_tables(internal_tables, form_code_length, form_code)
    logging.info(f"   Cleaned internal tables from {len(internal_tables)} sheets")
    
    # Create cross internal tables (2D cell mapping)
    msg = "🔧  Creating cross internal tables"
    logging.info(msg)
    initial_logs.append(("info", msg))
    cross_internal_tables = create_cross_internal_tables(internal_tables, column_internal_tables, form_code, form_code_length)
    logging.info(f"   Created cross internal tables for {len(cross_internal_tables)} sheets")
    
    # Process cell colors and formulas for all cells in cross internal tables
    msg = "🔧  Processing cell colors and formulas"
    logging.info(msg)
    initial_logs.append(("info", msg))
    cross_internal_tables = process_cell_colors_and_formulas(cross_internal_tables, source_wb, form_code)
    logging.info(f"   Processed cell colors and formulas for {len(cross_internal_tables)} sheets")
    
    # Calculate level column for internal tables
    internal_tables = calculate_level_column(internal_tables, form_code_length)
    
    # Remove leading spaces from rowLabelColumnNumbering
    internal_tables = remove_leading_spaces_from_labels(internal_tables)
    
    # Convert 6/5-digit row numbers to 4-digit (only if conversion_6_4_digits = 'Y')
    if conversion_6_4_digits == 'Y':
        # Path to mapping file (assumed to be in the same directory as the script)
        script_dir = os.path.dirname(os.path.abspath(__file__))
        mapping_file_path = os.path.join(script_dir, 'Mapping_Rub_6_4_digits.xlsx')
        
        if os.path.exists(mapping_file_path):
            internal_tables = convert_6_4_digit_row_numbers(internal_tables, mapping_file_path)
        else:
            logging.warning(f"   Mapping file not found: {mapping_file_path}")
            logging.warning(f"   Skipping 6/5-digit to 4-digit conversion")
    
    # Transform feeding rule patterns: remove spaces and replace L{digits} with {r{digits}}
    internal_tables = transform_feeding_rule_patterns(internal_tables)
    
    # Calculate computationRuleOrder based on feeding rule dependency depth
    internal_tables = calculate_computation_rule_order(internal_tables)
    
    # Process all sheets as a single output file (1:1 mapping)
    sheet_list = relevant_sheets
    generated_files = []  # Track generated files for summary
    
    logging.info("")
    logging.info("=====================================")
    logging.info("== 📢 PROCESSING INPUT FILE ==")
    logging.info(f" Input: '{input_filename}' with {len(sheet_list)} sheets")
    logging.info("=====================================")

    try:
        # Copy the template file once for this input file
        target_path, dedicated_logger, sa_code = copy_template_file(input_filename, output_dir, form_code, form_group_code)
        if not target_path:
            logging.error(f"Could not create target file for '{input_filename}'. Exiting.")
            return

        # Replay initial logs in this file's log FIRST
        if dedicated_logger and initial_logs:
            dedicated_logger.info("=" * 70)
            dedicated_logger.info("INITIAL SETUP LOGS")
            dedicated_logger.info("=" * 70)
            for level, msg in initial_logs:
                if level == "info":
                    dedicated_logger.info(msg)
                elif level == "debug":
                    dedicated_logger.debug(msg)
                elif level == "error":
                    dedicated_logger.error(msg)
            dedicated_logger.info("=" * 70)
            dedicated_logger.info("")

        # Initialize Excel logging for this file
        if EXCEL_LOGGER_AVAILABLE:
            excel_logger.excel_file_logger.log_excel_creation_start(
                target_path,
                source_file=source_path,
                root_name=input_filename,
                sheets=sheet_list,
                form_code=form_code,
                form_group_code=form_group_code
            )

        # Load the newly created target workbook
        target_wb = openpyxl.load_workbook(target_path, keep_vba=True)
        logging.debug(f"   Loaded target workbook: {target_path}")

        # Create the necessary sheets in the target workbook
        if EXCEL_LOGGER_AVAILABLE:
            excel_logger.log_excel_step(target_path, "Creating target sheets", f"Number of sheets: {len(sheet_list)}")
        sheet_name_mapping = create_empty_target_sheets(target_wb, sheet_list, sa_code)

        # Fill identification sheet and get base text for A2/B2 cells
        # This must be done BEFORE process_sheet_data() so that A2 is populated
        # when copy_internal_table_to_output() tries to read it for default rowLabelColumnNumbering
        logging.info("")
        logging.info("🔧  Filling identification sheet")
        if EXCEL_LOGGER_AVAILABLE:
            excel_logger.log_excel_step(target_path, "Filling identification sheet")
        base_text = fill_identification_and_labelTranslation(source_wb, target_wb, form_group_code, version, validity_starting_date, linked_domain, datatype, roundingUnit)
        
        # Fill A2/B2 cells in SA.*.* sheets with conditional text
        logging.info("🔧  Filling A2/B2 cells in SA.*.* sheets")
        if EXCEL_LOGGER_AVAILABLE:
            excel_logger.log_excel_step(target_path, "Filling A2/B2 cells in SA.*.* sheets")
        fill_sheet_labels(source_wb, target_wb, sheet_list, sheet_name_mapping, base_text)

        # Now, process each sheet (adjust headers, paste data, etc.)
        # Use the mapping to get the correct output sheet name
        logging.info("")
        logging.info("🔧  Processing sheet data")
        border_fixes = {}  # Collect border fix info from all sheets
        for input_sheet_name in sheet_list:
            logging.info(f"   Processing sheet: '{input_sheet_name}'")
            output_sheet_name = sheet_name_mapping.get(input_sheet_name, input_sheet_name)
            logging.info(f"   Output sheet name: '{output_sheet_name}'")
            border_info = process_sheet_data(source_wb, target_wb, input_sheet_name, output_sheet_name, internal_tables, column_internal_tables, cross_internal_tables, form_code, form_group_code, version, validity_starting_date, linked_domain, datatype, roundingUnit, form_code_length, target_path, dedicated_logger)
            logging.debug(f"   border_info returned from process_sheet_data: {border_info}")
            if border_info:
                border_fixes[output_sheet_name] = border_info
                logging.debug(f"   Added border fix for sheet '{output_sheet_name}': {border_info}")
            logging.info(f"   Completed processing sheet: '{input_sheet_name}'")

        # Process TEC sheets for applicable form codes (010, 020, 040, 050)
        if form_code in ['010', '020', '040', '050']:
            logging.info("")
            if EXCEL_LOGGER_AVAILABLE:
                excel_logger.log_excel_step(target_path, f"Processing TEC sheet for form code {form_code}")
            process_tec_sheets(target_wb, form_code, sa_code, cross_internal_tables)

        logging.info("")
        logging.info("🔧  Moving specific sheets to end")
        if EXCEL_LOGGER_AVAILABLE:
            excel_logger.log_excel_step(target_path, "Moving specific sheets to end")
        move_specific_sheets_to_end(target_wb)

        # Clear the COM cache to avoid issues with win32com
        logging.info("🔧  Cleaning up COM cache")
        if EXCEL_LOGGER_AVAILABLE:
            excel_logger.log_excel_step(target_path, "Cleaning up COM cache")
        gen_py_cache = os.path.join(os.environ.get('LOCALAPPDATA'), 'Temp', 'gen_py')
        if os.path.exists(gen_py_cache):
            shutil.rmtree(gen_py_cache)
        # Reset cache
        win32com.client.gencache.EnsureDispatch('Excel.Application')

        # Save and close the target workbook
        logging.info(f"🔧  Saving and closing target workbook: {target_path}")
        if EXCEL_LOGGER_AVAILABLE:
            excel_logger.log_excel_step(target_path, "Saving target workbook")
        target_wb.save(target_path)
        target_wb.close()
        
        # Give the system time to flush file buffers
        import time
        time.sleep(1)

        # Finalize with Excel for border compatibility
        logging.info("🔧  Finalizing with Excel Windows saving for border compatibility")
        if border_fixes:
            logging.debug(f"   Border fixes to apply: {border_fixes}")
        if EXCEL_LOGGER_AVAILABLE:
            excel_logger.log_excel_step(target_path, "Finalizing with Excel Windows for border compatibility")
        autosave_with_excel(target_path, fix_merged_column_borders=border_fixes if border_fixes else None)
        logging.debug(f"   Finalized file with Excel: {target_path}")

        # Finalize Excel logging
        if EXCEL_LOGGER_AVAILABLE:
            excel_logger.complete_excel_logging(
                target_path, 
                success=True,
                total_sheets=len(sheet_list),
                processed_sheets=sheet_list,
                form_code=form_code
            )

        generated_files.append(target_path)
        logging.info(f"✅ Successfully generated: {target_path}")

    except Exception as e:
        logging.error(f"❌ Error processing input file '{input_filename}': {str(e)}")
        if EXCEL_LOGGER_AVAILABLE and 'target_path' in locals():
            excel_logger.log_excel_error(target_path, f"Error processing input file '{input_filename}': {str(e)}", e)
            excel_logger.complete_excel_logging(target_path, success=False, error=str(e))

    # Close the source workbook
    logging.info("🔧  Closing source workbook")
    source_wb.close()    # Summary of generated files
    logging.info("")
    logging.info("=" * 70)
    logging.info("GENERATION SUMMARY")
    logging.info(f"Number of files generated: {len(generated_files)}")
    logging.info("Generated files:")
    for i, file_path in enumerate(generated_files, 1):
        logging.info(f"  {i}. {os.path.basename(file_path)}")
        logging.info(f"     Full path: {file_path}")
        if EXCEL_LOGGER_AVAILABLE:
            log_file = excel_logger.excel_file_logger.log_files.get(file_path)
            if log_file:
                logging.info(f"     Log file: {log_file}")
    
    if EXCEL_LOGGER_AVAILABLE and generated_files:
        logs_dir = os.path.join(output_dir, "logs")
        logging.info(f"Logs directory: {logs_dir}")
    
    logging.info("=" * 70)
    logging.info("Processing complete.")


if __name__ == "__main__":
    main()
