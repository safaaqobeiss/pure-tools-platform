"""
### 🔧 **Prompt 

**Create one file per sheet found (except the "Entry points" sheet) using an Excel template**

1. Create Python code to generate target Excel files based on each sheet in the source workbook.
   * Ask the user the path of the DPM source workbook (with file dialog)
   * Ask the user the formCode as a mandatory field and as optional fields : formGroupCode, version, validityStartingDate, linkedDomain, datatype (by default : euebacorep.positiveOrNullAmount)
   * Inform the user that the folder "generated_output" for generated files will be created if not exist at the grand-parent level of the source workbook (../)
   * The template file to be used is here "../template/pure-form-regulatoryOutputCode-formCode-(shortName).xlsm"

2. For each root sheet in the source workbook :
   * Get the sheet name.
   * Skip the sheet named "TOC" and "K_00.03".
   * Sheet  name containing only one "." will lead to only one Excel file created (as per the description below) 
        ==> rootName will be the sheet name
   * Sheet  name containing two "." and sharing same root name (eg. K_42.00 for K_42_00.a and K_42.00.b) will be grouped together in a single Excel file having multiple sheets. 
        ==> rootName will be here the real root name

   * Copy the template file.
   * Rename the new file using the naming pattern:
        `"pure-form-esgadhoc-[rootName].xlsm"`.
   * In the new file, copy (with data validation) and rename the `"code"` sheet as many times as there are different sheet names sharing same rootName. 
        ==> each initial `"code"` sheet will be called table
               
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
   * Get from source worksheet in A1 the text after the first "- " 
   * Copy this text to D2 cell and E2 cell of "identification" sheet from the generated workbook
   * In the G2 cell, write "direct"

6. Fill the labelTranslation sheet in the generated workbook :
   * Get from source worksheet in A1 the text after the first "- " 
   * Copy this text to A2 cell and B2 cell of "labelTranslation" sheet from the generated workbook
   * In the B1 cell, write "FR"

7. The final code must:
   * Include robust error handling
   * Provide detailed debugging information : there should be a debug mode that can be activated and asked to the user to provide maximum of information at each step
   * Properly handle file paths with spaces
   * Not alter any original files


   Could you split your answer into several logical parts (the size of your answer is limited to maximum 80 lines of Python code) ? I will ask you for the following parts in a next prompt



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
from tkinter import Tk, filedialog, messagebox, simpledialog, Label, Entry, Button
import win32com.client as win32
import win32com.client.gencache

# Import the feeding rule source processor and excel logger
sys.path.append(os.path.join(os.path.dirname(__file__), '..', 'common'))
try:
    import feedingRuleSource_in_forms as feeding_rules
    FEEDING_RULES_AVAILABLE = True
    VALIDATION_FILE_PATH = None  # Will be set once if feeding rules are used
except ImportError:
    FEEDING_RULES_AVAILABLE = False
    VALIDATION_FILE_PATH = None

# Import the Excel file logger
try:
    import excel_logger
    EXCEL_LOGGER_AVAILABLE = True
except ImportError:
    EXCEL_LOGGER_AVAILABLE = False
except ImportError:
    logging.warning("Feeding rules processor not available")
    FEEDING_RULES_AVAILABLE = False
    VALIDATION_FILE_PATH = None

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

    # 🟩 Select source file
    source_path = filedialog.askopenfilename(
        title="Select DPM source workbook from the regulator",
        filetypes=[("Excel files", "*.xlsx *.xlsm *.xls")]
    )
    if not source_path:
        messagebox.showerror("Error", "No file selected. Exiting.")
        exit(1)

    global DEBUG_MODE
    DEBUG_MODE = messagebox.askyesno("Debug Mode", "Enable debug mode?")
    configure_logging()
    

    # 🟩 Create a pop-up to collect all inputs together
    def ask_multiple_fields():
        # Create a new dialog window
        input_window = Tk()
        input_window.title("Form Details")     

        entries = {}

        labels = [
            ("formCode (optional)", False),
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
            form_code_val = entries["formCode (optional)"].get().strip() or ""
            # Collect all fields
            result = {
                "formCode": form_code_val,
                "formGroupCode": entries["formGroupCode (optional)"].get().strip() or "",
                "version": entries["version (optional)"].get().strip() or "",
                "validityStartingDate": entries["validityStartingDate (optional)"].get().strip() or "",
                "linkedDomain": entries["linkedDomain (optional)"].get().strip() or "",
                "datatype": entries["datatype (optional)"].get().strip() or "",
                "roundingUnit": entries["roundingUnit (optional)"].get().strip() or ""
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

    # 🟩 Return everything
    return source_path, user_inputs["formCode"], user_inputs["formGroupCode"], user_inputs["version"], user_inputs["validityStartingDate"], user_inputs["linkedDomain"], user_inputs["datatype"], user_inputs["roundingUnit"]



def create_output_directory(source_path):
    grandparent_dir = os.path.abspath(os.path.join(source_path, "../../"))
    output_dir = os.path.join(grandparent_dir, "generated_output")
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
        logging.info(f"Created output directory: {output_dir}")
    else:
        logging.debug(f"   Output directory already exists: {output_dir}")
    return output_dir


def copy_template_file(root_name, output_dir, form_code):
    # Get the directory where this script is located
    script_dir = os.path.dirname(os.path.abspath(__file__))
    template_path = os.path.join(script_dir, "..", "template", "pure-form-regulatoryOutputCode-formCode-(shortName).xlsm")
    template_path = os.path.abspath(template_path)  # Resolve any .. in the path
    
    target_filename = f"pure-form-{form_code}-{root_name}-({root_name.replace('_', ' ')}).xlsm"
    target_path = os.path.join(output_dir, target_filename)
    shutil.copy2(template_path, target_path)

    logging.debug(f"   Copied template from: {template_path}")
    logging.debug(f"   Copied template to: {target_path}")
    
    # Initialize Excel file logging if available
    dedicated_logger = None
    if EXCEL_LOGGER_AVAILABLE:
        try:
            dedicated_logger = excel_logger.setup_excel_file_logging(target_path, "forms_from_dpm")
            excel_logger.log_excel_step(target_path, "Template copied", 
                                      f"Template: {template_path} -> Target: {target_path}")
        except Exception as e:
            logging.warning(f"Error initializing Excel logging: {e}")
    
    return target_path, dedicated_logger


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
        if sheet not in ["TOC", "K_00.03", "Entry points"]:
            filtered_sheets.append(sheet)
            logging.debug(f"   Retained sheet: {sheet}")
        else :
            logging.debug(f"   Skipped sheet: {sheet}")
    return filtered_sheets

def group_sheets_by_root_name(sheets):
    grouped = {}
    for sheet in sheets:
        root_name = extract_root_name(sheet)
        if root_name not in grouped:
            grouped[root_name] = []
        grouped[root_name].append(sheet)
        logging.debug(f"   Grouping: {sheet} -> {root_name}")
    return grouped


def create_empty_target_sheets(target_wb, sheet_list):
    logging.debug("   === PROCESS WORKBOOKS ===")

    # Reference to the template sheet ('code')
    template_sheet = target_wb["code"]

    # Create copies for each target sheet
    for sheet_name in sheet_list:
        new_sheet = target_wb.copy_worksheet(template_sheet)
        new_sheet.title = sheet_name

        # 🟢 Reapply data validations manually (for openpyxl < 3.1)
        for dv in template_sheet.data_validations.dataValidation:
            # Clone data validation and reassign its range to the new sheet
            new_dv = copy(dv)  # Python's copy.copy
            new_sheet.add_data_validation(new_dv)

    # Remove the original 'code' sheet after duplication
    target_wb.remove(template_sheet)

    logging.debug(f"   Target workbook sheets after processing: {target_wb.sheetnames}")


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


def fill_identification_and_labelTranslation(source_sheet, target_wb,
                                             form_group_code, version, validity_starting_date,
                                             linked_domain, datatype, roundingUnit):
    # 🟢 Get the A1 value from the source sheet
    a1_value = source_sheet["A1"].value
   
    if not a1_value:
        logging.debug("   No value in A1.")
        return

    # 🟢 Extract text after the first '- '
    text_match = re.search(r"- (.*)", str(a1_value))
    if text_match:
        extracted_text = text_match.group(1).strip()
        logging.debug(f"   Extracted text: '{extracted_text}'")

        # 🟢 Fill identification sheet in the target workbook
        ident_sheet = target_wb["identification"]
        ident_sheet["D2"] = extracted_text
        ident_sheet["E2"] = extracted_text
        ident_sheet["G2"] = "direct"

        # Only write optional fields if user provided a value
        # (preserve template defaults when fields are empty)
        if form_group_code:
            ident_sheet["A2"] = form_group_code
        if version:
            ident_sheet["B2"] = version
        if validity_starting_date:
            ident_sheet["C2"] = validity_starting_date
        if linked_domain:
            ident_sheet["F2"] = linked_domain
        if datatype:
            ident_sheet["J2"] = datatype
        if roundingUnit:
            ident_sheet["L2"] = roundingUnit

        logging.debug("   identification sheet updated.")

        # 🟢 Fill labelTranslation sheet in the target workbook
        label_sheet = target_wb["labelTranslation"]
        label_sheet["A2"] = extracted_text
        label_sheet["B2"] = extracted_text
        label_sheet["B1"] = "FR"

        logging.debug("   labelTranslation sheet updated.")

    else:
        logging.debug("   No text found after '- ' in A1.")



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


def autosave_with_excel(path):
    try:
        excel = win32.gencache.EnsureDispatch('Excel.Application')
    except Exception:
        # Fallback to late binding if makepy/gencache fails
        excel = win32.Dispatch('Excel.Application')
    excel.DisplayAlerts = False
    excel.Visible = False
    wb = excel.Workbooks.Open(path)
    wb.Save()
    wb.Close()
    excel.Quit()



###############################################################################################################################################################
#                                                                                                                                                             #  
#       6th and last part: Main                                                                                                                               #  
#                                                                                                                                                             #  
#   Completes everything in a main() function!                                                                                                                #  
#                                                                                                                                                             #                                                                                                                                               #  
###############################################################################################################################################################

def process_sheet_data(source_wb, target_wb, sheet_name, form_group_code, version, validity_starting_date, linked_domain, datatype, roundingUnit, target_path=None, dedicated_logger=None):
    
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
        log_info(f"== 📄 PROCESSING SHEET DATA : {sheet_name} ==")
        
        # Log Excel step if logger is available
        if EXCEL_LOGGER_AVAILABLE and target_path:
            excel_logger.log_excel_step(target_path, f"Processing sheet: {sheet_name}")
        
        source_sheet = source_wb[sheet_name]
        table_sheet = target_wb[sheet_name]
        
        log_info("✅  Getting column headers")
        if EXCEL_LOGGER_AVAILABLE and target_path:
            excel_logger.log_excel_step(target_path, "Extracting column headers")
        header_cells, num_rows = get_column_headers(source_sheet)
        
        log_info("✅  Adjusting column labels")
        if EXCEL_LOGGER_AVAILABLE and target_path:
            excel_logger.log_excel_step(target_path, "Adjusting column labels", f"Number of header rows: {num_rows}")
        adjust_column_labels(table_sheet, num_rows)

        log_info("✅  Pasting column headers")
        if EXCEL_LOGGER_AVAILABLE and target_path:
            excel_logger.log_excel_step(target_path, "Pasting column headers")
        paste_column_headers(table_sheet, header_cells)

        log_info("✅  Pasting row headers")
        if EXCEL_LOGGER_AVAILABLE and target_path:
            excel_logger.log_excel_step(target_path, "Pasting row headers")
        paste_row_headers(table_sheet, source_sheet)
        
        log_info("✅  Finalizing table rectangle")
        if EXCEL_LOGGER_AVAILABLE and target_path:
            excel_logger.log_excel_step(target_path, "Finalizing table rectangle")
        finalize_table_rectangle(table_sheet)
        
        log_info("✅  Applying cell types and formatting")
        if EXCEL_LOGGER_AVAILABLE and target_path:
            excel_logger.log_excel_step(target_path, "Applying cell types and formatting")
        apply_cell_type(source_sheet, table_sheet)
        
        log_info("✅  Handling open rows")
        if EXCEL_LOGGER_AVAILABLE and target_path:
            excel_logger.log_excel_step(target_path, "Handling open rows")
        handle_open_rows(source_sheet, table_sheet)
        
        log_info("✅  Marking abstract rows")
        if EXCEL_LOGGER_AVAILABLE and target_path:
            excel_logger.log_excel_step(target_path, "Marking abstract rows")
        mark_abstract_rows(table_sheet)
        
        log_info("✅  Deleting empty rows")
        if EXCEL_LOGGER_AVAILABLE and target_path:
            excel_logger.log_excel_step(target_path, "Deleting empty rows")
        delete_empty_rows(table_sheet)
        
        log_info("✅  Removing extra right columns")
        if EXCEL_LOGGER_AVAILABLE and target_path:
            excel_logger.log_excel_step(target_path, "Removing extra right columns")
        remove_extra_right_columns(table_sheet)
        
        log_info("✅  Merging header cells")
        if EXCEL_LOGGER_AVAILABLE and target_path:
            excel_logger.log_excel_step(target_path, "Merging header cells")
        merge_header_cells(table_sheet)
        
        log_info("✅  Final table rectangle finalization")
        if EXCEL_LOGGER_AVAILABLE and target_path:
            excel_logger.log_excel_step(target_path, "Final table rectangle finalization")
        finalize_table_rectangle(table_sheet)
        
        log_info("✅  Filling identification and labelTranslation")
        if EXCEL_LOGGER_AVAILABLE and target_path:
            excel_logger.log_excel_step(target_path, "Filling identification and labelTranslation sheets")
        fill_identification_and_labelTranslation(source_sheet, target_wb, form_group_code, version, validity_starting_date, linked_domain, datatype, roundingUnit)
            
    finally:
        # Always restore original logging functions
        if dedicated_logger:
            logging.debug = original_debug
            logging.info = original_info
    

def main():
    # Reset global validation file path at the start of each run
    global VALIDATION_FILE_PATH
    VALIDATION_FILE_PATH = None

    # Capture initial logs to replay them in each file's log
    initial_logs = []
    
    # Get user input (source workbook path + debug mode)
    msg = "🔧  Getting user input"
    logging.info(msg)
    initial_logs.append(("info", msg))
    
    source_path, form_code, form_group_code, version, validity_starting_date, linked_domain, datatype, roundingUnit = get_user_input()
    output_dir = create_output_directory(source_path)

    # Setup validation rules file if needed
    msg = "🔧  Setting up validation rules file (optional - you can skip this step)"
    logging.info(msg)
    initial_logs.append(("info", msg))
    
    if FEEDING_RULES_AVAILABLE and VALIDATION_FILE_PATH is None:
        try:
            VALIDATION_FILE_PATH = feeding_rules.get_validation_rules_file()
            if VALIDATION_FILE_PATH:
                msg = f"   ✅ Validation rules file selected: {VALIDATION_FILE_PATH}"
                logging.info(msg)
                initial_logs.append(("info", msg))
            else:
                msg = "   ℹ️  No validation rules file selected - continuing without feeding rules"
                logging.info(msg)
                initial_logs.append(("info", msg))
        except Exception as e:
            msg = f"   ❌ Error getting validation rules file: {str(e)}"
            logging.error(msg)
            initial_logs.append(("error", msg))
            VALIDATION_FILE_PATH = None

    # Load the source workbook (read-only for data)
    msg = f"🔧  Loading source workbook: {source_path}"
    logging.info(msg)
    initial_logs.append(("info", msg))
    source_wb = openpyxl.load_workbook(source_path, data_only=True)
    logging.debug(f"   Loaded source workbook: {source_path}")

    # Get the relevant sheets and group them by root name
    msg = "🔧  Getting relevant sheets and grouping by root name"
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
    
    grouped_sheets = group_sheets_by_root_name(relevant_sheets)
    
    # Log grouping details
    for sheet in relevant_sheets:
        root_name = extract_root_name(sheet)
        msg = f"   Grouping: {sheet} -> {root_name}"
        logging.debug(msg)
        initial_logs.append(("debug", msg))

    # Process each root_name group
    generated_files = []  # Track generated files for summary
    
    for root_name, sheet_list in grouped_sheets.items():
        logging.info("")
        logging.info("=====================================")
        logging.info("== 📢 PROCESSING GROUP SHEET DATA ==")
        logging.info(f" Group '{root_name}' with sheets: {sheet_list}")
        logging.info("=====================================")

        try:
            # Copy the template file once for this group
            target_path, dedicated_logger = copy_template_file(root_name, output_dir, form_code)
            if not target_path:
                logging.error(f"Could not create target file for '{root_name}'. Skipping.")
                continue

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
                    root_name=root_name,
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
            create_empty_target_sheets(target_wb, sheet_list)

            # Now, process each sheet (adjust headers, paste data, etc.)
            for idx, sheet_name in enumerate(sheet_list):
                # Only pass optional field values for the first sheet
                # to avoid overwriting identification D2/E2 with the last sheet's name
                if idx == 0:
                    process_sheet_data(source_wb, target_wb, sheet_name, form_group_code, version, validity_starting_date, linked_domain, datatype, roundingUnit, target_path, dedicated_logger)
                else:
                    process_sheet_data(source_wb, target_wb, sheet_name, '', '', '', '', '', '', target_path, dedicated_logger)

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

            # Save and close the target workbook
            logging.info(f"🔧  Saving and closing target workbook: {target_path}")
            if EXCEL_LOGGER_AVAILABLE:
                excel_logger.log_excel_step(target_path, "Saving target workbook")
            target_wb.save(target_path)
            target_wb.close()

            # Apply feeding rules from validation rules if available
            if FEEDING_RULES_AVAILABLE and VALIDATION_FILE_PATH:
                logging.info("🔧  Applying feeding rules from pure validation rules")
                if EXCEL_LOGGER_AVAILABLE:
                    excel_logger.log_excel_step(target_path, "Applying validation rules")
                try:
                    success = feeding_rules.process_excel_with_validation_rules(
                        target_path, 
                        validation_file_path=VALIDATION_FILE_PATH,
                        debug_mode=DEBUG_MODE,
                        dedicated_logger=dedicated_logger
                    )
                    if success:
                        logging.info("   ✅ Feeding rules applied successfully")
                        if EXCEL_LOGGER_AVAILABLE:
                            excel_logger.log_excel_step(target_path, "Validation rules applied successfully")
                    else:
                        logging.warning("   ⚠️  Feeding rules processing completed with warnings")
                        if EXCEL_LOGGER_AVAILABLE:
                            excel_logger.log_excel_step(target_path, "Validation rules applied with warnings")
                except Exception as e:
                    logging.error(f"   ❌ Error applying feeding rules: {str(e)}")
                    if EXCEL_LOGGER_AVAILABLE:
                        excel_logger.log_excel_error(target_path, f"Error applying validation rules: {str(e)}", e)
            else:
                if FEEDING_RULES_AVAILABLE:
                    logging.info("🔧  Skipping feeding rules (no validation file selected)")
                else:
                    logging.info("🔧  Skipping feeding rules (processor not available)")

            # Finalize with Excel for border compatibility
            logging.info("🔧  Finalizing with Excel Windows saving for border compatibility")
            if EXCEL_LOGGER_AVAILABLE:
                excel_logger.log_excel_step(target_path, "Finalizing with Excel Windows for border compatibility")
            try:
                autosave_with_excel(target_path)
                logging.debug(f"   Finalized file with Excel: {target_path}")
            except Exception as e:
                logging.warning(f"  ⚠ Excel autosave failed: {e} (file saved by openpyxl)")
                if EXCEL_LOGGER_AVAILABLE:
                    excel_logger.log_excel_error(target_path, f"Excel autosave failed: {e}", e)

            # Finalize Excel logging
            if EXCEL_LOGGER_AVAILABLE:
                excel_logger.complete_excel_logging(
                    target_path, 
                    success=True,
                    total_sheets=len(sheet_list),
                    processed_sheets=sheet_list,
                    form_code=form_code,
                    feeding_rules_applied=FEEDING_RULES_AVAILABLE and VALIDATION_FILE_PATH is not None
                )

            generated_files.append(target_path)
            logging.info(f"✅ Successfully generated: {target_path}")

        except Exception as e:
            logging.error(f"❌ Error processing group '{root_name}': {str(e)}")
            if EXCEL_LOGGER_AVAILABLE and 'target_path' in locals():
                excel_logger.log_excel_error(target_path, f"Error processing group '{root_name}': {str(e)}", e)
                excel_logger.complete_excel_logging(target_path, success=False, error=str(e))
            continue

    # Close the source workbook
    logging.info("🔧  Closing source workbook")
    source_wb.close()
    
    # Summary of generated files
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
