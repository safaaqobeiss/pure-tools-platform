"""
### 🔧 **Prompt 

**Generate automatically the identical cells controls in the pure-regulation file for each model. 

1. Create Python code to generate a new target Excel file with a new sheet "identicalcells" from information of the source workbook.
   * Ask the user the path of the "identical_cells" source workbook (with file dialog)
   * Ask the user the path of the pure regulation target workbook to copy (with file dialog)
   * Ask the user the Module1 as a mandatory field and Module2, Module3 and Module4 more as optional fields, and "validityStartingDate" as mandatory field
   * Inform the user that the folder "generated_output" for generated files will be created if not exist at the grand-parent level of the source workbook (../) 
   * The new target file created will be a copy of the selected pure regulation target workbook, with an "_identical-cells" suffix added to its filename.
   * The template file to be used is here "../../template/pure-regulation-regulationCode.xlsx"

2. A new sheet
   * A new sheet named "identicalCells" must be added to this new target file. 
   * The sheet should copy the structure and formatting from the "validationRuleSetCode" sheet in the provided template file at: ../../template/pure-regulation-regulationCode.xlsx

3. Feed validationRuleSet sheet of the target file
    * Copy layout (border and all data validation associated to line 2)
    * Paste the layout at the end of the table (meaning all values of a row are empty) to create a new line
    * Add "identicalcells" as code (meaning in column A) 
    * Add in columns B, C and K same values as rhe row above (for origin, version and type column)

3. Read each line of the source unique worksheet
    * Select only the rows in the Module1 column the input of the user, and same thing if any values for Module2, Module3 and Module4 
    * Formula is composed of different blocks to compare (e.g. {C_34.09, r0010, c0010}=={K_06.00, r0010, c0010} or multiple blcoks as {D_07.00.a, r0320, c0060}=={D_09.01, r0010, c0060}=={K_47.00.a, r0320, c0060}=={K_49.01, r0010, c0060})
        * Parse each block of "Identical cells" value 
            * such as e.g. {C_34.09, r0010, c0010} ==> C_34.09 : Form, r0010 : Row and c0010 : Column
        * Transform and copy value within column "formula" 
            * for 2 blocks: {Form1, Row1, Column1} =  {Form2, Row2, Column2}
            * for muliple blocks : {Form1, Row1, Column1} =  {Form2, Row2, Column2} and {Form1, Row1, Column1} = {Form3, Row3, Column3} and  {Form1, Row1, Column1} = {Form4, Row4, Column4} 
    * Write in the column errorMessage (EN) :  
            * "The consistency validation {Form1, Row1, Column1} =  {Form2, Row2, Column2} for the scope : Listoftables is wrong"
            * by replacing the variables, e.g.: The consistency validation {C_34.09, r0010, c0010} = {K_06.00, r0010, c0010} for the scope : tables C_34.09 and K_06.00 is wrong
    * Write in the column errorMessage (FR) :  
            * "Le contrôle de cohérence {Form1, Row1, Column1} =  {Form2, Row2, Column2} = {Form3, Row3, Column3} pour le scope : Listoftables est faux"
            * by replacing the variables, e.g.: Le contrôle de cohérence {D_07.00.a, r0320, c0060}=={D_09.01, r0010, c0060}=={K_47.00.a, r0320, c0060} pour le scope : tables D_07.00.a and D_09.01 and K_47.00 and K_49.01 est faux
    * Each formula should be parsed block by block. r*, s*, or c* inside any block mean row_text, sheet_text, or column_text should be set to (All) respectively and be removed from the formula.
    * Fill all other cells from the target column mentionned below as such :
        * "code", "name(EN)" and "name(FR)"  ==> value of source "VariableVID" column with suffix "_i"
        * "status" ==> "enable" 
        * "severity" ==> "warning" 
        * "validityStartingDate" ==> user input value of same name
        * "missingValue" ==> "doNotRunRule" 
        * "arithmeticApproach" ==> "point" 

4. Post processings
    * Copy the controls, borders, alignment from row 2 of the target sheet identicalcells and paste it in all below non empty lines
    * Remove row 2

5. The final code must:
   * Include robust error handling
   * Provide detailed debugging information : there should be a debug mode that can be activated and asked to the user to provide maximum of information at each step
   * Properly handle file paths with spaces
   * Not alter any original files
   * Includes a main method
   * Includes at least 6 methods and is well split (functionally and technically speaking) 
   * Is split into at least 4 parts (as many as needed to group together few methods)
        * 1st part: the user input handling, directory preparation, and general setup (like debug mode) 
        * and last part : Main  and before last one : Post processing
   * Before each method, add a header such as :
   
###############################################################################################################################################################
#                                                                                                                                                             #
#       6th and last part: Main                                                                                                                               #
#                                                                                                                                                             #
#   Completes everything in a main() function!                                                                                                                #
#                                                                                                                                                             #                                                                                                                                               #
###############################################################################################################################################################


"""

import os
import shutil
import tkinter as tk
from tkinter import filedialog, simpledialog, messagebox
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from copy import copy
from copy import copy
import datetime
import re
import sys

# Import Excel logger
sys.path.append(os.path.join(os.path.dirname(__file__), '..', 'common'))
try:
    import excel_logger
    EXCEL_LOGGER_AVAILABLE = True
except ImportError:
    EXCEL_LOGGER_AVAILABLE = False


DEBUG_MODE = False

###############################################################################################################################################################
#                                                                                                                                                             #
#       1st part: User Input, Path Selection, Directory Preparation                                                                                           #
#                                                                                                                                                             #
#   Handles all input from user and sets up working directories                                                                                               #
#                                                                                                                                                             #
###############################################################################################################################################################

def get_user_inputs():
    global DEBUG_MODE

    root = tk.Tk()
    root.withdraw()

    identical_cells_path = filedialog.askopenfilename(title="Select source workbook (identical_cells)", filetypes=[("Excel files", "*.xlsx")])
    if not identical_cells_path:
        raise Exception("No source file selected.")

    target_pure_regulation_path = filedialog.askopenfilename(title="Select pure regulation target workbook", filetypes=[("Excel files", "*.xlsx")])
    if not target_pure_regulation_path:
        raise Exception("No target file selected.")

    module1, module2, module3, module4, validity_date = get_module_and_validity_inputs()
    debug_log(f"User inputs 1 : {module1, module2, module3, module4, validity_date}")

    debug_confirm = messagebox.askyesno("Debug Mode", "Enable debug mode for detailed logs?")
    DEBUG_MODE = debug_confirm

    return identical_cells_path, target_pure_regulation_path, module1, module2, module3, module4, validity_date


def get_module_and_validity_inputs(): 
    result = {}

    def on_submit():
        result["module1"] = entries["Module1"].get().strip()
        result["module2"] = entries["Module2"].get().strip()
        result["module3"] = entries["Module3"].get().strip()
        result["module4"] = entries["Module4"].get().strip()
        result["validity"] = entries["Validity Starting Date"].get().strip()
        form.quit()

    form = tk.Tk()
    form.title("Input Required Fields")
    form.geometry("450x260")
    form.resizable(False, False)

    labels = [
        "Module1 (optional)",
        "Module2 (optional)",
        "Module3 (optional)",
        "Module4 (optional)",
        "Validity Starting Date (mandatory)"
    ]
    keys = [
        "Module1",
        "Module2",
        "Module3",
        "Module4",
        "Validity Starting Date"
    ]
    entries = {}

    for i, label_text in enumerate(labels):
        tk.Label(form, text=label_text).grid(row=i, column=0, padx=10, pady=5, sticky="e")
        entry = tk.Entry(form, width=35)
        entry.grid(row=i, column=1, padx=10, pady=5)
        entries[keys[i]] = entry

    submit_btn = tk.Button(form, text="Submit", command=on_submit)
    submit_btn.grid(row=len(labels), column=0, columnspan=2, pady=20)

    form.mainloop()
    form.destroy()

    if not result:
        raise Exception("User closed the input form without submitting.")

    if not result["validity"]:
        raise Exception("Validity Starting Date is mandatory.")

    return (
        result["module1"] or None,
        result["module2"] or None,
        result["module3"] or None,
        result["module4"] or None,
        result["validity"]
    )




def prepare_output_directory(source_path):
    grandparent = os.path.abspath(os.path.join(source_path, "../../"))
    output_dir = os.path.join(grandparent, "generated_output")
    os.makedirs(output_dir, exist_ok=True)
    return output_dir

def debug_log(msg):
    if DEBUG_MODE:
        print(f"[DEBUG] {msg}")


###############################################################################################################################################################
#                                                                                                                                                             #
#       2nd part: Workbook and Template Preparation                                                                                                           #
#                                                                                                                                                             #
#   Loads workbooks, creates target copy from template, and prepares sheets                                                                                   #
#                                                                                                                                                             #
###############################################################################################################################################################

def create_target_workbook(template_path, regulation_workbook_path, source_path, output_dir):
    base_filename = os.path.splitext(os.path.basename(regulation_workbook_path))[0]
    output_file = os.path.join(output_dir, f"{base_filename}_identical-cells.xlsx")

    debug_log(f"Creating target workbook from regulation file: {regulation_workbook_path}")
    debug_log(f"Output file will be saved to: {output_file}")

    # Initialize Excel file logging if available
    if EXCEL_LOGGER_AVAILABLE:
        try:
            excel_logger.setup_excel_file_logging(output_file, "identical_cells")
            excel_logger.log_excel_step(output_file, "Creating target workbook", 
                                      f"Based on: {regulation_workbook_path}")
        except Exception as e:
            debug_log(f"Error initializing Excel logging: {e}")

    shutil.copy(regulation_workbook_path, output_file)
    debug_log(f"File successfully copied to: {output_file}")
    
    if EXCEL_LOGGER_AVAILABLE:
        excel_logger.log_excel_step(output_file, "Base file copied", 
                                  f"Output file: {output_file}")
    
    return output_file



def add_identical_cells_sheet(target_path, sheet_template_path):
    debug_log(f"Opening target workbook: {target_path}")
    target_wb = load_workbook(target_path)
    debug_log(f"Opening template workbook: {sheet_template_path}")
    template_wb = load_workbook(sheet_template_path, data_only=False)

    if "validationRuleSetCode" not in template_wb.sheetnames:
        raise Exception("Template missing 'validationRuleSetCode' sheet.")

    template_sheet = template_wb["validationRuleSetCode"]
    new_sheet = target_wb.create_sheet("identicalCells")
    debug_log("Created sheet 'identicalCells'")

    # Copy column widths
    for col in template_sheet.column_dimensions:
        new_sheet.column_dimensions[col].width = template_sheet.column_dimensions[col].width
    debug_log("Copied column widths")

    # Copy cell contents and styles
    for row in template_sheet.iter_rows():
        for cell in row:
            new_cell = new_sheet.cell(row=cell.row, column=cell.col_idx, value=cell.value)

            if cell.has_style:
                new_cell.font = copy(cell.font)
                new_cell.border = copy(cell.border)
                new_cell.fill = copy(cell.fill)
                new_cell.number_format = copy(cell.number_format)
                new_cell.protection = copy(cell.protection)
                new_cell.alignment = copy(cell.alignment)

    debug_log("Copied cell values and styles")

    # Copy data validations
    for dv in template_sheet.data_validations.dataValidation:
        copied_dv = DataValidation(
            type=dv.type,
            formula1=dv.formula1,
            formula2=dv.formula2,
            allow_blank=dv.allow_blank,
            showDropDown=dv.showDropDown,
            showInputMessage=dv.showInputMessage,
            showErrorMessage=dv.showErrorMessage,
            promptTitle=dv.promptTitle,
            prompt=dv.prompt,
            errorTitle=dv.errorTitle,
            error=dv.error,
            operator=dv.operator
        )
        for cell_range in dv.cells:
            copied_dv.add(cell_range)
        new_sheet.add_data_validation(copied_dv)

    debug_log("Copied data validations")

    template_wb.close()
    target_wb.save(target_path)
    target_wb.close()
    debug_log(f"'identicalCells' sheet fully copied and saved to: {target_path}")




###############################################################################################################################################################
#                                                                                                                                                             #
#       3rd part: Populate Validation Sheet and Extract Identical Cell Rules                                                                                  #
#                                                                                                                                                             #
#   Fill validationRuleSet sheet, extract valid identical formulas                                                                                            #
#                                                                                                                                                             #
###############################################################################################################################################################

def update_validation_ruleset(target_path):
    debug_log(f"Opening workbook to update 'validationRuleSet': {target_path}")
    wb = load_workbook(target_path)
    sheet = wb["validationRuleSet"]

    last_row = sheet.max_row + 1
    origin = sheet[f"B{last_row - 1}"].value
    version = sheet[f"C{last_row - 1}"].value
    type_val = sheet[f"K{last_row - 1}"].value

    debug_log(f"Appending new rule set entry at row {last_row}")
    debug_log(f"Origin: {origin}, Version: {version}, Type: {type_val}")

    for col in range(1, 15):
        source_cell = sheet.cell(row=2, column=col)
        target_cell = sheet.cell(row=last_row, column=col)
        target_cell._style = source_cell._style
        target_cell.data_type = source_cell.data_type

    sheet[f"A{last_row}"] = "identicalcells"
    sheet[f"B{last_row}"] = origin
    sheet[f"C{last_row}"] = version
    sheet[f"K{last_row}"] = type_val

    wb.save(target_path)
    wb.close()
    debug_log("ValidationRuleSet updated with 'identicalcells' entry")



def extract_and_populate_identical_cells(source_path, target_path, modules, validity_date):
    debug_log(f"Reading from source workbook: {source_path}")
    source_wb = load_workbook(source_path, data_only=True)
    source_ws = source_wb.active

    debug_log(f"Reading headers to determine column positions")
    header_row = next(source_ws.iter_rows(min_row=1, max_row=1))
    headers = {cell.value: idx for idx, cell in enumerate(header_row, start=1)}
    debug_log(f"Found headers: {headers}")

    required_columns = ["Module1", "Identical cells", "VariableVID"]
    for col in required_columns:
        if col not in headers:
            raise Exception(f"❌ Required column '{col}' not found in source workbook.")

    # Optional module columns: check only if user provides values
    optional_modules = ["Module2", "Module3", "Module4"]
    for i, mod in enumerate(modules[1:], start=2):
        if mod and f"Module{i}" not in headers:
            raise Exception(f"❌ Optional column 'Module{i}' provided by user, but not found in workbook.")

    debug_log(f"Opening target workbook to write identical cells: {target_path}")
    target_wb = load_workbook(target_path)
    target_ws = target_wb["identicalCells"]

    total_rows = 0
    matched_rows = 0
    written_rows = 0

    for row in source_ws.iter_rows(min_row=2):
        total_rows += 1
     
        if modules[0] and row[headers["Module1"] - 1].value != modules[0]: continue
        if modules[1] and row[headers["Module2"] - 1].value != modules[1]: continue
        if modules[2] and row[headers["Module3"] - 1].value != modules[2]: continue
        if modules[3] and row[headers["Module4"] - 1].value != modules[3]: continue

        matched_rows += 1
        formula_raw = row[headers["Identical cells"] - 1].value
        vid_code = row[headers["VariableVID"] - 1].value

        if not formula_raw or not vid_code:
            continue

        # Determine row_text and sheet_text based on presence of r* or s*
        row_text = sheet_text = column_text = None
        if "s*" in formula_raw:
            sheet_text = "(All)"
        if "r*" in formula_raw:
            row_text = "(All)"
        if "c*" in formula_raw:
            column_text = "(All)"

     
        # Parse the formula  
        blocks = re.findall(r"\{([^}]+)\}", formula_raw)

        if not blocks or len(blocks) < 2:
            debug_log(f"⚠️ Skipped row {total_rows}: invalid formula format → {formula_raw}")
            continue

        parsed_blocks = []
        forms_seen = []

        for b in blocks:
            parts = [p.strip() for p in b.split(',') ]
            form = parts[0]
            row_id = parts[1] if len(parts) > 1 else ""
            col_id = parts[2] if len(parts) > 2 else ""
            parsed_blocks.append(f"{{{form}, {row_id}, {col_id}}}")
            if form not in forms_seen:
                forms_seen.append(form)

        # ✅ Build final comparison text (for 'formula' column)
        if len(parsed_blocks) == 2:
            formula = f"{parsed_blocks[0]} = {parsed_blocks[1]}"
        else:
            formula = " and ".join([f"{parsed_blocks[0]} = {b}" for b in parsed_blocks[1:]])
        
        # ✅ formula_text within error description
        formula_text = " = ".join(parsed_blocks)

        # ✅ Remove wildcard arguments from formula_raw (e.g. ", s*", ", r*", ", c*")
       # formula = re.sub(r"(,\s*(r\*|s\*|c\*))(,)?", lambda m: "," if m.group(3) else "", formula)
        formula = re.sub(r",\s*(s\*|r\*|c\*)", "", formula)
        formula_text = re.sub(r",\s*(s\*|r\*|c\*)", "", formula_text)

        # ✅ forms_text with " and " / " et " separator (preserves order, removes duplicates)
        forms_text_en = " and ".join(forms_seen)
        forms_text_fr = " et ".join(forms_seen)

        # ✅ Used for error messages
        error_en = f"The consistency validation {formula_text} for the scope : tables {forms_text_en} is wrong"
        error_fr = f"Le contrôle de cohérence {formula_text} pour le scope : tables {forms_text_fr} est faux"

        new_row = [ 
            str(vid_code) + "_i",                      # code
            "enable",                             # status
            "warning",                            # severity 
            None, None, None,                     # currency, roundingUnit, tolerance
            validity_date,                        # validityStartingDate
            None, None,                           # validityEndingDate, type
            forms_text_en,                        # tables
            sheet_text,                           # sheets
            row_text,                             # rows
            column_text,                          # columns
            formula,                              # formula
            "doNotRunRule",                       # missingValue
            "point",                              # arithmeticApproach
            str(vid_code) + "_i",                 # name(EN), name(FR)
            str(vid_code) + "_i",                 # name(EN), name(FR)
            error_en,                             # errorMessage (EN)
            error_fr,                             # errorMessage (FR)
            None, None                            # errorDescription (EN), errorDescription (FR)
        ]
        target_ws.append(new_row)
        written_rows += 1
        debug_log(f"✅ Added row for VID {vid_code}: {formula}")

    debug_log(f"Scanned {total_rows} rows. Matched {matched_rows}. Rules written: {written_rows}.")

    source_wb.close()
    target_wb.save(target_path)
    target_wb.close()
    debug_log("Finished writing identical cell rules to target workbook.")



###############################################################################################################################################################
#                                                                                                                                                             #
#       4th part: Post-Processing                                                                                                                              #
#                                                                                                                                                             #
#   Can be expanded to copy rules into different format or workbook                                                                                           #
#                                                                                                                                                             #
###############################################################################################################################################################


def post_process_output(output_path):
    debug_log(f"Opening workbook for post-processing: {output_path}")
    wb = load_workbook(output_path)
    
    if "identicalCells" not in wb.sheetnames:
        raise Exception("❌ Sheet 'identicalCells' not found in the workbook.")

    ws = wb["identicalCells"]

    # Copy styles from row 2
    debug_log("Copying styles from row 2")
    style_template = {}
    for col in range(1, ws.max_column + 1):
        cell = ws.cell(row=2, column=col)
        style_template[col] = {
            "font": copy(cell.font),
            "border": copy(cell.border),
            "fill": copy(cell.fill),
            "alignment": copy(cell.alignment),
            "number_format": copy(cell.number_format),
            "protection": copy(cell.protection)
        }

    # Determine last row with content
    max_row = ws.max_row

    # Extract data validations from row 2 and clone them
    debug_log("Extracting data validations from row 2")
    new_validations = []
    if ws.data_validations:
        for dv in ws.data_validations.dataValidation:
            for cell_range in dv.cells:
                # For each column in row 2
                for col in range(1, ws.max_column + 1):
                    col_letter = get_column_letter(col)
                    if f"{col_letter}2" in cell_range:
                        target_range = f"{col_letter}3:{col_letter}{max_row}"
                        new_dv = DataValidation(
                            type=dv.type,
                            formula1=dv.formula1,
                            formula2=dv.formula2,
                            allow_blank=dv.allow_blank,
                            showDropDown=dv.showDropDown,
                            showInputMessage=dv.showInputMessage,
                            showErrorMessage=dv.showErrorMessage,
                            promptTitle=dv.promptTitle,
                            prompt=dv.prompt,
                            errorTitle=dv.errorTitle,
                            error=dv.error,
                            operator=dv.operator
                        )
                        new_dv.add(target_range)
                        new_validations.append(new_dv)

    # Apply styles from row 2 to all rows below
    debug_log("Applying styles to rows 3 and below")
    for row in ws.iter_rows(min_row=3, max_row=max_row):
        if all(cell.value is None for cell in row):
            continue
        for cell in row:
            col = cell.column
            style = style_template.get(col)
            if style:
                cell.font = style["font"]
                cell.border = style["border"]
                cell.fill = style["fill"]
                cell.alignment = style["alignment"]
                cell.number_format = style["number_format"]
                cell.protection = style["protection"]

    # Add cloned validations
    debug_log(f"Adding {len(new_validations)} data validations")
    for dv in new_validations:
        ws.add_data_validation(dv)

    # Delete row 2
    debug_log("Deleting row 2")
    ws.delete_rows(2)

    wb.save(output_path)
    wb.close()
    debug_log("Post-processing complete and saved.")



###############################################################################################################################################################
#                                                                                                                                                             #
#       6th and last part: Main                                                                                                                               #
#                                                                                                                                                             #
#   Completes everything in a main() function!                                                                                                                #
#                                                                                                                                                             #
###############################################################################################################################################################

def main():
    try:
        src_path, reg_path, m1, m2, m3, m4, validity = get_user_inputs()
        debug_log(f"User inputs : {src_path, reg_path, m1, m2, m3, m4, validity}")
        output_dir = prepare_output_directory(src_path)
        template_path = "../../template/pure-regulation-regulationCode.xlsx"

        output_file = create_target_workbook(template_path, reg_path, src_path, output_dir)
        
        # Log the start of creation if logger available
        if EXCEL_LOGGER_AVAILABLE:
            excel_logger.excel_file_logger.log_excel_creation_start(
                output_file,
                source_file=src_path,
                regulation_file=reg_path,
                modules=[m1, m2, m3, m4],
                validity_date=validity
            )

        if EXCEL_LOGGER_AVAILABLE:
            excel_logger.log_excel_step(output_file, "Adding identicalCells sheet")
        add_identical_cells_sheet(output_file, template_path)

        if EXCEL_LOGGER_AVAILABLE:
            excel_logger.log_excel_step(output_file, "Updating validationRuleSet sheet")
        update_validation_ruleset(output_file)
        
        if EXCEL_LOGGER_AVAILABLE:
            excel_logger.log_excel_step(output_file, "Extracting and populating identical cells", 
                                      f"Modules processed: {[m for m in [m1, m2, m3, m4] if m]}")
        extract_and_populate_identical_cells(src_path, output_file, [m1, m2, m3, m4], validity)
        
        if EXCEL_LOGGER_AVAILABLE:
            excel_logger.log_excel_step(output_file, "Post-processing (formatting and cleanup)")
        post_process_output(output_file)

        print(f"\n✅ Process completed. File saved at: {output_file}")
        
        # Complete Excel logging
        if EXCEL_LOGGER_AVAILABLE:
            modules_processed = [m for m in [m1, m2, m3, m4] if m]
            excel_logger.complete_excel_logging(
                output_file,
                success=True,
                modules_processed=modules_processed,
                validity_date=validity,
                source_file=src_path
            )

    except Exception as e:
        print(f"❌ Error: {str(e)}")
        
        # Log error if logger available and output_file was created
        if EXCEL_LOGGER_AVAILABLE and 'output_file' in locals():
            excel_logger.log_excel_error(output_file, f"Error during processing: {str(e)}", e)
            excel_logger.complete_excel_logging(output_file, success=False, error=str(e))

if __name__ == "__main__":
    main()
