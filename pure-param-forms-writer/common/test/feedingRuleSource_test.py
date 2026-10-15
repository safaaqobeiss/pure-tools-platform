import unittest
from unittest.mock import patch, MagicMock
import os
import sys
import openpyxl
import shutil
import tempfile
import logging

# Ensure the main script is importable
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

import feedingRuleSource_in_forms

def compare_excels(file1, file2, sheet_name=None, compare_formatting=True):
    """Compare two Excel files. If sheet_name is None, compare all sheets."""
    # Load workbooks without data_only to preserve formatting
    wb1 = openpyxl.load_workbook(file1, data_only=False)
    wb2 = openpyxl.load_workbook(file2, data_only=False)
    
    # Also load data-only versions for value comparison
    wb1_data = openpyxl.load_workbook(file1, data_only=True)
    wb2_data = openpyxl.load_workbook(file2, data_only=True)
    
    # Compare sheet names
    if wb1.sheetnames != wb2.sheetnames:
        wb1.close(); wb2.close(); wb1_data.close(); wb2_data.close()
        return False, f"Sheet names differ: {wb1.sheetnames} != {wb2.sheetnames}"
    
    # If specific sheet name provided, compare only that sheet
    if sheet_name:
        sheets_to_compare = [sheet_name]
    else:
        sheets_to_compare = wb1.sheetnames
    
    for sheet in sheets_to_compare:
        if sheet not in wb1.sheetnames or sheet not in wb2.sheetnames:
            wb1.close(); wb2.close(); wb1_data.close(); wb2_data.close()
            return False, f"Sheet '{sheet}' not found in one of the files"
            
        ws1 = wb1[sheet]
        ws2 = wb2[sheet]
        ws1_data = wb1_data[sheet]
        ws2_data = wb2_data[sheet]

        # Compare data values
        rows1 = list(ws1_data.iter_rows(values_only=True))
        rows2 = list(ws2_data.iter_rows(values_only=True))

        if len(rows1) != len(rows2):
            wb1.close(); wb2.close(); wb1_data.close(); wb2_data.close()
            return False, f"Row count differs in sheet '{sheet}': {len(rows1)} != {len(rows2)}"

        for i, (row1, row2) in enumerate(zip(rows1, rows2), 1):
            # Strict comparison - files must be identical
            if row1 != row2:
                wb1.close(); wb2.close(); wb1_data.close(); wb2_data.close()
                return False, f"Data difference in sheet '{sheet}' at row {i}: {row1} != {row2}"
        
        # Compare formatting if requested
        if compare_formatting:
            max_row = max(ws1.max_row, ws2.max_row)
            max_col = max(ws1.max_column, ws2.max_column)
            
            for row in range(1, max_row + 1):
                for col in range(1, max_col + 1):
                    cell1 = ws1.cell(row=row, column=col)
                    cell2 = ws2.cell(row=row, column=col)
                    
                    # Compare fill color
                    if str(cell1.fill) != str(cell2.fill):
                        wb1.close(); wb2.close(); wb1_data.close(); wb2_data.close()
                        return False, f"Fill color difference in sheet '{sheet}' at cell {cell1.coordinate}: {cell1.fill} != {cell2.fill}"
                    
                    # Compare font
                    if str(cell1.font) != str(cell2.font):
                        wb1.close(); wb2.close(); wb1_data.close(); wb2_data.close()
                        return False, f"Font difference in sheet '{sheet}' at cell {cell1.coordinate}: {cell1.font} != {cell2.font}"

    wb1.close()
    wb2.close()
    wb1_data.close()
    wb2_data.close()
    return True, "Excels are identical (data and formatting)"

class TestFeedingRuleSource(unittest.TestCase):
    def setUp(self):
        """Set up test environment before each test."""
        self.test_dir = os.path.dirname(__file__)
        self.input_dir = os.path.join(self.test_dir, "input")
        self.output_dir = os.path.join(self.test_dir, "generated_output")
        self.expected_dir = os.path.join(self.test_dir, "expected_result")
        
        # Clean output directory before each test
        if os.path.exists(self.output_dir):
            shutil.rmtree(self.output_dir)
        os.makedirs(self.output_dir, exist_ok=True)
        
        # Configure logging to suppress debug messages during tests
        logging.basicConfig(level=logging.ERROR)

    @patch('feedingRuleSource_in_forms.get_validation_rules_file')
    @patch('feedingRuleSource_in_forms.get_pure_form_file')
    def test_feeding_rule_source(self, mock_get_form_file, mock_get_validation_file):
        """Test the complete feeding rule source processing workflow."""
        # Configure logging
        import logging
        # logging.basicConfig(level=logging.DEBUG, format='[%(levelname)s] %(message)s', force=True)
        
        # Setup file paths
        form_file = os.path.join(self.input_dir, "pure-form-euebasolv-C_18.00-(C 18.00)_adapted_for_tests.xlsm")
        validation_file = os.path.join(self.input_dir, "pure-validationRules-eba-auto_adapted_for_tests.xlsx")
        expected_file = os.path.join(self.expected_dir, "pure-form-euebasolv-C_18.00-(C 18.00)_adapted_for_tests.xlsm")
        
        # Check if input files exist
        if not os.path.exists(form_file):
            self.skipTest(f"Form file not found: {form_file}")
        if not os.path.exists(validation_file):
            self.skipTest(f"Validation file not found: {validation_file}")
            
        # Mock the file selection functions
        mock_get_form_file.return_value = form_file
        mock_get_validation_file.return_value = validation_file
        
        # Run the main function with mocked file inputs and debug mode
        with patch('tkinter.messagebox.askyesno', return_value=True):  # Enable debug mode
            result = feedingRuleSource_in_forms.main()
        
        # Check that the script executed successfully
        self.assertEqual(result, 0, "Main function should return 0 for success")
        
        # The output directory is created in the test directory
        expected_output_dir = self.output_dir  # This is already correct: test/generated_output
        
        # Check that output directory was created
        self.assertTrue(os.path.exists(expected_output_dir), f"Output directory should be created: {expected_output_dir}")
        
        # Look for the generated file in the output directory
        generated_files = [f for f in os.listdir(expected_output_dir) if f.endswith('.xlsm')]
        self.assertGreater(len(generated_files), 0, "At least 1 Excel file should be generated")
        
        print("***** Test Result *****")
        print(f"Generated files: {generated_files}")
        print(f"Generated in: {expected_output_dir}")
        
        # Find the output file (should be the modified version of the input file)
        expected_output_filename = "pure-form-euebasolv-C_18.00-(C 18.00)_adapted_for_tests.xlsm"  # Modified name as per script logic
        output_file = os.path.join(expected_output_dir, expected_output_filename)
        
        # The output file should exist
        self.assertTrue(os.path.exists(output_file), f"Generated file should exist: {expected_output_filename} at path: {output_file}")
        
        # Compare with expected result if available
        if os.path.exists(expected_file):
            result, message = compare_excels(
                output_file, expected_file,
                compare_formatting=True
            )
            if result:
                print(f"[OK] {expected_output_filename}: Files are identical")
            else:
                print(f"[WARNING] {expected_output_filename}: {message}")
                
            # Assert strict comparison - files must be identical
            self.assertTrue(result, f"File comparison failed for {expected_output_filename}: {message}")
        else:
            print(f"[INFO] Expected file not found, skipping comparison: {expected_file}")
            
        print(f"[FILE] Generated file location: {output_file}")

    def tearDown(self):
        """Clean up after each test."""
        # Optionally clean up generated files
        # if os.path.exists(self.output_dir):
        #     shutil.rmtree(self.output_dir)
        pass

if __name__ == "__main__":
    # Set up logging for debugging
    logging.basicConfig(level=logging.INFO, format='[%(levelname)s] %(message)s')
    
    # Run tests
    unittest.main(verbosity=2)
