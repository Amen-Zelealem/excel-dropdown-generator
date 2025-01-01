import os
import unittest
from openpyxl import load_workbook
from src.excel_dropdown import create_excel_with_dropdown

class TestExcelDropdown(unittest.TestCase):
    def setUp(self):
        """Set up test environment."""
        self.test_file = "test_dropdown.xlsx"
        self.sheet_name = "TestSheet"
        self.dropdown_column = "B"
        self.dropdown_values = ["Option1", "Option2", "Option3"]

    def tearDown(self):
        """Clean up after tests."""
        if os.path.exists(self.test_file):
            os.remove(self.test_file)

    def test_create_excel_with_dropdown(self):
        """Test if the Excel file with dropdown is created successfully."""
        # Call the function to create the Excel file
        create_excel_with_dropdown(
            self.test_file,
            self.sheet_name,
            self.dropdown_column,
            self.dropdown_values
        )

        # Check if the file was created
        self.assertTrue(os.path.exists(self.test_file))

        # Open the file and verify the dropdown
        workbook = load_workbook(self.test_file)
        sheet = workbook[self.sheet_name]

        # Validate that the dropdown values are applied to the correct column
        data_validations = sheet.data_validations.dataValidation
        dropdown_found = False
        for dv in data_validations:
            if dv.sqref == f"{self.dropdown_column}1:{self.dropdown_column}1048576":
                dropdown_found = True
                self.assertEqual(dv.formula1, f'"{",".join(self.dropdown_values)}"')
                break

        self.assertTrue(dropdown_found, "Dropdown validation not found in the specified column")

if __name__ == "__main__":
    unittest.main()
