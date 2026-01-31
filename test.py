import unittest
from unittest.mock import patch
import pandas as pd
from main import excel_to_dict

# We use patch to replace the actual pandas.read_excel function with a
# mock object during the test execution, preventing file system access.

class TestExcelToDict(unittest.TestCase):

    # --- Test Case 1: Standard conversion with multiple rows and mixed types ---
    @patch('main.pd.read_excel')
    def test_standard_conversion(self, mock_read_excel):
        """Tests conversion of a typical DataFrame into a list of dictionaries (one dict per row)."""

        # 1. Define the mock data the function should "read" from the file
        mock_data = {
            'ID': [101, 102],
            'Product': ['Laptop', 'Monitor'],
            'Price': [1200.50, 350.00]
        }
        mock_df = pd.DataFrame(mock_data)

        # Configure the mock function to return the mock DataFrame
        mock_read_excel.return_value = mock_df

        # 2. Define the expected output structure
        expected_output = [
            {'ID': 101, 'Product': 'Laptop', 'Price': 1200.50},
            {'ID': 102, 'Product': 'Monitor', 'Price': 350.00}
        ]

        # 3. Call the function with a dummy path
        result = excel_to_dict("dummy_inventory.xlsx")

        # 4. Assertions
        # Verify the function was called with the correct argument
        mock_read_excel.assert_called_once_with("dummy_inventory.xlsx")
        # Verify the result matches the expected structure
        self.assertEqual(result, expected_output)

    # --- Test Case 2: Handling a DataFrame with a single row ---
    @patch('main.pd.read_excel')
    def test_single_row(self, mock_read_excel):
        """Tests conversion when the Excel file only contains a single row of data."""

        mock_data = {
            'Task': ['Review Report'],
            'Completed': [True],
            'Priority': [1]
        }
        mock_df = pd.DataFrame(mock_data)
        mock_read_excel.return_value = mock_df

        expected_output = [
            {'Task': 'Review Report', 'Completed': True, 'Priority': 1}
        ]

        result = excel_to_dict("single_task.xlsx")
        self.assertEqual(result, expected_output)

    # --- Test Case 3: Handling an empty Excel file ---
    @patch('main.pd.read_excel')
    def test_empty_dataframe(self, mock_read_excel):
        """Tests the function when the Excel file is empty, ensuring an empty list is returned."""

        # Define an empty DataFrame with column names (optional, but robust)
        mock_df = pd.DataFrame(columns=['A', 'B'])
        mock_read_excel.return_value = mock_df

        # Expected output is an empty list
        expected_output = []

        result = excel_to_dict("empty_file.xlsx")
        self.assertEqual(result, expected_output)

# To run these tests from your terminal, you would navigate to the directory
# containing main.py and test_main.py, and run:
# python -m unittest test_main.py
