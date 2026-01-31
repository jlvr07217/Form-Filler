import pandas as pd
import os

def create_excel_file(filename="dummy_inventory.xlsx"):
    """
    Creates a simple Excel file with mock inventory data for testing purposes.
    The data structure matches the data used in test_standard_conversion.
    """
    # Data matching the structure in test_main.py
    data = {
        'ID': [101, 102],
        'Product': ['Laptop', 'Monitor'],
        'Price': [1200.50, 350.00]
    }

    # Create the DataFrame
    df = pd.DataFrame(data)

    # Export to an Excel file
    try:
        df.to_excel(filename, index=False)
        print(f"Successfully created test Excel file: {os.path.abspath(filename)}")
    except Exception as e:
        print(f"An error occurred while creating the Excel file: {e}")

if __name__ == "__main__":
    create_excel_file()