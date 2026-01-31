# This program uses the fillpdf library to fill PDF forms.
# It will take data from an excel file and populate the corresponding fields in a PDF form.
# The excel data will be turned into a dictionary for easy access on separate file.
import pandas as pd
from fillpdf import fillpdfs
import datetime

def excel_to_dict(excel_file_path):
    # Read the excel file
    df = pd.read_excel(excel_file_path)

    # Convert each row to a dictionary and store in a list
    data_dicts = df.to_dict(orient='records')

    return data_dicts



# Function to get PDF fields and return them as a dictionary
def get_pdf_fields(pdf_file):
    return fillpdfs.get_form_fields(pdf_file)


# Function to fill PDF form fields with data from a dictionary
def fill_pdf_form(pdf_file, data_dict, output_pdf):
    pdf_fields = get_pdf_fields(pdf_file)
    print("Detected PDF fields:", list(pdf_fields.keys()))

    # Helper to normalize names (remove non-alnum, lowercase) for fuzzy matching
    def _normalize(name):
        if name is None:
            return ""
        return "".join(ch.lower() for ch in str(name) if ch.isalnum())

    # Build normalized lookup for incoming data keys (excel columns)
    data_norm_map = { _normalize(k): k for k in data_dict.keys() }

    for pdf_field_name in pdf_fields:
        norm_pdf = _normalize(pdf_field_name)

        # Try to find the corresponding key in the excel data
        data_key = None
        if norm_pdf in data_norm_map:
            data_key = data_norm_map[norm_pdf]
        else:
            # fallback: case-insensitive exact match on stripped names
            candidates = [k for k in data_dict.keys() if str(k).strip().lower() == str(pdf_field_name).strip().lower()]
            if candidates:
                data_key = candidates[0]

        if data_key is None:
            # nothing to fill for this pdf field
            # print(f"No matching excel column for PDF field '{pdf_field_name}'")
            continue

        value = data_dict.get(data_key)
        if pd.isna(value):
            # skip empty values
            # print(f"Skipping empty value for field '{pdf_field_name}' (key: {data_key})")
            continue

        # Attempt to coerce value to a datetime. If successful, format as MM/DD/YYYY.
        try:
            coerced = pd.to_datetime(value, errors='coerce')
        except Exception:
            coerced = pd.NaT

        if not pd.isna(coerced):
            try:
                pdf_fields[pdf_field_name] = coerced.strftime("%m/%d/%Y")
                print(f"Filled date field '{pdf_field_name}' with '{pdf_fields[pdf_field_name]}' from column '{data_key}'")
            except Exception as e:
                print(f"Error formatting date for {pdf_field_name}: {e}")
                pdf_fields[pdf_field_name] = str(value)
        else:
            # Not a date — write the string representation
            pdf_fields[pdf_field_name] = str(value)
            # print(f"Filled field '{pdf_field_name}' with '{pdf_fields[pdf_field_name]}'")

    fillpdfs.write_fillable_pdf(pdf_file, output_pdf, pdf_fields)
    return output_pdf



if __name__ == "__main__":
    # Example usage
    excel_file = "test.xlsx"
    pdf_file = "C:\\Users\\jlvr0\\Downloads\\Example Fillable PDF.pdf"
    data = excel_to_dict(excel_file)
    print(data)

    fill_pdf_form(pdf_file, data[0], "filled_form.pdf")
    