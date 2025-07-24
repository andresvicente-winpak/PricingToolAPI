import zipfile
import os
import tempfile
import pandas as pd

def process_file(input_path):
    # Create temporary directory
    temp_dir = tempfile.mkdtemp()
    zip_path = os.path.join(temp_dir, 'output.zip')
    
    # Load Excel file
    xls = pd.ExcelFile(input_path, engine='openpyxl')
    
    # Create a zip file
    with zipfile.ZipFile(zip_path, 'w') as zipf:
        for sheet_name in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet_name)
            csv_name = f"{sheet_name}.csv"
            csv_path = os.path.join(temp_dir, csv_name)
            df.to_csv(csv_path, index=False)
            zipf.write(csv_path, arcname=csv_name)

    return zip_path
