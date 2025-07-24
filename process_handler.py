import os
import zipfile
import tempfile
import pandas as pd

def process_file(input_path):
    # Create temp folder
    temp_dir = tempfile.mkdtemp()

    # Load Excel
    xls = pd.ExcelFile(input_path)

    # Define output files
    outputs = {
        "1-pricelist.csv": xls.parse("Pricelist"),
        "1-baseprice.csv": xls.parse("Base Price"),
        "1-matrixprice.csv": xls.parse("Matrix Price"),
        "1-gradprice.csv": xls.parse("Grad Price")
    }

    # Save CSVs
    for filename, df in outputs.items():
        df.to_csv(os.path.join(temp_dir, filename), index=False)

    # Zip it
    zip_path = os.path.join(temp_dir, 'output.zip')
    with zipfile.ZipFile(zip_path, 'w') as zipf:
        for filename in outputs.keys():
            zipf.write(os.path.join(temp_dir, filename), arcname=filename)

    return zip_path
