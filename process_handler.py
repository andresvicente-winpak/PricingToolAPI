import zipfile
import os
import tempfile
import pandas as pd  # make sure pandas is imported if you read Excel

def process_file(input_path):
    try:
        # Optional: print available sheets for debugging
        xls = pd.ExcelFile(input_path)
        print("✅ Available sheets:", xls.sheet_names)

        # Create a temporary output folder
        temp_dir = tempfile.mkdtemp()

        # Example: Save just the raw input file for now (for debugging)
        raw_copy_path = os.path.join(temp_dir, os.path.basename(input_path))
        with open(input_path, 'rb') as f_in, open(raw_copy_path, 'wb') as f_out:
            f_out.write(f_in.read())

        # Create a zip archive
        zip_path = os.path.join(temp_dir, 'output.zip')
        with zipfile.ZipFile(zip_path, 'w') as zipf:
            zipf.write(raw_copy_path, arcname=os.path.basename(input_path))

        print(f"📦 ZIP created at: {zip_path}")
        return zip_path

    except Exception as e:
        import traceback
        traceback.print_exc()
        raise e  # Let the error bubble up to Flask
