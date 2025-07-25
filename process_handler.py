from FixedLoadSheet import process_file as fls_process_file
import os
import tempfile
import pandas as pd
import zipfile

def process_file(input_path):
    try:

        # 🔍 DEBUG: Peek into first 100 bytes of the file
        with open(input_path, 'rb') as f:
            print("DEBUG - First 100 bytes of uploaded file:")
            print(f.read(100))
            
        # Debug check before reading as Excel
        if not zipfile.is_zipfile(input_path):
            raise ValueError(f"{input_path} is not a valid XLSX (ZIP) file — check Power Automate")
        
        # Create temp working directories
        temp_dir = tempfile.mkdtemp()
        output_dir = os.path.join(temp_dir, 'output')
        os.makedirs(output_dir, exist_ok=True)

        # Set paths
        config_path = os.path.join(os.getcwd(), 'configuration.xlsx')
        sample_dir = os.getcwd()  # sample files expected at repo root

        # Load configuration
        config_df = pd.read_excel(config_path, header=0, dtype=str, engine='openpyxl')
        config_df.columns = [str(col).strip() for col in config_df.columns]
        config_df = config_df.dropna(subset=['Dest_table', 'Dest_field'])

        # Run the real pricing logic
        fls_process_file(input_path, config_df, sample_dir, output_dir)

        # Return path to generated zip
        input_filename = os.path.splitext(os.path.basename(input_path))[0]
        zip_path = os.path.join(output_dir, f"{input_filename}.zip")
        return zip_path

    except Exception as e:
        import traceback
        traceback.print_exc()
        raise e
