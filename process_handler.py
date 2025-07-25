from FixedLoadSheet import process_file as fls_process_file
import os
import tempfile
import pandas as pd

def process_file(input_path):
    try:
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
