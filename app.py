from flask import Flask, request, send_file, jsonify
from FixedLoadSheet import process_file, load_excel_file
import pandas as pd
import os
import tempfile
import zipfile
from io import BytesIO

app = Flask(__name__)

@app.route('/process', methods=['POST'])
def process():
    try:
        print("📥 Receiving file...")
        uploaded_excel = BytesIO(request.get_data())
        print("✅ File received, loading Excel...")

        work_dir = tempfile.mkdtemp()
        input_dir = os.path.join(work_dir, "input")
        sample_dir = os.path.join(work_dir, ".")
        output_dir = os.path.join(work_dir, "output")

        os.makedirs(input_dir, exist_ok=True)
        os.makedirs(sample_dir, exist_ok=True)
        os.makedirs(output_dir, exist_ok=True)

        # Save uploaded Excel to input folder
        input_excel_path = os.path.join(input_dir, "input.xlsx")
        with open(input_excel_path, 'wb') as f:
            f.write(uploaded_excel.read())
        print("✅ Excel saved to disk.")

        # Load local configuration file
        config_path = os.path.join(os.getcwd(), "configuration.xlsx")
        print(f"🔍 Loading config from {config_path}")
        config_df = load_excel_file(config_path, header=0, dtype=str)
        config_df.columns = [str(col).strip() for col in config_df.columns]
        config_df = config_df.dropna(subset=['Dest_table', 'Dest_field'])

        # Define path to local sample CSVs
        local_sample_dir = os.getcwd()

        # Process the uploaded Excel
        process_file(input_excel_path, config_df, local_sample_dir, output_dir)

        # Locate subfolder inside output_dir
        subfolders = [d for d in os.listdir(output_dir) if os.path.isdir(os.path.join(output_dir, d))]
        if not subfolders:
            raise FileNotFoundError("No output files generated.")
        target_output = os.path.join(output_dir, subfolders[0])

        # Zip the processed subfolder
        zip_buffer = BytesIO()
        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for root, _, files in os.walk(target_output):
                for file in files:
                    full_path = os.path.join(root, file)
                    arcname = os.path.basename(full_path)
                    zipf.write(full_path, arcname)
        zip_buffer.seek(0)

        return send_file(zip_buffer, download_name='output.zip', as_attachment=True)

    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
