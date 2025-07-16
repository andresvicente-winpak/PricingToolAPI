from flask import Flask, request, send_file, jsonify, request
from FixedLoadSheet import process_file, load_excel_file
import pandas as pd
import os
import tempfile
import zipfile
from io import BytesIO
import shutil

app = Flask(__name__)

@app.route('/process', methods=['POST'])
def process():
    try:
        # Get uploaded Excel file
        uploaded_excel = BytesIO(request.get_data())

        # Setup working dirs
        work_dir = tempfile.mkdtemp()
        input_dir = os.path.join(work_dir, "input")
        sample_dir = os.path.join(work_dir, "sample")
        output_dir = os.path.join(work_dir, "output")
        os.makedirs(input_dir, exist_ok=True)
        os.makedirs(sample_dir, exist_ok=True)
        os.makedirs(output_dir, exist_ok=True)

        # Save uploaded Excel
        input_excel_path = os.path.join(input_dir, "input.xlsx")
        with open(input_excel_path, 'wb') as f:
            f.write(uploaded_excel.read())

        # Load config
        config_path = os.path.join(input_dir, "configuration.xlsx")
        shutil.copy("configuration.xlsx", config_path)
        config_df = load_excel_file(config_path, header=0, dtype=str)
        config_df.columns = [str(col).strip() for col in config_df.columns]
        config_df = config_df.dropna(subset=['Dest_table', 'Dest_field'])

        # Copy static sample files into sample dir
        for name in ["1-baseprice.csv", "1-gradprice.csv", "1-matrixprice.csv"]:
            shutil.copy(name, os.path.join(sample_dir, name))

        # Process file
        process_file(input_excel_path, config_df, sample_dir, output_dir)

        # Create zip of output
        mem_zip = BytesIO()
        with zipfile.ZipFile(mem_zip, 'w') as zf:
            for root, _, files in os.walk(output_dir):
                for f in files:
                    fp = os.path.join(root, f)
                    arc = os.path.relpath(fp, output_dir)
                    zf.write(fp, arc)
        mem_zip.seek(0)

        return send_file(mem_zip, download_name='output.zip', as_attachment=True)

    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
