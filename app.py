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
        uploaded_zip = BytesIO(request.get_data())

        work_dir = tempfile.mkdtemp()
        input_dir = os.path.join(work_dir, "input")
        sample_dir = os.path.join(work_dir, "sample")
        output_dir = os.path.join(work_dir, "output")

        os.makedirs(input_dir, exist_ok=True)
        os.makedirs(sample_dir, exist_ok=True)
        os.makedirs(output_dir, exist_ok=True)

        # Extract uploaded zip
        with zipfile.ZipFile(uploaded_zip, 'r') as z:
            z.extractall(input_dir)

        # Load config
        config_path = os.path.join(input_dir, "configuration.xlsx")
        config_df = load_excel_file(config_path, header=0, dtype=str)
        config_df.columns = [str(col).strip() for col in config_df.columns]
        config_df = config_df.dropna(subset=['Dest_table', 'Dest_field'])

        # Process all Excel files
        price_input_path = os.path.join(input_dir, "PriceList Input")
        for fname in os.listdir(price_input_path):
            if fname.endswith(('.xlsx', '.xls')):
                full_path = os.path.join(price_input_path, fname)
                process_file(full_path, config_df, sample_dir, output_dir)

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
