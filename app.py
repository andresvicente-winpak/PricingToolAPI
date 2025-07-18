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
        zip_output_dir = os.path.join(work_dir, "zips")

        os.makedirs(input_dir, exist_ok=True)
        os.makedirs(sample_dir, exist_ok=True)
        os.makedirs(output_dir, exist_ok=True)
        os.makedirs(zip_output_dir, exist_ok=True)

        # Extract uploaded zip
        with zipfile.ZipFile(uploaded_zip, 'r') as z:
            z.extractall(input_dir)

        # Load config
        config_path = os.path.join(input_dir, "configuration.xlsx")
        config_df = load_excel_file(config_path, header=0, dtype=str)
        config_df.columns = [str(col).strip() for col in config_df.columns]
        config_df = config_df.dropna(subset=['Dest_table', 'Dest_field'])

        # Process all Excel files and zip them individually
        price_input_path = os.path.join(input_dir, "PriceList Input")
        for fname in os.listdir(price_input_path):
            if fname.endswith(('.xlsx', '.xls')):
                full_path = os.path.join(price_input_path, fname)
                process_file(full_path, config_df, sample_dir, output_dir)

                # Zip that Excel's specific output folder
                excel_name = os.path.splitext(fname)[0]
                excel_output_folder = os.path.join(output_dir, excel_name)
                individual_zip_path = os.path.join(zip_output_dir, f"{excel_name}.zip")

                with zipfile.ZipFile(individual_zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
                    for root, _, files in os.walk(excel_output_folder):
                        for file in files:
                            full_path = os.path.join(root, file)
                            arcname = os.path.basename(full_path)
                            zipf.write(full_path, arcname)

        # Bundle all per-input ZIPs into one master ZIP
        final_zip = BytesIO()
        with zipfile.ZipFile(final_zip, 'w') as master_zip:
            for zip_name in os.listdir(zip_output_dir):
                zip_path = os.path.join(zip_output_dir, zip_name)
                master_zip.write(zip_path, arcname=zip_name)
        final_zip.seek(0)

        return send_file(final_zip, download_name='output.zip', as_attachment=True)

    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
