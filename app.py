from flask import Flask, request, send_file, jsonify
from FixedLoadSheet import process_file, load_excel_file
import pandas as pd
import os
import tempfile
import zipfile
from io import BytesIO
import traceback

app = Flask(__name__)

@app.route('/process', methods=['POST'])
def process():
    try:
        print("📥 Receiving uploaded Excel file...")
        uploaded_excel = BytesIO(request.get_data())

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
        print(f"✅ File saved to: {input_excel_path}")

        # Load config
        config_path = os.path.join(os.getcwd(), "configuration.xlsx")
        print(f"📄 Looking for config at: {config_path}")
        config_df = load_excel_file(config_path, header=0, dtype=str)
        config_df.columns = [str(col).strip() for col in config_df.columns]
        config_df = config_df.dropna(subset=['Dest_table', 'Dest_field'])
        print(f"✅ Config loaded. Rows: {len(config_df)}")

        # Use current directory for sample CSVs
        local_sample_dir = os.getcwd()
        print(f"📁 Sample dir: {local_sample_dir}")

        # Process
        print(f"⚙️ Running process_file...")
        process_file(input_excel_path, config_df, local_sample_dir, output_dir)

        subfolders = [d for d in os.listdir(output_dir) if os.path.isdir(os.path.join(output_dir, d))]
        if not subfolders:
            raise FileNotFoundError("No output subfolder found.")
        target_output = os.path.join(output_dir, subfolders[0])

        # Zip result
        print(f"📦 Zipping output from: {target_output}")
        zip_buffer = BytesIO()
        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for root, _, files in os.walk(target_output):
                for file in files:
                    full_path = os.path.join(root, file)
                    arcname = os.path.basename(full_path)
                    zipf.write(full_path, arcname)
        zip_buffer.seek(0)

        print("✅ Processing complete. Returning ZIP.")
        return send_file(zip_buffer, download_name='output.zip', as_attachment=True)

    except Exception as e:
        print("❌ Exception occurred during processing:")
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
