from flask import Flask, request, send_file, jsonify
from FixedLoadSheet import process_file, load_excel_file
import pandas as pd
import os
import tempfile
import zipfile
from io import BytesIO
import base64

app = Flask(__name__)

@app.route('/process', methods=['POST'])
def process():
    try:
        print("📥 Receiving file...")

        # Raw data from Power Automate
        raw_data = request.get_data()

        # DEBUG: Save raw data for inspection
        debug_path = os.path.join(os.getcwd(), "uploaded_raw_preview.xlsx")
        with open(debug_path, "wb") as f:
            f.write(raw_data)
        print(f"🧪 Debug file written: {debug_path}")
        print(f"📏 Raw data size: {len(raw_data)} bytes")
        print(f"📄 Content-Type: {request.headers.get('Content-Type')}")

        # Attempt base64 decode if needed
        try:
            decoded_data = base64.b64decode(raw_data)
            uploaded_excel = BytesIO(decoded_data)
            print("✅ Base64 decoded successfully.")
        except Exception:
            uploaded_excel = BytesIO(raw_data)
            print("⚠️ Base64 decode skipped or failed — assuming raw binary.")

        # Create temporary working directories
        work_dir = tempfile.mkdtemp()
        input_dir = os.path.join(work_dir, "input")
        sample_dir = os.path.join(work_dir, "sample")
        output_dir = os.path.join(work_dir, "output")
        os.makedirs(input_dir, exist_ok=True)
        os.makedirs(sample_dir, exist_ok=True)
        os.makedirs(output_dir, exist_ok=True)

        # Save file to disk
        input_excel_path = os.path.join(input_dir, "input.xlsx")
        with open(input_excel_path, 'wb') as f:
            f.write(uploaded_excel.read())
        print(f"📁 Excel saved to: {input_excel_path}")

        # Load config
        config_path = os.path.join(os.getcwd(), "configuration.xlsx")
        config_df = load_excel_file(config_path, header=0, dtype=str)
        config_df.columns = [str(col).strip() for col in config_df.columns]
        config_df = config_df.dropna(subset=['Dest_table', 'Dest_field'])

        # Run processing
        process_file(input_excel_path, config_df, os.getcwd(), output_dir)

        # Find the output subfolder
        subfolders = [d for d in os.listdir(output_dir) if os.path.isdir(os.path.join(output_dir, d))]
        if not subfolders:
            raise FileNotFoundError("No output files generated.")
        target_output = os.path.join(output_dir, subfolders[0])

        # Zip the output
        zip_buffer = BytesIO()
        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for root, _, files in os.walk(target_output):
                for file in files:
                    full_path = os.path.join(root, file)
                    arcname = os.path.basename(full_path)
                    zipf.write(full_path, arcname)
        zip_buffer.seek(0)

        print("✅ Processing complete. Sending ZIP back.")
        return send_file(zip_buffer, download_name='output.zip', as_attachment=True)

    except Exception as e:
        print(f"❌ Internal error: {e}")
        return jsonify({"error": str(e)}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
