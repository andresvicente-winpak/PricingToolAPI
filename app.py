from flask import Flask, request, send_file, jsonify
from FixedLoadSheet import process_file, load_excel_file
import pandas as pd
import os
import tempfile
import zipfile
from io import BytesIO
import base64
import json

app = Flask(__name__)

@app.route('/process', methods=['POST'])
def process():
    try:
        print("📥 Receiving file...")

        content_type = request.headers.get("Content-Type", "").lower()
        print(f"📄 Content-Type: {content_type}")

        if "application/json" in content_type or "application/octet-stream" in content_type:
            # Handle Power Automate base64 JSON format
            try:
                data = request.get_json(force=True)
                base64_data = data.get('contentBytes')

                if not base64_data:
                    raise ValueError("Missing 'contentBytes' in request body.")

                raw_bytes = base64.b64decode(base64_data)
                uploaded_excel = BytesIO(raw_bytes)
                print("✅ Base64 decoded from Power Automate JSON body.")

            except Exception as json_err:
                raise ValueError(f"Failed to decode JSON with contentBytes: {json_err}")
        else:
            # Handle raw binary just in case
            uploaded_excel = BytesIO(request.get_data())
            print("⚠️ Assuming raw binary payload (not JSON).")

        # Save debug file
        debug_path = os.path.join(os.getcwd(), "uploaded_raw_preview.xlsx")
        with open(debug_path, "wb") as debug_file:
            debug_file.write(uploaded_excel.getvalue())
        print(f"🧪 Debug file written: {debug_path}")
        print(f"📏 Raw data size: {len(uploaded_excel.getvalue())} bytes")

        # Set up temp working dirs
        work_dir = tempfile.mkdtemp()
        input_dir = os.path.join(work_dir, "input")
        output_dir = os.path.join(work_dir, "output")
        os.makedirs(input_dir, exist_ok=True)
        os.makedirs(output_dir, exist_ok=True)

        # Save uploaded file to temp
        input_excel_path = os.path.join(input_dir, "input.xlsx")
        with open(input_excel_path, "wb") as f:
            f.write(uploaded_excel.read())
        print(f"📁 Excel saved to: {input_excel_path}")

        # Load config
        config_path = os.path.join(os.getcwd(), "configuration.xlsx")
        config_df = load_excel_file(config_path, header=0, dtype=str)
        config_df.columns = [str(col).strip() for col in config_df.columns]
        config_df = config_df.dropna(subset=['Dest_table', 'Dest_field'])

        # Run processing
        process_file(input_excel_path, config_df, os.getcwd(), output_dir)

        # Locate output folder
        subfolders = [d for d in os.listdir(output_dir) if os.path.isdir(os.path.join(output_dir, d))]
        if not subfolders:
            raise FileNotFoundError("No output files generated.")
        final_output = os.path.join(output_dir, subfolders[0])

        # Zip output
        zip_buffer = BytesIO()
        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for root, _, files in os.walk(final_output):
                for file in files:
                    full_path = os.path.join(root, file)
                    arcname = os.path.basename(full_path)
                    zipf.write(full_path, arcname)
        zip_buffer.seek(0)

        print("✅ ZIP created and sent.")
        return send_file(zip_buffer, download_name='output.zip', as_attachment=True)

    except Exception as e:
        print(f"❌ Internal error: {e}")
        return jsonify({"error": str(e)}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
