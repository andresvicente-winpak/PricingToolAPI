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
        print("\n📥 Receiving file...")
        uploaded_excel = BytesIO(request.get_data())
        print("✅ File received, loading Excel...")

        work_dir = tempfile.mkdtemp()
        input_dir = os.path.join(work_dir, "input")
        sample_dir = os.getcwd()  # Use project root as sample folder
        output_dir = os.path.join(work_dir, "output")

        os.makedirs(input_dir, exist_ok=True)
        os.makedirs(output_dir, exist_ok=True)

        input_excel_path = os.path.join(input_dir, "input.xlsx")
        with open(input_excel_path, 'wb') as f:
            f.write(uploaded_excel.read())
        print("✅ Excel saved to:", input_excel_path)

        config_path = os.path.join(os.getcwd(), "configuration.xlsx")
        print(f"🔍 Loading config from {config_path}")
        if not os.path.exists(config_path):
            raise FileNotFoundError("❌ configuration.xlsx not found")

        try:
            config_df = load_excel_file(config_path, header=0, dtype=str)
            print("✅ Config loaded successfully")
        except Exception as e:
            print("❌ Failed to load config:", str(e))
            raise

        try:
            process_file(input_excel_path, config_df, sample_dir, output_dir)
            print("✅ File processed successfully")
        except Exception as e:
            print("❌ Failed during processing:", str(e))
            raise

        subfolders = [d for d in os.listdir(output_dir) if os.path.isdir(os.path.join(output_dir, d))]
        if not subfolders:
            raise FileNotFoundError("No output files generated.")
        target_output = os.path.join(output_dir, subfolders[0])

        zip_buffer = BytesIO()
        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for root, _, files in os.walk(target_output):
                for file in files:
                    full_path = os.path.join(root, file)
                    arcname = os.path.basename(full_path)
                    zipf.write(full_path, arcname)
        zip_buffer.seek(0)

        print("📦 Returning zip output")
        return send_file(zip_buffer, download_name='output.zip', as_attachment=True)

    except Exception as e:
        print("🔥 Internal error:", str(e))
        return jsonify({"error": str(e)}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
