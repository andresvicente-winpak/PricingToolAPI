from flask import Flask, request, jsonify
from FixedLoadSheet import process_file, load_excel_file, output_dir, log_path
import pandas as pd
import os
import tempfile
from io import BytesIO
import time

app = Flask(__name__)

@app.route('/process', methods=['POST'])
def process():
    try:
        uploaded_excel = BytesIO(request.get_data())

        work_dir = tempfile.mkdtemp()
        input_dir = os.path.join(work_dir, "input")
        local_output_dir = output_dir
        sample_dir = os.getcwd()  # ← this is the fix

        os.makedirs(input_dir, exist_ok=True)
        os.makedirs(local_output_dir, exist_ok=True)

        # Save uploaded Excel to input folder
        input_excel_path = os.path.join(input_dir, "input.xlsx")
        with open(input_excel_path, 'wb') as f:
            f.write(uploaded_excel.read())

        # Load configuration
        config_path = os.path.join(os.getcwd(), "configuration.xlsx")
        config_df = load_excel_file(config_path, header=0, dtype=str)
        config_df.columns = [str(col).strip() for col in config_df.columns]
        config_df = config_df.dropna(subset=['Dest_table', 'Dest_field'])

        # Sample files come from current working directory
        local_sample_dir = os.getcwd()

        # Process file
        process_file(input_excel_path, config_df, local_sample_dir, local_output_dir)

        # Wait briefly to ensure log file is written
        time.sleep(1)

        # Read and return log file content
        log_content = ""
        try:
            with open(log_path, "r", encoding="utf-8") as f:
                log_content = f.read()
        except Exception as e:
            log_content = f"⚠️ Failed to read log file: {str(e)}"

        return jsonify({
            "status": "success",
            "log": log_content
        })

    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
