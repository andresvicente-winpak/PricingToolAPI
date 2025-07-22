from flask import Flask, request, send_file, jsonify
from FixedLoadSheet import process_file, load_excel_file, output_dir, log_path
import pandas as pd
import os
import tempfile
import zipfile
from io import BytesIO

app = Flask(__name__)

@app.route('/process', methods=['POST'])
def process():
    try:
        uploaded_excel = BytesIO(request.get_data())

        work_dir = tempfile.mkdtemp()
        input_dir = os.path.join(work_dir, "input")
        sample_dir = os.path.join(work_dir, "sample")
        local_output_dir = output_dir

        os.makedirs(input_dir, exist_ok=True)
        os.makedirs(sample_dir, exist_ok=True)
        os.makedirs(local_output_dir, exist_ok=True)

        # Save uploaded Excel to input folder
        input_excel_path = os.path.join(input_dir, "input.xlsx")
        with open(input_excel_path, 'wb') as f:
            f.write(uploaded_excel.read())

        # Load local configuration file
        config_path = os.path.join(os.getcwd(), "configuration.xlsx")
        config_df = load_excel_file(config_path, header=0, dtype=str)
        config_df.columns = [str(col).strip() for col in config_df.columns]
        config_df = config_df.dropna(subset=['Dest_table', 'Dest_field'])

        # Use current working directory for sample CSVs
        local_sample_dir = os.getcwd()

        # Process the uploaded Excel
        process_file(input_excel_path, config_df, local_sample_dir, local_output_dir)

	# Wait briefly to ensure log is written before reading it
	time.sleep(1)
	
	if os.path.exists(log_path):
	    with open(log_path, "r", encoding="utf-8") as log_file:
	        log_content = log_file.read()
	else:
	    log_content = "❌ Log file not found (processing may have silently failed)."
	
	return jsonify({
	    "status": "success",
	    "log": log_content
	})
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
