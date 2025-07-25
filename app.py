from flask import Flask, request, send_file, jsonify
import os
import pandas as pd
import tempfile
import base64
from werkzeug.utils import secure_filename
from process_handler import process_file

app = Flask(__name__)

@app.route('/')
def index():
    return "Pricing Tool API is running."

@app.route('/process', methods=['POST'])
def process():
    content_type = request.headers.get('Content-Type', '')
    print(f"\U0001F4C2 Content-Type: {content_type}")

    # Try to parse as JSON with contentBytes and name
    if content_type == 'application/json':
        try:
            data = request.get_json()
            file_name = data.get('name', 'input.xlsx')
            file_bytes = base64.b64decode(data['contentBytes'])
            print("\u2705 Base64 decoded from Power Automate JSON body.")
        except Exception as e:
            return jsonify({"error": f"Failed to decode JSON with contentBytes: {e}"}), 400

    # Try to parse as raw binary upload (e.g. octet-stream)
    elif content_type == 'application/octet-stream':
        try:
            file_bytes = request.get_data()
            file_name = request.headers.get('X-Filename', 'input.xlsx')
            print("⚠️ Assuming raw binary payload (not JSON).")
        except Exception as e:
            return jsonify({"error": f"Failed to read binary data: {e}"}), 400

    # Unsupported type
    else:
        return jsonify({"error": f"Unsupported Content-Type: {content_type}"}), 415

    # Save input file
    try:
        input_dir = tempfile.mkdtemp()
        input_path = os.path.join(input_dir, secure_filename(file_name))
        with open(input_path, 'wb') as f:
            f.write(file_bytes)
        print(f"📁 Excel saved to: {input_path}")
    except Exception as e:
        return jsonify({"error": f"Failed to write file: {e}"}), 500

    # Process and return ZIP
    try:
        output_zip = process_file(input_path)
        print(f"📦 ZIP created and sent.")
        return send_file(
            output_zip,
            mimetype='application/zip',
            as_attachment=True,
            download_name='output.zip'
        )
    except Exception as e:
        return jsonify({"error": f"Failed to process file: {e}"}), 500

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)
