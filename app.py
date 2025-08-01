from flask import Flask, request, jsonify, send_file
import os
import tempfile
from process_handler import process_file
import base64
from datetime import datetime

app = Flask(__name__)

@app.route("/", methods=["GET"])
def home():
    return "Pricing Tool API is live"

@app.route("/process", methods=["POST"])
def process():
    try:
        # Parse base64-encoded content from JSON
        data = request.get_json()
        if not data or "$content" not in data:
            return jsonify({"error": "Missing base64 content."}), 400

        # Get filename from header or use default
        filename = request.headers.get("X-Filename", "uploaded.xlsx")

        # Save decoded file to temporary path
        temp_dir = tempfile.mkdtemp()
        input_path = os.path.join(temp_dir, filename)
        with open(input_path, "wb") as f:
            f.write(base64.b64decode(data["$content"]))

        try:
            # Try processing file
            zip_path = process_file(input_path)
            return send_file(zip_path, as_attachment=True)

        except Exception as e:
            # Write error to file and return it
            error_log_path = os.path.join(temp_dir, f"error-log.txt")
            with open(error_log_path, "w") as errfile:
                errfile.write("\u274C FAILED\n")
                errfile.write(f"Timestamp: {datetime.utcnow().isoformat()}Z\n")
                errfile.write(f"Filename: {filename}\n")
                errfile.write(f"Error Message: {str(e)}\n")
           
            # Zip the error file
            zip_path = os.path.join(temp_dir, f"error-log-{datetime.utcnow().strftime('%Y%m%d-%H%M%S')}.zip")
            with zipfile.ZipFile(zip_path, 'w') as zipf:
                zipf.write(error_log_path, arcname="error-log.txt")
            
            return send_file(zip_path, as_attachment=True, mimetype='application/zip'), 200

    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == "__main__":
    print("\u26a0\ufe0f Entered main block.")
    app.run(debug=True, port=10000, host="0.0.0.0")
