from flask import Flask, request, jsonify, send_file
import os
import tempfile
from process_handler import process_file
import base64

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

        # Process file
        try:
            zip_path = process_file(input_path)
        except Exception as e:
            return jsonify({"error": str(e)}), 500

        # Return zip as response
        return send_file(zip_path, as_attachment=True)

    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == "__main__":
    print("⚠️ Entered main block.")
    app.run(debug=True, port=10000, host="0.0.0.0")
