import base64
from flask import Flask, request, send_file
from process_handler import process_file

app = Flask(__name__)

@app.route("/")
def index():
    return "Pricing API running..."

@app.route("/process", methods=["POST"])
def process():
    try:
        # Power Automate sends content in JSON format with base64 content
        if request.is_json:
            payload = request.get_json()
            b64_content = payload.get("$content")
            filename = request.headers.get("X-Filename", "uploaded.xlsx")

            if not b64_content:
                return {"error": "No $content found"}, 400

            file_bytes = base64.b64decode(b64_content)

        else:
            return {"error": "Expected JSON input with base64 content"}, 415

        # Save to temporary file
        import tempfile, os
        temp_dir = tempfile.mkdtemp()
        input_path = os.path.join(temp_dir, filename)
        with open(input_path, "wb") as f:
            f.write(file_bytes)

        # Process
        output_path = process_file(input_path)

        return send_file(output_path, as_attachment=True)

    except Exception as e:
        return {"error": str(e)}, 500
