from flask import Flask, request, send_file, jsonify
from werkzeug.utils import secure_filename
from flask_cors import CORS
import os
from services.file_processor import process_file

app = Flask(__name__)
CORS(app)  # Enable CORS on all routes
app.config['UPLOAD_FOLDER'] = 'uploads/'
app.config['PROCESSED_FOLDER'] = 'processed/'
app.config['IMAGE_FOLDER'] = 'images/'

os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['PROCESSED_FOLDER'], exist_ok=True)
os.makedirs(app.config['IMAGE_FOLDER'], exist_ok=True)


@app.route('/')
def index():
    return "Welcome to the Excel Processor API!"


@app.route('/process', methods=['POST'])
def process():
    if 'file' not in request.files:
        return jsonify({"error": "No file part"}), 400
    file = request.files['file']
    if file.filename == '':
        return jsonify({"error": "No selected file"}), 400

    filename = secure_filename(file.filename)
    file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    file.save(file_path)

    try:
        processed_file, _ = process_file(file_path, app.config['PROCESSED_FOLDER'], app.config['IMAGE_FOLDER'])

        # Determine the download name and mimetype
        download_name = os.path.basename(processed_file)
        mimetype = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'  # MIME type for .xlsx

        return send_file(processed_file, as_attachment=True, download_name=download_name, mimetype=mimetype)
    except Exception as e:
        return jsonify({"error": str(e)}), 500


if __name__ == '__main__':
    app.run(debug=True, port=5001)
