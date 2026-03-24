#!/usr/bin/env python3
"""
VEROUDERD — Niet meer in gebruik.

Functionaliteit is overgenomen door backend/app.py:
  - POST /api/upload-report  → backend/app.py (opslaan in backend/storage/html/)
  - /api/kb/...              → backend/app.py (volledige KB implementatie)

De denjoy-upload systemd service en de bijbehorende Nginx routing zijn verwijderd
uit deploy/install.sh. Alle verkeer gaat nu naar app.py op poort 8787.
"""
import os
from flask import Flask, request, jsonify, send_from_directory, abort
from flask_cors import CORS
from kb_api import kb

ROOT = os.path.dirname(os.path.abspath(__file__))
WEB_DIR = os.path.join(ROOT, 'web')
HTML_DIR = os.path.join(WEB_DIR, 'html')

os.makedirs(HTML_DIR, exist_ok=True)

app = Flask(__name__, static_folder=WEB_DIR, static_url_path='')
app.config['MAX_CONTENT_LENGTH'] = int(os.getenv('UPLOAD_MAX_BYTES', 10 * 1024 * 1024))

default_origins = ['http://localhost:8080', 'http://127.0.0.1:8080']
allowed_origins = [
    origin.strip() for origin in os.getenv('UPLOAD_ALLOWED_ORIGINS', ','.join(default_origins)).split(',')
    if origin.strip()
]
CORS(app, resources={r"/*": {"origins": allowed_origins}})

app.register_blueprint(kb)


@app.route('/upload-report', methods=['POST'])
def upload_report():
    try:
        if not request.is_json:
            return jsonify({'error': 'Expected JSON'}), 400

        data = request.get_json()
        filename = data.get('filename') or 'M365-Complete-Baseline-latest.html'
        content = data.get('content')
        if not content:
            return jsonify({'error': 'No content provided'}), 400
        if not isinstance(content, str):
            return jsonify({'error': 'Content must be a string'}), 400

        # sanitize filename
        filename = os.path.basename(filename)
        if not filename.lower().endswith('.html'):
            return jsonify({'error': 'Only .html files are allowed'}), 400
        target_path = os.path.join(HTML_DIR, filename)

        with open(target_path, 'w', encoding='utf-8') as f:
            f.write(content)

        rel_path = f'/web/html/{filename}'
        return jsonify({'path': rel_path, 'filename': filename}), 200
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/web/<path:filename>')
def serve_web_file(filename):
    # Serve files from web/ directory
    file_path = os.path.join(WEB_DIR, filename)
    if not os.path.exists(file_path):
        abort(404)
    return send_from_directory(WEB_DIR, filename)


@app.route('/')
def index():
    return send_from_directory(WEB_DIR, 'index.html')


if __name__ == '__main__':
    import argparse

    parser = argparse.ArgumentParser()
    parser.add_argument('--host', default='127.0.0.1')
    parser.add_argument('--port', default=8080, type=int)
    args = parser.parse_args()

    print(f'Serving web UI from {WEB_DIR} on http://{args.host}:{args.port}')
    app.run(host=args.host, port=args.port)
