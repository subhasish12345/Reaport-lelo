"""
api/index.py — Vercel serverless Flask entry point
All routes live here. Templates are at ../templates/
"""
import io
import os
import sys

# Make report_generator importable (lives one level up from api/)
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

from flask import Flask, render_template, request, send_file, jsonify
from report_generator import generate_report_bytes

# template_folder points up one level so Vercel can find templates/
app = Flask(
    __name__,
    template_folder=os.path.join(ROOT, 'templates'),
)
app.config['MAX_CONTENT_LENGTH'] = 5 * 1024 * 1024   # 5 MB


# --------------------------------------------------------------------------
# Routes
# --------------------------------------------------------------------------

@app.route('/', methods=['GET'])
def index():
    return render_template('index.html')


@app.route('/generate', methods=['POST'])
def generate():
    """Accept paste or file upload → return .docx download."""
    content = None

    # Priority 1: uploaded .txt file
    if 'file' in request.files and request.files['file'].filename:
        raw = request.files['file'].read()
        content = raw.decode('utf-8', errors='replace')

    # Priority 2: pasted textarea content
    elif request.form.get('content', '').strip():
        content = request.form['content']

    if not content or not content.strip():
        return jsonify({'error': 'No content provided. Paste your report text or upload a .txt file.'}), 400

    # Build a safe filename from the project title
    title = request.form.get('project_title', 'Project_Report').strip() or 'Project_Report'
    safe  = ''.join(c if c.isalnum() or c in '-_ ' else '_' for c in title)
    filename = f"{safe.replace(' ', '_')}_Report.docx"

    try:
        docx_bytes = generate_report_bytes(content)
    except Exception as exc:
        return jsonify({'error': f'Generation failed: {exc}'}), 500

    return send_file(
        io.BytesIO(docx_bytes),
        as_attachment=True,
        download_name=filename,
        mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    )


# Vercel calls the WSGI app object directly as `app`
# Running locally: python api/index.py
if __name__ == '__main__':
    app.run(debug=True, port=5050)
