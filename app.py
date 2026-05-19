import os
import json
import tempfile
import fitz  # PyMuPDF
from flask import Flask, request, jsonify, send_file, render_template
from werkzeug.utils import secure_filename
import google.generativeai as genai
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 32 * 1024 * 1024  # 32MB

GEMINI_API_KEY = os.environ.get('GEMINI_API_KEY')
if not GEMINI_API_KEY:
    raise RuntimeError("GEMINI_API_KEY environment variable is not set")

genai.configure(api_key=GEMINI_API_KEY)

ALLOWED_EXTENSIONS = {'pdf', 'png', 'jpg', 'jpeg', 'webp', 'gif'}
IMAGE_MIME_TYPES = {
    'png': 'image/png',
    'jpg': 'image/jpeg',
    'jpeg': 'image/jpeg',
    'webp': 'image/webp',
    'gif': 'image/gif',
}

SYSTEM_INSTRUCTIONS = """You are a data extraction expert. Extract ALL data from the provided content without skipping any items.

Return ONLY a valid JSON object with this exact structure:
{
  "title": "descriptive title for the spreadsheet",
  "sheets": [
    {
      "name": "Data",
      "headers": ["Column1", "Column2", "Column3"],
      "rows": [
        ["value1", "value2", "value3"],
        ["value1", "value2", "value3"]
      ]
    }
  ]
}

Rules:
- Extract EVERY SINGLE item — do not truncate, summarize, or stop early. If there are 100 items, return 100 rows.
- If the document has categories or sections (e.g. "Beverages", "Food"), include the category as a column in every row it applies to — do NOT create separate sheets per category.
- Keep product/item names exactly as written — do NOT split them into sub-columns. Example: "IMPERIO PURO MALTE 350ML" is ONE value in one column, not two columns.
- Only create separate columns when the source document itself clearly separates values (e.g. a price column, a code column, a quantity column).
- Preserve numeric values as numbers (no quotes).
- Every row must have the same number of values as headers.
- Always put ALL data into a SINGLE sheet. Never create multiple sheets.
- Put all the headers in Brazilian Portuguese
- Return ONLY the JSON, no markdown, no explanation.
"""

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def get_extension(filename):
    return filename.rsplit('.', 1)[1].lower() if '.' in filename else ''

def pdf_to_images(filepath):
    """Render each PDF page to a PNG and return list of (bytes, mime_type)."""
    doc = fitz.open(filepath)
    pages = []
    for page in doc:
        mat = fitz.Matrix(3, 3)  # 3x zoom for high quality
        pix = page.get_pixmap(matrix=mat)
        pages.append((pix.tobytes("png"), "image/png"))
    doc.close()
    return pages

def _parse_gemini_response(raw):
    raw = raw.strip()
    if raw.startswith("```"):
        raw = raw.split("```")[1]
        if raw.startswith("json"):
            raw = raw[4:]
    return json.loads(raw.strip())

def call_gemini_images(image_list, user_prompt=""):
    """Send one or more images to Gemini vision."""
    model = genai.GenerativeModel('gemini-2.5-flash')
    prompt = SYSTEM_INSTRUCTIONS
    if user_prompt:
        prompt += f"\n\nAdditional instructions: {user_prompt}"
    parts = [prompt]
    for image_bytes, mime_type in image_list:
        parts.append({"mime_type": mime_type, "data": image_bytes})
    response = model.generate_content(parts)
    return _parse_gemini_response(response.text)

def call_gemini_text(content_text, user_prompt=""):
    model = genai.GenerativeModel('gemini-2.5-flash')
    prompt = f"{SYSTEM_INSTRUCTIONS}\n\nContent to extract:\n{content_text}"
    if user_prompt:
        prompt += f"\n\nAdditional instructions: {user_prompt}"
    response = model.generate_content(prompt)
    return _parse_gemini_response(response.text)

def build_excel(data: dict, output_path: str):
    wb = Workbook()
    wb.remove(wb.active)

    header_font = Font(name="Calibri", bold=True, size=11)
    data_font = Font(name="Calibri", size=11)
    header_fill = PatternFill("solid", fgColor="BFBFBF")

    sheets_data = data.get("sheets", [])
    if not sheets_data:
        raise ValueError("No sheets data found")

    for sheet_info in sheets_data:
        ws = wb.create_sheet(title=sheet_info["name"][:31])
        headers = sheet_info.get("headers", [])
        rows = sheet_info.get("rows", [])

        # Header row
        for col_idx, header in enumerate(headers, start=1):
            cell = ws.cell(row=1, column=col_idx, value=header)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center", vertical="center")

        # Data rows
        for row_idx, row_data in enumerate(rows, start=2):
            for col_idx, value in enumerate(row_data, start=1):
                cell = ws.cell(row=row_idx, column=col_idx, value=value)
                cell.font = data_font

        # Auto-fit columns
        for col_idx, header in enumerate(headers, start=1):
            col_letter = get_column_letter(col_idx)
            max_len = len(str(header))
            for row_data in rows:
                if col_idx - 1 < len(row_data):
                    max_len = max(max_len, len(str(row_data[col_idx - 1])))
            ws.column_dimensions[col_letter].width = min(max_len + 2, 50)

    wb.save(output_path)


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/api/convert', methods=['POST'])
def convert():
    user_prompt = request.form.get('prompt', '').strip()
    text_input = request.form.get('text_input', '').strip()

    uploaded_file = request.files.get('file')
    has_file = uploaded_file and uploaded_file.filename

    if not has_file and not text_input:
        return jsonify({'error': 'Please upload a file or paste some text'}), 400

    if has_file and not allowed_file(uploaded_file.filename):
        return jsonify({'error': 'Unsupported file type. Use PDF, PNG, JPG, WEBP, or GIF'}), 400

    try:
        data = None

        if has_file:
            ext = get_extension(uploaded_file.filename)
            with tempfile.NamedTemporaryFile(delete=False, suffix=f".{ext}") as tmp:
                uploaded_file.save(tmp.name)
                tmp_path = tmp.name

            try:
                if ext == 'pdf':
                    # Render PDF pages as images and send to Gemini vision
                    pages = pdf_to_images(tmp_path)
                    extra_prompt = user_prompt
                    if text_input:
                        extra_prompt += f"\n\nAdditional text context:\n{text_input}"
                    data = call_gemini_images(pages, extra_prompt)
                else:
                    with open(tmp_path, 'rb') as f:
                        image_bytes = f.read()
                    mime_type = IMAGE_MIME_TYPES[ext]
                    extra_prompt = user_prompt
                    if text_input:
                        extra_prompt += f"\n\nAdditional text context:\n{text_input}"
                    data = call_gemini_images([(image_bytes, mime_type)], extra_prompt)
            finally:
                os.unlink(tmp_path)

        else:
            data = call_gemini_text(text_input, user_prompt)

    except json.JSONDecodeError:
        return jsonify({'error': 'Gemini returned invalid data. Try again or simplify the content.'}), 500
    except Exception as e:
        return jsonify({'error': f'Processing error: {str(e)}'}), 500

    with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_out:
        out_path = tmp_out.name

    try:
        build_excel(data, out_path)
    except Exception as e:
        os.unlink(out_path)
        return jsonify({'error': f'Excel generation failed: {str(e)}'}), 500

    return send_file(
        out_path,
        as_attachment=True,
        download_name="extracted_data.xlsx",
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )


if __name__ == '__main__':
    app.run(debug=True, port=5000)