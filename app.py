import os
import json
import uuid
import threading
from pathlib import Path
from flask import Flask, render_template, request, jsonify, send_file
import pdfplumber
from groq import Groq
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import werkzeug.utils

app = Flask(__name__)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_FOLDER = '/tmp/uploads'
OUTPUT_FOLDER = '/tmp/outputs'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['OUTPUT_FOLDER'] = OUTPUT_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB

jobs = {}

# ─────────────────────────────────────────────────────────────────────────────
# SPATIAL EXTRACTION
# We use pdfplumber's word-level bounding boxes to understand WHERE on the page
# each piece of text sits. Words are then clustered into "lines" by their
# Y-position, and lines are grouped into "blocks" by vertical proximity.
# ─────────────────────────────────────────────────────────────────────────────

def extract_spatial_pages(pdf_path):
    """
    Returns a list of pages. Each page has:
      - page (int)
      - width, height (floats, PDF points)
      - blocks: list of text blocks, each with:
          - lines: list of { text, x0, y0, x1, y1 }
          - bbox: bounding box of whole block
      - tables: list of 2D arrays (raw table data) with position info
      - raw_text: plain text fallback
    """
    pages = []
    with pdfplumber.open(pdf_path) as pdf:
        for page_idx, page in enumerate(pdf.pages):
            pw = float(page.width)
            ph = float(page.height)

            # ── Extract tables first so we can exclude those regions from word extraction
            raw_tables = page.extract_tables() or []
            table_bboxes = []
            table_objects = []
            for t_obj in (page.find_tables() or []):
                try:
                    tb = t_obj.bbox  # (x0, top, x1, bottom)
                    table_bboxes.append(tb)
                    table_objects.append({
                        "bbox": tb,
                        "data": t_obj.extract()
                    })
                except Exception:
                    pass

            # ── Extract words with bounding boxes
            words = page.extract_words(
                x_tolerance=3,
                y_tolerance=3,
                keep_blank_chars=False,
                use_text_flow=False
            ) or []

            # Filter out words inside table bboxes
            def in_table(w):
                for tb in table_bboxes:
                    if (w['x0'] >= tb[0] - 2 and w['x1'] <= tb[2] + 2 and
                            w['top'] >= tb[1] - 2 and w['bottom'] <= tb[3] + 2):
                        return True
                return False

            text_words = [w for w in words if not in_table(w)]

            # ── Cluster words into lines by Y-position (top coordinate)
            LINE_Y_TOLERANCE = 4  # words within 4pt vertically = same line
            lines = []
            for w in sorted(text_words, key=lambda x: (round(x['top'] / LINE_Y_TOLERANCE), x['x0'])):
                placed = False
                for line in lines:
                    if abs(line['y'] - w['top']) <= LINE_Y_TOLERANCE:
                        line['words'].append(w)
                        line['x1'] = max(line['x1'], w['x1'])
                        line['y1'] = max(line['y1'], w['bottom'])
                        placed = True
                        break
                if not placed:
                    lines.append({
                        'y': w['top'],
                        'y1': w['bottom'],
                        'x0': w['x0'],
                        'x1': w['x1'],
                        'words': [w]
                    })

            # ── Build line dicts with assembled text
            line_dicts = []
            for line in sorted(lines, key=lambda l: l['y']):
                sorted_words = sorted(line['words'], key=lambda w: w['x0'])
                text = ' '.join(w['text'] for w in sorted_words)
                line_dicts.append({
                    'text': text,
                    'x0': sorted_words[0]['x0'],
                    'y0': line['y'],
                    'x1': line['x1'],
                    'y1': line['y1'],
                })

            # ── Cluster lines into blocks by vertical proximity
            BLOCK_GAP = 14  # lines more than 14pt apart = new block
            blocks = []
            for line in line_dicts:
                if not blocks or (line['y0'] - blocks[-1]['lines'][-1]['y1']) > BLOCK_GAP:
                    blocks.append({
                        'lines': [line],
                        'x0': line['x0'], 'y0': line['y0'],
                        'x1': line['x1'], 'y1': line['y1'],
                    })
                else:
                    b = blocks[-1]
                    b['lines'].append(line)
                    b['x0'] = min(b['x0'], line['x0'])
                    b['x1'] = max(b['x1'], line['x1'])
                    b['y1'] = line['y1']

            raw_text = page.extract_text() or ""

            pages.append({
                "page": page_idx + 1,
                "width": pw,
                "height": ph,
                "blocks": blocks,
                "tables": table_objects,
                "raw_text": raw_text.strip(),
            })

    return pages


# ─────────────────────────────────────────────────────────────────────────────
# GROQ AI ANALYSIS  (unchanged logic, uses raw_text from spatial pages)
# ─────────────────────────────────────────────────────────────────────────────

def categorize_with_groq(spatial_pages, api_key, filename, job_id):
    client = Groq(api_key=api_key)
    jobs[job_id]["status"] = "analyzing"
    jobs[job_id]["progress"] = 40

    full_text_parts = []
    for p in spatial_pages:
        chunk = f"[Page {p['page']}]\n{p['raw_text']}"
        full_text_parts.append(chunk)

    full_text = "\n\n".join(full_text_parts)
    if len(full_text) > 12000:
        full_text = full_text[:12000] + "\n...[truncated]"

    prompt = f"""Analyze this PDF document and extract structured data. Return ONLY valid JSON, no markdown.

Document: "{filename}"
Content:
{full_text}

Return exactly:
{{
  "document_summary": "2-3 sentence summary",
  "document_type": "Report/Invoice/Contract/Academic Paper/Manual/Letter/Resume/Presentation/Financial/Legal/Technical/Other",
  "main_topics": ["topic1", "topic2", "topic3"],
  "key_entities": [
    {{"name": "entity name", "type": "Person/Organization/Location/Date/Amount/Product", "context": "brief context"}}
  ],
  "key_facts": [
    {{"fact": "specific fact", "category": "Financial/Technical/Legal/General/Date/Contact", "page": 1}}
  ],
  "action_items": ["action if any"],
  "sentiment": "Positive/Neutral/Negative/Mixed",
  "language": "detected language",
  "estimated_date": "date if found or Unknown"
}}

Extract at least 5 key_entities and 8 key_facts if present."""

    response = client.chat.completions.create(
        model="llama-3.3-70b-versatile",
        messages=[{"role": "user", "content": prompt}],
        temperature=0.1,
        max_tokens=2000
    )

    raw = response.choices[0].message.content.strip()
    if raw.startswith("```"):
        raw = raw.split("```")[1]
        if raw.startswith("json"):
            raw = raw[4:]
    raw = raw.strip().rstrip("```").strip()

    result = json.loads(raw)
    jobs[job_id]["status"] = "building_excel"
    jobs[job_id]["progress"] = 70
    return result


# ─────────────────────────────────────────────────────────────────────────────
# EXCEL BUILDER — spatial layout approach
#
# Strategy per page:
#   • Map PDF X → Excel column, PDF Y → Excel row using a grid resolution
#   • Each page gets its own sheet
#   • Tables are written as proper Excel tables in their correct position
#   • Text blocks are written with merged cells reflecting their width
#   • A final "AI Analysis" sheet is appended at the end
# ─────────────────────────────────────────────────────────────────────────────

def build_spatial_excel(filename, spatial_pages, analysis, output_path):
    wb = Workbook()
    # Remove default sheet — we'll add pages ourselves
    wb.remove(wb.active)

    # ── Colors
    DARK     = "1A1A2E"
    ACCENT   = "E94560"
    HEADER   = "0F3460"
    MID      = "16213E"
    WHITE    = "FFFFFF"
    ALT      = "F0F4FF"
    NORM     = "FFFFFF"

    thin = Border(
        left=Side(style='thin', color='DDDDDD'),
        right=Side(style='thin', color='DDDDDD'),
        top=Side(style='thin', color='DDDDDD'),
        bottom=Side(style='thin', color='DDDDDD'),
    )
    thick_border = Border(
        left=Side(style='medium', color='888888'),
        right=Side(style='medium', color='888888'),
        top=Side(style='medium', color='888888'),
        bottom=Side(style='medium', color='888888'),
    )

    # Grid resolution: how many Excel rows/cols per PDF point
    # PDF pages are typically 612x792 pts (letter) or 595x842 (A4)
    # We map to ~100 cols x 150 rows per page for readability
    COL_SCALE = 0.13   # pts → col units
    ROW_SCALE = 0.19   # pts → row units
    COL_WIDTH  = 2.5   # Excel column width (narrow cols for positioning)
    ROW_HEIGHT = 12    # Excel row height

    def pdf_x_to_col(x, page_width):
        return max(1, int(x * COL_SCALE) + 1)

    def pdf_y_to_row(y, page_height):
        return max(1, int(y * ROW_SCALE) + 1)

    for page_data in spatial_pages:
        pnum = page_data["page"]
        pw   = page_data["width"]
        ph   = page_data["height"]
        sheet_name = f"Page {pnum}"
        ws = wb.create_sheet(title=sheet_name)
        ws.sheet_view.showGridLines = True

        # Set uniform narrow column widths for spatial accuracy
        max_col = max(2, int(pw * COL_SCALE) + 5)
        max_row = max(2, int(ph * ROW_SCALE) + 5)
        for c in range(1, max_col + 1):
            ws.column_dimensions[chr(64 + c) if c <= 26 else
                                 chr(64 + (c-1)//26) + chr(64 + (c-1)%26 + 1)].width = COL_WIDTH
        for r in range(1, max_row + 1):
            ws.row_dimensions[r].height = ROW_HEIGHT

        # Page header banner (row 1)
        ws.row_dimensions[1].height = 22
        end_col_letter = (chr(64 + min(max_col,26)))
        ws.merge_cells(f"A1:{end_col_letter}1")
        hdr_cell = ws.cell(1, 1, f"  PAGE {pnum}  —  {filename}")
        hdr_cell.font = Font(bold=True, size=10, color=WHITE, name="Calibri")
        hdr_cell.fill = PatternFill("solid", start_color=DARK)
        hdr_cell.alignment = Alignment(horizontal="left", vertical="center")

        # ── Write text blocks
        for block in page_data["blocks"]:
            for line in block["lines"]:
                col = pdf_x_to_col(line["x0"], pw)
                row = pdf_y_to_row(line["y0"], ph) + 2  # +2: skip header row + buffer

                # Estimate how many cols this line spans
                col_end = pdf_x_to_col(line["x1"], pw)
                if col_end > col:
                    try:
                        # Only merge if all cells in range are empty
                        can_merge = True
                        for mc in range(col, col_end + 1):
                            if ws.cell(row, mc).value is not None:
                                can_merge = False
                                break
                        if can_merge and col_end > col:
                            ws.merge_cells(
                                start_row=row, start_column=col,
                                end_row=row, end_column=col_end
                            )
                    except Exception:
                        pass

                cell = ws.cell(row, col, line["text"])
                cell.font = Font(size=9, name="Calibri", color="111111")
                cell.alignment = Alignment(
                    horizontal="left", vertical="center", wrap_text=False
                )

        # ── Write tables in their original position
        for tbl in page_data["tables"]:
            if not tbl["data"]:
                continue
            tbbox = tbl["bbox"]  # (x0, top, x1, bottom)
            start_col = pdf_x_to_col(tbbox[0], pw)
            start_row = pdf_y_to_row(tbbox[1], ph) + 2

            for r_idx, row_data in enumerate(tbl["data"]):
                for c_idx, cell_val in enumerate(row_data):
                    ecol = start_col + c_idx
                    erow = start_row + r_idx
                    cell = ws.cell(erow, ecol, str(cell_val or "").strip())
                    is_header = (r_idx == 0)
                    cell.font = Font(
                        size=9, bold=is_header, name="Calibri",
                        color=WHITE if is_header else "111111"
                    )
                    cell.fill = PatternFill("solid",
                        start_color=HEADER if is_header else (ALT if r_idx % 2 == 0 else NORM)
                    )
                    cell.alignment = Alignment(
                        horizontal="center" if is_header else "left",
                        vertical="center", wrap_text=True
                    )
                    cell.border = thin

    # ── AI Analysis sheet ────────────────────────────────────────────────────
    ws_ai = wb.create_sheet(title="AI Analysis")
    ws_ai.sheet_view.showGridLines = False

    def hdr(ws, row, col, value, bg=HEADER, fg=WHITE, bold=True, size=10):
        c = ws.cell(row, col, value)
        c.font = Font(bold=bold, size=size, color=fg, name="Calibri")
        c.fill = PatternFill("solid", start_color=bg)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        return c

    def val(ws, row, col, value, bg=NORM, fg="111111", bold=False, size=10, wrap=True, align="left"):
        c = ws.cell(row, col, value)
        c.font = Font(bold=bold, size=size, color=fg, name="Calibri")
        c.fill = PatternFill("solid", start_color=bg)
        c.alignment = Alignment(horizontal=align, vertical="center", wrap_text=wrap)
        return c

    ws_ai.column_dimensions["A"].width = 22
    ws_ai.column_dimensions["B"].width = 70

    # Title
    ws_ai.merge_cells("A1:B1")
    ws_ai.row_dimensions[1].height = 36
    t = ws_ai.cell(1, 1, "AI ANALYSIS REPORT")
    t.font = Font(bold=True, size=16, color=WHITE, name="Calibri")
    t.fill = PatternFill("solid", start_color=DARK)
    t.alignment = Alignment(horizontal="center", vertical="center")

    ws_ai.merge_cells("A2:B2")
    ws_ai.row_dimensions[2].height = 20
    s = ws_ai.cell(2, 1, f"Source: {filename}")
    s.font = Font(size=9, color="AAAAAA", name="Calibri")
    s.fill = PatternFill("solid", start_color=MID)
    s.alignment = Alignment(horizontal="center", vertical="center")

    r = 4
    summary_fields = [
        ("Document Type",  analysis.get("document_type", "Unknown")),
        ("Language",       analysis.get("language", "Unknown")),
        ("Sentiment",      analysis.get("sentiment", "Unknown")),
        ("Estimated Date", analysis.get("estimated_date", "Unknown")),
        ("Summary",        analysis.get("document_summary", "")),
    ]
    for label, value in summary_fields:
        ws_ai.row_dimensions[r].height = 36 if label == "Summary" else 20
        hdr(ws_ai, r, 1, label, bg=ACCENT)
        vc = val(ws_ai, r, 2, str(value), bg=ALT if r%2==0 else NORM)
        vc.border = thin
        r += 1

    # Topics
    r += 1
    ws_ai.merge_cells(f"A{r}:B{r}")
    hdr(ws_ai, r, 1, "MAIN TOPICS", bg=DARK, size=11)
    r += 1
    for i, topic in enumerate(analysis.get("main_topics", [])):
        ws_ai.row_dimensions[r].height = 18
        ws_ai.merge_cells(f"A{r}:B{r}")
        val(ws_ai, r, 1, f"  •  {topic}", bg=ALT if i%2==0 else NORM)
        r += 1

    # Entities
    r += 1
    ws_ai.merge_cells(f"A{r}:B{r}")
    hdr(ws_ai, r, 1, "KEY ENTITIES", bg=DARK, size=11)
    r += 1
    ws_ai.column_dimensions["A"].width = 28
    ws_ai.column_dimensions["B"].width = 20

    # Add a third col for context
    ws_ai.column_dimensions["C"].width = 50
    hdr(ws_ai, r, 1, "Name", bg=ACCENT)
    hdr(ws_ai, r, 2, "Type", bg=ACCENT)
    hdr(ws_ai, r, 3, "Context", bg=ACCENT)
    r += 1
    for i, ent in enumerate(analysis.get("key_entities", [])):
        bg = ALT if i%2==0 else NORM
        ws_ai.row_dimensions[r].height = 18
        val(ws_ai, r, 1, ent.get("name",""), bg=bg, bold=True)
        val(ws_ai, r, 2, ent.get("type",""), bg=bg, align="center")
        val(ws_ai, r, 3, ent.get("context",""), bg=bg)
        for c in range(1, 4):
            ws_ai.cell(r, c).border = thin
        r += 1

    # Key facts
    r += 1
    hdr(ws_ai, r, 1, "Fact", bg=DARK, size=11)
    hdr(ws_ai, r, 2, "Category", bg=DARK, size=11)
    hdr(ws_ai, r, 3, "Page", bg=DARK, size=11)
    r += 1
    for i, fact in enumerate(analysis.get("key_facts", [])):
        bg = ALT if i%2==0 else NORM
        ws_ai.row_dimensions[r].height = 24
        val(ws_ai, r, 1, fact.get("fact",""), bg=bg, wrap=True)
        val(ws_ai, r, 2, fact.get("category",""), bg=bg, align="center")
        val(ws_ai, r, 3, fact.get("page",""), bg=bg, align="center")
        for c in range(1, 4):
            ws_ai.cell(r, c).border = thin
        r += 1

    # Action items
    action_items = analysis.get("action_items", [])
    if action_items:
        r += 1
        ws_ai.merge_cells(f"A{r}:C{r}")
        hdr(ws_ai, r, 1, "ACTION ITEMS", bg=DARK, size=11)
        r += 1
        for i, item in enumerate(action_items):
            ws_ai.row_dimensions[r].height = 20
            ws_ai.merge_cells(f"A{r}:C{r}")
            val(ws_ai, r, 1, f"  >  {item}", bg=ALT if i%2==0 else NORM, wrap=True)
            r += 1

    wb.save(output_path)


# ─────────────────────────────────────────────────────────────────────────────
# JOB PROCESSOR
# ─────────────────────────────────────────────────────────────────────────────

def process_job(job_id, pdf_path, filename, api_key):
    try:
        jobs[job_id]["status"] = "extracting"
        jobs[job_id]["progress"] = 10

        spatial_pages = extract_spatial_pages(pdf_path)
        jobs[job_id]["pages"] = len(spatial_pages)
        jobs[job_id]["progress"] = 30

        analysis = categorize_with_groq(spatial_pages, api_key, filename, job_id)

        output_filename = f"{Path(filename).stem}_spatial_{job_id[:8]}.xlsx"
        output_path = os.path.join(app.config['OUTPUT_FOLDER'], output_filename)
        build_spatial_excel(filename, spatial_pages, analysis, output_path)

        # Prepare page data for web view
        # Serialize spatial_pages for JSON (strip heavy data, keep structure)
        pages_for_web = []
        for p in spatial_pages:
            pages_for_web.append({
                "page":   p["page"],
                "width":  p["width"],
                "height": p["height"],
                "blocks": [
                    {
                        "lines": [
                            {"text": l["text"], "x0": l["x0"], "y0": l["y0"],
                             "x1": l["x1"], "y1": l["y1"]}
                            for l in b["lines"]
                        ],
                        "x0": b["x0"], "y0": b["y0"],
                        "x1": b["x1"], "y1": b["y1"],
                    }
                    for b in p["blocks"]
                ],
                "tables": [
                    {"bbox": t["bbox"], "data": t["data"]}
                    for t in p["tables"]
                ],
            })

        jobs[job_id]["status"] = "complete"
        jobs[job_id]["progress"] = 100
        jobs[job_id]["output_file"] = output_filename
        jobs[job_id]["analysis"] = analysis
        jobs[job_id]["spatial_pages"] = pages_for_web

    except Exception as e:
        import traceback
        jobs[job_id]["status"] = "error"
        jobs[job_id]["error"] = str(e) + "\n" + traceback.format_exc()
    finally:
        if os.path.exists(pdf_path):
            os.remove(pdf_path)


# ─────────────────────────────────────────────────────────────────────────────
# FLASK ROUTES
# ─────────────────────────────────────────────────────────────────────────────

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/upload", methods=["POST"])
def upload():
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400
    f = request.files["file"]
    api_key = request.form.get("api_key", "").strip()
    if not api_key:
        return jsonify({"error": "Groq API key is required"}), 400
    if not f.filename.lower().endswith(".pdf"):
        return jsonify({"error": "Only PDF files are supported"}), 400

    job_id = str(uuid.uuid4())
    safe_name = werkzeug.utils.secure_filename(f.filename)
    pdf_path = os.path.join(app.config['UPLOAD_FOLDER'], f"{job_id}_{safe_name}")
    f.save(pdf_path)

    jobs[job_id] = {"status": "queued", "progress": 0, "filename": safe_name, "pages": 0}

    thread = threading.Thread(target=process_job, args=(job_id, pdf_path, safe_name, api_key))
    thread.daemon = True
    thread.start()

    return jsonify({"job_id": job_id})

@app.route("/status/<job_id>")
def status(job_id):
    job = jobs.get(job_id)
    if not job:
        return jsonify({"error": "Job not found"}), 404
    return jsonify(job)

@app.route("/download/<job_id>")
def download(job_id):
    job = jobs.get(job_id)
    if not job or job.get("status") != "complete":
        return jsonify({"error": "Not ready"}), 404
    path = os.path.join(app.config['OUTPUT_FOLDER'], job["output_file"])
    return send_file(path, as_attachment=True, download_name=job["output_file"])

if __name__ == "__main__":
    app.run(debug=False, port=5050)
