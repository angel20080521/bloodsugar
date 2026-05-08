import os
import io
import copy
import uuid
import logging
from datetime import datetime, timedelta

import openpyxl
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
from flask import (
    Flask,
    render_template,
    request,
    send_file,
    jsonify,
    flash,
    redirect,
    url_for,
)
import lxml.etree as etree

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = Flask(__name__)
_secret_key = os.environ.get("SECRET_KEY")
if not _secret_key:
    import secrets
    _secret_key = secrets.token_hex(32)
    logger.warning(
        "SECRET_KEY environment variable is not set. "
        "A random key has been generated for this session. "
        "Set SECRET_KEY to a fixed value in production."
    )
app.secret_key = _secret_key

UPLOAD_FOLDER = os.path.join(os.path.dirname(__file__), "uploads")
OUTPUT_FOLDER = os.path.join(os.path.dirname(__file__), "outputs")
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

ALLOWED_EXCEL = {"xlsx", "xls"}
WORD_TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), "血糖数据-01.docx")

_DATETIME_FORMATS = (
    "%Y-%m-%d %H:%M:%S",
    "%Y-%m-%d %H:%M",
    "%Y/%m/%d %H:%M",
    "%Y/%m/%d %H:%M:%S",
)


def allowed_file(filename, allowed):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in allowed


def get_word_template_path():
    if not os.path.isfile(WORD_TEMPLATE_PATH):
        raise FileNotFoundError(f"Word 模板文件未找到：{WORD_TEMPLATE_PATH}")
    return WORD_TEMPLATE_PATH


# ---------------------------------------------------------------------------
# Excel parsing
# ---------------------------------------------------------------------------

def parse_excel(filepath):
    """
    Parse the blood-sugar Excel file.

    Returns a dict:
        {
            "datetime_col": str,   # name of the datetime column
            "value_col": str,      # name of the blood-sugar value column
            "records": [           # list of (datetime, float_value) tuples
                (datetime_obj, value),
                ...
            ]
        }
    """
    wb = openpyxl.load_workbook(filepath, data_only=True)
    ws = wb.active

    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        raise ValueError("Excel 文件为空")

    headers = [str(h).strip() if h is not None else "" for h in rows[0]]

    # Try to auto-detect the datetime column and value column
    datetime_col_idx = None
    value_col_idx = None
    for i, h in enumerate(headers):
        if "时间" in h or "日期" in h or "date" in h.lower() or "time" in h.lower():
            datetime_col_idx = i
        if "血糖" in h or "mmol" in h.lower() or "glucose" in h.lower() or "值" in h:
            value_col_idx = i

    if datetime_col_idx is None:
        datetime_col_idx = 0
    if value_col_idx is None:
        value_col_idx = 1

    records = []
    for row in rows[1:]:
        raw_dt = row[datetime_col_idx]
        raw_val = row[value_col_idx]
        if raw_dt is None or raw_val is None:
            continue
        # Parse datetime
        if isinstance(raw_dt, datetime):
            dt = raw_dt
        else:
            dt_str = str(raw_dt).strip()
            for fmt in _DATETIME_FORMATS:
                try:
                    dt = datetime.strptime(dt_str, fmt)
                    break
                except ValueError:
                    continue
            else:
                continue
        # Parse value
        try:
            val = float(str(raw_val).strip())
        except ValueError:
            continue
        records.append((dt, val))

    if not records:
        raise ValueError("未能从 Excel 中解析到有效的血糖数据")

    return {
        "datetime_col": headers[datetime_col_idx],
        "value_col": headers[value_col_idx],
        "records": records,
    }


# ---------------------------------------------------------------------------
# Word template parsing
# ---------------------------------------------------------------------------

def parse_word_template(filepath):
    """
    Parse the Word template to extract:
      - time_points: list of "HH:MM" strings (column headers, skip first '日期' col)
      - title_text: the title paragraph text (used as reference for formatting)
    """
    doc = Document(filepath)
    if not doc.tables:
        raise ValueError("Word 模板中未找到表格")

    table = doc.tables[0]
    if not table.rows:
        raise ValueError("Word 模板中表格没有行")

    header_row = table.rows[0]
    time_points = []
    for cell in header_row.cells[1:]:  # skip the '日期' column
        text = cell.text.strip()
        if text:
            time_points.append(text)

    if not time_points:
        raise ValueError("无法从 Word 模板表头中读取时间点")

    title_text = doc.paragraphs[0].text if doc.paragraphs else ""
    return {"time_points": time_points, "title_text": title_text}


# ---------------------------------------------------------------------------
# Blood-sugar matching
# ---------------------------------------------------------------------------

def match_readings(records, time_points, window_minutes=30):
    """
    For each (date, time_point) pair find the reading whose timestamp is
    closest to `date + time_point`, within `window_minutes` on either side.

    Returns a dict:  {date_str: {time_point: value_str_or_empty}}
    """
    # Group records by date (YYYY-MM-DD)
    by_date = {}
    for dt, val in records:
        date_str = dt.strftime("%Y-%m-%d")
        by_date.setdefault(date_str, []).append((dt, val))

    # Parse time_points -> timedelta objects
    tp_deltas = {}
    for tp in time_points:
        try:
            t = datetime.strptime(tp, "%H:%M")
            tp_deltas[tp] = timedelta(hours=t.hour, minutes=t.minute)
        except ValueError:
            tp_deltas[tp] = None

    result = {}
    window = timedelta(minutes=window_minutes)

    for date_str, day_records in sorted(by_date.items()):
        day_values = {}
        base_date = datetime.strptime(date_str, "%Y-%m-%d")
        for tp in time_points:
            delta = tp_deltas.get(tp)
            if delta is None:
                day_values[tp] = ""
                continue
            target_dt = base_date + delta
            best_val = None
            best_diff = None
            for dt, val in day_records:
                diff = abs(dt - target_dt)
                if diff <= window:
                    if best_diff is None or diff < best_diff:
                        best_diff = diff
                        best_val = val
            if best_val is not None:
                # Format: remove trailing zeros after decimal when integer
                if best_val == int(best_val):
                    day_values[tp] = str(int(best_val))
                else:
                    day_values[tp] = str(round(best_val, 1))
            else:
                day_values[tp] = ""
        result[date_str] = day_values

    return result


# ---------------------------------------------------------------------------
# Word document generation
# ---------------------------------------------------------------------------

def _copy_row_format(source_row, target_row):
    """Copy XML properties from source_row to target_row (trPr)."""
    src_trPr = source_row._tr.find(qn("w:trPr"))
    tgt_trPr = target_row._tr.find(qn("w:trPr"))
    if src_trPr is not None:
        new_trPr = copy.deepcopy(src_trPr)
        if tgt_trPr is not None:
            target_row._tr.remove(tgt_trPr)
        target_row._tr.insert(0, new_trPr)


def _copy_cell_format(source_cell, target_cell):
    """Copy tcPr XML from source_cell to target_cell."""
    src_tcPr = source_cell._tc.find(qn("w:tcPr"))
    tgt_tcPr = target_cell._tc.find(qn("w:tcPr"))
    if src_tcPr is not None:
        new_tcPr = copy.deepcopy(src_tcPr)
        if tgt_tcPr is not None:
            target_cell._tc.remove(tgt_tcPr)
        target_cell._tc.insert(0, new_tcPr)


def generate_word(template_path, matched_data, time_points):
    """
    Generate a Word document by copying the template and replacing/adding
    data rows.

    Returns a BytesIO object containing the .docx file.
    """
    doc = Document(template_path)

    # ---- Update title paragraph ----
    if doc.paragraphs:
        dates = sorted(matched_data.keys())
        if dates:
            new_title = f"{dates[0]}至{dates[-1]}动态血糖记录"
            title_para = doc.paragraphs[0]
            # Preserve run formatting of first run
            if title_para.runs:
                run = title_para.runs[0]
                old_bold = run.bold
                old_size = run.font.size
                old_name = run.font.name
                # Clear paragraph then rewrite
                for r in title_para.runs:
                    r.text = ""
                run.text = new_title
                run.bold = old_bold
                if old_size:
                    run.font.size = old_size
                if old_name:
                    run.font.name = old_name
            else:
                title_para.text = new_title

    # ---- Rebuild table ----
    table = doc.tables[0]
    header_row = table.rows[0]

    # Get a reference data row to copy cell/row formatting from
    ref_data_row = table.rows[1] if len(table.rows) > 1 else None

    # Remove all rows except the header
    tbl_el = table._tbl
    rows_to_remove = list(tbl_el.findall(qn("w:tr")))[1:]
    for tr in rows_to_remove:
        tbl_el.remove(tr)

    # Add data rows
    dates = sorted(matched_data.keys())
    for date_str in dates:
        day_values = matched_data[date_str]

        # Add a new row by copying the tblPrEx/trPr from reference row
        new_row = table.add_row()

        if ref_data_row:
            _copy_row_format(ref_data_row, new_row)
            for ci, cell in enumerate(new_row.cells):
                if ci < len(ref_data_row.cells):
                    _copy_cell_format(ref_data_row.cells[ci], cell)

        # Fill cells
        cells = new_row.cells
        # Date cell
        date_para = cells[0].paragraphs[0]
        date_para.clear()
        date_run = date_para.add_run(date_str)
        if ref_data_row and ref_data_row.cells[0].paragraphs[0].runs:
            src_run = ref_data_row.cells[0].paragraphs[0].runs[0]
            if src_run.font.size:
                date_run.font.size = src_run.font.size
            if src_run.font.name:
                date_run.font.name = src_run.font.name

        # Value cells
        for ci, tp in enumerate(time_points, start=1):
            if ci >= len(cells):
                break
            val = day_values.get(tp, "")
            cell_para = cells[ci].paragraphs[0]
            cell_para.clear()
            val_run = cell_para.add_run(val)
            if ref_data_row and ci < len(ref_data_row.cells):
                src_run_list = ref_data_row.cells[ci].paragraphs[0].runs
                if src_run_list:
                    src_r = src_run_list[0]
                    if src_r.font.size:
                        val_run.font.size = src_r.font.size
                    if src_r.font.name:
                        val_run.font.name = src_r.font.name

    # Save to buffer
    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf


# ---------------------------------------------------------------------------
# Routes
# ---------------------------------------------------------------------------

@app.route("/", methods=["GET"])
def index():
    return render_template("index.html")


@app.route("/generate", methods=["POST"])
def generate():
    # Validate upload
    if "excel_file" not in request.files:
        flash("请上传 Excel 数据文件", "error")
        return redirect(url_for("index"))

    excel_file = request.files["excel_file"]

    if excel_file.filename == "":
        flash("请选择文件后再提交", "error")
        return redirect(url_for("index"))

    if not allowed_file(excel_file.filename, ALLOWED_EXCEL):
        flash("Excel 文件格式不正确，请上传 .xlsx 或 .xls 文件", "error")
        return redirect(url_for("index"))

    # Save uploaded Excel file
    uid = uuid.uuid4().hex
    excel_ext = excel_file.filename.rsplit(".", 1)[1].lower()
    excel_path = os.path.join(UPLOAD_FOLDER, f"{uid}_data.{excel_ext}")
    excel_file.save(excel_path)

    try:
        # Use fixed Word template
        word_path = get_word_template_path()

        # Parse
        excel_data = parse_excel(excel_path)
        template_info = parse_word_template(word_path)

        records = excel_data["records"]
        time_points = template_info["time_points"]

        window = int(request.form.get("window_minutes", 4))

        # Match readings to time points
        matched = match_readings(records, time_points, window_minutes=window)

        # Generate output Word document
        output_buf = generate_word(word_path, matched, time_points)

        # Build a descriptive filename
        dates = sorted(matched.keys())
        if dates:
            fname = f"血糖记录_{dates[0]}至{dates[-1]}.docx"
        else:
            fname = "血糖记录.docx"

        logger.info("Generated %s with %d date rows", fname, len(dates))

        return send_file(
            output_buf,
            as_attachment=True,
            download_name=fname,
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )

    except Exception as exc:
        logger.exception("Generation failed")
        flash(f"处理失败：{exc}", "error")
        return redirect(url_for("index"))

    finally:
        # Clean up uploaded Excel file
        try:
            os.remove(excel_path)
        except OSError:
            pass


@app.route("/health")
def health():
    return jsonify({"status": "ok"})


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5002, debug=False)
