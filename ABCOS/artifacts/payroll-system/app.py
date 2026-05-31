import io
import logging
import os
import sys
import tempfile
from contextlib import contextmanager

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from flask import Flask, jsonify, render_template, request, send_file
from werkzeug.exceptions import RequestEntityTooLarge

from utils.bc_generator import generate_bc_excel
from utils.errors import UserFacingError
from utils.excel_parser import extract_company_name, parse_a_file
from utils.pdf_generator import generate_pdf_zip, parse_bc_file
from utils.payslip_excel_generator import generate_payslip_excel_zip

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(name)s: %(message)s",
)
logger = logging.getLogger(__name__)

app = Flask(__name__)

_default_max_mb = 4 if os.environ.get("VERCEL") else 50
try:
    _max_mb = int(os.environ.get("MAX_UPLOAD_MB", _default_max_mb))
except ValueError:
    _max_mb = _default_max_mb
_max_mb = max(1, min(_max_mb, 100))
app.config["MAX_CONTENT_LENGTH"] = _max_mb * 1024 * 1024

ALLOWED_EXTENSIONS = {"xls", "xlsx"}


def allowed_file(filename: str) -> bool:
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


@contextmanager
def temp_path(suffix: str):
    path = None
    try:
        with tempfile.NamedTemporaryFile(suffix=suffix, delete=False) as tmp:
            path = tmp.name
        yield path
    finally:
        if path and os.path.exists(path):
            try:
                os.unlink(path)
            except OSError:
                logger.warning("Failed to delete temp file: %s", path)


def _read_and_delete(path: str) -> io.BytesIO:
    with open(path, "rb") as f:
        data = f.read()
    try:
        os.unlink(path)
    except OSError:
        logger.warning("Failed to delete output temp file: %s", path)
    return io.BytesIO(data)


def _json_error(message: str, status: int = 400):
    return jsonify({"error": message}), status


def _validate_upload(file):
    if file is None or not file.filename:
        raise UserFacingError("파일을 선택해주세요.")

    if not allowed_file(file.filename):
        raise UserFacingError("xls 또는 xlsx 파일만 업로드 가능합니다.")

    file.seek(0, os.SEEK_END)
    size = file.tell()
    file.seek(0)
    if size == 0:
        raise UserFacingError("빈 파일입니다. 다른 파일을 선택해 주세요.")

    max_bytes = app.config["MAX_CONTENT_LENGTH"]
    if size > max_bytes:
        max_mb = max_bytes // (1024 * 1024)
        raise UserFacingError(f"파일 크기가 {max_mb}MB를 초과합니다.")


def _handle_conversion_error(action: str, exc: Exception):
    if isinstance(exc, UserFacingError):
        return _json_error(exc.message, 400)
    logger.exception("%s failed", action)
    return _json_error(f"{action} 중 오류가 발생했습니다. 파일 형식을 확인해 주세요.", 500)


@app.errorhandler(413)
@app.errorhandler(RequestEntityTooLarge)
def handle_payload_too_large(_exc):
    max_mb = app.config["MAX_CONTENT_LENGTH"] // (1024 * 1024)
    return _json_error(f"파일 크기가 {max_mb}MB를 초과합니다.", 413)


@app.route("/")
def index():
    max_upload_mb = app.config["MAX_CONTENT_LENGTH"] // (1024 * 1024)
    return render_template("index.html", max_upload_mb=max_upload_mb)


@app.route("/healthz")
def health():
    return jsonify({"status": "ok"})


@app.route("/convert/bc", methods=["POST"])
def convert_bc():
    if "file" not in request.files:
        return _json_error("파일이 업로드되지 않았습니다.")

    file = request.files["file"]
    try:
        _validate_upload(file)
        suffix = "." + file.filename.rsplit(".", 1)[1].lower()

        with temp_path(suffix) as input_tmp:
            file.save(input_tmp)
            parsed = parse_a_file(input_tmp)

        year = parsed["year"]
        month = parsed["month"]
        company = extract_company_name(file.filename)

        with temp_path(".xlsx") as output_tmp:
            generate_bc_excel(parsed, output_tmp, company=company)
            buf = _read_and_delete(output_tmp)

        company_suffix = f"_{company}" if company else ""
        download_name = f"{year}-{month:02d}{company_suffix}.xlsx"
        return send_file(
            buf,
            as_attachment=True,
            download_name=download_name,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    except Exception as exc:
        return _handle_conversion_error("변환", exc)


@app.route("/convert/pdf", methods=["POST"])
def convert_pdf():
    if "file" not in request.files:
        return _json_error("파일이 업로드되지 않았습니다.")

    file = request.files["file"]
    try:
        _validate_upload(file)
        suffix = "." + file.filename.rsplit(".", 1)[1].lower()

        with temp_path(suffix) as input_tmp:
            file.save(input_tmp)
            parsed = parse_bc_file(input_tmp)

        year = parsed["year"]
        month = parsed["month"]

        with temp_path(".zip") as output_tmp:
            generate_pdf_zip(parsed, output_tmp)
            buf = _read_and_delete(output_tmp)

        download_name = f"급여명세서_{year}-{month:02d}.zip"
        return send_file(
            buf,
            as_attachment=True,
            download_name=download_name,
            mimetype="application/zip",
        )
    except Exception as exc:
        return _handle_conversion_error("PDF 생성", exc)


@app.route("/convert/payslip-excel", methods=["POST"])
def convert_payslip_excel():
    if "file" not in request.files:
        return _json_error("파일이 업로드되지 않았습니다.")

    file = request.files["file"]
    try:
        _validate_upload(file)
        suffix = "." + file.filename.rsplit(".", 1)[1].lower()

        with temp_path(suffix) as input_tmp:
            file.save(input_tmp)
            parsed = parse_bc_file(input_tmp)

        year = parsed["year"]
        month = parsed["month"]
        company = extract_company_name(file.filename)

        with temp_path(".zip") as output_tmp:
            generate_payslip_excel_zip(parsed, output_tmp, company=company)
            buf = _read_and_delete(output_tmp)

        company_suffix = f"_{company}" if company else ""
        download_name = f"{year}-{month:02d}{company_suffix}.zip"
        return send_file(
            buf,
            as_attachment=True,
            download_name=download_name,
            mimetype="application/zip",
        )
    except Exception as exc:
        return _handle_conversion_error("엑셀 급여명세서 생성", exc)


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
