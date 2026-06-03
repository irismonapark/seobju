import io
import logging
import os
import re
import shutil
import tempfile
import uuid
import zipfile

from template_config import (
    CERT_DATA_COLS,
    CERT_MAX_ROWS,
    CERT_START_ROW,
    CLEAR_SAMPLE_CELLS,
    CLEAR_SAMPLE_ROWS,
    EDU_DATA_COLS,
    EDU_MAX_ROWS,
    EDU_START_ROW,
    EXP_DATA_COLS,
    EXP_MAX_ROWS,
    EXP_START_ROW,
    EXTRA_CELL,
    INTRO_CELL,
    SCALAR_CELLS,
    SIGNATURE_CELL,
)
from flask import Flask, Response, jsonify, request, send_file
from werkzeug.exceptions import RequestEntityTooLarge

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

def _is_vercel():
    return bool(os.environ.get('VERCEL') or os.environ.get('VERCEL_ENV'))


app = Flask(__name__)
app.config['SECRET_KEY'] = os.environ.get('SESSION_SECRET', 'dev-secret-key')
app.config['MAX_CONTENT_LENGTH'] = 4 * 1024 * 1024
if _is_vercel():
    try:
        os.makedirs('/tmp', exist_ok=True)
    except OSError:
        pass
    app.config['UPLOAD_FOLDER'] = '/tmp'
else:
    app.config['UPLOAD_FOLDER'] = tempfile.gettempdir()

IS_PRODUCTION = _is_vercel() or os.environ.get('FLASK_ENV') == 'production'
MAX_PDF_PAGES = 30
MAX_TEXT_LENGTH = 10000
MIN_TEXT_LENGTH = 20

ALLOWED_PDF_EXTENSIONS = {'pdf'}

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_PATH = os.path.join(BASE_DIR, 'templates', 'resume_template.xlsx')
SAMPLE_SOURCE_PATH = os.path.join(BASE_DIR, 'templates', 'sample_re.xlsx')
SAMPLE_RE_DOWNLOAD = os.path.join(
    os.path.expanduser('~'),
    'Downloads',
    'ResumeConverter',
    'sample_re.xlsx',
)


def log_debug(message, *args):
    if not IS_PRODUCTION:
        logger.info(message, *args)


def safe_remove(path):
    if path and os.path.exists(path):
        try:
            os.remove(path)
        except OSError as exc:
            logger.warning('임시 파일 삭제 실패: %s (%s)', path, exc)


def allowed_file(filename, allowed_extensions):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in allowed_extensions


def extract_text_from_pdf(pdf_path):
    import pdfplumber

    try:
        with pdfplumber.open(pdf_path) as pdf:
            if len(pdf.pages) > MAX_PDF_PAGES:
                raise ValueError(
                    f'PDF 페이지가 너무 많습니다 (최대 {MAX_PDF_PAGES}페이지).'
                )

            text_parts = []
            for page in pdf.pages:
                page_text = page.extract_text()
                if page_text:
                    text_parts.append(page_text)

            text = '\n'.join(text_parts)
            if not text.strip():
                raise ValueError(
                    '해당 PDF파일은 변환을 지원하지 않습니다. '
                    '알바몬에서 브라우저 인쇄(Ctrl+P) → PDF로 저장 후 업로드해 주세요.'
                )

            log_debug('PDF 텍스트 추출 완료 (%d자)', len(text))
            return text
    except ValueError:
        raise
    except Exception as exc:
        raise ValueError(f'PDF 파일을 읽을 수 없습니다: {exc}') from exc


def validate_resume_text(text):
    cleaned = (text or '').strip()
    if len(cleaned) < MIN_TEXT_LENGTH:
        raise ValueError(
            f'이력서 내용이 너무 짧습니다 (최소 {MIN_TEXT_LENGTH}자).'
        )
    if len(cleaned) > MAX_TEXT_LENGTH:
        raise ValueError(
            f'이력서 내용이 너무 깁니다 (최대 {MAX_TEXT_LENGTH}자).'
        )
    return cleaned


def ensure_template_exists():
    from pathlib import Path

    template_path = Path(TEMPLATE_PATH)
    if template_path.is_file() and template_path.stat().st_size > 1000:
        return

    if os.environ.get('VERCEL') or os.environ.get('VERCEL_ENV'):
        raise ValueError(
            '이력서 양식(resume_template.xlsx)이 서버에 없습니다. 관리자에게 문의하세요.'
        )

    from repair_template import repair_xlsx, resolve_source_path

    source = resolve_source_path()
    if not source.is_file():
        if os.path.isfile(SAMPLE_RE_DOWNLOAD):
            Path(SAMPLE_SOURCE_PATH).parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(SAMPLE_RE_DOWNLOAD, SAMPLE_SOURCE_PATH)
            source = Path(SAMPLE_SOURCE_PATH)
        elif os.path.isfile(SAMPLE_SOURCE_PATH):
            source = Path(SAMPLE_SOURCE_PATH)

    template_path = Path(TEMPLATE_PATH)
    need_build = (
        not template_path.is_file()
        or (source.is_file() and source.stat().st_mtime > template_path.stat().st_mtime)
    )
    if need_build and source.is_file():
        try:
            repair_xlsx(source, template_path)
            return
        except Exception as exc:
            logger.warning('양식 생성 실패: %s', exc)

    if template_path.is_file():
        return
    raise ValueError('이력서 양식(sample_re)이 서버에 없습니다. 관리자에게 문의하세요.')


def extract_photo_from_pdf(pdf_path):
    """PDF에서 가장 큰 사진 추출 (없으면 None → sample 기본 사진 유지)"""
    try:
        import fitz
    except ImportError:
        log_debug('pymupdf 미설치 — sample 사진 유지')
        return None

    try:
        doc = fitz.open(pdf_path)
        best_bytes = None
        best_area = 0
        for page in doc:
            for info in page.get_images(full=True):
                xref = info[0]
                extracted = doc.extract_image(xref)
                area = extracted['width'] * extracted['height']
                if area > best_area and area > 4000:
                    best_area = area
                    best_bytes = extracted['image']
        doc.close()
        return prepare_photo_bytes(best_bytes) if best_bytes else None
    except Exception as exc:
        log_debug('PDF 사진 추출 실패: %s', exc)
        return None


def prepare_photo_bytes(raw_bytes):
    from PIL import Image

    image = Image.open(io.BytesIO(raw_bytes))
    if image.mode not in ('RGB', 'RGBA'):
        image = image.convert('RGB')
    image.thumbnail((120, 150))
    buffer = io.BytesIO()
    image.save(buffer, format='PNG')
    return buffer.getvalue()


def replace_photo_in_xlsx(xlsx_path, image_bytes):
    """sample 양식 내 사진 파일만 교체 (위치·크기 유지)"""
    if not image_bytes:
        return

    with zipfile.ZipFile(xlsx_path, 'r') as zin:
        archive = {name: zin.read(name) for name in zin.namelist()}

    media_files = [
        name for name in archive
        if name.startswith('xl/media/') and name.lower().endswith(('.png', '.jpg', '.jpeg'))
    ]
    if not media_files:
        return

    archive[media_files[0]] = image_bytes
    temp_path = xlsx_path + '.tmp'
    with zipfile.ZipFile(temp_path, 'w', compression=zipfile.ZIP_DEFLATED) as zout:
        for name, data in archive.items():
            zout.writestr(name, data)
    os.replace(temp_path, xlsx_path)


def make_output_filename(name):
    cleaned = re.sub(r'[\\/:*?"<>|\r\n]+', '', (name or '').strip()) or '이름없음'
    return f'알바몬_{cleaned}님_이력서.xlsx'


def split_experience_line(line):
    """경력 한 줄 → (기간, 회사/업무, 담당업무)"""
    period_match = re.search(
        r'\(?\s*(\d{4}[\.\/\-][\d\.\/\-\s]*?)\s*~\s*([\d\.\/\-\s]+?)\s*\)?\s*$',
        line,
    )
    if period_match:
        period = f'{period_match.group(1).strip()} ~ {period_match.group(2).strip()}'
        remainder = line[:period_match.start()].strip(' ()')
        return period, remainder, ''
    return '', line, ''


def validate_parsed_data(data):
    if not data.get('이름'):
        raise ValueError(
            '이력서에서 이름을 찾지 못했습니다. 알바몬 PDF인지 확인해주세요.'
        )
    if not data.get('연락처'):
        raise ValueError(
            '이력서에서 연락처(휴대폰)를 찾지 못했습니다.'
        )


def parse_resume_data(text):
    """알바몬 PDF 우선 파싱, 실패 시 일반 형식 폴백"""
    from albamon_parser import parse_albamon_resume

    albamon = parse_albamon_resume(text)
    if albamon.get('이름'):
        log_debug(
            '알바몬 파싱: 이름=%s, 학력=%d, 경력=%d',
            albamon['이름'],
            len(albamon.get('학력', [])),
            len(albamon.get('경력', [])),
        )
        return albamon
    return _parse_resume_data_generic(text)


def _parse_resume_data_generic(text):
    data = {
        '이름': '',
        '연락처': '',
        '이메일': '',
        '주소': '',
        '생년월일': '',
        '성별': '',
        '학력': [],
        '경력': [],
        '자격증': [],
    }

    name_patterns = [
        r'(?:이름|성명|姓名)[\s:：\t]*([가-힣]{2,4})',
        r'^([가-힣]{2,4})[\s\n]',
        r'지원자[\s:：]*([가-힣]{2,4})',
    ]
    for pattern in name_patterns:
        name_match = re.search(pattern, text, re.MULTILINE)
        if name_match:
            data['이름'] = name_match.group(1).strip()
            break

    phone_patterns = [
        r'(?:휴대폰|전화|연락처|H\.?P\.?|Mobile)[\s:：\t]*(\d{2,3}[-\s\.]?\d{3,4}[-\s\.]?\d{4})',
        r'(\d{3}[-\s\.]\d{4}[-\s\.]\d{4})',
        r'(010[-\s\.]\d{4}[-\s\.]\d{4})',
    ]
    for pattern in phone_patterns:
        phone_match = re.search(pattern, text)
        if phone_match:
            data['연락처'] = phone_match.group(1).strip()
            break

    email_patterns = [
        r'(?:이메일|E-?mail)[\s:：\t]*([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})',
        r'([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})',
    ]
    for pattern in email_patterns:
        email_match = re.search(pattern, text, re.IGNORECASE)
        if email_match:
            data['이메일'] = email_match.group(1).strip()
            break

    birth_patterns = [
        r'(?:생년월일|생일|출생)[\s:：\t]*(\d{4}[-./년\s]\d{1,2}[-./월\s]\d{1,2}일?)',
        r'(\d{4}년\s?\d{1,2}월\s?\d{1,2}일)',
        r'(\d{4}\.\d{1,2}\.\d{1,2})',
        r'(\d{6}[-]\d{7})',
    ]
    for pattern in birth_patterns:
        birth_match = re.search(pattern, text)
        if birth_match:
            data['생년월일'] = birth_match.group(1).strip()
            break

    address_patterns = [
        r'(?:주소|거주지|현주소)[\s:：\t]*([^\n]{10,150})',
        r'(?:서울|경기|인천|부산|대구|광주|대전|울산|세종|강원|충북|충남|전북|전남|경북|경남|제주).*?(?:시|군|구).*?(?:동|읍|면|로|길)[^\n]{0,50}',
    ]
    for pattern in address_patterns:
        address_match = re.search(pattern, text)
        if address_match:
            data['주소'] = (
                address_match.group(0)
                .replace('주소', '')
                .replace('거주지', '')
                .replace('현주소', '')
                .strip(':： \t')
                .strip()
            )
            break

    gender_match = re.search(r'(?:성별)[\s:：\t]*(남|여|남성|여성)', text)
    if gender_match:
        data['성별'] = gender_match.group(1).strip()

    def is_section_header(line, keywords, max_len=20):
        stripped = line.strip()
        if not stripped or len(stripped) > max_len:
            return False
        return any(
            stripped == kw or stripped.startswith(kw + ' ') or stripped.endswith(' ' + kw)
            for kw in keywords
        )

    in_education = False
    in_experience = False
    in_certificate = False

    for line in text.split('\n'):
        line = line.strip()
        if not line:
            continue

        if is_section_header(line, ['학력', '최종학력', '교육']):
            in_education, in_experience, in_certificate = True, False, False
            continue
        if is_section_header(line, ['경력', '근무경력', '경험'], max_len=25):
            in_education, in_experience, in_certificate = False, True, False
            continue
        if is_section_header(line, ['자격증', '자격사항', '면허']):
            in_education, in_experience, in_certificate = False, False, True
            continue
        if is_section_header(line, ['개인정보', '인적사항', '기본정보'], max_len=25):
            in_education, in_experience, in_certificate = False, False, False
            continue

        if len(line) <= 2:
            continue

        if in_education and any(
            word in line
            for word in ['대학', '고등학교', '중학교', '초등학교', '졸업', '재학', '수료', '학과', '전공']
        ):
            data['학력'].append(line)
        elif in_experience and (
            any(word in line for word in ['주식회사', '(주)', '회사', '근무', '재직', '퇴사', '~', '-', '년', '월'])
            or re.search(r'\d{4}', line)
        ):
            data['경력'].append(line)
        elif in_certificate and not any(skip in line for skip in ['항목', '구분', '번호']):
            data['자격증'].append(line)

    log_debug(
        '파싱 요약: 이름=%s, 학력=%d, 경력=%d, 자격증=%d',
        data['이름'] or '(없음)',
        len(data['학력']),
        len(data['경력']),
        len(data['자격증']),
    )
    return data


def clear_sample_data(ws):
    """sample 예시 데이터만 지우고 제목·컬럼명·병합 레이아웃은 유지"""
    from excel_layout import clear_cell

    edu_start, edu_end = CLEAR_SAMPLE_ROWS['edu']
    for row in range(edu_start, edu_end + 1):
        for col in EDU_DATA_COLS:
            clear_cell(ws, f'{col}{row}')

    exp_start, exp_end = CLEAR_SAMPLE_ROWS['exp']
    for row in range(exp_start, exp_end + 1):
        for col in EXP_DATA_COLS:
            clear_cell(ws, f'{col}{row}')

    cert_start, cert_end = CLEAR_SAMPLE_ROWS['cert']
    for row in range(cert_start, cert_end + 1):
        for col in CERT_DATA_COLS:
            clear_cell(ws, f'{col}{row}')

    for cell in CLEAR_SAMPLE_CELLS:
        clear_cell(ws, cell)

    for cell in SCALAR_CELLS.values():
        clear_cell(ws, cell)


def fill_excel_template(data, photo_bytes=None):
    """sample.xlsx 원본 양식 복사 후 PDF 데이터만 채움"""
    import openpyxl

    from albamon_parser import build_extra_text
    from excel_layout import (
        apply_workbook_layout,
        format_signature_line,
        split_education_line,
        write_cell,
    )

    ensure_template_exists()

    output_filename = make_output_filename(data.get('이름', ''))
    output_path = os.path.join(app.config['UPLOAD_FOLDER'], output_filename)
    shutil.copy2(TEMPLATE_PATH, output_path)

    wb = openpyxl.load_workbook(output_path)
    ws = wb.active

    clear_sample_data(ws)

    for field, cell in SCALAR_CELLS.items():
        if data.get(field):
            write_cell(ws, cell, data[field])

    edu_row = EDU_START_ROW
    for edu in data.get('학력', [])[:EDU_MAX_ROWS]:
        school, grad = split_education_line(edu)
        if school:
            write_cell(ws, f'E{edu_row}', school)
        if grad:
            write_cell(ws, f'O{edu_row}', grad)
        elif edu:
            write_cell(ws, f'E{edu_row}', edu)
        edu_row += 1

    exp_row = EXP_START_ROW
    for exp in data.get('경력', [])[:EXP_MAX_ROWS]:
        if isinstance(exp, dict):
            if exp.get('period'):
                write_cell(ws, f'B{exp_row}', exp['period'])
            if exp.get('company'):
                write_cell(ws, f'E{exp_row}', exp['company'])
            if exp.get('duty'):
                write_cell(ws, f'J{exp_row}', exp['duty'])
            if exp.get('note'):
                write_cell(ws, f'S{exp_row}', exp['note'])
        else:
            period, company, duty = split_experience_line(exp)
            if period:
                write_cell(ws, f'B{exp_row}', period)
            if company:
                write_cell(ws, f'E{exp_row}', company)
            if duty:
                write_cell(ws, f'J{exp_row}', duty)
            elif not period and not company:
                write_cell(ws, f'E{exp_row}', exp)
        exp_row += 1

    cert_row = CERT_START_ROW
    for cert in data.get('자격증', [])[:CERT_MAX_ROWS]:
        write_cell(ws, f'B{cert_row}', cert)
        cert_row += 1

    if data.get('자기소개'):
        write_cell(ws, INTRO_CELL, data['자기소개'])

    extra = build_extra_text(data)
    if extra:
        write_cell(ws, EXTRA_CELL, extra)

    name = data.get('이름', '')
    if name:
        write_cell(ws, SIGNATURE_CELL, format_signature_line(name))

    apply_workbook_layout(ws)

    wb.save(output_path)
    wb.close()

    if photo_bytes:
        replace_photo_in_xlsx(output_path, photo_bytes)

    log_debug('엑셀 저장 완료: %s', output_filename)
    return output_path, output_filename


@app.errorhandler(RequestEntityTooLarge)
def handle_request_too_large(_error):
    return jsonify({
        'success': False,
        'error': 'PDF 파일이 4MB를 초과합니다.',
    }), 413


@app.route('/health')
def health():
    template_ok = os.path.isfile(TEMPLATE_PATH)
    return jsonify({
        'status': 'ok',
        'service': 'resume-converter',
        'template': template_ok,
    }), 200


@app.route('/')
def index():
    # Vercel: public/는 CDN 전용이라 서버리스 번들에 없음 → templates 사용
    index_path = os.path.join(BASE_DIR, 'templates', 'index.html')
    if not os.path.isfile(index_path):
        index_path = os.path.join(BASE_DIR, 'public', 'index.html')
    if os.path.isfile(index_path):
        with open(index_path, 'rb') as html_file:
            return Response(html_file.read(), mimetype='text/html; charset=utf-8')
    return jsonify({'success': False, 'error': '화면 파일을 찾을 수 없습니다.'}), 404


@app.route('/convert', methods=['POST'])
def convert():
    temp_pdf = None
    output_path = None

    try:
        if not request.files.get('resume_pdf') or not request.files['resume_pdf'].filename:
            return jsonify({
                'success': False,
                'error': '알바몬 이력서 PDF 파일을 업로드해주세요.',
            }), 400

        pdf_file = request.files['resume_pdf']
        if not allowed_file(pdf_file.filename, ALLOWED_PDF_EXTENSIONS):
            return jsonify({'success': False, 'error': '올바른 PDF 파일이 아닙니다.'}), 400

        temp_pdf = os.path.join(
            app.config['UPLOAD_FOLDER'],
            f'{uuid.uuid4().hex}.pdf',
        )
        pdf_file.save(temp_pdf)
        resume_text = validate_resume_text(extract_text_from_pdf(temp_pdf))
        photo_bytes = extract_photo_from_pdf(temp_pdf)

        parsed_data = parse_resume_data(resume_text)
        validate_parsed_data(parsed_data)

        output_path, output_filename = fill_excel_template(parsed_data, photo_bytes)

        with open(output_path, 'rb') as excel_file:
            excel_bytes = io.BytesIO(excel_file.read())
        safe_remove(output_path)
        output_path = None

        excel_bytes.seek(0)
        return send_file(
            excel_bytes,
            as_attachment=True,
            download_name=output_filename,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        )

    except ValueError as exc:
        safe_remove(output_path)
        return jsonify({'success': False, 'error': str(exc)}), 400
    except Exception as exc:
        safe_remove(output_path)
        logger.exception('변환 실패')
        return jsonify({
            'success': False,
            'error': str(exc) if isinstance(exc, OSError) else (
                '변환 중 오류가 발생했습니다. 파일 형식과 내용을 확인해주세요.'
            ),
        }), 500
    finally:
        safe_remove(temp_pdf)


application = app

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5002))
    debug = os.environ.get('FLASK_DEBUG', '1').lower() in ('1', 'true', 'yes')
    print(f'이력서변환기: http://127.0.0.1:{port}')
    app.run(host='0.0.0.0', port=port, debug=debug)
