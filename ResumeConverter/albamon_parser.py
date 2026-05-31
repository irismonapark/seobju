"""알바몬 PDF 텍스트 → 구조화 데이터"""
import re

EXP_PERIOD_LINE = re.compile(
    r'^(\d{4}\.\d{1,2})\s*~\s*(\d{4}\.\d{1,2})\s+(.+)$'
)
EXP_SINGLE_LINE = re.compile(r'^(\d{4}\.\d{1,2})\s+(.+)$')
DURATION_SUFFIX = re.compile(r'\s+\d+년\s*\d*\s*개?월.*$')

HOPE_LABELS = (
    '근무지', '업직종', '근무형태', '근무기간', '근무일시', '급여',
)


def _clean_company(text):
    return DURATION_SUFFIX.sub('', text).strip()


def _lines_after_header(block):
    result = []
    for line in block.split('\n'):
        line = line.strip()
        if not line or line.startswith('25.'):
            continue
        result.append(line)
    return result


def _parse_hope_lines(block):
    items = []
    for line in _lines_after_header(block):
        if line.startswith('위 모든'):
            break
        matched = False
        for label in HOPE_LABELS:
            if line.startswith(label):
                items.append(line)
                matched = True
                break
        if not matched and line:
            items.append(line)
    return items


def build_extra_text(data):
    """특이사항 영역(B23) — 어학능력·스킬·희망근무 등"""
    sections = []

    if data.get('스킬'):
        sections.append('■ 나만의 스킬 · 강점')
        sections.extend(data['스킬'])

    if data.get('어학능력'):
        if sections:
            sections.append('')
        sections.append('■ 어학능력')
        sections.extend(data['어학능력'])

    if data.get('희망근무'):
        if sections:
            sections.append('')
        sections.append('■ 희망근무조건')
        sections.extend(data['희망근무'])

    return '\n'.join(sections).strip()


def parse_albamon_resume(text):
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
        '자기소개': '',
        '스킬': [],
        '어학능력': [],
        '희망근무': [],
    }

    lines = [ln.strip() for ln in text.split('\n') if ln.strip()]

    for line in lines:
        m = re.match(r'^([가-힣]{2,4})\s+여(?:자|성)\s*(\d+)세/?\s*(\d{4})년생', line)
        if m:
            data['이름'] = m.group(1)
            data['성별'] = '여' if '여' in line else '남'
            data['생년월일'] = f'{m.group(3)}년생'
            continue

        m = re.search(r'작성자\s+([가-힣]{2,4})', line)
        if m and not data['이름']:
            data['이름'] = m.group(1)

    for line in lines:
        if line.startswith('주소'):
            addr = line.replace('주소', '', 1).strip()
            addr = re.split(r'\s+이메일', addr)[0].strip()
            data['주소'] = addr

        phone = re.search(r'(?:휴대폰|전화)\s*(\d{3}[-\s]?\d{3,4}[-\s]?\d{4})', line)
        if phone:
            data['연락처'] = phone.group(1).replace(' ', '')

        email = re.search(
            r'이메일\s+([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})',
            line,
            re.I,
        )
        if email:
            data['이메일'] = email.group(1)

    if not data['생년월일']:
        birth = re.search(r'(\d{4})년생', text)
        if birth:
            data['생년월일'] = f'{birth.group(1)}년생'

    edu_match = re.search(r'학력\s*\n(.*?)(?=\n경력\s*\n|\n경력\n)', text, re.DOTALL)
    if edu_match:
        for line in edu_match.group(1).split('\n'):
            line = line.strip()
            if not line or re.match(r'^\d+년', line) or line == '학력':
                continue
            if any(k in line for k in ['대학', '고등', '중학', '졸업', '재학', '수료']):
                data['학력'].append(line)

    exp_match = re.search(
        r'경력\s*\n(.*?)(?=자기소개|나만의\s*스킬|어학능력|희망근무|$)',
        text,
        re.DOTALL,
    )
    if exp_match:
        data['경력'] = _parse_experience_block(exp_match.group(1))

    intro = re.search(
        r'자기소개\s*\n(.+?)(?=\n\d{2}\.\s*\d{2}\.|나만의|어학|희망|$)',
        text,
        re.DOTALL,
    )
    if intro:
        intro_lines = [
            ln.strip()
            for ln in intro.group(1).split('\n')
            if ln.strip() and not ln.startswith('25.')
        ]
        if intro_lines:
            data['자기소개'] = intro_lines[0]

    skill_match = re.search(
        r'나만의\s*스킬.*?\n(.*?)(?=어학능력|희망근무|$)',
        text,
        re.DOTALL,
    )
    if skill_match:
        for line in _lines_after_header(skill_match.group(1)):
            if line.startswith('나의 '):
                data['스킬'].append(line)
            elif line and 'MBTI' not in line and 'ENTP' not in line[:4]:
                if len(line) > 3:
                    data['스킬'].append(line)

    lang_match = re.search(r'어학능력\s*\n(.*?)(?=희망근무|$)', text, re.DOTALL)
    if lang_match:
        for line in _lines_after_header(lang_match.group(1)):
            if line and line != '어학능력':
                data['어학능력'].append(line)

    hope_match = re.search(r'희망근무조건\s*\n(.*?)(?=위 모든|$)', text, re.DOTALL)
    if hope_match:
        data['희망근무'] = _parse_hope_lines(hope_match.group(1))

    return data


def _parse_experience_block(block):
    experiences = []
    lines = [ln.strip() for ln in block.split('\n') if ln.strip()]
    pending = None

    for line in lines:
        if line == '경력' or re.match(r'^\d+년\s*\d*개?월', line):
            continue
        if '이런 경력' in line or 'PICK' in line or '근무' == line:
            continue

        m = EXP_PERIOD_LINE.match(line)
        if m:
            if pending:
                experiences.append(pending)
            pending = {
                'period': f'{m.group(1)} ~ {m.group(2)}',
                'company': _clean_company(m.group(3)),
                'duty': '',
                'note': '',
            }
            continue

        m = EXP_SINGLE_LINE.match(line)
        if m:
            if pending:
                experiences.append(pending)
            pending = {
                'period': m.group(1),
                'company': _clean_company(m.group(2)),
                'duty': '',
                'note': '',
            }
            continue

        if pending and not re.match(r'^\d{4}', line):
            if pending['duty']:
                pending['duty'] += ' / ' + line
            else:
                pending['duty'] = line

    if pending:
        experiences.append(pending)

    return experiences
