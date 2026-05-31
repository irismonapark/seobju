"""sample.xlsx → openpyxl 호환 + 서식/이미지 유지 resume_template.xlsx"""
import re
import shutil
import zipfile
from pathlib import Path

import openpyxl

BASE = Path(__file__).parent
SOURCE_CANDIDATES = [
    Path(r'c:\Users\paulc\Downloads\ResumeConverter\sample_re.xlsx'),
    BASE / 'templates' / 'sample_re.xlsx',
    BASE / 'templates' / 'sample.xlsx',
    Path(r'c:\Users\paulc\Downloads\sample.xlsx'),
]
TARGET = BASE / 'templates' / 'resume_template.xlsx'


def resolve_source_path():
    for path in SOURCE_CANDIDATES:
        if path.is_file():
            return path
    return SOURCE_CANDIDATES[1]

MINIMAL_STYLES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font><sz val="11"/><name val="맑은 고딕"/></font></fonts>
  <fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>"""


def strip_all_style_refs(sheet_xml: str) -> str:
    sheet_xml = re.sub(r'(<x:c\b[^>]*)\s+s="\d+"', r'\1', sheet_xml)
    sheet_xml = re.sub(r'(<x:col\b[^>]*)\s+style="\d+"', r'\1', sheet_xml)
    sheet_xml = re.sub(r'(<x:row\b[^>]*)\s+s="\d+"', r'\1', sheet_xml)
    return sheet_xml


def repair_xlsx(src: Path, dst: Path):
    """시트 style 참조만 제거(병합·열·이미지 유지). styles.xml은 가능하면 유지."""
    dst.parent.mkdir(parents=True, exist_ok=True)

    def write_repaired(replace_styles: bool):
        with zipfile.ZipFile(src, 'r') as zin:
            with zipfile.ZipFile(dst, 'w', compression=zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    data = zin.read(item.filename)
                    if replace_styles and item.filename == 'xl/styles.xml':
                        data = MINIMAL_STYLES.encode('utf-8')
                    elif item.filename == 'xl/worksheets/sheet1.xml':
                        data = strip_all_style_refs(data.decode('utf-8')).encode('utf-8')
                    zout.writestr(item, data)

    write_repaired(replace_styles=False)
    try:
        openpyxl.load_workbook(dst).close()
    except Exception:
        write_repaired(replace_styles=True)


def main():
    source = resolve_source_path()
    repair_xlsx(source, TARGET)
    wb = openpyxl.load_workbook(TARGET)
    ws = wb.active
    print('OK', ws.title)
    print('merged', len(list(ws.merged_cells.ranges)))
    wb.close()


if __name__ == '__main__':
    main()
