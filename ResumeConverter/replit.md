# 알바몬 → SJ 이력서 변환 서비스

## 프로젝트 개요
알바몬 지원자 이력서를 SJ 이력서 양식으로 자동 변환하는 웹 애플리케이션

## 최근 변경사항
- **2025-11-12**: 초기 프로젝트 생성
  - Flask 웹 애플리케이션 구조 구축
  - Material Design 기반 한글 UI 구현
  - PDF 텍스트 추출 및 파싱 기능
  - 엑셀 템플릿 매핑 및 변환 기능
  - 모바일 반응형 디자인 적용
  - **병합된 셀 자동 처리**: MergedCell 오류 방지를 위한 자동 병합 해제 기능 추가
  - **보안 강화**: Path traversal 취약점 수정, 파일명 검증 강화

## 기술 스택
- **Backend**: Python 3.11, Flask
- **PDF Processing**: pdfplumber
- **Excel Processing**: openpyxl
- **Frontend**: HTML5, CSS3 (Material Design), Vanilla JavaScript
- **Font**: Noto Sans KR (Google Fonts)
- **Icons**: Material Icons

## 프로젝트 구조
```
.
├── app.py                  # Flask 메인 애플리케이션
├── templates/
│   └── index.html         # 메인 페이지 템플릿
├── static/
│   ├── css/
│   │   └── style.css      # Material Design 스타일
│   └── js/
│       └── script.js      # 프론트엔드 로직
├── .gitignore
└── replit.md
```

## 주요 기능
1. **이중 입력 방식**
   - PDF 파일 업로드 (드래그 앤 드롭 지원)
   - 텍스트 직접 붙여넣기 (모바일 지원)

2. **자동 데이터 추출**
   - 이름, 연락처, 이메일, 주소, 생년월일
   - 학력 정보
   - 경력 정보
   - 자격증 정보

3. **엑셀 템플릿 변환**
   - SJ 양식에 자동 매핑
   - 즉시 다운로드 가능

4. **UI/UX**
   - Material Design 디자인 시스템
   - 한글 레이블 및 메시지
   - 모바일/데스크톱 반응형
   - 드래그 앤 드롭 파일 업로드
   - 실시간 입력 검증

## 환경 변수
- `SESSION_SECRET`: Flask 세션 시크릿 키 (자동 설정됨)

## 실행 방법
프로젝트는 자동으로 실행됩니다. 포트 5000에서 Flask 개발 서버가 구동됩니다.

## 파일 제한
- 최대 파일 크기: 10MB
- 지원 형식: PDF (.pdf), Excel (.xlsx, .xls)

## 데이터 매핑
알바몬 이력서에서 추출된 데이터를 SJ 엑셀 템플릿의 다음 위치에 매핑합니다:

| 필드 | 엑셀 셀 위치 |
|------|------------|
| 이름 | A2 |
| 연락처 | B2 |
| 이메일 | C2 |
| 주소 | D2 |
| 생년월일 | E2 |
| 학력 | A5:A9 |
| 경력 | A15:A24 |
| 자격증 | A30:A34 |

## 보안 고려사항
- 업로드된 파일은 임시 디렉토리에 저장
- 변환 후 자동 정리
- 파일 형식 검증
- 파일 크기 제한

## 향후 개선 사항
- AI/LLM 기반 정확도 향상
- 배치 변환 기능
- 필드 매핑 커스터마이징
- 미리보기 기능
- 변환 히스토리
