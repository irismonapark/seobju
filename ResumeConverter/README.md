# SJ - 이력서변환기

알바몬 이력서 **PDF**를 SJ **입사지원서 엑셀 양식**(`templates/sample.xlsx` 기반)으로 자동 변환합니다.

- **고정 유지**: 제목, 컬럼명, 표 레이아웃, sample 사진 위치
- **PDF에서 채움**: 이름·연락처·주소·학력·경력·자격증·사진(있을 경우)

## Git 이력 관리

```bash
# 최초 1회 (이미 되어 있으면 생략)
git init

# 변경 확인
git status
git diff

# 저장
git add .
git commit -m "변경 내용 요약"

# GitHub 연결 (최초 1회)
git remote add origin https://github.com/사용자명/ResumeConverter.git
git push -u origin main
```

Windows에서 `git` 명령이 안 되면 **Git Bash**를 쓰거나 PATH에 `C:\Program Files\Git\bin` 을 추가하세요.

## 로컬 실행

```bash
pip install -r requirements.txt
python app.py
```

- **이력서변환기:** http://127.0.0.1:5002
- 어브코스(5000)·포코스(5173)와 포트가 겹치지 않습니다.

## GitHub → Vercel 배포

코드는 [irismonapark/seobju](https://github.com/irismonapark/seobju) 의 **`ResumeConverter/`** 폴더에 있습니다.

### Vercel 대시보드 (권장)

1. [Vercel Import — seobju-resume](https://vercel.com/new/import?s=https://github.com/irismonapark/seobju&project-name=seobju-resume&root-directory=ResumeConverter) 접속
2. **Project Name:** `seobju-resume` · **Root Directory:** `ResumeConverter` 확인
3. Framework: **Flask** (자동 감지)
4. Environment Variables (권장): `SESSION_SECRET` = 임의 문자열
5. **Deploy** 클릭 → 완료 후 `https://프로젝트명.vercel.app` URL 확인

### Vercel CLI

```bash
cd ResumeConverter
npx vercel login
npx vercel deploy --prod
```

로그인 오류 시 [Vercel 토큰](https://vercel.com/account/tokens) 발급 후:

```bash
npx vercel deploy --prod --token YOUR_TOKEN
```

Vercel 대시보드 → **Functions**에서 PDF 변환용 `maxDuration: 60`, `memory: 1024` 권장. 엔트리포인트는 `pyproject.toml`의 `main:app` (→ `main.py`)입니다.

## 시스템별 헤더 색 (구분)

| 시스템 | 헤더 색 |
|--------|---------|
| 어브코스 | 남색 `#1A237E` |
| 포코스 | 청록 `#00695C` |
| 이력서변환기 | 보라 `#6A1B9A` |

## Vercel 참고

- 업로드 합계 **4MB 이하** 권장
- 변환 파일은 **같은 응답**으로 바로 다운로드 (서버리스 호환)
- 상태 확인: `/health`

## 안정성 (적용됨)

- 개인정보 `print` 제거 → 프로덕션 요약 로그만
- PDF 최대 30페이지, 텍스트 길이 검증
- 엑셀 양식 사전 검증, 파싱 결과 최소 필드 확인
- 임시 파일 자동 정리, 413/500 에러 처리
- 프론트: 4MB 사전 검사, 90초 타임아웃, XSS 방지
