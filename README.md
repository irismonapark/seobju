# seobju

| 폴더 | 설명 |
|---|---|
| **ABCOS/** | SJ - 어브코스 급여정산 자동화 시스템 |
| **pocos/** | SJ - 포코스 급여관리 시스템 (근무현황→청구서→급여명세서) |
| **ResumeConverter/** | SJ - 알바몬 이력서 PDF → 입사지원서 엑셀 변환 |

## Vercel 배포

### pocos (포코스 급여관리)
1. [Vercel Import — pocos](https://vercel.com/new/import?s=https://github.com/irismonapark/seobju&project-name=seobju-pocos&root-directory=pocos&framework=vite) 접속
2. **Root Directory:** `pocos` · **Framework:** Vite 확인
3. **Deploy** → URL 예: `https://seobju-pocos.vercel.app`

### ABCOS
1. [Vercel New Project](https://vercel.com/new) → `irismonapark/seobju` Import
2. **Root Directory:** `ABCOS`
3. Deploy

### pocos
1. 동일 저장소 Import
2. **Root Directory:** `pocos`
3. Deploy (Framework: Vite)

### ResumeConverter
1. 동일 저장소 Import
2. **Root Directory:** `ResumeConverter`
3. Deploy (Framework: Flask)
