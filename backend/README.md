# Dowple DMS Backend

Dowple MVP는 단일 HTML 프론트엔드(`dowple_dms.html`)와 FastAPI 백엔드로 구성된 로컬 개발용 문서 관리 프로토타입입니다.

## 현재 기능

- PDF/DOCX/TXT/HWPX/PPTX 실제 텍스트 추출 및 XLSX/CSV 제한 지원
- HWP 바이너리 파일은 `unsupported`로 fallback
- 이미지 파일은 `ocr_required`로 fallback
- SQLite DB 문서 메타데이터 저장
- `uploads/` 폴더에 원본 파일 저장
- SHA-256 기반 완전 중복 감지
- 파일명/evidence 기반 possible version 감지
- `version_of_id` 기반 새 버전 관계 저장
- 개발용 `team_id` 선택 기반 팀별 문서함
- 팀별 분류 기준 DB 동기화
- 팀별 대시보드/통계 화면
- 원본 파일 보기/다운로드 API
- 백엔드 미연결 시 프론트 localStorage fallback

## 지원 파일 형식

| 형식 | 상태 | 처리 방식 |
| --- | --- | --- |
| PDF | 지원 | PyMuPDF 실제 텍스트 추출 |
| DOCX | 지원 | python-docx로 paragraph와 table cell 텍스트 추출 |
| TXT | 지원 | utf-8-sig, utf-8, cp949, euc-kr 순서로 디코딩 |
| HWPX | 부분 지원 | zip 내부 XML/Preview 텍스트 추출 |
| PPTX | 지원 | python-pptx로 슬라이드 텍스트, 제목, 일반 shape, 표 cell 텍스트 추출 |
| XLSX | 제한 지원 | openpyxl로 시트명, 헤더 후보, 상위 행, 표 cell 텍스트 추출 |
| CSV | 제한 지원 | 인코딩을 자동 시도하고 헤더 후보와 상위 행 텍스트 추출 |
| XLS | 미지원 | 구형 바이너리 Excel은 향후 전용 파서 또는 변환 모듈 필요 |
| HWP | 미지원 | 향후 전용 파서 또는 변환 모듈 필요 |
| 이미지 | OCR 필요 | 향후 OCR 연결 필요 |

PPTX는 이미지 안에 박힌 글자까지 OCR하지는 않습니다. 이미지형 슬라이드처럼 추출 가능한 텍스트가 부족한 경우 `no_text_layer`로 분기하며 OCR 처리가 필요합니다.
XLSX/CSV는 전체 표를 완전 분석하지 않고, 제한된 행/열 범위에서 시트명, 헤더 후보, 상위 행, 주요 셀 값을 evidence로 구성합니다.

## 기본 분류 폴더

회사 문서관리용 기본 분류 기준은 아래 16개 폴더로 구성됩니다. 백엔드가 연결되어 있으면 이 기준이 팀별 `folder_rules` 테이블에 seed되고, 기존 DB에 새 폴더 ID가 없으면 서버 시작 또는 분류 기준 조회 시 자동으로 추가됩니다.

| 폴더 | 주요 기준 |
| --- | --- |
| 공고/양식 | 모집 공고, 신청서, 제출서류, 접수기간, 양식/서식 |
| 사업계획서/수행계획서 | 사업 목표, 수행 방법, 추진 전략, 예산, 일정, 기대효과 |
| 조사자료 | 설문, 인터뷰, 통계, 시장분석, 사례조사, 리서치 결과 |
| 발표자료 | 슬라이드, 중간/최종 발표, 발표 순서, 핵심 메시지 |
| 계약/정산 | 계약서, 대금 지급, 정산, 계좌, 거래 조건 |
| 기타 | 단서가 부족하거나 기준이 불명확해 검토가 필요한 문서 |
| 보고서/리포트 | 서론, 본론, 결론, 분석 결과, 참고문헌, 고찰 |
| 회의록/메모 | 회의 안건, 참석자, 결정사항, 후속 조치, 메모 |
| 교육/학습자료 | 교육자료, 매뉴얼, 사용법, 연습문제, 예제, 온보딩 |
| 인사/채용 | 채용, 입사, 면접, 근로계약, 휴가, 인사평가, 급여 |
| 재무/회계 | 매출, 비용, 예산, 재무제표, 손익계산서, 회계 보고 |
| 구매/발주 | 구매요청, 발주서, 견적, 납품, 입고, 공급업체 |
| 영업/고객관리 | 고객 상담, 제안, 요구사항, 거래처, CRM, 영업 기회 |
| 마케팅/홍보 | 캠페인, 광고, 콘텐츠, SNS, 브랜드, 보도자료 |
| 기술/개발문서 | API, 요구사항 정의, 시스템 설계, DB, 테스트, 배포 |
| 증명서/인증서 | 사업자등록증, 재직증명서, 자격증, 수료증, 납세증명서, 공식 증빙 |

## 팀원용 로컬 실행 가이드

### 1. 준비 사항

- Python 3.10 이상 설치가 필요합니다.
- Windows에서는 CMD 또는 PowerShell을 사용할 수 있습니다.
- 아래 명령어는 Windows CMD 기준입니다.

### 2. 백엔드 실행

프로젝트 루트에서 backend 폴더로 이동합니다.

```bat
cd backend
```

가상환경을 생성합니다.

```bat
python -m venv .venv
```

가상환경을 활성화합니다.

```bat
.venv\Scripts\activate.bat
```

필요 패키지를 설치합니다.

```bat
python -m pip install -r requirements.txt
```

FastAPI 서버를 실행합니다.

```bat
python -m uvicorn main:app --host 127.0.0.1 --port 8001
```

기본 프론트엔드는 `8000`과 `8001`을 자동으로 확인합니다. 이미 `8000`을 쓰는 프로그램이 있으면 위 예시처럼 `8001`을 사용해도 됩니다.

### 3. 서버 실행 확인

브라우저에서 아래 주소를 엽니다.

```text
http://127.0.0.1:8001/health
```

정상 응답:

```json
{"status":"ok"}
```

FastAPI 자동 문서도 확인할 수 있습니다.

```text
http://127.0.0.1:8001/docs
```

### 4. 프론트엔드 실행

백엔드 서버를 켠 뒤 프로젝트 루트의 `dowple_dms.html` 파일을 브라우저에서 직접 엽니다.

백엔드가 연결되면 문서 업로드 시 실제 텍스트 추출, DB 저장, 원본 파일 다운로드 기능을 사용할 수 있습니다. 백엔드가 꺼져 있으면 UI는 계속 열리지만 브라우저 localStorage 중심의 로컬 모드로 동작합니다.

## 간편 실행 배치 파일

`backend/run_local.bat`을 실행하면 가상환경을 활성화하고 `127.0.0.1:8001` 포트로 서버를 실행합니다.

가상환경이 없으면 먼저 아래 명령을 실행하세요.

```bat
cd backend
python -m venv .venv
.venv\Scripts\activate.bat
python -m pip install -r requirements.txt
```

## 주요 API

- `GET /health`: 서버 상태 확인
- `POST /api/analyze`: 파일 텍스트 추출 및 evidence_package 생성
- `POST /api/documents`: 파일 업로드, 텍스트 추출, 원본 저장, DB 저장
- `GET /api/documents?team_id=demo-team`: 팀별 문서 목록 조회
- `GET /api/documents/{document_id}?team_id=demo-team`: 문서 상세 조회
- `PATCH /api/documents/{document_id}?team_id=demo-team`: 문서 분류 결과, 요약, confidence, evidence_package, extraction_status, page_count 수정
- `GET /api/documents/{document_id}/download?team_id=demo-team`: 저장된 원본 파일 보기/다운로드
- `DELETE /api/documents/{document_id}?team_id=demo-team`: 해당 팀 문서 archived 처리
- `GET /api/folder-rules?team_id=demo-team`: 해당 팀의 분류 기준 조회
- `PUT /api/folder-rules?team_id=demo-team`: 해당 팀의 분류 기준 전체 저장
- `POST /api/folder-rules/reset?team_id=demo-team`: 해당 팀의 분류 기준 기본값 복원
- `GET /api/review-queue?team_id=demo-team`: 해당 팀의 pending 검토 필요 큐 조회
- `POST /api/review-queue?team_id=demo-team`: 검토 필요 항목 생성
- `PATCH /api/review-queue/{review_id}?team_id=demo-team`: 검토 필요 항목 상태 변경
- `DELETE /api/review-queue/{review_id}?team_id=demo-team`: 검토 필요 항목 ignored 처리

## 저장 위치

- SQLite DB: `backend/dowple.db`
- 업로드 원본 파일: `backend/uploads/`

서버를 처음 실행하면 `dowple.db`가 자동 생성되고 기본 `folder_rules` 데이터가 함께 들어갑니다.

## 팀 문서함 구조

현재 MVP는 실제 로그인 대신 프론트엔드의 팀 선택 드롭다운에서 `team_id`를 선택하는 방식입니다.

- 기본 팀: `demo-team`
- 추가 개발용 팀: `finance-team`, `project-team`
- 같은 파일이라도 다른 `team_id`에서는 별도 문서로 취급됩니다.
- 중복 감지는 같은 `team_id` 안에서만 수행됩니다.
- SaaS 전환 시 `team_id`는 로그인 사용자, 조직, 권한 정보와 연결할 수 있습니다.

## 팀별 분류 기준 관리

분류 기준 관리 화면에서 수정하는 폴더명, 설명, 대표 키워드, 문맥 단서는 백엔드가 연결되어 있으면 `folder_rules` 테이블에 팀별로 저장됩니다.

- `demo-team`, `finance-team`, `project-team`은 각각 독립적인 분류 기준을 가질 수 있습니다.
- 관리자가 키워드와 문맥 단서를 저장하면 이후 해당 팀의 업로드 분류에 반영됩니다.
- 기본값 복원은 현재 선택된 팀의 분류 기준만 초기화합니다.
- 백엔드가 꺼져 있으면 기존처럼 브라우저 localStorage 기준으로 동작하고, UI는 계속 사용할 수 있습니다.

## 대시보드/통계 화면

프론트 대시보드는 현재 선택된 `team_id` 기준으로 문서함과 검토 필요 큐 상태를 요약합니다. 백엔드가 연결되어 있으면 SQLite DB와 동기화된 현재 팀 문서 및 `review_queue` pending 항목을 기준으로 계산하고, 백엔드 미연결 상태에서는 localStorage fallback 데이터를 사용합니다.

- 주요 카드: 총 문서 수, 오늘 업로드, 자동 분류 저장, 사용자 확인 후 저장, pending 검토 큐, 평균 confidence, 중복/버전 의심, 학습 예시 수
- 폴더별 문서 수: 전체 분류 폴더별 건수와 Top 5 강조
- 파일 형식별 문서 수: PDF, DOCX, TXT, HWPX, PPTX, XLSX, CSV, 이미지, 기타
- 최근 업로드 문서: 최근 5개 문서와 상세 모달 연결
- 자동 분류 성과: 자동 분류 저장됨, 사용자 확인 후 저장됨, 검토 필요 비율과 평균 confidence

## 검색/필터 고도화

Dowple은 문서가 많아졌을 때 빠르게 찾을 수 있도록 문서 보관함과 자연어 검색 탭에서 확장 검색을 제공합니다.

- 문서 보관함 고급 필터: 검색어, 분류 폴더, 파일 형식, 상태, confidence 범위, 저장일 조건을 조합해 현재 팀 문서만 필터링합니다.
- 검색 대상: 파일명, 분류 폴더, 요약, `evidence_package`, 매칭 키워드, 맥락 단서, 상태, 파일 형식입니다.
- 자연어 검색 탭: “오늘 업로드한 문서”, “검토 필요한 문서”, “confidence 낮은 문서”, “엑셀 문서”, “발표자료”, “재무 자료”, “중복 문서” 같은 표현을 간단한 조건으로 해석합니다.
- 검색 결과 카드: 파일명, 분류 폴더, 파일 형식, confidence, 저장일, 상태, 검색어가 매칭된 위치, 짧은 요약을 보여주며 파일명 또는 상세 보기 버튼으로 문서 상세 모달을 엽니다.
- 대시보드 연동: 폴더별 문서 수와 파일 형식별 문서 수 항목을 클릭하면 문서 보관함으로 이동하고 해당 필터가 적용됩니다. 검토 필요 요약 버튼은 검토 필요 탭으로 이동합니다.
- 필터 상태는 `localStorage`에 저장되어 새로고침 후에도 유지되며, 필터 해제 버튼을 누르면 모든 조건이 초기화됩니다.

## 문서 상세 모달

Dowple의 문서 상세 모달은 단순 메타데이터 창이 아니라 AI-assisted document classification 결과를 설명하는 분석 화면으로 동작합니다.

- 상단 기본 정보: 파일명, 파일 형식, 현재 분류 폴더, confidence, 상태, 저장일, 현재 팀, DB 저장 여부, 원본 파일 열람 가능 여부
- 분류 결과 요약: 최종 분류 폴더, confidence, 자동 분류/사용자 확인/검토 필요 여부, 요약, 분류 이유
- Top 3 추천 후보: 후보 폴더명, 점수, 추천 신뢰도, 매칭 키워드, 문맥 단서, 추천 이유, 사용자 피드백 및 `feedbackScore` 반영 여부
- 점수 산정 근거: `keywordScore`, `contextScore`, `nameScore`, `feedbackScore`를 막대 형태로 표시
- Evidence / 추출 텍스트: 추출 방식, `extractionStatus`, 페이지/슬라이드/시트 정보, `evidence_package` 일부와 분류 키워드 강조, 더 보기/접기
- 문서 이력 및 무결성: 저장일, 업데이트일, `versionOfId`, 중복/버전 유사도, 재분류 반영 상태, `file_hash` 앞 12자리, 원본 파일 열람/다운로드 가능 여부
- 현재 기준으로 재분류: 상세 모달 안에서 해당 문서 1건만 재분류하고, `backendId`가 있으면 `PATCH /api/documents/{document_id}?team_id=...`로 DB에도 반영합니다. 백엔드 미연결 또는 로컬 문서는 localStorage fallback으로 반영됩니다.

## 기존 문서 재분류

문서 보관함의 `현재 기준으로 다시 분류` 버튼은 현재 팀 문서 중 `기타`, 낮은 confidence, 검토 필요 상태 문서를 최신 분류 기준으로 다시 평가합니다. `전체 문서 재분류` 버튼은 현재 팀의 모든 문서를 대상으로 같은 처리를 수행합니다.

`backendId`가 있는 문서는 `PATCH /api/documents/{document_id}?team_id=...` API로 DB의 `documents` 레코드에도 재분류 결과가 저장됩니다. 수정 대상은 분류 폴더, 요약, confidence, evidence_package, extraction_status, page_count입니다.

백엔드 미연결 상태이거나 `backendId`가 없는 로컬 문서는 기존처럼 프론트 localStorage fallback으로 반영됩니다. PATCH 실패 시에도 화면과 localStorage에는 재분류 결과를 유지하고, DB 업데이트 실패 안내를 표시합니다.

## 사용자 피드백 기반 경량 학습 구조

Dowple은 별도 머신러닝 모델을 학습시키는 단계는 아니지만, 사용자가 최종 선택한 분류 결과를 `learnedExamples`에 저장하고 이후 유사 문서의 분류 점수에 `feedbackScore`로 반영합니다. 이를 통해 human-in-the-loop 방식의 경량 학습 구조를 제공합니다.

- 학습 데이터는 현재 프론트 localStorage에 저장됩니다.
- 팀별로 학습 예시를 분리해 현재 팀 문서 분류에만 반영합니다.
- 용어 사전 화면의 학습 데이터 현황에서 누적 예시 수, 폴더별 예시 수, 최근 학습 문서를 확인할 수 있습니다.
- 학습 데이터 초기화는 `learnedExamples`만 비우며, 문서 보관함, 원본 파일, SQLite DB 문서는 삭제하지 않습니다.

## 자동 분류 / 사용자 확인 / 검토 필요 분기

Dowple은 분류 confidence와 후보 간 점수 차이를 기준으로 문서를 `자동 분류 가능`, `사용자 확인 필요`, `검토 필요` 상태로 나눕니다.

- 1위 후보가 80점 이상이고 2위와 10점 이상 차이나면 `자동 분류 가능`
- 1위 후보가 50점 이상 80점 미만이면 `사용자 확인 필요`
- 1위 후보가 50점 미만이거나, 1위와 2위 차이가 10점 미만이거나, 1위가 `기타`이면 `검토 필요`
- 자동 저장 모드를 켜면 `자동 분류 가능` 문서는 자동 저장되고, 애매한 문서는 사용자 확인 또는 검토 큐로 이동합니다.
- 중복/버전 후보가 발견된 문서는 자동 저장보다 사용자의 저장 방식 선택을 우선합니다.

## 검토 필요 큐 DB 동기화

Dowple은 자동 분류가 어려운 문서를 검토 필요 큐로 분리합니다. 백엔드가 연결되어 있으면 검토 큐 항목은 SQLite `review_queue` 테이블에 저장되고 `team_id` 기준으로 분리됩니다.

- 검토 필요 항목은 `GET /api/review-queue?team_id=...`로 현재 팀의 pending 항목만 불러옵니다.
- 업로드 결과가 검토 필요로 판단되면 프론트 localStorage에 먼저 표시하고, 백엔드 연결 상태에서는 `POST /api/review-queue?team_id=...`로 DB에도 저장합니다.
- 같은 팀에서 같은 파일명의 pending 항목이 이미 있으면 중복 생성하지 않고 기존 항목을 반환합니다.
- 담당자가 폴더를 확정해 저장하면 `PATCH /api/review-queue/{review_id}?team_id=...`로 `resolved` 상태와 최종 분류를 저장합니다.
- 백엔드 미연결 상태에서는 기존처럼 localStorage fallback으로 동작하며, 이후 서버 연결 시 현재 팀 기준 항목과 병합됩니다.

## Render 배포 준비 메모

Render에 백엔드를 올릴 때는 backend 폴더를 서비스 루트로 설정해야 할 수 있습니다.

권장 설정:

```text
Build Command: pip install -r requirements.txt
Start Command: uvicorn main:app --host 0.0.0.0 --port $PORT
```

주의 사항:

- Render의 임시 파일시스템에서는 SQLite DB와 `uploads/` 파일이 장기 보존되지 않을 수 있습니다.
- 장기 운영 시 Render Disk, PostgreSQL, S3/Supabase Storage/Google Cloud Storage 같은 영구 저장소가 필요합니다.
- 프론트의 `API_BASE_CANDIDATES`에 Render 백엔드 주소를 추가해야 합니다.

예:

```js
"https://dowple-api.onrender.com"
```

## Netlify/GitHub Pages/Vercel에 프론트만 배포할 때

`dowple_dms.html`만 정적 호스팅에 올리면 UI는 열립니다. 하지만 FastAPI 백엔드가 별도로 배포되어 있지 않으면 다음 기능은 동작하지 않거나 제한됩니다.

- 실제 텍스트 추출
- DB 저장
- 원본 파일 보기/다운로드
- 백엔드 기준 중복 검사
- 팀별 DB 문서함 동기화

프론트만 배포하는 경우에도 FastAPI 백엔드를 Render 등 별도 서비스로 배포하고, `dowple_dms.html`의 `API_BASE_CANDIDATES`에 배포 주소를 추가해야 합니다.

## 현재 MVP와 향후 SaaS 전환 계획

현재 MVP:

- 로컬 FastAPI 서버
- SQLite
- 로컬 `uploads/`
- 개발용 `team_id` 선택
- 단일 HTML 프론트엔드

향후 SaaS:

- 프론트 정적 배포
- 백엔드 클라우드 배포
- PostgreSQL 전환
- 클라우드 스토리지 전환
- 로그인/팀 권한 추가
- 운영용 모니터링과 백업 정책 추가
