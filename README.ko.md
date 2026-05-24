# Archivon DMS Backend

English version: [README.md](README.md)

Archivon은 문서 업로드, 실제 텍스트 추출, 자동 분류, 팀별 문서함, 검토 큐, 중복/버전 감지, 상세 분석 모달을 제공하는 로컬 개발용 문서관리 MVP입니다.

## 현재 구현 상태

- 프론트엔드: `dowple_dms.html` 단일 HTML 파일
- 백엔드: FastAPI + SQLite + `uploads/` 원본 파일 저장
- 지원 추출 형식: PDF, DOCX, TXT, HWPX, PPTX, XLSX, CSV
- fallback 형식: HWP 바이너리는 미지원, 이미지는 OCR 필요
- 팀별 문서함: 개발용 `team_id` 선택 방식
- 저장 구조: 백엔드 연결 시 SQLite DB 저장, 미연결 시 localStorage fallback
- 중복/버전 감지: SHA-256 exact duplicate 및 파일명/evidence 기반 possible version

## 지원 파일 형식

| 형식 | 상태 | 처리 방식 |
| --- | --- | --- |
| PDF | 지원 | PyMuPDF로 텍스트 레이어 추출 |
| DOCX | 지원 | python-docx로 paragraph/table cell 추출 |
| TXT | 지원 | utf-8-sig, utf-8, cp949, euc-kr 순서로 디코딩 |
| HWPX | 부분 지원 | zip 내부 XML/Preview 텍스트 추출 |
| PPTX | 지원 | python-pptx로 슬라이드 텍스트, 제목, shape, 표 cell 추출 |
| XLSX | 제한 지원 | openpyxl로 시트명, 헤더 후보, 상위 행, 주요 cell 추출 |
| CSV | 제한 지원 | 인코딩 자동 시도, 헤더 후보와 상위 행 추출 |
| HWP | 미지원 | 향후 전용 파서 또는 변환 모듈 필요 |
| 이미지 | OCR 필요 | 향후 OCR 연결 대상 |

## 주요 화면

- 대시보드: 현재 팀 기준 총 문서 수, 오늘 업로드, 자동 분류 저장, 검토 큐, 평균 confidence, 파일 형식/폴더별 통계, 최근 문서를 보여줍니다.
- 문서 업로드: 파일 업로드 후 텍스트 추출, 증거 패키지 생성, 자동 분류/사용자 확인/검토 필요 상태를 표시합니다.
- 문서 보관함: 팀별 문서 목록, 고급 검색/필터, 재분류, 상세 모달, 원본 보기/다운로드를 제공합니다.
- 검토 필요 큐: 애매한 문서를 pending 큐로 분리하고, 백엔드 연결 시 SQLite `review_queue`와 동기화합니다.
- 자연어 검색: “오늘 업로드한 PPTX”, “검토 필요한 문서”, “confidence 낮은 문서”, “재무 자료” 같은 표현을 조건으로 해석합니다.
- 용어 사전/분류 기준 관리: 팀별 분류 기준과 사용자 피드백 학습 현황을 관리합니다.
- 아카이브 제안: 90일 이상 미열람 추정 문서를 제안합니다.

## 기본 분류 폴더

현재 기본 분류 체계는 16개 폴더입니다.

1. 공고/양식
2. 사업계획서/수행계획서
3. 조사자료
4. 발표자료
5. 계약/정산
6. 기타
7. 보고서/리포트
8. 회의록/메모
9. 교육/학습자료
10. 인사/채용
11. 재무/회계
12. 구매/발주
13. 영업/고객관리
14. 마케팅/홍보
15. 기술/개발문서
16. 증명서/인증서

## 문서 상세 모달

문서 보관함이나 검색 결과에서 파일명을 클릭하면 상세 분석 모달이 열립니다.

- 기본 정보: 파일명, 파일 형식, 현재 분류 폴더, confidence, 상태, 저장일, 팀, DB 저장 여부
- 분류 결과 요약: 최종 분류, 처리 상태, 요약, 분류 이유
- Top 3 추천 후보: 후보 폴더, 점수, 신뢰도, 매칭 키워드, 문맥 단서, 피드백 반영 여부
- 점수 산정 근거: keywordScore, contextScore, nameScore, feedbackScore
- Evidence / 추출 텍스트: 추출 방식, extractionStatus, 페이지/슬라이드/시트 정보, 키워드 강조, 더 보기/접기
- 문서 이력/무결성: versionOfId, 중복 유형, file_hash 앞 12자리, 원본 파일 열람 가능 여부
- 현재 기준으로 재분류: 문서 1건만 재분류하고, backendId가 있으면 PATCH API로 DB에도 반영합니다.

## 로컬 실행

Windows CMD 기준:

```bat
cd backend
python -m venv .venv
.venv\Scripts\activate.bat
python -m pip install -r requirements.txt
python -m uvicorn main:app --host 127.0.0.1 --port 8001
```

간편 실행:

```bat
cd backend
run_local.bat
```

`run_local.bat`은 `.venv\Scripts\activate.bat`이 없으면 가상환경 생성/설치 안내를 출력하고 종료합니다. 가상환경이 준비되어 있으면 `127.0.0.1:8001`에서 FastAPI 서버를 실행합니다.

서버 확인:

```text
http://127.0.0.1:8001/health
http://127.0.0.1:8001/docs
```

프론트는 프로젝트 루트의 `dowple_dms.html`을 브라우저에서 열면 됩니다. 프론트는 `127.0.0.1:8000`, `127.0.0.1:8001` 순서로 백엔드 연결을 시도합니다.

## requirements.txt

현재 백엔드 필수 패키지:

- fastapi
- uvicorn
- python-multipart
- pymupdf
- sqlalchemy
- python-docx
- python-pptx
- openpyxl

더 자세한 백엔드 API, DB 위치, 배포 준비 메모는 [backend/README.md](backend/README.md)를 확인하세요.
