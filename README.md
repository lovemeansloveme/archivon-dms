# Archivon: AI-assisted Document Management System

Korean version: [README.ko.md](README.ko.md)

## Project overview

Archivon is a local MVP for an AI-assisted document management system. It extracts text from common office documents, builds a compact evidence package, classifies documents into business folders, stores metadata in SQLite, and keeps uploaded originals in a local uploads directory.

The Korean production-style prototype remains in `dowple_dms.html`. The English portfolio version is `dowple_dms_en.html`.

## Key features

- Single-file frontend prototype with dashboard, upload, repository, review queue, search, dictionary, rule management, and archive suggestions.
- FastAPI backend for real text extraction and SQLite persistence.
- Team workspace selector for local SaaS-style separation by `team_id`.
- Deterministic classification with folder keywords, context terms, semantic adjustment, and lightweight user-feedback scoring.
- Auto-classification, user-confirmation, and review-required routing based on confidence and score gaps.
- Duplicate and version detection using SHA-256 file hashes and text/name similarity.
- Document detail modal with classification explanation, Top 3 recommendations, score breakdown, evidence preview, file hash preview, reclassification, original-file view, and download.
- localStorage fallback when the backend is unavailable.

## Default classification folders

Archivon includes 16 default business folders: Notices / Forms, Business / Execution Plans, Research Materials, Presentation Decks, Contracts / Settlement, Other, Reports, Meeting Notes / Memos, Training / Learning Materials, HR / Recruiting, Finance / Accounting, Purchasing / Orders, Sales / CRM, Marketing / PR, Technical / Development Docs, and Certificates / Official Proofs.

## Supported file formats

| Format | Status | Extraction method |
| --- | --- | --- |
| PDF | Supported | PyMuPDF text extraction |
| DOCX | Supported | `python-docx` paragraphs and tables |
| TXT | Supported | Encoding fallback: UTF-8-SIG, UTF-8, CP949, EUC-KR |
| HWPX | Partial support | ZIP/XML text extraction |
| PPTX | Supported | `python-pptx` slide and table text extraction |
| XLSX | Limited support | `openpyxl` sheet names, headers, top rows, and key cell values |
| CSV | Limited support | Encoding detection, header extraction, and top rows |
| HWP | Not supported | Requires a future dedicated parser or conversion module |
| Images | OCR required | OCR is not implemented in this MVP |

## Architecture

- Frontend: `dowple_dms_en.html`, a standalone HTML/CSS/JavaScript dashboard.
- Backend: `backend/main.py`, a FastAPI server.
- Database: SQLite file `backend/dowple.db`.
- File storage: local `backend/uploads/` directory.
- Text extraction libraries: PyMuPDF, python-docx, python-pptx, openpyxl, and standard Python ZIP/XML/CSV utilities.
- Offline fallback: browser localStorage keeps documents, team settings, filters, dictionary terms, folder rules, review items, and lightweight learning examples.

## How to run locally

Windows CMD:

```bat
cd backend
python -m venv .venv
.venv\Scripts\activate.bat
python -m pip install -r requirements.txt
python -m uvicorn main:app --host 127.0.0.1 --port 8001
```

You can also use:

```bat
cd backend
run_local.bat
```

Then open `dowple_dms_en.html` in a browser. The frontend checks `http://127.0.0.1:8000` first and then `http://127.0.0.1:8001`.

Health checks:

- `http://127.0.0.1:8001/health`
- `http://127.0.0.1:8001/docs`

Expected health response:

```json
{"status":"ok"}
```

## Main APIs

- `GET /health`: backend health check.
- `POST /api/analyze`: upload a file and return an evidence package without saving it as a document.
- `POST /api/documents`: upload, extract, store original file, save metadata, and return duplicate/version candidates.
- `GET /api/documents?team_id=demo-team`: list documents for a team.
- `GET /api/documents/{document_id}?team_id=demo-team`: get one document.
- `PATCH /api/documents/{document_id}?team_id=demo-team`: update category, summary, confidence, evidence package, extraction status, and page count.
- `DELETE /api/documents/{document_id}?team_id=demo-team`: archive a document.
- `GET /api/documents/{document_id}/download?team_id=demo-team`: view or download the stored original file.
- `GET /api/folder-rules?team_id=demo-team`: load team-specific classification rules.
- `PUT /api/folder-rules?team_id=demo-team`: save team-specific classification rules.
- `POST /api/folder-rules/reset?team_id=demo-team`: reset team-specific rules to defaults.
- `GET /api/review-queue?team_id=demo-team`: list pending review items.
- `POST /api/review-queue?team_id=demo-team`: create a review queue item.
- `PATCH /api/review-queue/{review_id}?team_id=demo-team`: resolve or update a review queue item.
- `DELETE /api/review-queue/{review_id}?team_id=demo-team`: mark a review queue item as ignored.

## Dashboard, search, and detail modal

The dashboard summarizes the selected team's document count, today's uploads, auto-classified saves, user-confirmed saves, pending review items, average confidence, duplicate/version suspects, and learning examples. Folder and file-type charts link directly into repository filters.

The repository supports advanced filters by keyword, folder, file type, status, confidence range, and saved date. The natural-language search tab interprets simple phrases such as "uploaded today", "low confidence documents", "PPTX", "finance", and "duplicate/version suspected".

The detail modal explains the final category, confidence, processing state, Top 3 recommendations, score breakdown, extracted evidence, version/duplicate history, file hash preview, DB sync state, and original-file access. A single document can also be reclassified from the modal.

## Current MVP limitations

- This is a local development MVP, not a hosted SaaS product yet.
- Authentication and real organization permissions are not implemented; teams are selected with a development-only `team_id` dropdown.
- SQLite and local uploads are suitable for local testing, but not for durable multi-user production storage.
- Image OCR and scanned-PDF OCR are not implemented.
- Binary HWP parsing is not implemented.
- Classification is deterministic and heuristic-based; it is not a trained ML model.
- LLM/Gemini/OpenAI-based semantic classification is a future integration path, not part of the current backend.

## Future SaaS roadmap

- Deploy the frontend separately from the FastAPI backend.
- Move from SQLite to PostgreSQL.
- Move uploaded files from local disk to cloud storage such as S3, Supabase Storage, or Google Cloud Storage.
- Add authentication, organization membership, team roles, and permission checks.
- Add OCR for scanned PDFs and image documents.
- Add optional LLM-based evidence summarization and classification review.
- Add audit logs, admin analytics, retention policies, and production monitoring.
