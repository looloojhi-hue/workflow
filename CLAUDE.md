# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is a Google Apps Script (GAS) web application — **경영지원본부 통합 보고 포털** (Management Support Division Integrated Reporting Portal). It is deployed as a Google Workspace web app accessible only within the `coway.com` domain.

## Development Commands (CLASP)

This project uses [CLASP](https://github.com/google/clasp) for local development.

```bash
clasp login          # Authenticate with Google account
clasp push           # Push local files to Apps Script (the only "deploy" step for code changes)
clasp pull           # Pull the current remote state to local
clasp open           # Open the Apps Script project in the browser
clasp deploy         # Create a new versioned web app deployment
```

There is no build step. Changes to `Code.js` and `Index.html` take effect after `clasp push`.

To test backend functions, open the Apps Script editor (`clasp open`) and run functions directly from the editor's Run menu. Use `DEBUG_TEST_GET_DATA()` as a starting point for data flow debugging.

## Architecture

### Two-file structure
- **`Code.js`** — Backend: all server-side GAS functions. These are called from the frontend via `google.script.run`.
- **`Index.html`** — Frontend: a single-file React 18 SPA rendered by GAS's `HtmlService`. Uses Babel Standalone for JSX, Tailwind CSS (CDN), and SweetAlert2 for modals. All components are defined inline in this single file.

### Frontend → Backend communication
The frontend calls backend functions exclusively through the `google.script.run` bridge:
```js
google.script.run
  .withSuccessHandler(result => { ... })
  .withFailureHandler(error => { ... })
  .functionName(args);
```
There is no REST API — every backend call is a named GAS function.

### Data storage (Google Sheets)
All data lives in Google Sheets on the spreadsheet bound to this script. Sheet names and column layouts are hardcoded:

| Sheet | Key columns |
|---|---|
| **권한관리** | A=성함, B=이메일, C=직책, D=본부, E=소속실(office), F=소속팀(team) |
| **보고DB** | A=ID, B=타임스탬프, C=보고일자, D=이메일, E=성함, F=소속실, G=소속팀, H=제목, I=파일URL, J=AI요약, K=키워드, L=공유대상, M=상태, N=결재로그, O=To-do(JSON), P=읽음이메일, Q=전결대상직책, R=보고유형, S=연관보고서ID |
| **대시보드DB** | A=ID, B=등록일, C=등록자이메일, D=부서구분(내부/외부), E=팀명/출처, F=제목, G=설명, H=링크 |
| **피드백DB** | A=ID, B=접수일시, C=접수자, D=이메일, E=유형, F=제목, G=내용, H=진행상태, I=관리자메모 |

### AI integration (Vertex AI)
`callVertexAI(blob)` sends document blobs directly to Gemini 2.5 Flash on GCP project `hr-division-ai-rpa` (us-central1). Authentication uses a service account JWT built from script properties `CLIENT_EMAIL` and `PRIVATE_KEY` — these must exist in the Apps Script project's **Script Properties** (not environment variables).

### File storage (Google Drive)
Uploaded files go to the hardcoded folder `1ayuewDy5BDH5qepAxHtkfP4NOBii1A5z`. File permissions are granted silently (no notification email) using the Drive API v3 advanced service via `grantSilentPermission()`.

### Approval workflow
Report status flows through a fixed hierarchy: `[팀장 대기]` → `[실장 대기]` → `[본부장 대기]` → `[직책 승인 완료]`. The submitter's role determines the starting state (e.g., a 팀장 submitting bypasses the 팀장 step). Email notifications are sent via `triggerReportNotification()` at each transition using `MailApp`.

### Access control
Role-based visibility in `getMySubmissions()` and `getDashboards()`:
- **본부장**: sees all reports
- **실장**: sees reports from their 소속실(office) and all teams within it
- **팀장**: sees reports from their team only
- **팀원**: sees only their own submissions and reports they are shared on

## Required Setup (Script Properties)
Set these in the Apps Script editor under **Project Settings → Script Properties**:

| Key | Value |
|---|---|
| `CLIENT_EMAIL` | Service account email for Vertex AI |
| `PRIVATE_KEY` | Service account private key (with `\n` as newlines) |

## Advanced Services (appsscript.json)
The following Google APIs must be enabled in the Apps Script project:
- `AdminDirectory` (directory_v1)
- `Drive` (v3) — required for `grantSilentPermission` and file overwrites
- `Tasks` (v1)
