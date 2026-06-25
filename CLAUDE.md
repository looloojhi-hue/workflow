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

### React component tree
```
App  (activeMenu 상태로 화면 전환: 'dashboard' | 'new-report' | 'data-hub')
├── LeaderDashboard   — 팀장·실장·본부장 전용. 처리 대기 / 부서 현황 / 공유 보고서 3분류
├── EmployeeDashboard — 팀원 전용. 내 보고서(90일 이내) + 공유받은 보고서
├── NewReportForm     — 보고서 등록. 파일업로드/URL 탭, 연관 보고서 모달, 공유 조직도 모달 포함
├── ReportDetailView  — 상세·결재 화면. AI 요약 수정, To-Do 관리, 승인/반려/재승인, 프로젝트 히스토리
└── DashboardView     — 대시보드 탭. 내부(권한 필터) + 외부(개인) 대시보드 관리
```

### Data flow — report submission
```
파일 선택 → FileReader base64 인코딩 → submitReport()
  → Drive 업로드 → grantSilentPermission() (이메일 알림 없이 권한 부여)
  → callVertexAI() → 보고DB 시트 저장 → triggerReportNotification()
```

### Google Tasks OAuth (client-side)
To-Do의 "Task 등록" 버튼은 팝업 OAuth로 access_token을 획득해 `localStorage`에 저장한 뒤, `insertGoogleTaskServerSide(token, ...)` 를 통해 서버가 Tasks API를 대신 호출한다. OAuth Client ID와 Redirect URI가 `Index.html` 안에 하드코딩되어 있다.

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

## Hardcoded values (변경 시 주의)

코드 내에 직접 박혀 있는 값들. 환경이 바뀌거나 배포 URL이 바뀌면 두 파일 모두 수정해야 한다.

| 위치 | 값 | 설명 |
|---|---|---|
| Code.js:66 | `"1ayuewDy5BDH5qepAxHtkfP4NOBii1A5z"` | 업로드 폴더 ID |
| Code.js:1120 | `"https://script.google.com/a/..."` | 포털 배포 URL (이메일 링크에 사용) |
| Index.html:544 | `"81027032834-deoq6..."` | Tasks OAuth Client ID |
| Index.html:547 | `"https://script.google.com/a/..."` | OAuth Redirect URI |
| Code.js:51-55 | orgMap 객체 | 조직도 (권한관리 시트와 중복) |

## Known issues (알려진 버그)

- **`updateReportStatus` 컬럼 인덱스 오류** (Code.js:591-604): 결재 이메일 발송용 `reportObj` 생성 시 보고DB 컬럼 인덱스가 전부 틀려 있어, 이메일 본문에 엉뚱한 값이 출력됨.
- **`createNewFile` 미정의** (Code.js:712): `updateReportDocument`에서 기존 파일 URL 없을 때 호출하는 `createNewFile()` 함수가 코드베이스 어디에도 없음 → ReferenceError.
- **삭제된 보고서 미필터링** (`getMySubmissions`): `[보고서 삭제됨]` 상태 문서가 목록·연관 보고서 검색에 계속 노출됨.
- **읽음 처리 includes 오류** (Code.js:943): 이메일 부분 문자열 매칭으로 다른 사용자의 읽음 상태가 섞일 수 있음.

## Required Setup (Script Properties)
Set these in the Apps Script editor under **Project Settings → Script Properties**:

| Key | Value |
|---|---|
| `CLIENT_EMAIL` | Vertex AI 서비스 계정 이메일 |
| `PRIVATE_KEY` | 서비스 계정 비밀키 (`\n`을 줄바꿈으로 저장) |

## Advanced Services (appsscript.json)
The following Google APIs must be enabled in the Apps Script project:
- `AdminDirectory` (directory_v1)
- `Drive` (v3) — `grantSilentPermission` 및 파일 덮어쓰기에 필수
- `Tasks` (v1)
