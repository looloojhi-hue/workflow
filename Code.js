// =========================================================================
// 1. 초기 화면 및 프론트엔드 연동
// =========================================================================
function doGet(e) {
  return HtmlService.createTemplateFromFile('Index').evaluate()
    .setTitle('경영지원본부 통합 보고 포털')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// =========================================================================
// 2. [API] 폼 초기화 데이터 로드 (🚨 office, team 분리 버전 완벽 적용)
// =========================================================================
function getFormData() {
  const email = Session.getActiveUser().getEmail();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const authSheet = ss.getSheetByName('권한관리');
  const authData = authSheet.getDataRange().getValues();
  
  let isLeader = false;
  let userList = []; 
  
  // 권한관리 시트 순회 (A:성함, B:이메일, C:직책, D:본부, E:실, F:팀)
  for (let i = 1; i < authData.length; i++) {
    const rowName = authData[i][0];
    const rowEmail = authData[i][1];
    const role = authData[i][2];
    
    // 🚨 기존에 합쳐버리던 로직 대신, 명확하게 E열과 F열을 분리해서 변수에 담습니다.
    const rowOffice = authData[i][4] || ""; // E열: 소속실
    const rowTeam = authData[i][5] || "";   // F열: 소속팀
    
    // 이메일이 있는 모든 사용자를 수집
    if (rowEmail) {
      userList.push({ 
        name: rowName, 
        email: rowEmail, 
        role: role,
        // 🚨 프론트엔드로 전달할 때도 office와 team을 각각 분리해서 보내줍니다!
        office: rowOffice,
        team: rowTeam 
      });
    }
    
    // 접속자의 직책 확인 (본부장, 실장, 팀장이면 대시보드 접근 권한 부여)
    if (rowEmail === email && (role === '본부장' || role === '실장' || role === '팀장')) {
      isLeader = true;
    }
  }
  
  const orgMap = {
    "인사전략실": ["인사전략팀", "GHR팀", "인재개발팀"],
    "인사실": ["인사팀", "총무팀"],
    "ER실": ["ER전략팀"]
  };

  return { email: email, orgMap: orgMap, isLeader: isLeader, userList: userList };
}

// =========================================================================
// 3. [API] 보고서 제출 메인 로직 (파일, 권한, 스마트 전결 및 아카이브 통합)
// =========================================================================
function submitReport(formData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dbSheet = ss.getSheetByName('보고DB');
  const folderId = "1ayuewDy5BDH5qepAxHtkfP4NOBii1A5z"; // 보고서 폴더 ID
  
  const id = 'R-' + new Date().getTime();
  const timestamp = new Date();
  
  let finalFileUrl = formData.url || "";
  let targetBlob = null;

  // A. 파일 업로드 처리 (파일이 직접 넘어온 경우)
  if (formData.fileData) {
    const decodedFile = Utilities.base64Decode(formData.fileData.split(",")[1]);
    
    let extension = "";
    if (formData.fileName && formData.fileName.includes(".")) {
      extension = formData.fileName.substring(formData.fileName.lastIndexOf(".")); 
    }
    const safeTitle = formData.title.replace(/[\\/:*?"<>|]/g, " ");
    const newFileName = `[${formData.team}] ${safeTitle}_${formData.reportDate}_${formData.name}${extension}`;

    const blob = Utilities.newBlob(decodedFile, formData.fileContentType, newFileName);
    const folder = DriveApp.getFolderById(folderId);
    const file = folder.createFile(blob);
    
    file.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE);
    
    const accessEmails = new Set();
    accessEmails.add(formData.email); 
    if (formData.sharedUsers) {
      formData.sharedUsers.split(',').forEach(e => accessEmails.add(e.trim())); 
    }
    
    const authData = getAuthData();
    const superiors = getSuperiors(formData.office, formData.team, authData);
    superiors.forEach(e => accessEmails.add(e));

    // 🚨 [보안/알림 패치] 이메일 발송 없이 조용히 권한 부여
    const fileId = file.getId();
    accessEmails.forEach(email => {
      grantSilentPermission(fileId, email); 
    });

    finalFileUrl = file.getUrl();
    targetBlob = file.getBlob(); 
  } else if (formData.url) {
    // 🚨 [하이브리드 AI 패치] 유저가 웹페이지 HTML 소스코드를 주입한 경우 크롤링 장벽을 우회하여 다이렉트 덤프 분석을 실행합니다.
    if (formData.htmlSource && formData.htmlSource.trim() !== "") {
      targetBlob = Utilities.newBlob(formData.htmlSource, "text/html", "webpage.html");
    } else {
      targetBlob = extractTextFromUrl(formData.url);
    }
  }

  // B. AI 분석 진행
  let aiSummary = "";
  let aiKeywords = "";
  if (targetBlob && typeof targetBlob !== 'string') {
    const aiResult = callVertexAI(targetBlob);
    aiSummary = aiResult.summary;
    aiKeywords = aiResult.keywords;
  }

  // C. 접속자 직책 파악 (파일 업로드 분기에서 이미 로드한 authData 재사용, 없으면 로드)
  const authDataForRole = (typeof authData !== 'undefined') ? authData : getAuthData();
  let myRole = "팀원";
  for (let i = 1; i < authDataForRole.length; i++) {
    if (authDataForRole[i][1] === formData.email) {
      myRole = authDataForRole[i][2];
      break;
    }
  }

  // 🚨 [핵심 스마트 로직] 아카이브 및 전결 지정 상태 셋팅
  let initialStatus = "";
  let initialLog = "";

  if (formData.isArchiveMode) {
    initialStatus = `[${formData.targetRole} 승인 완료]`;
    initialLog = `${Utilities.formatDate(new Date(), "GMT+9", "yyyy-MM-dd HH:mm")} - ${formData.name}님이 과거 완료된 보고서를 보관함에 등록했습니다.`;
  } else {
    initialStatus = "[팀장 대기]";
    if (myRole === "팀장") initialStatus = "[실장 대기]";
    if (myRole === "실장") initialStatus = "[본부장 대기]";
    
    if (myRole === "본부장" || myRole === formData.targetRole || (myRole === "실장" && formData.targetRole === "팀장")) {
      initialStatus = `[${formData.targetRole} 승인 완료]`;
    }

    initialLog = `${Utilities.formatDate(new Date(), "GMT+9", "yyyy-MM-dd HH:mm")} - ${formData.name}님 보고서 제출`;
  }

  // D. 시트 저장
  const lastRow = dbSheet.getLastRow() + 1;
  const rowData = [
    id, timestamp, formData.reportDate, formData.email, formData.name, 
    formData.office, formData.team, formData.title, finalFileUrl, 
    aiSummary, aiKeywords, formData.sharedUsers || "",
    initialStatus, initialLog, "[]", "",
    formData.targetRole, // Q열
    formData.reportType, // R열
    formData.linkedReports || "" // 🚨 [추가] S열: 프론트에서 넘겨준 연관 보고서 ID 문자열 (예: "R-123,R-456")
  ];
  dbSheet.getRange(lastRow, 1, 1, rowData.length).setValues([rowData]);

  // 🚨 [신규 추가] 신규 보고서 제출 시 최초 결재선 직책자 알림 메일 전송 망 연동
  try {
    // 과거 완료 보고서 보관 모드(isArchiveMode)가 아닐 때만 메일 알림 작동
    if (!formData.isArchiveMode) {
      const reportObj = {
        id: id,
        dept: formData.office,
        team: formData.team,
        author: formData.name,
        authorEmail: formData.email,
        title: formData.title,
        url: finalFileUrl,
        reportType: formData.reportType,
        aiSummary: aiSummary,
        status: initialStatus
      };

      // 초기 셋팅된 대기 단계에 맞춰 상위 직책자 맞춤형 라우팅 전송
      if (initialStatus === "[팀장 대기]") {
        triggerReportNotification('SUBMIT', reportObj, authDataForRole);
      } else if (initialStatus === "[실장 대기]") {
        triggerReportNotification('APPROVE_BY_TEAM', reportObj, authDataForRole);
      } else if (initialStatus === "[본부장 대기]") {
        triggerReportNotification('APPROVE_BY_ROOM', reportObj, authDataForRole);
      }
    }
  } catch (mailError) {
    // 메일 발송 오류가 나더라도 메인 보고서 등록 트랜잭션이 끊기지 않도록 로그 예외 방어막 구성
    Logger.log("신규 제출 알림 메일 엔진 구동 중 예외 발생 (프로세스 우회 유지): " + mailError.toString());
  }

  return { success: true, id: id };
}

// =========================================================================
// 4. [내부 유틸] URL에서 파일 데이터(Blob) 추출 (Vertex AI 직접 전송용)
// =========================================================================
function extractTextFromUrl(url) {
  try {
    // 1. URL에서 파일 ID 추출
    const match = url.match(/[-\w]{25,}/);
    if (!match) return "[ERROR] 유효한 구글 드라이브 링크가 아닙니다.";
    const fileId = match[0];

    // 2. 파일 객체 가져오기
    const file = DriveApp.getFileById(fileId);
    
    // 🚨 중요: 텍스트 추출 대신 파일의 'Blob(원본 데이터)' 자체를 반환합니다.
    // Gemini 1.5/2.5 모델은 PDF/이미지를 직접 읽을 수 있습니다.
    return file.getBlob();

  } catch (e) {
    Logger.log("파일 로드 에러: " + e.message);
    return "[ERROR] " + e.message;
  }
}

// =========================================================================
// 5. [내부 유틸] GCP Vertex AI 호출 (강제 변환 제거 - PDF 전용 안내 롤백 버전)
// =========================================================================
function callVertexAI(blob) {
  const projectId = 'hr-division-ai-rpa';
  const location = 'us-central1';
  const modelId = 'gemini-2.5-flash'; // 2.5 버전 유지
  
  const accessToken = getAccessToken();
  const url = `https://${location}-aiplatform.googleapis.com/v1/projects/${projectId}/locations/${location}/publishers/google/models/${modelId}:generateContent`;

  const base64Data = Utilities.base64Encode(blob.getBytes());
  const mimeType = blob.getContentType();

  // 🚨 프롬프트 수정: 답변 형식을 더 엄격하게 제한 (생각 생략 유도)
  const systemInstruction = "당신은 전문 보고서 분석가입니다. 요약과 키워드 외의 불필요한 설명이나 마크다운 제목(##)은 절대 출력하지 마세요.";
  
  const payload = {
    "systemInstruction": { "parts": [{ "text": systemInstruction }] },
    "contents": [{
      "role": "user",
      "parts": [
        { "inline_data": { "mime_type": mimeType, "data": base64Data } },
        { "text": "첨부된 문서를 분석하여 반드시 아래 형식으로만 대답하세요.\n\n요약:\n1. (핵심 수치 포함 요약)\n2. (비용 증감 원인 요약)\n3. (향후 계획 또는 특이사항 요약)\n\n키워드: #키워드1 #키워드2 #키워드3" }
      ]
    }],
    "generationConfig": { 
      "temperature": 0.1,
      "maxOutputTokens": 4096, // 🚨 1024에서 4096으로 대폭 상향 (잘림 방지)
      "topP": 0.95
    }
  };

  const options = {
    "method": "post",
    "contentType": "application/json",
    "headers": { "Authorization": "Bearer " + accessToken },
    "payload": JSON.stringify(payload),
    "muteHttpExceptions": true
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const responseText = response.getContentText();
    const json = JSON.parse(responseText);
    
    Logger.log("🤖 AI 분석 완료");

    // 🚨 [핵심 수정] 여기서 영어 에러 대신 깔끔한 한국어 에러 메시지를 반환합니다.
    if (!json.candidates || !json.candidates[0].content) {
       return { summary: "[AI 분석 불가] PDF 형식의 파일만 자동 요약이 지원됩니다.", keywords: "" };
    }

    const resultText = json.candidates[0].content.parts[0].text;

    // 파싱 로직 (마크다운 제목 ## 등이 들어와도 잘 작동하게 보강)
    let summary = "";
    let keywords = "";
    
    const summaryMatch = resultText.match(/요약:?([\s\S]*?)(?=키워드:?|$)/i);
    const keywordMatch = resultText.match(/키워드:?([\s\S]*)$/i);
    
    // 파싱 실패 시에도 너무 긴 영어가 나오지 않게 처리
    summary = summaryMatch ? summaryMatch[1].replace(/[*#]/g, '').trim() : "요약 파싱 실패";
    keywords = keywordMatch ? keywordMatch[1].replace(/[*]/g, '').trim() : "";
    
    return { summary: summary, keywords: keywords };
    
  } catch (e) {
    return { summary: "[시스템 에러] " + e.message, keywords: "" };
  }
}

// =========================================================================
// 6. [인증] 서비스 계정 기반 토큰 생성 함수 (컨설턴트님 기존 코드)
// =========================================================================
function getAccessToken() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const clientEmail = scriptProperties.getProperty('CLIENT_EMAIL');
  const privateKey = scriptProperties.getProperty('PRIVATE_KEY');
  
  if (!clientEmail || !privateKey) {
    throw new Error("스크립트 속성에 CLIENT_EMAIL 또는 PRIVATE_KEY가 없습니다.");
  }
  
  const formattedPrivateKey = privateKey.replace(/\\n/g, '\n');

  const header = { alg: 'RS256', typ: 'JWT' };
  const now = Math.floor(Date.now() / 1000);
  const claimSet = {
    iss: clientEmail,
    scope: 'https://www.googleapis.com/auth/cloud-platform',
    aud: 'https://oauth2.googleapis.com/token',
    exp: now + 3600,
    iat: now
  };

  const toSign = Utilities.base64EncodeWebSafe(JSON.stringify(header)) + '.' +
                 Utilities.base64EncodeWebSafe(JSON.stringify(claimSet));
  const signatureBytes = Utilities.computeRsaSha256Signature(toSign, formattedPrivateKey);
  const signature = Utilities.base64EncodeWebSafe(signatureBytes);
  const jwt = toSign + '.' + signature;

  const params = {
    method: 'post',
    payload: { grant_type: 'urn:ietf:params:oauth:grant-type:jwt-bearer', assertion: jwt },
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch('https://oauth2.googleapis.com/token', params);
  const json = JSON.parse(response.getContentText());

  if (json.access_token) return json.access_token;
  else throw new Error('토큰 생성 실패: ' + JSON.stringify(json));
}

// =========================================================================
// 7. [관리자 전용] 최초 1회 권한 강제 승인용 함수
// =========================================================================
function forceTriggerAuth() {
  try {
    // 이 코드들은 실제로 뭘 하기 위함이 아니라, 
    // 구글 서버에 "나 드라이브랑 외부 통신 권한 필요해!"라고 신고해서 팝업을 띄우기 위한 미끼입니다.
    DriveApp.getFiles(); 
    Drive.Files.list({maxResults: 1}); 
    UrlFetchApp.fetch("https://www.google.com");
    Logger.log("권한 승인 완료!");
  } catch (e) {
    Logger.log("권한 승인 대기 중...");
  }
}

function testDriveAuth() {
  const testUrl = "https://drive.google.com/file/d/1YpFlyeaWJxx2JAz_M8893kGvhmx5dr2s/view?usp=drive_link"; 
  
  // 1. 파일 가져오기 (Blob 반환)
  const blob = extractTextFromUrl(testUrl);
  Logger.log("파일 로드 상태: " + blob); // 여기서 'Blob'이라고 뜨면 성공
  
  // 2. 가져온 파일을 AI에게 던지기
  const aiResult = callVertexAI(blob);
  
  // 3. AI의 최종 답변 확인
  Logger.log("--- AI 분석 결과 ---");
  Logger.log("요약: " + aiResult.summary);
  Logger.log("키워드: " + aiResult.keywords);
}

// =========================================================================
// 8. [API] 나의 제출 및 직책별 하위 부서 내역 전체 조회 로직 (데이터 누락 방지 통합 버전)
// =========================================================================
function getMySubmissions() {
  try {
    const email = Session.getActiveUser().getEmail().toLowerCase().trim();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dbSheet = ss.getSheetByName('보고DB');
    const authSheet = ss.getSheetByName('권한관리');
    
    const authData = authSheet.getDataRange().getValues();
    let myRole = "팀원";
    let myOffice = "";
    let myTeam = "";

    for (let i = 1; i < authData.length; i++) {
      if (String(authData[i][1]).toLowerCase().trim() === email) {
        myRole = String(authData[i][2] || "").trim() || "팀원";
        myOffice = String(authData[i][4] || "").trim();
        myTeam = String(authData[i][5] || "").trim();
        break;
      }
    }

    const myOfficeTeams = new Set();
    if (myRole === "실장") {
      for (let i = 1; i < authData.length; i++) {
        if (String(authData[i][4] || "").trim() === myOffice) {
          const t = String(authData[i][5] || "").trim();
          if (t) myOfficeTeams.add(t);
        }
      }
    }

    const data = dbSheet.getDataRange().getValues();
    const reports = [];
    
    for (let i = data.length - 1; i > 0; i--) {
      const row = data[i];
      if (!row[0]) continue;

      // 삭제된 보고서 제외
      if (String(row[12] || "") === "[보고서 삭제됨]") continue;

      const authorEmail = String(row[3] || "").toLowerCase().trim();
      const rawSharedUsers = String(row[11] || "").toLowerCase();
      const rOffice = String(row[5] || "").trim();
      const rTeam = String(row[6] || "").trim();

      let canView = false;
      if (authorEmail === email || rawSharedUsers.includes(email)) {
        canView = true; 
      } else if (myRole === "본부장") {
        canView = true; 
      } else if (myRole === "실장" && (myOfficeTeams.has(rTeam) || rOffice === myOffice)) {
        canView = true; 
      } else if (myRole === "팀장" && rTeam === myTeam) {
        canView = true; 
      }

      if (canView) {
        let formattedDate = "날짜 미상";
        let rawDateString = ""; 
        
        try {
          if (row[1]) {
            const dateObj = new Date(row[1]);
            if (!isNaN(dateObj.getTime())) {
              formattedDate = Utilities.formatDate(dateObj, Session.getScriptTimeZone(), "yyyy.MM.dd HH:mm");
              rawDateString = dateObj.toISOString();
            }
          }
        } catch (e) {}

        const status = String(row[12] || '상태 없음'); 
        let statusColor = 'bg-slate-100 text-slate-700 border-slate-200';
        if (status.includes('대기')) statusColor = 'bg-amber-100 text-amber-700 border-amber-200';
        else if (status.includes('반려')) statusColor = 'bg-red-100 text-red-700 border-red-200';
        else if (status.includes('승인')) statusColor = 'bg-emerald-100 text-emerald-700 border-emerald-200';
        else if (status.includes('검토중')) statusColor = 'bg-blue-100 text-blue-700 border-blue-200';

        let parsedTodos = [];
        try { parsedTodos = row[14] ? JSON.parse(row[14]) : []; } catch(e) { parsedTodos = []; }

        const readList = String(row[15] || "").toLowerCase().split(',').map(e => e.trim()).filter(Boolean);
        const isRead = readList.includes(email);

        // 🚨 보고 일자(C열) 포맷팅 로직 (주신 코드의 핵심 부분 유지)
        const reportDateRaw = row[2];
        let reportDateStr = "";
        if (reportDateRaw instanceof Date) {
          reportDateStr = Utilities.formatDate(reportDateRaw, "GMT+9", "yyyy-MM-dd");
        } else {
          reportDateStr = String(reportDateRaw || "");
        }

        reports.push({
          id: String(row[0]),
          rawDate: rawDateString,
          reportDate: reportDateStr, // 🚨 C열
          title: String(row[7] || ""),
          author: String(row[4] || ""),
          authorEmail: authorEmail,
          dept: rOffice,
          team: rTeam,
          date: formattedDate + " 제출",
          status: status,
          statusColor: statusColor,
          aiSummary: String(row[9] || ""),
          url: String(row[8] || ""),
          log: String(row[13] || ""),
          todos: parsedTodos, 
          isRead: isRead,
          // 🚨 [핵심 추가] R열(18번째 열, 인덱스 17)에서 보고 유형 데이터 로드
          reportType: String(row[17] || ""),
          // 🚨 [STEP 4 완결] S열(19번째 열, 인덱스 18)에서 연관 보고서 데이터 로드!
          linkedReports: String(row[18] || "") 
        });
      }
    }
    return reports;

  } catch (error) {
    // 🚨 주신 에러 핸들링 로직 그대로 유지
    return [{
      id: 'error-msg',
      rawDate: new Date().toISOString(),
      title: '🚨 시스템 데이터 로드 실패 (개발자 확인용)',
      author: '시스템 관리자',
      dept: 'System',
      team: 'Error',
      date: '오류 발생',
      status: '[에러]',
      statusColor: 'bg-red-100 text-red-700 border-red-200',
      aiSummary: '에러 상세: ' + error.toString() + ' (라인: ' + (error.lineNumber || '알수없음') + ')',
      url: '', log: '', todos: [], isRead: true
    }];
  }
}

// =========================================================================
// 9. [API] 보고서 상태 변경 및 결재 로직 (🚨 재승인 요청 및 완료 보고서 처리 로직 통합)
// =========================================================================
function updateReportStatus(reportId, action, comment) {
  const email = Session.getActiveUser().getEmail();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dbSheet = ss.getSheetByName('보고DB');
  const authSheet = ss.getSheetByName('권한관리');
  
  // 1. 접속자 정보 확인
  const authData = authSheet.getDataRange().getValues();
  let userInfo = null;
  for (let i = 1; i < authData.length; i++) {
    if (authData[i][1] === email) {
      userInfo = { name: authData[i][0], role: authData[i][2] };
      break;
    }
  }
  
  // 2. 해당 보고서 행 찾기
  const data = dbSheet.getDataRange().getValues();
  const rowIndex = findReportRow(data, reportId);

  const currentStatus = data[rowIndex-1][12]; // M열
  const currentLog = data[rowIndex-1][13];    // N열
  const targetRole = data[rowIndex-1][16];    // Q열: 최종 전결 대상
  const authorEmail = String(data[rowIndex-1][3] || ""); // D열: 기안자 이메일
  let nextStatus = currentStatus;

  // 3. 서버사이드 권한 검증
  const userRole = userInfo ? userInfo.role : "";
  if (action === 'APPROVE' || action === 'REJECT') {
    const isAuthorized =
      (currentStatus.includes("팀장")  && userRole === "팀장")  ||
      (currentStatus.includes("실장")  && userRole === "실장")  ||
      (currentStatus.includes("본부장") && userRole === "본부장");
    if (!isAuthorized) {
      throw new Error("결재 권한이 없습니다. 현재 결재 단계: " + currentStatus);
    }
  } else if (action === 'RESUBMIT' || action === 'SOFT_DELETE' || action === 'CHG_TO_ARCHIVE') {
    if (email.toLowerCase().trim() !== authorEmail.toLowerCase().trim()) {
      throw new Error("본인이 작성한 보고서에만 이 작업을 수행할 수 있습니다.");
    }
  } else if (action === 'FORCE_ARCHIVE') {
    const isLeader = (userRole === "팀장" || userRole === "실장" || userRole === "본부장");
    if (!isLeader) {
      throw new Error("팀장 이상의 리더 권한이 필요합니다.");
    }
  }

  if (action === 'APPROVE') {
    // 🚨 [핵심 로직] 현재 승인자가 최종 전결 대상인지 확인
    if (userInfo.role === targetRole) {
      nextStatus = `[${targetRole} 승인 완료]`;
    } else {
      // 다음 단계로 토스
      if (currentStatus.includes("팀장")) nextStatus = "[실장 대기]";
      else if (currentStatus.includes("실장")) nextStatus = "[본부장 대기]";
      else if (currentStatus.includes("본부장")) nextStatus = "[본부장 승인 완료]";
    }
  } else if (action === 'REJECT') {
    nextStatus = `[${userInfo.role} 반려]`;
  } else if (action === 'RESUBMIT') {
    // 재승인 시 다시 처음 단계부터 (targetRole 상관없이 팀장부터 혹은 본인 직책 다음부터)
    nextStatus = (userInfo.role === "팀장") ? "[실장 대기]" : 
                 (userInfo.role === "실장") ? "[본부장 대기]" : "[팀장 대기]";
  } else if (action === 'FORCE_ARCHIVE') {
    // 🚨 리더 권한으로 과거 오등록 문서를 즉시 최종 결재 완료 상태로 강제 전환
    nextStatus = `[${targetRole} 승인 완료]`;
  } else if (action === 'CHG_TO_ARCHIVE') {
    // 🚨 [보고자 패치] 최초 단계에서 보고자 자율로 최종 완료 처리 스위칭
    nextStatus = `[${targetRole} 승인 완료]`;
  } else if (action === 'SOFT_DELETE') {
    // 🚨 [논리 삭제 패치] 행 데이터 영구 유실 방지를 위해 상태값만 삭제 격리 처리
    nextStatus = `[보고서 삭제됨]`;
  }

  // 4. 로그 기록 생성
  const timeStr = Utilities.formatDate(new Date(), "GMT+9", "yyyy-MM-dd HH:mm");
  let actionText = "";
  if (action === 'APPROVE') actionText = '승인';
  else if (action === 'REJECT') actionText = '반려';
  else if (action === 'RESUBMIT') actionText = '재승인 요청';
  else if (action === 'FORCE_ARCHIVE') actionText = '완료 보고서 처리';
  else if (action === 'CHG_TO_ARCHIVE') actionText = '완료 보고서로 전환';
  else if (action === 'SOFT_DELETE') actionText = '보고서 삭제';

  // 🚨 접속자가 조직도에 없어서 userInfo가 null일 때를 고려해 폴백 가드라인 추가(기안자 본인 전용)
  const userDispName = userInfo ? `${userInfo.name}(${userInfo.role})` : "기안자 본인";
  const newLog = `${currentLog}\n${timeStr} - ${userDispName} ${actionText}: ${comment || '의견 없음'}`;

  // 5. 시트 업데이트
  dbSheet.getRange(rowIndex, 13).setValue(nextStatus); // M열 업데이트
  dbSheet.getRange(rowIndex, 14).setValue(newLog);    // N열 업데이트
  
  // 🚨 [신규 추가] 결재선 자동 이메일 알림 연동 인터셉터 망 구성
  try {
    const reportRow = data[rowIndex-1];
    
    // 이메일 템플릿용 실시간 데이터 바인딩 객체 조립 (보고DB 데이터 컬럼 매핑 스펙)
    const reportObj = {
      id: reportId,
      dept:        reportRow[5],   // F열: 소속실
      team:        reportRow[6],   // G열: 소속팀
      author:      reportRow[4],   // E열: 기안자 성명
      authorEmail: reportRow[3],   // D열: 기안자 이메일
      title:       reportRow[7],   // H열: 보고서 제목
      url:         reportRow[8],   // I열: 파일 URL
      reportType:  reportRow[17],  // R열: 보고 유형
      aiSummary:   reportRow[9],   // J열: AI 요약
      status: nextStatus
    };

    // 결재 액션 및 스위칭된 차기 상태코드 스니펫 판별 처리 (CHG_TO_ARCHIVE 및 SOFT_DELETE는 무음 처리하여 불필요한 직책자 알림 폭탄 방지)
    if (action === 'REJECT') {
      triggerReportNotification('REJECT', reportObj, authData);
    } else if (action === 'RESUBMIT') {
      triggerReportNotification('RESUBMIT', reportObj, authData);
    } else if (action === 'APPROVE') {
      if (nextStatus.includes("실장 대기")) {
        triggerReportNotification('APPROVE_BY_TEAM', reportObj, authData);
      } else if (nextStatus.includes("본부장 대기")) {
        triggerReportNotification('APPROVE_BY_ROOM', reportObj, authData);
      }
    }
  } catch (mailError) {
    // 혹시 모를 메일 API 예외 발생 시 시스템 결재 트랜잭션이 튕기지 않도록 로그 덤프 방어막 구성
    Logger.log("결재선 이메일 연동 처리 중 예외 발생 (무시하고 프로세스 진행): " + mailError.toString());
  }

  return { success: true, nextStatus: nextStatus };
}

// =========================================================================
// [내부 유틸] 신규 파일 생성 및 권한 부여 (기존 URL 없을 때 updateReportDocument에서 호출)
// =========================================================================
function createNewFile(blob, payload) {
  const folderId = "1ayuewDy5BDH5qepAxHtkfP4NOBii1A5z";
  const folder = DriveApp.getFolderById(folderId);
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE);

  const accessEmails = new Set();
  if (payload.email) accessEmails.add(payload.email);
  if (payload.sharedUsers) {
    payload.sharedUsers.split(',').forEach(e => accessEmails.add(e.trim()));
  }
  try {
    const superiors = getSuperiors(payload.office || "", payload.team || "");
    superiors.forEach(e => accessEmails.add(e));
  } catch (e) {
    Logger.log("createNewFile 권한 추출 실패: " + e.message);
  }

  const fileId = file.getId();
  accessEmails.forEach(email => grantSilentPermission(fileId, email));

  return file.getUrl();
}

// =========================================================================
// 10. [API] 제출된 보고서 파일/링크 업데이트 로직 (🚨 덮어쓰기 강제 & 날짜 픽스 & 권한 복구 & 연관보고서 추가)
// =========================================================================
function updateReportDocument(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dbSheet = ss.getSheetByName('보고DB');
  const data = dbSheet.getDataRange().getValues();
  
  const rowIndex = findReportRow(data, payload.id);
  const existingFileUrl = String(data[rowIndex-1][8] || "");

  // 🚨 [추가] 원본 보고서의 정보(이메일, 소속, 공유자 등)를 DB에서 가져옵니다.
  // (업데이트하는 사람이 원작자가 아닐 수 있기 때문에 기존 정보를 보존하여 권한을 재부여합니다)
  const originalEmail = String(data[rowIndex-1][3] || ""); // D열: 원작자 이메일
  const originalOffice = String(data[rowIndex-1][5] || ""); // F열: 원작자 실
  const originalTeam = String(data[rowIndex-1][6] || ""); // G열: 원작자 팀
  const originalShared = String(data[rowIndex-1][11] || ""); // L열: 공유 대상자

  let finalFileUrl = existingFileUrl; 
  let targetBlob = null;

  if (payload.fileData) {
    const decodedFile = Utilities.base64Decode(payload.fileData.split(",")[1]);
    
    let extension = "";
    if (payload.fileName && payload.fileName.includes(".")) {
      extension = payload.fileName.substring(payload.fileName.lastIndexOf(".")); 
    }
    const safeTitle = payload.title.replace(/[\\/:*?"<>|]/g, " ");
    
    // 🚨 프론트에서 넘어온 정확한 보고일자로 파일명 조합
    const newFileName = `[${payload.team}] ${safeTitle}_${payload.reportDate}_${payload.name}${extension}`;

    // 🚨 Drive API V3 통과를 위해 MimeType 명시적 강제 부여
    const mimeType = payload.fileContentType || 'application/pdf';
    const blob = Utilities.newBlob(decodedFile, mimeType, newFileName);
    targetBlob = blob; 

    // 안전한 정규식으로 고유 File ID 추출
    const match = existingFileUrl.match(/[-\w]{25,}/);
    
    if (existingFileUrl && match) {
      try {
        const fileId = match[0];
        
        // 1. 파일명 변경 (기본 DriveApp 사용이 가장 안전)
        DriveApp.getFileById(fileId).setName(newFileName);
        
        // 2. 내용 덮어쓰기 (Drive API V3)
        Drive.Files.update({ mimeType: mimeType }, fileId, blob, { supportsAllDrives: true });
        
        // 🚨 [추가] 덮어쓰기 성공 후, 누락되었던 권한을 완벽하게 다시 쏴줍니다!
        const accessEmails = new Set();
        if (originalEmail) accessEmails.add(originalEmail); // 원작자
        if (payload.email) accessEmails.add(payload.email); // 업데이트 한 사람
        if (originalShared) originalShared.split(',').forEach(e => accessEmails.add(e.trim())); // 기존 공유 대상자
        
        try {
          // 원래 소속을 기준으로 결재라인(팀장/실장 등) 다시 추출
          const superiors = getSuperiors(originalOffice, originalTeam);
          superiors.forEach(e => accessEmails.add(e));
        } catch (supErr) {
          Logger.log("상위권자 권한 추출 실패: " + supErr);
        }

        // 완성된 권한 리스트를 바탕으로 조용히 권한 부여 실행!
        accessEmails.forEach(email => {
          grantSilentPermission(fileId, email); 
        });

        finalFileUrl = existingFileUrl; 
      } catch (e) {
        // 🚨 실패 시 숨기지 않고 프론트로 에러를 던져 정확한 원인을 노출시킵니다!
        throw new Error("드라이브 덮어쓰기 실패 (권한 또는 API 문제): " + e.message); 
      }
    } else {
      finalFileUrl = createNewFile(blob, payload); 
    }

  } else if (payload.url) {
    finalFileUrl = payload.url;
    targetBlob = extractTextFromUrl(payload.url);
  }

  let aiSummary = "";
  let aiKeywords = "";
  if (targetBlob && typeof targetBlob !== 'string') {
    const aiResult = callVertexAI(targetBlob);
    aiSummary = aiResult.summary;
    aiKeywords = aiResult.keywords;
  }

  const timeStr = Utilities.formatDate(new Date(), "GMT+9", "yyyy-MM-dd HH:mm");
  const currentLog = data[rowIndex-1][13] || "";
  const newLog = `${currentLog}\n${timeStr} - ${payload.name}님이 첨부 문서를 최신 버전으로 업데이트했습니다.`;

  dbSheet.getRange(rowIndex, 9).setValue(finalFileUrl); // I열
  if (aiSummary) {
    dbSheet.getRange(rowIndex, 10).setValue(aiSummary); // J열
    dbSheet.getRange(rowIndex, 11).setValue(aiKeywords); // K열
  }
  dbSheet.getRange(rowIndex, 14).setValue(newLog); // N열
  
  // 🚨 [추가] S열(19번째 열): 연관 보고서 ID 업데이트 (데이터가 넘어왔을 때만)
  if (payload.linkedReports !== undefined) {
    dbSheet.getRange(rowIndex, 19).setValue(payload.linkedReports); 
  }
  
  return { success: true, newUrl: finalFileUrl, newSummary: aiSummary || data[rowIndex-1][9] };
}

// =========================================================================
// 11. [API] 보고서 요약 내용 수정 로직
// =========================================================================
function updateReportSummary(reportId, newSummary, modifierName) {
  const callerEmail = Session.getActiveUser().getEmail();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dbSheet = ss.getSheetByName('보고DB');
  const data = dbSheet.getDataRange().getValues();

  const rowIndex = findReportRow(data, reportId);

  // 작성자 본인만 AI 요약 수정 가능
  const reportAuthorEmail = String(data[rowIndex-1][3] || "");
  if (callerEmail.toLowerCase().trim() !== reportAuthorEmail.toLowerCase().trim()) {
    throw new Error("본인이 작성한 보고서의 AI 요약만 수정할 수 있습니다.");
  }

  // J열(10번째 열) 요약 내용 업데이트
  dbSheet.getRange(rowIndex, 10).setValue(newSummary);
  
  // N열(14번째 열) 결재 로그 업데이트
  const timeStr = Utilities.formatDate(new Date(), "GMT+9", "yyyy-MM-dd HH:mm");
  const currentLog = data[rowIndex-1][13] || "";
  const newLog = `${currentLog}\n${timeStr} - ${modifierName}님이 AI 요약 내용을 직접 수정했습니다.`;
  dbSheet.getRange(rowIndex, 14).setValue(newLog);

  return { success: true, newSummary: newSummary, newLog: newLog };
}

// =========================================================================
// 12. [API] 후속 조치(To-Do) 업데이트 로직 (신규)
// =========================================================================
function updateReportTodos(reportId, todosJSON) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dbSheet = ss.getSheetByName('보고DB');
  const data = dbSheet.getDataRange().getValues();
  const rowIndex = findReportRow(data, reportId);

  // O열(15번째 열) To-Do 리스트 업데이트
  dbSheet.getRange(rowIndex, 15).setValue(todosJSON);

  return { success: true };
}

// =========================================================================
// 13. [API] 새 대시보드 등록 로직
// =========================================================================
function submitDashboard(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('대시보드DB');
  
  const id = 'D-' + new Date().getTime();
  const timestamp = new Date();
  
  // payload: { email, type: '내부'|'외부', source, title, desc, url }
  // 시트 구조: ID(A), 등록일(B), 등록자(C), 부서구분(D), 팀명/출처(E), 대시보드명(F), 설명(G), 링크(H)
  const rowData = [
    id, 
    timestamp, 
    payload.email, 
    payload.type, 
    payload.source, 
    payload.title, 
    payload.desc, 
    payload.url
  ];
  
  sheet.appendRow(rowData);
  return { success: true, id: id };
}

// =========================================================================
// 14. [API] 권한별 대시보드 목록 조회 로직 (🚨 등록자 이름 동적 매핑 추가)
// =========================================================================
function getDashboards() {
  const email = Session.getActiveUser().getEmail();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. 접속자의 권한/소속 정보 가져오기 및 이름 매핑 객체 생성
  const authSheet = ss.getSheetByName('권한관리');
  const authData = authSheet.getDataRange().getValues();
  
  let myRole = "팀원";
  let myOffice = "";
  let myTeam = "";
  const emailToNameMap = {}; // 🚨 [신규 추가] 이메일 -> 이름 매핑 저장소
  
  for (let i = 1; i < authData.length; i++) {
    const rowName = authData[i][0];
    const rowEmail = authData[i][1];

    if (rowEmail) {
      emailToNameMap[rowEmail] = rowName; // 🚨 [신규 추가] 맵핑 데이터 구축
    }

    if (rowEmail === email) {
      myRole = authData[i][2];   // 직책 (본부장, 실장, 팀장, 팀원)
      myOffice = authData[i][4]; // 소속 실
      myTeam = authData[i][5];   // 소속 팀
      // 💡 내 정보는 찾았지만, 전체 이름 맵핑을 위해 break 하지 않고 끝까지 순회합니다.
    }
  }
  
  // 2. 소속 실(Office)에 속한 모든 하위 팀 목록 구하기 (실장님 뷰업용)
  const myOfficeTeams = new Set();
  if (myRole === "실장") {
    for (let i = 1; i < authData.length; i++) {
      if (authData[i][4] === myOffice && authData[i][5]) {
        myOfficeTeams.add(authData[i][5]);
      }
    }
  }

  // 3. 대시보드 데이터 읽기 및 권한별 필터링
  const dbSheet = ss.getSheetByName('대시보드DB');
  if (dbSheet.getLastRow() < 2) return []; // 데이터가 없으면 빈 배열 반환
  
  const data = dbSheet.getRange(2, 1, dbSheet.getLastRow() - 1, 8).getValues();
  const dashboards = [];
  
  for (let i = data.length - 1; i >= 0; i--) { // 최신 등록순으로 역순 조회
    const row = data[i];
    const dbId = row[0];
    const dbCreator = row[2]; // 등록자 이메일
    const dbType = row[3];    // 부서구분 ('내부' or '외부')
    const dbSource = row[4];  // 팀명/출처
    
    let canView = false;

    if (dbType === "내부") {
      // 💡 [내부 대시보드] 계층형 권한 로직
      if (myRole === "본부장") {
        canView = true; // 본부장: 전사(모든 팀) 대시보드 노출
      } else if (myRole === "실장") {
        // 실장: 본인 실(Office) 산하에 있는 팀들의 대시보드 노출
        if (myOfficeTeams.has(dbSource) || dbSource === myOffice) canView = true;
      } else {
        // 팀장/팀원: 본인 소속팀 대시보드만 노출
        if (dbSource === myTeam) canView = true;
      }
    } else {
      // 💡 [외부 연동 대시보드] 개인화 로직
      // 본인이 직접 등록한 외부 대시보드만 노출
      if (dbCreator === email) {
        canView = true;
      }
    }
    
    if (canView) {
      dashboards.push({
        id: dbId,
        type: dbType,   // '내부' or '외부'
        team: dbSource,
        title: row[5],
        desc: row[6],
        url: row[7],
        email: dbCreator,
        // 🚨 [신규 추가] 권한관리 시트에서 맵핑해둔 진짜 이름을 찾아서 넣어줍니다.
        creatorName: emailToNameMap[dbCreator] || dbCreator.split('@')[0], 
        isFavorite: false // 초기 즐겨찾기 상태 (향후 기능 확장 가능)
      });
    }
  }
  
  return dashboards;
}

// =========================================================================
// 15. [API] 보고서 읽음 처리 로직 (신규)
// =========================================================================
function markReportAsRead(reportId) {
  const email = Session.getActiveUser().getEmail();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dbSheet = ss.getSheetByName('보고DB');
  const data = dbSheet.getDataRange().getValues();

  let rowIndex;
  try { rowIndex = findReportRow(data, reportId); } catch (e) { return { success: false }; }

  // P열(16번째 열) 읽은 사람 이메일 업데이트
  let readUsers = data[rowIndex-1][15] || "";
  const alreadyRead = readUsers.split(',').map(e => e.trim()).filter(Boolean).includes(email);
  if (!alreadyRead) {
    readUsers = readUsers ? readUsers + "," + email : email;
    dbSheet.getRange(rowIndex, 16).setValue(readUsers);
  }
  
  return { success: true };
}

// =========================================================================
// 16. [API] 대시보드 삭제 로직 (신규)
// =========================================================================
function deleteDashboard(dashboardId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('대시보드DB'); // 🚨 등록 로직에 맞춰 시트명 정확히 연결
    
    // 데이터가 없으면 에러 처리
    if (sheet.getLastRow() < 2) {
      throw new Error("데이터가 없습니다.");
    }

    const data = sheet.getDataRange().getValues();
    
    // ID가 일치하는 행 찾아서 삭제 (역순으로 찾는 것이 안전하지만, 고유 ID이므로 순차 탐색)
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(dashboardId)) {
        sheet.deleteRow(i + 1); // 배열 인덱스는 0부터, 구글 시트 행은 1부터 시작하므로 +1
        return { success: true };
      }
    }
    
    throw new Error("삭제할 대시보드를 찾을 수 없습니다.");
  } catch (error) {
    throw new Error("삭제 중 오류가 발생했습니다: " + error.message);
  }
}

// =========================================================================
// 17. [API] 대시보드 정보 수정 로직 (신규)
// =========================================================================
function updateDashboard(payload) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('대시보드DB');
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(payload.id)) {
        // 시트 구조: ID(A), 등록일(B), 등록자(C), 부서구분(D), 팀명/출처(E), 대시보드명(F), 설명(G), 링크(H)
        sheet.getRange(i + 1, 4).setValue(payload.type);
        sheet.getRange(i + 1, 5).setValue(payload.source);
        sheet.getRange(i + 1, 6).setValue(payload.title);
        sheet.getRange(i + 1, 7).setValue(payload.desc);
        sheet.getRange(i + 1, 8).setValue(payload.url);
        return { success: true };
      }
    }
    throw new Error("수정할 대시보드를 찾을 수 없습니다.");
  } catch (error) {
    throw new Error("수정 중 오류 발생: " + error.message);
  }
}

// =========================================================================
// 18. [API] 상세 뷰에서 공유 대상자 추가
// =========================================================================
function addSharedUsersToReport(reportId, newSharedEmailsStr) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dbSheet = ss.getSheetByName('보고DB');
  const data = dbSheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === reportId) {
      const existingShared = String(data[i][11] || "").trim(); // L열
      const fileUrl = String(data[i][8] || ""); // I열 (파일 URL)
      
      // 1. 기존 공유자에 새 공유자 문자열 병합
      let updatedShared = existingShared;
      if (updatedShared && newSharedEmailsStr) {
        updatedShared += ", " + newSharedEmailsStr;
      } else if (newSharedEmailsStr) {
        updatedShared = newSharedEmailsStr;
      }
      
      // 2. DB L열(12번째) 업데이트
      dbSheet.getRange(i + 1, 12).setValue(updatedShared);
      
      // 3. 파일이 있으면 메일 알림 없이 조용히 권한 부여
      if (fileUrl.includes("drive.google.com/file/d/")) {
        const fileId = fileUrl.split("/d/")[1].split("/")[0];
        const newEmails = newSharedEmailsStr.split(",");
        newEmails.forEach(email => grantSilentPermission(fileId, email));
      }
      
      return { success: true };
    }
  }
  return { success: false, message: "보고서를 찾을 수 없습니다." };
}

// =========================================================================
// 19. [API] 피드백 및 개선 건의사항 접수 로직
// =========================================================================
function submitFeedback(payload) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('피드백DB');
    if (!sheet) throw new Error("피드백DB 시트가 존재하지 않습니다. 관리자에게 문의하세요.");

    const id = 'F-' + new Date().getTime();
    const timestamp = Utilities.formatDate(new Date(), "GMT+9", "yyyy-MM-dd HH:mm:ss");

    // 시트 구조: ID(A), 접수일시(B), 접수자(C), 이메일(D), 유형(E), 제목(F), 상세내용(G), 진행상태(H), 관리자메모(I)
    const rowData = [
      id,
      timestamp,
      payload.userName,
      payload.userEmail,
      payload.type,
      payload.title,
      payload.content,
      "[신규 접수]", // 초기 상태 자동 지정
      ""            // 관리자 메모 비워둠
    ];

    sheet.appendRow(rowData);
    return { success: true, id: id };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

// =========================================================================
// 20. [API] 구글 Tasks CORS 우회용 서버사이드 호출 엔진 (🚨 엔드포인트 안정성 패치 버전)
// =========================================================================
function insertGoogleTaskServerSide(token, title, notes) {
  try {
    // 🚨 [핵심 수정] 구글 서브도메인 라우팅 에러를 방지하기 위해 가장 안정적인 클래식 공식 주소로 교체합니다.
    const url = "https://www.googleapis.com/tasks/v1/lists/@default/tasks";
    
    const payload = { title: title, notes: notes };
    const options = {
      method: "POST",
      contentType: "application/json",
      headers: { "Authorization": "Bearer " + token },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };
    
    const response = UrlFetchApp.fetch(url, options);
    const responseText = response.getContentText();
    
    // 🚨 [방어 코드 추가] 구글 서버가 기습적으로 HTML 에러 페이지를 반환한 경우 파싱 전에 먼저 걸러냅니다.
    if (responseText.startsWith("<!DOCTYPE") || responseText.includes("<html")) {
      return { 
        success: false, 
        message: "구글 API 서버가 웹페이지(HTML)를 반환했습니다. (응답코드: " + response.getResponseCode() + "). GCP 콘솔에서 Tasks API가 활성화되어 있는지 확인이 필요할 수 있습니다." 
      };
    }
    
    const json = JSON.parse(responseText);
    
    if (json.id) {
      return { success: true };
    } else {
      return { success: false, message: json.error ? json.error.message : "구글 API 인증 처리 오류" };
    }
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

// =========================================================================
// 21. [API] 결재 워크플로우 단계별 이메일 자동 알림 전송 엔진 (🚨 실물 시트 인덱스 완벽 매칭본)
// =========================================================================
function triggerReportNotification(actionType, reportObj, authRows) {
  try {
    const PORTAL_URL = "https://script.google.com/a/macros/coway.com/s/AKfycbwI0kdNm0csBC-BLK8UYD7tdKqYX2zWtMkDgCndqa-ukYVsBv2oJ8215ZzUU9twAdx3YA/dev";

    const reportTypeStr = reportObj.reportType === 'written' ? '서면보고' : '대면보고';

    const userRows = authRows || getAuthData();
    
    let targetEmail = "";
    let emailSubject = "";
    let htmlBody = "";
    
    // ---------------------------------------------------------------------
    // [대상자 추출 로직] 실물 시트 구조 기반 이메일 라우팅 조립 (A=0, B=1, C=2...)
    // ---------------------------------------------------------------------
    if (actionType === 'REJECT') {
      targetEmail = reportObj.authorEmail; // 반려 시 기안자 이메일 다이렉트 매핑
    } else if (actionType === 'SUBMIT' || actionType === 'RESUBMIT') {
      // 신규/재승인 요청 시: 소속팀(F열=5번 인덱스)이 같고 직책(C열=2번 인덱스)이 "팀장"인 인원 추출
      for (let i = 1; i < userRows.length; i++) {
        if (userRows[i][5] === reportObj.team && String(userRows[i][2]).includes("팀장")) { 
          targetEmail = userRows[i][1]; // B열(1번 인덱스): 이메일 주소 탈취
          break;
        }
      }
    } else if (actionType === 'APPROVE_BY_TEAM') {
      // 팀장 승인 시: 소속실(E열=4번 인덱스)이 같고 직책(C열=2번 인덱스)이 "실장"인 인원 추출
      for (let i = 1; i < userRows.length; i++) {
        if (userRows[i][4] === reportObj.dept && String(userRows[i][2]).includes("실장")) { 
          targetEmail = userRows[i][1];
          break;
        }
      }
    } else if (actionType === 'APPROVE_BY_ROOM') {
      // 실장 승인 시: 직책(C열=2번 인덱스)이 "본부장"인 인원 추출
      for (let i = 1; i < userRows.length; i++) {
        if (String(userRows[i][2]).includes("본부장")) {
          targetEmail = userRows[i][1];
          break;
        }
      }
    }

    if (!targetEmail) {
      Logger.log("🚨 [알림 패스] 수신 대상 직책자 이메일을 시트에서 찾지 못했습니다. 액션타입: " + actionType);
      return { success: false, message: "수신 대상자를 찾을 수 없습니다." };
    }

    // ---------------------------------------------------------------------
    // [템플릿 빌드] HTML 카드 레이아웃 스타일링
    // ---------------------------------------------------------------------
    if (actionType === 'REJECT') {
      emailSubject = `[반려 알림] ${reportTypeStr}_${reportObj.title}`;
      htmlBody = `
        <div style="font-family:'Pretendard',sans-serif; max-w:550px; margin:0 auto; padding:24px; border:1px solid #e2e8f0; border-radius:16px; background-color:#ffffff;">
          <h2 style="color:#EF4444; font-size:20px; font-weight:bold; margin-top:0;">🚫 보고서 반려 안내</h2>
          <p style="font-size:14px; color:#475569; line-height:1.6;">경영지원본부 보고 시스템에서 알려드립니다.</p>
          <p style="font-size:15px; color:#1e293b; font-weight:bold; background-color:#FEF2F2; padding:12px; border-radius:8px; border-left:4px solid #EF4444;">
            제출하신 <strong>[${reportObj.title}]</strong> 보고서가 반려되었습니다.
          </p>
          <p style="font-size:13px; color:#64748b; margin-bottom:24px;">포털 내 리더십 피드백 진행 이력을 확인하시어 수정 후 재승인(등록) 요청을 진행해 주세요.</p>
          <a href="${PORTAL_URL}" target="_blank" style="display:inline-block; padding:12px 24px; background-color:#4F46E5; color:#ffffff; font-weight:bold; font-size:14px; text-decoration:none; border-radius:8px;">스마트 보고 포털 바로가기</a>
        </div>`;
    } else {
      emailSubject = `[보고 알림] ${reportTypeStr}_${reportObj.title}`;
      htmlBody = `
        <div style="font-family:'Pretendard',sans-serif; max-w:550px; margin:0 auto; padding:24px; border:1px solid #e2e8f0; border-radius:16px; background-color:#ffffff;">
          <h2 style="color:#4F46E5; font-size:20px; font-weight:bold; margin-top:0;">📬 보고서 승인 요청 알림</h2>
          <p style="font-size:14px; color:#475569; line-height:1.6;">경영지원본부 보고 시스템에서 알려드립니다.</p>
          <p style="font-size:15px; color:#1e293b; margin-bottom:20px;">
            <strong>${reportObj.author} 님</strong>이 아래 보고서의 승인을 요청하였습니다.
          </p>
          <div style="background-color:#f8fafc; padding:16px; border-radius:12px; border:1px solid #edf2f7; margin-bottom:24px; font-size:14px; line-height:1.6; color:#334155;">
            <div style="margin-bottom:8px;"><strong>• 보고 제목:</strong> ${reportObj.title}</div>
            <div style="margin-bottom:8px;"><strong>• 보고 유형:</strong> ${reportTypeStr}</div>
            <div style="margin-bottom:12px;"><strong>• 보고 요약 (AI):</strong></div>
            <div style="background-color:#ffffff; padding:12px; border-radius:8px; border:1px solid #e2e8f0; font-size:13px; color:#475569; white-space:pre-wrap;">${reportObj.aiSummary || "요약 정보 없음"}</div>
          </div>
          <a href="${PORTAL_URL}" target="_blank" style="display:inline-block; padding:12px 24px; background-color:#4F46E5; color:#ffffff; font-weight:bold; font-size:14px; text-decoration:none; border-radius:8px;">스마트 보고 포털 바로가기</a>
        </div>`;
    }

    // 🚨 [종합 패치] MailApp엔진 가동 (name 속성을 주어 발신인 가독성을 극대화합니다.)
    MailApp.sendEmail({
      to: targetEmail,
      subject: emailSubject,
      htmlBody: htmlBody,
      name: "스마트 보고시스템" // 👈 수신함에 '개발자 이름' 대신 이 이름으로 예쁘게 박힙니다!
    });

    return { success: true };
  } catch (e) {
    Logger.log("메일 인프라 전송 오류: " + e.toString());
    return { success: false, message: e.toString() };
  }
}

// =========================================================================
// [내부 유틸] 권한관리 시트 데이터 로드 (중복 로드 방지용 공통 헬퍼)
// =========================================================================
function getAuthData() {
  return SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName('권한관리')
    .getDataRange().getValues();
}

// =========================================================================
// [내부 유틸] 보고DB에서 reportId로 행 번호 반환 (없으면 에러)
// =========================================================================
function findReportRow(data, reportId) {
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === reportId) return i + 1;
  }
  throw new Error("보고서를 찾을 수 없습니다. (ID: " + reportId + ")");
}

// =========================================================================
// [내부 유틸] 소속 정보 기반 상위 소속장 이메일 리스트 가져오기
// =========================================================================
function getSuperiors(office, team, authRows) {
  const data = authRows || getAuthData();
  const superiors = [];

  for (let i = 1; i < data.length; i++) {
    const email = data[i][1];
    const role = data[i][2];
    const rowOffice = data[i][4];
    const rowTeam = data[i][5];

    if (rowTeam === team && role === '팀장') superiors.push(email);
    if (rowOffice === office && role === '실장') superiors.push(email);
    if (role === '본부장') superiors.push(email);
  }
  return superiors;
}

// =========================================================================
// [내부 유틸] 이메일 알림 없이 조용히 드라이브 권한만 부여하는 함수 (Drive API V3 필요)
// =========================================================================
function grantSilentPermission(fileId, email) {
  if (!email || email.trim() === "") return;
  try {
    Drive.Permissions.create(
      { role: 'commenter', type: 'user', emailAddress: email.trim() },
      fileId,
      { sendNotificationEmail: false, supportsAllDrives: true } // 🚨 핵심: 이메일 끄기 + 공유 드라이브 지원
    );
    Logger.log("권한 부여 성공: " + email);
  } catch (e) {
    Logger.log("권한 부여 실패 (" + email + "): " + e.message);
  }
}

// =========================================================================
// 🚨 [신규 API] 상세 페이지 전용: 연관 보고서(Project Link) 업데이트 로직
// =========================================================================
function updateLinkedReports(reportId, linkedIdsStr) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dbSheet = ss.getSheetByName('보고DB');
    const data = dbSheet.getDataRange().getValues();
    const rowIndex = findReportRow(data, reportId);

    // 🚨 S열(19번째 열)에 새로운 연관 보고서 ID 문자열만 단독으로 덮어쓰기
    dbSheet.getRange(rowIndex, 19).setValue(linkedIdsStr || "");
    
    return { success: true, message: "연관 보고서가 성공적으로 업데이트되었습니다." };
    
  } catch (error) {
    Logger.log("연관 보고서 업데이트 실패: " + error.message);
    return { success: false, message: error.message };
  }
}

// =========================================================================
// 🚨 [디버깅 전용] 내 데이터 조회 테스트 함수
// =========================================================================
function DEBUG_TEST_GET_DATA() {
  try {
    const email = Session.getActiveUser().getEmail().toLowerCase().trim();
    Logger.log("1. 접속 계정 이메일: " + email);

    const authSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('권한관리');
    const authData = authSheet.getDataRange().getValues();
    
    let myRole = "팀원", myOffice = "", myTeam = "";
    for (let i = 1; i < authData.length; i++) {
      if (String(authData[i][1]).toLowerCase().trim() === email) {
        myRole = String(authData[i][2] || "").trim();
        myOffice = String(authData[i][4] || "").trim();
        myTeam = String(authData[i][5] || "").trim();
        break;
      }
    }
    Logger.log(`2. 파악된 내 권한: [직책: ${myRole}], [실: ${myOffice}], [팀: ${myTeam}]`);

    // 백엔드 함수 직접 호출해보기
    const reports = getMySubmissions();
    
    if (reports === null) {
      Logger.log("🚨 [치명적 에러] getMySubmissions 함수가 null을 반환했습니다!");
    } else {
      Logger.log("3. 최종 리턴된 데이터 타입: " + typeof reports + " / 길이: " + reports.length);
      if (reports.length > 0) {
        Logger.log("4. 첫 번째 보고서 샘플: " + JSON.stringify(reports[0]));
      }
    }
  } catch (e) {
    Logger.log("🚨 [디버그 테스트 중 에러 발생]: " + e.toString() + "\n" + e.stack);
  }
}
