/**
 * ============================================
 * G-IMPACT 리포트 생성기 v4.0
 * ============================================
 * 
 * 아키텍처:
 * - Google Apps Script: AI 변환 + 오케스트레이션
 * - Python 서버: 고품질 PDF 생성
 * 
 * 산출물:
 * - 요약 보고서 (15페이지, 디자인된 PDF)
 * - 상세 보고서 (50-100페이지)
 * 
 * 실행 방식:
 * - 백그라운드 실행 → 완료 시 이메일 알림
 */

// ============================================
// [1] 설정값
// ============================================

var REPORT_CONFIG_V4 = {
  version: "4.0",
  
  // 폴더 설정
  folderName: "G-IMPACT_분석리포트_v4",
  
  // Python PDF 서버 URL (배포 후 변경)
  pdfServerUrl: "https://your-server.com/api/generate-pdf",
  
  // Gemini API 설정
  geminiModel: "gemini-1.5-flash",
  geminiApiKey: "", // PropertiesService에서 가져옴
  
  // 이메일 설정
  emailSubject: "[G-IMPACT] 분석 리포트가 생성되었습니다",
  
  // 분석 단계 정의
  analysisSteps: [
    { id: "pestel", name: "PESTEL 분석", source: "step_2_1_pestel" },
    { id: "scenario", name: "시나리오 분석", source: "step_2_2_scenario" },
    { id: "competition", name: "경쟁환경 분석", source: "step_2_3_competition" },
    { id: "customer", name: "고객 분석", source: "step_2_4_customer" },
    { id: "market", name: "시장 분석", source: "step_2_5_market" },
    { id: "diagnosis", name: "경영진단", source: "step_3_1_diagnosis" },
    { id: "vrio", name: "VRIO 분석", source: "step_3_2_vrio" },
    { id: "swot", name: "SWOT 분석", source: "step_3_3_swot" },
    { id: "tows", name: "TOWS 전략", source: "step_3_4_tows" }
  ],
  
  // 리포트 타입
  reportTypes: {
    summary: { name: "요약 보고서", pages: "15페이지" },
    detail: { name: "상세 보고서", pages: "50-100페이지" }
  }
};

// ============================================
// [2] 메뉴 등록
// ============================================

/**
 * 스프레드시트 열 때 메뉴 추가
 */
function onOpen_ReportV4(e) {
  addReportMenuV4();
}

/**
 * 메뉴 생성
 */
function addReportMenuV4() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('📊 G-IMPACT 리포트 v4')
    .addItem('🚀 리포트 생성 (자동)', 'showReportDialogV4')
    .addItem('📂 리포트 폴더 열기', 'openReportFolderV4')
    .addSeparator()
    .addItem('⚙️ API 키 설정', 'showApiKeySettingV4')
    .addItem('🔧 서버 URL 설정', 'showServerUrlSettingV4')
    .addItem('📧 이메일 설정', 'showEmailSettingV4')
    .addSeparator()
    .addItem('📋 생성 이력 보기', 'showReportHistoryV4')
    .addItem('🗑️ 캐시 초기화', 'clearReportCacheV4')
    .addToUi();
}

// ============================================
// [3] 메인 다이얼로그
// ============================================

/**
 * 리포트 생성 다이얼로그 표시
 */
function showReportDialogV4() {
  var html = HtmlService.createHtmlOutput(buildReportDialogHtml_V4())
    .setWidth(600)
    .setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(html, 'G-IMPACT 분석 리포트 생성기 v4.0');
}

/**
 * 다이얼로그 HTML 생성
 */
function buildReportDialogHtml_V4() {
  return `
<!DOCTYPE html>
<html>
<head>
  <base target="_top">
  <style>
    * { box-sizing: border-box; font-family: 'Google Sans', sans-serif; }
    body { margin: 0; padding: 20px; background: #f8f9fa; }
    
    .header {
      background: linear-gradient(135deg, #1a73e8 0%, #0d47a1 100%);
      color: white;
      padding: 20px;
      border-radius: 12px;
      margin-bottom: 20px;
      text-align: center;
    }
    .header h2 { margin: 0 0 5px 0; }
    .header small { opacity: 0.9; }
    
    .card {
      background: white;
      padding: 20px;
      border-radius: 12px;
      box-shadow: 0 2px 8px rgba(0,0,0,0.1);
      margin-bottom: 15px;
    }
    
    .form-group { margin-bottom: 15px; }
    .form-group label {
      display: block;
      margin-bottom: 5px;
      font-weight: 500;
      color: #333;
    }
    .form-group select, .form-group input {
      width: 100%;
      padding: 12px;
      border: 1px solid #ddd;
      border-radius: 8px;
      font-size: 14px;
    }
    .form-group select:focus, .form-group input:focus {
      outline: none;
      border-color: #1a73e8;
      box-shadow: 0 0 0 3px rgba(26,115,232,0.1);
    }
    
    .checkbox-group {
      display: flex;
      gap: 20px;
      margin-top: 10px;
    }
    .checkbox-item {
      display: flex;
      align-items: center;
      gap: 8px;
      padding: 12px 16px;
      background: #f8f9fa;
      border-radius: 8px;
      cursor: pointer;
      flex: 1;
      border: 2px solid transparent;
      transition: all 0.2s;
    }
    .checkbox-item:hover { background: #e8f0fe; }
    .checkbox-item.selected {
      background: #e8f0fe;
      border-color: #1a73e8;
    }
    .checkbox-item input { display: none; }
    
    .btn {
      padding: 12px 24px;
      border: none;
      border-radius: 8px;
      cursor: pointer;
      font-size: 14px;
      font-weight: 500;
      transition: all 0.2s;
    }
    .btn-primary {
      background: #1a73e8;
      color: white;
      width: 100%;
      font-size: 16px;
      padding: 15px;
    }
    .btn-primary:hover { background: #1557b0; }
    .btn-primary:disabled {
      background: #ccc;
      cursor: not-allowed;
    }
    
    .progress-container {
      display: none;
      margin-top: 20px;
    }
    .progress-bar {
      height: 8px;
      background: #e0e0e0;
      border-radius: 4px;
      overflow: hidden;
    }
    .progress-fill {
      height: 100%;
      background: linear-gradient(90deg, #1a73e8, #34a853);
      width: 0%;
      transition: width 0.5s;
    }
    .progress-text {
      text-align: center;
      margin-top: 10px;
      color: #666;
      font-size: 14px;
    }
    
    .status {
      padding: 12px;
      border-radius: 8px;
      margin-top: 15px;
      display: none;
    }
    .status-info { background: #e3f2fd; color: #1565c0; }
    .status-success { background: #e8f5e9; color: #2e7d32; }
    .status-error { background: #ffebee; color: #c62828; }
    
    .step-indicator {
      display: flex;
      justify-content: space-between;
      margin-bottom: 20px;
      padding: 0 10px;
    }
    .step {
      display: flex;
      flex-direction: column;
      align-items: center;
      flex: 1;
    }
    .step-circle {
      width: 30px;
      height: 30px;
      border-radius: 50%;
      background: #e0e0e0;
      display: flex;
      align-items: center;
      justify-content: center;
      font-size: 12px;
      font-weight: bold;
      color: #666;
      margin-bottom: 5px;
    }
    .step-circle.active { background: #1a73e8; color: white; }
    .step-circle.done { background: #34a853; color: white; }
    .step-label { font-size: 11px; color: #666; text-align: center; }
    
    .info-box {
      background: #fff3e0;
      border-left: 4px solid #ff9800;
      padding: 12px;
      border-radius: 0 8px 8px 0;
      margin-top: 15px;
      font-size: 13px;
    }
  </style>
</head>
<body>
  <div class="header">
    <h2>📊 G-IMPACT 분석 리포트</h2>
    <small>AI 기반 자동 생성 시스템 v4.0</small>
  </div>
  
  <div class="card">
    <div class="step-indicator">
      <div class="step">
        <div class="step-circle active" id="step1">1</div>
        <div class="step-label">기업 선택</div>
      </div>
      <div class="step">
        <div class="step-circle" id="step2">2</div>
        <div class="step-label">AI 변환</div>
      </div>
      <div class="step">
        <div class="step-circle" id="step3">3</div>
        <div class="step-label">PDF 생성</div>
      </div>
      <div class="step">
        <div class="step-circle" id="step4">4</div>
        <div class="step-label">완료</div>
      </div>
    </div>
    
    <div class="form-group">
      <label>📌 기업 선택</label>
      <select id="bizSelect" onchange="onBizChange()">
        <option value="">-- 기업을 선택하세요 --</option>
      </select>
    </div>
    
    <div class="form-group">
      <label>🏢 비즈니스 모델</label>
      <select id="bmSelect">
        <option value="ALL">전체</option>
      </select>
    </div>
    
    <div class="form-group">
      <label>📄 리포트 유형</label>
      <div class="checkbox-group">
        <label class="checkbox-item selected" onclick="toggleReport(this, 'summary')">
          <input type="checkbox" id="chkSummary" checked>
          <span>📋 요약 보고서</span>
          <small>(15페이지)</small>
        </label>
        <label class="checkbox-item selected" onclick="toggleReport(this, 'detail')">
          <input type="checkbox" id="chkDetail" checked>
          <span>📚 상세 보고서</span>
          <small>(50-100p)</small>
        </label>
      </div>
    </div>
    
    <div class="form-group">
      <label>📧 결과 수신 이메일</label>
      <input type="email" id="emailInput" placeholder="example@company.com">
    </div>
    
    <div class="info-box">
      💡 <strong>안내:</strong> 리포트 생성에는 약 5-10분이 소요됩니다.<br>
      완료되면 입력하신 이메일로 다운로드 링크가 발송됩니다.
    </div>
    
    <div class="progress-container" id="progressContainer">
      <div class="progress-bar">
        <div class="progress-fill" id="progressFill"></div>
      </div>
      <div class="progress-text" id="progressText">준비 중...</div>
    </div>
    
    <div class="status" id="statusMsg"></div>
    
    <button class="btn btn-primary" id="startBtn" onclick="startGeneration()" disabled>
      🚀 리포트 생성 시작
    </button>
  </div>
  
  <script>
    // 초기화
    function init() {
      google.script.run
        .withSuccessHandler(function(list) {
          var select = document.getElementById('bizSelect');
          if (list && list.length > 0) {
            list.forEach(function(name) {
              var opt = document.createElement('option');
              opt.value = name;
              opt.text = name;
              select.add(opt);
            });
          } else {
            showStatus('기업 목록이 없습니다.', 'error');
          }
        })
        .withFailureHandler(function(e) {
          showStatus('기업 목록 로드 실패: ' + e.message, 'error');
        })
        .getBusinessListForReportV4();
      
      // 현재 사용자 이메일 설정
      google.script.run
        .withSuccessHandler(function(email) {
          document.getElementById('emailInput').value = email || '';
        })
        .getCurrentUserEmail();
    }
    
    // 기업 선택 시
    function onBizChange() {
      var biz = document.getElementById('bizSelect').value;
      document.getElementById('startBtn').disabled = !biz;
      
      if (biz) {
        google.script.run
          .withSuccessHandler(function(list) {
            var select = document.getElementById('bmSelect');
            select.innerHTML = '<option value="ALL">전체</option>';
            if (list && list.length > 0) {
              list.forEach(function(bm) {
                if (bm && bm !== 'ALL') {
                  var opt = document.createElement('option');
                  opt.value = bm;
                  opt.text = bm;
                  select.add(opt);
                }
              });
            }
          })
          .getCompanyBMListV4(biz);
      }
    }
    
    // 리포트 유형 토글
    function toggleReport(el, type) {
      el.classList.toggle('selected');
      var chk = el.querySelector('input');
      chk.checked = !chk.checked;
    }
    
    // 생성 시작
    function startGeneration() {
      var biz = document.getElementById('bizSelect').value;
      var bm = document.getElementById('bmSelect').value;
      var email = document.getElementById('emailInput').value;
      var genSummary = document.getElementById('chkSummary').checked;
      var genDetail = document.getElementById('chkDetail').checked;
      
      if (!biz) {
        showStatus('기업을 선택하세요.', 'error');
        return;
      }
      if (!email) {
        showStatus('이메일을 입력하세요.', 'error');
        return;
      }
      if (!genSummary && !genDetail) {
        showStatus('최소 하나의 리포트 유형을 선택하세요.', 'error');
        return;
      }
      
      // UI 업데이트
      document.getElementById('startBtn').disabled = true;
      document.getElementById('startBtn').textContent = '⏳ 생성 중...';
      document.getElementById('progressContainer').style.display = 'block';
      updateProgress(5, '데이터 수집 중...');
      setStepActive(1);
      
      // 백그라운드 실행 시작
      var params = {
        businessName: biz,
        bm: bm,
        email: email,
        generateSummary: genSummary,
        generateDetail: genDetail
      };
      
      google.script.run
        .withSuccessHandler(function(result) {
          updateProgress(100, '완료!');
          setStepDone(4);
          showStatus('✅ 리포트 생성이 시작되었습니다.<br>완료되면 ' + email + '로 알림이 발송됩니다.', 'success');
          document.getElementById('startBtn').textContent = '✓ 요청 완료';
        })
        .withFailureHandler(function(e) {
          showStatus('오류 발생: ' + e.message, 'error');
          document.getElementById('startBtn').disabled = false;
          document.getElementById('startBtn').textContent = '🚀 리포트 생성 시작';
        })
        .startReportGenerationV4(params);
      
      // 진행 상태 폴링
      pollProgress();
    }
    
    // 진행 상태 폴링
    function pollProgress() {
      var interval = setInterval(function() {
        google.script.run
          .withSuccessHandler(function(status) {
            if (status) {
              updateProgress(status.progress, status.message);
              if (status.step) setStepActive(status.step);
              if (status.completed) {
                clearInterval(interval);
                setStepDone(4);
              }
            }
          })
          .getReportProgressStatusV4();
      }, 3000);
    }
    
    function updateProgress(percent, text) {
      document.getElementById('progressFill').style.width = percent + '%';
      document.getElementById('progressText').textContent = text;
    }
    
    function setStepActive(step) {
      for (var i = 1; i <= 4; i++) {
        var el = document.getElementById('step' + i);
        if (i < step) el.className = 'step-circle done';
        else if (i === step) el.className = 'step-circle active';
        else el.className = 'step-circle';
      }
    }
    
    function setStepDone(step) {
      document.getElementById('step' + step).className = 'step-circle done';
    }
    
    function showStatus(msg, type) {
      var el = document.getElementById('statusMsg');
      el.innerHTML = msg;
      el.className = 'status status-' + type;
      el.style.display = 'block';
    }
    
    init();
  </script>
</body>
</html>
  `;
}

// ============================================
// [4] 유틸리티 함수
// ============================================

/**
 * 현재 사용자 이메일 가져오기
 */
function getCurrentUserEmail() {
  return Session.getActiveUser().getEmail();
}

/**
 * 기업 목록 조회
 */
function getBusinessListForReportV4() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("ANALYSIS_PROGRESS");
    
    if (!sheet) {
      Logger.log("ANALYSIS_PROGRESS 시트가 없습니다.");
      return [];
    }
    
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    
    // 기업명 컬럼 찾기
    var possibleHeaders = ["business_name", "기업명", "회사명", "기업"];
    var businessIdx = -1;
    for (var h = 0; h < possibleHeaders.length; h++) {
      var idx = headers.indexOf(possibleHeaders[h]);
      if (idx !== -1) { businessIdx = idx; break; }
    }
    if (businessIdx === -1) businessIdx = 0;
    
    // 유니크한 기업명 추출
    var businesses = {};
    for (var i = 1; i < data.length; i++) {
      var name = String(data[i][businessIdx] || "").trim();
      if (name) businesses[name] = true;
    }
    
    return Object.keys(businesses).sort();
  } catch (e) {
    Logger.log("기업 목록 조회 오류: " + e.message);
    return [];
  }
}

/**
 * BM 목록 조회
 */
function getCompanyBMListV4(businessName) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("ANALYSIS_PROGRESS");
    
    if (!sheet) return [];
    
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    
    // 컬럼 인덱스 찾기
    var businessIdx = Math.max(0, headers.indexOf("business_name"), headers.indexOf("기업명"));
    var bmIdx = headers.indexOf("bm");
    if (bmIdx === -1) bmIdx = headers.indexOf("BM");
    
    var bms = {};
    for (var i = 1; i < data.length; i++) {
      var name = normalizeCompanyNameV4(String(data[i][businessIdx] || ""));
      if (name === normalizeCompanyNameV4(businessName)) {
        var bm = String(data[i][bmIdx] || "").trim();
        if (bm) bms[bm] = true;
      }
    }
    
    return Object.keys(bms).sort();
  } catch (e) {
    Logger.log("BM 목록 조회 오류: " + e.message);
    return [];
  }
}

/**
 * 회사명 정규화
 */
function normalizeCompanyNameV4(name) {
  return String(name || "")
    .replace(/[\(\)（）]/g, "")
    .replace(/\s+/g, "")
    .replace(/주식회사|㈜/g, "")
    .trim()
    .toLowerCase();
}

/**
 * 리포트 폴더 열기
 */
function openReportFolderV4() {
  var folder = getOrCreateReportFolderV4();
  var html = HtmlService.createHtmlOutput(
    '<script>window.open("' + folder.getUrl() + '", "_blank");google.script.host.close();</script>'
  ).setWidth(100).setHeight(50);
  SpreadsheetApp.getUi().showModalDialog(html, '폴더 열기...');
}

/**
 * 리포트 폴더 가져오기/생성
 */
function getOrCreateReportFolderV4() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var parentFolder = DriveApp.getFileById(ss.getId()).getParents().next();
  
  var folderName = REPORT_CONFIG_V4.folderName;
  var folders = parentFolder.getFoldersByName(folderName);
  
  if (folders.hasNext()) {
    return folders.next();
  }
  return parentFolder.createFolder(folderName);
}

// ============================================
// [5] API 키 설정
// ============================================

/**
 * API 키 설정 다이얼로그
 */
function showApiKeySettingV4() {
  var currentKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY') || '';
  var maskedKey = currentKey ? currentKey.substring(0, 8) + '...' : '(설정되지 않음)';
  
  var ui = SpreadsheetApp.getUi();
  var result = ui.prompt(
    'Gemini API 키 설정',
    '현재: ' + maskedKey + '\n\n새 API 키를 입력하세요:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (result.getSelectedButton() === ui.Button.OK) {
    var newKey = result.getResponseText().trim();
    if (newKey) {
      PropertiesService.getScriptProperties().setProperty('GEMINI_API_KEY', newKey);
      ui.alert('API 키가 저장되었습니다.');
    }
  }
}

/**
 * 서버 URL 설정 다이얼로그
 */
function showServerUrlSettingV4() {
  var currentUrl = PropertiesService.getScriptProperties().getProperty('PDF_SERVER_URL') || REPORT_CONFIG_V4.pdfServerUrl;
  
  var ui = SpreadsheetApp.getUi();
  var result = ui.prompt(
    'PDF 서버 URL 설정',
    '현재: ' + currentUrl + '\n\n새 URL을 입력하세요:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (result.getSelectedButton() === ui.Button.OK) {
    var newUrl = result.getResponseText().trim();
    if (newUrl) {
      PropertiesService.getScriptProperties().setProperty('PDF_SERVER_URL', newUrl);
      ui.alert('서버 URL이 저장되었습니다.');
    }
  }
}

/**
 * 이메일 설정 다이얼로그
 */
function showEmailSettingV4() {
  var currentEmail = PropertiesService.getScriptProperties().getProperty('REPORT_EMAIL') || '';
  
  var ui = SpreadsheetApp.getUi();
  var result = ui.prompt(
    '기본 수신 이메일 설정',
    '현재: ' + (currentEmail || '(기본값: 현재 사용자)') + '\n\n이메일 주소를 입력하세요:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (result.getSelectedButton() === ui.Button.OK) {
    var newEmail = result.getResponseText().trim();
    PropertiesService.getScriptProperties().setProperty('REPORT_EMAIL', newEmail);
    ui.alert('이메일이 저장되었습니다.');
  }
}

// ============================================
// [6] 메인 생성 프로세스
// ============================================

/**
 * 리포트 생성 시작 (백그라운드)
 */
function startReportGenerationV4(params) {
  Logger.log("리포트 생성 시작: " + JSON.stringify(params));
  
  // 진행 상태 초기화
  var progressKey = "REPORT_PROGRESS_" + params.businessName;
  PropertiesService.getScriptProperties().setProperty(progressKey, JSON.stringify({
    step: 1,
    progress: 5,
    message: "데이터 수집 중...",
    completed: false,
    startedAt: new Date().toISOString()
  }));
  
  // 트리거 생성하여 백그라운드 실행
  var trigger = ScriptApp.newTrigger('executeReportGenerationV4')
    .timeBased()
    .after(1000) // 1초 후 실행
    .create();
  
  // 파라미터 저장
  PropertiesService.getScriptProperties().setProperty('REPORT_PARAMS_' + trigger.getUniqueId(), JSON.stringify(params));
  
  return { status: "started", triggerId: trigger.getUniqueId() };
}

/**
 * 리포트 생성 실행 (트리거에서 호출)
 */
function executeReportGenerationV4(e) {
  var triggerId = e.triggerUid;
  var paramsJson = PropertiesService.getScriptProperties().getProperty('REPORT_PARAMS_' + triggerId);
  
  if (!paramsJson) {
    Logger.log("파라미터를 찾을 수 없습니다.");
    return;
  }
  
  var params = JSON.parse(paramsJson);
  var progressKey = "REPORT_PROGRESS_" + params.businessName;
  
  try {
    // Step 1: 데이터 수집
    updateProgressV4(progressKey, 1, 10, "HANDOFF 데이터 수집 중...");
    var rawData = collectAllHandoffsV4(params.businessName, params.bm);
    
    // Step 2: AI 변환
    updateProgressV4(progressKey, 2, 20, "AI 분석 변환 중... (1/9)");
    var transformedData = transformAllWithAI_V4(rawData, params.businessName, progressKey);
    
    // Step 3: PDF 생성 요청
    updateProgressV4(progressKey, 3, 80, "PDF 생성 중...");
    var pdfResult = requestPdfGenerationV4(rawData, transformedData, params);
    
    // Step 4: 완료 및 이메일 발송
    updateProgressV4(progressKey, 4, 95, "이메일 발송 중...");
    sendCompletionEmailV4(params.email, params.businessName, pdfResult);
    
    updateProgressV4(progressKey, 4, 100, "완료!", true);
    
  } catch (error) {
    Logger.log("리포트 생성 오류: " + error.message);
    updateProgressV4(progressKey, 0, 0, "오류: " + error.message, true);
    
    // 오류 이메일 발송
    sendErrorEmailV4(params.email, params.businessName, error.message);
  } finally {
    // 트리거 삭제
    deleteTriggerV4(triggerId);
    // 파라미터 삭제
    PropertiesService.getScriptProperties().deleteProperty('REPORT_PARAMS_' + triggerId);
  }
}

/**
 * 진행 상태 업데이트
 */
function updateProgressV4(key, step, progress, message, completed) {
  PropertiesService.getScriptProperties().setProperty(key, JSON.stringify({
    step: step,
    progress: progress,
    message: message,
    completed: completed || false,
    updatedAt: new Date().toISOString()
  }));
}

/**
 * 진행 상태 조회
 */
function getReportProgressStatusV4() {
  // 가장 최근 진행 상태 반환
  var props = PropertiesService.getScriptProperties();
  var keys = props.getKeys();
  
  for (var i = 0; i < keys.length; i++) {
    if (keys[i].indexOf("REPORT_PROGRESS_") === 0) {
      try {
        return JSON.parse(props.getProperty(keys[i]));
      } catch (e) {}
    }
  }
  return null;
}

/**
 * 트리거 삭제
 */
function deleteTriggerV4(triggerId) {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getUniqueId() === triggerId) {
      ScriptApp.deleteTrigger(triggers[i]);
      break;
    }
  }
}

// ============================================
// [7] 데이터 수집
// ============================================

/**
 * 모든 HANDOFF 데이터 수집
 */
function collectAllHandoffsV4(businessName, bm) {
  Logger.log("데이터 수집 시작: " + businessName + " / " + bm);
  
  var data = {
    meta: {
      business_name: businessName,
      bm: bm,
      collected_at: new Date().toISOString(),
      version: REPORT_CONFIG_V4.version
    },
    handoffs: {}
  };
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("ANALYSIS_PROGRESS");
  
  if (!sheet) {
    throw new Error("ANALYSIS_PROGRESS 시트를 찾을 수 없습니다.");
  }
  
  var sheetData = sheet.getDataRange().getValues();
  var headers = sheetData[0];
  
  // 컬럼 매핑
  var handoffMapping = {
    "2.1_PESTEL": "step_2_1_pestel",
    "2.2_시나리오": "step_2_2_scenario",
    "2.3_경쟁환경": "step_2_3_competition",
    "2.4_고객분석": "step_2_4_customer",
    "2.5_시장분석": "step_2_5_market",
    "3.1_경영진단": "step_3_1_diagnosis",
    "3.2_VRIO": "step_3_2_vrio",
    "3.3_SWOT": "step_3_3_swot",
    "3.4_TOWS": "step_3_4_tows"
  };
  
  // 기업명/BM 컬럼 인덱스
  var businessIdx = findColumnIndex(headers, ["business_name", "기업명", "회사명"]);
  var bmIdx = findColumnIndex(headers, ["bm", "BM"]);
  
  // 각 HANDOFF 수집
  for (var sheetHeader in handoffMapping) {
    var internalKey = handoffMapping[sheetHeader];
    var stepIdx = headers.indexOf(sheetHeader);
    
    if (stepIdx === -1) continue;
    
    // 역순 검색 (최신 데이터 우선)
    for (var i = sheetData.length - 1; i >= 1; i--) {
      var row = sheetData[i];
      var rowBusiness = normalizeCompanyNameV4(String(row[businessIdx] || ""));
      var rowBm = String(row[bmIdx] || "ALL");
      
      if (rowBusiness === normalizeCompanyNameV4(businessName)) {
        if (bm === "ALL" || rowBm === "ALL" || rowBm === bm) {
          var handoffStr = row[stepIdx];
          if (handoffStr && String(handoffStr).trim() !== "") {
            try {
              data.handoffs[internalKey] = JSON.parse(handoffStr);
              Logger.log("HANDOFF 로드: " + internalKey);
            } catch (e) {
              Logger.log("JSON 파싱 실패 (" + internalKey + "): " + e.message);
            }
            break;
          }
        }
      }
    }
  }
  
  Logger.log("데이터 수집 완료 - 수집된 HANDOFF: " + Object.keys(data.handoffs).length);
  return data;
}

/**
 * 컬럼 인덱스 찾기 헬퍼
 */
function findColumnIndex(headers, possibleNames) {
  for (var i = 0; i < possibleNames.length; i++) {
    var idx = headers.indexOf(possibleNames[i]);
    if (idx !== -1) return idx;
  }
  return 0;
}

// ============================================
// [8] AI 변환 (Gemini API)
// ============================================

/**
 * 모든 섹션 AI 변환
 */
function transformAllWithAI_V4(rawData, businessName, progressKey) {
  var transformed = {
    sections: {},
    executiveSummary: null
  };
  
  var steps = REPORT_CONFIG_V4.analysisSteps;
  var totalSteps = steps.length + 1; // +1 for executive summary
  
  // 각 섹션 변환
  for (var i = 0; i < steps.length; i++) {
    var step = steps[i];
    var progress = 20 + Math.floor((i / totalSteps) * 50);
    updateProgressV4(progressKey, 2, progress, "AI 분석 변환 중... (" + (i+1) + "/" + steps.length + ") " + step.name);
    
    var handoffData = rawData.handoffs[step.source];
    if (handoffData) {
      try {
        transformed.sections[step.id] = transformSectionWithAI_V4(step.id, handoffData, businessName);
        Logger.log("AI 변환 완료: " + step.id);
      } catch (e) {
        Logger.log("AI 변환 실패 (" + step.id + "): " + e.message);
        transformed.sections[step.id] = { error: e.message, original: handoffData };
      }
    }
    
    // API 속도 제한 방지
    Utilities.sleep(1000);
  }
  
  // 경영진 요약 생성
  updateProgressV4(progressKey, 2, 75, "경영진 요약 생성 중...");
  try {
    transformed.executiveSummary = generateExecutiveSummaryV4(rawData, transformed.sections, businessName);
    Logger.log("경영진 요약 생성 완료");
  } catch (e) {
    Logger.log("경영진 요약 생성 실패: " + e.message);
  }
  
  return transformed;
}

/**
 * 개별 섹션 AI 변환
 */
function transformSectionWithAI_V4(sectionId, handoffData, businessName) {
  var prompt = buildSectionPromptV4(sectionId, handoffData, businessName);
  var response = callGeminiAPI_V4(prompt);
  
  return {
    content: response,
    generatedAt: new Date().toISOString()
  };
}

/**
 * 섹션별 프롬프트 빌더
 */
function buildSectionPromptV4(sectionId, data, businessName) {
  var baseInstruction = `당신은 15년 경력의 경영 컨설턴트입니다.
중소기업 CEO "${businessName}" 대표님께 분석 결과를 설명합니다.

[작성 원칙]
1. 전문용어는 반드시 쉬운 말로 풀어서 설명하세요
2. "그래서 우리 회사에 어떤 의미인가?"를 반드시 포함하세요
3. 구체적 숫자와 사례를 활용하세요
4. 권고사항은 실행 가능하게 작성하세요
5. 마크다운 형식으로 작성하세요 (##, ###, 표, 불릿 등)

`;

  var sectionPrompts = {
    pestel: `[분석 유형: PESTEL 거시환경 분석]

아래 JSON 데이터를 "고객 친화적 언어"로 변환하여 상세 보고서를 작성하세요.

[데이터]
${JSON.stringify(data, null, 2)}

[출력 형식]
## 1. 거시환경 분석 (PESTEL)

### 핵심 메시지
[CEO가 30초 안에 파악할 수 있는 핵심 3줄]

### 1.1 정치·정책 환경 (Political)
**현황**: [쉬운 설명]
**귀사에 미치는 영향**: [구체적 해석]
**대응 방향**: [실행 가능한 제안]

### 1.2 경제 환경 (Economic)
[동일 형식]

### 1.3 사회·문화 환경 (Social)
[동일 형식]

### 1.4 기술 환경 (Technological)
[동일 형식]

### 1.5 환경·생태 (Environmental)
[동일 형식]

### 1.6 법률·규제 (Legal)
[동일 형식]

### 종합 시사점
**핵심 기회 TOP 3**:
| 순위 | 기회 | 영향도 | 활용 방안 |
|------|------|--------|----------|

**핵심 위협 TOP 3**:
| 순위 | 위협 | 긴급도 | 대응 방안 |
|------|------|--------|----------|
`,

    scenario: `[분석 유형: 시나리오 분석]

아래 JSON 데이터를 "고객 친화적 언어"로 변환하여 상세 보고서를 작성하세요.

[데이터]
${JSON.stringify(data, null, 2)}

[출력 형식]
## 2. 미래 시나리오 분석

### 핵심 메시지
[CEO가 30초 안에 파악할 수 있는 핵심]

### 시나리오 개요
| 시나리오 | 발생확률 | 핵심 특징 | 귀사 영향 |
|----------|----------|----------|----------|

### 시나리오 1: [이름]
**상황 설명**: [쉬운 설명]
**귀사에 미치는 영향**: [구체적 해석]
**대응 전략**: [실행 가능한 제안]

[나머지 시나리오도 동일 형식]

### 강건한 전략 (어떤 시나리오에서도 유효)
- 전략 1: [설명]
- 전략 2: [설명]
- 전략 3: [설명]
`,

    competition: `[분석 유형: 경쟁환경 분석]

아래 JSON 데이터를 "고객 친화적 언어"로 변환하여 상세 보고서를 작성하세요.

[데이터]
${JSON.stringify(data, null, 2)}

[출력 형식]
## 3. 경쟁환경 분석

### 핵심 메시지
[CEO가 30초 안에 파악할 수 있는 핵심]

### 3.1 산업 경쟁 강도 (Five Forces)
| 요소 | 강도 | 의미 | 대응 방향 |
|------|------|------|----------|
| 신규 진입 위협 | | | |
| 기존 경쟁 강도 | | | |
| 대체재 위협 | | | |
| 공급자 교섭력 | | | |
| 구매자 교섭력 | | | |

**종합 평가**: [쉬운 설명]

### 3.2 주요 경쟁사 분석
| 경쟁사 | 강점 | 약점 | 위협 수준 | 대응 전략 |
|--------|------|------|----------|----------|

### 3.3 경쟁 포지셔닝
**귀사의 현재 위치**: [설명]
**목표 포지션**: [설명]
**이동 전략**: [구체적 방안]
`,

    customer: `[분석 유형: 고객 분석]

아래 JSON 데이터를 "고객 친화적 언어"로 변환하여 상세 보고서를 작성하세요.

[데이터]
${JSON.stringify(data, null, 2)}

[출력 형식]
## 4. 고객 분석

### 핵심 메시지
[CEO가 30초 안에 파악할 수 있는 핵심]

### 4.1 고객 생태계
**구매자 (Payer)**: [누가 돈을 내는가?]
**사용자 (User)**: [누가 실제로 사용하는가?]
**영향자 (Influencer)**: [구매 결정에 영향을 미치는 사람은?]

### 4.2 핵심 고객 세그먼트
| 세그먼트 | 특성 | 니즈 | 공략 전략 |
|----------|------|------|----------|

### 4.3 신규 발견 고객
[새롭게 발견된 잠재 고객에 대한 설명]

### 4.4 고객 확보 전략
**단기 (3개월)**: [구체적 액션]
**중기 (6개월)**: [구체적 액션]
**장기 (1년)**: [구체적 액션]
`,

    market: `[분석 유형: 시장 분석]

아래 JSON 데이터를 "고객 친화적 언어"로 변환하여 상세 보고서를 작성하세요.

[데이터]
${JSON.stringify(data, null, 2)}

[출력 형식]
## 5. 시장 분석

### 핵심 메시지
[CEO가 30초 안에 파악할 수 있는 핵심]

### 5.1 시장 규모
| 구분 | 규모 | 설명 |
|------|------|------|
| TAM (전체 시장) | 억원 | [쉬운 설명] |
| SAM (접근 가능 시장) | 억원 | [쉬운 설명] |
| SOM (1년 목표) | 억원 | [쉬운 설명] |

### 5.2 시장 성장성
**과거 성장률**: [데이터와 설명]
**향후 전망**: [데이터와 설명]
**성장 동인**: [핵심 요인 설명]

### 5.3 시장 진입 전략
**권장 진입 방식**: [구체적 방안]
**예상 소요 기간**: [기간]
**필요 투자 규모**: [금액 범위]
`,

    diagnosis: `[분석 유형: 경영진단]

아래 JSON 데이터를 "고객 친화적 언어"로 변환하여 상세 보고서를 작성하세요.

[데이터]
${JSON.stringify(data, null, 2)}

[출력 형식]
## 6. 경영진단

### 핵심 메시지
[CEO가 30초 안에 파악할 수 있는 핵심]

### 6.1 영역별 진단 결과
| 영역 | 점수 | 상태 | 핵심 이슈 | 개선 방향 |
|------|------|------|----------|----------|
| 사회적가치 | /5 | | | |
| 경영일반 | /5 | | | |
| 영업마케팅 | /5 | | | |
| 재무 | /5 | | | |
| 인사조직 | /5 | | | |

### 6.2 강점 영역 (잘하고 있는 것)
[구체적 설명과 유지/강화 방안]

### 6.3 개선 필요 영역
| 우선순위 | 영역 | 이슈 | 개선 방안 | 기대 효과 |
|----------|------|------|----------|----------|

### 6.4 즉시 실행 과제
1. [과제명] - 담당: [누구], 기한: [언제]
2. [과제명] - 담당: [누구], 기한: [언제]
`,

    vrio: `[분석 유형: VRIO 분석]

아래 JSON 데이터를 "고객 친화적 언어"로 변환하여 상세 보고서를 작성하세요.

VRIO란?
- V (Valuable): 가치 있는가?
- R (Rare): 희소한가?
- I (Inimitable): 모방하기 어려운가?
- O (Organized): 조직이 활용하고 있는가?

[데이터]
${JSON.stringify(data, null, 2)}

[출력 형식]
## 7. 핵심 자원 분석 (VRIO)

### 핵심 메시지
[CEO가 30초 안에 파악할 수 있는 핵심]

### 7.1 보유 자원 현황
| 자원 | 유형 | V | R | I | O | 경쟁우위 |
|------|------|---|---|---|---|----------|

### 7.2 지속적 경쟁우위 자원
[VRIO 모두 충족하는 자원에 대한 상세 설명]

### 7.3 개발 필요 자원
[부족한 자원과 확보 방안]

### 7.4 자원 투자 우선순위
| 순위 | 자원 | 현재 상태 | 투자 방향 | 예상 효과 |
|------|------|----------|----------|----------|
`,

    swot: `[분석 유형: SWOT 분석]

아래 JSON 데이터를 "고객 친화적 언어"로 변환하여 상세 보고서를 작성하세요.

[데이터]
${JSON.stringify(data, null, 2)}

[출력 형식]
## 8. SWOT 분석

### 핵심 메시지
[CEO가 30초 안에 파악할 수 있는 핵심]

### 8.1 SWOT 매트릭스

#### 강점 (Strengths) - 우리가 잘하는 것
| 항목 | 설명 | 활용 방안 |
|------|------|----------|

#### 약점 (Weaknesses) - 개선이 필요한 것
| 항목 | 설명 | 개선 방안 |
|------|------|----------|

#### 기회 (Opportunities) - 외부의 좋은 변화
| 항목 | 설명 | 포착 방안 |
|------|------|----------|

#### 위협 (Threats) - 외부의 나쁜 변화
| 항목 | 설명 | 대응 방안 |
|------|------|----------|

### 8.2 핵심 인사이트
1. [인사이트 1]
2. [인사이트 2]
3. [인사이트 3]
`,

    tows: `[분석 유형: TOWS 전략]

아래 JSON 데이터를 "고객 친화적 언어"로 변환하여 상세 보고서를 작성하세요.

TOWS란?
- SO전략: 강점으로 기회를 살린다 (공격)
- WO전략: 약점을 보완하며 기회를 잡는다 (전환)
- ST전략: 강점으로 위협을 막는다 (방어)
- WT전략: 약점과 위협을 최소화한다 (생존)

[데이터]
${JSON.stringify(data, null, 2)}

[출력 형식]
## 9. TOWS 전략

### 핵심 메시지
[CEO가 30초 안에 파악할 수 있는 핵심]

### 9.1 전략 옵션 매트릭스
| 유형 | 전략명 | 핵심 가설 | 점수 | 우선순위 |
|------|--------|----------|------|----------|

### 9.2 최종 선정 전략 TOP 3

#### 🥇 1순위: [전략명]
**전략 유형**: [SO/WO/ST/WT]
**핵심 내용**: [쉬운 설명]
**왜 이 전략인가?**: [선정 근거]
**실행 방안**:
- 단기 (3개월): [구체적 액션]
- 중기 (6개월): [구체적 액션]
**기대 효과**: [정량적/정성적]
**필요 자원**: [인력, 예산 등]

#### 🥈 2순위: [전략명]
[동일 형식]

#### 🥉 3순위: [전략명]
[동일 형식]

### 9.3 전략 실행 로드맵
| 단계 | 기간 | 핵심 전략 | 목표 |
|------|------|----------|------|
| Phase 1 | 0-6개월 | | |
| Phase 2 | 6-12개월 | | |
| Phase 3 | 1-2년 | | |

### 9.4 즉시 실행 과제
1. **[과제명]** - 담당: [누구], 기한: [언제]
2. **[과제명]** - 담당: [누구], 기한: [언제]
`
  };

  return baseInstruction + (sectionPrompts[sectionId] || "데이터를 분석하여 보고서를 작성하세요.\n\n" + JSON.stringify(data, null, 2));
}

/**
 * 경영진 요약 생성
 */
function generateExecutiveSummaryV4(rawData, transformedSections, businessName) {
  var prompt = `당신은 15년 경력의 경영 컨설턴트입니다.
"${businessName}" 대표님께 전체 분석 결과를 1페이지로 요약합니다.

[작성 원칙]
1. CEO가 3분 안에 핵심을 파악할 수 있도록 작성
2. 전문용어는 모두 쉬운 말로 변환
3. 숫자와 구체적 사례 활용
4. 즉시 실행 가능한 액션 아이템 포함

[분석 데이터 요약]
${JSON.stringify(rawData.handoffs, null, 2).substring(0, 15000)}

[출력 형식]
## 경영진 요약 (Executive Summary)

### 1. 핵심 결론 (30초 요약)
1. [가장 중요한 결론]
2. [두 번째 중요한 결론]
3. [세 번째 중요한 결론]

### 2. 외부환경 (기회 vs 위협)
**주요 기회**: [핵심 기회 요약]
**주요 위협**: [핵심 위협 요약]

### 3. 내부역량 (강점 vs 약점)
**핵심 강점**: [강점 요약]
**핵심 약점**: [약점 요약]

### 4. 전략 방향
**추천 전략**: [핵심 전략 1-2문장]

| 순위 | 전략 | 유형 | 핵심 근거 |
|------|------|------|----------|
| 1 | | | |
| 2 | | | |
| 3 | | | |

### 5. 90일 실행 계획
| 단계 | 기간 | 핵심 과제 | 담당 | 목표 |
|------|------|----------|------|------|
| Phase 1 | 0-30일 | | | |
| Phase 2 | 30-60일 | | | |
| Phase 3 | 60-90일 | | | |

### 6. 핵심 리스크 및 대응
| 리스크 | 발생확률 | 영향도 | 예방 조치 |
|--------|----------|--------|----------|
`;

  return callGeminiAPI_V4(prompt);
}

/**
 * Gemini API 호출
 */
function callGeminiAPI_V4(prompt) {
  var apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  
  if (!apiKey) {
    throw new Error("Gemini API 키가 설정되지 않았습니다. 메뉴 > API 키 설정에서 설정하세요.");
  }
  
  var url = "https://generativelanguage.googleapis.com/v1beta/models/" + 
            REPORT_CONFIG_V4.geminiModel + ":generateContent?key=" + apiKey;
  
  var payload = {
    contents: [{
      parts: [{
        text: prompt
      }]
    }],
    generationConfig: {
      temperature: 0.7,
      maxOutputTokens: 8192
    }
  };
  
  var options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  
  var response = UrlFetchApp.fetch(url, options);
  var result = JSON.parse(response.getContentText());
  
  if (result.error) {
    throw new Error("Gemini API 오류: " + result.error.message);
  }
  
  if (result.candidates && result.candidates[0] && result.candidates[0].content) {
    return result.candidates[0].content.parts[0].text;
  }
  
  throw new Error("Gemini API 응답을 파싱할 수 없습니다.");
}


// ============================================
// [9] PDF 서버 연동
// ============================================

/**
 * Python PDF 서버에 생성 요청
 */
function requestPdfGenerationV4(rawData, transformedData, params) {
  var serverUrl = PropertiesService.getScriptProperties().getProperty('PDF_SERVER_URL') || REPORT_CONFIG_V4.pdfServerUrl;
  
  var payload = {
    // 메타 정보
    meta: rawData.meta,
    
    // 원본 HANDOFF 데이터 (차트 생성용)
    handoffs: rawData.handoffs,
    
    // AI 변환된 텍스트 (보고서 본문용)
    transformed: transformedData,
    
    // 생성 옵션
    options: {
      generateSummary: params.generateSummary,
      generateDetail: params.generateDetail,
      businessName: params.businessName,
      bm: params.bm
    }
  };
  
  var options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
    timeout: 300 // 5분 타임아웃
  };
  
  try {
    var response = UrlFetchApp.fetch(serverUrl + "/generate", options);
    var result = JSON.parse(response.getContentText());
    
    if (result.error) {
      throw new Error("PDF 서버 오류: " + result.error);
    }
    
    // PDF 파일을 Google Drive에 저장
    var folder = getOrCreateReportFolderV4();
    var savedFiles = {};
    
    if (result.summaryPdf) {
      var summaryBlob = Utilities.newBlob(Utilities.base64Decode(result.summaryPdf), 'application/pdf', 
        params.businessName + '_요약보고서_' + Utilities.formatDate(new Date(), "Asia/Seoul", "yyyyMMdd") + '.pdf');
      var summaryFile = folder.createFile(summaryBlob);
      savedFiles.summary = {
        id: summaryFile.getId(),
        url: summaryFile.getUrl(),
        name: summaryFile.getName()
      };
    }
    
    if (result.detailPdf) {
      var detailBlob = Utilities.newBlob(Utilities.base64Decode(result.detailPdf), 'application/pdf',
        params.businessName + '_상세보고서_' + Utilities.formatDate(new Date(), "Asia/Seoul", "yyyyMMdd") + '.pdf');
      var detailFile = folder.createFile(detailBlob);
      savedFiles.detail = {
        id: detailFile.getId(),
        url: detailFile.getUrl(),
        name: detailFile.getName()
      };
    }
    
    return savedFiles;
    
  } catch (e) {
    Logger.log("PDF 서버 요청 실패: " + e.message);
    
    // 폴백: Google Docs로 생성
    return generateFallbackReportV4(rawData, transformedData, params);
  }
}

/**
 * 폴백: Google Docs로 리포트 생성
 */
function generateFallbackReportV4(rawData, transformedData, params) {
  Logger.log("폴백 모드: Google Docs로 리포트 생성");
  
  var folder = getOrCreateReportFolderV4();
  var savedFiles = {};
  var timestamp = Utilities.formatDate(new Date(), "Asia/Seoul", "yyyyMMdd_HHmm");
  
  // 요약 보고서 생성
  if (params.generateSummary) {
    var summaryDoc = DocumentApp.create(params.businessName + "_요약보고서_" + timestamp);
    var summaryBody = summaryDoc.getBody();
    
    // 표지
    buildCoverPageV4(summaryBody, params.businessName, "요약 보고서");
    summaryBody.appendPageBreak();
    
    // 경영진 요약
    if (transformedData.executiveSummary) {
      appendMarkdownToDocV4(summaryBody, transformedData.executiveSummary);
    }
    
    summaryDoc.saveAndClose();
    
    // PDF 변환
    var summaryPdf = DriveApp.getFileById(summaryDoc.getId()).getAs('application/pdf');
    var summaryFile = folder.createFile(summaryPdf).setName(params.businessName + "_요약보고서_" + timestamp + ".pdf");
    
    savedFiles.summary = {
      id: summaryFile.getId(),
      url: summaryFile.getUrl(),
      name: summaryFile.getName()
    };
    
    // Docs 삭제 (PDF만 유지)
    DriveApp.getFileById(summaryDoc.getId()).setTrashed(true);
  }
  
  // 상세 보고서 생성
  if (params.generateDetail) {
    var detailDoc = DocumentApp.create(params.businessName + "_상세보고서_" + timestamp);
    var detailBody = detailDoc.getBody();
    
    // 표지
    buildCoverPageV4(detailBody, params.businessName, "상세 분석 보고서");
    detailBody.appendPageBreak();
    
    // 목차
    buildTableOfContentsV4(detailBody);
    detailBody.appendPageBreak();
    
    // 각 섹션 추가
    var sectionOrder = ['pestel', 'scenario', 'competition', 'customer', 'market', 'diagnosis', 'vrio', 'swot', 'tows'];
    for (var i = 0; i < sectionOrder.length; i++) {
      var sectionId = sectionOrder[i];
      if (transformedData.sections[sectionId] && transformedData.sections[sectionId].content) {
        appendMarkdownToDocV4(detailBody, transformedData.sections[sectionId].content);
        if (i < sectionOrder.length - 1) {
          detailBody.appendPageBreak();
        }
      }
    }
    
    detailDoc.saveAndClose();
    
    // PDF 변환
    var detailPdf = DriveApp.getFileById(detailDoc.getId()).getAs('application/pdf');
    var detailFile = folder.createFile(detailPdf).setName(params.businessName + "_상세보고서_" + timestamp + ".pdf");
    
    savedFiles.detail = {
      id: detailFile.getId(),
      url: detailFile.getUrl(),
      name: detailFile.getName()
    };
    
    // Docs 삭제
    DriveApp.getFileById(detailDoc.getId()).setTrashed(true);
  }
  
  return savedFiles;
}

/**
 * 표지 생성
 */
function buildCoverPageV4(body, businessName, reportType) {
  var titleStyle = {};
  titleStyle[DocumentApp.Attribute.FONT_SIZE] = 28;
  titleStyle[DocumentApp.Attribute.BOLD] = true;
  titleStyle[DocumentApp.Attribute.FOREGROUND_COLOR] = '#1a73e8';
  
  var subtitleStyle = {};
  subtitleStyle[DocumentApp.Attribute.FONT_SIZE] = 16;
  subtitleStyle[DocumentApp.Attribute.FOREGROUND_COLOR] = '#666666';
  
  body.appendParagraph("\n\n\n\n").setAttributes({});
  
  var title = body.appendParagraph("G-IMPACT 분석 리포트");
  title.setAttributes(titleStyle);
  title.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  
  body.appendParagraph("\n");
  
  var subtitle = body.appendParagraph(reportType);
  subtitle.setAttributes(subtitleStyle);
  subtitle.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  
  body.appendParagraph("\n\n\n");
  
  var companyTitle = body.appendParagraph(businessName);
  companyTitle.setAttributes({});
  companyTitle.setFontSize(24);
  companyTitle.setBold(true);
  companyTitle.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  
  body.appendParagraph("\n\n\n\n");
  
  var dateStr = Utilities.formatDate(new Date(), "Asia/Seoul", "yyyy년 M월 d일");
  var datePara = body.appendParagraph(dateStr);
  datePara.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  datePara.setFontSize(12);
  datePara.setForegroundColor('#999999');
}

/**
 * 목차 생성
 */
function buildTableOfContentsV4(body) {
  var tocTitle = body.appendParagraph("목차");
  tocTitle.setHeading(DocumentApp.ParagraphHeading.HEADING1);
  
  var tocItems = [
    "1. 거시환경 분석 (PESTEL)",
    "2. 미래 시나리오 분석",
    "3. 경쟁환경 분석",
    "4. 고객 분석",
    "5. 시장 분석",
    "6. 경영진단",
    "7. 핵심 자원 분석 (VRIO)",
    "8. SWOT 분석",
    "9. TOWS 전략"
  ];
  
  for (var i = 0; i < tocItems.length; i++) {
    var item = body.appendParagraph(tocItems[i]);
    item.setFontSize(12);
    item.setSpacingAfter(8);
  }
}

/**
 * 마크다운을 Google Docs에 추가
 */
function appendMarkdownToDocV4(body, markdown) {
  if (!markdown) return;
  
  var lines = markdown.split('\n');
  var inTable = false;
  var tableData = [];
  
  for (var i = 0; i < lines.length; i++) {
    var line = lines[i];
    
    // 테이블 처리
    if (line.trim().startsWith('|')) {
      if (!inTable) {
        inTable = true;
        tableData = [];
      }
      // 구분선 스킵
      if (line.indexOf('---') === -1) {
        var cells = line.split('|').filter(function(c) { return c.trim() !== ''; });
        tableData.push(cells.map(function(c) { return c.trim(); }));
      }
      continue;
    } else if (inTable && tableData.length > 0) {
      // 테이블 종료, 테이블 생성
      try {
        var table = body.appendTable(tableData);
        styleTableV4(table);
      } catch (e) {
        Logger.log("테이블 생성 오류: " + e.message);
      }
      inTable = false;
      tableData = [];
    }
    
    // 헤딩 처리
    if (line.startsWith('## ')) {
      var h2 = body.appendParagraph(line.substring(3));
      h2.setHeading(DocumentApp.ParagraphHeading.HEADING2);
      h2.setForegroundColor('#1a73e8');
      continue;
    }
    if (line.startsWith('### ')) {
      var h3 = body.appendParagraph(line.substring(4));
      h3.setHeading(DocumentApp.ParagraphHeading.HEADING3);
      continue;
    }
    if (line.startsWith('#### ')) {
      var h4 = body.appendParagraph(line.substring(5));
      h4.setBold(true);
      continue;
    }
    
    // 일반 텍스트
    if (line.trim()) {
      var para = body.appendParagraph(line);
      
      // Bold 처리 (**text**)
      var boldRegex = /\*\*([^*]+)\*\*/g;
      var text = para.editAsText();
      var match;
      while ((match = boldRegex.exec(line)) !== null) {
        var start = line.indexOf(match[0]);
        if (start >= 0) {
          text.setBold(start, start + match[0].length - 1, true);
        }
      }
    }
  }
  
  // 마지막 테이블 처리
  if (inTable && tableData.length > 0) {
    try {
      var table = body.appendTable(tableData);
      styleTableV4(table);
    } catch (e) {}
  }
}

/**
 * 테이블 스타일 적용
 */
function styleTableV4(table) {
  if (!table || table.getNumRows() === 0) return;
  
  // 헤더 행 스타일
  var headerRow = table.getRow(0);
  for (var i = 0; i < headerRow.getNumCells(); i++) {
    var cell = headerRow.getCell(i);
    cell.setBackgroundColor('#1a73e8');
    cell.editAsText().setForegroundColor('#ffffff').setBold(true);
  }
  
  // 데이터 행 스타일
  for (var r = 1; r < table.getNumRows(); r++) {
    var row = table.getRow(r);
    var bgColor = (r % 2 === 0) ? '#f8f9fa' : '#ffffff';
    for (var c = 0; c < row.getNumCells(); c++) {
      row.getCell(c).setBackgroundColor(bgColor);
    }
  }
  
  // 테두리
  table.setBorderWidth(1);
  table.setBorderColor('#e0e0e0');
}

// ============================================
// [10] 이메일 발송
// ============================================

/**
 * 완료 이메일 발송
 */
function sendCompletionEmailV4(email, businessName, pdfResult) {
  var subject = REPORT_CONFIG_V4.emailSubject + " - " + businessName;
  
  var body = `
안녕하세요,

${businessName}의 G-IMPACT 분석 리포트가 생성되었습니다.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

📄 생성된 리포트

`;

  if (pdfResult.summary) {
    body += `✅ 요약 보고서: ${pdfResult.summary.name}
   다운로드: ${pdfResult.summary.url}

`;
  }

  if (pdfResult.detail) {
    body += `✅ 상세 보고서: ${pdfResult.detail.name}
   다운로드: ${pdfResult.detail.url}

`;
  }

  body += `
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

📂 리포트 폴더: ${getOrCreateReportFolderV4().getUrl()}

생성일시: ${Utilities.formatDate(new Date(), "Asia/Seoul", "yyyy-MM-dd HH:mm:ss")}

---
G-IMPACT 분석 시스템 v${REPORT_CONFIG_V4.version}
`;

  GmailApp.sendEmail(email, subject, body);
  Logger.log("완료 이메일 발송: " + email);
}

/**
 * 오류 이메일 발송
 */
function sendErrorEmailV4(email, businessName, errorMessage) {
  var subject = "[오류] G-IMPACT 리포트 생성 실패 - " + businessName;
  
  var body = `
안녕하세요,

${businessName}의 G-IMPACT 분석 리포트 생성 중 오류가 발생했습니다.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

❌ 오류 내용:
${errorMessage}

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

문제가 지속되면 관리자에게 문의하세요.

발생일시: ${Utilities.formatDate(new Date(), "Asia/Seoul", "yyyy-MM-dd HH:mm:ss")}

---
G-IMPACT 분석 시스템 v${REPORT_CONFIG_V4.version}
`;

  GmailApp.sendEmail(email, subject, body);
  Logger.log("오류 이메일 발송: " + email);
}

// ============================================
// [11] 기타 유틸리티
// ============================================

/**
 * 리포트 이력 보기
 */
function showReportHistoryV4() {
  var folder = getOrCreateReportFolderV4();
  var files = folder.getFiles();
  
  var history = [];
  while (files.hasNext()) {
    var file = files.next();
    if (file.getMimeType() === 'application/pdf') {
      history.push({
        name: file.getName(),
        date: file.getDateCreated(),
        url: file.getUrl()
      });
    }
  }
  
  // 날짜 역순 정렬
  history.sort(function(a, b) { return b.date - a.date; });
  
  var html = '<html><head><style>';
  html += 'body{font-family:sans-serif;padding:20px;}';
  html += 'table{width:100%;border-collapse:collapse;}';
  html += 'th,td{padding:10px;text-align:left;border-bottom:1px solid #ddd;}';
  html += 'th{background:#1a73e8;color:white;}';
  html += 'a{color:#1a73e8;}';
  html += '</style></head><body>';
  html += '<h2>📋 리포트 생성 이력</h2>';
  html += '<table><tr><th>파일명</th><th>생성일</th><th>링크</th></tr>';
  
  for (var i = 0; i < Math.min(history.length, 20); i++) {
    var h = history[i];
    html += '<tr>';
    html += '<td>' + h.name + '</td>';
    html += '<td>' + Utilities.formatDate(h.date, "Asia/Seoul", "yyyy-MM-dd HH:mm") + '</td>';
    html += '<td><a href="' + h.url + '" target="_blank">다운로드</a></td>';
    html += '</tr>';
  }
  
  html += '</table></body></html>';
  
  var output = HtmlService.createHtmlOutput(html).setWidth(700).setHeight(500);
  SpreadsheetApp.getUi().showModalDialog(output, '리포트 이력');
}

/**
 * 캐시 초기화
 */
function clearReportCacheV4() {
  var ui = SpreadsheetApp.getUi();
  var result = ui.alert(
    '캐시 초기화',
    '진행 중인 작업 상태를 초기화하시겠습니까?',
    ui.ButtonSet.YES_NO
  );
  
  if (result === ui.Button.YES) {
    var props = PropertiesService.getScriptProperties();
    var keys = props.getKeys();
    
    var deleted = 0;
    for (var i = 0; i < keys.length; i++) {
      if (keys[i].indexOf('REPORT_PROGRESS_') === 0 || keys[i].indexOf('REPORT_PARAMS_') === 0) {
        props.deleteProperty(keys[i]);
        deleted++;
      }
    }
    
    ui.alert('초기화 완료', deleted + '개의 캐시가 삭제되었습니다.', ui.ButtonSet.OK);
  }
}

// ============================================
// [12] 호환성 함수
// ============================================

/**
 * 이전 버전 호환
 */
function showReportGeneratorV4() {
  showReportDialogV4();
}

