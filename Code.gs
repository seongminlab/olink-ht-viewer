// Olink HT Viewer — Google Apps Script API
// Deploy as Web App: Execute as Me, Access: Anyone (including anonymous)

var SPREADSHEET_ID = "YOUR_SPREADSHEET_ID_HERE"; // ← 스프레드시트 ID 입력

// ── Settings 시트 구조 ──────────────────────────────────────────
// 1행 헤더: Project status | Project ID | description
// 2행~:    데이터 (status: "hidden" 이면 API 목록에서 제외)
//
// ── 요청 방식 ────────────────────────────────────────────────────
// 파라미터 없음           → 프로젝트 목록 반환
// ?project=PROJ001       → 해당 프로젝트 샘플 데이터 반환

function doGet(e) {
  var params = (e && e.parameter) ? e.parameter : {};
  var projectId = params.project || '';

  try {
    var result = projectId ? getProjectData(projectId) : getProjectList();
    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Settings 시트에서 프로젝트 목록 반환
// status가 "hidden"(대소문자 무관)인 행은 제외
function getProjectList() {
  var ss       = SpreadsheetApp.openById(SPREADSHEET_ID);
  var settings = ss.getSheetByName("Settings");
  if (!settings) throw new Error("Settings 시트를 찾을 수 없습니다.");

  var lastRow = settings.getLastRow();
  var projects = [];

  if (lastRow >= 2) {
    var rows = settings.getRange(2, 1, lastRow - 1, 3).getValues();
    rows.forEach(function(row) {
      var status      = String(row[0]).trim();
      var id          = String(row[1]).trim();
      var description = String(row[2]).trim();
      if (!id) return;                                   // 빈 행 무시
      if (status.toLowerCase() === 'hidden') return;    // hidden 제외
      projects.push({ status: status, id: id, description: description });
    });
  }

  return { projects: projects };
}

// 특정 프로젝트 시트의 샘플 데이터 반환 (A~E 컬럼만, F~H 비공개)
function getProjectData(projectId) {
  var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(projectId);
  if (!sheet) throw new Error("프로젝트 시트 '" + projectId + "'를 찾을 수 없습니다.");

  var lastRow = sheet.getLastRow();
  var samples = [];

  if (lastRow >= 2) {
    var data = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
    data.forEach(function(row) {
      var requestCode = String(row[0]).trim();
      if (!requestCode) return;

      samples.push({
        request_code: requestCode,
        sample_type:  String(row[1]).trim(),
        sample_count: parseInt(row[2], 10) || 0,
        request_date: formatDateCell(row[3]),
        receive_date: formatDateCell(row[4])
      });
    });
  }

  return { project: projectId, samples: samples };
}

function formatDateCell(val) {
  if (!val) return "";
  if (val instanceof Date && !isNaN(val)) {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  return String(val).trim();
}
