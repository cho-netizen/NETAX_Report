/**
 * NETAX 자문보고서 열람 시스템
 * Google Apps Script - Web App (doPost)
 *
 * 배포 방법:
 * 1. Google Sheets 열고 확장 프로그램 > Apps Script
 * 2. 이 코드 붙여넣기
 * 3. SHEET_ID를 실제 스프레드시트 ID로 교체 (또는 SpreadsheetApp.getActiveSpreadsheet() 사용 가능)
 * 4. 관리자 우회 코드는 "스크립트 속성"에 저장 (아래 SETUP 함수 실행 or 직접 설정)
 * 5. 배포 > 새 배포 > 유형: 웹앱 > 액세스 권한: "모든 사용자" 로 설정
 * 6. 배포 후 나오는 웹앱 URL을 프론트(GitHub Pages)의 fetch 대상으로 사용
 */

// ===== 설정 =====
const SHEET_ID = '1fE0Vm33n8ivSzO0bFV6xwK6Bxav-Xdi92yHvYVpqZOc';
const SHEET_CUSTOMER = 'Reports';    // 시트1 이름
const SHEET_LOG = 'AccessLog';       // 시트2 이름

// ===== 최초 1회 실행: 관리자 우회 코드 설정 =====
// Apps Script 편집기에서 이 함수를 한 번 실행하면 스크립트 속성에 저장됩니다.
function SETUP_setAdminCode() {
  PropertiesService.getScriptProperties().setProperty('ADMIN_CODE', '여기에_원하는_관리자코드_입력');
}

// ===== 메인 진입점 =====
function doPost(e) {
  try {
    const params = JSON.parse(e.postData.contents);
    const reportId = (params.report_id || '').trim();
    const passwordHash = params.password_hash || '';
    const adminCode = params.admin_code || '';

    if (!reportId) {
      return jsonResponse({ success: false, message: 'report_id가 필요합니다.' });
    }

    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(SHEET_CUSTOMER);
    const data = sheet.getDataRange().getValues();
    const headers = data[0];

    const colName = headers.indexOf('고객명');
    const colReportId = headers.indexOf('report_id');
    const colIssued = headers.indexOf('발급일');
    const colHash = headers.indexOf('비밀번호해시');
    const colExpiry = headers.indexOf('만료일');
    const colLink = headers.indexOf('자료링크');
    const colContent = headers.indexOf('보고서원문');

    let targetRow = null;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][colReportId]).trim() === reportId) {
        targetRow = data[i];
        break;
      }
    }

    if (!targetRow) {
      return jsonResponse({ success: false, message: '존재하지 않는 report_id입니다.' });
    }

    // 관리자 우회 체크 (비밀번호 절차 스킵)
    const storedAdminCode = PropertiesService.getScriptProperties().getProperty('ADMIN_CODE');
    const isAdmin = storedAdminCode && adminCode === storedAdminCode;

    const rowIndex = data.indexOf(targetRow) + 1; // 시트상 실제 행 번호 (1-indexed, 헤더 포함)
    const storedHash = String(targetRow[colHash]).trim();
    const hasPassword = storedHash.length > 0;

    if (!isAdmin) {
      // 만료일 체크
      const expiry = new Date(targetRow[colExpiry]);
      const today = new Date();
      today.setHours(0, 0, 0, 0);
      if (today > expiry) {
        return jsonResponse({ success: false, message: '열람 기간이 만료되었습니다.' });
      }

      if (!hasPassword) {
        // ===== 최초 접속: 비밀번호가 아직 없으면 이번 요청으로 설정 =====
        if (!passwordHash) {
          // 비밀번호 없이 report_id만 조회한 경우 -> "최초 설정 필요" 신호만 보냄
          return jsonResponse({ success: false, needs_setup: true, customer_name: targetRow[colName], message: '최초 접속입니다. 비밀번호를 설정해주세요.' });
        }
        sheet.getRange(rowIndex, colHash + 1).setValue(passwordHash);
      } else {
        // ===== 이미 비밀번호가 설정된 경우: 일반 로그인 =====
        if (!passwordHash) {
          return jsonResponse({ success: false, needs_login: true, customer_name: targetRow[colName], message: '비밀번호를 입력해주세요.' });
        }
        if (passwordHash !== storedHash) {
          return jsonResponse({ success: false, message: '비밀번호가 일치하지 않습니다.' });
        }
      }
    }

    // 접속 로그 기록
    logAccess(ss, reportId);

    // 성공 응답 - 보고서원문(마크다운)이 있으면 그것을, 없으면 기존 자료링크를 반환
    const mdContent = colContent >= 0 ? String(targetRow[colContent]).trim() : '';
    const response = {
      success: true,
      customer_name: targetRow[colName],
      is_admin: isAdmin
    };
    if (mdContent) {
      response.report_content = mdContent;
    } else {
      response.material_link = targetRow[colLink];
    }
    return jsonResponse(response);

  } catch (err) {
    return jsonResponse({ success: false, message: '서버 오류: ' + err.message });
  }
}

// ===== 접속 로그 기록 =====
function logAccess(ss, reportId) {
  const logSheet = ss.getSheetByName(SHEET_LOG);
  logSheet.appendRow([reportId, new Date()]);
}

// ===== SHA-256 해시 (프론트와 동일 알고리즘) =====
function sha256(text) {
  const rawHash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, text, Utilities.Charset.UTF_8);
  return rawHash.map(byte => {
    const v = (byte < 0 ? byte + 256 : byte).toString(16);
    return v.length === 1 ? '0' + v : v;
  }).join('');
}

// ===== JSON 응답 헬퍼 =====
function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
