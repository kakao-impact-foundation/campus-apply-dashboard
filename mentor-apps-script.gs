/**
 * 테크포임팩트 캠퍼스 26-2 멘토 신청현황 대시보드용 데이터 API
 * (mentor.html 전용 — 사회혁신가 대시보드의 apps-script.gs와는 별개예요)
 *
 * 설치 방법 (한 번만 하면 돼요):
 * 1. "[테크포임팩트 캠퍼스] 26-2학기 멘토 신청서(응답)" 시트 열기
 *    → 상단 메뉴 [확장 프로그램] → [Apps Script]
 * 2. 기존 코드 지우고 이 파일 내용 전체를 붙여넣기 → 저장(💾)
 * 3. 오른쪽 위 [배포] → [새 배포] → 유형 '웹 앱' 선택
 *    - 실행 계정: 나
 *    - 액세스 권한: "링크가 있는 모든 사용자"
 * 4. [배포] 클릭 → 나오는 웹 앱 URL 복사
 * 5. 복사한 URL을 mentor.html 상단의 DATA_URL = "" 안에 붙여넣기
 *
 * ※ 시트 자체의 공유 설정은 바꿀 필요 없어요 (웹 앱만 링크 공개).
 * ※ 코드를 수정한 뒤에는 [배포] → [배포 관리] → 연필 → 버전 '새 버전'으로
 *    다시 배포해야 반영돼요.
 */

const SHEET_GID = 679275159; // 응답 탭의 gid
const ASSIGN_SHEET = '멘토배정'; // 배정 보드 저장 탭 (없으면 자동 생성)

// GitHub Pages 공개 배포용이라 연락처(전화·이메일)는 내보내지 않아요.
// 연락처가 필요하면 시트에서 직접 확인하세요.

function doGet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheets().find(s => s.getSheetId() === SHEET_GID) || ss.getSheets()[0];
  const values = sheet.getDataRange().getValues();
  const header = values[0];

  // 헤더 텍스트로 열을 찾아서, 열 순서가 바뀌어도 안전하게
  const col = (keyword) => header.findIndex(h => String(h).includes(keyword));
  const iTs      = 0;                     // A열: 타임스탬프
  const iName    = col('성함');
  const iLdap    = header.findIndex(h => String(h).trim() === 'LDAP'); // '추천 크루 LDAP' 열과 구분
  const iCompany = col('공동체명');
  const iDept    = col('부서명');
  const iJob     = col('직군');
  const iExp     = col('참여 경험');
  const iMotive  = col('지원 동기');
  const iWork    = col('맡고 계신 업무');
  const iCoach   = col('코칭 가능한');
  const iMentor  = col('멘토링 경험');
  const iSchools = col('희망하는 학교');
  const iAlma    = col('모교 희망');
  const iRef     = col('추천 크루');

  const tz = ss.getSpreadsheetTimeZone();
  const str = (row, i) => i >= 0 ? String(row[i] || '').trim() : '';

  const rows = values.slice(1)
    .map(r => ({
      ts: r[iTs] instanceof Date
        ? Utilities.formatDate(r[iTs], tz, 'yyyy-MM-dd HH:mm')
        : String(r[iTs]),
      name: str(r, iName),
      ldap: str(r, iLdap),
      company: str(r, iCompany),
      dept: str(r, iDept),
      job: str(r, iJob),
      exp: str(r, iExp),
      motive: str(r, iMotive),
      work: str(r, iWork),
      coach: str(r, iCoach),
      mentorExp: str(r, iMentor),
      schools: str(r, iSchools),
      alma: str(r, iAlma),
      referral: str(r, iRef)
    })).filter(r => r.name);

  const payload = {
    updatedAt: Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm'),
    rows: rows,
    assignments: readAssignments(ss)
  };

  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ── 배정 보드 (대시보드 [배정 보드] 탭) ─────────────────────
 * 칩을 옮기면 doPost로 들어와 '멘토배정' 탭에 (LDAP, 이름, 배정대학) 저장돼요.
 * 같은 LDAP은 덮어쓰기, 배정대학 ''(미배정)은 행 삭제.
 */

function readAssignments(ss) {
  const sh = ss.getSheetByName(ASSIGN_SHEET);
  const out = {};
  if (!sh || sh.getLastRow() < 2) return out;
  sh.getDataRange().getValues().slice(1).forEach(r => {
    const ldap = String(r[1] || '').trim();
    const school = String(r[3] || '').trim();
    if (ldap && school) out[ldap] = school;
  });
  return out;
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const body = JSON.parse(e.postData.contents);
    if (body.action !== 'assign') return jsonOut({ ok: false, error: 'invalid action' });
    const ldap = String(body.ldap || '').trim().slice(0, 100);
    const name = String(body.name || '').trim().slice(0, 50);
    const school = String(body.school || '').trim().slice(0, 60);
    if (!ldap) return jsonOut({ ok: false, error: 'invalid ldap' });

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sh = ss.getSheetByName(ASSIGN_SHEET);
    if (!sh) {
      sh = ss.insertSheet(ASSIGN_SHEET);
      sh.appendRow(['수정시각', 'LDAP', '이름', '배정대학']);
    }
    const data = sh.getDataRange().getValues();
    let rowIdx = -1;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][1]).trim() === ldap) { rowIdx = i + 1; break; }
    }
    if (school === '') {
      if (rowIdx > 0) sh.deleteRow(rowIdx);
    } else if (rowIdx > 0) {
      sh.getRange(rowIdx, 1, 1, 4).setValues([[new Date(), ldap, name, school]]);
    } else {
      sh.appendRow([new Date(), ldap, name, school]);
    }
    return jsonOut({ ok: true });
  } catch (err) {
    return jsonOut({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
