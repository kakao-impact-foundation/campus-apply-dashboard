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
      referral: str(r, iRef)
    })).filter(r => r.name);

  const payload = {
    updatedAt: Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm'),
    rows: rows
  };

  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}
