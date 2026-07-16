/**
 * 테크포임팩트 캠퍼스 사회혁신가 신청현황 대시보드용 데이터 API
 *
 * 설치 방법 (한 번만 하면 돼요):
 * 1. 신청 응답 시트 열기 → 상단 메뉴 [확장 프로그램] → [Apps Script]
 * 2. 기존 코드 지우고 이 파일 내용 전체를 붙여넣기 → 저장(💾)
 * 3. 오른쪽 위 [배포] → [새 배포] → 유형 '웹 앱' 선택
 *    - 실행 계정: 나
 *    - 액세스 권한: "링크가 있는 모든 사용자"
 * 4. [배포] 클릭 → 나오는 웹 앱 URL 복사
 * 5. 복사한 URL을 index.html 상단의 DATA_URL = "" 안에 붙여넣기
 *
 * ※ 코드를 수정한 뒤에는 [배포] → [배포 관리] → 연필 → 버전 '새 버전'으로
 *    다시 배포해야 반영돼요.
 */

const SHEET_GID = 229370971; // 응답 탭의 gid

// GitHub Pages 공개 배포용이라 연락처(전화·이메일)는 내보내지 않아요.
// 연락처가 필요하면 시트에서 직접 확인하세요.
const INCLUDE_CONTACTS = false;

function doGet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheets().find(s => s.getSheetId() === SHEET_GID) || ss.getSheets()[0];
  const values = sheet.getDataRange().getValues();
  const header = values[0];

  // 헤더 텍스트로 열을 찾아서, 열 순서가 바뀌어도 안전하게
  const col = (keyword) => header.findIndex(h => String(h).includes(keyword));
  const iTs     = 0;                    // A열: 타임스탬프
  const iName   = col('성함');           // Q1
  const iPhone  = col('휴대폰');          // Q2
  const iEmail  = col('이메일 주소를');    // Q3
  const iRole   = col('직함');           // Q4
  const iPath   = col('알게 된 경로');     // Q5
  const iOrg    = col('조직명');          // Q6
  const iType   = col('조직 유형');        // Q7
  const iYears  = col('활동 기간');        // Q8
  const iFields = col('활동 분야');        // Q9
  const iRegion = col('활동 지역');        // Q10 (L열)
  const iIntro  = col('주요 활동을');      // Q11
  const iLink   = col('홈페이지');         // Q12
  const iReason = col('참여하고 싶은 이유'); // Q14
  const iQuest  = col('질문');            // Q15
  const iOutput = col('결과물');          // Q16

  // H열(index 7) 배경색으로 기참여 여부 판별
  const dataRange = sheet.getRange(2, 8, values.length - 1, 1); // H열, 데이터 행만
  const bgColors = dataRange.getBackgrounds().map(r => r[0]);

  const tz = ss.getSpreadsheetTimeZone();
  const str = (row, i) => i >= 0 ? String(row[i] || '').trim() : '';

  // 노란색=기참여, 초록색=네트워크, 파란색=제안메일, 그 외(보라/없음 등)=공모
  // 노란색=기참여, 초록색=네트워크, 파란색=제안메일, 그 외=공모
  const colorTag = (hex) => {
    const h = (hex || '').toLowerCase();
    if (h === '#ffd966') return '기참여';
    if (h === '#93c47d') return '네트워크';
    if (h === '#6fa8dc') return '제안메일';
    return '공모';
  };

  const rows = values.slice(1)
    .map((r, i) => ({
      ts: r[iTs] instanceof Date
        ? Utilities.formatDate(r[iTs], tz, 'yyyy-MM-dd HH:mm')
        : String(r[iTs]),
      name: str(r, iName),
      phone: INCLUDE_CONTACTS ? str(r, iPhone) : '',
      email: INCLUDE_CONTACTS ? str(r, iEmail) : '',
      role: str(r, iRole),
      path: str(r, iPath),
      org: str(r, iOrg),
      type: str(r, iType),
      years: str(r, iYears),
      fields: str(r, iFields),
      regions: str(r, iRegion).split(',').map(s => s.trim()).filter(Boolean),
      intro: str(r, iIntro),
      link: str(r, iLink),
      reason: str(r, iReason),
      questions: str(r, iQuest),
      outcome: str(r, iOutput),
      tag: colorTag(bgColors[i] || '')
    })).filter(r => r.org);

  const payload = {
    updatedAt: Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm'),
    rows: rows
  };

  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}
