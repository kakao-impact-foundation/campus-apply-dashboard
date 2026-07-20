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
const RATING_SHEET = '운영진평가'; // 평점 저장 탭 (없으면 자동 생성)
const RATERS = ['헤이븐', '조안', '준'];

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

  const iTag    = col('참가자 구분');    // U열

  const tz = ss.getSpreadsheetTimeZone();
  const str = (row, i) => i >= 0 ? String(row[i] || '').trim() : '';

  // U열 값에서 순위와 하위 카테고리 분리: "1순위 - 기참여" → { rank: "1순위", sub: "기참여" }
  const parseTag = (val) => {
    const s = String(val || '').trim();
    const m = s.match(/^(\d순위)\s*-\s*(.+)/);
    if (m) return { rank: m[1], sub: m[2].trim() };
    return { rank: '4순위', sub: '공개모집' };
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
      tag: parseTag(str(r, iTag)).rank,
      tagSub: parseTag(str(r, iTag)).sub
    })).filter(r => r.org);

  const payload = {
    updatedAt: Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm'),
    rows: rows,
    ratings: readRatings(ss)
  };

  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ── 운영진 평가 (헤이븐·조안·준) ─────────────────────────
 * 대시보드에서 별점을 누르면 doPost로 들어와
 * '운영진평가' 탭에 (조직키, 평가자, 점수) 한 줄씩 저장돼요.
 * 같은 조직·같은 평가자는 덮어쓰기, 점수 0은 평가 취소(삭제).
 */

function readRatings(ss) {
  const sh = ss.getSheetByName(RATING_SHEET);
  const out = {};
  if (!sh || sh.getLastRow() < 2) return out;
  sh.getDataRange().getValues().slice(1).forEach(r => {
    const key = String(r[1] || '').trim();
    const rater = String(r[2] || '').trim();
    const score = Number(r[3]);
    if (!key || !rater || !(score >= 1 && score <= 5)) return;
    (out[key] = out[key] || {})[rater] = score;
  });
  return out;
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const body = JSON.parse(e.postData.contents);
    const key = String(body.orgKey || '').trim().slice(0, 200);
    const rater = String(body.rater || '').trim();
    const score = Number(body.score);
    if (!key || RATERS.indexOf(rater) === -1 || !(score >= 0 && score <= 5)) {
      return jsonOut({ ok: false, error: 'invalid' });
    }
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sh = ss.getSheetByName(RATING_SHEET);
    if (!sh) {
      sh = ss.insertSheet(RATING_SHEET);
      sh.appendRow(['수정시각', '조직키', '평가자', '점수']);
    }
    const data = sh.getDataRange().getValues();
    let rowIdx = -1;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][1]).trim() === key && String(data[i][2]).trim() === rater) {
        rowIdx = i + 1;
        break;
      }
    }
    if (score === 0) {
      if (rowIdx > 0) sh.deleteRow(rowIdx);
    } else if (rowIdx > 0) {
      sh.getRange(rowIdx, 1, 1, 4).setValues([[new Date(), key, rater, score]]);
    } else {
      sh.appendRow([new Date(), key, rater, score]);
    }
    return jsonOut({ ok: true, ratings: readRatings(ss)[key] || {} });
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
