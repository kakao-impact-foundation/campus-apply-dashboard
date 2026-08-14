/**
 * 테크포임팩트 캠퍼스 26-2 대학 운영 관리 대시보드용 데이터 API
 * (campus.html 전용 — 사회혁신가/멘토 대시보드의 Apps Script와는 별개예요)
 *
 * 설치 방법 (한 번만 하면 돼요):
 * 1. "캠퍼스 운영 관리" 마스터 시트(학교 목록·시트 주소가 있는 시트) 열기
 *    → 상단 메뉴 [확장 프로그램] → [Apps Script]
 * 2. 기존 코드 지우고 이 파일 내용 전체를 붙여넣기 → 저장(💾)
 * 3. 오른쪽 위 [배포] → [새 배포] → 유형 '웹 앱' 선택
 *    - 실행 계정: 나
 *    - 액세스 권한: "링크가 있는 모든 사용자"
 * 4. [배포] 클릭 → 처음엔 권한 승인 창이 떠요 (외부 스프레드시트 열람 권한)
 * 5. 나오는 웹 앱 URL을 campus.html 상단의 DATA_URL = "" 안에 붙여넣기
 *
 * ※ 학교별 시트는 이 스크립트를 실행하는 계정이 열람 가능하기만 하면 돼요
 *    (마스터 시트의 '시트 주소' 열에 링크를 추가하면 대시보드에 자동 반영).
 * ※ 코드를 수정한 뒤에는 [배포] → [배포 관리] → 연필 → 버전 '새 버전'으로
 *    다시 배포해야 반영돼요.
 */

const MASTER_GID = 2000827737; // 학교 목록 탭의 gid (학교명 · 시트 주소)

/* 각 학교 시트에서 읽어올 탭 이름 (템플릿 공통) */
const TAB_NAMES = [
  '체크리스트',
  '수업 및 운영진 정보',
  '예산 계획 및 기부금 처리',
  '수강 인원 및 팀 정보',
  '성과발표회'
];

/* 연락처(전화·이메일)·학번 마스킹 여부 — 운영팀 결정으로 어드민에서는 원본 표시(true).
 * ※ 웹앱 URL을 아는 사람은 비밀번호 게이트 없이도 이 데이터를 볼 수 있어요.
 *   민감도가 올라가면 false로 되돌리고 새 버전 배포. */
const INCLUDE_CONTACTS = true;

const CACHE_SEC = 300;   // 학교별 캐시 (초) — 12개 시트를 매번 열면 느려서
const MAX_ROWS = 200;    // 탭당 최대 행
const MAX_COLS = 30;     // 탭당 최대 열

function doGet(e) {
  const fresh = e && e.parameter && e.parameter.fresh === '1';
  const only = e && e.parameter && String(e.parameter.school || '').trim();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = ss.getSpreadsheetTimeZone();
  let list = readMasterList(ss);
  if (only) list = list.filter(s => s.name.indexOf(only) !== -1);

  const schools = list.map(s => fetchSchool(s, fresh));

  const payload = {
    v: 1, // 코드 버전 (배포 확인용)
    updatedAt: Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm'),
    masked: !INCLUDE_CONTACTS,
    schools: schools
  };

  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ── 마스터 시트에서 학교 목록 읽기 ──
 * '시트 주소' 열에서 구글시트 링크가 있는 행만 추려요.
 * 학교명은 링크 왼쪽의 첫 번째 비어있지 않은 셀. */
function readMasterList(ss) {
  const sheet = ss.getSheets().find(s => s.getSheetId() === MASTER_GID) || ss.getSheets()[0];
  const values = sheet.getDataRange().getValues();
  const out = [];
  values.forEach(row => {
    const urlIdx = row.findIndex(c => /docs\.google\.com\/spreadsheets\//.test(String(c)));
    if (urlIdx < 1) return;
    let name = '';
    for (let i = urlIdx - 1; i >= 0; i--) {
      if (String(row[i]).trim()) { name = String(row[i]).trim(); break; }
    }
    if (name && name !== '학교목록') out.push({ name: name, url: String(row[urlIdx]).trim() });
  });
  return out;
}

/* ── 학교 시트 하나 읽기 (캐시 사용) ── */
function fetchSchool(s, fresh) {
  const cache = CacheService.getScriptCache();
  const key = 'campus1_' + Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, s.url)
    .map(b => ((b & 0xFF) + 0x100).toString(16).slice(1)).join('');

  if (!fresh) {
    const hit = cache.get(key);
    if (hit) { try { return JSON.parse(hit); } catch (err) { /* 캐시 파손 → 새로 읽기 */ } }
  }

  const out = { name: s.name, url: s.url, ok: false, tabs: {} };
  try {
    const ss = SpreadsheetApp.openByUrl(s.url);
    TAB_NAMES.forEach(t => {
      const sh = ss.getSheetByName(t);
      out.tabs[t] = sh ? trimGrid(sh.getDataRange().getDisplayValues()) : null;
    });
    out.ok = true;
  } catch (err) {
    out.error = '시트를 열 수 없어요 — 스크립트 실행 계정에 열람 권한이 있는지 확인해 주세요';
  }

  try {
    const json = JSON.stringify(out);
    if (json.length < 95000) cache.put(key, json, CACHE_SEC);
  } catch (err) { /* 캐시 실패는 무시 */ }
  return out;
}

/* 뒤쪽의 빈 행·열 제거 + 크기 제한 + 연락처 마스킹 */
function trimGrid(grid) {
  let lastRow = -1, lastCol = -1;
  grid.forEach((row, r) => {
    row.forEach((c, i) => {
      if (String(c).trim()) { lastRow = Math.max(lastRow, r); lastCol = Math.max(lastCol, i); }
    });
  });
  if (lastRow < 0) return [];
  return grid.slice(0, Math.min(lastRow + 1, MAX_ROWS))
    .map(row => row.slice(0, Math.min(lastCol + 1, MAX_COLS)).map(maskCell));
}

function maskCell(v) {
  let s = String(v == null ? '' : v);
  if (!INCLUDE_CONTACTS) {
    /* 이메일: k***@gachon.ac.kr — 단, 셀 전체가 시트/드라이브 링크면 그대로 둠 */
    if (!/^https?:\/\//.test(s.trim())) {
      s = s.replace(/([A-Za-z0-9._%+-])[A-Za-z0-9._%+-]*@([A-Za-z0-9.-]+\.[A-Za-z]{2,})/g, '$1***@$2');
      s = s.replace(/\b(01[016789])[-.\s]?\d{3,4}[-.\s]?(\d{4})\b/g, '$1-****-$2');
      s = s.replace(/\b(\d{4})\d{4,}\b/g, '$1*****'); // 학번 등 8자리 이상 숫자
    }
  }
  return s;
}
