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
const STATUS_SHEET = '멘토심사'; // 합불 심사(합격·보류·불합격) 저장 탭 (없으면 자동 생성)
const ROSTER_SHEET = '멘토 참여자 정보'; // 26-1 명단 형식의 자동 생성 탭
const SCHOOL_COLORS = { /* [배경, 글자색] — 26-1 명단의 매칭 학교 열 톤 */
  '가천대': ['#F4CCCC', '#990000'], '경운대': ['#FCE5CD', '#B45F06'], '고려대(세종)': ['#FFF2CC', '#7F6000'],
  '동국대': ['#D9EAD3', '#274E13'], '부산외대': ['#D0E0E3', '#0C343D'], '서울대': ['#C9DAF8', '#1C4587'],
  '시립대': ['#CFE2F3', '#073763'], '서울여대': ['#EAD1DC', '#741B47'], '한라대': ['#D9D2E9', '#20124D'],
  'GIST': ['#E6B8AF', '#85200C'], 'KAIST': ['#B6D7A8', '#38761D'], 'UNIST': ['#A2C4C9', '#134F5C'],
  '미배정': ['#EFEFEF', '#666666']
};
const SCHOOL_SHORT = {
  '가천대학교': '가천대', '경운대학교': '경운대', '고려대학교 세종캠퍼스': '고려대(세종)',
  '동국대학교': '동국대', '부산외국어대학교': '부산외대', '서울대학교': '서울대',
  '서울시립대학교': '시립대', '서울여자대학교': '서울여대', '한라대학교': '한라대',
  '광주과학기술원 GIST': 'GIST', '한국과학기술원 KAIST': 'KAIST', '울산과학기술원 UNIST': 'UNIST'
};

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
  /* '추천 크루 LDAP'가 별도 열일 때만 내보내요 (iRef와 같은 열이면 중복 방지) */
  const iRefLdap0 = header.findIndex(h => String(h).includes('추천 크루') && String(h).includes('LDAP'));
  const iRefLdap = iRefLdap0 === iRef ? -1 : iRefLdap0;

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
      referral: str(r, iRef),
      refLdap: str(r, iRefLdap)
    })).filter(r => r.name);

  const payload = {
    v: 7, // 코드 버전 (배포 확인용 — 7: 합불 심사 저장 + 추천 크루 LDAP)
    updatedAt: Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm'),
    rows: rows,
    assignments: readAssignments(ss),
    statuses: readStatuses(ss)
  };

  /* 신청 인원이 달라졌으면 '멘토 참여자 정보' 탭도 갱신 */
  try {
    const nPeople = new Set(rows.map(r => r.ldap || r.name)).size;
    const roster = ss.getSheetByName(ROSTER_SHEET);
    if (!roster || roster.getLastRow() - 2 !== nPeople) rebuildRoster(ss);
  } catch (err) { /* 명단 갱신 실패해도 대시보드 응답은 정상 반환 */ }

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

/* ── 합불 심사 (대시보드 심사 버튼) ─────────────────────────
 * 합격·보류·불합격 버튼을 누르면 doPost로 들어와 '멘토심사' 탭에 저장돼요.
 * 같은 LDAP은 덮어쓰기, 상태 ''(해제)는 행 삭제.
 */

function readStatuses(ss) {
  const sh = ss.getSheetByName(STATUS_SHEET);
  const out = {};
  if (!sh || sh.getLastRow() < 2) return out;
  sh.getDataRange().getValues().slice(1).forEach(r => {
    const ldap = String(r[1] || '').trim();
    const status = String(r[3] || '').trim();
    if (ldap && status) out[ldap] = status;
  });
  return out;
}

/* 시트 탭에서 LDAP으로 행을 찾아 upsert/삭제 (멘토배정·멘토심사 공용) */
function upsertByLdap(ss, sheetName, headerRow, ldap, values, remove) {
  let sh = ss.getSheetByName(sheetName);
  if (!sh) {
    sh = ss.insertSheet(sheetName);
    sh.appendRow(headerRow);
  }
  const data = sh.getDataRange().getValues();
  let rowIdx = -1;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1]).trim() === ldap) { rowIdx = i + 1; break; }
  }
  if (remove) {
    if (rowIdx > 0) sh.deleteRow(rowIdx);
  } else if (rowIdx > 0) {
    sh.getRange(rowIdx, 1, 1, values.length).setValues([values]);
  } else {
    sh.appendRow(values);
  }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const body = JSON.parse(e.postData.contents);
    const ldap = String(body.ldap || '').trim().slice(0, 100);
    const name = String(body.name || '').trim().slice(0, 50);
    if (!ldap) return jsonOut({ ok: false, error: 'invalid ldap' });
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    if (body.action === 'assign') {
      const school = String(body.school || '').trim().slice(0, 60);
      upsertByLdap(ss, ASSIGN_SHEET, ['수정시각', 'LDAP', '이름', '배정대학'],
        ldap, [new Date(), ldap, name, school], school === '');
      try { rebuildRoster(ss); } catch (err) { /* 명단 갱신 실패는 무시 */ }
      return jsonOut({ ok: true });
    }

    if (body.action === 'status') {
      const status = String(body.status || '').trim();
      if (!['합격', '보류', '불합격', ''].includes(status)) {
        return jsonOut({ ok: false, error: 'invalid status' });
      }
      upsertByLdap(ss, STATUS_SHEET, ['수정시각', 'LDAP', '이름', '심사상태'],
        ldap, [new Date(), ldap, name, status], status === '');
      return jsonOut({ ok: true });
    }

    return jsonOut({ ok: false, error: 'invalid action' });
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

/* ── '멘토 참여자 정보' 탭 자동 생성 ─────────────────────────
 * 26-1 카카오멘토 명단 형식. 배정 변경(doPost)·인원 변동(doGet) 때마다 다시 생성돼요.
 * 자동 생성 탭이므로 직접 수정하지 마세요 — 예외로 '팀장' 열은 LDAP 기준으로 보존됩니다.
 */

function rebuildRoster(ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheets().find(s => s.getSheetId() === SHEET_GID) || ss.getSheets()[0];
  const values = sheet.getDataRange().getValues();
  const header = values[0];
  const col = (keyword) => header.findIndex(h => String(h).includes(keyword));
  const iName = col('성함');
  const iLdap = header.findIndex(h => String(h).trim() === 'LDAP');
  const iCompany = col('공동체명'), iDept = col('부서명'), iJob = col('직군');
  const iExp = col('참여 경험'), iPhone = col('핸드폰'), iEmail = col('회사 Email');
  const iAlma = col('모교 희망'), iWork = col('맡고 계신 업무'), iCoach = col('코칭 가능한');
  const str = (row, i) => i >= 0 ? String(row[i] || '').trim() : '';

  /* 같은 LDAP 재제출은 최신 응답만 (뒤 행이 최신) */
  const byLdap = new Map();
  values.slice(1).forEach(r => { if (str(r, iName)) byLdap.set(str(r, iLdap) || str(r, iName), r); });
  const assign = readAssignments(ss);

  /* 기존 탭의 '팀장' 열(F) 값을 LDAP 기준으로 보존 */
  const existing = ss.getSheetByName(ROSTER_SHEET);
  const lead = {};
  if (existing && existing.getLastRow() > 2) {
    existing.getDataRange().getValues().slice(2).forEach(r => {
      const l = String(r[3] || '').trim();
      if (l && String(r[5] || '').trim()) lead[l] = String(r[5]).trim();
    });
  }

  /* 카테고리: 직군 + 업무·코칭 내용 키워드로 자동 분류 (26-1은 AI 분류였음) */
  const category = (job, txt) => {
    if (/기획|디자/.test(job)) return '기획 · 디자인';
    if (/AI/i.test(job)) return 'AI/ML';
    if (/ios|android|안드로이드|모바일|스위프트|swift|kotlin.*(앱|클라이언트)|앱 개발/i.test(txt)) return '모바일 (Android/iOS)';
    if (/프론트|front|react|vue|next\.js|웹 ?개발|javascript|typescript/i.test(txt)) return '프론트엔드 (FE)';
    if (/데이터 ?(엔지니어|파이프라인|분석|플랫폼)|hadoop|spark|kafka/i.test(txt)) return '데이터 엔지니어링';
    if (/개발/.test(job)) return '서버/플랫폼 (BE/Infra)';
    return '기타';
  };

  const rows = [...byLdap.values()].map(r => {
    const ldap = str(r, iLdap);
    const full = assign[ldap] || '';
    return {
      school: full ? (SCHOOL_SHORT[full] || full) : '미배정',
      unassigned: !full,
      name: str(r, iName), ldap: ldap,
      company: str(r, iCompany), dept: str(r, iDept), job: str(r, iJob),
      cat: category(str(r, iJob), str(r, iWork) + ' ' + str(r, iCoach)),
      phone: str(r, iPhone), email: str(r, iEmail),
      re: str(r, iExp).includes('예') ? 'O' : '',
      alma: str(r, iAlma)
    };
  });
  const grp = s => /^[가-힣]/.test(s) ? 0 : 1; /* 한글 대학 먼저, 영문 뒤 */
  rows.sort((a, b) =>
    (a.unassigned - b.unassigned) ||
    (grp(a.school) - grp(b.school)) ||
    a.school.localeCompare(b.school, 'ko') ||
    a.name.localeCompare(b.name, 'ko'));

  const out = [
    ['26-2학기 테크포임팩트 캠퍼스 카카오멘토 명단 (자동 생성 — 배정은 대시보드 배정 보드에서)', '', '', '', '', '', '', '', '', '', '', '', '', ''],
    ['순번', '매칭 학교', '이름', 'LDAP', '이름 (LDAP)', '팀장', '공동체', '소속 부서 (최하위 조직명)', '직군', '카테고리', '전화번호', '이메일', '재참여', '모교 희망']
  ];
  rows.forEach((p, i) => out.push([
    i + 1, p.school, p.name, p.ldap, p.name + ' (' + p.ldap + ')', lead[p.ldap] || '',
    p.company, p.dept, p.job, p.cat, p.phone, p.email, p.re, p.alma
  ]));

  const sh = existing || ss.insertSheet(ROSTER_SHEET);
  const nRows = rows.length;
  const NC = 14;

  /* ── 빠른 경로: 인원 수가 그대로면 값 + 매칭 학교 색만 갱신 (서식은 유지됨) ── */
  const prevRows = existing ? existing.getLastRow() - 2 : -1;
  if (existing && prevRows === nRows && existing.getLastColumn() >= NC) {
    sh.getRange(1, 1, out.length, NC).setValues(out);
    const bg = rows.map(p => [(SCHOOL_COLORS[p.school] || SCHOOL_COLORS['미배정'])[0]]);
    const fc = rows.map(p => [(SCHOOL_COLORS[p.school] || SCHOOL_COLORS['미배정'])[1]]);
    sh.getRange(3, 2, nRows, 1).setBackgrounds(bg).setFontColors(fc);
    return;
  }

  /* ── 전체 재생성 (인원 변동 시): 값 + 26-1 명단 스타일 서식 ── */
  sh.clear();
  sh.getRange(1, 1, out.length, 14).setValues(out);
  /* 제목 */
  sh.getRange(1, 1).setFontWeight('bold').setFontSize(12);
  /* 헤더: 검정 배경 + 흰 볼드 + 가운데 */
  sh.getRange(2, 1, 1, NC).setBackground('#000000').setFontColor('#FFFFFF')
    .setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle');
  if (nRows > 0) {
    const body = sh.getRange(3, 1, nRows, NC);
    body.setVerticalAlignment('middle');
    /* 배경색: 매칭 학교(B) 파스텔, 이름(C)·이름(LDAP)(E) 크림, LDAP(D) 시안 */
    const bgs = [], fcs = [], fws = [], has = [];
    rows.forEach(p => {
      const c = SCHOOL_COLORS[p.school] || SCHOOL_COLORS['미배정'];
      bgs.push(['#FFFFFF', c[0], '#FFF2CC', '#5FF0F0', '#FFF2CC', '#FFFFFF', '#FFFFFF', '#FFFFFF', '#FFFFFF', '#FFFFFF', '#FFFFFF', '#FFFFFF', '#FFFFFF', '#FFFFFF']);
      fcs.push(['#000000', c[1], '#000000', '#000000', '#000000', '#000000', '#000000', '#000000', '#000000', '#000000', '#000000', '#000000', '#000000', '#000000']);
      fws.push(['normal', 'bold', 'normal', 'bold', 'normal', 'normal', 'normal', 'normal', 'normal', 'normal', 'normal', 'normal', 'normal', 'normal']);
      has.push(['center', 'center', 'center', 'left', 'left', 'center', 'center', 'center', 'center', 'center', 'center', 'left', 'center', 'center']);
    });
    body.setBackgrounds(bgs).setFontColors(fcs).setFontWeights(fws).setHorizontalAlignments(has);
  }
  /* 테두리: 표 전체 얇은 검정 */
  sh.getRange(2, 1, nRows + 1, NC)
    .setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  /* 열 너비 + 머리글 고정 */
  const widths = [50, 110, 70, 110, 165, 50, 135, 185, 95, 155, 115, 210, 55, 135];
  widths.forEach((w, i) => sh.setColumnWidth(i + 1, w));
  sh.setFrozenRows(2);

  sh.getRange(1, 1).setNote('이 탭은 자동 생성돼요 — 배정 보드에서 배정을 바꾸거나 새 신청이 들어오면 다시 만들어집니다.\n직접 수정하면 지워져요 (예외: "팀장" 열은 LDAP 기준으로 보존).');
}
