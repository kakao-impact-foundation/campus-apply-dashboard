# 캠퍼스 신청현황 대시보드 (사회혁신가 + 멘토)

테크포임팩트 캠퍼스 26-2 모집 현황을 보여주는 내부용 대시보드 2종.

- **사회혁신가**: https://kakao-impact-foundation.github.io/campus-apply-dashboard/ (비밀번호 게이트)
- **멘토**: https://kakao-impact-foundation.github.io/campus-apply-dashboard/mentor.html (같은 비밀번호)
- 내부 공유용: 화면 앞 비밀번호(캐주얼 차단) + `noindex` 메타 + `robots.txt`로 검색 색인 차단
- 소스에는 신청 데이터가 들어있지 않음 (모든 데이터는 열람 시점에 Apps Script에서 로드)

- **데이터 원본**: [사회혁신가 응답 시트](https://docs.google.com/spreadsheets/d/13T-rtHCFIKb5VHrFL5AMgqNB5KmIHSyEc0KB-C8nrac/edit?gid=229370971) · [멘토 응답 시트](https://docs.google.com/spreadsheets/d/14GF-Du7sDDciiUCKJcsazNzy1ae8nf1MRrzgsfIJYuk/edit?gid=679275159)
- **목표**: 사회혁신가 24팀 (조직 기준, 중복 신청 1팀 카운트) · 멘토 60명
- **디자인**: [campus_2026_spring 파트너 페이지](https://kakao-impact-foundation.github.io/campus_2026_spring/partner/) 톤 — 사회혁신가는 핑크, 멘토는 앰버 액센트로 구분

## 파일

| 파일 | 역할 |
|---|---|
| `index.html` | 사회혁신가 대시보드 본체 |
| `apps-script.gs` | 사회혁신가 시트 연동용 Apps Script (사회혁신가 응답 시트에 설치) |
| `mentor.html` | 멘토 대시보드 본체 (목표 60명 · 직군/소속 분포 · 12개 대학별 희망 현황) |
| `mentor-apps-script.gs` | 멘토 시트 연동용 Apps Script (멘토 응답 시트에 설치) |

## 시트 자동 연동 켜기

1. `apps-script.gs` 상단 주석의 5단계 따라 웹 앱 배포
2. 배포 URL을 `index.html`의 `const DATA_URL = ""` 에 붙여넣기
3. 이후 페이지를 열 때마다 시트 최신 데이터가 자동 반영됨

연동 전에는 `SNAPSHOT`(코드에 박아둔 마지막 저장본)이 표시된다.
클로드에게 "대시보드 스냅샷 업데이트해줘"라고 하면 시트를 읽어 갱신해준다.

## 주요 기능

- 지역 클릭 → 해당 지역 신청 조직 + 조직 리스트 필터링 (신청 수에 따라 지도 음영)
- 활동 분야 키워드 칩 → 조직 리스트 필터링 (지역 필터와 중첩 가능)
- 조직 카드 클릭 → 시트의 모든 응답(소개·참여 이유·풀고 싶은 질문·기대 결과물·연락처 등)을 모달로 표시
- 같은 조직 중복 신청은 1팀으로 합산 (모달에서 신청 건별로 구분 표시)

## 개인정보 주의

GitHub Pages 공개 배포에 맞춰 **연락처(전화·이메일)는 어디에도 싣지 않는다.**
`apps-script.gs`의 `INCLUDE_CONTACTS = false` 상태를 유지할 것 — 연락처가 필요하면 시트에서 직접 확인.
비밀번호 게이트는 검색·우연 유입을 막는 캐주얼 차단이며 실질 인증이 아니므로,
민감 정보를 추가해야 하면 Vercel 등 서버 검증 방식으로 옮길 것.
