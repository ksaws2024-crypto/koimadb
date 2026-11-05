"""HS 코드 정비 워크북(v16) 데이터정비 스크립트.

사용법: python3 clean_hs_workbook.py <입력.xlsx> <출력.xlsx>

정비 항목
  1. 일련번호: 시트마다 숫자/문자 혼재 → 모든 시트 정수로 통일
  2. 업체명(한글): 전각 괄호·공백을 반각으로, 연속 공백 1칸으로, 괄호 안쪽 공백 제거
  3. 자동 필터 범위: 실제 데이터 끝 행과 어긋난 시트 보정
  4. 1차_적재(반영) 중 HS 품목이 있는데 12대=12(기타 서비스/물류)인 행 → 주 HS류 기준 분류로 교정
  5. 보류 확인 계획: 업체 수를 고정값 → COUNTIFS 수식으로 교체(+합계)
  6. 신규 시트: '정비 내역'(요약·변경 로그·확인 필요 사항), '적재_통합'(DB 반영 대상 일원화)
원본 서식(머리글 색, 틀 고정, 열 너비)은 그대로 둔다.
"""
import re
import sys
from collections import Counter, defaultdict

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

SRC, DST = sys.argv[1], sys.argv[2]
DATA_SHEETS = ["1차_적재(반영)", "1차_적재_12대(미처리)", "보류(협회 확인 후)", "미처리(현재값 유지)"]
PLAN = "보류 확인 계획"
HOLD = "보류(협회 확인 후)"
LOAD = "1차_적재(반영)"
CAT12 = "1차_적재_12대(미처리)"
APP_SCALES = {
    "1억불 초과", "5000만불~1억불", "3000만불~5000만불", "1000만불~3000만불", "700만불~1000만불",
    "500만불~700만불", "300만불~500만불", "100만불~300만불", "50만불~100만불", "30만불~50만불",
    "10만불~30만불", "5만불~10만불", "3만불~5만불", "1만불~3만불", "1만불 미만",
}

HEAD_FONT = Font(name="Calibri", bold=True)
HEAD_FILL = PatternFill("solid", fgColor="E1EFD3")
NOTE_FILL = PatternFill("solid", fgColor="FFF2CC")


def norm_name(v: str) -> str:
    s = v.replace("（", "(").replace("）", ")").replace("　", " ")
    s = re.sub(r"\s{2,}", " ", s)
    s = re.sub(r"\(\s+", "(", s)
    s = re.sub(r"\s+\)", ")", s)
    return s.strip()


def header_map(ws):
    return {c.value: c.column for c in ws[1] if c.value is not None}


def last_data_row(ws):
    r = ws.max_row
    while r > 1 and all(c.value in (None, "") for c in ws[r]):
        r -= 1
    return r


wb = openpyxl.load_workbook(SRC)
log = []  # (시트, 행, 일련번호, 업체명, 항목, 변경 전, 변경 후, 사유)
stats = Counter()
missing_code = defaultdict(int)
scale_issues = Counter()

for name in DATA_SHEETS:
    ws = wb[name]
    hm = header_map(ws)
    last = last_data_row(ws)
    cs, cn, cc = hm["일련번호"], hm["업체명(한글)"], hm["업체 코드"]
    for r in range(2, last + 1):
        sc = ws.cell(r, cs)
        if isinstance(sc.value, str) and sc.value.strip().isdigit():
            sc.value = int(sc.value.strip())
            stats[(name, "일련번호 문자→숫자")] += 1
        nc = ws.cell(r, cn)
        if isinstance(nc.value, str):
            new = norm_name(nc.value)
            if new != nc.value:
                log.append((name, r, sc.value, new, "업체명(한글)", nc.value, new, "공백·괄호 정리"))
                nc.value = new
                stats[(name, "업체명 공백·괄호 정리")] += 1
        if ws.cell(r, cc).value in (None, ""):
            missing_code[name] += 1
    # 수입규모 라벨 점검(앱 필터 라벨과 다른 구간)
    for col in ("수입규모", "수입규모 코드명(한글)"):
        if col in hm:
            for r in range(2, last + 1):
                v = ws.cell(r, hm[col]).value
                if v and v not in APP_SCALES:
                    scale_issues[(name, v)] += 1
    # 자동 필터 범위 보정
    want = f"A1:{get_column_letter(ws.max_column)}{last}"
    if ws.auto_filter.ref != want:
        log.append((name, "", "", "", "자동 필터 범위", ws.auto_filter.ref, want, "실제 데이터 범위와 불일치"))
        ws.auto_filter.ref = want
        stats[(name, "자동 필터 범위 보정")] += 1

# 4. 1차_적재(반영): 품목이 있는데 12대=12 → 주 HS류(2자리)의 다수 분류로 교정
ws = wb[LOAD]
hm = header_map(ws)
last = last_data_row(ws)
ch_cat = defaultdict(Counter)
code_of = {}
for r in range(2, last + 1):
    hs = ws.cell(r, hm["제안 HS 전체(대표순)"]).value or ""
    cat = ws.cell(r, hm["제안 12대 분류"]).value
    code_of[cat] = ws.cell(r, hm["제안 12대 코드"]).value
    for t in hs.split(","):
        if t.strip():
            ch_cat[t.strip()[:2]][cat] += 1
for r in range(2, last + 1):
    code = ws.cell(r, hm["제안 12대 코드"]).value
    if str(code) == "12":
        hs = ws.cell(r, hm["제안 HS 전체(대표순)"]).value
        votes = Counter()
        for t in hs.split(","):
            for k, v in ch_cat[t.strip()[:2]].items():
                if k != "기타 서비스/물류":
                    votes[k] += v
        cat = votes.most_common(1)[0][0]
        old = f"{code} {ws.cell(r, hm['제안 12대 분류']).value}"
        ws.cell(r, hm["제안 12대 코드"]).value = type(code)(code_of[cat]) if not isinstance(code, int) else int(code_of[cat])
        ws.cell(r, hm["제안 12대 분류"]).value = cat
        log.append((LOAD, r, ws.cell(r, hm["일련번호"]).value, ws.cell(r, hm["업체명(한글)"]).value, "제안 12대",
                    old, f"{code_of[cat]} {cat}", f"HS {hs} 품목 보유 — 12(물품 근거 없음) 부적합, 해당 HS류 다수 분류 적용"))
        stats[(LOAD, "12대 분류 교정")] += 1

# 5. 보류 확인 계획: 업체 수 → COUNTIFS 수식
wp = wb[PLAN]
hold_last = last_data_row(wb[HOLD])
rng_a = f"'{HOLD}'!$A$2:$A${hold_last}"
rng_b = f"'{HOLD}'!$B$2:$B${hold_last}"
first = last_row = None
for r in range(2, wp.max_row + 1):
    a, b, c = wp.cell(r, 1).value, wp.cell(r, 2).value, wp.cell(r, 3).value
    if a is None:
        break
    if isinstance(c, (int, float)):
        wp.cell(r, 3).value = f'=COUNTIFS({rng_a},$A{r},{rng_b},$B{r})'
        first = first or r
        last_row = r
tot = last_row + 1
wp.insert_rows(tot)
wp.cell(tot, 1).value = "합계"
wp.cell(tot, 1).font = HEAD_FONT
wp.cell(tot, 3).value = f"=SUM(C{first}:C{last_row})"
wp.cell(tot, 3).font = HEAD_FONT
wp.cell(tot, 2).value = f"검증: 보류 시트 전체 행 수"
wp.cell(tot, 4).value = f"=COUNTA({rng_a})"
wp.cell(tot, 5).value = f'=IF(C{tot}=D{tot},"일치","불일치")'
wp.auto_filter.ref = f"A1:C{last_row}"
stats[(PLAN, "업체 수 고정값→COUNTIFS 수식")] += last_row - first + 1

# 6-a. 적재_통합: DB 반영 대상을 한 시트로
wi = wb.create_sheet("적재_통합", 0)
cols = ["반영 구분", "업체 코드", "일련번호", "업체명(한글)", "HS 코드(대표순)", "세부 상품명(한글)",
        "세부 상품명(영문)", "12대 코드", "12대 분류", "HS코드 출처", "신뢰도", "원본 시트"]
wi.append(cols)
n_hs = n_12 = 0
for r in range(2, last + 1):
    g = lambda k: ws.cell(r, hm[k]).value
    wi.append(["HS·상품명·12대 교체", g("업체 코드"), g("일련번호"), g("업체명(한글)"), g("제안 HS 전체(대표순)"),
               g("제안 세부 상품명(한글)"), g("제안 세부 상품명(영문)"), int(g("제안 12대 코드")), g("제안 12대 분류"),
               g("HS코드 출처(제안)"), g("신뢰도"), LOAD])
    n_hs += 1
w2 = wb[CAT12]
h2 = header_map(w2)
for r in range(2, last_data_row(w2) + 1):
    g = lambda k: w2.cell(r, h2[k]).value
    wi.append(["12대만 교체(HS 유지)", g("업체 코드"), g("일련번호"), g("업체명(한글)"), None, None, None,
               int(g("제안 12대 코드")), g("제안 12대 분류"), g("HS코드 출처(제안)"), None, CAT12])
    n_12 += 1
for c in wi[1]:
    c.font, c.fill = HEAD_FONT, HEAD_FILL
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
for i, w in enumerate([20, 12, 10, 26, 40, 40, 50, 9, 20, 11, 8, 22], 1):
    wi.column_dimensions[get_column_letter(i)].width = w
wi.freeze_panes = "E2"
wi.auto_filter.ref = f"A1:L{wi.max_row}"

# 6-b. 정비 내역
wl = wb.create_sheet("정비 내역", 0)
bold = Font(name="Calibri", bold=True)
rows = [["HS 코드 정비 워크북 — 데이터정비 내역 (v16 → v17, 2026-11-05)"], [],
        ["1. 정비 요약"], ["시트", "정비 항목", "건수"]]
for (s, k), v in sorted(stats.items()):
    rows.append([s, k, v])
rows += [[], ["2. 적재_통합 시트 구성"], ["반영 구분", "건수", "비고"],
         ["HS·상품명·12대 교체", n_hs, f"{LOAD} 전체"],
         ["12대만 교체(HS 유지)", n_12, f"{CAT12} 전체 — HS·상품명은 현재값 유지"],
         ["합계", n_hs + n_12, "보류·미처리(현재값 유지) 시트는 적재 대상 아님"],
         [], ["3. 확인 필요 사항(자동 수정하지 않음)"], ["구분", "내용", "건수"]]
for s in DATA_SHEETS:
    if missing_code[s]:
        rows.append(["업체 코드 없음", f"{s} — DB 매칭 키가 없어 일련번호·업체명으로 매칭 필요", missing_code[s]])
for (s, v), n in sorted(scale_issues.items()):
    rows.append(["수입규모 라벨 불일치", f"{s} — '{v}' (사이트 필터 라벨에 없는 구간)", n])
rows += [[], ["4. 변경 로그"], ["시트", "행", "일련번호", "업체명(한글)", "항목", "변경 전", "변경 후", "사유"]]
rows += [list(x) for x in log]
for row in rows:
    wl.append(row)
wl["A1"].font = Font(name="Calibri", bold=True, size=14)
for r in range(1, wl.max_row + 1):
    v = wl.cell(r, 1).value
    if isinstance(v, str) and re.match(r"^\d\. ", v):
        wl.cell(r, 1).font = Font(name="Calibri", bold=True, size=12)
    elif v in ("시트", "반영 구분", "구분"):
        for c in wl[r]:
            if c.value is not None:
                c.font, c.fill = bold, HEAD_FILL
    elif v in ("업체 코드 없음", "수입규모 라벨 불일치"):
        for c in wl[r][:3]:
            c.fill = NOTE_FILL
for i, w in enumerate([26, 60, 12, 30, 16, 40, 40, 60], 1):
    wl.column_dimensions[get_column_letter(i)].width = w

wb.save(DST)
print("stats", dict(stats))
print("적재_통합", n_hs, n_12, "log", len(log), "missing", dict(missing_code))
print("scale issues", dict(scale_issues))
