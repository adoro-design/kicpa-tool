"""
docgen.py — 문서 자동 생성 모듈
손익분석서(xlsx) / 프로젝트 프로파일(docx) / 개발요청서(docx + xlsx)
"""
import io, re, calendar, zipfile
from datetime import date, timedelta
from typing import List, Optional

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from docx import Document as DocxDocument
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ── 상수 ──────────────────────────────────────────────────────────────
MONTHLY_SALARY    = 8_850_000
PM_RATE           = 0.01
PROD_RATE         = 0.15
STUDIO_UNIT_PRICE = 45_000

PRICE_NEW       = 500_000
PRICE_PORTING   = 50_000
PRICE_EDIT_PORT = 160_000
PRICE_TRAVEL_HR = 100_000

PM_NAME    = "이상현"
PROD_NAME  = "염왕도"
DEPT_CODE  = "SS"
DEPT_FULL  = "에듀테크서비스본부"
COMPANY_NAME = "㈜디유넷"
COMPANY_ADDR = "서울 서대문구 충정로3가 139"
COMPANY_CEO  = "김평국"
CLIENT_NAME  = "공인회계사회 회계연수원"
CLIENT_ADDR  = "서울 서대문구 충정로 2가 185-10"
CLIENT_CEO   = "최운열"
DELIVERY_PLACE = "한국공인회계사회"

# ── 날짜 헬퍼 ─────────────────────────────────────────────────────────
def get_weekday_count(start: date, end: date) -> int:
    count, d = 0, start
    while d <= end:
        if d.weekday() < 5:
            count += 1
        d += timedelta(days=1)
    return count

def get_last_business_day(year: int, month: int) -> date:
    d = date(year, month, calendar.monthrange(year, month)[1])
    while d.weekday() >= 5:
        d -= timedelta(days=1)
    return d

def get_next_month_5th_weekday(year: int, month: int) -> date:
    ny, nm = (year + 1, 1) if month == 12 else (year, month + 1)
    d = date(ny, nm, 5)
    while d.weekday() >= 5:
        d += timedelta(days=1)
    return d

def get_month_number(month_str: str) -> int:
    m = re.search(r'(\d{1,2})월', str(month_str or ""))
    return int(m.group(1)) if m else 1

def fmt_kr(d: date) -> str:
    """2026. 03. 04"""
    return f"{d.year}. {d.month:02d}. {d.day:02d}"

def fmt_kr2(d: date) -> str:
    """2026년 03월 04일"""
    return f"{d.year}년 {d.month:02d}월 {d.day:02d}일"

# ── 단가 헬퍼 ─────────────────────────────────────────────────────────
def classify_fmt(fmt: str):
    """(is_new, is_porting, is_edit_porting)"""
    if not fmt:
        return True, False, False
    if "포팅" in fmt:
        if "편집" in fmt and "무편집" not in fmt:
            return False, True, True
        return False, True, False
    return True, False, False

def get_unit_price_for(content_row, price_tbl: dict) -> int:
    if content_row.custom_price:
        return content_row.custom_price
    fmt = content_row.shooting_format or ""
    if "포팅" in fmt:
        if "편집" in fmt and "무편집" not in fmt:
            return price_tbl.get("편집포팅", PRICE_EDIT_PORT) or PRICE_EDIT_PORT
        return price_tbl.get("포팅", PRICE_PORTING) or PRICE_PORTING
    if "출장" in fmt:
        return price_tbl.get("FullVod (출장)", PRICE_NEW) or PRICE_NEW
    for k, v in price_tbl.items():
        if k and k in fmt:
            return v or 0
    return PRICE_NEW

def get_travel_for(content_row, travel_rate: int = PRICE_TRAVEL_HR) -> int:
    if not content_row.shooting_format or "출장" not in content_row.shooting_format:
        return 0
    if content_row.travel_expense is not None:
        return content_row.travel_expense
    if content_row.travel_hours:
        return content_row.travel_hours * travel_rate
    return travel_rate

# ── 공통 계산 ─────────────────────────────────────────────────────────
def calc_period(courses, year: int, month_num: int):
    starts = [c.shooting_date for c in courses if c.shooting_date]
    ends   = [c.open_date     for c in courses if c.open_date]
    s = min(starts) if starts else date(year, month_num, 1)
    e = max(ends)   if ends   else get_next_month_5th_weekday(year, month_num)
    return s, e

def calc_labor(ps: date, pe: date, year: int, month_num: int):
    m_start = date(year, month_num, 1)
    m_end   = date(year, month_num, calendar.monthrange(year, month_num)[1])
    total_wd  = get_weekday_count(m_start, m_end)
    period_wd = get_weekday_count(ps, pe)
    ratio = period_wd / max(total_wd, 1)
    return (round(PM_RATE * MONTHLY_SALARY * ratio),
            round(PROD_RATE * MONTHLY_SALARY * ratio),
            period_wd, total_wd)

def build_project_name(courses) -> str:
    new_s  = sum(c.session_count or 0 for c in courses if classify_fmt(c.shooting_format or "")[0])
    prt_c  = sum(c.chapter_count or 0 for c in courses if classify_fmt(c.shooting_format or "")[1]
                 and not classify_fmt(c.shooting_format or "")[2])
    eprt_c = sum(c.chapter_count or 0 for c in courses if classify_fmt(c.shooting_format or "")[2])
    parts  = []
    if new_s:  parts.append(f"신규{new_s}차시")
    if prt_c:  parts.append(f"포팅{prt_c}챕터")
    if eprt_c: parts.append(f"편집포팅{eprt_c}챕터")
    return f"한국공인회계사 콘텐츠 개발({'·'.join(parts) if parts else '신규'})"

def calc_revenue(courses, price_tbl: dict) -> int:
    tr = price_tbl.get("1 ~ 4시간", PRICE_TRAVEL_HR)
    return sum(
        get_unit_price_for(c, price_tbl) * (c.session_count or c.chapter_count or 0)
        + get_travel_for(c, tr)
        for c in courses
    )

# ═══════════════════════════════════════════════════════════════════════
# 1. 손익분석서 (Excel)
# ═══════════════════════════════════════════════════════════════════════
def gen_pnl_excel(courses, dept: str, month_str: str, year: int,
                  price_tbl: dict, studio_hours: int, include_studio: bool) -> bytes:
    month_num = get_month_number(month_str)
    ps, pe    = calc_period(courses, year, month_num)
    pm_a, prod_a, pwd, twd = calc_labor(ps, pe, year, month_num)
    write_dt  = get_last_business_day(year, month_num)
    revenue   = calc_revenue(courses, price_tbl)
    studio_a  = studio_hours * STUDIO_UNIT_PRICE if include_studio else 0
    labor_tot = pm_a + prod_a
    exp_tot   = studio_a
    cost_tot  = labor_tot + exp_tot
    profit    = revenue - cost_tot
    profit_r  = profit / revenue if revenue else 0
    outsrc_r  = exp_tot / revenue if revenue else 0
    proj_name = build_project_name(courses)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "손익분석서_V1.0"

    # 열 너비
    for col, w in zip("ABCDEFGHI", [14, 20, 12, 13, 13, 8, 8, 14, 16]):
        ws.column_dimensions[col].width = w

    thin   = Side(style="thin", color="AAAAAA")
    bd     = Border(top=thin, bottom=thin, left=thin, right=thin)
    C      = Alignment(horizontal="center", vertical="center", wrap_text=True)
    R      = Alignment(horizontal="right",  vertical="center")
    L      = Alignment(horizontal="left",   vertical="center", wrap_text=True)
    BL     = PatternFill("solid", start_color="2F5496")
    GR     = PatternFill("solid", start_color="D9E1F2")
    NUM    = '#,##0'
    PCT    = '0.0%'
    DT     = 'YYYY-MM-DD'

    def mc(r, c1, c2=None, val=None, bold=False, sz=9, color="000000", fill=None, align=None, fmt=None):
        """merge + style 헬퍼"""
        if c2 and c2 > c1:
            ws.merge_cells(start_row=r, start_column=c1, end_row=r, end_column=c2)
        cell = ws.cell(r, c1, val)
        cell.font      = Font(bold=bold, size=sz, color=color)
        cell.alignment = align or C
        cell.border    = bd
        if fill: cell.fill = fill
        if fmt:  cell.number_format = fmt
        for col in range(c1+1, (c2 or c1)+1):
            ws.cell(r, col).border = bd
        return cell

    def cv(r, c, val, bold=False, fmt=None, align=None):
        cell = ws.cell(r, c, val)
        cell.font      = Font(bold=bold, size=9)
        cell.border    = bd
        cell.alignment = align or L
        if fmt: cell.number_format = fmt
        return cell

    # ── 제목 ──
    r = 1
    ws.row_dimensions[r].height = 22
    mc(r, 1, 9, f"{year} 손익분석서  ", bold=True, sz=13, align=C)
    r += 1

    # 메타 4행
    for label, value, vfmt in [
        ("버전", 1, None),
        ("프로젝트명", proj_name, None),
        ("작성일", write_dt, DT),
        ("작성자", PM_NAME, None),
    ]:
        mc(r, 1, 4, label, bold=True, align=C)
        cell = mc(r, 5, 9, value, align=L)
        if vfmt: cell.number_format = vfmt
        r += 1

    r += 1  # 공백

    # ── 컬럼 헤더 ──
    mc(r, 1, 3, "항   목", bold=True, color="FFFFFF", fill=BL)
    mc(r, 4, 8, "내   역", bold=True, color="FFFFFF", fill=BL)
    mc(r, 9, 9, "금액(VAT별도)", bold=True, color="FFFFFF", fill=BL)
    r += 1

    # ── 매출 ──
    mc(r, 1, 3, "매출①", bold=True)
    mc(r, 4, 8, "")
    cv(r, 9, revenue, bold=True, fmt=NUM, align=R)
    r += 1

    # ── 직접인건비 제목 ──
    mc(r, 1, 1, "직접\n인건비", bold=True, fill=GR)
    ws.row_dimensions[r].height = 28
    for ci, h in enumerate(["역할","성명","시작일시","종료일시","비율","본부","단가","금액"], 2):
        mc(r, ci, ci, h, bold=True, fill=GR, align=C)
    r += 1

    # PM
    ws.cell(r, 1).border = bd
    for ci, (v, fmt, al) in enumerate([
        ("PM",     None, C), (PM_NAME, None, C),
        (ps, DT, C), (pe, DT, C),
        (PM_RATE, PCT, C), (DEPT_CODE, None, C),
        (MONTHLY_SALARY, NUM, R), (pm_a, NUM, R),
    ], 2):
        cv(r, ci, v, fmt=fmt, align=al)
    r += 1

    # 촬영편집
    ws.cell(r, 1).border = bd
    for ci, (v, fmt, al) in enumerate([
        ("영상촬영 및 편집, 포팅", None, L), (PROD_NAME, None, C),
        (ps, DT, C), (pe, DT, C),
        (PROD_RATE, PCT, C), (DEPT_CODE, None, C),
        (MONTHLY_SALARY, NUM, R), (prod_a, NUM, R),
    ], 2):
        cv(r, ci, v, fmt=fmt, align=al)
    r += 1

    # 인건비 소계
    mc(r, 1, 8, "직접인건비 소계", bold=True,
       align=Alignment(horizontal="right", vertical="center"))
    cv(r, 9, labor_tot, bold=True, fmt=NUM, align=R)
    r += 1

    # ── 직접경비 ──
    mc(r, 1, 1, "직접경비", bold=True, fill=GR)
    ws.row_dimensions[r].height = 20
    for ci, h in enumerate(["구분","담당","단가","투입","단위","본부","비고","금액"], 2):
        mc(r, ci, ci, h, bold=True, fill=GR, align=C)
    r += 1

    if include_studio and studio_hours > 0:
        ws.cell(r, 1).border = bd
        for ci, (v, fmt, al) in enumerate([
            ("스튜디오 대관", None, C), ("주승돈", None, C),
            (STUDIO_UNIT_PRICE, NUM, R), (studio_hours, None, C),
            ("", None, C), (DEPT_CODE, None, C),
            ("", None, C), (studio_a, NUM, R),
        ], 2):
            cv(r, ci, v, fmt=fmt, align=al)
        r += 1
    else:
        mc(r, 2, 9, "")
        r += 1

    # 경비 소계
    mc(r, 1, 8, "직접경비 소계", bold=True,
       align=Alignment(horizontal="right", vertical="center"))
    cv(r, 9, exp_tot, bold=True, fmt=NUM, align=R)
    r += 1

    # ── 손익 요약 ──
    for label, value, fmt in [
        ("비용 합계②", cost_tot, NUM),
        ("외주율",      outsrc_r, PCT),
        ("손익(①-②)", profit,   NUM),
        ("손익률",      profit_r, PCT),
    ]:
        mc(r, 1, 8, label, bold=True,
           align=Alignment(horizontal="right", vertical="center"))
        cv(r, 9, value, bold=True, fmt=fmt, align=R)
        r += 1

    r += 1  # 공백

    # ── 본부별 배분 ──
    for ci, h in enumerate(["구분","코드","본부별배분","본부별 월 인건비"], 2):
        mc(r, ci, ci, h, bold=True, color="FFFFFF", fill=BL, align=C)
    ws.cell(r, 1).border = bd
    r += 1

    for label, code, distr in [(DEPT_FULL, DEPT_CODE, cost_tot), ("총합", "TOT", cost_tot)]:
        ws.cell(r, 1).border = bd
        cv(r, 2, label, align=L)
        cv(r, 3, code,  align=C)
        cv(r, 4, distr, fmt=NUM, align=R)
        cv(r, 5, MONTHLY_SALARY, fmt=NUM, align=R)
        r += 1

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ═══════════════════════════════════════════════════════════════════════
# 2. 개발요청서 Excel 첨부
# ═══════════════════════════════════════════════════════════════════════
def gen_devreq_excel(courses, dept: str, month_str: str, year: int,
                     price_tbl: dict) -> bytes:
    month_num = get_month_number(month_str)
    wb = openpyxl.Workbook()

    # ── Sheet1: 콘텐츠개발내역 ──
    ws1 = wb.active
    ws1.title = "콘텐츠개발내역"

    hfill = PatternFill("solid", start_color="2F5496")
    thin  = Side(style="thin", color="AAAAAA")
    bd    = Border(top=thin, bottom=thin, left=thin, right=thin)
    C     = Alignment(horizontal="center", vertical="center", wrap_text=True)
    R     = Alignment(horizontal="right",  vertical="center")
    NUM   = '#,##0'

    ws1.merge_cells("A1:N1")
    ws1["A1"] = f"{month_num}월 콘텐츠 개발 내역"
    ws1["A1"].font = Font(bold=True, size=12)
    ws1["A1"].alignment = Alignment(horizontal="left", vertical="center")
    ws1.row_dimensions[1].height = 22

    ws1.merge_cells("A2:N2")
    ws1["A2"] = f"* {dept}"
    ws1["A2"].font = Font(bold=True, size=10)

    headers = ["연번","구분","강좌명","강사명","강사소속","시간(차시)","챕터수",
               "등록일\n(수정일)","제작유형","단가기준 유형","퀴즈","단가","금액","총액(VAT 포함)"]
    col_w = [5, 8, 35, 12, 12, 8, 8, 12, 10, 10, 6, 12, 12, 14]
    for i, (h, w) in enumerate(zip(headers, col_w), 1):
        ws1.column_dimensions[get_column_letter(i)].width = w
        cell = ws1.cell(3, i, h)
        cell.font      = Font(bold=True, color="FFFFFF", size=9)
        cell.fill      = hfill
        cell.alignment = C
        cell.border    = bd
    ws1.row_dimensions[3].height = 32

    tr = price_tbl.get("1 ~ 4시간", PRICE_TRAVEL_HR)
    total_sessions = total_chapters = total_amount = 0
    for idx, c in enumerate(courses, 1):
        is_new, is_p, is_ep = classify_fmt(c.shooting_format or "")
        구분    = "포팅" if (is_p or is_ep) else "신규"
        unit_p  = get_unit_price_for(c, price_tbl)
        qty     = c.session_count or c.chapter_count or 0
        amount  = unit_p * qty + get_travel_for(c, tr)
        fmt_name = c.shooting_format or ""

        row_data = [idx, 구분, c.course_name, c.instructor, "",
                    c.session_count or "", c.chapter_count or "",
                    c.open_date or "", fmt_name, fmt_name,
                    "", unit_p, amount, round(amount * 1.1)]

        total_sessions += c.session_count or 0
        total_chapters += c.chapter_count or 0
        total_amount   += amount

        for ci, v in enumerate(row_data, 1):
            cell = ws1.cell(idx + 3, ci, v)
            cell.font      = Font(size=9)
            cell.border    = bd
            cell.alignment = C if ci in (1, 2, 6, 7, 8, 9, 10, 11) else \
                             R if ci in (12, 13, 14) else \
                             Alignment(horizontal="left", vertical="center")
            if ci in (12, 13, 14): cell.number_format = NUM
            if ci == 8 and isinstance(v, date):
                cell.number_format = "YYYY-MM-DD"

    # 합계행
    tr_row = len(courses) + 4
    ws1.merge_cells(f"A{tr_row}:E{tr_row}")
    ws1.cell(tr_row, 1, "전체 합계").font = Font(bold=True, size=9)
    ws1.cell(tr_row, 1).alignment = C
    ws1.cell(tr_row, 1).border = bd
    for ci in range(2, 6):
        ws1.cell(tr_row, ci).border = bd
    for ci, v in [(6, total_sessions), (7, total_chapters),
                  (13, total_amount), (14, round(total_amount * 1.1))]:
        cell = ws1.cell(tr_row, ci, v)
        cell.font      = Font(bold=True, size=9)
        cell.border    = bd
        cell.number_format = NUM
        cell.alignment = C
    for ci in [8, 9, 10, 11, 12]:
        ws1.cell(tr_row, ci).border = bd

    # ── Sheet2: 단가 ──
    ws2 = wb.create_sheet("단가")
    ws2.column_dimensions["B"].width = 30
    ws2.column_dimensions["C"].width = 18
    ws2.merge_cells("B1:D1")
    ws2["B1"] = "콘텐츠 유형별 개발 단가"
    ws2["B1"].font = Font(bold=True, size=10)

    for r_idx, (label, val_) in enumerate([
        ("콘텐츠유형", "단가(vat별도)/차시"),
        ("크로마키",   500000), ("FullVod(출장)", 500000),
        ("태블릿형",   500000), ("전자칠판형",    500000),
        ("", ""),
        ("외부과정 포팅 단가(삼일, 조세일보, 휴넷 등)", ""),
        ("콘텐츠유형", "단가(vat별도)/챕터"),
        ("포팅(동영상 무편집)", 50000), ("포팅(동영상 편집)", 160000),
        ("", ""),
        ("출장비(서울소재 - 촬영감독 1명)", ""),
        ("기준시간(1일)", "단가(vat별도)/챕터"),
        ("1~4시간", 100000), ("4시간 초과", "별도 협의"),
    ], 2):
        ws2.cell(r_idx, 2, label).font = Font(size=9)
        c = ws2.cell(r_idx, 3, val_)
        c.font = Font(size=9)
        if isinstance(val_, int):
            c.number_format = NUM
            c.alignment = R

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ═══════════════════════════════════════════════════════════════════════
# 3. 개발요청서 Word
# ═══════════════════════════════════════════════════════════════════════
def _set_cell(cell, text, bold=False, size=10, align=WD_ALIGN_PARAGRAPH.LEFT,
              bg_color=None):
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    p = cell.paragraphs[0]
    p.alignment = align
    run = p.add_run(text)
    run.bold = bold
    run.font.size = Pt(size)
    if bg_color:
        tc   = cell._tc
        tcPr = tc.get_or_add_tcPr()
        shd  = OxmlElement("w:shd")
        shd.set(qn("w:val"), "clear")
        shd.set(qn("w:color"), "auto")
        shd.set(qn("w:fill"), bg_color)
        tcPr.append(shd)

def _set_col_width(table, col_idx, width_cm):
    for row in table.rows:
        row.cells[col_idx].width = Cm(width_cm)

def gen_devreq_docx(courses, dept: str, month_str: str, year: int,
                    ps: date, pe: date, write_dt: date) -> bytes:
    month_num = get_month_number(month_str)
    doc = DocxDocument()

    # 페이지 여백
    for section in doc.sections:
        section.top_margin    = Cm(2.0)
        section.bottom_margin = Cm(2.0)
        section.left_margin   = Cm(2.5)
        section.right_margin  = Cm(2.5)

    # ── 헤더 표 ──
    tbl = doc.add_table(rows=7, cols=4)
    tbl.style = "Table Grid"
    tbl.autofit = False

    # 행1: 제목 (병합)
    row0 = tbl.rows[0]
    row0.cells[0].merge(row0.cells[3])
    _set_cell(row0.cells[0], "개  발  요  청  서", bold=True, size=14,
              align=WD_ALIGN_PARAGRAPH.CENTER, bg_color="D9E1F2")

    # 행2: 발주처 / 개발사
    for ci, txt in [(0, "발  주  처"), (2, "개  발  사")]:
        tbl.rows[1].cells[ci].merge(tbl.rows[1].cells[ci+1])
        _set_cell(tbl.rows[1].cells[ci], txt, bold=True,
                  align=WD_ALIGN_PARAGRAPH.CENTER, bg_color="D9E1F2")

    # 행3~5: 발주처/개발사 정보
    info = [
        ("회사명", CLIENT_NAME, "회사명", COMPANY_NAME),
        ("주소",   CLIENT_ADDR, "주소",   COMPANY_ADDR),
        ("대표자", CLIENT_CEO,  "대표자", COMPANY_CEO),
    ]
    for ri, (l1, v1, l2, v2) in enumerate(info, 2):
        _set_cell(tbl.rows[ri].cells[0], l1, bold=True,
                  align=WD_ALIGN_PARAGRAPH.CENTER, bg_color="F2F2F2")
        _set_cell(tbl.rows[ri].cells[1], v1, size=9)
        _set_cell(tbl.rows[ri].cells[2], l2, bold=True,
                  align=WD_ALIGN_PARAGRAPH.CENTER, bg_color="F2F2F2")
        _set_cell(tbl.rows[ri].cells[3], v2, size=9)

    # 행6: 시행일
    tbl.rows[5].cells[0].merge(tbl.rows[5].cells[1])
    tbl.rows[5].cells[2].merge(tbl.rows[5].cells[3])
    _set_cell(tbl.rows[5].cells[0], "시행일", bold=True,
              align=WD_ALIGN_PARAGRAPH.CENTER, bg_color="F2F2F2")
    _set_cell(tbl.rows[5].cells[2], fmt_kr2(write_dt), size=9,
              align=WD_ALIGN_PARAGRAPH.CENTER)

    # 행7: 제목
    tbl.rows[6].cells[0].merge(tbl.rows[6].cells[1])
    tbl.rows[6].cells[2].merge(tbl.rows[6].cells[3])
    _set_cell(tbl.rows[6].cells[0], "제목", bold=True,
              align=WD_ALIGN_PARAGRAPH.CENTER, bg_color="F2F2F2")
    _set_cell(tbl.rows[6].cells[2], "교육콘텐츠 개발요청", size=9)

    doc.add_paragraph()  # 간격

    # ── 본문 ──
    def add_para(text, bold=False, size=10, indent_cm=0):
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Cm(indent_cm)
        run = p.add_run(text)
        run.bold = bold
        run.font.size = Pt(size)
        return p

    add_para(f"*{dept}", bold=True)
    doc.add_paragraph()
    add_para(f"● 개발기간 : {fmt_kr2(ps)} ~ {fmt_kr2(pe)}", indent_cm=1)
    add_para(f"● 납품장소 : {DELIVERY_PLACE}", indent_cm=1)
    doc.add_paragraph()
    add_para("1. 귀 사의 일익 번창함을 기원합니다.", indent_cm=1)
    add_para("2. 상기와 같이 동영상 제작을 요청하오니 기한 내에 납품하여 주시기 바랍니다.",
             indent_cm=1)
    doc.add_paragraph()

    # 날짜 + 서명
    p_date = doc.add_paragraph()
    p_date.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_date.add_run(fmt_kr2(write_dt)).font.size = Pt(10)

    p_sign = doc.add_paragraph()
    p_sign.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_sign.add_run(DELIVERY_PLACE).bold = True

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


# ═══════════════════════════════════════════════════════════════════════
# 4. 프로젝트 프로파일 Word
# ═══════════════════════════════════════════════════════════════════════
def gen_profile_docx(courses, dept: str, month_str: str, year: int,
                     price_tbl: dict, studio_hours: int, include_studio: bool,
                     customer_contact=None) -> bytes:
    month_num = get_month_number(month_str)
    ps, pe    = calc_period(courses, year, month_num)
    pm_a, prod_a, pwd, twd = calc_labor(ps, pe, year, month_num)
    write_dt  = get_last_business_day(year, month_num)
    revenue   = calc_revenue(courses, price_tbl)
    studio_a  = studio_hours * STUDIO_UNIT_PRICE if include_studio else 0
    proj_name = build_project_name(courses)

    # 세부내역 구성
    new_s  = sum(c.session_count or 0 for c in courses if classify_fmt(c.shooting_format or "")[0])
    prt_c  = sum(c.chapter_count or 0 for c in courses
                 if classify_fmt(c.shooting_format or "")[1] and not classify_fmt(c.shooting_format or "")[2])
    eprt_c = sum(c.chapter_count or 0 for c in courses if classify_fmt(c.shooting_format or "")[2])
    travel_h = sum((c.travel_hours or 0) for c in courses
                   if c.shooting_format and "출장" in c.shooting_format)

    detail_lines = ["1. 사업내용", "   (1) 한공회 콘텐츠 개발"]
    if new_s:
        detail_lines.append(f"   - 신규: {new_s}차시(단가: ₩500,000, VAT별도) / 유상개발")
    if prt_c:
        detail_lines.append(f"   - 포팅(무편집): {prt_c}챕터(단가: ₩50,000, VAT별도) / 유상개발")
    if eprt_c:
        detail_lines.append(f"   - 포팅(편집): {eprt_c}챕터(단가: ₩160,000, VAT별도) / 유상개발")
    if travel_h:
        detail_lines.append(f"   - 출장비: {travel_h}시간(단가: ₩100,000, VAT별도)")
    detail_lines.append(f"   (2) 개발기간 : {fmt_kr(ps)} ~ {fmt_kr(pe)}")
    detail_text = "\n".join(detail_lines)

    contact_name  = customer_contact.contact_name if customer_contact else ""
    contact_phone = customer_contact.phone        if customer_contact else ""
    contact_email = customer_contact.email        if customer_contact else ""

    doc = DocxDocument()
    for section in doc.sections:
        section.top_margin    = Cm(2.0)
        section.bottom_margin = Cm(2.0)
        section.left_margin   = Cm(2.5)
        section.right_margin  = Cm(2.5)

    # ── 제목 ──
    h = doc.add_paragraph()
    h.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = h.add_run("프로젝트 프로파일")
    r.bold = True
    r.font.size = Pt(14)
    r.underline = True

    doc.add_paragraph()

    # 문서번호/사업명/작성자/작성일 표
    meta_tbl = doc.add_table(rows=2, cols=4)
    meta_tbl.style = "Table Grid"
    for ci, (l, v) in enumerate([("문서번호", ""), ("사업명", "한공회 콘텐츠 개발")]):
        _set_cell(meta_tbl.rows[0].cells[ci*2],   l, bold=True,
                  align=WD_ALIGN_PARAGRAPH.CENTER, bg_color="D9E1F2", size=9)
        _set_cell(meta_tbl.rows[0].cells[ci*2+1], v, size=9)
    for ci, (l, v) in enumerate([("작성자(SR)", PM_NAME), ("작성일", fmt_kr(write_dt))]):
        _set_cell(meta_tbl.rows[1].cells[ci*2],   l, bold=True,
                  align=WD_ALIGN_PARAGRAPH.CENTER, bg_color="D9E1F2", size=9)
        _set_cell(meta_tbl.rows[1].cells[ci*2+1], v, size=9)

    doc.add_paragraph()

    # ── 1. 개요 ──
    h1 = doc.add_paragraph()
    h1.add_run("1. 개요").bold = True

    note_p = doc.add_paragraph()
    note_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    note_p.add_run("(단위 : 원, VAT별도)").font.size = Pt(8)

    # 개요 표 (프로젝트, 기간, 금액, 세부내역, 고객정보)
    ov = doc.add_table(rows=8, cols=4)
    ov.style = "Table Grid"
    ov.autofit = False

    # 프로젝트명
    ov.rows[0].cells[0].merge(ov.rows[0].cells[1])
    ov.rows[0].cells[2].merge(ov.rows[0].cells[3])
    _set_cell(ov.rows[0].cells[0], "프로젝트명", bold=True,
              align=WD_ALIGN_PARAGRAPH.CENTER, bg_color="D9E1F2", size=9)
    _set_cell(ov.rows[0].cells[2], proj_name, size=9)

    # 기간
    ov.rows[1].cells[0].merge(ov.rows[1].cells[1])
    ov.rows[1].cells[2].merge(ov.rows[1].cells[3])
    _set_cell(ov.rows[1].cells[0], "기간", bold=True,
              align=WD_ALIGN_PARAGRAPH.CENTER, bg_color="D9E1F2", size=9)
    _set_cell(ov.rows[1].cells[2], f"{fmt_kr(ps)} ~ {fmt_kr(pe)}", size=9)

    # 계약금액
    ov.rows[2].cells[0].merge(ov.rows[2].cells[1])
    ov.rows[2].cells[2].merge(ov.rows[2].cells[3])
    _set_cell(ov.rows[2].cells[0], "계약금액", bold=True,
              align=WD_ALIGN_PARAGRAPH.CENTER, bg_color="D9E1F2", size=9)
    _set_cell(ov.rows[2].cells[2], f"₩{revenue:,}", size=9)

    # 금액 구분
    _set_cell(ov.rows[3].cells[0], "금액", bold=True,
              align=WD_ALIGN_PARAGRAPH.CENTER, bg_color="D9E1F2", size=9)
    _set_cell(ov.rows[3].cells[1], "선급금", bold=True,
              align=WD_ALIGN_PARAGRAPH.CENTER, bg_color="F2F2F2", size=9)
    _set_cell(ov.rows[3].cells[2], "중도금", bold=True,
              align=WD_ALIGN_PARAGRAPH.CENTER, bg_color="F2F2F2", size=9)
    _set_cell(ov.rows[3].cells[3], "잔금",   bold=True,
              align=WD_ALIGN_PARAGRAPH.CENTER, bg_color="F2F2F2", size=9)

    _set_cell(ov.rows[4].cells[0], "", size=9)
    _set_cell(ov.rows[4].cells[1], "계약체결시", size=9,
              align=WD_ALIGN_PARAGRAPH.CENTER)
    _set_cell(ov.rows[4].cells[2], "중간보고시", size=9,
              align=WD_ALIGN_PARAGRAPH.CENTER)
    _set_cell(ov.rows[4].cells[3], "납품완료시", size=9,
              align=WD_ALIGN_PARAGRAPH.CENTER)

    _set_cell(ov.rows[5].cells[0], "", size=9)
    _set_cell(ov.rows[5].cells[1], "", size=9)
    _set_cell(ov.rows[5].cells[2], "", size=9)
    _set_cell(ov.rows[5].cells[3], f"₩{revenue:,}", size=9,
              align=WD_ALIGN_PARAGRAPH.RIGHT)

    # 세부내역
    _set_cell(ov.rows[6].cells[0], "세부내역", bold=True,
              align=WD_ALIGN_PARAGRAPH.CENTER, bg_color="D9E1F2", size=9)
    ov.rows[6].cells[1].merge(ov.rows[6].cells[3])
    _set_cell(ov.rows[6].cells[1], detail_text, size=9)

    # 고객정보
    ov.rows[7].cells[0].merge(ov.rows[7].cells[3])
    _set_cell(ov.rows[7].cells[0], "고객정보", bold=True,
              align=WD_ALIGN_PARAGRAPH.CENTER, bg_color="D9E1F2", size=9)

    # 고객 상세 (별도 표)
    cust = doc.add_table(rows=2, cols=6)
    cust.style = "Table Grid"
    for ci, h_ in enumerate(["고객명", "담당자", "설명", "연락처", "핸드폰", "E-Mail"]):
        _set_cell(cust.rows[0].cells[ci], h_, bold=True,
                  align=WD_ALIGN_PARAGRAPH.CENTER, bg_color="D9E1F2", size=9)
    _set_cell(cust.rows[1].cells[0], "한국공인회계사회", size=9)
    _set_cell(cust.rows[1].cells[1], contact_name, size=9)
    _set_cell(cust.rows[1].cells[2], dept, size=9)
    _set_cell(cust.rows[1].cells[3], contact_phone, size=9)
    _set_cell(cust.rows[1].cells[4], "", size=9)
    _set_cell(cust.rows[1].cells[5], contact_email, size=9)

    doc.add_paragraph()

    # ── 2. 참여인력 ──
    h2 = doc.add_paragraph()
    h2.add_run("2. 참여인력(계획)").bold = True

    pm_rate_pct  = round(PM_RATE   * 100 * pwd / max(twd, 1))
    prod_rate_pct = round(PROD_RATE * 100 * pwd / max(twd, 1))
    period_str = f"{str(ps.year)[2:]}.{ps.month:02d}.{ps.day:02d} ~ {str(pe.year)[2:]}.{pe.month:02d}.{pe.day:02d}"

    staff_tbl = doc.add_table(rows=4, cols=5)
    staff_tbl.style = "Table Grid"
    for ci, h_ in enumerate(["역할","성명","소속","참여기간","참여율(%)"]):
        _set_cell(staff_tbl.rows[0].cells[ci], h_, bold=True,
                  align=WD_ALIGN_PARAGRAPH.CENTER, bg_color="D9E1F2", size=9)
    for ri, (role, name, rate) in enumerate([
        ("PM", PM_NAME, f"{pm_rate_pct}%"),
        ("영상촬영 및 편집,\n포팅", PROD_NAME, f"{prod_rate_pct}%"),
        ("", "", ""),
    ], 1):
        _set_cell(staff_tbl.rows[ri].cells[0], role, size=9,
                  align=WD_ALIGN_PARAGRAPH.CENTER)
        _set_cell(staff_tbl.rows[ri].cells[1], name, size=9,
                  align=WD_ALIGN_PARAGRAPH.CENTER)
        _set_cell(staff_tbl.rows[ri].cells[2], "서비스 운영팀", size=9,
                  align=WD_ALIGN_PARAGRAPH.CENTER)
        _set_cell(staff_tbl.rows[ri].cells[3], period_str if name else "", size=9,
                  align=WD_ALIGN_PARAGRAPH.CENTER)
        _set_cell(staff_tbl.rows[ri].cells[4], rate, size=9,
                  align=WD_ALIGN_PARAGRAPH.CENTER)

    doc.add_paragraph()

    # ── 3. 외부협력 ──
    h3 = doc.add_paragraph()
    h3.add_run("3. 외부협력(계획)").bold = True
    note3 = doc.add_paragraph()
    note3.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    note3.add_run("(단위 : 원, VAT별도)").font.size = Pt(8)

    ext_tbl = doc.add_table(rows=5, cols=6)
    ext_tbl.style = "Table Grid"
    for ci, h_ in enumerate(["부문","업체","담당자","사무실","핸드폰","E-Mail"]):
        _set_cell(ext_tbl.rows[0].cells[ci], h_, bold=True,
                  align=WD_ALIGN_PARAGRAPH.CENTER, bg_color="D9E1F2", size=9)
    for ri in range(1, 5):
        for ci in range(6):
            _set_cell(ext_tbl.rows[ri].cells[ci], "", size=9)

    doc.add_paragraph()

    # ── 4. 장비 및 시설 ──
    h4 = doc.add_paragraph()
    h4.add_run("4. 장비 및 시설").bold = True
    note4 = doc.add_paragraph()
    note4.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    note4.add_run("(단위 : 원, VAT별도)").font.size = Pt(8)

    eq_tbl = doc.add_table(rows=7, cols=5)
    eq_tbl.style = "Table Grid"
    for ci, h_ in enumerate(["내용","수량(시간)","단가","금액","비고"]):
        _set_cell(eq_tbl.rows[0].cells[ci], h_, bold=True,
                  align=WD_ALIGN_PARAGRAPH.CENTER, bg_color="D9E1F2", size=9)

    if include_studio and studio_hours > 0:
        _set_cell(eq_tbl.rows[1].cells[0], "스튜디오 대관", size=9)
        _set_cell(eq_tbl.rows[1].cells[1], str(studio_hours), size=9,
                  align=WD_ALIGN_PARAGRAPH.CENTER)
        _set_cell(eq_tbl.rows[1].cells[2], f"{STUDIO_UNIT_PRICE:,}", size=9,
                  align=WD_ALIGN_PARAGRAPH.RIGHT)
        _set_cell(eq_tbl.rows[1].cells[3], f"{studio_a:,}", size=9,
                  align=WD_ALIGN_PARAGRAPH.RIGHT)
        _set_cell(eq_tbl.rows[1].cells[4], "", size=9)
        start_row = 2
    else:
        start_row = 1

    for ri in range(start_row, 7):
        for ci in range(5):
            _set_cell(eq_tbl.rows[ri].cells[ci], "", size=9)

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


# ═══════════════════════════════════════════════════════════════════════
# 전체 ZIP 생성
# ═══════════════════════════════════════════════════════════════════════
def generate_all(courses, dept: str, month_str: str, year: int,
                 price_tbl: dict, studio_hours: int, include_studio: bool,
                 customer_contact=None) -> bytes:
    month_num = get_month_number(month_str)
    ps, pe    = calc_period(courses, year, month_num)
    write_dt  = get_last_business_day(year, month_num)
    m2        = f"{month_num:02d}"
    dept_short = dept.replace(" ", "")

    pnl_xlsx    = gen_pnl_excel(courses, dept, month_str, year,
                                price_tbl, studio_hours, include_studio)
    req_xlsx    = gen_devreq_excel(courses, dept, month_str, year, price_tbl)
    req_docx    = gen_devreq_docx(courses, dept, month_str, year, ps, pe, write_dt)
    profile_docx = gen_profile_docx(courses, dept, month_str, year,
                                    price_tbl, studio_hours, include_studio,
                                    customer_contact)

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr(f"손익분석서_{dept_short}_{year}{m2}.xlsx", pnl_xlsx)
        zf.writestr(f"프로젝트프로파일_{dept_short}_{year}{m2}.docx", profile_docx)
        zf.writestr(f"개발요청서_{dept_short}_{year}{m2}.docx", req_docx)
        zf.writestr(f"개발요청서첨부_{dept_short}_{year}{m2}.xlsx", req_xlsx)
    return buf.getvalue()
