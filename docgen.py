"""
docgen.py — 문서 자동 생성 모듈 (템플릿 기반)
손익분석서(xlsx) / 프로젝트 프로파일(docx) / 개발요청서(docx + xlsx)
"""
import io, re, calendar, zipfile, os
from copy import copy
from datetime import date, timedelta
from typing import Optional

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from docx import Document as DocxDocument
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

TEMPLATE_DIR = os.path.join(os.path.dirname(__file__), 'doc_templates')

# ── 상수 ──────────────────────────────────────────────────────────────
MONTHLY_SALARY    = 8_850_000
PM_RATE           = 0.01      # 기본 참여율 (조정 가능)
PROD_RATE         = 0.15
STUDIO_UNIT_PRICE = 45_000
TARGET_PROFIT     = 0.40      # 최소 손익률

PRICE_NEW       = 500_000
PRICE_PORTING   = 50_000
PRICE_EDIT_PORT = 160_000
PRICE_TRAVEL_HR = 100_000

PM_NAME    = "이상현"
PROD_NAME  = "염왕도"
DEPT_CODE  = "SS"
DEPT_FULL  = "에듀테크서비스본부"
COMPANY_NAME  = "㈜디유넷"
COMPANY_ADDR  = "서울 서대문구 충정로3가 139"
COMPANY_CEO   = "김평국"
CLIENT_ADDR   = "서울 서대문구 충정로 2가 185-10"
CLIENT_CEO    = "최운열"
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

def fmt_kr(d: date) -> str:   return f"{d.year}. {d.month:02d}. {d.day:02d}"
def fmt_kr2(d: date) -> str:  return f"{d.year}년 {d.month:02d}월 {d.day:02d}일"
def fmt_short(d: date) -> str:
    return f"{str(d.year)[2:]}.{d.month:02d}.{d.day:02d}"

# ── 단가 헬퍼 ─────────────────────────────────────────────────────────
def classify_fmt(fmt: str):
    if not fmt: return True, False, False
    if "포팅" in fmt:
        if "편집" in fmt and "무편집" not in fmt: return False, True, True
        return False, True, False
    return True, False, False

def get_unit_price_for(content_row, price_tbl: dict) -> int:
    if content_row.custom_price: return content_row.custom_price
    fmt = content_row.shooting_format or ""
    if "포팅" in fmt:
        if "편집" in fmt and "무편집" not in fmt:
            return price_tbl.get("편집포팅", PRICE_EDIT_PORT) or PRICE_EDIT_PORT
        return price_tbl.get("포팅", PRICE_PORTING) or PRICE_PORTING
    if "출장" in fmt:
        return price_tbl.get("FullVod (출장)", PRICE_NEW) or PRICE_NEW
    for k, v in price_tbl.items():
        if k and k in fmt: return v or 0
    return PRICE_NEW

def get_travel_for(content_row, travel_rate: int = PRICE_TRAVEL_HR) -> int:
    if not content_row.shooting_format or "출장" not in content_row.shooting_format: return 0
    if content_row.travel_expense is not None: return content_row.travel_expense
    if content_row.travel_hours: return content_row.travel_hours * travel_rate
    return travel_rate

# ── 공통 계산 ─────────────────────────────────────────────────────────
def calc_period(courses, year: int, month_num: int):
    starts = [c.shooting_date for c in courses if c.shooting_date]
    ends   = [c.open_date     for c in courses if c.open_date]
    s = min(starts) if starts else date(year, month_num, 1)
    e = max(ends)   if ends   else get_next_month_5th_weekday(year, month_num)
    return s, e

def calc_revenue(courses, price_tbl: dict) -> int:
    tr = price_tbl.get("1 ~ 4시간", PRICE_TRAVEL_HR)
    return sum(
        get_unit_price_for(c, price_tbl) * (c.session_count or c.chapter_count or 0)
        + get_travel_for(c, tr)
        for c in courses
    )

def calc_labor_amounts(ps: date, pe: date, pm_rate: float, prod_rate: float):
    """템플릿 공식과 동일: (pe-ps).days / 30 × rate × 월인건비 + 1"""
    d = (pe - ps).days
    pm_a   = round(d / 30 * pm_rate   * MONTHLY_SALARY + 1)
    prod_a = round(d / 30 * prod_rate * MONTHLY_SALARY + 1)
    return pm_a, prod_a

def adjust_rates(revenue: int, studio_a: int, ps: date, pe: date) -> tuple:
    """손익률 40% 이상이 되도록 참여율 자동 조정. 두 문서에 동일 적용."""
    pm_a, prod_a = calc_labor_amounts(ps, pe, PM_RATE, PROD_RATE)
    total_cost   = pm_a + prod_a + studio_a
    if revenue > 0 and (revenue - total_cost) / revenue >= TARGET_PROFIT:
        return PM_RATE, PROD_RATE  # 기본 비율로 달성

    # 비율 축소
    max_labor = (1 - TARGET_PROFIT) * revenue - studio_a - 2
    period_d  = (pe - ps).days
    base_labor = period_d / 30 * (PM_RATE + PROD_RATE) * MONTHLY_SALARY
    if max_labor <= 0 or base_labor <= 0:
        return PM_RATE, PROD_RATE
    scale     = max_labor / base_labor
    new_pm    = max(round(PM_RATE   * scale, 4), 0.001)
    new_prod  = max(round(PROD_RATE * scale, 4), 0.001)
    return new_pm, new_prod

def build_project_name(courses) -> str:
    new_s  = sum(c.session_count or 0 for c in courses if classify_fmt(c.shooting_format or "")[0])
    prt_c  = sum(c.chapter_count or 0 for c in courses
                 if classify_fmt(c.shooting_format or "")[1]
                 and not classify_fmt(c.shooting_format or "")[2])
    eprt_c = sum(c.chapter_count or 0 for c in courses if classify_fmt(c.shooting_format or "")[2])
    parts  = []
    if new_s:  parts.append(f"신규{new_s}차시")
    if prt_c:  parts.append(f"포팅{prt_c}챕터")
    if eprt_c: parts.append(f"편집포팅{eprt_c}챕터")
    return f"한국공인회계사 콘텐츠 개발({'·'.join(parts) if parts else '신규'})"

# ── openpyxl 스타일 복사 헬퍼 ─────────────────────────────────────────
def _copy_cell_style(src, dst):
    dst.font        = copy(src.font)
    dst.border      = copy(src.border)
    dst.fill        = copy(src.fill)
    dst.number_format = src.number_format
    dst.alignment   = copy(src.alignment)

# ── Word 텍스트 치환 헬퍼 ─────────────────────────────────────────────
def _replace_para(para, old: str, new: str) -> bool:
    """단락 전체 텍스트에서 old→new 치환. 첫 번째 런에 모아서 처리."""
    full = para.text
    if old not in full:
        return False
    replaced = full.replace(old, new)
    if para.runs:
        para.runs[0].text = replaced
        for run in para.runs[1:]:
            run.text = ""
    else:
        para.add_run(replaced)
    return True

def _replace_doc(doc, old: str, new: str):
    """문서 전체(단락 + 표 셀)에서 텍스트 치환"""
    for para in doc.paragraphs:
        _replace_para(para, old, new)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    _replace_para(para, old, new)


# ═══════════════════════════════════════════════════════════════════════
# 1. 손익분석서 (Excel) — 템플릿 기반
# ═══════════════════════════════════════════════════════════════════════
def gen_pnl_excel(courses, dept: str, month_str: str, year: int,
                  price_tbl: dict, studio_hours: int, include_studio: bool,
                  pm_rate: float, prod_rate: float) -> bytes:
    month_num = get_month_number(month_str)
    ps, pe    = calc_period(courses, year, month_num)
    write_dt  = get_last_business_day(year, month_num)
    revenue   = calc_revenue(courses, price_tbl)
    studio_a  = studio_hours * STUDIO_UNIT_PRICE if include_studio else 0
    proj_name = build_project_name(courses)

    wb = openpyxl.load_workbook(os.path.join(TEMPLATE_DIR, 'tpl_pnl.xlsx'))
    ws = wb.active

    # 프로젝트명 & 작성일
    ws['F3'] = proj_name
    ws['F4'] = write_dt
    ws['F4'].number_format = 'YYYY"년" MM"월" DD"일"'

    # 매출
    ws['I8'] = revenue

    # PM 인건비
    ws['D10'] = ps;  ws['D10'].number_format = 'YYYY-MM-DD'
    ws['E10'] = pe;  ws['E10'].number_format = 'YYYY-MM-DD'
    ws['F10'] = pm_rate

    # 촬영편집 인건비
    ws['D11'] = ps;  ws['D11'].number_format = 'YYYY-MM-DD'
    ws['E11'] = pe;  ws['E11'].number_format = 'YYYY-MM-DD'
    ws['F11'] = prod_rate

    # 스튜디오 대관 (직접경비)
    ws['D17'] = STUDIO_UNIT_PRICE if include_studio else 0
    ws['E17'] = studio_hours       if include_studio else 0

    # 월 인건비 기준 (수식 참조용)
    ws['E32'] = MONTHLY_SALARY

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ═══════════════════════════════════════════════════════════════════════
# 2. 개발요청서 Excel 첨부 — 템플릿 기반
# ═══════════════════════════════════════════════════════════════════════
def gen_devreq_excel(courses, dept: str, month_str: str, year: int,
                     price_tbl: dict) -> bytes:
    month_num = get_month_number(month_str)
    write_dt  = get_last_business_day(year, month_num)

    wb = openpyxl.load_workbook(os.path.join(TEMPLATE_DIR, 'tpl_devreq.xlsx'))
    ws = wb['콘텐츠개발내역']

    # 헤더 업데이트
    ws['A1'] = f"{month_num}월 콘텐츠 개발 내역"
    ws['N1'] = f"작성일 : {write_dt.strftime('%Y-%m-%d')}"
    ws['A2'] = f"* {dept}"

    # 기존 데이터 행 개수 파악 (row 4부터 합계행 전까지)
    tpl_data_start = 4
    tpl_total_row  = tpl_data_start
    while ws.cell(tpl_total_row, 1).value != '전체 합계':
        tpl_total_row += 1
    tpl_data_count = tpl_total_row - tpl_data_start

    # 스타일 참조용 셀 저장 (첫 데이터 행)
    style_row = tpl_data_start

    n = len(courses)

    # 행 수 조정
    if n > tpl_data_count:
        ws.insert_rows(tpl_total_row, n - tpl_data_count)
    elif n < tpl_data_count:
        ws.delete_rows(tpl_data_start + n, tpl_data_count - n)

    # 새 데이터 작성
    tr = price_tbl.get("1 ~ 4시간", PRICE_TRAVEL_HR)
    for idx, c in enumerate(courses):
        r = tpl_data_start + idx
        is_new, is_p, is_ep = classify_fmt(c.shooting_format or "")
        unit_p = get_unit_price_for(c, price_tbl)
        qty    = c.session_count if not (is_p or is_ep) else c.chapter_count or 0

        row_vals = {
            1: idx + 1,
            2: "포팅" if (is_p or is_ep) else "신규",
            3: c.course_name,
            4: c.instructor,
            5: "",
            6: c.session_count or "",
            7: c.chapter_count or "",
            8: c.open_date,
            9: c.shooting_format or "",
            10: c.shooting_format or "",
            11: "",
            12: unit_p,
            13: f"=L{r}*F{r}",
            14: f"=M{r}*1.1",
        }
        for col, val in row_vals.items():
            cell = ws.cell(r, col, val)
            if col == 8 and isinstance(val, date):
                cell.number_format = 'YYYY-MM-DD'
            if col == 12:
                cell.number_format = '#,##0'

    # 합계 행
    total_row = tpl_data_start + n
    end_row   = total_row - 1
    ws.cell(total_row, 1, "전체 합계")
    ws.cell(total_row, 6, f"=SUM(F{tpl_data_start}:F{end_row})")
    ws.cell(total_row, 7, f"=SUM(G{tpl_data_start}:G{end_row})")
    ws.cell(total_row, 13, f"=SUM(M{tpl_data_start}:M{end_row})").number_format = '#,##0'
    ws.cell(total_row, 14, f"=SUM(N{tpl_data_start}:N{end_row})").number_format = '#,##0'

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ═══════════════════════════════════════════════════════════════════════
# 3. 개발요청서 Word — 템플릿 기반 (텍스트 치환)
# ═══════════════════════════════════════════════════════════════════════
def gen_devreq_docx(courses, dept: str, month_str: str, year: int,
                    ps: date, pe: date, write_dt: date) -> bytes:
    doc = DocxDocument(os.path.join(TEMPLATE_DIR, 'tpl_devreq.docx'))
    table = doc.tables[0]

    # 시행일 (row 5, cells[1..3])
    old_date = "2026년 03월 03일"
    for cell in table.rows[5].cells[1:]:
        for para in cell.paragraphs:
            _replace_para(para, old_date, fmt_kr2(ps))

    # 본문 (row 7, all cells)
    replacements = [
        ("조세지원본부",   dept),
        ("2026년 03월 04일", fmt_kr2(ps)),
        ("2026년 03월 27일", fmt_kr2(pe)),
        ("2026년 03월 16일", fmt_kr2(write_dt)),
    ]
    for cell in table.rows[7].cells:
        for old, new in replacements:
            for para in cell.paragraphs:
                _replace_para(para, old, new)

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


# ═══════════════════════════════════════════════════════════════════════
# 4. 프로젝트 프로파일 Word — python-docx 생성 (템플릿 손상으로 직접 생성)
# ═══════════════════════════════════════════════════════════════════════
def _cell(cell, text, bold=False, size=9, align=WD_ALIGN_PARAGRAPH.LEFT, bg=None):
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    p = cell.paragraphs[0]
    p.alignment = align
    run = p.add_run(text)
    run.bold      = bold
    run.font.size = Pt(size)
    if bg:
        tc   = cell._tc
        tcPr = tc.get_or_add_tcPr()
        shd  = OxmlElement("w:shd")
        shd.set(qn("w:val"), "clear")
        shd.set(qn("w:color"), "auto")
        shd.set(qn("w:fill"), bg)
        tcPr.append(shd)

HDR_BG  = "D9E1F2"
SUB_BG  = "F2F2F2"
CTR = WD_ALIGN_PARAGRAPH.CENTER
RGT = WD_ALIGN_PARAGRAPH.RIGHT

def gen_profile_docx(courses, dept: str, month_str: str, year: int,
                     price_tbl: dict, studio_hours: int, include_studio: bool,
                     pm_rate: float, prod_rate: float,
                     customer_contact=None) -> bytes:
    month_num = get_month_number(month_str)
    ps, pe    = calc_period(courses, year, month_num)
    write_dt  = get_last_business_day(year, month_num)
    revenue   = calc_revenue(courses, price_tbl)
    studio_a  = studio_hours * STUDIO_UNIT_PRICE if include_studio else 0
    proj_name = build_project_name(courses)
    period_d  = (pe - ps).days

    # 참여율 (손익분석서와 동일한 비율 사용)
    pm_pct   = round(pm_rate   * 100, 1)
    prod_pct = round(prod_rate * 100, 1)
    period_str = f"{fmt_short(ps)} ~ {fmt_short(pe)}"

    # 세부내역
    new_s  = sum(c.session_count or 0 for c in courses if classify_fmt(c.shooting_format or "")[0])
    prt_c  = sum(c.chapter_count or 0 for c in courses
                 if classify_fmt(c.shooting_format or "")[1]
                 and not classify_fmt(c.shooting_format or "")[2])
    eprt_c = sum(c.chapter_count or 0 for c in courses if classify_fmt(c.shooting_format or "")[2])
    travel_h = sum((c.travel_hours or (1 if c.shooting_format and "출장" in c.shooting_format else 0))
                   for c in courses if c.shooting_format and "출장" in c.shooting_format)

    detail_lines = []
    if new_s:   detail_lines.append(f"   - 신규: {new_s}차시(단가: ₩500,000, VAT별도) / 유상개발")
    if prt_c:   detail_lines.append(f"   - 포팅(무편집): {prt_c}챕터(단가: ₩50,000, VAT별도) / 유상개발")
    if eprt_c:  detail_lines.append(f"   - 포팅(편집): {eprt_c}챕터(단가: ₩160,000, VAT별도) / 유상개발")
    if travel_h: detail_lines.append(f"   - 출장비: {travel_h}시간(단가: ₩100,000, VAT별도)")
    detail_text = (
        f"1. 사업내용\n   (1) 한공회 콘텐츠 개발\n"
        + "\n".join(detail_lines)
        + f"\n   (2) 개발기간 : {fmt_kr(ps)} ~ {fmt_kr(pe)}"
    )

    contact_name  = customer_contact.contact_name if customer_contact else ""
    contact_phone = customer_contact.phone        if customer_contact else ""
    contact_email = customer_contact.email        if customer_contact else ""

    doc = DocxDocument()
    for sec in doc.sections:
        sec.top_margin = sec.bottom_margin = Cm(2.0)
        sec.left_margin = sec.right_margin = Cm(2.5)

    # ── 제목 ──────────────────────────────────────────────────────────
    title_p = doc.add_paragraph()
    title_p.alignment = CTR
    r = title_p.add_run("프로젝트 프로파일")
    r.bold = True; r.font.size = Pt(14); r.underline = True
    doc.add_paragraph()

    # 문서번호/사업명/작성자/작성일
    m = doc.add_table(rows=2, cols=4)
    m.style = "Table Grid"
    for ci, (l, v) in enumerate([("문서번호", ""), ("사업명", "한공회 콘텐츠 개발")]):
        _cell(m.rows[0].cells[ci*2],   l, bold=True, align=CTR, bg=HDR_BG)
        _cell(m.rows[0].cells[ci*2+1], v)
    for ci, (l, v) in enumerate([("작성자(SR)", PM_NAME), ("작성일", fmt_kr(write_dt))]):
        _cell(m.rows[1].cells[ci*2],   l, bold=True, align=CTR, bg=HDR_BG)
        _cell(m.rows[1].cells[ci*2+1], v)
    doc.add_paragraph()

    # ── 1. 개요 ───────────────────────────────────────────────────────
    doc.add_paragraph().add_run("1. 개요").bold = True
    note = doc.add_paragraph()
    note.alignment = RGT
    note.add_run("(단위 : 원, VAT별도)").font.size = Pt(8)

    ov = doc.add_table(rows=8, cols=4)
    ov.style = "Table Grid"

    def ov_merge(row_idx, label, value_text):
        ov.rows[row_idx].cells[0].merge(ov.rows[row_idx].cells[1])
        ov.rows[row_idx].cells[2].merge(ov.rows[row_idx].cells[3])
        _cell(ov.rows[row_idx].cells[0], label, bold=True, align=CTR, bg=HDR_BG)
        _cell(ov.rows[row_idx].cells[2], value_text)

    ov_merge(0, "프로젝트명", proj_name)
    ov_merge(1, "기간",       f"{fmt_kr(ps)} ~ {fmt_kr(pe)}")
    ov_merge(2, "계약금액",   f"₩{revenue:,}")

    # 금액 행
    _cell(ov.rows[3].cells[0], "금액", bold=True, align=CTR, bg=HDR_BG)
    for ci, h in enumerate(["선급금","중도금","잔금"], 1):
        _cell(ov.rows[3].cells[ci], h, bold=True, align=CTR, bg=SUB_BG)
    _cell(ov.rows[4].cells[0], "", bg=HDR_BG)
    for ci, h in enumerate(["계약체결시","중간보고시","납품완료시"], 1):
        _cell(ov.rows[4].cells[ci], h, align=CTR)
    _cell(ov.rows[5].cells[0], "", bg=HDR_BG)
    _cell(ov.rows[5].cells[1], "")
    _cell(ov.rows[5].cells[2], "")
    _cell(ov.rows[5].cells[3], f"₩{revenue:,}", align=RGT)

    # 세부내역
    _cell(ov.rows[6].cells[0], "세부내역", bold=True, align=CTR, bg=HDR_BG)
    ov.rows[6].cells[1].merge(ov.rows[6].cells[3])
    _cell(ov.rows[6].cells[1], detail_text)

    # 고객정보
    ov.rows[7].cells[0].merge(ov.rows[7].cells[3])
    _cell(ov.rows[7].cells[0], "고객정보", bold=True, align=CTR, bg=HDR_BG)

    cust = doc.add_table(rows=2, cols=6)
    cust.style = "Table Grid"
    for ci, h in enumerate(["고객명","담당자","설명","연락처(☎)","핸드폰","E-Mail"]):
        _cell(cust.rows[0].cells[ci], h, bold=True, align=CTR, bg=HDR_BG)
    _cell(cust.rows[1].cells[0], "한국공인회계사회")
    _cell(cust.rows[1].cells[1], contact_name)
    _cell(cust.rows[1].cells[2], dept)
    _cell(cust.rows[1].cells[3], contact_phone)
    _cell(cust.rows[1].cells[4], "")
    _cell(cust.rows[1].cells[5], contact_email)
    doc.add_paragraph()

    # ── 2. 참여인력 ───────────────────────────────────────────────────
    doc.add_paragraph().add_run("2. 참여인력(계획)").bold = True
    staff = doc.add_table(rows=4, cols=5)
    staff.style = "Table Grid"
    for ci, h in enumerate(["역할","성명","소속","참여기간","참여율(%)"]):
        _cell(staff.rows[0].cells[ci], h, bold=True, align=CTR, bg=HDR_BG)
    people = [
        ("PM",                       PM_NAME,   f"{pm_pct}%"),
        ("영상촬영 및 편집,\n포팅",    PROD_NAME, f"{prod_pct}%"),
        ("",                          "",        ""),
    ]
    for ri, (role, name, pct) in enumerate(people, 1):
        _cell(staff.rows[ri].cells[0], role,            align=CTR)
        _cell(staff.rows[ri].cells[1], name,            align=CTR)
        _cell(staff.rows[ri].cells[2], "서비스 운영팀" if name else "", align=CTR)
        _cell(staff.rows[ri].cells[3], period_str if name else "", align=CTR)
        _cell(staff.rows[ri].cells[4], pct,             align=CTR)
    doc.add_paragraph()

    # ── 3. 외부협력 ───────────────────────────────────────────────────
    doc.add_paragraph().add_run("3. 외부협력(계획)").bold = True
    note3 = doc.add_paragraph(); note3.alignment = RGT
    note3.add_run("(단위 : 원, VAT별도)").font.size = Pt(8)
    ext = doc.add_table(rows=5, cols=6)
    ext.style = "Table Grid"
    for ci, h in enumerate(["부문","업체","담당자","사무실","핸드폰","E-Mail"]):
        _cell(ext.rows[0].cells[ci], h, bold=True, align=CTR, bg=HDR_BG)
    for ri in range(1, 5):
        for ci in range(6): _cell(ext.rows[ri].cells[ci], "")
    doc.add_paragraph()

    # ── 4. 장비 및 시설 ───────────────────────────────────────────────
    doc.add_paragraph().add_run("4. 장비 및 시설").bold = True
    note4 = doc.add_paragraph(); note4.alignment = RGT
    note4.add_run("(단위 : 원, VAT별도)").font.size = Pt(8)
    eq = doc.add_table(rows=7, cols=5)
    eq.style = "Table Grid"
    for ci, h in enumerate(["내용","수량(시간)","단가","금액","비고"]):
        _cell(eq.rows[0].cells[ci], h, bold=True, align=CTR, bg=HDR_BG)
    if include_studio and studio_hours > 0:
        _cell(eq.rows[1].cells[0], "스튜디오 대관")
        _cell(eq.rows[1].cells[1], str(studio_hours), align=CTR)
        _cell(eq.rows[1].cells[2], f"{STUDIO_UNIT_PRICE:,}", align=RGT)
        _cell(eq.rows[1].cells[3], f"{studio_a:,}", align=RGT)
        _cell(eq.rows[1].cells[4], "")
        start = 2
    else:
        start = 1
    for ri in range(start, 7):
        for ci in range(5): _cell(eq.rows[ri].cells[ci], "")

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
    revenue   = calc_revenue(courses, price_tbl)
    studio_a  = studio_hours * STUDIO_UNIT_PRICE if include_studio else 0

    # 손익률 40% 보장하는 참여율 계산 (두 문서 공통)
    pm_rate, prod_rate = adjust_rates(revenue, studio_a, ps, pe)

    m2         = f"{month_num:02d}"
    dept_short = dept.replace(" ", "").replace("•", "")

    pnl      = gen_pnl_excel(courses, dept, month_str, year, price_tbl,
                              studio_hours, include_studio, pm_rate, prod_rate)
    req_xlsx = gen_devreq_excel(courses, dept, month_str, year, price_tbl)
    req_docx = gen_devreq_docx(courses, dept, month_str, year, ps, pe, write_dt)
    profile  = gen_profile_docx(courses, dept, month_str, year, price_tbl,
                                 studio_hours, include_studio,
                                 pm_rate, prod_rate, customer_contact)

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr(f"손익분석서_{dept_short}_{year}{m2}.xlsx",  pnl)
        zf.writestr(f"프로젝트프로파일_{dept_short}_{year}{m2}.docx", profile)
        zf.writestr(f"개발요청서_{dept_short}_{year}{m2}.docx",  req_docx)
        zf.writestr(f"개발요청서첨부_{dept_short}_{year}{m2}.xlsx", req_xlsx)
    return buf.getvalue()
