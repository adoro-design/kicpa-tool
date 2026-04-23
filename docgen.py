"""
docgen.py - 문서 자동 생성 모듈 (템플릿 기반)
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

# 상수
MONTHLY_SALARY    = 8_850_000
PM_RATE           = 0.01
PROD_RATE         = 0.15
STUDIO_UNIT_PRICE = 45_000
TARGET_PROFIT     = 0.40

PRICE_NEW       = 500_000
PRICE_PORTING   = 50_000
PRICE_EDIT_PORT = 160_000
PRICE_TRAVEL_HR = 100_000

PM_NAME        = "이상현"
PROD_NAME      = "염왕도"
DEPT_CODE      = "SS"
DEPT_FULL      = "에듀테크서비스본부"
COMPANY_NAME   = "주디유넷"
COMPANY_ADDR   = "서울 서대문구 충정로3가 139"
COMPANY_CEO    = "김평국"
CLIENT_ADDR    = "서울 서대문구 충정로 2가 185-10"
CLIENT_CEO     = "최운열"
DELIVERY_PLACE = "한국공인회계사회"


# 날짜 헬퍼
def get_weekday_count(start, end):
    count, d = 0, start
    while d <= end:
        if d.weekday() < 5:
            count += 1
        d += timedelta(days=1)
    return count

def get_last_business_day(year, month):
    d = date(year, month, calendar.monthrange(year, month)[1])
    while d.weekday() >= 5:
        d -= timedelta(days=1)
    return d

def get_next_month_5th_weekday(year, month):
    ny, nm = (year + 1, 1) if month == 12 else (year, month + 1)
    d = date(ny, nm, 5)
    while d.weekday() >= 5:
        d += timedelta(days=1)
    return d

def get_month_number(month_str):
    m = re.search(r'(\d{1,2})월', str(month_str or ""))
    return int(m.group(1)) if m else 1

def fmt_kr(d):   return f"{d.year}. {d.month:02d}. {d.day:02d}"
def fmt_kr2(d):  return f"{d.year}년 {d.month:02d}월 {d.day:02d}일"
def fmt_short(d): return f"{str(d.year)[2:]}.{d.month:02d}.{d.day:02d}"

def get_sijengil(ps):
    """시행일: 개발기간 시작일 2~3일 전 평일"""
    d = ps - timedelta(days=2)
    while d.weekday() >= 5:
        d -= timedelta(days=1)
    return d


# 단가 헬퍼
def classify_fmt(fmt):
    if not fmt: return True, False, False
    if "포팅" in fmt:
        if "편집" in fmt and "무편집" not in fmt: return False, True, True
        return False, True, False
    return True, False, False

def get_unit_price_for(content_row, price_tbl):
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

def get_travel_for(content_row, travel_rate=PRICE_TRAVEL_HR):
    if not content_row.shooting_format or "출장" not in content_row.shooting_format: return 0
    if content_row.travel_expense is not None: return content_row.travel_expense
    if content_row.travel_hours: return content_row.travel_hours * travel_rate
    return travel_rate


# 공통 계산
def calc_period(courses, year, month_num):
    starts = [c.shooting_date for c in courses if c.shooting_date]
    ends   = [c.open_date     for c in courses if c.open_date]
    s = min(starts) if starts else date(year, month_num, 1)
    e = max(ends)   if ends   else get_next_month_5th_weekday(year, month_num)
    return s, e

def calc_revenue(courses, price_tbl):
    tr = price_tbl.get("1 ~ 4시간", PRICE_TRAVEL_HR)
    return sum(
        get_unit_price_for(c, price_tbl) * (c.session_count or c.chapter_count or 0)
        + get_travel_for(c, tr)
        for c in courses
    )

def calc_labor_amounts(ps, pe, pm_rate, prod_rate):
    """템플릿 공식: (pe-ps).days / 30 * rate * 월인건비 + 1"""
    d = (pe - ps).days
    pm_a   = round(d / 30 * pm_rate   * MONTHLY_SALARY + 1)
    prod_a = round(d / 30 * prod_rate * MONTHLY_SALARY + 1)
    return pm_a, prod_a

def adjust_rates(revenue, studio_a, ps, pe):
    """손익률 40% 이상 + PM 최소 1% 보장. 두 문서에 동일 적용."""
    pm_a, prod_a = calc_labor_amounts(ps, pe, PM_RATE, PROD_RATE)
    total_cost   = pm_a + prod_a + studio_a
    if revenue > 0 and (revenue - total_cost) / revenue >= TARGET_PROFIT:
        return PM_RATE, PROD_RATE

    period_d  = (pe - ps).days
    max_labor = (1 - TARGET_PROFIT) * revenue - studio_a - 2

    # PM 최소 1% 유지
    fixed_pm   = max(PM_RATE, 0.01)
    pm_a_fixed = round(period_d / 30 * fixed_pm * MONTHLY_SALARY + 1)

    # 남은 예산을 PROD에 할당
    max_prod_a  = max_labor - pm_a_fixed
    base_prod_a = period_d / 30 * PROD_RATE * MONTHLY_SALARY
    if max_prod_a <= 0:
        return fixed_pm, 0.0
    if base_prod_a <= 0 or max_prod_a >= base_prod_a:
        return fixed_pm, PROD_RATE
    new_prod = round(max_prod_a / (period_d / 30 * MONTHLY_SALARY), 4) if period_d > 0 else 0
    return fixed_pm, max(new_prod, 0.0)

def build_project_name(courses):
    new_s  = sum(c.session_count or 0 for c in courses if classify_fmt(c.shooting_format or "")[0])
    prt_c  = sum(c.chapter_count or 0 for c in courses
                 if classify_fmt(c.shooting_format or "")[1]
                 and not classify_fmt(c.shooting_format or "")[2])
    eprt_c = sum(c.chapter_count or 0 for c in courses if classify_fmt(c.shooting_format or "")[2])
    parts  = []
    if new_s:  parts.append(f"신규{new_s}차시")
    if prt_c:  parts.append(f"포팅{prt_c}챕터")
    if eprt_c: parts.append(f"편집포팅{eprt_c}챕터")
    return "한국공인회계사 콘텐츠 개발(" + ("·".join(parts) if parts else "신규") + ")"


# Word 텍스트 치환 헬퍼
def _replace_para(para, old, new):
    """단락 전체 텍스트에서 치환. 첫 번째 런에 모아서 처리."""
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

def _replace_doc(doc, old, new):
    """문서 전체(단락 + 표 셀)에서 텍스트 치환"""
    for para in doc.paragraphs:
        _replace_para(para, old, new)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    _replace_para(para, old, new)


# 1. 손익분석서 (Excel) - 템플릿 기반
def gen_pnl_excel(courses, dept, month_str, year, price_tbl,
                  studio_hours, include_studio, pm_rate, prod_rate):
    month_num = get_month_number(month_str)
    ps, pe    = calc_period(courses, year, month_num)
    write_dt  = get_last_business_day(year, month_num)
    revenue   = calc_revenue(courses, price_tbl)
    proj_name = build_project_name(courses)

    wb = openpyxl.load_workbook(os.path.join(TEMPLATE_DIR, 'tpl_pnl.xlsx'))
    ws = wb.active

    ws['F3'] = proj_name
    ws['F4'] = write_dt
    ws['F4'].number_format = 'YYYY"년" MM"월" DD"일"'
    ws['I8'] = revenue
    ws['D10'] = ps;  ws['D10'].number_format = 'YYYY-MM-DD'
    ws['E10'] = pe;  ws['E10'].number_format = 'YYYY-MM-DD'
    ws['F10'] = pm_rate
    ws['D11'] = ps;  ws['D11'].number_format = 'YYYY-MM-DD'
    ws['E11'] = pe;  ws['E11'].number_format = 'YYYY-MM-DD'
    ws['F11'] = prod_rate
    ws['D17'] = STUDIO_UNIT_PRICE if include_studio else 0
    ws['E17'] = studio_hours       if include_studio else 0
    ws['E32'] = MONTHLY_SALARY

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# 2. 개발요청서 Excel 첨부 - 템플릿 기반
def gen_devreq_excel(courses, dept, month_str, year, price_tbl):
    month_num = get_month_number(month_str)
    write_dt  = get_last_business_day(year, month_num)

    wb = openpyxl.load_workbook(os.path.join(TEMPLATE_DIR, 'tpl_devreq.xlsx'))
    ws = wb['콘텐츠개발내역']

    ws['A1'] = f"{month_num}월 콘텐츠 개발 내역"
    ws['N1'] = f"작성일 : {write_dt.strftime('%Y-%m-%d')}"
    ws['A2'] = f"* {dept}"

    # 기존 데이터행 위치 파악
    tpl_data_start = 4
    tpl_total_row  = tpl_data_start
    while tpl_total_row <= ws.max_row:
        if ws.cell(tpl_total_row, 1).value == '전체 합계':
            break
        tpl_total_row += 1
    tpl_data_count = tpl_total_row - tpl_data_start

    # 행 조작 전에 데이터 영역 병합 해제
    for mr in list(ws.merged_cells.ranges):
        if mr.min_row >= tpl_data_start:
            ws.unmerge_cells(str(mr))

    n = len(courses)
    if n > tpl_data_count:
        ws.insert_rows(tpl_total_row, n - tpl_data_count)
    elif n < tpl_data_count:
        ws.delete_rows(tpl_data_start + n, tpl_data_count - n)

    # 데이터 작성
    tr = price_tbl.get("1 ~ 4시간", PRICE_TRAVEL_HR)
    for idx, c in enumerate(courses):
        r = tpl_data_start + idx
        is_new, is_p, is_ep = classify_fmt(c.shooting_format or "")
        unit_p = get_unit_price_for(c, price_tbl)

        vals = {
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
        for col, val in vals.items():
            cell = ws.cell(r, col, val)
            if col == 8 and isinstance(val, date):
                cell.number_format = 'YYYY-MM-DD'
            if col == 12:
                cell.number_format = '#,##0'

    # 합계 행
    total_row = tpl_data_start + n
    end_row   = total_row - 1

    ws.cell(total_row, 1, "전체 합계")
    ws.merge_cells(f"A{total_row}:E{total_row}")
    ws.cell(total_row, 6,  f"=SUM(F{tpl_data_start}:F{end_row})")
    ws.cell(total_row, 7,  f"=SUM(G{tpl_data_start}:G{end_row})")
    ws.cell(total_row, 13, f"=SUM(M{tpl_data_start}:M{end_row})").number_format = '#,##0'
    ws.cell(total_row, 14, f"=SUM(N{tpl_data_start}:N{end_row})").number_format = '#,##0'

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# 3. 개발요청서 Word - 템플릿 기반 (텍스트 치환)
def gen_devreq_docx(courses, dept, month_str, year, ps, pe, write_dt):
    doc   = DocxDocument(os.path.join(TEMPLATE_DIR, 'tpl_devreq.docx'))
    table = doc.tables[0]

    # 시행일 = 개발기간 시작일 2~3일 전 평일
    sijengil = get_sijengil(ps)

    # 시행일 행 (row 5, cells[1..3])
    for cell in table.rows[5].cells[1:]:
        for para in cell.paragraphs:
            _replace_para(para, "2026년 03월 03일", fmt_kr2(sijengil))

    # 본문 (row 7) - 중복 셀 처리 방지
    replacements = [
        ("조세지원본부",      dept),
        ("2026년 03월 04일",  fmt_kr2(ps)),
        ("2026년 03월 27일",  fmt_kr2(pe)),
        ("2026년 03월 16일",  fmt_kr2(write_dt)),
    ]
    seen = set()
    for cell in table.rows[7].cells:
        cid = id(cell._tc)
        if cid in seen:
            continue
        seen.add(cid)
        for old, new in replacements:
            for para in cell.paragraphs:
                _replace_para(para, old, new)

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


# 4. 프로젝트 프로파일 Word - 템플릿 기반 (텍스트 치환)
def _set_cell_text(cell, new_text):
    """셀 텍스트 교체 (기존 서식 최대한 유지)"""
    for para in cell.paragraphs:
        if para.runs:
            para.runs[0].text = new_text
            for run in para.runs[1:]:
                run.text = ""
            return
    cell.paragraphs[0].add_run(new_text)

def gen_profile_docx(courses, dept, month_str, year, price_tbl,
                     studio_hours, include_studio, pm_rate, prod_rate,
                     customer_contact=None):
    month_num = get_month_number(month_str)
    ps, pe    = calc_period(courses, year, month_num)
    write_dt  = get_last_business_day(year, month_num)
    revenue   = calc_revenue(courses, price_tbl)
    studio_a  = studio_hours * STUDIO_UNIT_PRICE if include_studio else 0
    proj_name = build_project_name(courses)

    pm_pct   = round(pm_rate   * 100, 1)
    prod_pct = round(prod_rate * 100, 1)
    period_str = f"{fmt_short(ps)} ~ {fmt_short(pe)}"

    # 세부내역
    new_s  = sum(c.session_count or 0 for c in courses if classify_fmt(c.shooting_format or "")[0])
    prt_c  = sum(c.chapter_count or 0 for c in courses
                 if classify_fmt(c.shooting_format or "")[1]
                 and not classify_fmt(c.shooting_format or "")[2])
    eprt_c = sum(c.chapter_count or 0 for c in courses if classify_fmt(c.shooting_format or "")[2])
    travel_h = sum((c.travel_hours or 1)
                   for c in courses if c.shooting_format and "출장" in c.shooting_format)

    detail_parts = []
    if new_s:    detail_parts.append(f"- 신규 : {new_s}차시(단가 : \\500,000, VAT별도) / 유상개발")
    if prt_c:    detail_parts.append(f"- 포팅(무편집) : {prt_c}챕터(단가 : \\50,000, VAT별도) / 유상개발")
    if eprt_c:   detail_parts.append(f"- 포팅(편집) : {eprt_c}챕터(단가 : \\160,000, VAT별도) / 유상개발")
    if travel_h: detail_parts.append(f"- 출장비 : {travel_h}시간(단가 : \\100,000, VAT별도)")
    new_detail = (
        "1. 사업내용\n(1) 한공회 콘텐츠 개발\n"
        + "\n".join(detail_parts)
        + f"\n(2) 개발기간 : {fmt_kr(ps)} ~ {fmt_kr(pe)}"
    )

    contact_name  = customer_contact.contact_name if customer_contact else ""
    contact_phone = customer_contact.phone        if customer_contact else ""
    contact_email = customer_contact.email        if customer_contact else ""

    doc = DocxDocument(os.path.join(TEMPLATE_DIR, 'tpl_profile.docx'))
    t0, t1, t2, t3, t4 = doc.tables[0], doc.tables[1], doc.tables[2], doc.tables[3], doc.tables[4]

    # 표 0: 작성일
    _set_cell_text(t0.rows[1].cells[3], fmt_kr(write_dt))

    # 표 1: 개요
    _replace_doc(doc, "한공회 콘텐츠 개발(신규14차시)", proj_name)
    _replace_doc(doc, "2026. 03. 04 ~ 2026. 03. 27", f"{fmt_kr(ps)} ~ {fmt_kr(pe)}")
    _replace_doc(doc, "7,000,000", f"{revenue:,}")
    _set_cell_text(t1.rows[8].cells[1], new_detail)
    _set_cell_text(t1.rows[12].cells[1], contact_name)
    _set_cell_text(t1.rows[12].cells[2], contact_phone)
    _set_cell_text(t1.rows[12].cells[5], contact_email)

    # 표 2: 참여인력 - 참여율
    _replace_doc(doc, "5%",  f"{pm_pct:.4g}%")
    _replace_doc(doc, "25%", f"{prod_pct:.4g}%")
    _replace_doc(doc, "26. 03. 04 ~ 26. 03. 27", period_str)

    # 표 4: 장비 및 시설 - 스튜디오
    if include_studio and studio_hours > 0:
        _set_cell_text(t4.rows[2].cells[1], str(studio_hours))
        _set_cell_text(t4.rows[2].cells[2], f"{STUDIO_UNIT_PRICE:,}")
        _set_cell_text(t4.rows[2].cells[3], f"{studio_hours * STUDIO_UNIT_PRICE:,}")
    else:
        for ci in range(5):
            _set_cell_text(t4.rows[2].cells[ci], "")

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


# 전체 ZIP 생성
def generate_all(courses, dept, month_str, year, price_tbl,
                 studio_hours, include_studio, customer_contact=None):
    month_num = get_month_number(month_str)
    ps, pe    = calc_period(courses, year, month_num)
    write_dt  = get_last_business_day(year, month_num)
    revenue   = calc_revenue(courses, price_tbl)
    studio_a  = studio_hours * STUDIO_UNIT_PRICE if include_studio else 0

    # 참여율 계산 (PM 최소 1%, 손익 40% 보장)
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
        zf.writestr(f"손익분석서_{dept_short}_{year}{m2}.xlsx",     pnl)
        zf.writestr(f"프로젝트프로파일_{dept_short}_{year}{m2}.docx", profile)
        zf.writestr(f"개발요청서_{dept_short}_{year}{m2}.docx",      req_docx)
        zf.writestr(f"개발요청서첨부_{dept_short}_{year}{m2}.xlsx",   req_xlsx)
    return buf.getvalue()
