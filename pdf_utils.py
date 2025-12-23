from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib.utils import ImageReader
from reportlab.lib import colors
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image, Table, TableStyle, PageBreak
import io
import datetime
from typing import Optional
from scheduler import build_supervisor_table

def _is_winter(month: int) -> bool:
    # winter: Oct-Jan (10-1), summer: Mar-Jul (3-7)
    return month in [10,11,12,1]

def generate_duty_pdf(supervisor_name: str, schedule_df, staff_df, start_date, end_date, exam_type: str, college_logo_bytes: Optional[bytes]=None, uni_logo_bytes: Optional[bytes]=None, sign_bytes: Optional[bytes]=None) -> bytes:
    buf = io.BytesIO()
    # Use platypus SimpleDocTemplate so text and tables flow across A4 pages correctly
    doc = SimpleDocTemplate(buf, pagesize=A4, leftMargin=30, rightMargin=30, topMargin=40, bottomMargin=40)
    width, height = A4

    styles = getSampleStyleSheet()
    normal = ParagraphStyle('Normal', parent=styles['Normal'], fontName='Helvetica', fontSize=10, leading=15)
    heading = ParagraphStyle('Heading', parent=styles['Heading1'], alignment=1, fontName='Helvetica-Bold', fontSize=12)
    title = ParagraphStyle('Title', parent=styles['Title'], alignment=1, fontName='Helvetica-Bold', fontSize=14)

    story = []

    # Header with logos and centered lines
    left_img = None
    right_img = None
    try:
        if college_logo_bytes:
            left_img = Image(io.BytesIO(college_logo_bytes), width=40*mm, height=40*mm)
    except Exception:
        left_img = None
    try:
        if uni_logo_bytes:
            right_img = Image(io.BytesIO(uni_logo_bytes), width=40*mm, height=40*mm)
    except Exception:
        right_img = None

    now = datetime.date.today()
    season = 'Winter' if _is_winter(now.month) else 'Summer'
    # Header with smaller institute/university font sizes and a slightly larger title
    center_html = (
        '<font size="10"><b>V. V. P. INSTITUTE OF ENGINEERING AND TECHNOLOGY,</b></font><br/>'
        '<font size="10"><b>SOLAPUR</b></font><br/>'
        '<font size="9"><b>Dr. BABASAHEB AMBEDKAR TECHNOLOGICAL UNIVERSITY,</b></font><br/>'
        '<font size="9"><b>LONERE</b></font><br/>'
        f'<font size="9">{season} Exam Regular And Supplementary ({now.year})</font><br/>'
        '<font size="12"><b>DUTY ALLOTMENT SHEET</b></font>'
    )
    center_para = Paragraph(center_html, ParagraphStyle('center', parent=getSampleStyleSheet()['Normal'], alignment=1, leading=18))

    # Build a table for header to place left logo, centered text, right logo
    header_data = [
        [left_img if left_img else '', center_para, right_img if right_img else '']
    ]
    header_tbl = Table(header_data, colWidths=[40*mm, (width-80*mm), 40*mm])
    header_tbl.setStyle(TableStyle([
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('ALIGN', (1,0), (1,0), 'CENTER')
    ]))
    story.append(header_tbl)
    story.append(Spacer(1, 6))
    # Date on right
    story.append(Paragraph(now.strftime('%Y-%m-%d'), ParagraphStyle('Right', parent=styles['Normal'], alignment=2, fontSize=9)))
    story.append(Spacer(1, 10))

    # Salutation
    salutation_text = f"To,<br/>The Invigilators/supervisor,<br/><br/><b>{supervisor_name}</b><br/><br/>"
    salutation_text += ("Following is the schedule of your Jr. Supervisions for the DBATU, LONERE examination at VVPIET, Solapur center. "
                      "You are requested to go through the instructions written overleaf.<br/><br/>"
                      "In accordance with the request of University authority, we bring to your notice the provisions of the amendment of the section 32 (g) of Maharashtra Universities act 1994. "
                      "32 (g): &ldquo;It shall be obligatory on every teacher and on the non-teaching employee of the University, affiliated, conducted or autonomous college or recognized institutions to render necessary assistance and service in respect of examinations of the University. If any teacher or non-teaching employee fails to comply with the order of the University or college or institution in this respect, It shall be treated as misconduct and the employee shall be liable for disciplinary action.&rdquo; "
                      "Kindly acknowledge the receipt of the duty allotment.")
    story.append(Paragraph(salutation_text, normal))
    story.append(Spacer(1, 12))

    # Supervisor table
    table_df = build_supervisor_table(supervisor_name, schedule_df)
    if table_df.empty:
        story.append(Paragraph('No duties assigned.', normal))
    else:
        # Use Paragraph in cells for wrapping in morning/evening columns
        data = [[Paragraph('<b>Sr. No.</b>', normal), Paragraph('<b>Date</b>', normal), Paragraph('<b>Morning (10.00 a.m. to 01.00 p.m.)</b>', normal), Paragraph('<b>Evening (02.00 p.m. to 05.00 p.m.)</b>', normal)]]
        for i, row in table_df.iterrows():
            m_cell = Paragraph(row["Morning"] or '', normal)
            e_cell = Paragraph(row["Evening"] or '', normal)
            data.append([row["Sr. No."], row["Date"], m_cell, e_cell])
        tbl = Table(data, colWidths=[20*mm, 35*mm, 65*mm, 65*mm], repeatRows=1)
        style = TableStyle([
            ('GRID', (0,0), (-1,-1), 0.5, colors.black),
            ('BACKGROUND', (0,0), (-1,0), colors.lightgrey),
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ])
        tbl.setStyle(style)
        story.append(tbl)
        story.append(Spacer(1, 12))

    # Instructions
    instr_lines = [
        '<b>INSTRUCTIONS TO INVIGILATORS/SUPERVISOR</b>',
        'All the Invigilators/supervisor are informed to observe following points strictly.',
        '01. Report exam office 30 min. prior to starting time of examination.',
        '02. No substitute arrangements be done without principal/Office In charge permission.',
        '03. Check the identity card and Exam fee receipt/hall ticket during every examination.',
        '04. All the books, note books and any other material brought by the students should be kept outside the hall.',
        '05. Students are not allowed to communicate with other students, exchange the calculators or any other material during examination period.',
        '06. Programmable calculators are not allowed.',
        '07. Students are not allowed to use colored pencil/pen and make any objectionable marks on the answer sheet.',
        '08. Student should not write anything on question paper.',
        '09. Invigilators/supervisor shall make two copies of their report for each paper of two sections. For composite blocks Jr. Supervisors should give separate report for every paper. Also the O/C has to be separate for every paper.',
        '10. Invigilators/supervisor should not give any kind of explanation or interpretation to students in connection with the question paper.',
        '11. Students should not be allowed to leave exam hall within half an hour after commencement of examination.'
    ]
    for li in instr_lines:
        story.append(Paragraph(li, normal))
        story.append(Spacer(1, 6))

    # Signature (image if provided)
    story.append(Spacer(1, 24))
    if sign_bytes:
        try:
            sig = Image(io.BytesIO(sign_bytes), width=40*mm, height=20*mm)
            # Right-align the signature image above the Office In charge lines
            sig_tbl = Table([["", sig]], colWidths=[width - (40*mm) - 20, 40*mm])
            sig_tbl.setStyle(TableStyle([('ALIGN', (1,0), (1,0), 'RIGHT'), ('VALIGN', (0,0), (-1,-1), 'MIDDLE')]))
            story.append(sig_tbl)
        except Exception:
            pass
    story.append(Paragraph('Office In charge', ParagraphStyle('sig', parent=styles['Normal'], alignment=2)))
    story.append(Paragraph('VVPIET CENTER, SOLAPUR', ParagraphStyle('sig2', parent=styles['Normal'], alignment=2)))

    # Build document
    doc.build(story)
    buf.seek(0)
    return buf.read()


def _build_story_for_supervisor(supervisor_name: str, schedule_df, staff_df, start_date, end_date, exam_type: str, college_logo_bytes: Optional[bytes]=None, uni_logo_bytes: Optional[bytes]=None, sign_bytes: Optional[bytes]=None):
    """Return a list of flowables for a single supervisor's duty order (without building the PDF)."""
    styles = getSampleStyleSheet()
    normal = ParagraphStyle('Normal', parent=styles['Normal'], fontName='Helvetica', fontSize=10, leading=15)

    story = []

    # Header
    left_img = None
    right_img = None
    try:
        if college_logo_bytes:
            left_img = Image(io.BytesIO(college_logo_bytes), width=40*mm, height=40*mm)
    except Exception:
        left_img = None
    try:
        if uni_logo_bytes:
            right_img = Image(io.BytesIO(uni_logo_bytes), width=40*mm, height=40*mm)
    except Exception:
        right_img = None

    now = datetime.date.today()
    season = 'Winter' if _is_winter(now.month) else 'Summer'
    center_html = (
        '<font size="10"><b>V. V. P. INSTITUTE OF ENGINEERING AND TECHNOLOGY,</b></font><br/>'
        '<font size="10"><b>SOLAPUR</b></font><br/>'
        '<font size="9"><b>Dr. BABASAHEB AMBEDKAR TECHNOLOGICAL UNIVERSITY,</b></font><br/>'
        '<font size="9"><b>LONERE</b></font><br/>'
        f'<font size="9">{season} Exam Regular And Supplementary ({now.year})</font><br/>'
        '<font size="12"><b>DUTY ALLOTMENT SHEET</b></font>'
    )
    center_para = Paragraph(center_html, ParagraphStyle('center', parent=getSampleStyleSheet()['Normal'], alignment=1, leading=18))

    header_data = [[left_img if left_img else '', center_para, right_img if right_img else '']]
    width, _ = A4
    header_tbl = Table(header_data, colWidths=[40*mm, (width-80*mm), 40*mm])
    header_tbl.setStyle(TableStyle([('VALIGN', (0,0), (-1,-1), 'MIDDLE'), ('ALIGN', (1,0), (1,0), 'CENTER')]))
    story.append(header_tbl)
    story.append(Spacer(1, 6))
    story.append(Paragraph(now.strftime('%Y-%m-%d'), ParagraphStyle('Right', parent=styles['Normal'], alignment=2, fontSize=9)))
    story.append(Spacer(1, 10))

    # Salutation
    salutation_text = f"To,<br/>The Invigilators/supervisor,<br/><br/><b>{supervisor_name}</b><br/><br/>"
    salutation_text += ("Following is the schedule of your Jr. Supervisions for the DBATU, LONERE examination at VVPIET, Solapur center. "
                      "You are requested to go through the instructions written overleaf.<br/><br/>"
                      "In accordance with the request of University authority, we bring to your notice the provisions of the amendment of the section 32 (g) of Maharashtra Universities act 1994. "
                      "32 (g): &ldquo;It shall be obligatory on every teacher and on the non-teaching employee of the University, affiliated, conducted or autonomous college or recognized institutions to render necessary assistance and service in respect of examinations of the University. If any teacher or non-teaching employee fails to comply with the order of the University or college or institution in this respect, It shall be treated as misconduct and the employee shall be liable for disciplinary action.&rdquo; "
                      "Kindly acknowledge the receipt of the duty allotment.")
    story.append(Paragraph(salutation_text, normal))
    story.append(Spacer(1, 12))

    # Supervisor table
    table_df = build_supervisor_table(supervisor_name, schedule_df)
    if table_df.empty:
        story.append(Paragraph('No duties assigned.', normal))
    else:
        data = [[Paragraph('<b>Sr. No.</b>', normal), Paragraph('<b>Date</b>', normal), Paragraph('<b>Morning (10.00 a.m. to 01.00 p.m.)</b>', normal), Paragraph('<b>Evening (02.00 p.m. to 05.00 p.m.)</b>', normal)]]
        for i, row in table_df.iterrows():
            m_cell = Paragraph(row["Morning"] or '', normal)
            e_cell = Paragraph(row["Evening"] or '', normal)
            data.append([row["Sr. No."], row["Date"], m_cell, e_cell])
        tbl = Table(data, colWidths=[20*mm, 35*mm, 65*mm, 65*mm], repeatRows=1)
        style = TableStyle([('GRID', (0,0), (-1,-1), 0.5, colors.black), ('BACKGROUND', (0,0), (-1,0), colors.lightgrey), ('ALIGN', (0,0), (-1,-1), 'CENTER'), ('VALIGN', (0,0), (-1,-1), 'MIDDLE')])
        tbl.setStyle(style)
        story.append(tbl)
        story.append(Spacer(1, 12))

    # Instructions
    instr_lines = ['<b>INSTRUCTIONS TO INVIGILATORS/SUPERVISOR</b>', 'All the Invigilators/supervisor are informed to observe following points strictly.', '01. Report exam office 30 min. prior to starting time of examination.', '02. No substitute arrangements be done without principal/Office In charge permission.', '03. Check the identity card and Exam fee receipt/hall ticket during every examination.', '04. All the books, note books and any other material brought by the students should be kept outside the hall.', '05. Students are not allowed to communicate with other students, exchange the calculators or any other material during examination period.', '06. Programmable calculators are not allowed.', '07. Students are not allowed to use colored pencil/pen and make any objectionable marks on the answer sheet.', '08. Student should not write anything on question paper.', '09. Invigilators/supervisor shall make two copies of their report for each paper of two sections. For composite blocks Jr. Supervisors should give separate report for every paper. Also the O/C has to be separate for every paper.', '10. Invigilators/supervisor should not give any kind of explanation or interpretation to students in connection with the question paper.', '11. Students should not be allowed to leave exam hall within half an hour after commencement of examination.']
    for li in instr_lines:
        story.append(Paragraph(li, normal))
        story.append(Spacer(1, 6))

    # Signature (image if provided)
    story.append(Spacer(1, 24))
    if sign_bytes:
        try:
            sig = Image(io.BytesIO(sign_bytes), width=40*mm, height=20*mm)
            # Right-align the signature image above the Office In charge lines
            sig_tbl = Table([["", sig]], colWidths=[width - (40*mm) - 20, 40*mm])
            sig_tbl.setStyle(TableStyle([('ALIGN', (1,0), (1,0), 'RIGHT'), ('VALIGN', (0,0), (-1,-1), 'MIDDLE')]))
            story.append(sig_tbl)
        except Exception:
            pass
    story.append(Paragraph('Office In charge', ParagraphStyle('sig', parent=getSampleStyleSheet()['Normal'], alignment=2)))
    story.append(Paragraph('VVPIET CENTER, SOLAPUR', ParagraphStyle('sig2', parent=getSampleStyleSheet()['Normal'], alignment=2)))

    return story


def generate_absence_memo(supervisor_name: str, absences: list, staff_df, college_logo_bytes: Optional[bytes]=None, uni_logo_bytes: Optional[bytes]=None, sign_bytes: Optional[bytes]=None) -> bytes:
    """Generate a short memo PDF for a supervisor listing absences. """
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, leftMargin=30, rightMargin=30, topMargin=40, bottomMargin=40)
    width, _ = A4
    styles = getSampleStyleSheet()
    normal = ParagraphStyle('Normal', parent=styles['Normal'], fontName='Helvetica', fontSize=10, leading=15)
    heading = ParagraphStyle('Heading', parent=styles['Heading1'], alignment=1, fontName='Helvetica-Bold', fontSize=12)

    story = []
    # Header
    # Top institution line (bold, centered)
    inst_html = '<font size="12"><b>VVP Institute of Engineering and Technology, Solapur</b></font>'
    story.append(Paragraph(inst_html, ParagraphStyle('inst', parent=styles['Normal'], alignment=1)))
    story.append(Spacer(1, 6))

    center_html = '<font size="12"><b>DUTY ABSENCE MEMO</b></font>'
    story.append(Paragraph(center_html, ParagraphStyle('center', parent=styles['Normal'], alignment=1)))
    # Add current date below the title (right aligned)
    today = datetime.date.today()
    story.append(Paragraph(today.strftime('%Y-%m-%d'), ParagraphStyle('date_right', parent=styles['Normal'], alignment=2, fontSize=9)))
    story.append(Spacer(1, 12))

    # Body
    date_lines = '<br/>'.join([f"- {d.strftime('%Y-%m-%d')} ({s})" for d, s in absences])
    body = f"To,<br/><b>{supervisor_name}</b><br/><br/>This is to inform you that you were absent for invigilation duty on the following date(s)/session(s):<br/>{date_lines}<br/><br/>You are requested to explain the absence within one day and acknowledge receipt of this memo.<br/><br/>"
    story.append(Paragraph(body, normal))
    story.append(Spacer(1, 24))

    # Signature image if provided
    if sign_bytes:
        try:
            sig = Image(io.BytesIO(sign_bytes), width=40*mm, height=20*mm)
            # Right-align the signature image above the Office In charge lines
            sig_tbl = Table([["", sig]], colWidths=[width - (40*mm) - 20, 40*mm])
            sig_tbl.setStyle(TableStyle([('ALIGN', (1,0), (1,0), 'RIGHT'), ('VALIGN', (0,0), (-1,-1), 'MIDDLE')]))
            story.append(sig_tbl)
        except Exception:
            pass

    story.append(Spacer(1, 12))
    story.append(Paragraph('Office In charge', ParagraphStyle('sig', parent=styles['Normal'], alignment=2)))
    story.append(Paragraph('VVPIET CENTER, SOLAPUR', ParagraphStyle('sig2', parent=styles['Normal'], alignment=2)))

    doc.build(story)
    buf.seek(0)
    return buf.read()


def generate_combined_duty_pdf(supervisor_names: list, schedule_df, staff_df, start_date, end_date, exam_type: str, college_logo_bytes: Optional[bytes]=None, uni_logo_bytes: Optional[bytes]=None, sign_bytes: Optional[bytes]=None) -> bytes:
    """Generate a single combined PDF containing duty orders for all supervisors in supervisor_names (each starts on a new page)."""
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, leftMargin=30, rightMargin=30, topMargin=40, bottomMargin=40)
    story = []
    for i, name in enumerate(supervisor_names):
        story.extend(_build_story_for_supervisor(name, schedule_df, staff_df, start_date, end_date, exam_type, college_logo_bytes, uni_logo_bytes, sign_bytes))
        if i < len(supervisor_names) - 1:
            story.append(PageBreak())
    doc.build(story)
    buf.seek(0)
    return buf.read()

def combine_pdfs_bytes(list_of_pdf_bytes: list) -> bytes:
    # Try pypdf first, then PyPDF2, otherwise raise an informative error
    merger = None
    try:
        from pypdf import PdfMerger
        merger = PdfMerger()
        for b in list_of_pdf_bytes:
            f = io.BytesIO(b)
            f.seek(0)
            merger.append(f)
        out = io.BytesIO()
        merger.write(out)
        merger.close()
        return out.getvalue()
    except Exception:
        pass

    try:
        # PyPDF2 compatibility
        from PyPDF2 import PdfMerger as PyPdfMerger
        merger = PyPdfMerger()
        for b in list_of_pdf_bytes:
            f = io.BytesIO(b)
            f.seek(0)
            merger.append(f)
        out = io.BytesIO()
        merger.write(out)
        merger.close()
        return out.getvalue()
    except Exception:
        pass

    # Last resort: return concatenation of bytes (may not be valid PDF)
    raise RuntimeError("Unable to combine PDFs: install 'pypdf' or 'PyPDF2' to enable PDF merging (e.g., pip install pypdf).")
