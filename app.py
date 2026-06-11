import os
import uuid
import io
import json
import time
import threading
from pathlib import Path
from flask import Flask, render_template, request, redirect, url_for, flash, send_file
from scheduler import parse_excel, organize_by_day, build_grid
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.units import mm
from PIL import Image, ImageDraw, ImageFont

app = Flask(__name__)
app.secret_key = 'schedule-manager-secret-key-change-in-production'
app.config['UPLOAD_FOLDER'] = os.path.join(os.path.dirname(__file__), 'uploads')
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024
app.config['UPLOAD_CLEANUP_AGE'] = 30 * 60  # 30 minutes

ALLOWED_EXTENSIONS = {'xlsx', 'xls'}

os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)


def cleanup_uploads(age=None):
    folder = Path(app.config['UPLOAD_FOLDER'])
    now = time.time()
    for f in folder.iterdir():
        if f.is_file() and f.suffix.lower() in ('.xlsx', '.xls'):
            if age is None or now - f.stat().st_mtime > age:
                f.unlink(missing_ok=True)


def cleanup_older_uploads():
    cleanup_uploads(age=app.config['UPLOAD_CLEANUP_AGE'])


def start_cleanup_thread():
    def run():
        while True:
            time.sleep(600)
            cleanup_older_uploads()
    thread = threading.Thread(target=run, daemon=True)
    thread.start()


cleanup_uploads()  # cleanup stale files on startup
start_cleanup_thread()

DAY_COLORS = {
    'Sunday': '#E3F2FD',
    'Monday': '#F3E5F5',
    'Tuesday': '#E8F5E9',
    'Wednesday': '#FFF3E0',
    'Thursday': '#FCE4EC',
    'Friday': '#F5F5F5',
    'Saturday': '#E0F7FA',
}

DAY_SHORT = {
    'Sunday': 'Sun',
    'Monday': 'Mon',
    'Tuesday': 'Tue',
    'Wednesday': 'Wed',
    'Thursday': 'Thu',
    'Friday': 'Fri',
    'Saturday': 'Sat',
}


def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


def process_file(filepath):
    student_info, courses = parse_excel(filepath)
    schedule = organize_by_day(courses)
    return student_info, schedule


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        flash('No file selected')
        return redirect(url_for('index'))

    file = request.files['file']
    if file.filename == '':
        flash('No file selected')
        return redirect(url_for('index'))

    if not allowed_file(file.filename):
        flash('Please upload an Excel file (.xlsx or .xls)')
        return redirect(url_for('index'))

    cleanup_uploads()  # remove all previous uploads

    file_id = str(uuid.uuid4())
    ext = file.filename.rsplit('.', 1)[1].lower()
    filename = f'{file_id}.{ext}'
    filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    file.save(filepath)

    try:
        student_info, schedule = process_file(filepath)
        if not schedule:
            flash('Could not find any course schedule in the file')
            return redirect(url_for('index'))
        active_days, grid_rows = build_grid(schedule)
        unique_courses = []
        seen = set()
        for row in grid_rows:
            for day in active_days:
                c = row['courses'].get(day)
                if c and c['course'] not in seen:
                    seen.add(c['course'])
                    unique_courses.append(c)
        return render_template('schedule.html',
                               student=student_info,
                               schedule=schedule,
                               active_days=active_days,
                               grid_rows=grid_rows,
                               day_short=DAY_SHORT,
                               day_colors=DAY_COLORS,
                               file_id=file_id,
                               unique_courses=unique_courses)
    except Exception as e:
        flash(f'Error reading file: {str(e)}')
        return redirect(url_for('index'))


@app.route('/download/pdf/<file_id>')
def download_pdf(file_id):
    filepath = os.path.join(app.config['UPLOAD_FOLDER'], f'{file_id}.xlsx')
    if not os.path.exists(filepath):
        alt_path = os.path.join(app.config['UPLOAD_FOLDER'], f'{file_id}.xls')
        if os.path.exists(alt_path):
            filepath = alt_path
        else:
            flash('File not found')
            return redirect(url_for('index'))

    student_info, schedule = process_file(filepath)
    active_days, grid_rows = build_grid(schedule)
    faculty_raw = request.args.get('faculty', '')
    try:
        faculty_map = json.loads(faculty_raw) if faculty_raw else {}
    except (json.JSONDecodeError, TypeError):
        faculty_map = {}
    buf = generate_pdf(student_info, active_days, grid_rows, faculty_map)
    return send_file(buf, mimetype='application/pdf',
                     as_attachment=True,
                     download_name=f'schedule_{student_info["name"] or "routine"}.pdf')


@app.route('/download/image/<file_id>')
def download_image(file_id):
    filepath = os.path.join(app.config['UPLOAD_FOLDER'], f'{file_id}.xlsx')
    if not os.path.exists(filepath):
        alt_path = os.path.join(app.config['UPLOAD_FOLDER'], f'{file_id}.xls')
        if os.path.exists(alt_path):
            filepath = alt_path
        else:
            flash('File not found')
            return redirect(url_for('index'))

    student_info, schedule = process_file(filepath)
    active_days, grid_rows = build_grid(schedule)
    faculty_raw = request.args.get('faculty', '')
    try:
        faculty_map = json.loads(faculty_raw) if faculty_raw else {}
    except (json.JSONDecodeError, TypeError):
        faculty_map = {}
    buf = generate_image(student_info, active_days, grid_rows, faculty_map)
    return send_file(buf, mimetype='image/png',
                     as_attachment=True,
                     download_name=f'schedule_{student_info["name"] or "routine"}.png')


def get_faculty(c, faculty_map):
    return faculty_map.get(c['course'], '')

def make_cell_text(c, faculty_map):
    sec = c.get('sec', '')
    parts = [f"<b>{c['course']}</b>"]
    sec_line = ''
    if sec:
        sec_line = f"Sec-{sec}"
        f = get_faculty(c, faculty_map)
        if f:
            sec_line += f"({f})"
    if sec_line:
        parts.append(sec_line)
    parts.append(c['room'])
    return '<br/>'.join(parts)

def generate_pdf(student_info, active_days, grid_rows, faculty_map={}):
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=landscape(A4),
                            leftMargin=15*mm, rightMargin=15*mm,
                            topMargin=15*mm, bottomMargin=15*mm)

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle('Title2', parent=styles['Title'], fontSize=18, spaceAfter=6)
    info_style = ParagraphStyle('Info', parent=styles['Normal'], fontSize=11, spaceAfter=6)
    cell_style = ParagraphStyle('Cell', parent=styles['Normal'], fontSize=10, leading=14, spaceAfter=0, spaceBefore=0, alignment=TA_CENTER)

    elements = []
    title = f"Class Routine - {student_info['semester']}"
    elements.append(Paragraph(title, title_style))
    info = f"{student_info['name']} | {student_info['id']} | {student_info['semester']}"
    elements.append(Paragraph(info, info_style))
    elements.append(Spacer(1, 6*mm))

    if not active_days:
        elements.append(Paragraph("No schedule data found.", styles['Normal']))
        doc.build(elements)
        buf.seek(0)
        return buf

    header = ['Time'] + [DAY_SHORT.get(d, d[:3]) for d in active_days]
    table_data = [header]

    page_w = landscape(A4)[0]
    available = page_w - 30*mm

    from reportlab.pdfbase.pdfmetrics import stringWidth
    time_max_w = max(stringWidth(f"{r['start']} - {r['end']}", 'Helvetica', 10) for r in grid_rows)
    time_col_w = max(time_max_w + 14, 65)

    day_max_w = max(stringWidth(DAY_SHORT.get(d, d[:3]), 'Helvetica-Bold', 11) for d in active_days)
    for row in grid_rows:
        for day in active_days:
            c = row['courses'].get(day)
            if c:
                cw = stringWidth(f"{c['course']}", 'Helvetica-Bold', 10)
                sec = c.get('sec', '')
                sec_line = ''
                if sec:
                    sec_line = f"Sec-{sec}"
                    f = get_faculty(c, faculty_map)
                    if f:
                        sec_line += f"({f})"
                sw = stringWidth(sec_line, 'Helvetica', 10) if sec_line else 0
                rw = stringWidth(c['room'], 'Helvetica', 10)
                day_max_w = max(day_max_w, cw, sw, rw)
    day_col_w = min(max(day_max_w + 14, 110), int((available - time_col_w) / len(active_days)))
    if len(active_days) * day_col_w + time_col_w > available:
        day_col_w = int((available - time_col_w) / len(active_days))
    if day_col_w < 90:
        day_col_w = 90
        time_col_w = int(available - day_col_w * len(active_days))
        if time_col_w < 50:
            time_col_w = 50
            day_col_w = int((available - 50) / len(active_days))

    col_widths = [time_col_w] + [day_col_w] * len(active_days)

    for row in grid_rows:
        row_vals = [Paragraph(f"{row['start']} - {row['end']}", cell_style)]
        for day in active_days:
            c = row['courses'].get(day)
            if c:
                row_vals.append(Paragraph(make_cell_text(c, faculty_map), cell_style))
            else:
                row_vals.append('')
        table_data.append(row_vals)

    table = Table(table_data, colWidths=col_widths, repeatRows=1)

    style_cmds = [
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1a73e8')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 11),
        ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 1), (-1, -1), 10),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#f8f9fa')]),
        ('TOPPADDING', (0, 0), (-1, -1), 6),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
    ]

    for i, day in enumerate(active_days, 1):
        day_color = DAY_COLORS.get(day, '#FFFFFF')
        style_cmds.append(('BACKGROUND', (i, 1), (i, -1), colors.HexColor(day_color)))

    table.setStyle(TableStyle(style_cmds))
    elements.append(table)
    doc.build(elements)
    buf.seek(0)
    return buf


def generate_image(student_info, active_days, grid_rows, faculty_map={}):
    if not active_days:
        buf = io.BytesIO()
        Image.new('RGB', (600, 200), 'white').save(buf, 'PNG')
        buf.seek(0)
        return buf

    try:
        font_bold = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf", 18)
        font_header = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf", 14)
        font_course = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf", 13)
        font_sec = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", 11)
        font_room = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", 11)
        font_time = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", 11)
        font_info = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", 13)
    except Exception:
        fonts = [ImageFont.load_default()] * 7
        font_bold, font_header, font_course, font_sec, font_room, font_time, font_info = fonts

    tmp_img = Image.new('RGB', (1, 1))
    meas = ImageDraw.Draw(tmp_img)

    pad = 12
    line_gap = 4
    max_time_w = max(meas.textbbox((0, 0), f"{r['start']} - {r['end']}", font=font_time)[2] for r in grid_rows)
    time_col_w = max(max_time_w + pad * 2, 80)

    def make_sec_line(c, faculty_map):
        sec = c.get('sec', '')
        if not sec:
            return ''
        s = f"Sec-{sec}"
        f = get_faculty(c, faculty_map)
        if f:
            s += f"({f})"
        return s

    def cell_lines(c, faculty_map, max_w):
        course_w = meas.textbbox((0, 0), c['course'], font=font_course)[2]
        lines = [('course', c['course'])]
        sec_line = make_sec_line(c, faculty_map)
        if sec_line:
            lines.append(('sec', sec_line))
        room_max = max_w - pad * 2
        if meas.textbbox((0, 0), c['room'], font=font_room)[2] <= room_max:
            lines.append(('room', c['room']))
        else:
            words = c['room'].split(' ')
            current = ''
            for w in words:
                test = current + (' ' if current else '') + w
                if meas.textbbox((0, 0), test, font=font_room)[2] <= room_max:
                    current = test
                else:
                    if current:
                        lines.append(('room', current))
                    current = w
            if current:
                lines.append(('room', current))
        return lines

    day_col_w = 0
    for day in active_days:
        dw = meas.textbbox((0, 0), DAY_SHORT.get(day, day[:3]), font=font_header)[2]
        day_col_w = max(day_col_w, dw)
    for row in grid_rows:
        for day in active_days:
            c = row['courses'].get(day)
            if c:
                for ltype, text in cell_lines(c, faculty_map, 9999):
                    f = font_course if ltype == 'course' else (font_sec if ltype == 'sec' else font_room)
                    w = meas.textbbox((0, 0), text, font=f)[2]
                    day_col_w = max(day_col_w, w)
    day_col_w = min(max(day_col_w + pad * 2, 130), 260)

    col_widths = [time_col_w] + [day_col_w] * len(active_days)
    total_width = sum(col_widths)

    row_heights = []
    for row in grid_rows:
        max_h = 50
        for day in active_days:
            c = row['courses'].get(day)
            if c:
                cls = cell_lines(c, faculty_map, day_col_w)
                h = len(cls) * 22 + 10
                max_h = max(max_h, h)
        row_heights.append(max_h)

    header_height = 44
    title_height = 70
    total_height = title_height + header_height + sum(row_heights) + 30

    img = Image.new('RGB', (total_width + 60, total_height), 'white')
    draw = ImageDraw.Draw(img)

    y = 14
    draw.text((30, y), f"Class Routine - {student_info['semester']}", fill='#1a73e8', font=font_bold)
    y += 28
    draw.text((30, y), f"{student_info['name']}  |  {student_info['id']}  |  {student_info['semester']}", fill='#555555', font=font_info)
    y += 38

    x_start = 30
    x = x_start
    draw.rectangle([x, y, x + total_width, y + header_height], fill='#1a73e8')
    tx = x + (col_widths[0] - meas.textbbox((0, 0), 'Time', font=font_header)[2]) / 2
    draw.text((tx, y + 11), 'Time', fill='white', font=font_header)
    x += col_widths[0]
    for day in active_days:
        dx = x + (col_widths[1] - meas.textbbox((0, 0), DAY_SHORT.get(day, day[:3]), font=font_header)[2]) / 2
        draw.text((dx, y + 11), DAY_SHORT.get(day, day[:3]), fill='white', font=font_header)
        x += col_widths[1]

    y += header_height
    line_heights = {'course': 22, 'sec': 18, 'room': 18}
    line_fonts = {'course': font_course, 'sec': font_sec, 'room': font_room}
    line_colors = {'course': '#1a1a2e', 'sec': '#555555', 'room': '#555555'}

    for ri, row in enumerate(grid_rows):
        rh = row_heights[ri]
        x = x_start
        bg = '#F8F9FA' if ri % 2 == 0 else '#FFFFFF'
        draw.rectangle([x, y, x + col_widths[0], y + rh], fill=bg, outline='#CCCCCC')
        tw = meas.textbbox((0, 0), f"{row['start']} - {row['end']}", font=font_time)[2]
        draw.text((x + (col_widths[0] - tw) / 2, y + (rh - 16) / 2), f"{row['start']} - {row['end']}", fill='#333333', font=font_time)
        x += col_widths[0]

        for day in active_days:
            c = row['courses'].get(day)
            day_color = DAY_COLORS.get(day, '#FFFFFF')
            draw.rectangle([x, y, x + col_widths[1], y + rh], fill=day_color, outline='#CCCCCC')
            if c:
                cls = cell_lines(c, faculty_map, day_col_w)
                total_text_h = sum(line_heights.get(t, 18) for t, _ in cls)
                cy = y + (rh - total_text_h) // 2
                for ltype, text in cls:
                    lh = line_heights.get(ltype, 18)
                    lf = line_fonts.get(ltype, font_room)
                    lc = line_colors.get(ltype, '#555555')
                    tw2 = meas.textbbox((0, 0), text, font=lf)[2]
                    draw.text((x + (col_widths[1] - tw2) / 2, cy), text, fill=lc, font=lf)
                    cy += lh
            x += col_widths[1]
        y += rh

    buf = io.BytesIO()
    img.save(buf, 'PNG')
    buf.seek(0)
    return buf


if __name__ == '__main__':
    import sys
    debug = '--debug' in sys.argv
    app.run(debug=debug, host='0.0.0.0', port=5000)
