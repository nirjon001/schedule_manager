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

COLORS_LIGHT = ['#E3F2FD', '#F3E5F5', '#E8F5E9', '#FFF3E0', '#FCE4EC', '#F5F5F5', '#E0F7FA']
COLORS_DARK = ['#BBDEFB', '#E1BEE7', '#C8E6C9', '#FFE0B2', '#F8BBD0', '#E0E0E0', '#B2EBF2']

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
        active_days, time_slots, grid_rows = build_grid(schedule)
        unique_courses = []
        seen = set()
        for row in grid_rows:
            for ts in time_slots:
                c = row['courses'].get(ts)
                if c and c['course'] not in seen:
                    seen.add(c['course'])
                    unique_courses.append(c)
        return render_template('schedule.html',
                               student=student_info,
                               schedule=schedule,
                               active_days=active_days,
                               time_slots=time_slots,
                               grid_rows=grid_rows,
                               day_short=DAY_SHORT,
                               day_colors=DAY_COLORS,
                               colors_light=COLORS_LIGHT,
                               colors_dark=COLORS_DARK,
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
    active_days, time_slots, grid_rows = build_grid(schedule)
    faculty_raw = request.args.get('faculty', '')
    try:
        faculty_map = json.loads(faculty_raw) if faculty_raw else {}
    except (json.JSONDecodeError, TypeError):
        faculty_map = {}
    buf = generate_pdf(student_info, active_days, time_slots, grid_rows, faculty_map)
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
    active_days, time_slots, grid_rows = build_grid(schedule)
    faculty_raw = request.args.get('faculty', '')
    try:
        faculty_map = json.loads(faculty_raw) if faculty_raw else {}
    except (json.JSONDecodeError, TypeError):
        faculty_map = {}
    buf = generate_image(student_info, active_days, time_slots, grid_rows, faculty_map)
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

def generate_pdf(student_info, active_days, time_slots, grid_rows, faculty_map={}):
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=landscape(A4),
                            leftMargin=15*mm, rightMargin=15*mm,
                            topMargin=15*mm, bottomMargin=15*mm)

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle('Title2', parent=styles['Title'], fontSize=18, spaceAfter=6)
    info_style = ParagraphStyle('Info', parent=styles['Normal'], fontSize=11, spaceAfter=6)
    cell_style = ParagraphStyle('Cell', parent=styles['Normal'], fontSize=10, leading=14, spaceAfter=0, spaceBefore=0, alignment=TA_CENTER)
    day_style = ParagraphStyle('DayCell', parent=cell_style, fontSize=11, fontName='Helvetica-Bold')

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

    time_labels = [f"{ts[0]} - {ts[1]}" for ts in time_slots]
    header = ['Day / Time'] + time_labels
    table_data = [header]

    page_w = landscape(A4)[0]
    available = page_w - 30*mm

    from reportlab.pdfbase.pdfmetrics import stringWidth
    day_name_max_w = max(stringWidth(d, 'Helvetica-Bold', 11) for d in active_days)
    for row in grid_rows:
        for ts in time_slots:
            c = row['courses'].get(ts)
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
                day_name_max_w = max(day_name_max_w, cw, sw, rw)
    day_col_w = max(day_name_max_w + 14, 90)

    time_max_w = max(stringWidth(tl, 'Helvetica-Bold', 11) for tl in time_labels)
    time_col_w = max(time_max_w + 14, 80)

    if len(time_slots) * time_col_w + day_col_w > available:
        time_col_w = int((available - day_col_w) / len(time_slots))
    if time_col_w < 60:
        time_col_w = 60
        day_col_w = int(available - time_col_w * len(time_slots))
        if day_col_w < 50:
            day_col_w = 50
            time_col_w = int((available - 50) / len(time_slots))

    col_widths = [day_col_w] + [time_col_w] * len(time_slots)

    for row in grid_rows:
        row_vals = [Paragraph(row['day'], day_style)]
        for ts in time_slots:
            c = row['courses'].get(ts)
            if c:
                row_vals.append(Paragraph(make_cell_text(c, faculty_map), cell_style))
            else:
                row_vals.append('')
        table_data.append(row_vals)

    table = Table(table_data, colWidths=col_widths, repeatRows=1)

    style_cmds = [
        ('BACKGROUND', (0, 0), (0, 0), colors.HexColor('#1a73e8')),
        ('TEXTCOLOR', (0, 0), (0, 0), colors.white),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 11),
        ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 1), (-1, -1), 10),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
        ('TOPPADDING', (0, 0), (-1, -1), 6),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
    ]

    for i in range(len(time_slots)):
        col = i + 1
        c = COLORS_DARK[i % len(COLORS_DARK)]
        style_cmds.append(('BACKGROUND', (col, 0), (col, 0), colors.HexColor(c)))
        style_cmds.append(('TEXTCOLOR', (col, 0), (col, 0), colors.HexColor('#333333')))

    style_cmds.append(('BACKGROUND', (0, 1), (0, -1), colors.HexColor('#BBDEFB')))
    style_cmds.append(('TEXTCOLOR', (0, 1), (0, -1), colors.HexColor('#333333')))
    for i in range(len(time_slots)):
        col = i + 1
        c = COLORS_LIGHT[i % len(COLORS_LIGHT)]
        style_cmds.append(('BACKGROUND', (col, 1), (col, -1), colors.HexColor(c)))

    table.setStyle(TableStyle(style_cmds))
    elements.append(table)
    doc.build(elements)
    buf.seek(0)
    return buf


def generate_image(student_info, active_days, time_slots, grid_rows, faculty_map={}):
    if not active_days:
        buf = io.BytesIO()
        Image.new('RGB', (600, 200), 'white').save(buf, 'PNG')
        buf.seek(0)
        return buf

    def load_font(path, size):
        try:
            return ImageFont.truetype(path, size)
        except Exception:
            return None

    font_paths = [
        ("C:/Windows/Fonts/arialbd.ttf", "C:/Windows/Fonts/arial.ttf"),
        ("C:/Windows/Fonts/segoeuib.ttf", "C:/Windows/Fonts/segoeui.ttf"),
        ("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf", "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"),
    ]

    font_bold = font_header = font_course = font_day = None
    font_sec = font_room = font_info = None
    for bold_path, regular_path in font_paths:
        fb = load_font(bold_path, 18)
        if fb:
            font_bold = fb
            font_header = load_font(bold_path, 14)
            font_course = load_font(bold_path, 13)
            font_day = load_font(bold_path, 13)
            font_sec = load_font(regular_path, 11)
            font_room = load_font(regular_path, 11)
            font_info = load_font(regular_path, 13)
            break

    if not font_bold:
        fonts = [ImageFont.load_default()] * 7
        font_bold, font_header, font_course, font_sec, font_room, font_day, font_info = fonts

    tmp_img = Image.new('RGB', (1, 1))
    meas = ImageDraw.Draw(tmp_img)

    pad = 12
    time_labels = [f"{ts[0]} - {ts[1]}" for ts in time_slots]

    day_name_max_w = max(meas.textbbox((0, 0), d, font=font_day)[2] for d in active_days)
    day_col_w = max(day_name_max_w + pad * 2, 90)

    time_header_w = max(meas.textbbox((0, 0), tl, font=font_header)[2] for tl in time_labels)
    time_col_w = max(time_header_w + pad * 2, 80)

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

    for row in grid_rows:
        for ts in time_slots:
            c = row['courses'].get(ts)
            if c:
                for ltype, text in cell_lines(c, faculty_map, 9999):
                    f = font_course if ltype == 'course' else (font_sec if ltype == 'sec' else font_room)
                    w = meas.textbbox((0, 0), text, font=f)[2]
                    time_col_w = max(time_col_w, w + pad * 2)

    col_widths = [day_col_w] + [time_col_w] * len(time_slots)
    total_width = sum(col_widths)

    row_heights = []
    for row in grid_rows:
        max_h = 50
        for ts in time_slots:
            c = row['courses'].get(ts)
            if c:
                cls = cell_lines(c, faculty_map, time_col_w)
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
    draw.rectangle([x, y, x + col_widths[0], y + header_height], fill='#1a73e8')
    dx = x + (col_widths[0] - meas.textbbox((0, 0), 'Day', font=font_header)[2]) / 2
    draw.text((dx, y + 11), 'Day / Time', fill='white', font=font_header)
    x += col_widths[0]
    for i, tl in enumerate(time_labels):
        c = COLORS_DARK[i % len(COLORS_DARK)]
        draw.rectangle([x, y, x + col_widths[1], y + header_height], fill=c, outline='#CCCCCC')
        dx2 = x + (col_widths[1] - meas.textbbox((0, 0), tl, font=font_header)[2]) / 2
        draw.text((dx2, y + 11), tl, fill='#333333', font=font_header)
        x += col_widths[1]

    y += header_height
    line_heights = {'course': 22, 'sec': 18, 'room': 18}
    line_fonts = {'course': font_course, 'sec': font_sec, 'room': font_room}
    line_colors = {'course': '#1a1a2e', 'sec': '#555555', 'room': '#555555'}

    for ri, row in enumerate(grid_rows):
        rh = row_heights[ri]
        x = x_start
        draw.rectangle([x, y, x + col_widths[0], y + rh], fill='#BBDEFB', outline='#CCCCCC')
        dw = meas.textbbox((0, 0), row['day'], font=font_day)[2]
        draw.text((x + (col_widths[0] - dw) / 2, y + (rh - 16) / 2), row['day'], fill='#333333', font=font_day)
        x += col_widths[0]
        for ci, ts in enumerate(time_slots):
            c = row['courses'].get(ts)
            col_color = COLORS_LIGHT[ci % len(COLORS_LIGHT)]
            draw.rectangle([x, y, x + col_widths[1], y + rh], fill=col_color, outline='#CCCCCC')
            if c:
                cls = cell_lines(c, faculty_map, time_col_w)
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
