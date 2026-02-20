# -*- coding: utf-8 -*-
"""
Ứng dụng Flask quản lý trích xuất và nạp bản dịch cho file Excel, PowerPoint và Word
"""

import os
import json
import zipfile
import uuid
import shutil
import io
from datetime import datetime, timedelta
from urllib.parse import quote
from flask import Flask, render_template, request, send_file, jsonify, session, redirect, url_for
from werkzeug.utils import secure_filename
from openpyxl import load_workbook
from pptx import Presentation
from docx import Document
from functools import wraps

# Khởi tạo ứng dụng Flask
app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 60 * 1024 * 1024  # Giới hạn 50MB
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['SECRET_KEY'] = os.urandom(24)  # Secret key cho session
app.config['PERMANENT_SESSION_LIFETIME'] = timedelta(hours=8)  # Session timeout 8h

# Các định dạng file được phép
ALLOWED_EXTENSIONS = {'xlsx', 'pptx', 'docx'}

# Đọc password từ file
PASSWORD_FILE = 'password.txt'

# File lưu Prompt Templates
TEMPLATES_FILE = 'prompt_templates.json'

# ==================== HELPER: PROMPT TEMPLATES ====================
def get_default_templates():
    """Trả về danh sách template mặc định"""
    return [
        {
            "id": "formal",
            "name": "Dịch chính xác (Formal)",
            "content": "Hãy dịch các giá trị (values) trong file JSON này sang {TARGET_LANG}.\n\nPhong cách: Chính xác, chuyên nghiệp, dùng trong tài liệu kinh doanh.\n\nQuy tắc bắt buộc:\n1. Giữ nguyên 100% các keys\n2. CHỈ dịch nội dung bên trong values\n3. KHÔNG dịch từ/cụm từ đã là ngôn ngữ đích\n4. KHÔNG dịch mã kỹ thuật, placeholder, tên biến\n5. KHÔNG dịch số, ngày tháng, ký hiệu đặc biệt\n6. Giữ nguyên format JSON chuẩn\n\n⚠️ QUY TẮc về dấu ngoặc kép: CHỈ dùng \" (U+0022). KHÔNG dùng \u201c \u201d \u201e \u201f \u00ab \u00bb\nTrích dẫn: dùng \u300c \u300dhoặc 'đơn'\n\nOutput: Trả về ĐÚNG cấu trúc JSON, KHÔNG thêm giải thích."
        },
        {
            "id": "casual",
            "name": "Dịch tự nhiên (Casual)",
            "content": "Hãy dịch các giá trị (values) trong file JSON này sang {TARGET_LANG}.\n\nPhong cách: Tự nhiên, thân thiện, dễ đọc - phù hợp cho giao diện người dùng.\n\nQuy tắc bắt buộc:\n1. Giữ nguyên 100% các keys\n2. CHỈ dịch nội dung bên trong values\n3. KHÔNG dịch từ/cụm từ đã là ngôn ngữ đích\n4. KHÔNG dịch mã kỹ thuật, placeholder, tên biến\n5. KHÔNG dịch số, ngày tháng, ký hiệu đặc biệt\n6. Giữ nguyên format JSON chuẩn\n\n⚠️ QUY TẮc về dấu ngoặc kép: CHỈ dùng \" (U+0022). KHÔNG dùng \u201c \u201d \u201e \u201f \u00ab \u00bb\n\nOutput: Trả về ĐÚNG cấu trúc JSON, KHÔNG thêm giải thích."
        },
        {
            "id": "technical",
            "name": "Dịch kỹ thuật (Technical)",
            "content": "Hãy dịch các giá trị (values) trong file JSON này sang {TARGET_LANG}.\n\nPhong cách: Kỹ thuật, chính xác cao, giữ nguyên thuật ngữ IT.\n\nQuy tắc bắt buộc:\n1. Giữ nguyên 100% các keys\n2. CHỈ dịch nội dung bên trong values\n3. KHÔNG dịch từ/cụm từ đã là ngôn ngữ đích\n4. KHÔNG dịch placeholder ({0}, %s, $n...), tên biến\n5. KHÔNG dịch số, ngày tháng, ký hiệu đặc biệt\n6. Giữ nguyên thuật ngữ IT tiếng Anh nếu không có từ tương đương chính xác\n7. Giữ nguyên format JSON chuẩn\n\n⚠️ QUY TẮc về dấu ngoặc kép: CHỈ dùng \" (U+0022). KHÔNG dùng \u201c \u201d \u201e \u201f \u00ab \u00bb\n\nOutput: Trả về ĐÚNG cấu trúc JSON, KHÔNG thêm giải thích."
        }
    ]

def load_templates(lang='default'):
    """Đọc prompt templates cho một ngôn ngữ cụ thể (fallback về default)"""
    try:
        with open(TEMPLATES_FILE, 'r', encoding='utf-8') as f:
            data = json.load(f)
        # Tương thích ngược: nếu data là array thì đó là format cũ
        if isinstance(data, list):
            return data
        return data.get(lang) or data.get('default') or get_default_templates()
    except Exception:
        return get_default_templates()


def build_dedup_data(extracted_data, chunk_size=400):
    """
    Gộp các keys có cùng value để giảm số lượng cần dịch.
    Returns: (dedup_files, mapping, stats)
      - dedup_files: list of {name, content} – các chunk dedup (giống format files thường)
      - mapping: {dedup_key: [orig_key1, orig_key2, ...]}
      - stats: {total, unique, saved, percent_saved}
    """
    # Group keys by value (giữ order)
    value_to_keys = {}
    for key, value in extracted_data.items():
        if value not in value_to_keys:
            value_to_keys[value] = []
        value_to_keys[value].append(key)

    # Build dedup dict và mapping
    dedup_data = {}
    mapping = {}  # dedup_key → [original_keys]
    for idx, (value, keys) in enumerate(value_to_keys.items(), 1):
        dk = f'dedup_{idx}'
        dedup_data[dk] = value
        mapping[dk] = keys

    total = len(extracted_data)
    unique = len(dedup_data)
    saved = total - unique
    percent = round(saved * 100 / total) if total > 0 else 0
    stats = {'total': total, 'unique': unique, 'saved': saved, 'percent_saved': percent}

    # Chia thành các chunk
    items = list(dedup_data.items())
    num_chunks = max(1, (unique + chunk_size - 1) // chunk_size)
    dedup_files = []
    for i in range(num_chunks):
        chunk = dict(items[i * chunk_size:(i + 1) * chunk_size])
        dedup_files.append({
            'name': f'dedup_part{i+1:02d}_of_{num_chunks:02d}.json',
            'content': json.dumps(chunk, ensure_ascii=False, indent=2)
        })

    return dedup_files, mapping, stats


def expand_dedup_data(json_data, session_folder):
    """
    Mở rộng dedup JSON (dedup_N → value) thành keys gốc dựa trên mapping đã lưu.
    Nếu không tìm thấy mapping file thì trả về nguyên.
    """
    mapping_path = os.path.join(session_folder, 'dedup_mapping.json')
    if not os.path.exists(mapping_path):
        return json_data
    try:
        with open(mapping_path, 'r', encoding='utf-8') as f:
            mapping = json.load(f)
        expanded = {}
        for key, value in json_data.items():
            if key.startswith('dedup_') and key in mapping:
                for orig_key in mapping[key]:
                    expanded[orig_key] = value
            else:
                expanded[key] = value  # key thường, giữ nguyên
        return expanded
    except Exception:
        return json_data


def get_password():
    """Đọc password từ file password.txt"""
    try:
        with open(PASSWORD_FILE, 'r', encoding='utf-8') as f:
            return f.read().strip()
    except FileNotFoundError:
        # Nếu file không tồn tại, tạo file với password mặc định
        default_password = 'admin123'
        with open(PASSWORD_FILE, 'w', encoding='utf-8') as f:
            f.write(default_password)
        return default_password

def get_machine_id():
    """Lấy ID máy (dựa trên UUID node)"""
    return hex(uuid.getnode())

def create_session_id():
    """Tạo session ID dựa trên machine ID + timestamp"""
    machine_id = get_machine_id()
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    return f"{machine_id}_{timestamp}"

def get_session_folder():
    """Lấy đường dẫn folder của session hiện tại"""
    if 'session_id' not in session:
        session['session_id'] = create_session_id()
    
    session_folder = os.path.join(app.config['UPLOAD_FOLDER'], session['session_id'])
    os.makedirs(session_folder, exist_ok=True)
    return session_folder

def cleanup_old_sessions():
    """Xóa tất cả folder của các phiên từ hôm qua trở về trước"""
    try:
        upload_folder = app.config['UPLOAD_FOLDER']
        if not os.path.exists(upload_folder):
            return
        
        # Lấy ngày hiện tại (không có giờ phút giây)
        today = datetime.now().date()
        
        # Duyệt qua tất cả các folder trong uploads
        for folder_name in os.listdir(upload_folder):
            folder_path = os.path.join(upload_folder, folder_name)
            
            if os.path.isdir(folder_path):
                try:
                    # Parse timestamp từ tên folder (format: machine_YYYYMMDD_HHMMSS)
                    parts = folder_name.split('_')
                    if len(parts) >= 2:
                        date_str = parts[-2]  # YYYYMMDD
                        folder_date = datetime.strptime(date_str, '%Y%m%d').date()
                        
                        # Nếu folder từ hôm qua trở về trước, xóa đi
                        if folder_date < today:
                            shutil.rmtree(folder_path)
                            print(f"Đã xóa folder cũ: {folder_name}")
                except (ValueError, IndexError):
                    # Nếu không parse được, bỏ qua
                    continue
    except Exception as e:
        print(f"Lỗi khi cleanup old sessions: {e}")

def login_required(f):
    """Decorator để yêu cầu đăng nhập"""
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if 'logged_in' not in session:
            # Nếu là request AJAX/JSON, trả về JSON thay vì redirect
            if request.path.startswith('/api') or request.is_json or request.path in ['/extract', '/inject', '/clear-uploads']:
                return jsonify({'error': 'Chưa đăng nhập hoặc phiên đã hết hạn'}), 401
            return redirect(url_for('login'))
        return f(*args, **kwargs)
    return decorated_function

def allowed_file(filename):
    """
    Kiểm tra xem file có phải định dạng được phép không
    """
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def set_download_headers(response, display_name, default_ascii_name):
    """Set Content-Disposition hỗ trợ tên file Unicode (RFC 5987)."""
    encoded_filename = quote(display_name, safe='')
    ascii_filename = secure_filename(display_name) or default_ascii_name
    response.headers['Content-Disposition'] = (
        f"attachment; filename=\"{ascii_filename}\"; filename*=UTF-8''{encoded_filename}"
    )
    return response

# Error Handlers
@app.errorhandler(400)
def bad_request(error):
    """Xử lý lỗi 400 Bad Request"""
    if request.path.startswith('/api') or request.is_json or request.path in ['/extract', '/inject', '/clear-uploads']:
        return jsonify({'error': str(error) or 'Yêu cầu không hợp lệ'}), 400
    return str(error), 400

@app.errorhandler(401)
def unauthorized(error):
    """Xử lý lỗi 401 Unauthorized"""
    if request.path.startswith('/api') or request.is_json or request.path in ['/extract', '/inject', '/clear-uploads']:
        return jsonify({'error': 'Chưa đăng nhập hoặc phiên đã hết hạn'}), 401
    return redirect(url_for('login'))

@app.errorhandler(404)
def not_found(error):
    """Xử lý lỗi 404 Not Found"""
    if request.path.startswith('/api') or request.is_json:
        return jsonify({'error': 'Không tìm thấy tài nguyên'}), 404
    return str(error), 404

@app.errorhandler(413)
def request_entity_too_large(error):
    """Xử lý lỗi 413 Request Entity Too Large"""
    return jsonify({'error': 'File quá lớn. Giới hạn 50MB'}), 413

@app.errorhandler(500)
def internal_server_error(error):
    """Xử lý lỗi 500 Internal Server Error"""
    if request.path.startswith('/api') or request.is_json or request.path in ['/extract', '/inject', '/clear-uploads']:
        return jsonify({'error': f'Lỗi máy chủ: {str(error)}'}), 500
    return str(error), 500

def extract_text_from_shape(shape, shape_path, extracted_data):
    """
    Hàm đệ quy để trích xuất text từ shape, bao gồm cả grouped shapes
    shape_path: đường dẫn đến shape, ví dụ "Shape1" hoặc "Shape1_2_3"
    """
    # Trích xuất text từ text frame của shape hiện tại
    if hasattr(shape, "text") and shape.text:
        text_content = shape.text.strip()
        if text_content:  # Chỉ lấy nội dung không rỗng
            extracted_data[shape_path] = text_content
    
    # Trích xuất text từ table nếu có
    if hasattr(shape, "has_table") and shape.has_table:
        table = shape.table
        for row_idx, row in enumerate(table.rows, start=1):
            for col_idx, cell in enumerate(row.cells, start=1):
                if cell.text.strip():
                    key = f"{shape_path}!Table_R{row_idx}C{col_idx}"
                    extracted_data[key] = cell.text.strip()
    
    # Kiểm tra xem shape có phải là GroupShape không (chứa các shape con)
    if hasattr(shape, "shapes"):
        # Đây là grouped shape, duyệt qua các shape con
        for child_idx, child_shape in enumerate(shape.shapes, start=1):
            child_path = f"{shape_path}_{child_idx}"
            extract_text_from_shape(child_shape, child_path, extracted_data)

def extract_text_from_pptx(filepath):
    """
    Trích xuất text từ file PPTX, bao gồm cả text trong grouped shapes
    Trả về dictionary với format: {"SlideX!ShapeY": "Content"}
    Với nested shapes: {"SlideX!ShapeY_Z": "Content"} (Z là shape con)
    """
    extracted_data = {}
    prs = Presentation(filepath)
    
    for slide_idx, slide in enumerate(prs.slides, start=1):
        for shape_idx, shape in enumerate(slide.shapes, start=1):
            shape_path = f"Slide{slide_idx}!Shape{shape_idx}"
            extract_text_from_shape(shape, shape_path, extracted_data)
    
    return extracted_data

def inject_text_to_shape(shape, shape_indices, translated_value, is_table_cell=False, table_pos=None):
    """
    Hàm đệ quy để nạp text vào shape, bao gồm cả grouped shapes
    shape_indices: list các index để navigate đến shape đúng, ví dụ [2, 3] cho Shape2_3
    is_table_cell: có phải là table cell không
    table_pos: tuple (row_idx, col_idx) nếu là table cell
    """
    # Nếu là shape cuối cùng trong path
    if len(shape_indices) == 0:
        if is_table_cell and table_pos:
            # Nạp vào table cell - giữ nguyên định dạng
            row_idx, col_idx = table_pos
            if hasattr(shape, "has_table") and shape.has_table:
                table = shape.table
                if row_idx < len(table.rows) and col_idx < len(table.rows[row_idx].cells):
                    cell = table.rows[row_idx].cells[col_idx]
                    # Thay thế text trong từng paragraph/run để giữ định dạng
                    if cell.text_frame:
                        replace_text_keep_format(cell.text_frame, translated_value)
        else:
            # Nạp vào text frame của shape - giữ nguyên định dạng
            if hasattr(shape, "text_frame") and shape.text_frame:
                replace_text_keep_format(shape.text_frame, translated_value)
        return True
    
    # Navigate đến shape con
    if hasattr(shape, "shapes"):
        next_idx = shape_indices[0]
        if next_idx <= len(shape.shapes):
            child_shape = shape.shapes[next_idx - 1]  # Chuyển từ 1-indexed sang 0-indexed
            return inject_text_to_shape(child_shape, shape_indices[1:], translated_value, is_table_cell, table_pos)
    
    return False

def replace_text_keep_format(text_frame, new_text):
    """
    Thay thế text trong text_frame nhưng giữ nguyên định dạng (font, màu, gạch chân, bold, italic...)
    Chiến lược:
    1. Nếu toàn bộ text frame chỉ có 1 paragraph và 1 run -> thay text của run đó
    2. Nếu có nhiều runs/paragraphs -> xóa text của tất cả runs, gán text mới vào run đầu tiên với định dạng gốc
    """
    if not text_frame.paragraphs:
        return
    
    # Thu thập tất cả runs từ tất cả paragraphs
    all_runs = []
    for paragraph in text_frame.paragraphs:
        for run in paragraph.runs:
            all_runs.append(run)
    
    if not all_runs:
        # Không có run nào, tạo mới
        if text_frame.paragraphs:
            text_frame.paragraphs[0].text = new_text
        return
    
    # Lưu định dạng của run đầu tiên
    first_run = all_runs[0]
    
    # Xóa text của tất cả runs
    for run in all_runs:
        run.text = ""
    
    # Gán text mới vào run đầu tiên (giữ nguyên định dạng)
    first_run.text = new_text

def inject_text_to_pptx(filepath, json_data):
    """
    Nạp text đã dịch vào file PPTX, bao gồm cả grouped shapes
    """
    prs = Presentation(filepath)
    
    for key, translated_value in json_data.items():
        try:
            # Parse key format: 
            # "SlideX!ShapeY" hoặc "SlideX!ShapeY_Z" (nested) 
            # hoặc "SlideX!ShapeY!Table_RxCy" hoặc "SlideX!ShapeY_Z!Table_RxCy"
            if '!' not in key:
                continue
            
            parts = key.split('!')
            if len(parts) < 2:
                continue
            
            # Lấy slide index
            slide_part = parts[0]
            if not slide_part.startswith('Slide'):
                continue
            slide_idx = int(slide_part.replace('Slide', '')) - 1
            
            if slide_idx >= len(prs.slides):
                continue
            
            slide = prs.slides[slide_idx]
            
            # Parse shape path: "Shape2" hoặc "Shape2_3_1" (nested)
            shape_part = parts[1]
            if not shape_part.startswith('Shape'):
                continue
            
            # Tách các indices: "Shape2_3_1" -> [2, 3, 1]
            shape_str = shape_part.replace('Shape', '')
            shape_indices = [int(idx) for idx in shape_str.split('_')]
            
            # Lấy shape đầu tiên (top-level shape)
            first_shape_idx = shape_indices[0] - 1  # Chuyển sang 0-indexed
            if first_shape_idx >= len(slide.shapes):
                continue
            
            shape = slide.shapes[first_shape_idx]
            
            # Kiểm tra xem có phải table cell không
            is_table_cell = False
            table_pos = None
            
            if len(parts) == 3 and parts[2].startswith('Table_R'):
                # Parse table cell position
                is_table_cell = True
                table_part = parts[2].replace('Table_R', '').split('C')
                row_idx = int(table_part[0]) - 1
                col_idx = int(table_part[1]) - 1
                table_pos = (row_idx, col_idx)
            
            # Navigate và nạp text (bỏ qua index đầu tiên vì đã lấy shape rồi)
            inject_text_to_shape(shape, shape_indices[1:], translated_value, is_table_cell, table_pos)
            
        except (ValueError, IndexError, AttributeError) as e:
            # Bỏ qua các key không hợp lệ
            continue
    
    return prs

def extract_text_from_docx(filepath):
    """
    Trích xuất text từ file DOCX, bao gồm paragraphs, tables, headers, footers
    Trả về dictionary với format:
    - Paragraphs: {"ParagraphX": "Content"}
    - Tables: {"TableX!RyC z": "Content"}
    - Headers: {"Header_SectionX!ParagraphY": "Content"}
    - Footers: {"Footer_SectionX!ParagraphY": "Content"}
    """
    extracted_data = {}
    doc = Document(filepath)
    
    # 1. Trích xuất text từ các paragraph thông thường (không trong table)
    paragraph_idx = 0
    for para in doc.paragraphs:
        text_content = para.text.strip()
        if text_content:  # Chỉ lấy paragraph không rỗng
            paragraph_idx += 1
            key = f"Paragraph{paragraph_idx}"
            extracted_data[key] = text_content
    
    # 2. Trích xuất text từ các bảng
    for table_idx, table in enumerate(doc.tables, start=1):
        for row_idx, row in enumerate(table.rows, start=1):
            for col_idx, cell in enumerate(row.cells, start=1):
                text_content = cell.text.strip()
                if text_content:
                    key = f"Table{table_idx}!R{row_idx}C{col_idx}"
                    extracted_data[key] = text_content
    
    # 3. Trích xuất text từ headers
    for section_idx, section in enumerate(doc.sections, start=1):
        header = section.header
        for para_idx, para in enumerate(header.paragraphs, start=1):
            text_content = para.text.strip()
            if text_content:
                key = f"Header_Section{section_idx}!Paragraph{para_idx}"
                extracted_data[key] = text_content
        
        # Trích xuất từ table trong header (nếu có)
        for table_idx, table in enumerate(header.tables, start=1):
            for row_idx, row in enumerate(table.rows, start=1):
                for col_idx, cell in enumerate(row.cells, start=1):
                    text_content = cell.text.strip()
                    if text_content:
                        key = f"Header_Section{section_idx}!Table{table_idx}!R{row_idx}C{col_idx}"
                        extracted_data[key] = text_content
    
    # 4. Trích xuất text từ footers
    for section_idx, section in enumerate(doc.sections, start=1):
        footer = section.footer
        for para_idx, para in enumerate(footer.paragraphs, start=1):
            text_content = para.text.strip()
            if text_content:
                key = f"Footer_Section{section_idx}!Paragraph{para_idx}"
                extracted_data[key] = text_content
        
        # Trích xuất từ table trong footer (nếu có)
        for table_idx, table in enumerate(footer.tables, start=1):
            for row_idx, row in enumerate(table.rows, start=1):
                for col_idx, cell in enumerate(row.cells, start=1):
                    text_content = cell.text.strip()
                    if text_content:
                        key = f"Footer_Section{section_idx}!Table{table_idx}!R{row_idx}C{col_idx}"
                        extracted_data[key] = text_content
    
    return extracted_data

def replace_text_keep_format_docx(paragraph, new_text):
    """
    Thay thế text trong paragraph của Word nhưng giữ nguyên định dạng (font, màu, bold, italic...)
    Chiến lược:
    1. Lưu định dạng của run đầu tiên
    2. Xóa text của tất cả runs
    3. Gán text mới vào run đầu tiên (giữ nguyên định dạng)
    """
    if not paragraph.runs:
        # Không có run nào, tạo mới
        paragraph.text = new_text
        return
    
    # Lưu định dạng của run đầu tiên
    first_run = paragraph.runs[0]
    
    # Xóa text của tất cả runs
    for run in paragraph.runs:
        run.text = ""
    
    # Gán text mới vào run đầu tiên (giữ nguyên định dạng)
    first_run.text = new_text

def inject_text_to_docx(filepath, json_data):
    """
    Nạp text đã dịch vào file DOCX
    Giữ nguyên định dạng (font, màu, size, bold, italic...)
    """
    doc = Document(filepath)
    
    for key, translated_value in json_data.items():
        try:
            # 1. Xử lý Paragraph thông thường: "ParagraphX"
            if key.startswith('Paragraph') and '!' not in key:
                para_num = int(key.replace('Paragraph', ''))
                # Đếm lại các paragraph không rỗng để map đúng index
                current_para_idx = 0
                for para in doc.paragraphs:
                    if para.text.strip():  # Chỉ đếm paragraph không rỗng
                        current_para_idx += 1
                        if current_para_idx == para_num:
                            replace_text_keep_format_docx(para, translated_value)
                            break
            
            # 2. Xử lý Table: "TableX!RyCz"
            elif key.startswith('Table') and '!' in key and not key.startswith('Header_') and not key.startswith('Footer_'):
                parts = key.split('!')
                if len(parts) != 2:
                    continue
                
                # Parse table index
                table_part = parts[0]
                table_idx = int(table_part.replace('Table', '')) - 1
                
                if table_idx >= len(doc.tables):
                    continue
                
                table = doc.tables[table_idx]
                
                # Parse cell position: "R2C3"
                cell_part = parts[1]
                if not cell_part.startswith('R'):
                    continue
                
                cell_parts = cell_part.replace('R', '').split('C')
                if len(cell_parts) != 2:
                    continue
                
                row_idx = int(cell_parts[0]) - 1
                col_idx = int(cell_parts[1]) - 1
                
                if row_idx < len(table.rows) and col_idx < len(table.rows[row_idx].cells):
                    cell = table.rows[row_idx].cells[col_idx]
                    # Thay thế text trong paragraph đầu tiên của cell
                    if cell.paragraphs:
                        replace_text_keep_format_docx(cell.paragraphs[0], translated_value)
            
            # 3. Xử lý Header: "Header_SectionX!ParagraphY" hoặc "Header_SectionX!TableY!RzCw"
            elif key.startswith('Header_Section'):
                parts = key.split('!')
                if len(parts) < 2:
                    continue
                
                # Parse section index
                section_part = parts[0].replace('Header_Section', '')
                section_idx = int(section_part) - 1
                
                if section_idx >= len(doc.sections):
                    continue
                
                header = doc.sections[section_idx].header
                
                # Check if it's a paragraph or table
                if parts[1].startswith('Paragraph'):
                    para_num = int(parts[1].replace('Paragraph', ''))
                    para_idx = 0
                    for para in header.paragraphs:
                        if para.text.strip():
                            para_idx += 1
                            if para_idx == para_num:
                                replace_text_keep_format_docx(para, translated_value)
                                break
                
                elif parts[1].startswith('Table') and len(parts) == 3:
                    # Header table cell
                    table_idx = int(parts[1].replace('Table', '')) - 1
                    if table_idx >= len(header.tables):
                        continue
                    
                    table = header.tables[table_idx]
                    cell_part = parts[2]
                    if not cell_part.startswith('R'):
                        continue
                    
                    cell_parts = cell_part.replace('R', '').split('C')
                    if len(cell_parts) != 2:
                        continue
                    
                    row_idx = int(cell_parts[0]) - 1
                    col_idx = int(cell_parts[1]) - 1
                    
                    if row_idx < len(table.rows) and col_idx < len(table.rows[row_idx].cells):
                        cell = table.rows[row_idx].cells[col_idx]
                        if cell.paragraphs:
                            replace_text_keep_format_docx(cell.paragraphs[0], translated_value)
            
            # 4. Xử lý Footer: "Footer_SectionX!ParagraphY" hoặc "Footer_SectionX!TableY!RzCw"
            elif key.startswith('Footer_Section'):
                parts = key.split('!')
                if len(parts) < 2:
                    continue
                
                # Parse section index
                section_part = parts[0].replace('Footer_Section', '')
                section_idx = int(section_part) - 1
                
                if section_idx >= len(doc.sections):
                    continue
                
                footer = doc.sections[section_idx].footer
                
                # Check if it's a paragraph or table
                if parts[1].startswith('Paragraph'):
                    para_num = int(parts[1].replace('Paragraph', ''))
                    para_idx = 0
                    for para in footer.paragraphs:
                        if para.text.strip():
                            para_idx += 1
                            if para_idx == para_num:
                                replace_text_keep_format_docx(para, translated_value)
                                break
                
                elif parts[1].startswith('Table') and len(parts) == 3:
                    # Footer table cell
                    table_idx = int(parts[1].replace('Table', '')) - 1
                    if table_idx >= len(footer.tables):
                        continue
                    
                    table = footer.tables[table_idx]
                    cell_part = parts[2]
                    if not cell_part.startswith('R'):
                        continue
                    
                    cell_parts = cell_part.replace('R', '').split('C')
                    if len(cell_parts) != 2:
                        continue
                    
                    row_idx = int(cell_parts[0]) - 1
                    col_idx = int(cell_parts[1]) - 1
                    
                    if row_idx < len(table.rows) and col_idx < len(table.rows[row_idx].cells):
                        cell = table.rows[row_idx].cells[col_idx]
                        if cell.paragraphs:
                            replace_text_keep_format_docx(cell.paragraphs[0], translated_value)
        
        except (ValueError, IndexError, AttributeError) as e:
            # Bỏ qua các key không hợp lệ
            continue
    
    return doc

@app.route('/login', methods=['GET', 'POST'])
def login():
    """Trang đăng nhập"""
    if request.method == 'POST':
        password = request.form.get('password', '')
        correct_password = get_password()
        
        if password == correct_password:
            session.permanent = True
            session['logged_in'] = True
            session['session_id'] = create_session_id()
            
            # Cleanup old sessions khi đăng nhập
            cleanup_old_sessions()
            
            return redirect(url_for('index'))
        else:
            return render_template('login.html', error='Mật khẩu không đúng!')
    
    return render_template('login.html')

@app.route('/logout')
def logout():
    """Đăng xuất"""
    # Xóa folder của session hiện tại
    if 'session_id' in session:
        session_folder = os.path.join(app.config['UPLOAD_FOLDER'], session['session_id'])
        if os.path.exists(session_folder):
            shutil.rmtree(session_folder)
    
    session.clear()
    return redirect(url_for('login'))

@app.route('/')
@login_required
def index():
    """
    Trang chủ hiển thị dashboard với 2 chức năng Extract và Inject
    """
    # Cleanup old sessions mỗi khi load trang
    cleanup_old_sessions()
    return render_template('index.html')

@app.route('/api/languages', methods=['GET'])
@login_required
def get_languages():
    """
    Trả về danh sách ngôn ngữ đích từ file languages.json
    """
    languages_file = os.path.join(os.path.dirname(__file__), 'languages.json')
    try:
        with open(languages_file, 'r', encoding='utf-8') as f:
            languages = json.load(f)
        return jsonify(languages)
    except FileNotFoundError:
        # Fallback nếu file không tồn tại
        return jsonify([
            {"code": "ja", "name": "tiếng Nhật",  "label": "🇯🇵 Tiếng Nhật (Japanese)"},
            {"code": "en", "name": "tiếng Anh",   "label": "🇺🇸 Tiếng Anh (English)"},
            {"code": "vi", "name": "tiếng Việt",  "label": "🇻🇳 Tiếng Việt (Vietnamese)"},
            {"code": "zh", "name": "tiếng Trung", "label": "🇨🇳 Tiếng Trung (Chinese)"},
            {"code": "ko", "name": "tiếng Hàn",   "label": "🇰🇷 Tiếng Hàn (Korean)"}
        ])
    except Exception as e:
        return jsonify({'error': str(e)}), 500

# ==================== API: PROMPT TEMPLATES ====================

@app.route('/api/templates', methods=['GET'])
@login_required
def api_get_templates():
    """Trả về danh sách prompt templates cho ngôn ngữ được chỉ định"""
    lang = request.args.get('lang', 'default')
    return jsonify(load_templates(lang))

@app.route('/api/templates', methods=['POST'])
@login_required
def api_save_templates():
    """Lưu danh sách prompt templates cho ngôn ngữ được chỉ định"""
    lang = request.args.get('lang', 'default')
    new_templates = request.json
    if not isinstance(new_templates, list):
        return jsonify({'error': 'Dữ liệu phải là array'}), 400
    try:
        try:
            with open(TEMPLATES_FILE, 'r', encoding='utf-8') as f:
                all_data = json.load(f)
            if isinstance(all_data, list):
                all_data = {'default': all_data}
        except Exception:
            all_data = {}
        all_data[lang] = new_templates
        with open(TEMPLATES_FILE, 'w', encoding='utf-8') as f:
            json.dump(all_data, f, ensure_ascii=False, indent=2)
        return jsonify({'success': True})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/extract', methods=['POST'])
@login_required
def extract():
    """
    Chức năng 1: Trích xuất các cell chứa string từ file Excel, PPTX hoặc DOCX
    Bỏ qua các cell chứa số và công thức (bắt đầu bằng '=') trong Excel
    Trả về file JSON với format: {"SheetName!CellCoordinate": "Content"} hoặc {"SlideX!ShapeY": "Content"} hoặc {"ParagraphX": "Content"}
    """
    # Kiểm tra xem có file được upload không
    if 'file' not in request.files:
        return jsonify({'error': 'Không có file được upload'}), 400
    
    file = request.files['file']
    
    # Kiểm tra xem file có được chọn không
    if file.filename == '':
        return jsonify({'error': 'Không có file được chọn'}), 400
    
    # Kiểm tra định dạng file
    if not allowed_file(file.filename):
        return jsonify({'error': 'Chỉ chấp nhận file .xlsx, .pptx hoặc .docx'}), 400
    
    try:
        # Lấy session folder
        session_folder = get_session_folder()
        
        # Lưu tên file gốc (giữ nguyên tiếng Nhật, ký tự đặc biệt)
        original_filename = file.filename
        
        # Lấy extension từ tên file gốc
        if '.' in original_filename:
            original_ext = original_filename.rsplit('.', 1)[1].lower()
        else:
            return jsonify({'error': 'Tên file phải có đuôi mở rộng (.xlsx hoặc .pptx)'}), 400
        
        # Tạo tên file tạm an toàn hoàn toàn từ timestamp (không dùng tên gốc)
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        safe_temp_filename = f"temp_{timestamp}.{original_ext}"
        filepath = os.path.join(session_folder, safe_temp_filename)
        file.save(filepath)
        
        # Xác định loại file và trích xuất
        file_ext = original_ext
        
        if file_ext == 'xlsx':
            # Mở file Excel bằng openpyxl
            workbook = load_workbook(filepath)
            
            # Dictionary để lưu kết quả
            extracted_data = {}
            
            # Duyệt qua tất cả các sheet
            for sheet_name in workbook.sheetnames:
                sheet = workbook[sheet_name]
                
                # Duyệt qua tất cả các cell trong sheet
                for row in sheet.iter_rows():
                    for cell in row:
                        # Bỏ qua cell rỗng
                        if cell.value is None:
                            continue
                        
                        # Chỉ lấy cell chứa string
                        if isinstance(cell.value, str):
                            # Bỏ qua công thức (bắt đầu bằng '=')
                            if not cell.value.startswith('='):
                                # Tạo key theo format "SheetName!CellCoordinate"
                                key = f"{sheet_name}!{cell.coordinate}"
                                extracted_data[key] = cell.value
            
            # Đóng workbook
            workbook.close()
        
        elif file_ext == 'pptx':
            # Trích xuất text từ PPTX
            extracted_data = extract_text_from_pptx(filepath)
        
        elif file_ext == 'docx':
            # Trích xuất text từ DOCX
            extracted_data = extract_text_from_docx(filepath)
        
        # Tách dữ liệu thành nhiều file, mỗi file 400 cặp key-value
        CHUNK_SIZE = 400
        data_items = list(extracted_data.items())
        total_items = len(data_items)
        num_files = (total_items + CHUNK_SIZE - 1) // CHUNK_SIZE  # Làm tròn lên
        
        # Lấy tên file gốc không có extension (giữ nguyên tiếng Nhật)
        base_filename = os.path.splitext(original_filename)[0]
        
        # Nếu base_filename rỗng, dùng tên mặc định
        if not base_filename or base_filename.strip() == '':
            base_filename = f"file_{timestamp}"
        # Tên safe cho filesystem (dùng timestamp)
        safe_base_filename = f"extracted_{timestamp}"
        
        # Tên folder trong ZIP (giữ nguyên tiếng Nhật)
        folder_name = f"{base_filename}_json_to_translate"
        # Tên folder tạm trong filesystem (dùng safe filename)
        safe_folder_name = f"{safe_base_filename}_temp_{timestamp}"
        
        # Tạo thư mục tạm để chứa các file JSON (dùng tên safe cho filesystem)
        temp_dir = os.path.join(session_folder, safe_folder_name)
        os.makedirs(temp_dir, exist_ok=True)
        
        json_files = []
        json_display_names = []  # Lưu tên hiển thị với tiếng Nhật
        
        # Tạo các file JSON nhỏ
        for i in range(num_files):
            start_idx = i * CHUNK_SIZE
            end_idx = min((i + 1) * CHUNK_SIZE, total_items)
            chunk_data = dict(data_items[start_idx:end_idx])
            
            # Tên file hiển thị (giữ nguyên tiếng Nhật)
            json_display_name = f"{base_filename}_part{i+1:02d}_of_{num_files:02d}.json"
            json_display_names.append(json_display_name)
            
            # Tên file an toàn cho filesystem
            safe_json_filename = f"{safe_base_filename}_part{i+1:02d}.json"
            json_filepath = os.path.join(temp_dir, safe_json_filename)
            
            # Lưu dữ liệu vào file JSON với encoding UTF-8
            with open(json_filepath, 'w', encoding='utf-8') as json_file:
                json.dump(chunk_data, json_file, ensure_ascii=False, indent=2)
            
            json_files.append(json_filepath)
        
        # Đọc nội dung từng file JSON để trả về cho frontend
        files_data = []
        for idx, json_filepath in enumerate(json_files):
            with open(json_filepath, 'r', encoding='utf-8') as f:
                files_data.append({
                    'name': json_display_names[idx],
                    'content': f.read()
                })
        
        # Tạo file ZIP chứa folder và các file JSON
        zip_display_name = f"{base_filename}_json_to_translate.zip"  # Tên hiển thị
        safe_zip_filename = f"{safe_base_filename}_json_{timestamp}.zip"  # Tên file trong filesystem
        zip_filepath = os.path.join(session_folder, safe_zip_filename)
        
        # Dùng ZIP_STORED để không nén file JSON
        with zipfile.ZipFile(zip_filepath, 'w', zipfile.ZIP_STORED) as zipf:
            for idx, json_filepath_item in enumerate(json_files):
                arcname = os.path.join(folder_name, json_display_names[idx])
                zipf.write(json_filepath_item, arcname)
        
        # Lưu thông tin ZIP vào session để download sau
        session['extract_zip'] = {
            'path': zip_filepath,
            'display_name': zip_display_name,
            'input_path': filepath,
            'json_files': json_files,
            'temp_dir': temp_dir
        }
        
        # Tính toán dedup data (gộp keys có cùng value)
        dedup_files, dedup_mapping, dedup_stats = build_dedup_data(extracted_data, CHUNK_SIZE)

        # Lưu dedup mapping vào session folder để dùng khi inject
        dedup_mapping_path = os.path.join(session_folder, 'dedup_mapping.json')
        with open(dedup_mapping_path, 'w', encoding='utf-8') as f:
            json.dump(dedup_mapping, f, ensure_ascii=False, indent=2)

        # Trả về JSON response với danh sách file để frontend hiển thị
        return jsonify({
            'success': True,
            'total_files': num_files,
            'total_items': total_items,
            'files': files_data,
            'zip_display_name': zip_display_name,
            'dedup_files': dedup_files,
            'dedup_stats': dedup_stats
        })
        
    except Exception as e:
        # Xử lý lỗi
        return jsonify({'error': f'Lỗi khi xử lý file: {str(e)}'}), 500

@app.route('/download-zip', methods=['GET'])
@login_required
def download_zip():
    """
    Serve file ZIP đã được tạo từ /extract.
    Xóa tất cả file tạm sau khi gửi xong.
    """
    zip_info = session.get('extract_zip')
    if not zip_info:
        return jsonify({'error': 'Không tìm thấy file ZIP. Vui lòng trích xuất lại.'}), 404
    
    zip_filepath = zip_info.get('path')
    zip_display_name = zip_info.get('display_name', 'download.zip')
    
    if not zip_filepath or not os.path.exists(zip_filepath):
        return jsonify({'error': 'File ZIP không còn tồn tại. Vui lòng trích xuất lại.'}), 404
    
    # Xóa thông tin ZIP trong session
    session.pop('extract_zip', None)
    
    # Trả về file ZIP
    response = send_file(zip_filepath, mimetype='application/zip')
    response = set_download_headers(response, zip_display_name, 'download.zip')
    
    # Xóa tất cả file tạm sau khi gửi
    input_path = zip_info.get('input_path')
    json_files = zip_info.get('json_files', [])
    temp_dir = zip_info.get('temp_dir')
    
    @response.call_on_close
    def cleanup():
        import time
        import gc
        gc.collect()
        time.sleep(0.1)
        
        try:
            if zip_filepath and os.path.exists(zip_filepath):
                os.remove(zip_filepath)
        except Exception as e:
            print(f"Warning: Không thể xóa ZIP: {e}")
        
        try:
            if input_path and os.path.exists(input_path):
                os.remove(input_path)
        except Exception as e:
            print(f"Warning: Không thể xóa input file: {e}")
        
        for jf in json_files:
            try:
                if os.path.exists(jf):
                    os.remove(jf)
            except Exception as e:
                print(f"Warning: Không thể xóa JSON file: {e}")
        
        try:
            if temp_dir and os.path.exists(temp_dir):
                os.rmdir(temp_dir)
        except Exception as e:
            print(f"Warning: Không thể xóa temp dir: {e}")
    
    return response

@app.route('/inject', methods=['POST'])
@login_required
def inject():
    """
    Chức năng 2: Nạp dữ liệu từ file JSON đã dịch vào file Excel, PPTX hoặc DOCX gốc
    Giữ nguyên định dạng, màu sắc của file gốc
    Hỗ trợ nhiều file JSON riêng lẻ hoặc file ZIP chứa nhiều file JSON
    """
    # Kiểm tra xem có file được upload không
    if 'excel_file' not in request.files:
        return jsonify({'error': 'Cần upload file Excel, PPTX hoặc DOCX'}), 400
    
    excel_file = request.files['excel_file']
    
    # Lấy pasted JSON data nếu có
    pasted_json_data = request.form.get('pasted_json_data', None)
    
    # Kiểm tra xem có file JSON được upload hoặc có pasted JSON không
    json_files = request.files.getlist('json_files') if 'json_files' in request.files else []
    
    # Kiểm tra xem có ít nhất một nguồn JSON
    has_json_files = len(json_files) > 0 and any(f.filename != '' for f in json_files)
    has_pasted_json = pasted_json_data is not None and pasted_json_data.strip() != ''
    
    if not has_json_files and not has_pasted_json:
        return jsonify({'error': 'Cần upload ít nhất 1 file JSON/ZIP hoặc paste JSON'}), 400
    
    # Kiểm tra xem file excel có được chọn không
    if excel_file.filename == '':
        return jsonify({'error': 'Cần chọn file Excel/PowerPoint/Word'}), 400
    
    # Kiểm tra định dạng file
    if not allowed_file(excel_file.filename):
        return jsonify({'error': 'File phải có định dạng .xlsx, .pptx hoặc .docx'}), 400
    
    try:
        # Lấy session folder
        session_folder = get_session_folder()
        
        # Lưu tên file gốc (giữ nguyên tiếng Nhật, ký tự đặc biệt)
        original_excel_filename = excel_file.filename
        
        # Lấy extension từ tên file gốc
        if '.' in original_excel_filename:
            original_ext = original_excel_filename.rsplit('.', 1)[1].lower()
        else:
            return jsonify({'error': 'Tên file phải có đuôi mở rộng (.xlsx hoặc .pptx)'}), 400
        
        # Tạo tên file tạm an toàn hoàn toàn từ timestamp
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        safe_temp_filename = f"temp_{timestamp}.{original_ext}"
        excel_filepath = os.path.join(session_folder, safe_temp_filename)
        excel_file.save(excel_filepath)
        
        # Đọc và gộp dữ liệu JSON từ tất cả các file
        json_data = {}
        temp_files = []
        
        for json_file in json_files:
            if json_file.filename == '':
                continue
                
            json_filename = json_file.filename.lower()
            is_zip = json_filename.endswith('.zip')
            is_json = json_filename.endswith('.json')
            
            if not (is_zip or is_json):
                continue
            
            if is_zip:
                # Xử lý file ZIP
                temp_zip_filename = f"temp_{timestamp}_{secure_filename(json_file.filename)}"
                zip_filepath = os.path.join(session_folder, temp_zip_filename)
                json_file.save(zip_filepath)
                temp_files.append(zip_filepath)
                
                # Giải nén và đọc tất cả các file JSON
                with zipfile.ZipFile(zip_filepath, 'r') as zipf:
                    for file_info in zipf.namelist():
                        if file_info.lower().endswith('.json'):
                            with zipf.open(file_info) as f:
                                try:
                                    content = f.read().decode('utf-8')
                                    chunk_data = json.loads(content)
                                    json_data.update(chunk_data)
                                except UnicodeDecodeError:
                                    # Thử với encoding khác
                                    f.seek(0)
                                    content = f.read().decode('utf-8-sig')
                                    chunk_data = json.loads(content)
                                    json_data.update(chunk_data)
            else:
                # Xử lý file JSON đơn lẻ
                try:
                    json_content = json_file.stream.read().decode('utf-8')
                    chunk_data = json.loads(json_content)
                    json_data.update(chunk_data)
                except UnicodeDecodeError:
                    # Thử với encoding khác nếu UTF-8 thất bại
                    json_file.stream.seek(0)
                    json_content = json_file.stream.read().decode('utf-8-sig')
                    chunk_data = json.loads(json_content)
                    json_data.update(chunk_data)
                except json.JSONDecodeError as e:
                    return jsonify({'error': f'File JSON "{json_file.filename}" không hợp lệ: {str(e)}'}), 400
        
        # Xử lý pasted JSON data
        if pasted_json_data:
            try:
                pasted_data_list = json.loads(pasted_json_data)
                
                # pasted_data_list là danh sách các JSON objects
                if isinstance(pasted_data_list, list):
                    for idx, pasted_obj in enumerate(pasted_data_list):
                        if isinstance(pasted_obj, dict):
                            json_data.update(pasted_obj)
                        else:
                            return jsonify({'error': f'Pasted JSON #{idx + 1} không phải là object'}), 400
                else:
                    return jsonify({'error': 'Pasted JSON data phải là danh sách các objects'}), 400
                    
            except json.JSONDecodeError as e:
                return jsonify({'error': f'Pasted JSON không hợp lệ: {str(e)}'}), 400

        # Mở rộng dedup keys nếu có (khi user dịch từ dedup JSON)
        if any(k.startswith('dedup_') for k in json_data):
            json_data = expand_dedup_data(json_data, session_folder)

        # Xác định loại file và nạp dữ liệu (dùng tên file gốc)
        file_ext = original_excel_filename.rsplit('.', 1)[1].lower()
        
        if file_ext == 'xlsx':
            # Mở file Excel bằng openpyxl
            workbook = load_workbook(excel_filepath)
            
            # Duyệt qua từng entry trong JSON
            for key, translated_value in json_data.items():
                # Parse key theo format "SheetName!CellCoordinate"
                if '!' not in key:
                    continue
                
                sheet_name, cell_coordinate = key.split('!', 1)
                
                # Kiểm tra xem sheet có tồn tại không
                if sheet_name not in workbook.sheetnames:
                    continue
                
                # Lấy sheet
                sheet = workbook[sheet_name]
                
                # Nạp dữ liệu đã dịch vào cell
                # openpyxl tự động giữ nguyên định dạng của cell
                sheet[cell_coordinate] = translated_value
            
            # Tạo tên file output
            base_filename = os.path.splitext(original_excel_filename)[0]  # Tên gốc với tiếng Nhật
            
            output_display_name = f"{base_filename}_translated.xlsx"  # Tên hiển thị
            safe_output_filename = f"output_{timestamp}.xlsx"  # Tên file trong filesystem
            output_filepath = os.path.join(session_folder, safe_output_filename)
            
            # Lưu file Excel đã được nạp dữ liệu
            workbook.save(output_filepath)
            workbook.close()
            del workbook  # Giải phóng memory
            
            output_mimetype = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        
        elif file_ext == 'pptx':
            # Nạp text vào PPTX
            prs = inject_text_to_pptx(excel_filepath, json_data)
            
            # Tạo tên file output
            base_filename = os.path.splitext(original_excel_filename)[0]  # Tên gốc với tiếng Nhật
            
            output_display_name = f"{base_filename}_translated.pptx"  # Tên hiển thị
            safe_output_filename = f"output_{timestamp}.pptx"  # Tên file trong filesystem
            output_filepath = os.path.join(session_folder, safe_output_filename)
            
            # Lưu file PPTX đã được nạp dữ liệu
            prs.save(output_filepath)
            del prs  # Giải phóng memory
            
            output_mimetype = 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
        
        elif file_ext == 'docx':
            # Nạp text vào DOCX
            doc = inject_text_to_docx(excel_filepath, json_data)
            
            # Tạo tên file output
            base_filename = os.path.splitext(original_excel_filename)[0]  # Tên gốc với tiếng Nhật
            
            output_display_name = f"{base_filename}_translated.docx"  # Tên hiển thị
            safe_output_filename = f"output_{timestamp}.docx"  # Tên file trong filesystem
            output_filepath = os.path.join(session_folder, safe_output_filename)
            
            # Lưu file DOCX đã được nạp dữ liệu
            doc.save(output_filepath)
            del doc  # Giải phóng memory
            
            output_mimetype = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        
        # Trả về file đã được nạp dữ liệu (dùng tên hiển thị)
        response = send_file(
            output_filepath,
            mimetype=output_mimetype
        )
        
        default_ascii_name = 'download.docx' if file_ext == 'docx' else ('download.pptx' if file_ext == 'pptx' else 'download.xlsx')
        response = set_download_headers(response, output_display_name, default_ascii_name)
        
        # Xóa tất cả file tạm sau khi gửi response
        @response.call_on_close
        def cleanup():
            import time
            import gc
            
            # Force garbage collection để giải phóng file handles
            gc.collect()
            time.sleep(0.1)  # Delay nhỏ để đảm bảo file được giải phóng
            
            # Xóa file output
            try:
                if os.path.exists(output_filepath):
                    os.remove(output_filepath)
            except Exception as e:
                print(f"Warning: Không thể xóa output file: {e}")
            
            # Xóa file Excel/PPTX/DOCX tạm
            try:
                if os.path.exists(excel_filepath):
                    os.remove(excel_filepath)
            except Exception as e:
                print(f"Warning: Không thể xóa file tạm: {e}")
            
            # Xóa tất cả file ZIP tạm
            for temp_file in temp_files:
                try:
                    if os.path.exists(temp_file):
                        os.remove(temp_file)
                except Exception as e:
                    print(f"Warning: Không thể xóa file ZIP tạm: {e}")
        
        return response
        
    except Exception as e:
        # Xử lý lỗi
        return jsonify({'error': f'Lỗi khi xử lý file: {str(e)}'}), 500

@app.route('/clear-uploads', methods=['POST'])
@login_required
def clear_uploads():
    """
    Xóa tất cả file trong thư mục session hiện tại
    """
    try:
        session_folder = get_session_folder()
        
        # Kiểm tra xem thư mục có tồn tại không
        if not os.path.exists(session_folder):
            return jsonify({'success': True, 'message': 'Không có file nào để xóa'}), 200
        
        # Đếm số file đã xóa
        deleted_count = 0
        
        # Duyệt qua tất cả file trong session folder
        for item in os.listdir(session_folder):
            item_path = os.path.join(session_folder, item)
            
            try:
                if os.path.isfile(item_path):
                    # Xóa file
                    os.remove(item_path)
                    deleted_count += 1
                elif os.path.isdir(item_path):
                    # Xóa thư mục con và tất cả nội dung bên trong
                    shutil.rmtree(item_path)
                    deleted_count += 1
            except Exception as e:
                print(f"Không thể xóa {item_path}: {str(e)}")
        
        return jsonify({
            'success': True,
            'message': f'Đã xóa thành công {deleted_count} file trong phiên của bạn',
            'deleted_count': deleted_count
        }), 200
        
    except Exception as e:
        return jsonify({'error': f'Lỗi khi xóa file: {str(e)}'}), 500

# Đảm bảo thư mục uploads tồn tại
if not os.path.exists(app.config['UPLOAD_FOLDER']):
    os.makedirs(app.config['UPLOAD_FOLDER'])

if __name__ == '__main__':
    # Chạy ứng dụng Flask ở chế độ debug
    app.run(debug=True, host='0.0.0.0', port=5001)
