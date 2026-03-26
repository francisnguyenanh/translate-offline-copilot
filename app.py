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
import re
import csv
from datetime import datetime, timedelta
from urllib.parse import quote
from flask import Flask, render_template, request, send_file, jsonify, session, redirect, url_for
from werkzeug.utils import secure_filename
from copy import deepcopy, copy
from openpyxl import load_workbook
from functools import wraps

# Khởi tạo ứng dụng Flask
app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 60 * 1024 * 1024  # Giới hạn 50MB

# DEBUG: enable debug-level logging for OCR coordinate diagnostics
import logging
app.logger.setLevel(logging.DEBUG)
logging.basicConfig(level=logging.DEBUG)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['SECRET_KEY'] = os.urandom(24)  # Secret key cho session
app.config['PERMANENT_SESSION_LIFETIME'] = timedelta(hours=5)  # Session timeout 5h

# Các định dạng file được phép
ALLOWED_EXTENSIONS = {'xlsx'}

# Đọc password từ file
PASSWORD_FILE = 'password.txt'


# File lưu Prompt Templates (Tab 1/3)
TEMPLATES_FILE = 'prompt_templates.json'
GLOSSARY_DIR = 'glossaries'   # thư mục lưu các file CSV chuyên ngành
os.makedirs(GLOSSARY_DIR, exist_ok=True)

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


def apply_glossary(extracted_data: dict, glossary_ids: list) -> dict:
    """
    Thay thế các cụm từ trong extracted_data theo các file glossary được chọn.

    Logic:
    - Với mỗi cặp (src, dst) trong glossary: tìm src (case-insensitive, exact match)
      trong từng value của extracted_data, thay bằng dst.
    - "Exact match" = cụm từ đứng độc lập, không phải substring nằm giữa ký tự chữ.
      Dùng regex word-boundary \\b kết hợp re.IGNORECASE.
    - Ưu tiên thay thế cụm dài trước (tránh thay một phần của cụm dài hơn).
    - Nếu glossary_ids rỗng hoặc không có file nào, trả về nguyên extracted_data.

    CSV format: cột A = ngôn ngữ đích (dst), cột B = ngôn ngữ gốc (src)
    """
    if not glossary_ids:
        return extracted_data

    # Thu thập tất cả cặp (src, dst) từ các file glossary được chọn
    pairs = []
    for gid in glossary_ids:
        filepath = os.path.join(GLOSSARY_DIR, f'{gid}.csv')
        if not os.path.exists(filepath):
            continue
        with open(filepath, 'r', encoding='utf-8-sig') as f:
            reader = csv.reader(f)
            for row in reader:
                if len(row) >= 2:
                    dst = row[0].strip()   # cột A: ngôn ngữ đích
                    src = row[1].strip()   # cột B: ngôn ngữ gốc
                    if src and dst:
                        pairs.append((src, dst))

    if not pairs:
        return extracted_data

    # Sắp xếp: cụm dài trước để tránh thay nhầm phần của cụm dài
    pairs.sort(key=lambda x: len(x[0]), reverse=True)

    # Build regex patterns (một lần, tái sử dụng)
    # Dùng \b nếu src bắt đầu/kết thúc bằng ký tự word; fallback lookaround nếu không
    def make_pattern(src: str):
        escaped = re.escape(src)
        prefix = r'\b' if re.match(r'\w', src[0])  else r'(?<![^\s])'
        suffix = r'\b' if re.match(r'\w', src[-1]) else r'(?![^\s])'
        return re.compile(prefix + escaped + suffix, re.IGNORECASE)

    compiled = [(make_pattern(src), dst) for src, dst in pairs]

    # Áp dụng thay thế lên từng value
    result = {}
    for key, value in extracted_data.items():
        if not isinstance(value, str):
            result[key] = value
            continue
        for pattern, dst in compiled:
            value = pattern.sub(dst, value)
        result[key] = value

    return result


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


def _get_rgb6(cell) -> str:
    """Đọc mã RGB 6 ký tự từ font.color, trả rỗng nếu không xác định được."""
    try:
        color = cell.font.color if (cell.font and cell.font.color) else None
        if color and color.type == 'rgb' and color.rgb:
            return str(color.rgb).upper().lstrip('F')[-6:].zfill(6)
    except Exception:
        pass
    return ''


# ==================== COLOR FILTER HELPERS ====================

def _get_font_rgb_xlsx(cell) -> str:
    """Đọc màu chữ HEX 6 ký tự từ openpyxl cell. Trả '' nếu không xác định."""
    try:
        color = cell.font.color if (cell.font and cell.font.color) else None
        if color and color.type == 'rgb' and color.rgb:
            return str(color.rgb).upper()[-6:]  # lấy 6 ký tự cuối của ARGB
    except Exception:
        pass
    return ''


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

def _xlsx_get_sheet_drawing_map(z):
    """
    Từ ZipFile đang mở, trả về dict: sheet_name → list các đường dẫn drawing XML
    Ví dụ: {'Sheet1': ['xl/drawings/drawing1.xml']}
    """
    names_set = set(z.namelist())

    # Đọc workbook.xml để lấy tên sheet và rId
    wb_xml = z.read('xl/workbook.xml')
    wb_root = _etree.fromstring(wb_xml)

    # Đọc rels của workbook để map rId → target file
    wb_rels_xml = z.read('xl/_rels/workbook.xml.rels')
    wb_rels_root = _etree.fromstring(wb_rels_xml)
    rid_to_target = {rel.get('Id'): rel.get('Target') for rel in wb_rels_root}

    # Thu thập (sheet_name, sheet_path)
    sheet_info = []
    for sheet_el in wb_root.iter(f'{{{_NS_WB}}}sheet'):
        name = sheet_el.get('name')
        rid  = sheet_el.get(f'{{{_NS_R}}}id')
        target = rid_to_target.get(rid, '')
        # Chuẩn hóa path: "worksheets/sheet1.xml" → "xl/worksheets/sheet1.xml"
        if target.startswith('../'):
            sheet_path = 'xl/' + target[3:]
        elif not target.startswith('xl/'):
            sheet_path = 'xl/' + target
        else:
            sheet_path = target
        sheet_info.append((name, sheet_path))

    result = {}
    for sheet_name, sheet_path in sheet_info:
        if sheet_path not in names_set:
            continue
        parts = sheet_path.rsplit('/', 1)
        sheet_dir  = parts[0] if len(parts) > 1 else ''
        sheet_file = parts[-1]
        rels_path  = f"{sheet_dir}/_rels/{sheet_file}.rels"
        if rels_path not in names_set:
            continue

        rels_xml  = z.read(rels_path)
        rels_root = _etree.fromstring(rels_xml)

        drawing_paths = []
        for rel in rels_root:
            if 'drawing' in rel.get('Type', '').lower():
                tgt = rel.get('Target', '')
                if tgt.startswith('../'):
                    draw_path = 'xl/' + tgt[3:]
                elif tgt.startswith('/'):
                    draw_path = tgt.lstrip('/')
                else:
                    draw_path = f"{sheet_dir}/{tgt.lstrip('./')}"
                if draw_path in names_set:
                    drawing_paths.append(draw_path)
        if drawing_paths:
            result[sheet_name] = drawing_paths

    return result


def _collect_sp_elements(drawing_root):
    """
    Thu thập tất cả phần tử <xdr:sp> (shape/text-box) theo thứ tự cây (tree-order),
    bao gồm cả sp bên trong group-shapes (xdr:grpSp).
    """
    sp_list = []
    _walk_tags = {
        f'{{{_NS_XDR}}}grpSp',
        f'{{{_NS_XDR}}}twoCellAnchor',
        f'{{{_NS_XDR}}}oneCellAnchor',
        f'{{{_NS_XDR}}}absoluteAnchor',
    }

    def walk(el):
        for child in el:
            if child.tag == f'{{{_NS_XDR}}}sp':
                sp_list.append(child)
            elif child.tag in _walk_tags:
                walk(child)

    walk(drawing_root)
    return sp_list


def _get_sp_text(sp):
    """Lấy toàn bộ text trong txBody của một shape."""
    txBody = sp.find(f'{{{_NS_XDR}}}txBody')
    if txBody is None:
        return ''
    parts = []
    for para in txBody.findall(f'{{{_NS_A}}}p'):
        # Lấy tất cả a:t trong paragraph (bao gồm cả text trong a:fld)
        para_text = ''.join((t.text or '') for t in para.findall(f'.//{{{_NS_A}}}t'))
        if para_text:
            parts.append(para_text)
    return '\n'.join(parts)


def _set_sp_text(sp, new_text):
    """
    Thay thế text trong txBody của shape theo cách tối thiểu để tương thích Excel:
    - Giữ nguyên cấu trúc paragraph/run/fld hiện có
    - Chỉ thay text ở node a:t đầu tiên, các node a:t còn lại set rỗng
    - Nếu chưa có a:t thì tạo mới tối thiểu trong paragraph đầu
    """
    txBody = sp.find(f'{{{_NS_XDR}}}txBody')
    if txBody is None:
        return

    paras = txBody.findall(f'{{{_NS_A}}}p')
    if not paras:
        return

    all_t_nodes = txBody.findall(f'.//{{{_NS_A}}}t')
    new_text = '' if new_text is None else str(new_text)

    if all_t_nodes:
        first_t = all_t_nodes[0]
        first_t.text = new_text
        if new_text != new_text.strip() or '\n' in new_text:
            first_t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
        else:
            first_t.attrib.pop('{http://www.w3.org/XML/1998/namespace}space', None)

        for t in all_t_nodes[1:]:
            t.text = ''
        return

    # Không có a:t nào, tạo tối thiểu trong paragraph đầu
    first_para = paras[0]
    first_run = _etree.SubElement(first_para, f'{{{_NS_A}}}r')
    first_t = _etree.SubElement(first_run, f'{{{_NS_A}}}t')
    first_t.text = new_text
    if new_text != new_text.strip() or '\n' in new_text:
        first_t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')


def extract_xlsx_shapes(filepath):
    """
    Trích xuất text từ tất cả shape/object (text-box) trong file xlsx.
    Trả về dict: {"SheetName!XLShape{n}": "text"}
    - n là thứ tự shape (đếm TẤT CẢ sp, bao gồm cả sp không có text) → index ổn định.
    """
    shapes_data = {}
    try:
        with zipfile.ZipFile(filepath, 'r') as z:
            sheet_drawing_map = _xlsx_get_sheet_drawing_map(z)
            for sheet_name, drawing_paths in sheet_drawing_map.items():
                for drawing_path in drawing_paths:
                    drawing_xml  = z.read(drawing_path)
                    drawing_root = _etree.fromstring(drawing_xml)
                    for shape_idx, sp in enumerate(_collect_sp_elements(drawing_root), start=1):
                        text = _get_sp_text(sp)
                        if text.strip():
                            key = f"{sheet_name}!XLShape{shape_idx}"
                            shapes_data[key] = text.strip()
    except Exception as e:
        print(f"Warning: Không thể trích xuất shapes từ xlsx: {e}")
    return shapes_data


def _xlsx_sheet_path_map(files):
    """Trả về map: sheet_name -> sheet_xml_path từ workbook + workbook.rels."""
    workbook_root = _etree.fromstring(files['xl/workbook.xml'])
    rels_root = _etree.fromstring(files['xl/_rels/workbook.xml.rels'])
    rid_to_target = {rel.get('Id'): rel.get('Target') for rel in rels_root}

    sheet_map = {}
    for sheet_el in workbook_root.iter(f'{{{_NS_WB}}}sheet'):
        sheet_name = sheet_el.get('name')
        rid = sheet_el.get(f'{{{_NS_R}}}id')
        target = rid_to_target.get(rid, '')
        if target.startswith('/'):
            path = target.lstrip('/')
        elif target.startswith('xl/'):
            path = target
        else:
            path = f'xl/{target}'
        sheet_map[sheet_name] = path
    return sheet_map


def _xlsx_sheet_drawing_map_from_files(files, sheet_map):
    """Trả về map: sheet_name -> [drawing_xml_paths]."""
    result = {}
    names_set = set(files.keys())

    for sheet_name, sheet_path in sheet_map.items():
        parts = sheet_path.rsplit('/', 1)
        sheet_dir = parts[0] if len(parts) > 1 else ''
        sheet_file = parts[-1]
        rels_path = f"{sheet_dir}/_rels/{sheet_file}.rels"
        if rels_path not in names_set:
            continue

        rels_root = _etree.fromstring(files[rels_path])
        drawing_paths = []
        for rel in rels_root:
            if 'drawing' in (rel.get('Type') or '').lower():
                tgt = rel.get('Target', '')
                if tgt.startswith('/'):
                    draw_path = tgt.lstrip('/')
                else:
                    draw_path = os.path.normpath(os.path.join(sheet_dir, tgt)).replace('\\', '/')
                if draw_path in names_set:
                    drawing_paths.append(draw_path)

        if drawing_paths:
            result[sheet_name] = drawing_paths

    return result


def _xlsx_set_cell_inline_text(sheet_root, cell_ref, text_value):
    """Gán text vào một cell theo kiểu inlineStr, giữ nguyên style của cell nếu có."""
    m = re.fullmatch(r'([A-Za-z]+)(\d+)', cell_ref or '')
    if not m:
        return

    row_num = int(m.group(2))
    sheet_data = sheet_root.find(f'{{{_NS_WB}}}sheetData')
    if sheet_data is None:
        return

    row_elem = None
    rows = sheet_data.findall(f'{{{_NS_WB}}}row')
    for row in rows:
        if int(row.get('r', '0') or 0) == row_num:
            row_elem = row
            break

    if row_elem is None:
        row_elem = _etree.Element(f'{{{_NS_WB}}}row', r=str(row_num))
        inserted = False
        for idx, row in enumerate(rows):
            existing = int(row.get('r', '0') or 0)
            if existing > row_num:
                sheet_data.insert(idx, row_elem)
                inserted = True
                break
        if not inserted:
            sheet_data.append(row_elem)

    cell_elem = None
    for cell in row_elem.findall(f'{{{_NS_WB}}}c'):
        if (cell.get('r') or '').upper() == cell_ref.upper():
            cell_elem = cell
            break

    if cell_elem is None:
        cell_elem = _etree.Element(f'{{{_NS_WB}}}c', r=cell_ref.upper())
        row_elem.append(cell_elem)

    text_str = '' if text_value is None else str(text_value)

    # Check if cell already has an <is> with rich-text runs
    existing_is = cell_elem.find(f'{{{_NS_WB}}}is')
    r_tag = f'{{{_NS_WB}}}r'
    t_tag = f'{{{_NS_WB}}}t'
    xml_space_attr = '{http://www.w3.org/XML/1998/namespace}space'

    if existing_is is not None:
        runs = existing_is.findall(r_tag)
        if len(runs) > 1:
            # Rich text: distribute new_text across runs by original char-count ratio
            orig_lengths = [len((r.findtext(t_tag) or '')) for r in runs]
            total_orig = sum(orig_lengths)
            new_total = len(text_str)
            pos = 0
            for i, (run_elem, orig_len) in enumerate(zip(runs, orig_lengths)):
                t_elem = run_elem.find(t_tag)
                if t_elem is None:
                    t_elem = _etree.SubElement(run_elem, t_tag)
                if i == len(runs) - 1:
                    part = text_str[pos:]
                else:
                    if total_orig > 0:
                        count = round(new_total * orig_len / total_orig)
                        count = min(count, new_total - pos - (len(runs) - i - 1))
                        count = max(count, 0)
                    else:
                        count = 0
                    part = text_str[pos:pos + count]
                    pos += count
                t_elem.text = part
                if part != part.strip() or '\n' in part:
                    t_elem.set(xml_space_attr, 'preserve')
                elif xml_space_attr in t_elem.attrib:
                    del t_elem.attrib[xml_space_attr]
            # Ensure cell type is inlineStr
            cell_elem.set('t', 'inlineStr')
            for child in list(cell_elem):
                if child.tag in {f'{{{_NS_WB}}}v', f'{{{_NS_WB}}}f'}:
                    cell_elem.remove(child)
            return
        elif len(runs) == 1:
            # Single run: just replace the text element, keep rPr
            t_elem = runs[0].find(t_tag)
            if t_elem is None:
                t_elem = _etree.SubElement(runs[0], t_tag)
            t_elem.text = text_str
            if text_str != text_str.strip() or '\n' in text_str:
                t_elem.set(xml_space_attr, 'preserve')
            elif xml_space_attr in t_elem.attrib:
                del t_elem.attrib[xml_space_attr]
            cell_elem.set('t', 'inlineStr')
            for child in list(cell_elem):
                if child.tag in {f'{{{_NS_WB}}}v', f'{{{_NS_WB}}}f'}:
                    cell_elem.remove(child)
            return
        # else: plain <is><t>...</t></is> — fall through to rewrite below

    # No existing <is>, or plain <is><t> only: rewrite from scratch
    cell_elem.set('t', 'inlineStr')
    for child in list(cell_elem):
        if child.tag in {
            f'{{{_NS_WB}}}v',
            f'{{{_NS_WB}}}f',
            f'{{{_NS_WB}}}is',
        }:
            cell_elem.remove(child)

    is_elem = _etree.SubElement(cell_elem, f'{{{_NS_WB}}}is')
    t_elem = _etree.SubElement(is_elem, t_tag)
    t_elem.text = text_str
    if text_str != text_str.strip() or '\n' in text_str:
        t_elem.set(xml_space_attr, 'preserve')


def inject_xlsx_shapes(source_filepath, output_filepath, json_data):
    """
    ZIP-level patch cho XLSX:
    - Không dùng openpyxl.save
    - Patch trực tiếp cell XML
    - Patch trực tiếp drawing XML (shape text)
    - Giữ nguyên toàn bộ parts khác của file gốc
    """
    with zipfile.ZipFile(source_filepath, 'r') as z_src:
        infos = {info.filename: info for info in z_src.infolist()}
        order = [info.filename for info in z_src.infolist()]
        files = {name: z_src.read(name) for name in order}

    sheet_map = _xlsx_sheet_path_map(files)

    cell_updates = {}
    shape_updates = {}
    for key, translated_value in json_data.items():
        if '!' not in key:
            continue
        sheet_name, second_part = key.split('!', 1)
        if sheet_name not in sheet_map:
            continue

        if second_part.startswith('XLShape'):
            try:
                shape_idx = int(second_part.replace('XLShape', ''))
                shape_updates[(sheet_name, shape_idx)] = '' if translated_value is None else str(translated_value)
            except ValueError:
                continue
        else:
            cell_updates.setdefault(sheet_name, {})[second_part] = translated_value

    for sheet_name, updates in cell_updates.items():
        sheet_path = sheet_map.get(sheet_name)
        if not sheet_path or sheet_path not in files:
            continue
        sheet_root = _etree.fromstring(files[sheet_path])
        for cell_ref, value in updates.items():
            _xlsx_set_cell_inline_text(sheet_root, cell_ref, value)
        # Sync hyperlink display attributes to match updated cell values
        hyperlinks_elem = sheet_root.find(f'{{{_NS_WB}}}hyperlinks')
        if hyperlinks_elem is not None:
            for hl in hyperlinks_elem.findall(f'{{{_NS_WB}}}hyperlink'):
                ref = hl.get('ref', '')
                if ref in updates and hl.get('display') is not None:
                    hl.set('display', str(updates[ref]))
        files[sheet_path] = _etree.tostring(sheet_root, xml_declaration=True, encoding='UTF-8', standalone=True)

    if shape_updates:
        drawing_map = _xlsx_sheet_drawing_map_from_files(files, sheet_map)
        for sheet_name, drawing_paths in drawing_map.items():
            wanted = {
                idx: val
                for (sn, idx), val in shape_updates.items()
                if sn == sheet_name
            }
            if not wanted:
                continue

            for drawing_path in drawing_paths:
                if drawing_path not in files:
                    continue
                drawing_root = _etree.fromstring(files[drawing_path])
                for shape_idx, sp in enumerate(_collect_sp_elements(drawing_root), start=1):
                    if shape_idx in wanted:
                        _set_sp_text(sp, wanted[shape_idx])
                files[drawing_path] = _etree.tostring(
                    drawing_root,
                    xml_declaration=True,
                    encoding='UTF-8',
                    standalone=True,
                )

    with zipfile.ZipFile(output_filepath, 'w') as z_out:
        written = set()
        for name in order:
            content = files.get(name)
            if content is None:
                continue
            original = infos[name]
            zi = zipfile.ZipInfo(name, original.date_time)
            zi.compress_type = original.compress_type
            zi.external_attr = original.external_attr
            zi.create_system = original.create_system
            z_out.writestr(zi, content)
            written.add(name)

        for name, content in files.items():
            if name in written:
                continue
            z_out.writestr(name, content)


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


# ==================== API: GLOSSARY ====================

@app.route('/api/glossaries', methods=['GET'])
@login_required
def api_list_glossaries():
    """Trả về danh sách tất cả file glossary đã lưu."""
    result = []
    for fname in sorted(os.listdir(GLOSSARY_DIR)):
        if not fname.endswith('.csv'):
            continue
        gid = fname[:-4]
        meta_path = os.path.join(GLOSSARY_DIR, f'{gid}.meta.json')
        display_name = gid
        if os.path.exists(meta_path):
            with open(meta_path, 'r', encoding='utf-8') as f:
                meta = json.load(f)
                display_name = meta.get('name', gid)
        # Đếm số dòng
        try:
            with open(os.path.join(GLOSSARY_DIR, fname), 'r', encoding='utf-8-sig') as f:
                row_count = sum(1 for r in csv.reader(f) if any(r))
        except Exception:
            row_count = 0
        result.append({'id': gid, 'name': display_name, 'rows': row_count})
    return jsonify(result)


@app.route('/api/glossaries', methods=['POST'])
@login_required
def api_upload_glossary():
    """Upload file CSV mới. Form fields: file (CSV), name (tên hiển thị)."""
    import time
    if 'file' not in request.files:
        return jsonify({'error': 'Không có file'}), 400
    f = request.files['file']
    if not f.filename.endswith('.csv'):
        return jsonify({'error': 'Chỉ chấp nhận file .csv'}), 400
    display_name = request.form.get('name', '').strip() or os.path.splitext(f.filename)[0]
    gid = f'glossary_{int(time.time())}'
    csv_path  = os.path.join(GLOSSARY_DIR, f'{gid}.csv')
    meta_path = os.path.join(GLOSSARY_DIR, f'{gid}.meta.json')
    f.save(csv_path)
    with open(meta_path, 'w', encoding='utf-8') as mf:
        json.dump({'name': display_name}, mf, ensure_ascii=False)
    # Đếm dòng
    with open(csv_path, 'r', encoding='utf-8-sig') as cf:
        row_count = sum(1 for r in csv.reader(cf) if any(r))
    return jsonify({'success': True, 'id': gid, 'name': display_name, 'rows': row_count})


@app.route('/api/glossaries/<gid>', methods=['GET'])
@login_required
def api_get_glossary(gid):
    """Trả về toàn bộ nội dung glossary dưới dạng list of {src, dst}."""
    csv_path = os.path.join(GLOSSARY_DIR, f'{gid}.csv')
    if not os.path.exists(csv_path):
        return jsonify({'error': 'Không tìm thấy'}), 404
    rows = []
    with open(csv_path, 'r', encoding='utf-8-sig') as f:
        for r in csv.reader(f):
            if len(r) >= 2:
                rows.append({'dst': r[0].strip(), 'src': r[1].strip()})
            elif len(r) == 1 and r[0].strip():
                rows.append({'dst': r[0].strip(), 'src': ''})
    return jsonify(rows)


@app.route('/api/glossaries/<gid>', methods=['PUT'])
@login_required
def api_update_glossary(gid):
    """
    Cập nhật nội dung glossary.
    Body JSON: { "name": "...", "rows": [{"src": "...", "dst": "..."}, ...] }
    """
    csv_path  = os.path.join(GLOSSARY_DIR, f'{gid}.csv')
    meta_path = os.path.join(GLOSSARY_DIR, f'{gid}.meta.json')
    if not os.path.exists(csv_path):
        return jsonify({'error': 'Không tìm thấy'}), 404
    data = request.get_json()
    if not data:
        return jsonify({'error': 'Body rỗng'}), 400
    # Cập nhật tên
    if 'name' in data:
        with open(meta_path, 'w', encoding='utf-8') as mf:
            json.dump({'name': data['name']}, mf, ensure_ascii=False)
    # Cập nhật rows
    if 'rows' in data:
        with open(csv_path, 'w', encoding='utf-8-sig', newline='') as f:
            w = csv.writer(f)
            for row in data['rows']:
                if row.get('src') or row.get('dst'):
                    w.writerow([row.get('dst', ''), row.get('src', '')])
    return jsonify({'success': True})


@app.route('/api/glossaries/<gid>', methods=['DELETE'])
@login_required
def api_delete_glossary(gid):
    """Xóa glossary và meta file."""
    csv_path  = os.path.join(GLOSSARY_DIR, f'{gid}.csv')
    meta_path = os.path.join(GLOSSARY_DIR, f'{gid}.meta.json')
    for p in [csv_path, meta_path]:
        try:
            os.remove(p)
        except FileNotFoundError:
            pass
    return jsonify({'success': True})


# ==================== API: EXTRACT COLORS ====================

@app.route('/api/extract-colors', methods=['POST'])
@login_required
def api_extract_colors():
    """
    Trả về danh sách màu chữ duy nhất có trong file xlsx.
    Luôn bao gồm '000000' để đại diện cho màu đen/auto (mặc định).
    """
    if 'file' not in request.files or not request.files['file'].filename:
        return jsonify({'error': 'Không có file'}), 400

    f = request.files['file']
    if not allowed_file(f.filename):
        return jsonify({'error': 'Chỉ chấp nhận file .xlsx'}), 400

    file_bytes = f.read()
    colors = set()

    try:
        wb = load_workbook(io.BytesIO(file_bytes))
        for ws in wb.worksheets:
            for row in ws.iter_rows():
                for cell in row:
                    if cell.value is not None:
                        c = _get_font_rgb_xlsx(cell)
                        if c:
                            colors.add(c)
        wb.close()
    except Exception as e:
        return jsonify({'error': f'Lỗi đọc màu: {str(e)}'}), 500

    # Luôn bao gồm 000000 cho màu đen/auto
    colors.add('000000')

    return jsonify({'colors': sorted(colors)})


@app.route('/extract', methods=['POST'])
@login_required
def extract():
    """
    Chức năng 1: Trích xuất các cell chứa string từ file Excel (.xlsx)
    Bỏ qua các cell chứa số và công thức (bắt đầu bằng '=')
    Trả về file JSON với format: {"SheetName!CellCoordinate": "Content"}
    """
    # Kiểm tra file upload hoặc fallback từ Smart Update
    su_info = session.get('tab1_from_smart_update')
    if 'file' in request.files and request.files['file'].filename:
        file = request.files['file']
        use_session_file = False
        if not allowed_file(file.filename):
            return jsonify({'error': 'Chỉ chấp nhận file .xlsx'}), 400
    elif su_info and os.path.exists(su_info.get('filepath', '')):
        file = None
        use_session_file = True
    else:
        return jsonify({'error': 'Không có file được upload'}), 400

    try:
        # Lấy session folder
        session_folder = get_session_folder()
        
        if not use_session_file:
            # Lưu tên file gốc (giữ nguyên tiếng Nhật, ký tự đặc biệt)
            original_filename = file.filename

            # Lấy extension từ tên file gốc
            if '.' in original_filename:
                original_ext = original_filename.rsplit('.', 1)[1].lower()
            else:
                return jsonify({'error': 'Tên file phải có đuôi mở rộng (.xlsx)'}), 400

            # Tạo tên file tạm an toàn hoàn toàn từ timestamp (không dùng tên gốc)
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            safe_temp_filename = f"temp_{timestamp}.{original_ext}"
            filepath = os.path.join(session_folder, safe_temp_filename)
            file.save(filepath)
        else:
            original_filename = su_info['display_name']
            original_ext = original_filename.rsplit('.', 1)[1].lower() if '.' in original_filename else 'xlsx'
            filepath = su_info['filepath']
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            # Không pop session ở đây — giữ lại để /inject vẫn dùng được cùng file

        # Xác định loại file và trích xuất
        file_ext = original_ext

        # Đọc color filter nếu có (chuỗi HEX cách nhau bằng dấu phẩy)
        color_filter_raw = request.form.get('color_filter', '')
        color_list = [c.strip() for c in color_filter_raw.split(',') if c.strip()]
        color_filter = _normalize_color_filter(color_list) if color_list else None
        
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
                                if color_filter is None or (_get_font_rgb_xlsx(cell) or '000000') in color_filter:
                                    # Tạo key theo format "SheetName!CellCoordinate"
                                    key = f"{sheet_name}!{cell.coordinate}"
                                    extracted_data[key] = cell.value
            
            # Trích xuất text từ shapes/objects (text-box) trong xlsx
            # TODO: color filter for xlsx shapes not yet implemented
            shapes_from_xlsx = extract_xlsx_shapes(filepath)
            extracted_data.update(shapes_from_xlsx)

            # Đóng workbook
            workbook.close()
        
        # Áp dụng glossary nếu có
        glossary_ids_raw = request.form.get('glossary_ids', '')
        glossary_ids = [g.strip() for g in glossary_ids_raw.split(',') if g.strip()]
        if glossary_ids:
            extracted_data = apply_glossary(extracted_data, glossary_ids)

        # Tách dữ liệu thành nhiều file, mỗi file 400 cặp key-value
        CHUNK_SIZE = 300
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


# ==================== HELPER: core extract logic ====================

def _run_extract(filepath, original_filename, glossary_ids, session_folder, color_filter=None):
    """
    Chạy toàn bộ logic extract từ cột filepath.
    Trả về dict cho jsonify (cùng format như route /extract).
    Ném Exception nếu có lỗi.
    color_filter: set HEX strings hoặc None (không lọc)
    """
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    original_ext = original_filename.rsplit('.', 1)[-1].lower() if '.' in original_filename else 'xlsx'

    if original_ext == 'xlsx':
        workbook = load_workbook(filepath)
        extracted_data = {}
        for sheet_name in workbook.sheetnames:
            sheet = workbook[sheet_name]
            for row in sheet.iter_rows():
                for cell in row:
                    if cell.value is None:
                        continue
                    if isinstance(cell.value, str) and not cell.value.startswith('='):
                        if color_filter is None or (_get_font_rgb_xlsx(cell) or '000000') in color_filter:
                            extracted_data[f"{sheet_name}!{cell.coordinate}"] = cell.value
        # TODO: color filter for xlsx shapes not yet implemented
        extracted_data.update(extract_xlsx_shapes(filepath))
        workbook.close()
    else:
        raise ValueError(f'Không hỗ trợ định dạng .{original_ext}')

    if glossary_ids:
        extracted_data = apply_glossary(extracted_data, glossary_ids)

    CHUNK_SIZE = 400
    data_items  = list(extracted_data.items())
    total_items = len(data_items)
    num_files   = max(1, (total_items + CHUNK_SIZE - 1) // CHUNK_SIZE)

    base_filename = os.path.splitext(original_filename)[0] or f'file_{timestamp}'
    safe_base     = f'extracted_{timestamp}'
    folder_name   = f'{base_filename}_json_to_translate'
    temp_dir      = os.path.join(session_folder, f'{safe_base}_temp_{timestamp}')
    os.makedirs(temp_dir, exist_ok=True)

    json_files = []
    json_display_names = []
    for i in range(num_files):
        chunk_data  = dict(data_items[i*CHUNK_SIZE:(i+1)*CHUNK_SIZE])
        disp_name   = f'{base_filename}_part{i+1:02d}_of_{num_files:02d}.json'
        safe_name   = f'{safe_base}_part{i+1:02d}.json'
        jpath       = os.path.join(temp_dir, safe_name)
        with open(jpath, 'w', encoding='utf-8') as jf:
            json.dump(chunk_data, jf, ensure_ascii=False, indent=2)
        json_files.append(jpath)
        json_display_names.append(disp_name)

    files_data = []
    for idx, jpath in enumerate(json_files):
        with open(jpath, 'r', encoding='utf-8') as f:
            files_data.append({'name': json_display_names[idx], 'content': f.read()})

    zip_display = f'{base_filename}_json_to_translate.zip'
    safe_zip    = f'{safe_base}_json_{timestamp}.zip'
    zip_path    = os.path.join(session_folder, safe_zip)
    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_STORED) as zf:
        for idx, jpath in enumerate(json_files):
            zf.write(jpath, os.path.join(folder_name, json_display_names[idx]))

    session['extract_zip'] = {
        'path': zip_path,
        'display_name': zip_display,
        'input_path': filepath,
        'json_files': json_files,
        'temp_dir': temp_dir,
    }

    dedup_files, dedup_mapping, dedup_stats = build_dedup_data(extracted_data, CHUNK_SIZE)
    dedup_map_path = os.path.join(session_folder, 'dedup_mapping.json')
    with open(dedup_map_path, 'w', encoding='utf-8') as f:
        json.dump(dedup_mapping, f, ensure_ascii=False, indent=2)

    return {
        'success': True,
        'total_files': num_files,
        'total_items': total_items,
        'files': files_data,
        'zip_display_name': zip_display,
        'dedup_files': dedup_files,
        'dedup_stats': dedup_stats,
    }


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
    Chức năng 2: Nạp dữ liệu từ file JSON đã dịch vào file Excel (.xlsx) gốc
    Giữ nguyên định dạng, màu sắc của file gốc
    Hỗ trợ nhiều file JSON riêng lẻ hoặc file ZIP chứa nhiều file JSON
    """
    # Kiểm tra file upload
    su_info_inject = session.get('tab1_from_smart_update')

    if 'excel_file' in request.files and request.files['excel_file'].filename:
        excel_file = request.files['excel_file']
        use_session_file_inject = False
        if not allowed_file(excel_file.filename):
            return jsonify({'error': 'File phải có định dạng .xlsx'}), 400
    elif su_info_inject and os.path.exists(su_info_inject.get('filepath', '')):
        excel_file = None
        use_session_file_inject = True
    else:
        return jsonify({'error': 'Cần upload file Excel (.xlsx)'}), 400

    # Lấy pasted JSON data nếu có
    pasted_json_data = request.form.get('pasted_json_data', None)

    # Kiểm tra xem có file JSON được upload hoặc có pasted JSON không
    json_files = request.files.getlist('json_files') if 'json_files' in request.files else []

    # Kiểm tra xem có ít nhất một nguồn JSON
    has_json_files = len(json_files) > 0 and any(f.filename != '' for f in json_files)
    has_pasted_json = pasted_json_data is not None and pasted_json_data.strip() != ''

    if not has_json_files and not has_pasted_json:
        return jsonify({'error': 'Cần upload ít nhất 1 file JSON/ZIP hoặc paste JSON'}), 400

    try:
        # Lấy session folder
        session_folder = get_session_folder()
        
        if not use_session_file_inject:
            # Lưu tên file gốc (giữ nguyên tiếng Nhật, ký tự đặc biệt)
            original_excel_filename = excel_file.filename

            # Lấy extension từ tên file gốc
            if '.' in original_excel_filename:
                original_ext = original_excel_filename.rsplit('.', 1)[1].lower()
            else:
                return jsonify({'error': 'Tên file phải có đuôi mở rộng (.xlsx)'}), 400

            # Tạo tên file tạm an toàn hoàn toàn từ timestamp
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            safe_temp_filename = f"temp_{timestamp}.{original_ext}"
            excel_filepath = os.path.join(session_folder, safe_temp_filename)
            excel_file.save(excel_filepath)
        else:
            original_excel_filename = su_info_inject['display_name']
            original_ext = original_excel_filename.rsplit('.', 1)[1].lower() if '.' in original_excel_filename else 'xlsx'
            excel_filepath = su_info_inject['filepath']
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            session.pop('tab1_from_smart_update', None)

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
            # Tạo tên file output
            base_filename = os.path.splitext(original_excel_filename)[0]  # Tên gốc với tiếng Nhật
            
            output_display_name = f"{base_filename}_translated.xlsx"  # Tên hiển thị
            safe_output_filename = f"output_{timestamp}.xlsx"  # Tên file trong filesystem
            output_filepath = os.path.join(session_folder, safe_output_filename)

            # ZIP-level patch: nạp cả cell + shape trực tiếp trên package gốc
            inject_xlsx_shapes(excel_filepath, output_filepath, json_data)
            
            output_mimetype = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        
        # Trả về file đã được nạp dữ liệu (dùng tên hiển thị)
        response = send_file(
            output_filepath,
            mimetype=output_mimetype
        )
        
        response = set_download_headers(response, output_display_name, 'download.xlsx')
        
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
    # app.run(host='0.0.0.0', port=5020)
    app.run(debug=True,host='0.0.0.0', port=5020)
