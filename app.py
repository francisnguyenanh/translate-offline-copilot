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
from datetime import datetime, timedelta
from urllib.parse import quote
from flask import Flask, render_template, request, send_file, jsonify, session, redirect, url_for
from werkzeug.utils import secure_filename
from copy import deepcopy
from openpyxl import load_workbook
from pptx import Presentation
from docx import Document
from functools import wraps
from lxml import etree as _etree

# Khởi tạo ứng dụng Flask
app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 60 * 1024 * 1024  # Giới hạn 50MB
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['SECRET_KEY'] = os.urandom(24)  # Secret key cho session
app.config['PERMANENT_SESSION_LIFETIME'] = timedelta(hours=5)  # Session timeout 5h

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


# ==================== SMART UPDATE (MATRIX MAPPING INHERITANCE) ====================

import tempfile as _tempfile
from openpyxl.cell import MergedCell as _MergedCell

# Mặc định: màu xanh lá trong VN_1.1 = nội dung mới/sửa
DEFAULT_NEW_COLORS = {'38761D', '00FF00', '008000', '006100', '00B050'}
DEFAULT_RED_COLORS  = {'FF0000', 'FF0000'}


def _get_rgb6(cell) -> str:
    """Đọc mã RGB 6 ký tự từ font.color, trả rỗng nếu không xác định được."""
    try:
        color = cell.font.color if (cell.font and cell.font.color) else None
        if color and color.type == 'rgb' and color.rgb:
            return str(color.rgb).upper().lstrip('F')[-6:].zfill(6)
    except Exception:
        pass
    return ''


def is_new_or_modified_cell(cell_vn11, new_colors=None, red_colors=None) -> bool:
    """
    Trả về True nếu ô được đánh dấu màu xanh lá (nội dung mới/sửa).
    Trả về False nếu màu đỏ (marker) hoặc đen/auto (nội dung cũ).
    None → không xác định qua màu, cần dng fallback.
    """
    if new_colors is None:
        new_colors = DEFAULT_NEW_COLORS
    if red_colors is None:
        red_colors = DEFAULT_RED_COLORS
    rgb = _get_rgb6(cell_vn11)
    if not rgb:
        return None  # không có màu rõ ràng → fallback
    if rgb in new_colors:
        return True
    if rgb in red_colors:
        return False  # marker, bỏ qua
    return None  # màu khác → fallback sang so sánh nội dung


def is_content_changed(cell_vn10, cell_vn11) -> bool:
    """Fallback: so sánh nội dung text giữa VN_1.0 và VN_1.1 tại cùng tọa độ."""
    v0 = str(cell_vn10.value).strip() if cell_vn10.value is not None else ''
    v1 = str(cell_vn11.value).strip() if cell_vn11.value is not None else ''
    return bool(v1) and v1 != v0


def _should_skip_cell(cell) -> bool:
    """Bỏ qua ô công thức, merged non-top-left, hoặc rỗng."""
    if isinstance(cell, _MergedCell):
        return True
    if cell.value is None:
        return True
    if isinstance(cell.value, str):
        if cell.value.strip() == '':
            return True
        if cell.value.startswith('='):
            return True
    return False


def _safe_set_value(ws, coord, value):
    """
    Ghi giá trị vào ô worksheet, bỏ qua nếu ô đó là MergedCell (non-top-left).
    Tránh lỗi ‘MergedCell’ object attribute ‘value’ is read-only.
    """
    cell = ws[coord]
    if isinstance(cell, _MergedCell):
        return  # chỉ top-left của merge mới ghi được
    cell.value = value


def _build_coord_content_map(ws_vn10, ws_jp10) -> dict:
    """
    Map: { (coordinate, vn_text) → jp_text }

    Key là TUPLE (tọa độ, nội dung VN) — chỉ kế thừa JP khi CẢ HAI đều khớp:
    - Tọa độ giống nhau (coord)
    - Nội dung VN tại tọa độ đó không thay đổi (vn_text)

    Đây là tầng chính xác nhất: nếu VN_1.1[A5] == VN_1.0[A5] thì JP_1.0[A5]
    chắc chắn là bản dịch đúng cho ô đó, bất kể text có trùng với ô khác không.
    """
    result = {}
    for row in ws_vn10.iter_rows():
        for cell_vn in row:
            if isinstance(cell_vn, _MergedCell) or not cell_vn.value:
                continue
            vn_text = str(cell_vn.value).strip()
            if not vn_text or vn_text.startswith('='):
                continue
            jp_cell = ws_jp10[cell_vn.coordinate]
            if isinstance(jp_cell, _MergedCell) or not jp_cell.value:
                continue
            jp_text = str(jp_cell.value).strip()
            if jp_text:
                result[(cell_vn.coordinate, vn_text)] = jp_text
    return result


def _build_vn_jp_content_map(ws_vn10, ws_jp10) -> dict:
    """
    Map: { vn_text → jp_text } — fallback khi tọa độ đã thay đổi.

    Với mỗi VN text, đếm số lần xuất hiện của từng bản dịch JP tương ứng,
    rồi chọn JP xuất hiện NHIỀU NHẤT (dominant).

    Lý do dùng dominant thay vì loại bỏ conflict: bản dịch JP xuất hiện
    nhiều nhất là bản khách đã chấp nhận nhiều lần — đáng tin cậy hơn.

    Ví dụ:
      "Varchar" → {"varchar": 6, "数値": 1} → trả về "varchar"  (dominant)
      "Label"   → {"ラベル": 5}              → trả về "ラベル"
    """
    from collections import defaultdict
    vn_to_jp_counter = defaultdict(lambda: defaultdict(int))

    for row in ws_vn10.iter_rows():
        for cell_vn in row:
            if isinstance(cell_vn, _MergedCell) or not cell_vn.value:
                continue
            vn_text = str(cell_vn.value).strip()
            if not vn_text or vn_text.startswith('='):
                continue
            jp_cell = ws_jp10[cell_vn.coordinate]
            if isinstance(jp_cell, _MergedCell) or not jp_cell.value:
                continue
            jp_text = str(jp_cell.value).strip()
            if jp_text:
                vn_to_jp_counter[vn_text][jp_text] += 1

    return {
        vn_text: max(jp_counter, key=jp_counter.get)
        for vn_text, jp_counter in vn_to_jp_counter.items()
    }


def _clone_vn11_as_base(path_vn11: str, tmp_path: str) -> None:
    """
    Clone VN_1.1 thành file base cho JP_1.1.
    Giữ nguyên TOÀN BỘ cấu trúc (merge cells, row heights, col widths, styles).
    Sau đó xóa sạch text trong các cell (giữ format) để điền JP vào.
    """
    import shutil
    shutil.copy2(path_vn11, tmp_path)
    wb = load_workbook(tmp_path)
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if isinstance(cell, _MergedCell):
                    continue
                if cell.value is not None:
                    cell.value = None
    wb.save(tmp_path)
    wb.close()


def _copy_sheet_structure(ws_source, ws_dest):
    """Copy column widths, row heights, merge cells từ source sang dest."""
    for merge in ws_source.merged_cells.ranges:
        try:
            ws_dest.merge_cells(str(merge))
        except Exception:
            pass
    for col_letter, col_dim in ws_source.column_dimensions.items():
        ws_dest.column_dimensions[col_letter].width = col_dim.width
    for row_num, row_dim in ws_source.row_dimensions.items():
        ws_dest.row_dimensions[row_num].height = row_dim.height


def smart_update_excel(path_vn10, path_vn11, path_jp10, new_colors=None, red_colors=None):
    """
    Tạo JP_1.1 bằng cách:
      - Clone VN_1.1 làm base (cấu trúc đúng nhất: merge, rows, cols)
      - Ô màu đỏ (marker)       → giữ text VN gốc
      - Ô màu xanh (new/sửa)    → giữ text VN_1.1, thêm vào to_translate
      - Ô đen/auto               → tầng 1 (coord+text) → tầng 2 (dominant JP) → giữ text VN_1.1
      - Sheet mới trong VN_1.1   → toàn bộ giữ text VN_1.1 + to_translate
      - Sheet bỏ trong VN_1.1    → bỏ qua (JP_1.1 theo cấu trúc VN_1.1)
      - Sheet có trong VN nhưng JP_1.0 thiếu → dùng text VN_1.1 thức đẩy to_translate
    Returns: (wb_jp11, to_translate_dict, stats_dict)
    """
    import shutil

    wb_vn10 = load_workbook(path_vn10, data_only=True)
    wb_vn11 = load_workbook(path_vn11, data_only=True)
    wb_jp10 = load_workbook(path_jp10, data_only=True)

    # Clone VN_1.1 → base của JP_1.1 (cấu trúc đúng nhất)
    tmp_path = _tempfile.mktemp(suffix='_jp11.xlsx')
    _clone_vn11_as_base(path_vn11, tmp_path)
    wb_jp11 = load_workbook(tmp_path)

    to_translate = {}
    inherited = 0
    sheet_stats = {}   # per-sheet breakdown

    for sheet_name in wb_vn11.sheetnames:
        ws_vn11 = wb_vn11[sheet_name]
        ws_vn10 = wb_vn10[sheet_name] if sheet_name in wb_vn10.sheetnames else None
        ws_jp10 = wb_jp10[sheet_name] if sheet_name in wb_jp10.sheetnames else None
        ws_jp11 = wb_jp11[sheet_name] if sheet_name in wb_jp11.sheetnames \
            else wb_jp11.create_sheet(sheet_name)

        sh_inherited   = 0
        sh_to_translate = 0

        # Case A: Sheet mới trong VN_1.1 (không có trong VN_1.0 và/hoặc JP_1.0)
        #         → toàn bộ ô giữ text VN_1.1 và được đánh dấu cần dịch
        if ws_vn10 is None:
            for row in ws_vn11.iter_rows():
                for cell in row:
                    if _should_skip_cell(cell):
                        continue
                    key = f"{sheet_name}!{cell.coordinate}"
                    to_translate[key] = str(cell.value).strip()
                    _safe_set_value(ws_jp11, cell.coordinate, cell.value)  # giữ text VN_1.1
                    sh_to_translate += 1
            sheet_stats[sheet_name] = {'inherited': 0, 'to_translate': sh_to_translate, 'status': 'new_sheet'}
            continue

        # Case B: Sheet tồn tại trong cả VN_1.0 và VN_1.1 → xử lý kế thừa
        # ws_jp10 có thể là None (sheet chưa tồn tại trong JP_1.0) → map rỗng → tất cả đen→translate
        coord_content_map = _build_coord_content_map(ws_vn10, ws_jp10) \
            if (ws_vn10 and ws_jp10) else {}
        vn_jp_content_map = _build_vn_jp_content_map(ws_vn10, ws_jp10) \
            if (ws_vn10 and ws_jp10) else {}

        for row in ws_vn11.iter_rows():
            for cell_vn11 in row:
                if _should_skip_cell(cell_vn11):
                    continue

                coord     = cell_vn11.coordinate
                text_vn11 = str(cell_vn11.value).strip()
                key       = f"{sheet_name}!{coord}"

                color_result = is_new_or_modified_cell(cell_vn11, new_colors, red_colors)

                if color_result is False:
                    # Màu đỏ = marker (①②...) → giữ nguyên text VN trong JP_1.1
                    _safe_set_value(ws_jp11, coord, cell_vn11.value)
                    continue

                if color_result is True:
                    # Màu xanh = nội dung mới/sửa → cần dịch mới
                    to_translate[key] = text_vn11
                    _safe_set_value(ws_jp11, coord, cell_vn11.value)  # giữ text VN_1.1
                    sh_to_translate += 1
                    continue

                # Không màu rõ ràng → fallback theo 2 tầng

                # Tầng 1: coord + vn_text đều khớp VN_1.0 → JP_1.0[coord] chắc chắn đúng
                jp_val = coord_content_map.get((coord, text_vn11))
                if jp_val:
                    _safe_set_value(ws_jp11, coord, jp_val)
                    inherited += 1
                    sh_inherited += 1
                    continue

                # Tầng 2: vn_text có trong VN_1.0 → lấy JP dominant (bản khách đã chấp nhận)
                jp_val = vn_jp_content_map.get(text_vn11)
                if jp_val:
                    _safe_set_value(ws_jp11, coord, jp_val)
                    inherited += 1
                    sh_inherited += 1
                    continue

                # Không tìm được → text mới hoàn toàn → cần dịch mới
                to_translate[key] = text_vn11
                _safe_set_value(ws_jp11, coord, cell_vn11.value)  # giữ text VN_1.1
                sh_to_translate += 1

        sheet_stats[sheet_name] = {
            'inherited': sh_inherited,
            'to_translate': sh_to_translate,
            'status': 'updated' if ws_jp10 else 'no_jp10',
        }

    wb_vn10.close()
    wb_vn11.close()
    wb_jp10.close()

    try:
        os.remove(tmp_path)
    except Exception:
        pass

    stats = {
        'inherited': inherited,
        'to_translate': len(to_translate),
        'total': inherited + len(to_translate),
        'sheets': sheet_stats,
    }
    return wb_jp11, to_translate, stats


# giữ alias để bảo tương thích ngược với code cũ (nếu có chỗ nào gọi)
compare_and_inherit_excel = smart_update_excel


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

# ==================== XLSX SHAPE / OBJECT SUPPORT ====================
# Namespaces dùng trong drawing XML của xlsx
_NS_XDR = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'
_NS_A   = 'http://schemas.openxmlformats.org/drawingml/2006/main'
_NS_R   = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
_NS_WB  = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'


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

    cell_elem.set('t', 'inlineStr')
    for child in list(cell_elem):
        if child.tag in {
            f'{{{_NS_WB}}}v',
            f'{{{_NS_WB}}}f',
            f'{{{_NS_WB}}}is',
        }:
            cell_elem.remove(child)

    is_elem = _etree.SubElement(cell_elem, f'{{{_NS_WB}}}is')
    t_elem = _etree.SubElement(is_elem, f'{{{_NS_WB}}}t')
    text_str = '' if text_value is None else str(text_value)
    t_elem.text = text_str
    if text_str != text_str.strip() or '\n' in text_str:
        t_elem.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')


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
    # Kiểm tra file upload hoặc fallback từ Smart Update
    su_info = session.get('tab1_from_smart_update')
    if 'file' in request.files and request.files['file'].filename:
        file = request.files['file']
        use_session_file = False
        if not allowed_file(file.filename):
            return jsonify({'error': 'Chỉ chấp nhận file .xlsx, .pptx hoặc .docx'}), 400
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
                return jsonify({'error': 'Tên file phải có đuôi mở rộng (.xlsx hoặc .pptx)'}), 400

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
            
            # Trích xuất text từ shapes/objects (text-box) trong xlsx
            shapes_from_xlsx = extract_xlsx_shapes(filepath)
            extracted_data.update(shapes_from_xlsx)

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
    # Kiểm tra file upload hoặc fallback từ Smart Update
    su_info_inject = session.get('tab1_from_smart_update')
    if 'excel_file' in request.files and request.files['excel_file'].filename:
        excel_file = request.files['excel_file']
        use_session_file_inject = False
        if not allowed_file(excel_file.filename):
            return jsonify({'error': 'File phải có định dạng .xlsx, .pptx hoặc .docx'}), 400
    elif su_info_inject and os.path.exists(su_info_inject.get('filepath', '')):
        excel_file = None
        use_session_file_inject = True
    else:
        return jsonify({'error': 'Cần upload file Excel, PPTX hoặc DOCX'}), 400

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
                return jsonify({'error': 'Tên file phải có đuôi mở rộng (.xlsx hoặc .pptx)'}), 400

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

# ==================== SMART UPDATE ROUTES ====================

@app.route('/smart-update', methods=['POST'])
@login_required
def smart_update_route():
    """
    Smart Update: So sánh và kế thừa bản dịch Excel từ version cũ.
    Input (form-data):
        file_vn_old / file_vn10  : VN_1.0 (chưa dịch, phiên bản cũ)
        file_vn_new / file_vn11  : VN_1.1 (chưa dịch, phiên bản mới)
        file_jp_old / file_jp10  : JP_1.0 (đã dịch, phiên bản cũ)
        new_colors (optional)    : danh sách RGB6 cách nhau dấu phẩy, VD "38761D,00B050"
    Output JSON: stats + link tải file
    """
    # Hỗ trợ cả 2 bộ tên trường (cũ và mới)
    def _get_file(new_name, old_name):
        f = request.files.get(new_name) or request.files.get(old_name)
        return f if (f and f.filename) else None

    file_vn10 = _get_file('file_vn_old', 'file_vn10')
    file_vn11 = _get_file('file_vn_new', 'file_vn11')
    file_jp10 = _get_file('file_jp_old', 'file_jp10')

    missing = [n for n, f in [('VN cũ', file_vn10), ('VN mới', file_vn11), ('JP cũ', file_jp10)] if not f]
    if missing:
        return jsonify({'error': f'Thiếu file: {", ".join(missing)}'}), 400

    for f in [file_vn10, file_vn11, file_jp10]:
        if not f.filename.lower().endswith('.xlsx'):
            return jsonify({'error': 'Smart Update chỉ hỗ trợ file .xlsx'}), 400

    # Parse màu xanh tùy chỉnh từ form (nếu có)
    raw_colors = request.form.get('new_colors', '').strip()
    custom_new_colors = None
    if raw_colors:
        custom_new_colors = set(
            c.strip().upper().lstrip('#').lstrip('FF')[-6:].zfill(6)
            for c in raw_colors.split(',') if c.strip()
        ) | DEFAULT_NEW_COLORS  # merge với mặc định

    try:
        session_folder = get_session_folder()
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')

        path_vn10 = os.path.join(session_folder, f'su_vn10_{timestamp}.xlsx')
        path_vn11 = os.path.join(session_folder, f'su_vn11_{timestamp}.xlsx')
        path_jp10 = os.path.join(session_folder, f'su_jp10_{timestamp}.xlsx')
        file_vn10.save(path_vn10)
        file_vn11.save(path_vn11)
        file_jp10.save(path_jp10)

        result_wb, to_translate, stats = smart_update_excel(
            path_vn10, path_vn11, path_jp10,
            new_colors=custom_new_colors
        )

        # Lưu workbook kết quả JP_1.1
        original_name = os.path.splitext(file_vn11.filename)[0]
        result_display_name = f"{original_name}_JP_1_1.xlsx"
        safe_result_path = os.path.join(session_folder, f'su_result_{timestamp}.xlsx')
        result_wb.save(safe_result_path)
        result_wb.close()

        # Chia các ô cần dịch thành JSON chunks (≤400/file)
        CHUNK_SIZE = 400
        items = list(to_translate.items())
        num_chunks = max(1, (len(items) + CHUNK_SIZE - 1) // CHUNK_SIZE)

        temp_json_dir = os.path.join(session_folder, f'su_json_{timestamp}')
        os.makedirs(temp_json_dir, exist_ok=True)

        json_display_names = []
        json_paths = []
        files_data = []

        for i in range(num_chunks):
            chunk = dict(items[i * CHUNK_SIZE:(i + 1) * CHUNK_SIZE])
            jname_display = f"{original_name}_to_translate_part{i+1:02d}_of_{num_chunks:02d}.json"
            jname_safe = f"su_json_part{i+1:02d}_{timestamp}.json"
            jpath = os.path.join(temp_json_dir, jname_safe)
            with open(jpath, 'w', encoding='utf-8') as jf:
                json.dump(chunk, jf, ensure_ascii=False, indent=2)
            json_display_names.append(jname_display)
            json_paths.append(jpath)
            with open(jpath, 'r', encoding='utf-8') as jf:
                files_data.append({'name': jname_display, 'content': jf.read()})

        # Tạo ZIP chứa tất cả JSON cần dịch
        zip_display_name = f"{original_name}_to_translate.zip"
        zip_path = os.path.join(session_folder, f'su_json_{timestamp}.zip')
        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_STORED) as zf:
            for idx, jpath_item in enumerate(json_paths):
                zf.write(jpath_item, json_display_names[idx])

        session['smart_update'] = {
            'result_path': safe_result_path,
            'result_display_name': result_display_name,
            'result_filename': os.path.basename(safe_result_path),
            'path_vn11': path_vn11,
            'zip_path': zip_path,
            'zip_display_name': zip_display_name,
            'temp_files': [path_vn10, path_vn11, path_jp10] + json_paths,
            'temp_dir': temp_json_dir,
        }

        return jsonify({
            'success': True,
            'stats': stats,
            'result_display_name': result_display_name,
            'zip_display_name': zip_display_name,
            'files': files_data,
        })

    except Exception as e:
        import traceback
        return jsonify({'error': f'Lỗi xử lý: {str(e)}', 'detail': traceback.format_exc()}), 500


@app.route('/download-smart-excel', methods=['GET'])
@login_required
def download_smart_excel():
    """Tải file Excel JP 1.1 hỗn hợp từ Smart Update"""
    info = session.get('smart_update')
    if not info or not os.path.exists(info.get('result_path', '')):
        return jsonify({'error': 'Không tìm thấy file. Vui lòng chạy Smart Update lại.'}), 404
    response = send_file(
        info['result_path'],
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )
    response = set_download_headers(response, info['result_display_name'], 'result_mixed.xlsx')
    return response


@app.route('/download-smart-zip', methods=['GET'])
@login_required
def download_smart_zip():
    """Tải ZIP các file JSON cần dịch từ Smart Update và dọn dẹp file tạm"""
    info = session.get('smart_update')
    if not info or not os.path.exists(info.get('zip_path', '')):
        return jsonify({'error': 'Không tìm thấy file ZIP. Vui lòng chạy Smart Update lại.'}), 404

    zip_path = info['zip_path']
    zip_display_name = info['zip_display_name']
    session.pop('smart_update', None)

    response = send_file(zip_path, mimetype='application/zip')
    response = set_download_headers(response, zip_display_name, 'new_strings.zip')

    temp_files = info.get('temp_files', [])
    temp_dir = info.get('temp_dir')
    result_path = info.get('result_path')

    @response.call_on_close
    def cleanup():
        import time
        import gc
        gc.collect()
        time.sleep(0.1)
        for p in [zip_path, result_path] + temp_files:
            try:
                if p and os.path.exists(p):
                    os.remove(p)
            except Exception:
                pass
        try:
            if temp_dir and os.path.exists(temp_dir):
                shutil.rmtree(temp_dir)
        except Exception:
            pass

    return response


@app.route('/smart-update/use-as-input', methods=['POST'])
@login_required
def smart_update_use_as_input():
    """
    Chuyển file kết quả Smart Update (JP_1.1) thành file input cho Tab Dịch tài liệu.
    Frontend gọi route này, sau đó chuyển sang Tab 1 — file đã sẵn sàng để Extract/Inject.
    """
    info = session.get('smart_update')
    if not info or not os.path.exists(info.get('result_path', '')):
        return jsonify({'error': 'Không tìm thấy file Smart Update. Vui lòng chạy lại.'}), 404

    result_path  = info['result_path']
    display_name = info['result_display_name']

    # Ghi vào session của Tab 1 — giống như user vừa upload file này
    session['tab1_from_smart_update'] = {
        'filepath':     result_path,
        'display_name': display_name,
    }

    return jsonify({
        'success':      True,
        'display_name': display_name,
    })


# Đảm bảo thư mục uploads tồn tại
if not os.path.exists(app.config['UPLOAD_FOLDER']):
    os.makedirs(app.config['UPLOAD_FOLDER'])


# ==================== IMAGE TRANSLATION ====================
ALLOWED_IMAGE_EXTENSIONS = {'png', 'jpg', 'jpeg', 'gif', 'webp', 'bmp'}

def allowed_image(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_IMAGE_EXTENSIONS


@app.route('/img-translate/upload', methods=['POST'])
@login_required
def img_translate_upload():
    """Upload ảnh vào session folder, trả về filename."""
    if 'image' not in request.files:
        return jsonify({'error': 'Không tìm thấy file ảnh'}), 400
    file = request.files['image']
    if file.filename == '':
        return jsonify({'error': 'Chưa chọn file'}), 400
    if not allowed_image(file.filename):
        return jsonify({'error': 'Định dạng không hỗ trợ. Chỉ nhận: PNG, JPG, JPEG, GIF, WEBP, BMP'}), 400

    session_folder = get_session_folder()
    filename = secure_filename(file.filename)
    ts = datetime.now().strftime('%H%M%S')
    name, ext = os.path.splitext(filename)
    safe_filename = f"img_{ts}_{name[:30]}{ext}"
    filepath = os.path.join(session_folder, safe_filename)
    file.save(filepath)
    return jsonify({'success': True, 'filename': safe_filename}), 200


@app.route('/img-translate/image/<filename>', methods=['GET'])
@login_required
def img_translate_serve(filename):
    """Serve ảnh đã upload từ session folder."""
    session_folder = get_session_folder()
    safe = secure_filename(filename)
    filepath = os.path.join(session_folder, safe)
    if not os.path.exists(filepath):
        return jsonify({'error': 'File không tồn tại'}), 404
    return send_file(filepath)


@app.route('/img-translate/prompt', methods=['POST'])
@login_required
def img_translate_prompt():
    """Sinh Instruction Prompt yêu cầu AI phân tích ảnh và trả JSON overlay."""
    data = request.get_json() or {}
    target_lang = data.get('target_lang', 'tiếng Nhật')
    source_lang = data.get('source_lang', '').strip()
    img_w = int(data.get('image_width', 0))
    img_h = int(data.get('image_height', 0))

    source_note = f" (ngôn ngữ gốc trong ảnh: {source_lang})" if source_lang else ""

    if img_w > 0 and img_h > 0:
        dim_note = f"\n\nKích thước ảnh CHÍNH XÁC: {img_w} × {img_h} pixel (rộng × cao).\n" \
                   f"→ Quy đổi tọa độ pixel sang %: left_pct = pixel_x / {img_w} × 100, top_pct = pixel_y / {img_h} × 100\n" \
                   f"→ Quy đổi kích thước sang %: width_pct = pixel_w / {img_w} × 100, height_pct = pixel_h / {img_h} × 100\n" \
                   f"→ font_size_pct = chiều_cao_font_pixel / {img_h} × 100"
        font_example = round(24 / img_h * 100, 2) if img_h else 2.5
        dim_font_note = f"(ví dụ: chữ 24px trong ảnh {img_h}px cao → font_size_pct = {font_example})"
    else:
        dim_note = ""
        font_example = 2.5
        dim_font_note = "(ví dụ: 2.5)"

    prompt = f"""Bạn là chuyên gia OCR và dịch thuật chuyên nghiệp. Tôi sẽ gửi cho bạn một bức ảnh{source_note}.{dim_note}

NHIỆM VỤ:
1. Nhận diện (OCR) TẤT CẢ các vùng có văn bản trong ảnh.
2. Dịch toàn bộ sang {target_lang} một cách tự nhiên, chính xác.
3. Với mỗi vùng văn bản, xác định CÁC GIÁ TRỊ SAU:

   top_pct    : tọa độ mép TRÊN của text box, tính bằng % chiều CAO ảnh (0 = trên cùng, 100 = dưới cùng)
   left_pct   : tọa độ mép TRÁI của text box, tính bằng % chiều RỘNG ảnh (0 = trái, 100 = phải)
   width_pct  : chiều RỘNG text box, % chiều rộng ảnh
   height_pct : chiều CAO text box, % chiều cao ảnh
   bg_color   : mã HEX màu nền THỰC TẾ ngay phía sau văn bản (để che chữ cũ)
   text_color : mã HEX màu chữ phù hợp để đọc được trên bg_color
   font_size_pct : cỡ chữ tính bằng % chiều cao ảnh {dim_font_note}
                   → PHẢI xấp xỉ bằng chiều cao thực tế của 1 dòng chữ trong ảnh
                   → KHÔNG được nhỏ hơn 60% height_pct (nếu block là 1 dòng)
                   → Với block nhiều dòng: font_size_pct ≈ height_pct / số_dòng × 0.8

⚠️ YÊU CẦU BẮT BUỘC:
- Tọa độ phải bao phủ CHÍNH XÁC vùng chứa văn bản, sai số không quá 1%
- Các box KHÔNG được chồng lên nhau (trừ khi text thực sự chồng trong ảnh)
- bg_color lấy từ màu nền thực trong ảnh, KHÔNG đặt màu tùy ý
- Trả về JSON THUẦN TÚY — KHÔNG giải thích, KHÔNG bọc markdown code block ```

FORMAT JSON TRẢ VỀ (giữ đúng cấu trúc này):
{{
  "text_blocks": [
    {{
      "original": "Văn bản gốc trong ảnh",
      "translated": "Bản dịch sang {target_lang}",
      "top_pct": 10.5,
      "left_pct": 5.2,
      "width_pct": 30.0,
      "height_pct": 5.0,
      "bg_color": "#FFFFFF",
      "text_color": "#000000",
      "font_size_pct": {font_example}
    }}
  ]
}}"""

    return jsonify({'prompt': prompt}), 200


if __name__ == '__main__':
    # Chạy ứng dụng Flask ở chế độ debug
    app.run(debug=True, host='0.0.0.0', port=5001)
