# -*- coding: utf-8 -*-
"""
Ứng dụng Flask quản lý trích xuất và nạp bản dịch cho file Excel
"""

import os
import json
import zipfile
from flask import Flask, render_template, request, send_file, jsonify
from werkzeug.utils import secure_filename
from openpyxl import load_workbook
from pptx import Presentation
from datetime import datetime

# Khởi tạo ứng dụng Flask
app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # Giới hạn 16MB
app.config['UPLOAD_FOLDER'] = 'uploads'

# Các định dạng file được phép
ALLOWED_EXTENSIONS = {'xlsx', 'pptx'}

def allowed_file(filename):
    """
    Kiểm tra xem file có phải định dạng được phép không
    """
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def allowed_json_file(filename):
    """
    Kiểm tra xem file có phải định dạng JSON không
    """
    return '.' in filename and filename.rsplit('.', 1)[1].lower() == 'json'

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
            # Nạp vào table cell
            row_idx, col_idx = table_pos
            if hasattr(shape, "has_table") and shape.has_table:
                table = shape.table
                if row_idx < len(table.rows) and col_idx < len(table.rows[row_idx].cells):
                    cell = table.rows[row_idx].cells[col_idx]
                    cell.text = translated_value
        else:
            # Nạp vào text frame của shape
            if hasattr(shape, "text_frame"):
                shape.text_frame.clear()
                p = shape.text_frame.paragraphs[0]
                p.text = translated_value
        return True
    
    # Navigate đến shape con
    if hasattr(shape, "shapes"):
        next_idx = shape_indices[0]
        if next_idx <= len(shape.shapes):
            child_shape = shape.shapes[next_idx - 1]  # Chuyển từ 1-indexed sang 0-indexed
            return inject_text_to_shape(child_shape, shape_indices[1:], translated_value, is_table_cell, table_pos)
    
    return False

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

@app.route('/')
def index():
    """
    Trang chủ hiển thị dashboard với 2 chức năng Extract và Inject
    """
    return render_template('index.html')

@app.route('/extract', methods=['POST'])
def extract():
    """
    Chức năng 1: Trích xuất các cell chứa string từ file Excel hoặc PPTX
    Bỏ qua các cell chứa số và công thức (bắt đầu bằng '=') trong Excel
    Trả về file JSON với format: {"SheetName!CellCoordinate": "Content"} hoặc {"SlideX!ShapeY": "Content"}
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
        return jsonify({'error': 'Chỉ chấp nhận file .xlsx hoặc .pptx'}), 400
    
    try:
        # Lưu file tạm thời
        filename = secure_filename(file.filename)
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        temp_filename = f"temp_{timestamp}_{filename}"
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], temp_filename)
        file.save(filepath)
        
        # Xác định loại file và trích xuất
        file_ext = filename.rsplit('.', 1)[1].lower()
        
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
        
        # Xóa file tạm
        os.remove(filepath)
        
        # Tách dữ liệu thành nhiều file, mỗi file 400 cặp key-value
        CHUNK_SIZE = 400
        data_items = list(extracted_data.items())
        total_items = len(data_items)
        num_files = (total_items + CHUNK_SIZE - 1) // CHUNK_SIZE  # Làm tròn lên
        
        # Tên folder trong ZIP
        folder_name = f"to_translate_{timestamp}"
        
        # Tạo thư mục tạm để chứa các file JSON
        temp_dir = os.path.join(app.config['UPLOAD_FOLDER'], folder_name)
        os.makedirs(temp_dir, exist_ok=True)
        
        json_files = []
        
        # Tạo các file JSON nhỏ
        for i in range(num_files):
            start_idx = i * CHUNK_SIZE
            end_idx = min((i + 1) * CHUNK_SIZE, total_items)
            chunk_data = dict(data_items[start_idx:end_idx])
            
            # Tên file với số thứ tự
            json_filename = f"to_translate_{timestamp}_part{i+1:02d}_of_{num_files:02d}.json"
            json_filepath = os.path.join(temp_dir, json_filename)
            
            # Lưu dữ liệu vào file JSON với encoding UTF-8
            with open(json_filepath, 'w', encoding='utf-8') as json_file:
                json.dump(chunk_data, json_file, ensure_ascii=False, indent=2)
            
            json_files.append(json_filepath)
        
        # Tạo file ZIP chứa folder và các file JSON
        zip_filename = f"to_translate_{timestamp}.zip"
        zip_filepath = os.path.join(app.config['UPLOAD_FOLDER'], zip_filename)
        
        # Dùng ZIP_STORED để không nén file JSON (giữ nguyên text có thể đọc được)
        with zipfile.ZipFile(zip_filepath, 'w', zipfile.ZIP_STORED) as zipf:
            for json_filepath in json_files:
                # Thêm file vào ZIP với đường dẫn folder/filename
                arcname = os.path.join(folder_name, os.path.basename(json_filepath))
                zipf.write(json_filepath, arcname)
        
        # Xóa các file JSON tạm và thư mục tạm
        for json_filepath in json_files:
            if os.path.exists(json_filepath):
                os.remove(json_filepath)
        if os.path.exists(temp_dir):
            os.rmdir(temp_dir)
        
        # Trả về file ZIP và xóa sau khi gửi
        response = send_file(
            zip_filepath,
            as_attachment=True,
            download_name=zip_filename,
            mimetype='application/zip'
        )
        
        # Xóa file JSON sau khi gửi (sử dụng after_request để đảm bảo file đã được gửi)
        @response.call_on_close
        def cleanup():
            try:
                if os.path.exists(zip_filepath):
                    os.remove(zip_filepath)
            except Exception:
                pass
        
        return response
        
    except Exception as e:
        # Xử lý lỗi
        return jsonify({'error': f'Lỗi khi xử lý file: {str(e)}'}), 500

@app.route('/inject', methods=['POST'])
def inject():
    """
    Chức năng 2: Nạp dữ liệu từ file JSON đã dịch vào file Excel hoặc PPTX gốc
    Giữ nguyên định dạng, màu sắc của file gốc
    Hỗ trợ nhiều file JSON riêng lẻ hoặc file ZIP chứa nhiều file JSON
    """
    # Kiểm tra xem có file được upload không
    if 'excel_file' not in request.files:
        return jsonify({'error': 'Cần upload file Excel hoặc PPTX'}), 400
    
    excel_file = request.files['excel_file']
    
    # Kiểm tra xem có file JSON được upload không
    if 'json_files' not in request.files:
        return jsonify({'error': 'Cần upload ít nhất 1 file JSON hoặc ZIP'}), 400
    
    json_files = request.files.getlist('json_files')
    
    # Kiểm tra xem các file có được chọn không
    if excel_file.filename == '' or len(json_files) == 0:
        return jsonify({'error': 'Cần chọn đủ file và JSON'}), 400
    
    # Kiểm tra định dạng file
    if not allowed_file(excel_file.filename):
        return jsonify({'error': 'File phải có định dạng .xlsx hoặc .pptx'}), 400
    
    try:
        # Lưu file Excel tạm thời
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        excel_filename = secure_filename(excel_file.filename)
        temp_excel_filename = f"temp_{timestamp}_{excel_filename}"
        excel_filepath = os.path.join(app.config['UPLOAD_FOLDER'], temp_excel_filename)
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
                zip_filepath = os.path.join(app.config['UPLOAD_FOLDER'], temp_zip_filename)
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
        
        
        # Xác định loại file và nạp dữ liệu
        file_ext = excel_filename.rsplit('.', 1)[1].lower()
        
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
            output_filename = f"output_translated_{timestamp}.xlsx"
            output_filepath = os.path.join(app.config['UPLOAD_FOLDER'], output_filename)
            
            # Lưu file Excel đã được nạp dữ liệu
            workbook.save(output_filepath)
            workbook.close()
            
            output_mimetype = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        
        elif file_ext == 'pptx':
            # Nạp text vào PPTX
            prs = inject_text_to_pptx(excel_filepath, json_data)
            
            # Tạo tên file output
            output_filename = f"output_translated_{timestamp}.pptx"
            output_filepath = os.path.join(app.config['UPLOAD_FOLDER'], output_filename)
            
            # Lưu file PPTX đã được nạp dữ liệu
            prs.save(output_filepath)
            
            output_mimetype = 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
        
        # Xóa file Excel tạm
        os.remove(excel_filepath)
        
        # Xóa tất cả file ZIP tạm
        for temp_file in temp_files:
            if os.path.exists(temp_file):
                os.remove(temp_file)
        
        # Trả về file đã được nạp dữ liệu và xóa tất cả file tạm sau khi gửi
        response = send_file(
            output_filepath,
            as_attachment=True,
            download_name=output_filename,
            mimetype=output_mimetype
        )
        
        # Xóa file output sau khi gửi
        @response.call_on_close
        def cleanup():
            try:
                if os.path.exists(output_filepath):
                    os.remove(output_filepath)
            except Exception:
                pass
        
        return response
        
    except Exception as e:
        # Xử lý lỗi
        return jsonify({'error': f'Lỗi khi xử lý file: {str(e)}'}), 500

@app.route('/clear-uploads', methods=['POST'])
def clear_uploads():
    """
    Xóa tất cả file và thư mục trong thư mục uploads
    """
    try:
        upload_folder = app.config['UPLOAD_FOLDER']
        
        # Kiểm tra xem thư mục có tồn tại không
        if not os.path.exists(upload_folder):
            return jsonify({'message': 'Thư mục uploads không tồn tại'}), 200
        
        # Đếm số file và thư mục đã xóa
        deleted_count = 0
        
        # Duyệt qua tất cả file và thư mục trong uploads
        for item in os.listdir(upload_folder):
            item_path = os.path.join(upload_folder, item)
            
            try:
                if os.path.isfile(item_path):
                    # Xóa file
                    os.remove(item_path)
                    deleted_count += 1
                elif os.path.isdir(item_path):
                    # Xóa thư mục và tất cả nội dung bên trong
                    import shutil
                    shutil.rmtree(item_path)
                    deleted_count += 1
            except Exception as e:
                print(f"Không thể xóa {item_path}: {str(e)}")
        
        return jsonify({
            'success': True,
            'message': f'Đã xóa thành công {deleted_count} file/thư mục',
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
