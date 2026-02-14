# Excel & PowerPoint Translation Manager

Ứng dụng web Flask để quản lý việc trích xuất và nạp bản dịch cho file Excel (.xlsx) và PowerPoint (.pptx).

## Tính năng

### 1. Trích xuất (Extract)
- Upload file Excel hoặc PowerPoint gốc
- **Excel**: Tự động trích xuất tất cả các cell chứa text (bỏ qua số và công thức)
- **PowerPoint**: Trích xuất text từ shapes, text boxes, tables, và cả **grouped shapes (shapes lồng nhau)**
- Tạo file JSON với format:
  - Excel: `{"SheetName!CellCoordinate": "Content"}`
  - PowerPoint: 
    - Shape thông thường: `{"SlideX!ShapeY": "Content"}`
    - Shape trong group: `{"SlideX!ShapeY_Z": "Content"}` (Z là shape con bên trong)
    - Table cell: `{"SlideX!ShapeY!Table_RxCy": "Content"}`
- Tải về file ZIP chứa nhiều file JSON (mỗi file 400 cặp key-value)

### 2. Nạp bản dịch (Inject)
- Upload file gốc và file JSON đã dịch (hoặc file ZIP chứa nhiều file JSON)
- Tự động nạp bản dịch vào đúng vị trí
- Giữ nguyên định dạng, màu sắc ban đầu
- Tải về file `output_translated.xlsx` hoặc `output_translated.pptx`

## Cài đặt

### 1. Cài đặt các thư viện cần thiết

```bash
pip install -r requirements.txt
```

### 2. Chạy ứng dụng


```bash
python app.py
```

### 3. Mở trình duyệt

Truy cập: `http://localhost:5000`

## Cấu trúc thư mục

```
translate-offline-copilot/
│
├── app.py                 # Logic Flask chính
├── requirements.txt       # Các thư viện cần thiết
├── templates/            
│   └── index.html        # Giao diện web
└── uploads/              # Thư mục tạm để xử lý file
```

## Hướng dẫn sử dụng

### Trích xuất nội dung cần dịch

1. Ở "Bước 1": Upload file Excel (.xlsx) hoặc PowerPoint (.pptx) cần dịch
2. Ở "Bước 2": Nhấn "Trích xuất và Tải về JSON"
3. File ZIP chứa nhiều file JSON sẽ được tải về
4. Giải nén ZIP và dịch từng file JSON (có thể dùng AI như ChatGPT, Copilot)
5. Sử dụng prompt AI được cung cấp trong giao diện để dịch chính xác

### Nạp bản dịch vào file gốc

1. Ở "Bước 1": Upload file gốc (Excel hoặc PowerPoint)
2. Ở "Bước 3": Upload các file JSON đã dịch (hoặc file ZIP chứa các file đã dịch)
3. Nhấn "Nạp bản dịch và Tải về"
4. File đã dịch sẽ được tải về với nội dung đã được cập nhật

## Lưu ý

- Hỗ trợ định dạng file: `.xlsx` (Excel 2007+) và `.pptx` (PowerPoint 2007+)
- Hỗ trợ encoding UTF-8 cho tiếng Việt và tiếng Nhật
- Giữ nguyên định dạng, màu sắc, font chữ của file gốc
- **Excel**: Chỉ trích xuất các cell chứa text, bỏ qua số và công thức (bắt đầu với '=')
- **PowerPoint**: 
  - Trích xuất text từ tất cả shapes có text và table cells
  - Hỗ trợ **grouped shapes** - tự động trích xuất text từ shapes lồng bên trong shapes khác
  - Nested shapes được đánh dấu bằng underscore: `Shape2_3` có nghĩa là shape con thứ 3 bên trong shape 2
- File JSON được tách thành nhiều file nhỏ (400 cặp/file) để dễ dàng xử lý với AI

## Yêu cầu hệ thống

- Python 3.7+
- Flask 3.0.0
- openpyxl 3.1.2
- python-pptx 0.6.23

## Tác giả
, openpyxl, và python-pptx
Được phát triển bằng Flask và openpyxl.
