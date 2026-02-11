# Excel Translation Manager

Ứng dụng web Flask để quản lý việc trích xuất và nạp bản dịch cho file Excel (.xlsx).

## Tính năng

### 1. Trích xuất (Extract)
- Upload file Excel gốc
- Tự động trích xuất tất cả các cell chứa text (bỏ qua số và công thức)
- Tạo file JSON với format: `{"SheetName!CellCoordinate": "Content"}`
- Tải về file `to_translate.json`

### 2. Nạp bản dịch (Inject)
- Upload file Excel gốc và file JSON đã dịch
- Tự động nạp bản dịch vào đúng vị trí trong Excel
- Giữ nguyên định dạng, màu sắc ban đầu
- Tải về file `output_translated.xlsx`

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

1. Chọn chức năng "Extract" ở bên trái
2. Upload file Excel (.xlsx) cần dịch
3. Nhấn "Trích xuất và Tải về JSON"
4. File `to_translate.json` sẽ được tải về
5. Dịch nội dung trong file JSON (có thể dùng công cụ dịch tự động hoặc thủ công)

### Nạp bản dịch vào Excel

1. Chọn chức năng "Inject" ở bên phải
2. Upload file Excel gốc
3. Upload file JSON đã dịch
4. Nhấn "Nạp bản dịch và Tải về Excel"
5. File `output_translated.xlsx` sẽ được tải về với nội dung đã dịch

## Lưu ý

- Hỗ trợ định dạng file: `.xlsx` (Excel 2007+)
- Hỗ trợ encoding UTF-8 cho tiếng Việt và tiếng Nhật
- Giữ nguyên định dạng, màu sắc, font chữ của file Excel gốc
- Chỉ trích xuất các cell chứa text, bỏ qua:
  - Các cell chứa số
  - Các cell chứa công thức (bắt đầu với '=')
  - Các cell trống

## Yêu cầu hệ thống

- Python 3.7+
- Flask 3.0.0
- openpyxl 3.1.2

## Tác giả

Được phát triển bằng Flask và openpyxl.
