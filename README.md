# Excel, PowerPoint, Word & PDF Translation Manager

Ứng dụng web Flask để quản lý việc trích xuất và nạp bản dịch cho file Excel (.xlsx), PowerPoint (.pptx), Word (.docx) và PDF (.pdf).

## Tính năng

### 1. Trích xuất (Extract)
- Upload file Excel, PowerPoint, Word hoặc PDF gốc
- **Excel**: Tự động trích xuất tất cả các cell chứa text (bỏ qua số và công thức)
- **PowerPoint**: Trích xuất text từ shapes, text boxes, tables, và cả **grouped shapes (shapes lồng nhau)**
- **Word**: Trích xuất text từ paragraphs, tables, headers và footers
- **PDF**: Tự động chuyển đổi sang Word rồi trích xuất (giữ được phần lớn format, màu sắc, bảng)
- Tạo file JSON với format:
  - Excel: `{"SheetName!CellCoordinate": "Content"}`
  - PowerPoint: 
    - Shape thông thường: `{"SlideX!ShapeY": "Content"}`
    - Shape trong group: `{"SlideX!ShapeY_Z": "Content"}` (Z là shape con bên trong)
    - Table cell: `{"SlideX!ShapeY!Table_RxCy": "Content"}`
  - Word:
    - Paragraph: `{"ParagraphX": "Content"}`
    - Table cell: `{"TableX!RyCz": "Content"}`
    - Header: `{"Header_SectionX!ParagraphY": "Content"}`
    - Footer: `{"Footer_SectionX!ParagraphY": "Content"}`
- Tải về file ZIP chứa nhiều file JSON (mỗi file 400 cặp key-value)

### 2. Nạp bản dịch (Inject)
- Upload file gốc và file JSON đã dịch (hoặc file ZIP chứa nhiều file JSON)
- Tự động nạp bản dịch vào đúng vị trí
- Giữ nguyên định dạng, màu sắc ban đầu
- **PDF**: Sẽ được chuyển sang Word, nạp bản dịch, và trả về file Word đã dịch (không thể nạp lại vào PDF)
- Tải về file `output_translated.xlsx`, `output_translated.pptx` hoặc `output_translated.docx`

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

1. Ở "Bước 1": Upload file Excel (.xlsx), PowerPoint (.pptx), Word (.docx) hoặc PDF (.pdf) cần dịch
2. Ở "Bước 2": Nhấn "Trích xuất và Tải về JSON"
3. File ZIP chứa nhiều file JSON sẽ được tải về
4. Giải nén ZIP và dịch từng file JSON (có thể dùng AI như ChatGPT, Copilot)
5. Sử dụng prompt AI được cung cấp trong giao diện để dịch chính xác

### Nạp bản dịch vào file gốc

1. Ở "Bước 1": Upload file gốc (Excel, PowerPoint, Word hoặc PDF)
2. Ở "Bước 3": Upload các file JSON đã dịch (hoặc file ZIP chứa các file đã dịch)
3. Nhấn "Nạp bản dịch và Tải về"
4. File đã dịch sẽ được tải về với nội dung đã được cập nhật
5. **Lưu ý**: Nếu upload file PDF, bạn sẽ nhận lại file Word (.docx) đã dịch

## Lưu ý

- Hỗ trợ định dạng file: `.xlsx` (Excel 2007+), `.pptx` (PowerPoint 2007+), `.docx` (Word 2007+) và `.pdf`
- Hỗ trợ encoding UTF-8 cho tiếng Việt và tiếng Nhật
- Giữ nguyên định dạng, màu sắc, font chữ của file gốc
- **Excel**: Chỉ trích xuất các cell chứa text, bỏ qua số và công thức (bắt đầu với '=')
- **PowerPoint**: 
  - Trích xuất text từ tất cả shapes có text và table cells
  - Hỗ trợ **grouped shapes** - tự động trích xuất text từ shapes lồng bên trong shapes khác
  - Nested shapes được đánh dấu bằng underscore: `Shape2_3` có nghĩa là shape con thứ 3 bên trong shape 2
- **Word**:
  - Trích xuất text từ paragraphs, tables, headers và footers
  - Giữ nguyên định dạng (font, size, bold, italic, color...)
  - Hỗ trợ nhiều section với headers/footers khác nhau
- **PDF**:
  - Tự động chuyển đổi sang Word bằng thư viện `pdf2docx` (miễn phí)
  - Giữ được phần lớn format, màu sắc, bảng (không hoàn hảo 100% với PDF phức tạp)
  - File đầu ra sẽ là Word (.docx), không thể nạp lại vào PDF
  - Nên kiểm tra lại file Word sau khi chuyển đổi
- File JSON được tách thành nhiều file nhỏ (400 cặp/file) để dễ dàng xử lý với AI

## Yêu cầu hệ thống

- Python 3.7+
- Flask 3.0.0
- openpyxl 3.1.2
- python-pptx 0.6.23
- python-docx 1.1.0
- pdf2docx 0.5.8

## Tác giả
, openpyxl, và python-pptx
Được phát triển bằng Flask và openpyxl.
