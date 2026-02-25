# Hướng dẫn Sử dụng Hệ thống Dịch thuật Offline

Chào mừng bạn đến với hệ thống quản lý trích xuất và nạp bản dịch cho các tài liệu Excel, PowerPoint và Word. Hệ thống này giúp bạn dịch thuật tài liệu một cách chuyên nghiệp bằng cách kết hợp sức mạnh của AI với quy trình kiểm soát nội dung chính xác.

## 1. Đăng nhập
Khi truy cập vào hệ thống, bạn cần nhập mật khẩu để bắt đầu sử dụng.
- **Mật khẩu mặc định**: Vui lòng kiểm tra file `password.txt` hoặc liên hệ quản trị viên (Thường là `admin123`).

## 2. Quy trình dịch thuật (4 Bước)

Hệ thống hoạt động theo một quy trình khép kín giúp đảm bảo định dạng file gốc không bị thay đổi.

### Bước 1: Tải lên tài liệu
- Tại khung **"Bước 1: Tải lên tài liệu"**, hãy kéo thả hoặc chọn file từ máy tính của bạn.
- Hệ thống hỗ trợ các định dạng: **.xlsx** (Excel), **.pptx** (PowerPoint), và **.docx** (Word).
- Nhấn **"Tải lên và Trích xuất"**.

### Bước 2: Sao chép dữ liệu và Dịch thuật với AI
- Sau khi tải lên, hệ thống sẽ phân tích văn bản có trong file.
- **Lựa chọn ngôn ngữ & Phong cách:**
    1. Chọn **Ngôn ngữ đích** (Target Language) muốn dịch sang.
    2. Chọn **Mẫu câu lệnh (Prompt Template)** phù hợp (Chuyên nghiệp, Tự nhiên, hoặc Kỹ thuật).
- **Sao chép nội dung:**
    - **ƯU TIÊN:** Hãy sử dụng các phân đoạn trong mục **"Dữ liệu thu gọn (Dedup)"**. Đây là những nội dung đã loại bỏ câu trùng lặp, giúp tiết kiệm thời gian và số lượng từ.
    - Nhấn nút **"Copy"** cho từng phần (ví dụ: `dedup_part01.json`). 
    - **Lưu ý:** Nút Copy này đã bao gồm cả **Câu lệnh hướng dẫn (Prompt)** và **Nội dung JSON** trong bộ nhớ tạm.
- **Dịch với AI (Copilot/ChatGPT):**
    - Mở cửa sổ AI, dán nội dung vừa copy (Ctrl+V) và gửi cho AI.
    - Đợi AI xử lý và trả về khối JSON đã được dịch (chỉ dịch các giá trị, giữ nguyên các Key).

### Bước 3: Nạp lại bản dịch
- Quay lại hệ thống, dán nội dung mà AI vừa trả về vào khung ở **"Bước 3: Nạp bản dịch"**.
- Nhấn **"Kiểm tra và Nạp bản dịch"**.
- Hệ thống sẽ xác nhận nếu cấu trúc JSON hợp lệ và lưu vào bộ nhớ tạm.

### Bước 4: Tạo file và Tải về
- Sau khi đã nạp đủ các phần cần dịch, nhấn nút **"Tạo file kết quả"**.
- Hệ thống sẽ tự động ghép các phần dịch vào đúng vị trí trong file gốc.
- Nhấn **"Tải file kết quả"** để lưu file đã dịch về máy.

## 3. Các tính năng nâng cao

### Khử trùng lặp (Dedup)
Tính năng này tự động tìm các câu/cụm từ giống nhau và chỉ yêu cầu bạn dịch một lần duy nhất. 
- Giúp giảm đáng kể số lượng từ cần dịch.
- Tiết kiệm chi phí và thời gian khi làm việc với AI.
- Đảm bảo tính nhất quán của thuật ngữ trong toàn bộ tài liệu.

### Mẫu câu lệnh AI (Prompt Templates)
Bạn có thể tùy chỉnh cách AI dịch thuật thông qua các mẫu câu lệnh:
- **Dịch chính xác (Formal)**: Phù hợp cho hợp đồng, báo cáo kinh doanh.
- **Dịch tự nhiên (Casual)**: Phù hợp cho giao tiếp thông thường, UI ứng dụng.
- **Dịch kỹ thuật (Technical)**: Giữ nguyên các thuật ngữ chuyên ngành IT.

## 4. Lưu ý quan trọng
- **Cấu trúc JSON**: Tuyệt đối không thay đổi các "Key" (phần bên trái) trong file JSON khi dịch, AI chỉ được phép dịch phần "Value" (phần bên phải).
- **Kích thước file**: Hệ thống giới hạn tải lên tối đa là 50MB.
- **Bảo mật**: Các phiên làm việc (session) cũ sẽ tự động được xóa sau mỗi ngày để bảo vệ dữ liệu của bạn.
- **Dấu ngoặc kép**: AI nên sử dụng dấu ngoặc kép chuẩn `"` thay vì các dấu ngoặc kép đặc biệt khác để tránh lỗi định dạng JSON.

---
*Hệ thống được phát triển bởi Tomo Translation.*
*Hỗ trợ: [eikitomobe@gmail.com](mailto:eikitomobe@gmail.com)*
