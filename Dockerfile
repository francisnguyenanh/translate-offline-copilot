# Sử dụng Python phiên bản nhẹ
FROM python:3.9-slim

# Tạo thư mục làm việc trong Docker
WORKDIR /app

# Copy file requirements và cài đặt thư viện
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy toàn bộ code vào (Chỉ dùng khi build chính thức)
COPY . .

# Chạy Flask (hoặc FastAPI)
# Host 0.0.0.0 là bắt buộc để có thể truy cập từ bên ngoài Docker
CMD ["python", "app.py"]