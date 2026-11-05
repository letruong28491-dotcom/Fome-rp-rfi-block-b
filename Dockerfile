FROM python:3.11-slim

WORKDIR /app

# Cài đặt thư viện phụ thuộc
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Sao chép toàn bộ mã nguồn và biểu mẫu
COPY . .

# Thiết lập biến môi trường cho Cloud/Docker
ENV FOME_WEB_HOST=0.0.0.0
ENV PORT=8765

EXPOSE 8765

CMD ["python", "fome_web.py"]
