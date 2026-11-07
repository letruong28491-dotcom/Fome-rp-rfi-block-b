FROM python:3.11-slim

WORKDIR /app

# Cài đặt LibreOffice Calc headless và font chữ tiếng Việt cho xuất PDF chuẩn xác 100%
RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice-calc \
    libreoffice-writer \
    fonts-dejavu \
    fonts-liberation \
    fontconfig \
    libgl1 \
    libglib2.0-0 \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/*

# Cài đặt thư viện Python phụ thuộc
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Sao chép toàn bộ mã nguồn và biểu mẫu
COPY . .

# Thiết lập biến môi trường cho Cloud/Docker
ENV FOME_WEB_HOST=0.0.0.0
ENV PORT=8765
ENV SAL_USE_VCLPLUGIN=gen

EXPOSE 8765

CMD ["python", "fome_web.py"]
