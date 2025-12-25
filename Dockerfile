# Используем образ с CUDA 13.0
FROM nvidia/cuda:13.1.0-base-ubuntu24.04

# Устанавливаем Python
RUN apt-get update && apt-get install -y \
    python3.11 \
    python3-pip \
    python3.11-venv \
    libmagic1 \
    && rm -rf /var/lib/apt/lists/*

RUN ln -s /usr/bin/python3.11 /usr/bin/python

WORKDIR /app

# Копируем зависимости
COPY requirements.txt .

# Устанавливаем Python-пакеты
RUN pip3 install --no-cache-dir -r requirements.txt

# Проверяем CUDA
RUN python3 -c "import torch; \
    print(f'===== Проверка наличия CUDA =====');\
    print(f'PyTorch: {torch.__version__}'); \
    print(f'CUDA: {torch.version.cuda}'); \
    print(f'CUDA доступна: {torch.cuda.is_available()}')"

# Копируем код
COPY . .

# Порт для Streamlit
EXPOSE 8501

# Запуск
CMD ["streamlit", "run", "frontend/app.py", "--server.port=8501", "--server.address=0.0.0.0"]
