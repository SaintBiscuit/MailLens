# Базовый образ с Python
FROM python:3.11-slim

# Устанавливаем системные зависимости
RUN apt-get update && apt-get install -y \
    git \
    libglib2.0-0 \
    libsm6 \
    libgl1-mesa-glx \
    && rm -rf /var/lib/apt/lists/*

# Рабочая папка
WORKDIR /app

# Копируем всё
COPY . .

# Устанавливаем зависимости
RUN pip install --no-cache-dir -r requirements.txt

# Порт для Streamlit
EXPOSE 8501

# Запуск
CMD ["streamlit", "run", "frontend/app.py", "--server.port=8501", "--server.address=0.0.0.0"]