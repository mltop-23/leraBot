# Используйте официальный образ Python
FROM python:3.8

# Установка рабочей директории внутри контейнера
WORKDIR /app

# Копирование зависимостей и кода бота в контейнер
COPY requirements.txt /app/
COPY main.py /app/
COPY example.xlsx /app/
# Установка зависимостей
RUN pip install --no-cache-dir openpyxl telebot
RUN pip install --upgrade openpyxl

# Запуск бота при старте контейнера
CMD ["python", "main.py"]