FROM python:3.11-slim

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

ENV PORT=8080
EXPOSE ${PORT}

CMD gunicorn app:app --bind 0.0.0.0:${PORT} --timeout 120 --workers 2
