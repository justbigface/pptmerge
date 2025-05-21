FROM python:3.11-slim

WORKDIR /app

COPY . .
RUN pip install --no-cache-dir -r requirements.txt && pip install gunicorn

EXPOSE 8080

ENV PYTHONUNBUFFERED=1

CMD ["gunicorn", "-b", "0.0.0.0:8080", "-w", "2", "--threads", "4", "--graceful-timeout", "60", "--timeout", "120", "app.ppt_merge_service:app"]