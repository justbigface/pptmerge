FROM python:3.11-slim

WORKDIR /app

RUN groupadd -r appuser && useradd -r -g appuser appuser
USER appuser

COPY requirements.txt ./
RUN pip install --no-cache-dir -r requirements.txt && pip install gunicorn

COPY app ./app

RUN chown -R appuser:appuser /app

EXPOSE 8080

ENV PYTHONUNBUFFERED=1

CMD ["gunicorn", "-b", "0.0.0.0:8080", "-w", "2", "--threads", "4", "app.ppt_merge_service:app"]