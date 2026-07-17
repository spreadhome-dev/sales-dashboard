FROM python:3.9-slim

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

EXPOSE 8050

CMD ["gunicorn", "--worker-class", "sync", "--timeout", "300", "--workers", "1", "--bind", "0.0.0.0:8050", "sales_dashboard_live:app"]