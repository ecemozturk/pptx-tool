FROM python:3.12-slim
RUN apt-get update && apt-get install -y --no-install-recommends libreoffice-impress fonts-dejavu-core && rm -rf /var/lib/apt/lists/*
WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY converter.py web.py ./
EXPOSE 8000
CMD ["uvicorn", "web:app", "--host", "0.0.0.0", "--port", "8000", "--workers", "1"]
