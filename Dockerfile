FROM python:3.11-slim

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

# La app lee PORT del entorno (por defecto 80 si no se especifica).
EXPOSE 80

CMD ["python", "app.py"]
