FROM python:3.12-slim

# poppler-utils даёт pdftoppm (PDF -> PNG)
RUN apt-get update \
    && apt-get install -y --no-install-recommends poppler-utils \
    && rm -rf /var/lib/apt/lists/*

RUN pip install --no-cache-dir flask openpyxl pillow gunicorn

WORKDIR /app
COPY build_parts_table.py app.py ./
RUN mkdir -p /work/input /work/output

EXPOSE 8000
CMD ["gunicorn", "-b", "0.0.0.0:8000", "-w", "2", "-t", "300", "app:app"]
