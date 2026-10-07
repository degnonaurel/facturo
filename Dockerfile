# Facturo — image de déploiement (Hugging Face Spaces, Render, etc.)
FROM python:3.11-slim

# Dépendances système minimales (rendu PDF, polices)
RUN apt-get update && apt-get install -y --no-install-recommends \
        libglib2.0-0 libgl1 \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY facture_vers_excel.py interface_web.py releves_bancaires.py ./

# Le port par défaut de Hugging Face Spaces est 7860 ; Render fournit $PORT.
ENV PORT=7860
EXPOSE 7860

CMD ["sh", "-c", "uvicorn interface_web:app --host 0.0.0.0 --port ${PORT:-7860}"]
