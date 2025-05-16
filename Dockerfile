# Use uma imagem base do Python
FROM python:3.11-slim

# Instala as dependências do LibreOffice
RUN apt-get update && apt-get install -y \
    libreoffice \
    && apt-get clean

# Instala as dependências do Python
WORKDIR /app
COPY . .

RUN pip install --no-cache-dir -r requirements.txt

EXPOSE 8000

CMD ["uvicorn", "main:app", "--host", "0.0.0.0", "--port", "8000"]
