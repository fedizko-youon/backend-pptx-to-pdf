# Use uma imagem base do Python (ajuste a versão conforme necessário)
FROM python:3.9-slim

# Instalar o LibreOffice e dependências do sistema
RUN apt-get update && \
    apt-get install -y libreoffice && \
    apt-get clean && \
    rm -rf /var/lib/apt/lists/*

# Instalar as dependências do Python
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copiar o código da aplicação para dentro do contêiner
COPY . /app
WORKDIR /app

# Expor a porta para o servidor FastAPI
EXPOSE 8000

# Comando para rodar a aplicação FastAPI
CMD ["uvicorn", "main:app", "--host", "0.0.0.0", "--port", "8000"]
