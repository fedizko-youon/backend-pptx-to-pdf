# API de Edições PPTX

API simples em python com FASTAPI, que permite realizar edições de arquivos powerpoint baseado em um JSON.

## Funcionalidades

- Faz edições em powerpoints enviados por requisição do usuário. A aplicação é ideal para realizar modificações em arquivos power point **mantendo sua formatação**, de forma rápida e em alta quantidade de alterações de texto.

## Pré-requisitos

- Python 3.8+
- **pip** (gerenciador de pacotes do python)
- `fastapi`: O framework web para construir a API
- `python-pptx`: Biblioteca para ler e escrever arquivos `.pptx`
- `uvicorn`: Servidor ASGI para rodar aplicações FastAPI
- `python-multipart`: Necessário para lidar com uploads de arquivos (formulários multipart)

**Para instalar todas elas:**

```bash
pip install fastapi python-pptx uvicorn python-multipart
```

## Como Rodar a Aplicação

Após instalar as dependências, você pode iniciar o servidor Uvicorn:

```bash
uvicorn main:app --host 0.0.0.0 --port 8000 --reload
```

A API estará disponível em http://127.0.0.1:8000 (ou http://localhost:8000).

## Endpoints da API (Como Usar a API)

`POST/editar`

Este endpoint recebe um **JSON**  com as substituições desejadas e um **arquivo PowerPoint** que será editado, processa o arquivo e retorna uma versão modificada dele.

- **Método**: POST
- **Headers**: Form Data (`Content-Type`: `multipart/form-data`)
- **Parâmetros do Formulário (Form Data)**:
  - `file` (tipo `File`): O arquivo .pptx que você deseja editar.
  - `substituicoes_json` (tipo `Form` - string JSON): Uma string JSON contendo os pares chave-valor para as substituições.
    As chaves devem ser os textos a serem encontrados no PowerPoint e os valores serão os textos pelos quais eles serão substituídos.
- **Exempo de `substituicoes_json`**:
  
  ```json
  {
    "{{NOME_CLIENTE}}": "Empresa Exemplo Ltda.",
    "{{DATA_ATUAL}}": "28 de Maio de 2025",
    "{{VALOR_PROPOSTA}}": "R$ 15.000,00",
    "{{CONTATO_RESPONSAVEL}}": "João da Silva"
  }
  ```
- **Observação**: O código sanitiza o nome_cliente para o nome do arquivo de saída, então caracteres especiais serão removidos.
    - A API utiliza arquivos temporários para processar o PowerPoint, que são automaticamente removidos após a resposta.
