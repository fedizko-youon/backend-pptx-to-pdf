from fastapi import FastAPI, HTTPException, BackgroundTasks, UploadFile, File, Form
from fastapi.responses import FileResponse
from pptx import Presentation
from pptx.shapes.graphfrm import GraphicFrame
import comtypes.client
import uuid
import os
import platform
import json

app = FastAPI()

# Substituição nos shapes
def substituir_texto_em_shape(shape, substituicoes):
    if shape.has_text_frame:
        for paragraph in shape.text_frame.paragraphs:
            for run in paragraph.runs:
                texto_original = run.text
                texto_limpo = texto_original.strip()
                for chave, valor in substituicoes.items():
                    if chave in texto_limpo:
                        run.text = texto_original.replace(chave, valor).strip()
    elif isinstance(shape, GraphicFrame) and shape.has_table:
        table = shape.table
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.text_frame.paragraphs:
                    for run in paragraph.runs:
                        texto_original = run.text
                        texto_limpo = texto_original.strip()
                        for chave, valor in substituicoes.items():
                            if chave in texto_limpo:
                                run.text = texto_original.replace(chave, valor).strip()


# Gera novo pptx com substituições
def substituir_em_apresentacao(caminho_entrada, caminho_saida, substituicoes):
    prs = Presentation(caminho_entrada)
    for slide in prs.slides:
        for shape in slide.shapes:
            substituir_texto_em_shape(shape, substituicoes)
    prs.save(caminho_saida)


# Converte pptx -> pdf usando PowerPoint via COM (Windows only)
def converter_para_pdf(pptx_path):
    if platform.system() != "Windows":
        raise Exception("Conversão via PowerPoint só é suportada no Windows.")

    powerpoint = comtypes.client.CreateObject("PowerPoint.Application")
    powerpoint.Visible = 1

    caminho_absoluto = os.path.abspath(pptx_path)
    caminho_pdf = caminho_absoluto.replace(".pptx", ".pdf")

    presentation = powerpoint.Presentations.Open(caminho_absoluto, WithWindow=False)
    presentation.SaveAs(caminho_pdf, 32)  # 32 = PDF
    presentation.Close()
    powerpoint.Quit()

    return caminho_pdf


# Limpa arquivos depois da resposta
def remover_arquivos(*caminhos):
    for caminho in caminhos:
        if os.path.exists(caminho):
            os.remove(caminho)


# Novo endpoint para receber o arquivo .pptx e o JSON com substituições
@app.post("/editar/")
async def editar_pptx_upload(
    background_tasks: BackgroundTasks,
    file: UploadFile = File(...),
    substituicoes_json: str = Form(...)
):
    try:
        # Lê e salva o arquivo recebido
        pptx_filename = f"upload_{uuid.uuid4()}.pptx"
        with open(pptx_filename, "wb") as f:
            f.write(await file.read())

        # Converte string JSON para dicionário
        try:
            substituicoes = json.loads(substituicoes_json)
        except json.JSONDecodeError:
            raise HTTPException(status_code=400, detail="JSON de substituições inválido.")

        # Gera novo arquivo pptx com substituições
        pptx_editado = f"editado_{uuid.uuid4()}.pptx"
        substituir_em_apresentacao(pptx_filename, pptx_editado, substituicoes)

        # Converte para PDF
        pdf_path = converter_para_pdf(pptx_editado)

        # Adiciona arquivos temporários para limpeza
        background_tasks.add_task(remover_arquivos, pptx_filename, pptx_editado, pdf_path)

        nome_cliente = substituicoes.get("nome_cliente", "Cliente")  # fallback caso não venha no JSON
        nome_cliente_sanitizado = "".join(c for c in nome_cliente if c.isalnum() or c in (" ", "_", "-")).strip()
        nome_pdf = f"Proposta Comercial {nome_cliente_sanitizado}.pdf"

        return FileResponse(pdf_path, filename=nome_pdf, media_type="application/pdf")

    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Erro: {str(e)}")