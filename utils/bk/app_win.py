# teste versão com windows e Word
from fastapi import FastAPI
from fastapi.responses import FileResponse, JSONResponse
from pydantic import BaseModel
from docx import Document
from docx2pdf import convert
from num2words import num2words
from datetime import datetime
import os
import uuid
import tempfile
import locale

# Força a localidade para português (para formatar data e número corretamente)
try:
    locale.setlocale(locale.LC_TIME, 'pt_BR.utf8')
except locale.Error:
    pass  # fallback se o sistema não suportar pt_BR

app = FastAPI()

TEMPLATE_MAP = {
    "autorizacao_corretor": "templates/autorizacao-de-venda-corretor-3.docx",
    "autorizacao_imobiliaria": "templates/autorizacao-de-venda-imob.docx"
}

class PayloadAutorizacao(BaseModel):
    vendedor: str
    cpf_mask: str
    razao_corretor: str
    cnpj_mask_corretor: str
    creci_corretor: str
    cartorio_number: str
    mat_number: str
    valor: float
    corretagem_number: int
    pendencia: bool
    pendencia_texto: str = ""
    tipo_template: str  # exemplo: "autorizacao_corretor" ou "autorizacao_imobiliaria"

def gerar_data_extenso():
    meses = {
        1: "janeiro", 2: "fevereiro", 3: "março", 4: "abril",
        5: "maio", 6: "junho", 7: "julho", 8: "agosto",
        9: "setembro", 10: "outubro", 11: "novembro", 12: "dezembro"
    }
    hoje = datetime.now()
    return f"{hoje.day} de {meses[hoje.month]} de {hoje.year}"

def gerar_texto_4_autorizacao(pendencia: bool, pendencia_texto: str):
    if pendencia and pendencia_texto.strip():
        return f"""O CONTRATANTE declara que o imóvel se encontra livre e desembaraçado de todos e
quaisquer ônus judicial, extrajudicial, hipoteca legal ou convencional, foro ou pensão e está quite
com todos os impostos, taxas, inclusive contribuições condominiais, se houver, até a presente
data, à exceção de {pendencia_texto}."""
    else:
        return """O CONTRATANTE declara que o imóvel se encontra livre e desembaraçado de todos e
quaisquer ônus judicial, extrajudicial, hipoteca legal ou convencional, foro ou pensão e está quite
com todos os impostos, taxas, inclusive contribuições condominiais, se houver, até a presente
data, sem exceção."""

@app.post("/gerar-pdf/autorizacao")
def gerar_pdf_autorizacao(dados: PayloadAutorizacao):
    if dados.tipo_template not in TEMPLATE_MAP:
        return JSONResponse(status_code=400, content={"status": "erro", "mensagem": "Tipo de template inválido."})

    template_path = TEMPLATE_MAP[dados.tipo_template]

    valor_mask_brl = f"{dados.valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    corretagem_text = num2words(dados.corretagem_number, lang='pt_BR')
    data_completa = gerar_data_extenso()
    text_4 = gerar_texto_4_autorizacao(dados.pendencia, dados.pendencia_texto)

    variaveis = {
        "vendedor": dados.vendedor,
        "cpf_mask": dados.cpf_mask,
        "razao_corretor": dados.razao_corretor,
        "cnpj_mask_corretor": dados.cnpj_mask_corretor,
        "creci_corretor": dados.creci_corretor,
        "cartorio_number": dados.cartorio_number,
        "mat_number": dados.mat_number,
        "valor_mask_brl": valor_mask_brl,
        "corretagem_number": str(dados.corretagem_number),
        "corretagem_text": corretagem_text,
        "data_completa": data_completa,
        "text_4": text_4
    }

    temp_dir = tempfile.gettempdir()
    unique_id = str(uuid.uuid4())
    docx_path = os.path.join(temp_dir, f"{unique_id}.docx")
    pdf_path = os.path.join(temp_dir, f"{unique_id}.pdf")

    doc = Document(template_path)
    for p in doc.paragraphs:
        for k, v in variaveis.items():
            chave = f"{{{k}}}"
            if chave in p.text:
                p.text = p.text.replace(chave, v)

    doc.save(docx_path)
    convert(docx_path)

    pdf_name = os.path.basename(pdf_path)
    return {
        "status": "sucesso",
        "tipo": "autorizacao-de-venda",        
        "pdf_name": pdf_name,
        "pdf_url": f"http://localhost:8000/download/{pdf_name}"
    }

@app.get("/download/{pdf_name}")
def baixar_pdf(pdf_name: str):
    pdf_path = os.path.join(tempfile.gettempdir(), pdf_name)
    if not os.path.exists(pdf_path):
        return JSONResponse(status_code=404, content={"status": "erro", "mensagem": "Arquivo não encontrado."})
    return FileResponse(pdf_path, media_type="application/pdf", filename=pdf_name)
