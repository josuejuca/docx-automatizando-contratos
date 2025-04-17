# teste versão com linux e libreoffice
from fastapi import FastAPI
from fastapi.responses import FileResponse, JSONResponse
from pydantic import BaseModel
from docx import Document
from num2words import num2words
from datetime import datetime
import os
import uuid
import tempfile
import locale
import subprocess
from fastapi.middleware.cors import CORSMiddleware # CORS
from contrato_de_corretagem import preencher_contrato
from typing import List, Optional, Dict
from declaracao_de_visita import preencher_declaracao_visita
# Força a localidade para português (para formatar data e número corretamente)
try:
    locale.setlocale(locale.LC_TIME, 'pt_BR.utf8')
except locale.Error:
    pass

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # Permite todas as origens (você pode restringir se necessário)
    allow_credentials=True,
    allow_methods=["*"],  # Permite todos os métodos (GET, POST, etc.)
    allow_headers=["*"],  # Permite todos os cabeçalhos
)

TEMPLATE_MAP = {
    "autorizacao_corretor": "templates/autorizacao-de-venda-corretor.docx",
    "autorizacao_imobiliaria": "templates/autorizacao-de-venda-imob.docx",
    "contrato_corretagem": "templates/contrato-de-corretagem.docx",
}
# Autorizacao de venda
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
    tipo_template: str

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

@app.get("/")
def root():
    return {"message": "Hello Clancy!"}

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

    # Conversão com LibreOffice em modo headless (Linux)
    subprocess.run([
        "libreoffice", "--headless", "--convert-to", "pdf:writer_pdf_Export", "--outdir", temp_dir, docx_path
    ], check=True)

    pdf_name = os.path.basename(pdf_path)
    docx_name = os.path.basename(docx_path)
    return {
        "status": "sucesso",
        "tipo": "autorizacao-de-venda",
        "pdf_name": pdf_name,
        "docx_name": docx_name,
        "docx_url": f"https://docx.imogo.com.br/download/{docx_name}",
        "pdf_url": f"https://docx.imogo.com.br/download/{pdf_name}"
    }

# fim autorizaçao de venda
# Contrato de corretagem
class Contratante(BaseModel):
    nome: str
    email: str
    endereco: str
    cpf: str
    telefone: str
    cidade: str
    cep: str
    uf: str

class Corretor(BaseModel):
    nome: str
    cnpj: str
    endereco: str
    telefone: str
    creci: str
    participacao: float

class Testemunhas(BaseModel):
    nome: str
    rg: str
    cpf: str    

class DadosContrato(BaseModel):
    endereco_imovel: str
    valor_venda: float
    porcentagem_corretagem: float
    contratantes: List[Contratante]
    corretores: List[Corretor]
    testemunhas: List[Testemunhas]

@app.post("/gerar-pdf/contrato-corretagem")
async def gerar_pdf_contrato_corretagem(dados: DadosContrato):
    try:
        # Gera o documento .docx em memória
        buffer = preencher_contrato(
            dados.endereco_imovel,
            [c.dict() for c in dados.contratantes],
            [c.dict() for c in dados.corretores],
            dados.valor_venda,
            dados.porcentagem_corretagem,
            dados.testemunhas,
            modelo_path="templates/contrato-de-corretagem.docx"  # já está no padrão certo
        )

        # Salva o .docx temporariamente
        temp_dir = tempfile.gettempdir()
        unique_id = str(uuid.uuid4())
        docx_path = os.path.join(temp_dir, f"{unique_id}.docx")
        pdf_path = os.path.join(temp_dir, f"{unique_id}.pdf")

        with open(docx_path, "wb") as f:
            f.write(buffer.read())

        # Converte o .docx em PDF usando LibreOffice headless
        subprocess.run([
            "libreoffice", "--headless", "--convert-to", "pdf:writer_pdf_Export", "--outdir", temp_dir, docx_path
        ], check=True)

        pdf_name = os.path.basename(pdf_path)
        docx_name = os.path.basename(docx_path)
        return {
            "status": "sucesso",
            "tipo": "contrato-de-corretagem",
            "pdf_name": pdf_name,
            "docx_name": docx_name,
            "docx_url": f"https://docx.imogo.com.br/download/{docx_name}",
            "pdf_url": f"https://docx.imogo.com.br/download/{pdf_name}"
        }
    except Exception as e:
        return JSONResponse(status_code=500, content={"status": "erro", "mensagem": str(e)})

# Fim contrato de corretagem

# declaração de visita
# Models
class Visitante(BaseModel):
    nome: str
    cpf: str
    email: str
    tel: str

class DeclaracaoVisitaPayload(BaseModel):
    endereco_imovel: str    
    visitantes: List[Visitante]
    nome_corretor: str
    creci_corretor: str
    isImob: bool = False
    name_imob: Optional[str] = None
    number_creci: Optional[str] = None
    avaliacao_nps: Optional[Dict[str, int]] = None

@app.post("/gerar-pdf/declaracao-visita")
async def gerar_pdf_declaracao_visita(dados: DeclaracaoVisitaPayload):
    try:
        # Gera o documento .docx preenchido e salvo temporariamente
        docx_path = preencher_declaracao_visita(dados.dict(), "templates/declaracao-de-visita.docx")

        # Define caminhos temporários
        temp_dir = tempfile.gettempdir()
        unique_id = str(uuid.uuid4())
        new_docx_path = os.path.join(temp_dir, f"{unique_id}.docx")
        pdf_path = os.path.join(temp_dir, f"{unique_id}.pdf")

        # Copia o arquivo gerado
        os.rename(docx_path, new_docx_path)

        # Converte para PDF com LibreOffice
        subprocess.run([
            "libreoffice", "--headless", "--convert-to", "pdf:writer_pdf_Export", "--outdir", temp_dir, new_docx_path
        ], check=True)

        pdf_name = os.path.basename(pdf_path)
        docx_name = os.path.basename(new_docx_path)

        return {
            "status": "sucesso",
            "tipo": "declaracao-de-visita",
            "pdf_name": pdf_name,
            "docx_name": docx_name,
            "docx_url": f"https://docx.imogo.com.br/download/{docx_name}",
            "pdf_url": f"https://docx.imogo.com.br/download/{pdf_name}"
        }

    except Exception as e:
        return JSONResponse(status_code=500, content={
            "status": "erro",
            "mensagem": str(e)
        })
# fim declaração de visita
# Downlaods
@app.get("/download/{pdf_name}")
def baixar_pdf(pdf_name: str):
    pdf_path = os.path.join(tempfile.gettempdir(), pdf_name)
    if not os.path.exists(pdf_path):
        return JSONResponse(status_code=404, content={"status": "erro", "mensagem": "Arquivo não encontrado."})
    return FileResponse(pdf_path, media_type="application/pdf", filename=pdf_name)
