import pdfplumber
import re
import pandas as pd
import pytesseract
import os
from pdf2image import convert_from_path
from PIL import Image

# Configurações
CAMINHO_PDFS = "/root/dev/PdfOcrProject/CRECHE"
DADOS_EXTRAIDOS = "dados_recibos.csv"

"""
DADOS:

Nome pagante
Nome aluno
Mês/competência
Valor do recibo

Número CNPJ
Nome Escola
Atividade Econômica
Educação Infantil
Data de emissão do documento

"""

def extrair_texto_com_ocr(imagem_path):
    return pytesseract.image_to_string(imagem_path, lang='por')

def processar_recibo(texto):
    dados = {}
    
    # Expressões regulares robustas (ignoram variações ortográficas)
    padrao_aluno = r"(?i)(?:aluno|aluna)[\s:]*([^\n]+)"
    padrao_serie = r"(?i)s[ée]rie[\s:]*([^\n]+)"
    padrao_turno = r"(?i)turno[\s:]*([^\n]+)"
    padrao_pagante = r"(?i)recebemos do\(a\):?\s*([^\n]+)"
    padrao_mes = r"(?i)m[êe]s:?\s*([^\n]+)"
    padrao_cnpj = r"\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}"
    padrao_cep = r"\d{2}\.\d{3}-\d{3}"
    
    # Extração com fallbacks
    dados['aluno'] = re.search(padrao_aluno, texto).group(1).strip() if re.search(padrao_aluno, texto) else "N/A"
    dados['serie'] = re.search(padrao_serie, texto).group(1).strip() if re.search(padrao_serie, texto) else "N/A"
    dados['turno'] = re.search(padrao_turno, texto).group(1).strip() if re.search(padrao_turno, texto) else "N/A"
    dados['pagante'] = re.search(padrao_pagante, texto).group(1).strip() if re.search(padrao_pagante, texto) else "N/A"
    dados['mes'] = re.search(padrao_mes, texto).group(1).strip() if re.search(padrao_mes, texto) else "N/A"
    dados['cnpj'] = re.search(padrao_cnpj, texto).group(0) if re.search(padrao_cnpj, texto) else "N/A"
    dados['cep'] = re.search(padrao_cep, texto).group(0) if re.search(padrao_cep, texto) else "N/A"
    
    # Extração de endereço (busca contextual)
    linhas = texto.split('\n')
    endereco = ""
    for i, linha in enumerate(linhas):
        if "AV." in linha and "CEP" in linhas[i+1]:
            endereco = linha.strip() + " " + linhas[i+1].strip()
            break
    dados['endereco'] = endereco if endereco else "N/A"
    
    return dados

# Processamento principal
resultados = []
for arquivo in os.listdir(CAMINHO_PDFS):
    if arquivo.endswith(".pdf"):
        # Converter PDF para imagem
        imagens = convert_from_path(os.path.join(CAMINHO_PDFS, arquivo))
        texto_completo = ""
        
        for img in imagens:
            texto_completo += extrair_texto_com_ocr(img)
        
        dados = processar_recibo(texto_completo)
        dados['arquivo'] = arquivo
        resultados.append(dados)

# Exportar para CSV
pd.DataFrame(resultados).to_csv(DADOS_EXTRAIDOS, index=False)
