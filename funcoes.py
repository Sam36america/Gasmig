from pdf2image import convert_from_path
from PIL import Image
import pytesseract
import pandas as pd
import numpy as np
import os
import re
from sqlalchemy import create_engine, select, MetaData, Table
from sqlalchemy.orm import sessionmaker
import shutil

#in/out 3100, 1300, 3800, 1450
usuario_conectado = 'samuel.santos'
# Configure o caminho do executável do Tesseract
pytesseract.pytesseract.tesseract_cmd = fr'C:\Users\{usuario_conectado}\AppData\Local\Programs\Tesseract-OCR\tesseract.exe'

def get_first_pdf(folder_path):
    pdf_files = [f for f in os.listdir(folder_path) if f.lower().endswith('.pdf')]
    pdf_files.sort()
    return os.path.join(folder_path, pdf_files[0]) if pdf_files else None

def get_second_pdf(folder_path):
    pdf_files = [f for f in os.listdir(folder_path) if f.lower().endswith('.pdf')]
    pdf_files.sort()
    return os.path.join(folder_path, pdf_files[1]) if len(pdf_files) > 1 else None

def pdf_ocr(image):
    # Select the first page
    config = pytesseract.pytesseract.tesseract_cmd
    
    # Define o idioma para o reconhecimento de texto (por exemplo, português)
    config += r'--oem 3 --psm 6 -l por'
    config += r'--psm 6 outputbase alphanumeric'

    return pytesseract.image_to_string(image,config=config)

def pdf_to_image(pdf_path):

    # Converte o PDF em uma lista de imagens
    images = convert_from_path(pdf_path, 500, poppler_path=r'C:\poppler-0.68.0\bin')
    imagem = images[0]
    #imagem.show()
    return imagem
      
def dados_excel(cnpj, valor_total,volume_total, data_emissao, data_inicio, data_fim, numero_fatura, valor_icms, correcao_pcs, dist, nome_arquivo):
   
    dados = {
           'CNPJ': [cnpj],
           'VALOR TOTAL': valor_total,
           'VOLUME TOTAL': [volume_total],
           'DATA DA EMISSÃO': data_emissao,
           'DATA INICIO': [data_inicio],
           'DATA FIM': [data_fim],     
           'NUMERO FATURA':[numero_fatura],
           'VALOR ICMS': [valor_icms],
           'CORREÇÃO PCS': [correcao_pcs],
           'DISTRIBUIDORA': [dist],
           'NOME DO ARQUIVO': [nome_arquivo]
     }
    try:    
        df = pd.DataFrame(dados)
    except:
        dados = {
           'CNPJ':'CNPJ não encontrado', 
           'VALOR TOTAL': 'valor_total não econtrado',
           'VOLUME TOTAL': 'volume_total não econtrado',
           'DATA DA EMISSÃO': 'data_emissao não econtrado',
           'DATA INICIO': 'data_inicio não econtrado',
           'DATA FIM': 'data_fim não econtrado',     
           'NUMERO FATURA':'numero_fatura não econtrado]',
           'VALOR ICMS': 'valor_icms não econtrado',
           'CORREÇÃO DO PCS': 'correcao_pcs não encontrado',
           'DISTRIBUIDORA': 'Distribuidora não econtrado',
           'NOME ARQUIVO': 'Nome do arquivo não econtrado'
            }
        
        indice = ['1']
        df = pd.DataFrame(dados,index=indice)  
    
    return df
          
def adicionar_dados_excel(dados, novos_dados):
    try:
        df_existente = pd.read_excel(dados)
        
    except FileNotFoundError:
        print(f"O arquivo '{dados}' não foi encontrado. Criando um novo.")
        df_existente = pd.DataFrame()

    df_novos_dados = pd.DataFrame(novos_dados)
    df_resultante = pd.concat([df_existente, df_novos_dados], ignore_index=True)
    df_resultante.to_excel(dados, index=False)

    print(f"Dados adicionados com sucesso na planilha '{dados}'")

def listar_pdfs_com_referencia_na_pasta(pasta, referencia):
    arquivos_pdf = []
    for arquivo in os.listdir(pasta):
        if arquivo.endswith('.pdf'):
            nome_distribuidora = re.findall(r'_GN_([A-ZÁ]+)_',arquivo)
            if nome_distribuidora:
                nome_distribuidora = nome_distribuidora[0]
                
            arquivos_pdf.append(arquivo)
    return arquivos_pdf

def verificar_fatura_existe(session, tabela_faturas, numero_fatura):
    stmt = select([tabela_faturas.c.numero_fatura]).where(tabela_faturas.c.numero_fatura == numero_fatura)
    result = session.execute(stmt).fetchone()
    return result is not None

def verificar_download(cnpj, data_inicio, data_fim, excel_path):
    # Carregar o arquivo Excel
    df = pd.read_excel(excel_path, sheet_name='Sheet1')
    
    cnpj = int(cnpj)

    # Filtrar as linhas que correspondem aos critérios
    df_filtrado = df[
        (df['CNPJ'] == cnpj) &
        (df['DATA INICIO'] == data_inicio) &
        (df['DATA FIM'] == data_fim)
    ]
    
    # Verificar se há pelo menos uma linha que atenda aos critérios
    if len(df_filtrado) > 0:
        return False
    else:
        return True
    
def mover_faturas_lidas(origem, destino):
    try:
        if not os.path.exists(destino):
            os.makedirs(destino)
        for arquivo in os.listdir(origem):
            if arquivo.lower().endswith('.pdf'):
                caminho_completo = os.path.join(origem, arquivo)
                shutil.move(caminho_completo, destino)
                print(f'Arquivo {arquivo} movido para {destino}.')
    except Exception as e:
        print(f'Erro ao mover arquivos: {e}')
# Caminho da pasta onde os PDFs estão localizados
pasta_pdfs = r'G:\QUALIDADE\Códigos\Leitura de Faturas Gás\Códigos\Gasmig\Faturas'

# Referência que queremos encontrar nos nomes dos arquivos
referencia = 'GÁSMIG'  # Exemplo: todos os arquivos que contêm 'FATURA' no nome

# Listar todos os PDFs na pasta que contêm a referência no nome
pdfs_com_referencia = listar_pdfs_com_referencia_na_pasta(pasta_pdfs, referencia)

# Exibir os PDFs encontrados
for pdf in pdfs_com_referencia:
    print(pdf)



