from PIL import Image
import pytesseract
import re
import os
import pandas as pd
import numpy as np
from config import corte_gasmig, caminho_excel
from funcoes import pdf_ocr, pdf_to_image, dados_excel, adicionar_dados_excel, verificar_download, mover_faturas_lidas, get_first_pdf, get_second_pdf   

#in/out 3100, 1300, 3800, 1450
#X-Y X-Y
usuario_conectado = 'samuel.santos'
# Configure o caminho do executável do Tesseract
pytesseract.pytesseract.tesseract_cmd = fr'C:\Users\{usuario_conectado}\AppData\Local\Programs\Tesseract-OCR\tesseract.exe'

DIST = 'GÁSMIG'

def extrator_cnpj(folder_path, corte):
    segundo_pdf = get_second_pdf(folder_path)
    if not segundo_pdf:
        return False

    imagem = pdf_to_image(segundo_pdf)
    cnpj_crop = imagem.crop(corte['cnpj'])
    cnpj_crop.show()
    cnpj_text = pdf_ocr(cnpj_crop)
    cnpj = cnpj_text.replace(',', '').replace('/', '').replace('-', '').replace('.', '')
    cnpj_match = re.findall(r'(\d+)', cnpj)

    return cnpj_match[0] if cnpj_match else False

def extrator_valor_total(folder_path, corte):
    segundo_pdf = get_second_pdf(folder_path)
    if not segundo_pdf:
        return False

    imagem = pdf_to_image(segundo_pdf)
    valor_total_crop = imagem.crop(corte['valor_total'])
    valor_total_text = pdf_ocr(valor_total_crop)
    valor_total_match = re.findall(r'(\d{1,3}\.?\d{1,3}\.?\s?\,?\d{1,2})', valor_total_text)

    return valor_total_match[0].strip() if valor_total_match else False
    
def extrator_volume_total(folder_path, corte):
    segundo_pdf = get_second_pdf(folder_path)
    if not segundo_pdf:
        return False

    imagem = pdf_to_image(segundo_pdf)
    volume_total_crop = imagem.crop(corte['volume_total'])
    volume_total_text = pdf_ocr(volume_total_crop)
    volume_total_match = re.findall(r'\s?(\d+\.?\,?\s?\d+\.?\,?\s?\d+)', volume_total_text)
    
    return volume_total_match[0].strip() if volume_total_match else False
    
def extrator_data_emissao(folder_path, corte):
    primeiro_pdf = get_first_pdf(folder_path)
    if not primeiro_pdf:
        return False

    imagem = pdf_to_image(primeiro_pdf)
    data_emissao_crop = imagem.crop(corte['data_emissao'])
    data_emissao_text = pdf_ocr(data_emissao_crop)
    data_emissao_match = re.findall(r'\d{2}\.?\/?\d{2}\.?\/?\d{2,4}', data_emissao_text)
    
    return data_emissao_match[0].strip() if data_emissao_match else False
    
def extrator_data_inicio(folder_path, corte):
    segundo_pdf = get_second_pdf(folder_path)
    if not segundo_pdf:
        return False

    imagem = pdf_to_image(segundo_pdf)
    data_inicio_crop = imagem.crop(corte['data_inicio'])
    data_inicio_text = pdf_ocr(data_inicio_crop)
    data_inicio_match = re.findall(r'(\d{2}\/\d{2}\/\d{4})', data_inicio_text)
    
    return data_inicio_match[0].strip() if data_inicio_match else False

def extrator_data_fim(folder_path, corte):
    segundo_pdf = get_second_pdf(folder_path)
    if not segundo_pdf:
        return False

    imagem = pdf_to_image(segundo_pdf)
    data_fim_crop = imagem.crop(corte['data_fim'])
    data_fim_text = pdf_ocr(data_fim_crop)
    data_fim_match = re.findall(r'a\s?(\d{2}\/\d{2}\/\d{4})', data_fim_text)
    
    return data_fim_match[0].strip() if data_fim_match else False
        
def extrator_numero_fatura(folder_path, corte):
    segundo_pdf = get_second_pdf(folder_path)
    if not segundo_pdf:
        return False

    imagem = pdf_to_image(segundo_pdf)
    numero_fatura_crop = imagem.crop(corte['numero_fatura'])
    numero_fatura_text = pdf_ocr(numero_fatura_crop)
    numero_fatura_match = re.findall(r'\d+', numero_fatura_text)
    
    return numero_fatura_match[0] if numero_fatura_match else False

def extrator_valor_icms(folder_path, corte):
    segundo_pdf = get_second_pdf(folder_path)
    if not segundo_pdf:
        return False

    imagem = pdf_to_image(segundo_pdf)
    valor_icms_crop = imagem.crop(corte['valor_icms'])
    valor_icms_text = pdf_ocr(valor_icms_crop)
    valor_icms_match = re.findall(r'\d+\.?\,?\d+\.?\,?\d+\.?\,?', valor_icms_text)
    
    return valor_icms_match[0] if valor_icms_match else False

def extrator_correcao_pcs(folder_path, corte):
    segundo_pdf = get_second_pdf(folder_path)
    if not segundo_pdf:
        return False

    imagem = pdf_to_image(segundo_pdf)
    correcao_pcs_crop = imagem.crop(corte['correcao_pcs'])
    #correcao_pcs_crop.show()
    correcao_pcs_text = pdf_ocr(correcao_pcs_crop)
    correcao_pcs_match = re.findall(r'\d+\.?\,?\d+\.?\,?\d+\.?\,?', correcao_pcs_text)
    
    return correcao_pcs_match[0] if correcao_pcs_match else False

def main(folder_path):
    corte = corte_gasmig()  # Isso retorna um dicionário com os cortes

    # Extrair CNPJ
    cnpj = extrator_cnpj(folder_path, corte)
    if cnpj == False:
        corte_ajustado = corte.copy()
        corte_ajustado['cnpj'] = corte['cnpj_ajustado']
        cnpj = extrator_cnpj(folder_path, corte_ajustado)
        if cnpj == False:    
            corte_ajustado['cnpj'] = corte['cnpj_ajustado2']
            cnpj = extrator_cnpj(folder_path, corte_ajustado)
            if cnpj == False:
                corte_ajustado['cnpj'] = corte['cnpj_ajustado3']
                cnpj = extrator_cnpj(folder_path, corte_ajustado)

    # Extrair valor total
    valor_total = extrator_valor_total(folder_path, corte)
    if valor_total == False:
        corte_ajustado = corte.copy()
        corte_ajustado['valor_total'] = corte['valor_total_ajustado']
        valor_total = extrator_valor_total(folder_path, corte_ajustado)

    # Extrair volume total
    volume_total = extrator_volume_total(folder_path, corte)
    if volume_total == False:
        corte_ajustado = corte.copy()
        corte_ajustado['volume_total'] = corte['volume_total_ajustado']
        volume_total = extrator_volume_total(folder_path, corte_ajustado)
        if volume_total == False:
            corte_ajustado['volume_total'] = corte['volume_total_ajustado2']
            volume_total = extrator_volume_total(folder_path, corte_ajustado)

    # Extrair data emissão
    data_emissao = extrator_data_emissao(folder_path, corte)
    if data_emissao == False:
        corte_ajustado = corte.copy()
        corte_ajustado['data_emissao'] = corte['data_emissao_ajustado']
        data_emissao = extrator_data_emissao(folder_path, corte_ajustado)

    # Extrair data início
    data_inicio = extrator_data_inicio(folder_path, corte)
    if data_inicio == False:
        corte_ajustado = corte.copy()
        corte_ajustado['data_inicio'] = corte['data_inicio_ajustado']
        data_inicio = extrator_data_inicio(folder_path, corte_ajustado)
        if data_inicio == False:
            corte_ajustado['data_inicio'] = corte['data_inicio_ajustado2']
            data_inicio = extrator_data_inicio(folder_path, corte_ajustado)

    # Extrair data fim
    data_fim = extrator_data_fim(folder_path, corte)
    if data_fim == False:
        corte_ajustado = corte.copy()
        corte_ajustado['data_fim'] = corte['data_fim_ajustado']
        data_fim = extrator_data_fim(folder_path, corte_ajustado)
        if data_fim == False:
            corte_ajustado['data_fim'] = corte['data_fim_ajustado2']
            data_fim = extrator_data_fim(folder_path, corte_ajustado)

    # Extrair número fatura
    numero_fatura = extrator_numero_fatura(folder_path, corte)
    if numero_fatura == False:
        corte_ajustado = corte.copy()
        corte_ajustado['numero_fatura'] = corte['numero_fatura_ajustado']
        numero_fatura = extrator_numero_fatura(folder_path, corte_ajustado)

    # Extrair valor ICMS
    valor_icms = extrator_valor_icms(folder_path, corte)
    if valor_icms == False:
        corte_ajustado = corte.copy()
        corte_ajustado['valor_icms'] = corte['valor_icms_ajustado']
        valor_icms = extrator_valor_icms(folder_path, corte_ajustado)

    # Extrair correção PCS
    correcao_pcs = extrator_correcao_pcs(folder_path, corte)
    if correcao_pcs == False:
        corte_ajustado = corte.copy()
        corte_ajustado['correcao_pcs'] = corte['correcao_pcs_ajustado']
        correcao_pcs = extrator_correcao_pcs(folder_path, corte_ajustado)

    # Verificar se todos os dados foram extraídos corretamente
    if not cnpj or not valor_total or not volume_total or not data_emissao or not data_inicio or not data_fim or not numero_fatura or not valor_icms or not correcao_pcs:
        print('Fatura não movida devido a dados incompletos.')
    else: 
        verificar = verificar_download(cnpj, data_inicio, data_fim, caminho_excel)
        if verificar:
            data_frame = dados_excel(cnpj, valor_total, volume_total, data_emissao, data_inicio, data_fim, numero_fatura, valor_icms, correcao_pcs, DIST, segundo_arquivo)
            adicionar_dados_excel(caminho_excel, data_frame)
            print(f'Dados da fatura {segundo_arquivo} adicionados com sucesso na planilha.')
            mover_faturas_lidas(folder_path, diretorio_destino)
            print(f'Fatura {segundo_arquivo} movida para {diretorio_destino}.')
        else:
            print('Dados já inseridos!')

# Exemplo de uso
file_path = r'G:\QUALIDADE\Códigos\Leitura de Faturas Gás\Códigos\Gasmig\Faturas'
diretorio_destino = r'G:\QUALIDADE\Códigos\Leitura de Faturas Gás\Códigos\Gasmig\Lidos'

pdf_files = [f for f in os.listdir(file_path) if f.lower().endswith('.pdf')]

# Ordena a lista de arquivos (opcional, dependendo de como você quer definir o "segundo" arquivo)
pdf_files.sort()

if len(pdf_files) >= 1:
    # Pega o primeiro arquivo (índice 0)
    primeiro_arquivo = pdf_files[0]
    # Verifica se há um segundo arquivo (índice 1)
    if len(pdf_files) >= 2:
        segundo_arquivo = pdf_files[1]
    else:
        segundo_arquivo = primeiro_arquivo
    arquivo_completo = os.path.join(file_path, segundo_arquivo)
    main(file_path)
else:
    print("Não há arquivos PDF na pasta.")