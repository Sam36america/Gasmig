# COORDENADAS PARA EXTRAÇÃO DE DADOS USANDO OCR

def corte_gasmig():
    
    corte = {
    'cnpj': (2155, 700, 2622, 800), #17/09
    'cnpj_ajustado': (2155, 500, 2622, 1200),
    'cnpj_ajustado2': (2155, 500, 2722, 2200), #17/09
    'cnpj_ajustado3': (2155, 1100, 3222, 2400), #17/09

    'valor_total': (3120, 3780, 3432, 3950), #17/09
    'valor_total_ajustado':(),

    'volume_total': (2900, 1320, 3430, 1380), # 17/09
    'volume_total_ajustado':  (2040, 2990, 2285, 3100),
    'volume_total_ajustado2': (2000, 2890, 2385, 3200),

    'data_emissao': (3400, 1495, 3820, 1625), #  17/09
    'data_emissao_ajustado': (),

    'data_inicio': (1200, 1030, 3000, 1400), # 17/09
    'data_inicio_ajustado': (505, 4680, 1800, 4830),
    'data_inicio_ajustado2': (100, 4650, 1700, 4750),

    'data_fim': (1200, 1030, 3000, 1400), # 17/09
    'data_fim_ajustado': (272, 4680, 1800, 4830),
    'data_fim_ajustado2': (300, 4650, 1700, 4750),

    'numero_fatura': (2820, 120, 3200, 280), # 17/09
    'numero_fatura_ajustado': (000, 430, 3750, 590),

    'valor_icms': (1140, 3870, 1510, 3950), # 17/09
    'valor_icms_ajustado': (000, 430, 3750, 590),

    'correcao_pcs': (2620, 1320, 2930, 1380),
    }
    return corte

# CAMINHO DA PLANILHA

caminho_excel = r'G:\QUALIDADE\Códigos\Leitura de Faturas Gás\Códigos\00 Faturas Lidas\GASMIG.xlsx'

'''while True
    for Player(get_moeda):
        if player(moeda >= 3):
            brilhar(dourado)
        else:
            brilhar(branco)'''


