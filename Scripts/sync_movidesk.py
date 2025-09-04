import requests
import json
import pandas as pd
import time
import re
import numpy as np

# --- CONFIGURAÇÕES API---
MOVIDESK_API_TOKEN = "5514b190-8715-4587-8ee7-8ad802c86dcc" # Lembre-se de colocar seu token real
BASE_URL = "https://api.movidesk.com/public/v1"
TAMANHO_PAGINA = 100

# --- PARÂMETROS DA REQUISIÇÃO ---
params = {
    'token': MOVIDESK_API_TOKEN,
    '$select': 'id,subject,status,owner,clients,createdDate,resolvedIn',
    '$orderby': 'resolvedIn desc',
    '$filter': "(status eq 'Fechado' or status eq 'Resolvido') and resolvedIn ge 2025-07-01T00:00:00Z",
    '$expand': 'owner($select=businessName),clients($select=businessName)',
    '$top': TAMANHO_PAGINA,
    '$skip': 0
}

# --- FUNÇÃO PERSONALIZADA PARA TRATAR OS CLIENTES (sem alterações) ---
def processar_clientes(row):
    client_string = row['clientsName']
    adicionais, loja_numero, nome_loja_final = '', pd.NA, client_string
    if not isinstance(client_string, str) or not client_string: return adicionais, loja_numero, nome_loja_final
    if ',' in client_string:
        parts = [p.strip() for p in client_string.split(',')]
        nomes_pessoas = [p for p in parts if 'Drogamais' not in p]
        string_da_loja = next((p for p in parts if 'Drogamais' in p), '')
        adicionais = ', '.join(nomes_pessoas)
    else:
        string_da_loja = client_string
    match = re.search(r'(\d+)\s*-\s*(.*)', string_da_loja)
    if match:
        loja_numero = int(match.group(1))
        nome_loja_final = f"Drogamais {match.group(2).strip()}"
    return adicionais, loja_numero, nome_loja_final

# --- LOOP DE PAGINAÇÃO (sem alterações) ---
todos_os_tickets = []
pagina_atual = 1
print("Iniciando a busca de tickets Fechados/Resolvidos a partir de Julho/2025...")
print("-" * 40)
while True:
    try:
        print(f"Buscando página {pagina_atual}...")
        response = requests.get(f"{BASE_URL}/tickets", params=params)
        response.raise_for_status()
        tickets_da_pagina = response.json()
        if not tickets_da_pagina:
            print("\nBusca finalizada.")
            break
        todos_os_tickets.extend(tickets_da_pagina)
        print(f"  > Sucesso! Total acumulado: {len(todos_os_tickets)}")
        params['$skip'] += TAMANHO_PAGINA
        pagina_atual += 1
        time.sleep(6)
    except Exception as e:
        print(f"\nOcorreu um erro: {e}")
        break
print("-" * 40)

# --- TRATAMENTO E EXPORTAÇÃO PARA EXCEL ---
if not todos_os_tickets:
    print("Nenhum ticket foi encontrado com os critérios especificados.")
else:
    print(f"Total de {len(todos_os_tickets)} tickets baixados. Iniciando tratamento de dados...")
    df = pd.DataFrame(todos_os_tickets)
    
    # Tratamento de Datas
    print("Tratando e ajustando colunas de data...")
    df['createdDate'] = pd.to_datetime(df['createdDate'], errors='coerce').dt.tz_localize('UTC').dt.tz_convert('America/Sao_Paulo')
    df['resolvedIn'] = pd.to_datetime(df['resolvedIn'], errors='coerce').dt.tz_localize('UTC').dt.tz_convert('America/Sao_Paulo')
    df['dataConclusao'] = df['resolvedIn'].dt.strftime('%Y-%m-%d')

    # Extração de Nomes
    if 'owner' in df.columns:
        df['ownerName'] = df['owner'].apply(lambda x: x['businessName'].strip() if isinstance(x, dict) and x.get('businessName') else '')
    else:
        df['ownerName'] = ''
    if 'clients' in df.columns:
        df['clientsName'] = df['clients'].apply(lambda l: ', '.join([c['businessName'] for c in l]) if isinstance(l, list) and l else '')
    else:
        df['clientsName'] = ''
    df = df.drop(columns=['owner', 'clients'], errors='ignore')

    # Tratamento Avançado de Clientes
    print("Aplicando tratamento avançado na coluna de clientes...")
    df[['Adicionais', 'Loja_numero', 'clientsName']] = df.apply(processar_clientes, axis=1, result_type='expand')
    df['Loja_numero'] = df['Loja_numero'].astype('Int64').fillna(0)

    # Lógica de Categorização
    print("Criando colunas padrão e categorizando clientes...")
    df['Tarefa'] = 'Atendimento Geral'
    df.loc[df['Loja_numero'] == 0, 'clientsName'] = 'Outro'
    df['Tipo'] = 'MOVIDESK'
    df['Acao'] = 'PASSIVO'

    # --- INÍCIO DA NOVA SEÇÃO: RENOMEAR E REORDENAR COLUNAS ---
    print("Renomeando e reordenando colunas para o padrão final...")

    # 1. Dicionário para mapear nomes antigos para os novos
    mapa_para_renomear = {
        'createdDate': 'data_criacao',
        'dataConclusao': 'data_conclusao',
        'ownerName': 'responsavel',
        'clientsName': 'cliente',
        'Loja_numero': 'loja_numero',
        'Adicionais': 'adjunto',
        'Tarefa': 'tarefa',
        'Tipo': 'tipo',
        'Acao': 'acao',
        'subject': 'assunto'
    }
    df = df.rename(columns=mapa_para_renomear)

    # 2. Lista com a ordem final exata das colunas
    ordem_final = [
        'id', 'status', 'data_criacao', 'data_conclusao', 'resolvedIn',
        'responsavel', 'cliente', 'loja_numero', 'adjunto',
        'tarefa', 'tipo', 'acao', 'assunto'
    ]
    df_final = df[ordem_final]
    # --- FIM DA NOVA SEÇÃO ---
    
    # Salvar em arquivo Excel
    nome_do_arquivo_excel = 'tickets_concluidos_julho_2025_final.xlsx'
    try:
        # Formata as colunas de data para texto ANTES de salvar
        df_final['data_criacao'] = df_final['data_criacao'].dt.strftime('%Y-%m-%d %H:%M:%S')
        df_final['resolvedIn'] = df_final['resolvedIn'].dt.strftime('%Y-%m-%d %H:%M:%S')
        # data_conclusao já está como texto, não precisa formatar

        df_final.to_excel(nome_do_arquivo_excel, index=False, sheet_name='Tickets Concluídos')
        print(f"\nSUCESSO! Dados salvos no arquivo Excel '{nome_do_arquivo_excel}'")
    except Exception as e:
        print(f"\nERRO: Ocorreu um problema ao salvar o arquivo Excel: {e}")