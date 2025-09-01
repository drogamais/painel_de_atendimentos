import pandas as pd
import mysql.connector
import hashlib
import glob
import openpyxl
import sys
import os
import requests # Importa a biblioteca para requisições web
from datetime import datetime

# --- CONFIGURAÇÕES DO TELEGRAM E BANCO DE DADOS ---
# ATENÇÃO: Preencha com suas informações reais.
TELEGRAM_BOT_TOKEN = "8096205039:AAGz3TqmfyXGI__NGdyvf6TnMDNA--pvAWc"
TELEGRAM_CHAT_ID = "7035974555"

# ATENÇÃO: Substitua este dicionário pela sua configuração de banco de dados real.
# É recomendado carregar essas credenciais de um local seguro, não diretamente no código.
db_config = {
        "user": "drogamais",
        "password": "dB$MYSql@2119",
        "host": "10.48.12.20",
        "port": "3306",
        "database": "dbSults"
    }
# ----------------------------------------------------


# --- FUNÇÃO HÍBRIDA PARA ENVIAR NOTIFICAÇÃO AO TELEGRAM ---
def enviar_notificacao_hibrida_telegram(log_path, status):
    """
    Envia uma notificação híbrida para o Telegram:
    1. Uma mensagem de texto com o status final (Sucesso/Erro).
    2. O arquivo de log completo como anexo.
    """
    # Define a mensagem de status baseada no resultado da execução
    if status == "SUCESSO":
        status_message = "✅ *Automação finalizada com SUCESSO!*"
    else:
        status_message = "❌ *Automação finalizada com ERRO!*"
    
    # Tenta enviar a mensagem de texto primeiro
    try:
        url_message = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendMessage"
        payload = {
            'chat_id': TELEGRAM_CHAT_ID,
            'text': status_message,
            'parse_mode': 'Markdown'
        }
        response_msg = requests.post(url_message, data=payload)
        response_msg.raise_for_status()
        print("Mensagem de status enviada ao Telegram.")
    except requests.exceptions.RequestException as e:
        print(f"Erro ao enviar mensagem de status para o Telegram: {e}")
        # Mesmo com erro na mensagem, prossegue para tentar enviar o arquivo
    
    # Agora, tenta enviar o arquivo de log
    try:
        if not os.path.exists(log_path):
            print(f"Aviso: Arquivo de log não encontrado em '{log_path}'. Não será enviado.")
            return

        with open(log_path, 'rb') as arquivo_log:
            url_document = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendDocument"
            files = {'document': arquivo_log}
            data = {'chat_id': TELEGRAM_CHAT_ID, 'caption': 'Segue o log completo da execução.'}
            
            response_doc = requests.post(url_document, data=data, files=files)
            response_doc.raise_for_status()
            print("Arquivo de log enviado com sucesso para o Telegram.")

    except requests.exceptions.RequestException as e:
        print(f"Erro ao enviar arquivo de log para o Telegram: {e}")
    except Exception as e:
        print(f"Ocorreu um erro inesperado ao tentar enviar o log: {e}")


# --- CLASSE DE LOG PERSONALIZADA ---
# Responsável por mostrar as mensagens no console e salvá-las no log.txt ao mesmo tempo.
class Logger:
    def __init__(self, log_path, mode):
        self.terminal = sys.__stdout__
        self.log_file = open(log_path, mode, encoding="utf-8")

    def write(self, message):
        self.terminal.write(message)
        self.log_file.write(message)

    def flush(self):
        self.terminal.flush()
        self.log_file.flush()
    
    def close(self):
        self.log_file.close()

def run_sync():
    """
    Função principal que encapsula toda a lógica de sincronização.
    """
    # --- CONFIGURAÇÕES ---
    
    #Caminho definitivo das planilhas
    caminho_da_pasta = r"T:\APF\Atendimento ao Associado\LANÇAMENTO DE ATENDIMENTO DIARIOS"

    #Caminho temporário para testes das planilhas
    #caminho_da_pasta = r"T:\APF\Inteligência de Mercado\Códigos Python\Projeto Tarefas Focais\Planilhas Projeto"
    
    padroes_de_arquivos = ["Focal_*.xlsx", "Auditor_*.xlsx", "Adm.xlsx", "Coordenador.xlsx"]
    nome_tabela_atendimentos = "tb_atendimentos"

    # --- LÓGICA PARA BUSCAR ARQUIVOS ---
    print(f"Buscando arquivos na pasta: {caminho_da_pasta}")
    lista_de_arquivos = []
    for padrao in padroes_de_arquivos:
        caminho_completo_padrao = os.path.join(caminho_da_pasta, padrao)
        lista_de_arquivos.extend(glob.glob(caminho_completo_padrao))

    if not lista_de_arquivos:
        raise FileNotFoundError("Nenhum arquivo encontrado com os padrões especificados.")

    print(f"✅ Encontrados {len(lista_de_arquivos)} arquivos:")
    for f in sorted(lista_de_arquivos):
        print(f"  - {os.path.basename(f)}")

    # --- FUNÇÕES AUXILIARES ---
    def gerar_hash_conteudo(linha):
        string_unificada = "".join(str(valor) for valor in linha)
        return hashlib.sha256(string_unificada.encode("utf-8")).hexdigest()

    def limpar_e_preparar_planilha(caminho_arquivo, num_linhas=1000):
        try:
            print(f"  -> Limpando o arquivo: {os.path.basename(caminho_arquivo)}...")
            workbook = openpyxl.load_workbook(caminho_arquivo)
            sheet = workbook.active
            for row_index in range(sheet.max_row, 1, -1):
                sheet.delete_rows(row_index)
            print(f"  -> Aplicando fórmula de ID para {num_linhas} linhas...")
            for i in range(2, num_linhas + 2):
                formula_id = f'=IF(COUNTA(B{i}:G{i})=6, COUNTA(B$2:B{i}), "")'
                sheet[f"A{i}"] = formula_id
            workbook.save(caminho_arquivo)
            print(f"  -> Arquivo limpo e preparado com sucesso.")
            return True
        except Exception as e:
            print(f"  -> ❌ Erro ao limpar o arquivo {os.path.basename(caminho_arquivo)}: {e}")
            return False

    # --- LEITURA DE MÚLTIPLOS ARQUIVOS ---
    lista_dfs = []
    for caminho_do_arquivo in lista_de_arquivos:
        try:
            print(f"\nLendo a planilha: {os.path.basename(caminho_do_arquivo)}...")
            df_temp = pd.read_excel(caminho_do_arquivo, engine="openpyxl")
            if df_temp.empty:
                print(f"  -> ⚠️  Aviso: O arquivo {os.path.basename(caminho_do_arquivo)} está vazio e será ignorado.")
                continue
            nome_base_arquivo, _ = os.path.splitext(os.path.basename(caminho_do_arquivo))
            df_temp["arquivo_origem"] = nome_base_arquivo
            lista_dfs.append(df_temp)
        except Exception as e:
            print(f"❌ Erro ao ler o arquivo {os.path.basename(caminho_do_arquivo)}: {e}. Pulando.")
            continue

    if not lista_dfs:
        print("\nNenhuma linha de dados encontrada em nenhum dos arquivos. Encerrando.")
        return

    df = pd.concat(lista_dfs, ignore_index=True)
    print(f"\n✅ Sucesso! Total de {len(df)} linhas lidas de {len(lista_dfs)} arquivos válidos.")

    # --- LIMPEZA DO DATAFRAME COMBINADO ---
    df = df.rename(columns={
        "ID": "id_planilha", "arquivo_origem": "funcao", "Data": "data", "Tarefa": "tarefa",
        "Responsável": "responsavel", "Loja": "loja", "Tipo": "tipo",
        "Ação": "acao", "Assunto": "assunto",
    })
    colunas_essenciais = ["id_planilha", "data", "tarefa", "responsavel", "loja", "tipo", "acao"]
    colunas_faltando = [col for col in colunas_essenciais if col not in df.columns]
    if colunas_faltando:
        raise ValueError(f"❌ Erro Crítico: Colunas essenciais faltando: {colunas_faltando}.")
    df = df.astype(object).where(pd.notna(df), None)
    if "assunto" not in df.columns:
        df["assunto"] = None
    print("✅ DataFrame combinado foi limpo e padronizado.")

    # --- LÓGICA DE SINCRONIZAÇÃO ---
    connection = None
    try:
        connection = mysql.connector.connect(**db_config, collation="utf8mb4_unicode_ci")
        
        df["data"] = pd.to_datetime(df["data"], errors="coerce")
        df["id_planilha"] = pd.to_numeric(df["id_planilha"], errors="coerce")
        df.dropna(subset=["data", "id_planilha"], inplace=True)
        if df.empty:
            print("\nNenhuma linha válida após a limpeza de datas e IDs.")
            return
        df["id_planilha"] = df["id_planilha"].astype(int)
        
        print("\nCriando chaves de negócio e hashes de conteúdo...")
        df["data_str"] = df["data"].dt.strftime("%Y-%m-%d")
        df["chave_id"] = df["data_str"] + "-" + df["id_planilha"].astype(str) + "-" + df["funcao"]
        colunas_conteudo = ["data", "tarefa", "responsavel", "loja", "tipo", "acao", "assunto"]
        df["conteudo_hash"] = df[colunas_conteudo].apply(gerar_hash_conteudo, axis=1)

        print("Buscando registros existentes no banco de dados...")
        cursor = connection.cursor(dictionary=True)
        cursor.execute(f"SELECT chave_id, conteudo_hash FROM {nome_tabela_atendimentos}")
        db_data = {row["chave_id"]: row["conteudo_hash"] for row in cursor.fetchall()}
        cursor.close()
        print(f"✅ Encontrados {len(db_data)} registros no banco.")

        print("Comparando dados e preparando lotes...")
        para_inserir, para_atualizar = [], []
        for index, row in df.iterrows():
            chave, conteudo_atual = row["chave_id"], row["conteudo_hash"]
            if chave not in db_data:
                para_inserir.append(row)
            elif db_data[chave] != conteudo_atual:
                para_atualizar.append(row)
        
        print(f"🔎 Verificação concluída: {len(para_inserir)} para INSERIR, {len(para_atualizar)} para ATUALIZAR.")

        if para_inserir:
            df_inserir = pd.DataFrame(para_inserir)
            print(f"\nInserindo {len(df_inserir)} novos registros...")
            cursor = connection.cursor()
            colunas_db = ["chave_id", "id_planilha", "funcao"] + colunas_conteudo + ["conteudo_hash"]
            placeholders = ", ".join(["%s"] * len(colunas_db))
            query_insert = f"INSERT INTO {nome_tabela_atendimentos} ({', '.join(f'`{c}`' for c in colunas_db)}) VALUES ({placeholders})"
            dados_inserir = [tuple(r) for r in df_inserir[colunas_db].to_numpy()]
            cursor.executemany(query_insert, dados_inserir)
            cursor.close()

        if para_atualizar:
            df_atualizar = pd.DataFrame(para_atualizar)
            print(f"\nAtualizando {len(df_atualizar)} registros existentes...")
            cursor = connection.cursor()
            update_cols = ["id_planilha", "funcao"] + colunas_conteudo + ["conteudo_hash"]
            update_set = ", ".join([f"`{col}`=%s" for col in update_cols])
            query_update = f"UPDATE {nome_tabela_atendimentos} SET {update_set} WHERE `chave_id` = %s"
            for _, row in df_atualizar.iterrows():
                valores = list(row[update_cols]) + [row["chave_id"]]
                cursor.execute(query_update, valores)
            cursor.close()

        connection.commit()
        print("\n✅ Sincronização de dados concluída com sucesso!")

        # A função weekday() retorna 0 para segunda e 4 para sexta.
        if datetime.today().weekday() == 4:
            print("\n--- TAREFA DE SEXTA-FEIRA ---")
            print("Preparando planilhas para a próxima semana...")
            sucesso_total = True
            for arquivo in lista_de_arquivos:
                if not limpar_e_preparar_planilha(arquivo):
                    sucesso_total = False
            if sucesso_total:
                print("\n✅ Todas as planilhas foram limpas e preparadas!")
            else:
                print("\n⚠️ Atenção: Ocorreram erros ao limpar uma ou mais planilhas.")
        else:
            print("\nNenhuma ação de limpeza necessária hoje.")
            
    finally:
        if connection and connection.is_connected():
            connection.close()
            print("\nConexão com o banco de dados fechada.")

# --- BLOCO DE EXECUÇÃO PRINCIPAL ---
if __name__ == "__main__":
    original_stdout = sys.stdout
    original_stderr = sys.stderr

    if getattr(sys, "frozen", False):
        base_dir = os.path.dirname(sys.executable)
    else:
        base_dir = os.path.dirname(os.path.abspath(__file__))

    log_path = os.path.join(base_dir, "log.txt")
    # Zera o log na segunda-feira (0), senão concatena.
    mode = "w" if datetime.today().weekday() == 0 else "a"

    logger = Logger(log_path, mode)
    sys.stdout = logger
    sys.stderr = logger

    status_final = "SUCESSO"

    try:
        print("\n--- Execução iniciada:", datetime.now(), "---")
        run_sync()
        print("\n--- Execução finalizada com sucesso:", datetime.now(), "---")

    except Exception as e:
        status_final = "ERRO"
        print(f"\n--- ERRO CRÍTICO NA EXECUÇÃO: {e} ---")
        print("\n--- Execução finalizada com erro:", datetime.now(), "---")

    finally:
        # Restaura a saída padrão para que as mensagens de envio do log apareçam no console
        sys.stdout = original_stdout
        sys.stderr = original_stderr
        logger.close() # Garante que o arquivo de log está fechado antes de ser lido

        # Chama a função para enviar a notificação HÍBRIDA ao Telegram
        enviar_notificacao_hibrida_telegram(log_path, status_final)

        if status_final == "SUCESSO":
            print("\n=======================================================")
            print("  Processo finalizado com sucesso!")
            print("  Todos os detalhes foram salvos em log.txt")
            print("=======================================================")
        else:
            print("\n=======================================================")
            print("  Houve um erro durante a execução.")
            print("  Por favor, verifique o arquivo log.txt para detalhes.")
            print("=======================================================")
