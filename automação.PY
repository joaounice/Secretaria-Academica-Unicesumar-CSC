import pandas as pd
import os
import glob
import time 

# --- Configurações da TAREFA ORIGINAL (Credenciais) ---
# Caminho completo do arquivo Excel de origem (source)
caminho_excel_credenciais = r"C:\Users\Joao.silverio\OneDrive - Uniasselvi\BI - CSC SCRA UniCesumar - Arquivos - Arquivos\RPA\site_secretaria\CREDENCIAIS.xlsx"

# Diretório de destino (target) - SEM subpasta
diretorio_saida = r"C:\Users\Joao.silverio\OneDrive - Uniasselvi\BI - CSC SCRA UniCesumar - Arquivos - Arquivos\RPA\site_secretaria"

# Nome do arquivo JSON
nome_arquivo_json_credenciais = "credenciais_login_csc.json"

# Cria o caminho completo de destino
caminho_saida_json_credenciais = os.path.join(diretorio_saida, nome_arquivo_json_credenciais)

# --- Configurações da NOVA TAREFA (Análise Semanal - Arquivo Mais Recente) ---

# Diretório de origem para buscar o arquivo mais recente (contém CSVs)
diretorio_origem_fila = r"C:\Users\Joao.silverio\OneDrive - Uniasselvi\BI - CSC SCRA UniCesumar - Arquivos - Arquivos\FilaAtendimentos"

# Nome do arquivo JSON de saída para a Análise Semanal
nome_arquivo_json_analise = "analise semanal.json"

# Cria o caminho completo de destino
caminho_saida_json_analise = os.path.join(diretorio_saida, nome_arquivo_json_analise)

# Limite de tamanho desejado em Megabytes para o arquivo de análise
limite_mb_analise = 3.0 # 3 MB


# --- FUNÇÕES AUXILIARES ---

def encontrar_ultimo_arquivo(diretorio, extensao):
    """
    Busca o arquivo com a extensão especificada (.csv) mais recentemente modificado em um diretório.
    """
    # Cria a máscara de busca
    busca = os.path.join(diretorio, f'*.{extensao}')

    # Lista todos os arquivos que correspondem à máscara
    todos_arquivos = glob.glob(busca)

    if not todos_arquivos:
        return None

    # Encontra o arquivo mais recente usando a data de última modificação (getmtime)
    caminho_arquivo_mais_recente = max(todos_arquivos, key=os.path.getmtime)

    return caminho_arquivo_mais_recente


def processar_e_salvar_json(caminho_origem, caminho_json_destino, limite_mb_max=9999):
    """
    Lê o arquivo (Excel ou CSV), converte para JSON com orient='records' (leve) e salva.
    A função de leitura é escolhida com base na extensão do arquivo de origem.
    """
    nome_arquivo = os.path.basename(caminho_origem)
    extensao = nome_arquivo.split('.')[-1].lower()
    df = None

    try:
        print(f"Lendo o arquivo de {extensao.upper()}: {nome_arquivo}")
        
        if extensao in ['xlsx', 'xls']:
            # Leitura para arquivos Excel (usado no Bloco 1)
            df = pd.read_excel(caminho_origem)
        elif extensao == 'csv':
            # Leitura para arquivos CSV (usado no Bloco 2)
            # Tenta ler com os parâmetros comuns para CSVs brasileiros: separador ';', encoding 'latin1'
            try:
                df = pd.read_csv(caminho_origem, sep=';', encoding='latin1')
            except pd.errors.ParserError:
                 # Se falhar com ';', tenta o padrão ','
                df = pd.read_csv(caminho_origem, sep=',', encoding='latin1')
            except UnicodeDecodeError:
                # Se falhar com latin1, tenta utf-8
                df = pd.read_csv(caminho_origem, sep=';', encoding='utf-8')
        else:
            print(f"ERRO: Extensão de arquivo não suportada para leitura: .{extensao}")
            return

        # 1. Conversão para JSON
        print("Convertendo para JSON (orient='records')...")
        # orient='records' (lista de dicionários) é o formato mais leve e comum
        json_data = df.to_json(orient='records', force_ascii=False)

        # 2. Salvamento do arquivo JSON
        with open(caminho_json_destino, 'w', encoding='utf-8') as f:
            f.write(json_data)

        # 3. Verificação de Tamanho (feedback)
        tamanho_bytes = os.path.getsize(caminho_json_destino)
        tamanho_mb = tamanho_bytes / (1024 * 1024)

        print("\n--- Concluído ---")
        print(f"Arquivo JSON salvo com sucesso em:")
        print(caminho_json_destino)
        print(f"Tamanho do arquivo JSON: {tamanho_mb:.2f} MB")

        if tamanho_mb < limite_mb_max:
            print(f"O arquivo está abaixo do limite de {limite_mb_max:.2f} MB.")
        else:
            print(f"ATENÇÃO: O arquivo JSON pode ter {limite_mb_max:.2f} MB ou mais. Tente otimizar o CSV/Excel de origem.")

    except PermissionError:
        # Tratamento específico para o erro de Permissão Negada
        print("\n!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!")
        print(f"ERRO DE PERMISSÃO: O arquivo '{nome_arquivo}' não pode ser acessado.")
        print("A causa mais comum é que o arquivo está ABERTO em outro programa (como o Excel).")
        print("FECHE O ARQUIVO e execute o script novamente.")
        print("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!")
    except FileNotFoundError:
        print(f"\nERRO: Arquivo de origem não encontrado em: {caminho_origem}")
    except Exception as e:
        print(f"\nOcorreu um erro INESPERADO durante o processamento de {nome_arquivo}: {e}")


# ====================================================================
# --- BLOCO 1: PROCESSAMENTO CREDENCIAIS.xlsx (Seu código original) ---
# ====================================================================

print("=========================================================")
print("--- INICIANDO PROCESSO 1: CREDENCIAIS.xlsx ---")
print("=========================================================")

# 1. Garante que o diretório de saída exista
os.makedirs(diretorio_saida, exist_ok=True)

# Chamada da função para processar o arquivo fixo (Excel)
processar_e_salvar_json(caminho_excel_credenciais, caminho_saida_json_credenciais, limite_mb_max=30.0)


# ====================================================================
# --- BLOCO 2: PROCESSO ANÁLISE SEMANAL (Arquivo CSV Mais Recente) ---
# ====================================================================

print("\n\n=========================================================")
print(f"--- INICIANDO PROCESSO 2: ANÁLISE SEMANAL (Último CSV em {os.path.basename(diretorio_origem_fila)}) ---")
print("=========================================================")

# 1. Encontrar o último arquivo CSV no diretório da fila
caminho_ultimo_csv = encontrar_ultimo_arquivo(diretorio_origem_fila, 'csv')

if caminho_ultimo_csv:
    data_modificacao = os.path.getmtime(caminho_ultimo_csv)
    data_formatada = time.strftime('%d/%m/%Y %H:%M:%S', time.localtime(data_modificacao))
    print(f"\nArquivo CSV mais recente encontrado: {caminho_ultimo_csv}")
    print(f"Data de Última Modificação: {data_formatada}")
    print(f"Limite de tamanho para JSON: {limite_mb_analise:.2f} MB")

    # 2. Processar e salvar o JSON
    processar_e_salvar_json(caminho_ultimo_csv, caminho_saida_json_analise, limite_mb_analise)
else:
    print(f"\nAVISO: Não foram encontrados arquivos CSV (.csv) no diretório: {diretorio_origem_fila}")

print("\n--- FIM DA EXECUÇÃO DE TODOS OS PROCESSOS ---")

#dajdhajdhajdhajdadlhadd