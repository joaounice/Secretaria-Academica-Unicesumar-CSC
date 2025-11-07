import pandas as pd
import os
import glob
import time 

# Diretório de destino (target) - SEM subpasta
# O arquivo JSON de saída será salvo neste diretório
diretorio_saida = r"C:\Users\Joao.silverio\OneDrive - Uniasselvi\BI - CSC SCRA UniCesumar - Arquivos - Arquivos\RPA\site_secretaria"

# --- Configurações da TAREFA (Análise Semanal - Arquivo Mais Recente) ---

# Diretório de origem para buscar o arquivo mais recente (contém CSVs da fila)
diretorio_origem_fila = r"C:\Users\Joao.silverio\OneDrive - Uniasselvi\BI - CSC SCRA UniCesumar - Arquivos - Arquivos\FilaAtendimentos"

# Nome do arquivo JSON de saída para a Análise Semanal
nome_arquivo_json_analise = "analise semanal.json"

# Cria o caminho completo de destino para o JSON de análise
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
            # Leitura para arquivos Excel
            df = pd.read_excel(caminho_origem)
        elif extensao == 'csv':
            # Leitura para arquivos CSV
            # Tenta ler com os parâmetros comuns para CSVs brasileiros: separador ';', encoding 'latin1'
            try:
                df = pd.read_csv(caminho_origem, sep=';', encoding='latin1', low_memory=False)
            except pd.errors.ParserError:
                # Se falhar com ';', tenta o padrão ','
                df = pd.read_csv(caminho_origem, sep=',', encoding='latin1', low_memory=False)
            except UnicodeDecodeError:
                # Se falhar com latin1, tenta utf-8
                try:
                    df = pd.read_csv(caminho_origem, sep=';', encoding='utf-8', low_memory=False)
                except:
                    df = pd.read_csv(caminho_origem, sep=',', encoding='utf-8', low_memory=False)
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
# --- PROCESSO PRINCIPAL: ANÁLISE SEMANAL (Arquivo CSV Mais Recente) ---
# ====================================================================

print("=========================================================")
print(f"--- INICIANDO PROCESSO: ANÁLISE SEMANAL (Último CSV em {os.path.basename(diretorio_origem_fila)}) ---")
print("=========================================================")

# 1. Garante que o diretório de saída exista
os.makedirs(diretorio_saida, exist_ok=True)

# 2. Encontrar o último arquivo CSV no diretório da fila
caminho_ultimo_csv = encontrar_ultimo_arquivo(diretorio_origem_fila, 'csv')

if caminho_ultimo_csv:
    data_modificacao = os.path.getmtime(caminho_ultimo_csv)
    data_formatada = time.strftime('%d/%m/%Y %H:%M:%S', time.localtime(data_modificacao))
    print(f"\nArquivo CSV mais recente encontrado: {caminho_ultimo_csv}")
    print(f"Data de Última Modificação: {data_formatada}")
    print(f"Limite de tamanho para JSON: {limite_mb_analise:.2f} MB")

    # 3. Processar e salvar o JSON
    processar_e_salvar_json(caminho_ultimo_csv, caminho_saida_json_analise, limite_mb_analise)
else:
    print(f"\nAVISO: Não foram encontrados arquivos CSV (.csv) no diretório: {diretorio_origem_fila}")

print("\n--- FIM DA EXECUÇÃO DO PROCESSO ---")