import glob
import shutil
import pandas as pd
import requests
import os
from datetime import datetime

class FileTreatmentUseCase:
    @staticmethod
    def alteracao_cabecalho(arquivo):
        # Define o novo cabeçalho
        cabecalho = ['data', 'cod', 'tipo', 'moeda', 'Compra_1', 'Venda_1', 'Compra_2', 'Venda_2']

        # Lê o arquivo CSV
        df = pd.read_csv(arquivo, sep=";")

        # Define o novo cabeçalho no DataFrame
        df.columns = cabecalho

        # Salva o DataFrame com o novo cabeçalho de volta ao arquivo
        df.to_csv(arquivo, index=False, sep=";")

    @staticmethod
    def realizar_downloads_excel(file_path, output_directory):
        try:
            # Lê o arquivo Excel
            df = pd.read_excel(file_path)

            # Extrai os links da coluna 'Link_Download_CSV' e a coluna 'Moeda'
            links_moeda = df[['Link_Download_CSV', 'Moeda']]

            # Obtém a data atual no formato YYYYMMDD
            data_atual = datetime.now().strftime('%d%m%Y')

            # Define o nome da extração (você pode ajustar isso conforme necessário)
            nome_extracao = 'Extração -'

            # Diretório onde você deseja salvar os arquivos CSV
            diretorio_destino = os.path.join(output_directory, f'{nome_extracao} {data_atual}')

            # Cria as pastas Bronze, Silver e Gold
            diretorio_bronze = os.path.join(diretorio_destino, 'Bronze')
            diretorio_silver = os.path.join(diretorio_destino, 'Silver')
            diretorio_gold = os.path.join(diretorio_destino, 'Gold')

            os.makedirs(diretorio_bronze, exist_ok=True)
            os.makedirs(diretorio_silver, exist_ok=True)
            os.makedirs(diretorio_gold, exist_ok=True)

            # Lista para armazenar todos os DataFrames baixados
            dataframes = []

            # Itera pelos links e faz o download de cada arquivo
            for index, row in links_moeda.iterrows():
                link = row['Link_Download_CSV']
                moeda = row['Moeda']

                # Limpa o nome da moeda removendo caracteres inválidos
                nome_moeda = ''.join(c for c in moeda if c.isalnum() or c in [' ', '-', '_'])
                # Cria o diretório com o nome da moeda na pasta Bronze
                diretorio_moeda_bronze = os.path.join(diretorio_bronze, nome_moeda)
                os.makedirs(diretorio_moeda_bronze, exist_ok=True)

                try:
                    # Envia uma solicitação GET para o link
                    response = requests.get(link)
                    # Verifica se a solicitação foi bem-sucedida
                    if response.status_code == 200:
                        # Define o nome do arquivo com base na coluna 'Moeda'
                        nome_arquivo = f'{nome_moeda}.csv'
                        # Define o caminho completo para salvar o arquivo na pasta Bronze
                        caminho_arquivo_bronze = os.path.join(diretorio_moeda_bronze, nome_arquivo)
                        # Salva o conteúdo do arquivo no caminho especificado
                        with open(caminho_arquivo_bronze, 'wb') as arquivo:
                            arquivo.write(response.content)

                        print(f'Download do arquivo {nome_arquivo} concluído com sucesso na pasta Bronze.')

                        # Lê o arquivo baixado em um DataFrame e adiciona à lista
                        df_moeda = pd.read_csv(caminho_arquivo_bronze, sep=";")
                        dataframes.append(df_moeda)

                        # Copia o arquivo da pasta Bronze para a pasta Silver
                        caminho_arquivo_silver = os.path.join(diretorio_silver, nome_arquivo)
                        shutil.copyfile(caminho_arquivo_bronze, caminho_arquivo_silver)

                        print(f'Arquivo {nome_arquivo} copiado para a pasta Silver.')

                        # Chamando a função para tratar o cabeçalho do arquivo copiado para a pasta Silver
                        FileTreatmentUseCase.alteracao_cabecalho(caminho_arquivo_silver)
                        
                        # Move o arquivo concatenado sem as tratativas do Silver para a pasta Gold
                        caminho_arquivo_gold = os.path.join(diretorio_gold, nome_arquivo)
                        if os.path.exists(caminho_arquivo_gold):
                            shutil.move(caminho_arquivo_gold, caminho_arquivo_silver)
                        else:
                            print(f'O arquivo {nome_arquivo} não foi encontrado na pasta Gold.')
                    else:
                        print(f'Erro ao acessar o link {link}')

                except requests.exceptions.RequestException as req_err:
                    print(f'Erro ao fazer a solicitação GET para {link}: {str(req_err)}')

                except Exception as e:
                    print(f'Erro desconhecido: {str(e)}')

            # Concatena todos os DataFrames em um único DataFrame
            df_concatenado = pd.concat(dataframes, ignore_index=True)

            # Salva o DataFrame combinado em um arquivo CSV na pasta Silver
            arquivo_concatenado = os.path.join(diretorio_silver, 'dados_concatenados.csv')
            df_concatenado.to_csv(arquivo_concatenado, index=False)
            print(f'Dados concatenados salvos em {arquivo_concatenado} na pasta Silver.')

            # Concatenar todos os arquivos da pasta Silver em um único DataFrame
            arquivos_silver = os.listdir(diretorio_silver)
            dataframes_silver = []

            for arquivo in arquivos_silver:
                caminho_arquivo_silver = os.path.join(diretorio_silver, arquivo)
                if arquivo.endswith('.csv'):
                    df_silver = pd.read_csv(caminho_arquivo_silver, sep=";")
                    dataframes_silver.append(df_silver)

            # Concatenar os DataFrames da pasta Silver em um único DataFrame
            df_concatenado_silver = pd.concat(dataframes_silver, ignore_index=True)

            # Remove as 4 últimas colunas da direita
            df_concatenado_silver = df_concatenado_silver.iloc[:, :-1]

            # Salvar o DataFrame combinado em um arquivo CSV na pasta Gold
            arquivo_concatenado_silver = os.path.join(diretorio_gold, 'dados_concatenados_silver.csv')
            df_concatenado_silver.to_csv(arquivo_concatenado_silver, index=False)
            print(f'Todos os arquivos da pasta Silver foram concatenados e salvos em {arquivo_concatenado_silver} na pasta Gold.')

            print('Todos os downloads foram concluídos.')

        except Exception as e:
            print(f'Erro ao processar o arquivo Excel: {str(e)}')

# Exemplo de uso da função:
if __name__ == "__main__":
    arquivo_excel = 'Extracao.xlsx'
    output_directory = '/home/gabriel/cotacao_moeda'
    tratamento = FileTreatmentUseCase()
    tratamento.realizar_downloads_excel(arquivo_excel, output_directory)
