import glob
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
    def realizar_downloads_excel(file_path):
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
            diretorio_destino = f'/home/gabriel/{nome_extracao} {data_atual}/'

            # Lista para armazenar todos os DataFrames baixados
            dataframes = []

            # Itera pelos links e faz o download de cada arquivo
            for index, row in links_moeda.iterrows():
                link = row['Link_Download_CSV']
                moeda = row['Moeda']

                # Limpa o nome da moeda removendo caracteres inválidos
                nome_moeda = ''.join(c for c in moeda if c.isalnum() or c in [' ', '-', '_'])
                # Cria o diretório com o nome da moeda se ele não existir
                diretorio_moeda = os.path.join(diretorio_destino, nome_moeda)
                os.makedirs(diretorio_moeda, exist_ok=True)

                try:
                    # Envia uma solicitação GET para o link
                    response = requests.get(link)
                    # Verifica se a solicitação foi bem-sucedida
                    if response.status_code == 200:
                        # Define o nome do arquivo com base na coluna 'Moeda'
                        nome_arquivo = f'{nome_moeda}.csv'
                        # Define o caminho completo para salvar o arquivo
                        caminho_arquivo = os.path.join(diretorio_moeda, nome_arquivo)
                        # Salva o conteúdo do arquivo no caminho especificado
                        with open(caminho_arquivo, 'wb') as arquivo:
                            arquivo.write(response.content)

                        # Chama a função para alterar o cabeçalho do arquivo baixado
                        FileTreatmentUseCase.alteracao_cabecalho(caminho_arquivo)

                        print(f'Download do arquivo {nome_arquivo} concluído com sucesso.')

                        # Lê o arquivo baixado em um DataFrame e adiciona à lista
                        df_moeda = pd.read_csv(caminho_arquivo, sep=";")
                        dataframes.append(df_moeda)
                    else:
                        print(f'Erro ao acessar o link {link}')
                except Exception as e:
                    print(f'Erro ao fazer o download do arquivo {link}: {str(e)}')

            # Concatena todos os DataFrames em um único DataFrame
            df_concatenado = pd.concat(dataframes, ignore_index=True)

            # Salva o DataFrame combinado em um arquivo CSV
            arquivo_concatenado = os.path.join(diretorio_destino, 'dados_concatenados.csv')
            df_concatenado.to_csv(arquivo_concatenado, index=False)
            print(f'Dados concatenados salvos em {arquivo_concatenado}')

            print('Todos os downloads foram concluídos.')
        except Exception as e:
            print(f'Erro ao processar o arquivo Excel: {str(e)}')

# Exemplo de uso da função:
if __name__ == "__main__":
    arquivo_excel = 'Extracao.xlsx'
    tratamento = FileTreatmentUseCase()
    tratamento.realizar_downloads_excel(arquivo_excel)
