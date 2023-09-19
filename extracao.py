from playwright.sync_api import sync_playwright
import pandas as pd
import datetime

# Função para acessar o site do Banco Central, buscar a cotação das moedas e salvar em um arquivo Excel
def acessar_site_banco_central_e_salvar_cotacao():
    # Inicialize o Playwright
    with sync_playwright() as p:
        # Abra uma nova instância do navegador
        browser = p.chromium.launch()
        # Crie uma nova página no navegador
        page = browser.new_page()

        # Acesse o site do Banco Central
        page.goto('https://www.bcb.gov.br/conversao')

        # Espere até que a página carregue completamente (você pode ajustar esse tempo)
        page.wait_for_selector('.sabct220 > table')

        # Extraia a tabela de cotação das moedas
        tabela_moedas = page.query_selector('.sabct220 > table')
        cotacao_moedas = tabela_moedas.inner_text()

        # Feche o navegador
        browser.close()

        # Transforme os dados de texto em um DataFrame do pandas
        df = pd.read_html(cotacao_moedas)[0]

        # Adicione uma coluna de data ao DataFrame com a data atual
        data_atual = datetime.date.today()
        df['Data'] = data_atual

        # Verifique se o arquivo Excel já existe
        try:
            df_existente = pd.read_excel('cotacao_moedas.xlsx')
            # Concatene o DataFrame existente com os novos dados
            df_concatenado = pd.concat([df_existente, df], ignore_index=True)
            # Salve o DataFrame concatenado em um arquivo Excel
            df_concatenado.to_excel('cotacao_moedas.xlsx', index=False)
            print("Dados concatenados e salvos em 'cotacao_moedas.xlsx'.")
        except FileNotFoundError:
            # Se o arquivo Excel não existir, crie um novo
            df.to_excel('cotacao_moedas.xlsx', index=False)
            print("Dados salvos em 'cotacao_moedas.xlsx'.")

# Chame a função para acessar o site, buscar a cotação das moedas e salvar no arquivo Excel
acessar_site_banco_central_e_salvar_cotacao()
