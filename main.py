from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import pandas as pd
from IPython.display import display

navegador = webdriver.Chrome(r'C:\Program Files\JetBrains\PyCharm Community Edition 2020.1.2\bin\chromedriver.exe')

# pesquisar a cotação do dolar
navegador.get('https://www.google.com.br')
# pegar a cotação do dolar
navegador.find_element_by_xpath(r'/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys(
    'cotação dolar')
navegador.find_element_by_xpath(r'/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys(
    (Keys.ENTER))
cotacao_dolar = navegador.find_element_by_xpath(
    r'//*[@id="knowledge-currency__updatable-data-column"]/div[1]/div[2]/span[1]').get_attribute('data-value')
print(f'A cotação do dolar é: {cotacao_dolar}')

# pesquisar a cotação do euro
navegador.get('https://www.google.com.br')
# pegar a cotação do euro
navegador.find_element_by_xpath(r'/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input') \
    .send_keys('cotacao euro')
navegador.find_element_by_xpath(r'/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input') \
    .send_keys(Keys.ENTER)
cotacao_euro = navegador.find_element_by_xpath(r'//*[@id="knowledge-currency__updatable-data-column"]/div[1]/div[2]/span[1]')\
    .get_attribute('data-value')
print(f'A cotação do Euro é: {cotacao_euro}')

# pesquisar o preço do ouro
site = 'https://www.melhorcambio.com/ouro-hoje'
# Caso o site melhor cambio não estivesse na aba "ouro", o código para clicar no link do outro seria o seguinte:
# navegador.find_element_by_xpath('XPATH DO LINK DO OURO').click()
navegador.get(site)

# pegar a cotação do ouro
cotacao_ouro = navegador.find_element_by_xpath('//*[@id="comercial"]').get_attribute('value')
cotacao_ouro = cotacao_ouro.replace(',', '.')
print(f'A cotação do ouro é: {cotacao_ouro}')

# sou houvesse download de algum arquivo, pode-se utilizar time.sleep(tempo em segundos)

# importar os dados
pd.set_option('display.max_columns', None)
tabela_produtos = pd.read_excel('Produtos.xlsx')
# print(tabela_produtos)

# Atualizar as cotações  -  tabela_produtos.loc(linha, coluna)
tabela_produtos.loc[tabela_produtos['Moeda'] == 'Dólar', 'Cotação'] = float(cotacao_dolar)
tabela_produtos.loc[tabela_produtos['Moeda'] == 'Euro', 'Cotação'] = float(cotacao_euro)
tabela_produtos.loc[tabela_produtos['Moeda'] == 'Ouro', 'Cotação'] = float(cotacao_ouro)
# print(tabela_produtos)

# Atualizar os preços na tabela
tabela_produtos['Preço Base Reais'] = tabela_produtos['Preço Base Original'] * tabela_produtos['Cotação']
tabela_produtos['Preço Final'] = tabela_produtos['Preço Base Reais'] * tabela_produtos['Ajuste']
print(tabela_produtos)

# Exportando em um novo arquivo
tabela_produtos.to_excel('Produtos Atualizados.xlsx', index=False)  # index = False retira a primeira coluna de index

