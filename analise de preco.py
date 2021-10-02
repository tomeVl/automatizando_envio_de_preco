import pandas as pd
from selenium import webdriver
import time
import win32com.client as win32


def tranforma_texto(texto):
    return float(texto.replace('R$', '').replace('.', '').replace(',', '.'))


produtos = pd.read_excel(r'Produtos.xlsx')
produtos = produtos.fillna('-')
produtos = produtos[['Link Produto', 'Amazon', 'Mercado Livre', 'Casas Bahia', 'Preço Original', 'Preço Atual', 'Local']]


driver = webdriver.Chrome(executable_path=r'./chromedriver.exe')
driver.set_window_position(-10000, 0)

enviar_email = False
desconto_mini = 0.2

for i, linha in produtos.iterrows():
# pega Amazon e tratar
    driver.get(linha['Amazon'])
    time.sleep(2)
    try:
        preco_amazon = driver.find_element_by_class_name('a-color-price').text
        preco_amazon = tranforma_texto(preco_amazon)
    except:
        try:
            preco_amazon = driver.find_element_by_id('priceblock_ourprice').text
            preco_amazon = tranforma_texto(preco_amazon)
        except:
            print('Produto {} não Disponivel na Amazon'.format(linha['Link Produto']))
            preco_amazon = linha['Preço Original'] * 3

    time.sleep(2)

# pega Mercado livre e tratar
    driver.get(linha['Mercado Livre'])
    time.sleep(2)
    try:
        preco_mercado_livre = driver.find_element_by_class_name("price-tag-fraction").text
        preco_mercado_livre = tranforma_texto(preco_mercado_livre)
    except:
        print('Produto {} não Disponivel na Mercado Livre'.format(linha['Link Produto']))
        preco_mercado_livre = linha['Preço Original'] * 3

# pega Casas Bahia e tratar
    driver.get(linha['Casas Bahia'])
    time.sleep(2)
    try:
        preco_bahia = driver.find_element_by_id('product-price').text
        preco_bahia = tranforma_texto(preco_bahia)
    except:
        print('Produto {} não Disponivel na Casas Bahia'.format(linha['Link Produto']))
        preco_bahia = linha['Preço Original'] * 3

# print(preco_amazon, preco_mercado_livre, preco_bahia)
    preco_original = linha['Preço Original']

    lista_preco = [(preco_amazon, 'Amazon'), (preco_mercado_livre, 'Mercado Livre'), (preco_bahia, 'Casas Bahia'),
                   (preco_original, 'Preço Original')]

    lista_preco.sort()

    produtos.loc[i, 'Preço Atual'] = lista_preco[0][0]
    produtos.loc[i, 'Local'] = lista_preco[0][1]

    if lista_preco[0][0] <= preco_original * (1 - desconto_mini):
        enviar_email = True

outlook = win32.Dispatch('outlook.application')

# Salva O arquivo

produtos.to_excel('Produtos.xlsx')

# ENVIAR EMAIL

if enviar_email:
    mail = outlook.CreateItem(0)
    mail.To = 'rubenstome15@gmail.com'
    mail.Subject = f'Iphone 12 com {desconto_mini:.0%} De DESCONTO'

# FILTRA A TABELA DE PRODUTOS
    tabela_filtrada = produtos.loc[produtos['Preço Atual'] <= produtos['Preço Original'] * (1 - desconto_mini), :]

    mail.HTMLBody = f'<p>Esse é o Iphone 12 com {desconto_mini:.0%} de DESCONTO</p>{tabela_filtrada.to_html()}'

    mail.Send()

print('Fim da Análise')
driver.quit()
