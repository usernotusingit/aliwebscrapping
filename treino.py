from tkinter import *
from tkinter import filedialog, messagebox
import time
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import pandas as pd
from datetime import date
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


def webscraping():
    options = webdriver.ChromeOptions()
    options.add_experimental_option("detach", True)
    servico = Service(ChromeDriverManager().install())

    
    navegador = webdriver.Chrome(options=options, service=servico)
    navegador.get("https://pt.aliexpress.com/")

    search_input = WebDriverWait(navegador,10).until(
        EC.element_to_be_clickable((By.XPATH, '/html/body/div[7]/div/div[2]/div[3]/div[1]'))
    )

    navegador.find_element(By.XPATH ,'/html/body/div[7]/div/div[2]/div[3]/div[1]').click()


    navegador.find_element(by=By.ID, value="search-words").send_keys(item.get(), Keys.ENTER)

    if int(preco_max.get()) == 0:
        pass
    else:
        navegador.find_element(By.XPATH,
                               '//*[@id="root"]/div[1]/div/div[2]/div/div[2]/div[1]/div/div[1]/div[1]/div[1]/input[1]').send_keys(
            preco_min.get())
        navegador.find_element(By.XPATH,
                               '//*[@id="root"]/div[1]/div/div[2]/div/div[2]/div[1]/div/div[1]/div[1]/div[1]/input[2]').send_keys(
            preco_max.get())
        navegador.find_element(By.CLASS_NAME, 'priceInput--ok--2apR64x').click()

        while len(navegador.find_elements(By.CLASS_NAME, 'priceInput--ok--2apR64x')) >= 1:
            time.sleep(1)
    x = 0
    while True:
        x += 1
        navegador.execute_script(f"window.scrollBy(0, 300)")
        time.sleep(0.5)
        if x > 100:
            break

    lista_elementos = navegador.find_elements(By.CLASS_NAME, 'manhattan--container--1lP57Ag')
    if len(lista_elementos) < 1:
        messagebox.showerror(title='Erro', message='Produto não encontrado, tente buscar sem filtros')
        tela.destroy()
    else:
        pass
    produtos = []
    for elemento in lista_elementos:
        nome = elemento.find_element(By.CLASS_NAME, 'cards--title--2rMisuY').text
        preco = elemento.find_element(By.CLASS_NAME, 'manhattan--price-sale--1CCSZfK').text
        try:
            n_vendas = elemento.find_element(By.CLASS_NAME, 'manhattan--trade--2PeJIEB').text
        except:
            n_vendas = '0 vendas'
        link = elemento.get_attribute('href')

        produtos.append([nome, preco, n_vendas, link])

    tabela_produtos = pd.DataFrame(produtos, columns=['Produto', 'Preço', 'Numero de Vendas', 'Link'])
    pasta = filedialog.askdirectory()
    hoje = date.today().strftime("%d-%m-%Y")
    nome_arquivo = f'{item.get()}-{hoje}'
    tabela_produtos.to_excel(f'{pasta}\%s.xlsx' % nome_arquivo, 'w')


tela = Tk()

tela.title('Sistema de web scraping - Aliexpress')

tela.rowconfigure(0, weight=1)
tela.columnconfigure([0, 1], weight=1)


nome_programa = Label(text='Web Scraping - Aliexpress', fg='white', bg='black', width=20, height=3)
nome_programa.grid(row=0, column=0, columnspan=2, sticky='NSEW')

mensagem_produto = Label(text='Digite o produto desejado: ')
mensagem_produto.grid(row=1, column=0)

item = Entry()
item.grid(row=1, column=1)

mensagem_min = Label(text='Valor mínimo:')
preco_min = Entry()

mensagem_max = Label(text='Valor máximo:')
preco_max = Entry()

mensagem_min.grid(row=2, column=0)
preco_min.grid(row=2, column=1)

mensagem_max.grid(row=3, column=0)
preco_max.grid(row=3, column=1)

filtro_zero = Label(text='Valor máx = 0 para não filtrar preço')

buscar = Button(text='Buscar', command=webscraping)
buscar.grid(row=4, column=1)

filtro_zero.grid(row=4, column=0)

tela.mainloop()