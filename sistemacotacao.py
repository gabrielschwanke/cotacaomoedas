import tkinter as tk
from tkinter import ttk
from tkcalendar import DateEntry
from tkinter.filedialog import askopenfilename
import pandas as pd
import requests
from datetime import datetime
import numpy as np

#para pegar a cotação das moedas, vou usar a API do awesomeapi
requisicao = requests.get('https://economia.awesomeapi.com.br/json/all')
dicionario_moedas = requisicao.json()#para transformar o dicionario json em dicionario py
#print(dicionario_moedas)
lista_moedas = list(dicionario_moedas.keys())#keys para pegar as chaves do dicionario

def pegar_cotacao():
    moeda = combobox_selecionar_moeda.get()
    data_cotacao = calendario_moeda.get()
    ano = data_cotacao[-4:]
    mes = data_cotacao[3:5]
    dia = data_cotacao[:2]#pegando até o indice dois mas não pega o ultimo indice
    link = f'https://economia.awesomeapi.com.br/json/daily/{moeda}-BRL/?start_date={ano}{mes}{dia}&end_date={ano}{mes}{dia}'
    requisicao_moeda = requests.get(link)
    cotacao = requisicao_moeda.json()
    valor_moeda = cotacao[0]['bid']#uma lista de um item só, por isso pego o indice [0] item ['bid']
    label_texto_cotacao['text'] = f'A cotação da {moeda} no dia {data_cotacao} foi de: R${valor_moeda}'

def selecionar_arquivo():
    caminho_arquivo = askopenfilename(title='Selecione o Arquivo de Moeda')
    var_caminho_arquivo.set(caminho_arquivo)#mudando a variavel dentro da função
    if caminho_arquivo:
        label_arquivo_selecionado['text'] = f'Arquivo Selecionado: {caminho_arquivo}'


def atualizar_cotacoes():
    try:
        #ler o dataframe de moedas
        df = pd.read_excel(var_caminho_arquivo.get())
        moedas = df.iloc[:, 0]#iloc para localizar, [linhas, colunas] no iloc passamos os indices
        #pegando as informações de tada
        data_inicial = calendario_datainicial.get()
        data_final = calendario_datafinal.get()
        ano_inicial = data_inicial[-4:]
        mes_inicial = data_inicial[3:5]
        dia_inicial = data_inicial[:2]

        ano_final = data_final[-4:]
        mes_final = data_final[3:5]
        dia_final = data_final[:2]

        for moeda in moedas:
            link = f'https://economia.awesomeapi.com.br/json/daily/{moeda}-BRL/?' \
                   f'start_date={ano_inicial}{mes_inicial}{dia_inicial}&end_date={ano_final}{mes_final}{dia_final}'
            requisicao_moeda = requests.get(link)
            cotacoes = requisicao_moeda.json()
            for cotacao in cotacoes:
                timestamp = int(cotacao['timestamp'])#para pegar a data da cotação e a informação do timestamp é um texto, por isso tem que transformar em numero
                bid = float(cotacao['bid'])
                data = datetime.fromtimestamp(timestamp)
                data = data.strftime('%d/%m/%Y')
                if data not in df:
                    df[data] = np.nan

                df.loc[df.iloc[:, 0] == moeda, data] = bid

        df.to_excel('Teste.xlsx')
        label_atualizar_cotacoes['text'] = 'Arquivo Atualizado com Sucesso'

    except:
        label_atualizar_cotacoes['text'] = 'Selecione um arquivo excel no formato correto'

janela = tk.Tk()
janela.title('Ferramenta de Cotação de Moedas')
janela.config(bg='#C0C0C0')

label_cotacao_moeda = tk.Label(text='Cotação de uma moeda especifica', borderwidth=2, relief='solid')
label_cotacao_moeda.grid(row=0, column=0, padx=10, pady=10, sticky='nsew', columnspan=3)
#padx e pady é a distância entre a mensagem e a borda
#columnspan para pegar três colunas
label_selecionar_moeda = tk.Label(text='Selecionar Moeda', borderwidth=2, anchor='e')
label_selecionar_moeda.grid(row=1, column=0, padx=10, pady=10, sticky='nsew', columnspan=2)
label_selecionar_moeda.config(bg='#C0C0C0')

combobox_selecionar_moeda = ttk.Combobox(values=lista_moedas)
combobox_selecionar_moeda.grid(row=1, column=2, padx=10, pady=10, sticky='nsew')

label_selecionar_dia = tk.Label(text='Selecione o dia que deseja pegar a cotação', borderwidth=2,anchor='e')
label_selecionar_dia.grid(row=2, column=0, padx=10, pady=10, sticky='nsew', columnspan=2)
label_selecionar_dia.config(bg='#C0C0C0')

calendario_moeda = DateEntry(year=2022, locale='pt_br')
calendario_moeda.grid(row=2, column=2, padx=10, pady=10, sticky='nsew')

label_texto_cotacao = tk.Label(text='')
label_texto_cotacao.grid(row=3, column=0, columnspan=2, padx=10, pady=10, sticky='nsew')

botao_pegar_cotacao = tk.Button(text='Pegar Cotação', command=pegar_cotacao)
botao_pegar_cotacao.grid(row=3, column=2, padx=10, pady=10, sticky='nsew')

#multiplas moedas
label_cotacaovariasmoedas = tk.Label(text='Cotação Múltiplas Moedas', borderwidth=2, relief='solid')
label_cotacaovariasmoedas.grid(row=4, column=0, padx=10, pady=10, sticky='nsew', columnspan=3)

label_selecionar_arquivo = tk.Label(text='Selecione um arquivo em Excel com as moedas na coluna A')
label_selecionar_arquivo.grid(row=5, column=0, columnspan=2, padx=10, pady=10, sticky='nsew')
label_selecionar_arquivo.config(bg='#C0C0C0')

var_caminho_arquivo = tk.StringVar()#criando a variavel fora da função para poder utilizar em outras funções

botao_selecionar_arquivo = tk.Button(text='Clique para Selecionar', command=selecionar_arquivo)
botao_selecionar_arquivo.grid(row=5, column=2, padx=10, pady=10, sticky='nsew')

label_arquivo_selecionado = tk.Label(text='Nenhum Arquivo Selecionado', anchor='e')
label_arquivo_selecionado.grid(row=6, column=0, columnspan=3, padx=10, pady=10, sticky='nsew')

label_data_inicial = tk.Label(text='Data Inicial', anchor='e')
label_data_final = tk.Label(text='Data Final', anchor='e')
label_data_inicial.grid(row=7, column=0, padx=10, pady=10, sticky='nsew')
label_data_final.grid(row=8, column=0, padx=10, pady=10, sticky='nsew')
label_data_inicial.config(bg='#C0C0C0')
label_data_final.config(bg='#C0C0C0')


calendario_datainicial = DateEntry(year=2022, locale='pt_br')
calendario_datafinal = DateEntry(year=2022, locale='pt_br')
calendario_datainicial.grid(row=7, column=1, padx=10, pady=10, sticky='nsew')
calendario_datafinal.grid(row=8, column=1, padx=10, pady=10, sticky='nsew')

botao_atualizar_cotacoes = tk.Button(text='Atualizar Cotações', command=atualizar_cotacoes)
botao_atualizar_cotacoes.grid(row=9, column=0, padx=10, pady=10, sticky='nsew')

label_atualizar_cotacoes = tk.Label(text='')
label_atualizar_cotacoes.grid(row=9, column=1, columnspan=2, padx=10, pady=10, sticky='nsew')

botao_fechar = tk.Button(text='Fechar', command=janela.quit)
botao_fechar.grid(row=10, column=2, padx=10, pady=10, sticky='nsew')

janela.mainloop()