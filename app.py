"""
PRECISO AUTOMATIZAR MINHAS MENSAGENS P/ MEUS CLIENTES GOSTARIA DE SABER VALORES, E GOSTARIA QUE ESNTRASSEM
EM CONTATO COMIGO P/ EXPLICAR MELHOR, QUERO PODER MANDAR MENSAGENS DE COBRANÇA EM DETERMINADO DIA COM CLIENTES
COM FENCIMENTO DIFERENTE

"""

import openpyxl
from urllib.parse import quote
import webbrowser
from time import sleep

webbrowser.open('https://web.whatsapp.com')
sleep(10)
# ler planilha e guardar informações sobre nome, telefone e data de vencimento

workbook = openpyxl.load_workbook('clientes.xlsx')
pagina_clientes = workbook['Sheet1']


for linha in pagina_clientes.iter_rows( min_col=1 , max_col=12 , min_row=2 , max_row=2 , values_only=False):
   #nome, telefone, vencimento
    nome = linha[0].value 
    telefone = linha[1].value
    vencimento = linha[2].value

print(nome)
print(telefone)
print(vencimento)    
    # print('teste')


mensagem = f'Olá {nome} seu boleto vence no dia {vencimento.strftime("%d/%m/%Y")}. favor pagar no link https://www.link_do_pagamento.com'

# Criar links personalizados do whatsapp e enviar mensagens para cada cliente
link_mensagem_whatsapp = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'

    
# com base nos dados da planilha
webbrowser.open(link_mensagem_whatsapp)

input('')