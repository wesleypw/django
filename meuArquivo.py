import pandas as pd 
import win32com.client as win32


#importar a base de dados
tabela_de_vendas = pd.read_excel('Vendas.xlsx')


pd.set_option('display.max_columns', None)
print(tabela_de_vendas)

faturamento = tabela_de_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()

print(faturamento)

quantidade = tabela_de_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()

print(quantidade)

#visualizar a base de dados

# faturamento por loja 

#quantidade de produtos vendidos por loja

#ticket medio por produto em cada loja

print("-" * 50)

ticket_medio = (faturamento['Valor Final'] / quantidade['Quantidade']).to_frame()
ticket_medio = ticket_medio.rename(columns={0: 'Ticket medio'})
print(ticket_medio)

#enviear um email com o relatorio

outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'to address'
mail.Subject = 'Massege subject'
mail.HTMLBody =f'''

<h2>HTML Message body</h2>

<p> esse conteudo é uma mistura de html e python</p>

<p> que envia um email"outlook" e enviando</p>


'''


mail.Send()
