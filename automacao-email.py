import pandas as pd

import win32com.client as win23


# importar a base de dados
tabela_vendas = pd.read_excel('Vendas.xlsx')


# vizualizar a base de dados
pd.set_option('display.max_columns', None)


# faturamento por loja
faturamento = tabela_vendas[['ID Loja', 'Valor Final']].groupby(
    'ID Loja').sum()

print(faturamento)
# quantidade de produtos vendido por loja
quantidade = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()

print(quantidade)
# ticket medio por produto em cada loja
ticket_medio = (faturamento['Valor Final'] /quantidade['Quantidade']).to_frame()
ticket_medio = ticket_medio.rename(columns={0: 'ticket medio'})
print(ticket_medio)
# enviar um email com relatorio

outlook = win23.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'gabrielicristina276@gmail.com'
mail.Subject = 'relatorio de vendas por loja'
mail.HTMLbody = f'''
<p>prezados,</p>

<p>segue o relatorio de vendas por cada loja.</p>

<p>faturamento:</p>
{faturamento.to_html(formatters={'Valor Final':'R${:,.2f}'.format})}

<p>quantidade vendida:</p>
{quantidade.to_html()}

<p>ticket medio dos produtos em cada loja:</p>
{ticket_medio.to_html(formatters={'ticket medio':'R${:,.2f}'.format})}

<p>quarquer duvida estou a disposicao.</p>

<p>att.,</p>
<p>jean_volkoff</p>
'''

mail.Send()
print('email enviado')
