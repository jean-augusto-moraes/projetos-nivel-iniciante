import pandas as pd

tabela = pd.read_excel('Tabela 6588.xlsx', skiprows=[0,1,2,3])

indice_coluna = 5
tabela = tabela.drop(indice_coluna)
tabela = tabela.rename(columns={'Unnamed: 0': 'Mês'})

pd.options.display.float_format = '{:,.0f}'.format

tabela = tabela.transpose()
#tabela[['Mês','Ano']] = tabela[' '].str.split(' ', expand=True)

display(tabela)

tabela.to_excel('tabela_tratada.xlsx')
