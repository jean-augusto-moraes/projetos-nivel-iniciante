import pandas as pd

tabela = pd.read_excel('tabela_tratada.xlsx' , skiprows=[0])

tabela = tabela.to_series()

tabela[['Mês','Ano']] = tabela['Mês'].str.split(' ', expand=True)
coluna = tabela.pop('Ano')
tabela.insert(1, 'Ano', coluna)
tabela = tabela.style.set_properties(**{'text-align': 'center'})
tabela.set_table_styles([{'selector': 'th', 'props': [('text-align', 'center')]}])

display(tabela)
tabela.to_excel('tabela_tratada_v2.xlsx')
