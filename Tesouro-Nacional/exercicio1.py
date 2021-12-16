"""
@author: Fernando Nast
"""
import titulos

############################## UNINDO AS ABAS DAS RESPECTIVAS TABELAS E CRIANDO APENAS UMA
lft = titulos.table_LFT()                                   
ltn = titulos.table_LTN()
ntnb = titulos.table_NTNb()
ntnbp = titulos.table_NTNbp()
ntnc = titulos.table_NTNc()
ntnf = titulos.table_NTNf()


############################# UNINDO TODAS AS TABELAS E CRIANDO APENAS UMA
dataset = titulos.table_final(lft,ltn,ntnb,ntnbp,ntnc,ntnf)


############################# SALVANDO O DATASET EM ARQUIVO
dataset.to_csv('dataset.csv')