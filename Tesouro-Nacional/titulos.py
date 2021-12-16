"""
@author: Fernando Nast
"""
import pandas as pd

############################### TITULO LFT DE 2002/2020 ###############################
def table_LFT():
    LFT = pd.DataFrame()

    for tab in range(2002,2021):
        filepath = 'C://Users//Fernando Nast//turma-2020-2//data//intermediate//tesouro//LFT_'+str(tab)+'.xls'
        lft = pd.ExcelFile(filepath)
    
        lft_0 = pd.DataFrame()                                           # Inicializando Dataframes
        lft_1 = pd.DataFrame()                          
    
        for aba in lft.sheet_names:
            lft_0 = pd.read_excel(filepath,header=None,sheet_name=aba)   # Leitura do arquivo, considerando as abas
            lft_0['Maturity'] = pd.to_datetime(lft_0[1][0])              # Adicionando a coluna Maturity agora pois varia de aba para aba
            lft_1 = pd.concat([lft_1,lft_0],axis=0)                      # Concatenando as as abas da tabela
        
        LFT = pd.concat([LFT,lft_1],axis=0,sort=False)                   # Concatenando as tabelas
    
    LFT = LFT.drop([0,1])                                                # Apagando as duas primeiras linhas, pois as informações ja foram retiradas
    LFT['Data'] = pd.to_datetime(LFT[0])                                 # Adicionando a coluna DATA do tipo Datetime
    LFT = LFT.drop(0,axis=1)                                             # Retirando a coluna 0, pois foi substituida pela DATA
    LFT['Year'] = LFT['Data'].dt.year                                    # Acrescentando a coluna YEAR, também com tipo Datetime
    LFT['Bond'] = filepath[-12:-9]                                       # Acrescentando coluna BOND com o nome do titulo
    
    return LFT

############################### TITULO LTN DE 2002/2020 ###############################
def table_LTN():
    LTN = pd.DataFrame()
    
    for tab in range(2002,2021):
        filepath_ltn = 'C://Users//Fernando Nast//turma-2020-2//data//intermediate//tesouro//LTN_'+str(tab)+'.xls'
        ltn = pd.ExcelFile(filepath_ltn)
    
        ltn_0 = pd.DataFrame()
        ltn_1 = pd.DataFrame()
    
        for aba in ltn.sheet_names:
            ltn_0 = pd.read_excel(filepath_ltn,header=None,sheet_name=aba)     
            ltn_0['Maturity'] = pd.to_datetime(ltn_0[1][0])
            ltn_1 = pd.concat([ltn_1,ltn_0],axis=0)                       
        
        LTN = pd.concat([LTN,ltn_1],axis=0,sort=False)
    
    LTN = LTN.drop([0,1])
    LTN['Data'] = pd.to_datetime(LTN[0])  
    LTN = LTN.drop(0,axis=1)
    LTN['Year'] = LTN['Data'].dt.year                 
    LTN['Bond'] = filepath_ltn[-12:-9]
    
    return LTN

############################### TITULO NTN-B DE 2003/2020 ###############################
def table_NTNb():
    NTNb = pd.DataFrame()
    
    for tab in range(2003,2021):
        filepath_ntnb = 'C://Users//Fernando Nast//turma-2020-2//data//intermediate//tesouro//NTN-B_'+str(tab)+'.xls'
        ntnb = pd.ExcelFile(filepath_ntnb)
    
        ntnb_0 = pd.DataFrame()
        ntnb_1 = pd.DataFrame()
    
        for aba in ntnb.sheet_names:
            ntnb_0 = pd.read_excel(filepath_ntnb,header=None,sheet_name=aba) 
            ntnb_0['Maturity'] = pd.to_datetime(ntnb_0[1][0])
            ntnb_1 = pd.concat([ntnb_1,ntnb_0],axis=0)                             
        
        NTNb = pd.concat([NTNb,ntnb_1],axis=0,sort=False)
    
    NTNb = NTNb.drop([0,1])
    NTNb['Data'] = pd.to_datetime(NTNb[0])  
    NTNb = NTNb.drop(0,axis=1)
    NTNb['Year'] = NTNb['Data'].dt.year                 
    NTNb['Bond'] = filepath_ntnb[-14:-9]
    
    return NTNb

############################### TITULO NTN-B-Principal DE 2005/2020 ###############################
def table_NTNbp():
    NTNbp = pd.DataFrame()
    
    for tab in range(2005,2021):
        filepath_ntnbp = 'C://Users//Fernando Nast//turma-2020-2//data//intermediate//tesouro//NTN-B_Principal_'+str(tab)+'.xls'
        ntnbp = pd.ExcelFile(filepath_ntnbp)
    
        ntnbp_0 = pd.DataFrame()
        ntnbp_1 = pd.DataFrame()
    
        for aba in ntnbp.sheet_names:
            ntnbp_0 = pd.read_excel(filepath_ntnbp,header=None,sheet_name=aba)     
            ntnbp_0['Maturity'] = pd.to_datetime(ntnbp_0[1][0])
            ntnbp_1 = pd.concat([ntnbp_1,ntnbp_0],axis=0)                       
        
        NTNbp = pd.concat([NTNbp,ntnbp_1],axis=0,sort=False)
    
    NTNbp = NTNbp.drop([0,1])
    NTNbp['Data'] = pd.to_datetime(NTNbp[0])  
    NTNbp = NTNbp.drop(0,axis=1)
    NTNbp['Year'] = NTNbp['Data'].dt.year                 
    NTNbp['Bond'] = filepath_ntnbp[-24:-9]
    
    return NTNbp

############################### TITULO NTN-C DE 2008/2020 ###############################
def table_NTNc():
    NTNc = pd.DataFrame()
    
    for tab in range(2002,2021):
        filepath_ntnc = 'C://Users//Fernando Nast//turma-2020-2//data//intermediate//tesouro//NTN-C_'+str(tab)+'.xls'
        ntnc = pd.ExcelFile(filepath_ntnc)
    
        ntnc_0 = pd.DataFrame()
        ntnc_1 = pd.DataFrame()
    
        for aba in ntnc.sheet_names:
            ntnc_0 = pd.read_excel(filepath_ntnc,header=None,sheet_name=aba)
            ntnc_0['Maturity'] = pd.to_datetime(ntnc_0[1][0])
            ntnc_1 = pd.concat([ntnc_1,ntnc_0],axis=0)                  
        
        NTNc = pd.concat([NTNc,ntnc_1],axis=0,sort=False)
    
    NTNc = NTNc.drop([0,1])
    NTNc['Data'] = pd.to_datetime(NTNc[0])  
    NTNc = NTNc.drop(0,axis=1)
    NTNc['Year'] = NTNc['Data'].dt.year                 
    NTNc['Bond'] = filepath_ntnc[-14:-9]
    
    return NTNc

############################### TITULO NTN-F DE 2004/2020 ###############################
def table_NTNf():
    NTNf = pd.DataFrame()
    
    for tab in range(2004,2021):
        filepath_ntnf = 'C://Users//Fernando Nast//turma-2020-2//data//intermediate//tesouro//NTN-F_'+str(tab)+'.xls'
        ntnf = pd.ExcelFile(filepath_ntnf)
    
        ntnf_0 = pd.DataFrame()
        ntnf_1 = pd.DataFrame()
    
        for aba in ntnf.sheet_names:
            ntnf_0 = pd.read_excel(filepath_ntnf,header=None,sheet_name=aba) 
            ntnf_0['Maturity'] = pd.to_datetime(ntnf_0[1][0])
            ntnf_1 = pd.concat([ntnf_1,ntnf_0],axis=0)                     
        
        NTNf = pd.concat([NTNf,ntnf_1],axis=0,sort=False)
    
    NTNf = NTNf.drop([0,1])
    NTNf['Data'] = pd.to_datetime(NTNf[0])  
    NTNf = NTNf.drop(0,axis=1)
    NTNf['Year'] = NTNf['Data'].dt.year                 
    NTNf['Bond'] = filepath_ntnf[-14:-9]
    
    return NTNf

############################### TABELA FINAL ###############################
def table_final(a,b,c,d,e,f):
    final = pd.concat([a,b,c,d,e,f],axis=0,sort=False)
    
    final = final.rename(columns={1:'Taxa Compra Manhã',2:'Taxa Venda Manhã',               # Renomeando as colunas
                              3:'PU Compra Manhã',4:'PU Venda Manhã',5:'PU Base Manhã'})
    
    final = final.set_index('Data')                                                         # Indexando a coluna DATA
    
    final = final[['Taxa Compra Manhã', 'Taxa Venda Manhã', 'PU Compra Manhã',              # Organizando a ordem das colunas
                'PU Venda Manhã', 'PU Base Manhã', 'Maturity', 'Year', 'Bond']]
    
    return final