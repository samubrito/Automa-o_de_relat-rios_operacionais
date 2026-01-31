import pandas as pd
import os
import glob
import numpy as np

caminho = r"C:\Users\samue\Desktop\Operação Logística"

def converteNumero(df:pd.DataFrame,colunas=[]) -> pd.DataFrame:
    for coluna in colunas:
        df[coluna] = (df[coluna].replace("NaN", 0).astype(str).str.replace(",", ".", regex=False)
        .str.replace(r"[^\d.]", "", regex=True).replace("", np.nan)
        .astype(float).replace(np.nan, 0))
    
    return df
    

def trataPlanilha(nome_arquivo,nome_planilha):
    df = pd.read_excel(nome_arquivo, sheet_name=nome_planilha)
    mes = df.iloc[0,1]
    df = df.drop(df.index[0:6]).T
    df.columns = df.iloc[0]
    df = df.drop(df.index[0])
    df = converteNumero(df,["Tempo_h","Km","Custo"])
    df = df.groupby("Equipe", as_index=False).sum()

    df.insert(0,"Filial",nome_planilha)
    df.insert(0,"Mês",mes)
    return df

def trataArquivos(arquivo):
    planilhas = []
    for planilha in pd.ExcelFile(arquivo).sheet_names:
        planilhaTratada = trataPlanilha(arquivo,planilha)
        planilhas.append(planilhaTratada)
    df = pd.concat(planilhas, ignore_index=True)
    return df


def TrataPastas(camihno):
    pastas = os.listdir(caminho)
    dfs = []
    dfs_arquivos = []
    for i in pastas:
        arquivos = glob.glob(os.path.join(fr"{caminho}\{i}","*.xlsx"))
        
        for arquivo in arquivos:
            nome_arquivo = arquivo
            df = trataArquivos(arquivo)
            df.insert(0,"Estado",nome_arquivo)
            dfs_arquivos.append(df)
            
    dfs = pd.concat(dfs_arquivos, ignore_index=True)
    return dfs

print(TrataPastas(caminho))