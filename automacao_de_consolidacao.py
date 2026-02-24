import pandas as pd
import os
from glob import glob
import numpy as np

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
    df["Mês"] = pd.to_datetime(df["Mês"], errors="coerce")
    return df

def trataArquivos(arquivo):
    planilhas = []
    for planilha in pd.ExcelFile(arquivo).sheet_names:
        try:
            planilhaTratada = trataPlanilha(arquivo,planilha)
            planilhas.append(planilhaTratada)
        except Exception as e:
            print(f"Erro {type(e).__name__} na planilha {planilha} do arquivo {arquivo}:\n{e}\n")
    df = pd.concat(planilhas, ignore_index=True)
    return df


def trataPastas(caminho) -> pd.DataFrame:
    pastas = os.listdir(caminho)
    dfs = []
    dfs_arquivos = []
    for i in pastas:
        arquivos = glob(os.path.join(caminho,i,"*.xlsx"))
        
        for arquivo in arquivos:
            try:
                estado = os.path.basename(arquivo)
                estado = os.path.splitext(estado)[0].split("_")[-1]
                df = trataArquivos(arquivo)
                df.insert(0,"Estado",estado)
                dfs_arquivos.append(df)
            except Exception as e:
                print(f"Erro {type(e).__name__} no arquivo {arquivo}:\n{e}\n")

    dfs = pd.concat(dfs_arquivos, ignore_index=True)
    

    return dfs

def main():
    pasta_atual = os.getcwd()
    caminho = os.path.join(pasta_atual, "operacao_logistica")
    arquivo_final = trataPastas(caminho)
    arquivo_final.to_excel("arquivo_final.xlsx", index=False)
    print(arquivo_final)

if __name__ == "__main__":
    main()
