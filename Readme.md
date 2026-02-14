<h1>Automação de Consolidação: Relatório Operacional Logístico</h1>
<p align="center">
  <img src="automacao_relatorios.png" alt="Automação de Consolidação: Relatório Operacional Logístico" width="300px">
</p>
<p>Este projeto consiste em um pipeline de ETL (Extração, Transformação e Carga) desenvolvido em Python para consolidar múltiplos relatórios de operação logística. O script é ideal para cenários onde os dados estão distribuídos em diversas pastas por região, arquivos por estado e abas por filial.</p>

<details open="open">

<summary><h2>📋 Sumário</h2></summary>

- [1. Visão Geral](#visao_geral)
- [2. Estrutura de Dados](#estrutura)
- [3. Arquitetura do Processamento](#arquitetura)
- [4. Requisitos](#requisitos)
- [5. Como Executar](#executar)
- [6. Tratamento de Erros](#tratamento)

</details>

<h2 id="visao_geral">1. Visão Geral 🔍</h2>
<p>A automação resolve o problema de relatórios manuais que possuem cabeçalhos complexos e formatação inconsistente. O código varre diretórios, entra em cada aba de cada arquivo Excel, limpa os dados numéricos e gera um arquivo mestre consolidado para análise em BI ou Dashboards.
</p>
<p align="center">
  <img src="automacao_relatorios.png" alt="Automação de Consolidação: Relatório Operacional Logístico" width="400px">
</p>

<h2 id="estrutura">2. Estrutura de Dados 📊</h2>
O script espera uma hierarquia específica para atribuir os metadados corretamente:

<ul>
    <li>Estado: Extraído da última parte do nome do arquivo (ex: Relatorio_SP.xlsx → "SP").</li>
    <li>Mês: Extraído da célula B1 (índice 0,1) de cada aba (Formato esperado: Texto, ex: "Janeiro").</li>
    <li>Filial: Extraído automaticamente do nome da aba (sheet name).</li>
    <li>Colunas Métricas: Tempo_h, Km e Custo.</li>
</ul>

<h2 id="arquitetura">3. Arquitetura do Processamento ⚙️</h2>
O fluxo de tratamento segue estas etapas técnicas:

<h3>Mapeamento de Pastas:</h3> Utiliza os.listdir e glob para localizar arquivos .xlsx em subpastas regionais.


```python
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
```

<h3>Limpeza de Cabeçalho (Skip Rows):</h3> O script pula as primeiras 6 linhas e realiza uma transposição (.T), transformando o que eram rótulos de linha em colunas.

<h3>Padronização Numérica:</h3> * Remove símbolos monetários e caracteres especiais via Regex; Converter o padrão brasileiro (vírgula) para o padrão computacional (ponto); Trata valores ausentes (NaN) como 0 para evitar erros de cálculo.

```python
def converteNumero(df:pd.DataFrame,colunas=[]) -> pd.DataFrame:
    for coluna in colunas:
        df[coluna] = (df[coluna].replace("NaN", 0).astype(str).str.replace(",", ".", regex=False)
        .str.replace(r"[^\d.]", "", regex=True).replace("", np.nan)
        .astype(float).replace(np.nan, 0))
    
    return df
```

<h3>Agregação:</h3> Consolida os dados utilizando .groupby("Equipe").sum(), garantindo que cada equipe tenha apenas uma linha de resumo por filial.

```python
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
```

<h2 id="requisitos">4. Requisitos 🛠</h2>
<ul>
    <li>Python 3.8+</li>
    <li>Pandas: Para manipulação de DataFrames.</li>
    <li>Openpyxl: Engine necessária para leitura de arquivos Excel modernos.</li>
    <li>Numpy: Para tratamento de valores nulos e operações vetoriais.</li>
</ul>

Para as instalações necessárias abra o terminal e digite:

```bash
pip install pandas openpyxl numpy
```
<h2 id="executar">5. Como Executar 🚀</h2>
Certifique-se de que a pasta Automação Relatório Operacional/operacao_logistica está no mesmo diretório que o script.

Coloque seus arquivos .xlsx dentro das subpastas de regionais.

Execute o comando:

```bash
pip python nome_do_seu_arquivo.py
```

O arquivo arquivo_final.xlsx será gerado na raiz do projeto.

<h2 id="tratamento">6. Tratamento de Erros ⚠️</h2>
O script possui blocos try-except robustos para garantir que, caso uma aba específica ou um arquivo esteja corrompido ou fora do padrão, o processamento não seja interrompido. O erro será logado no console informando o local exato do problema para correção manual posterior.