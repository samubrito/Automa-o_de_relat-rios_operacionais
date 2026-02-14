<h1>Automação de Consolidação: Relatório Operacional Logístico</h1>
<p align="center">
  <img src="automacao_relatorios.png" alt="Automação de Consolidação: Relatório Operacional Logístico" width="400px">
</p>
Este projeto consiste em um pipeline de ETL (Extração, Transformação e Carga) desenvolvido em Python para consolidar múltiplos relatórios de operação logística. O script é ideal para cenários onde os dados estão distribuídos em diversas pastas por região, arquivos por estado e abas por filial.

<details open="open">
<summary>📋 Sumário</summary>

- [Visão Geral](#visao_geral)
- [Estrutura de Dados](#estrutura)
- [Arquitetura do Processamento](#arquitetura)
- [Requisitos](#requisitos)
- [Como Executar](#executar)
- [Tratamento de Erros](#tratamento)

</details>

<h2 id="visao_geral">🔍 Visão Geral</h2>
<p>A automação resolve o problema de relatórios manuais que possuem cabeçalhos complexos e formatação inconsistente. O código varre diretórios, entra em cada aba de cada arquivo Excel, limpa os dados numéricos e gera um arquivo mestre consolidado para análise em BI ou Dashboards.</p>

<h2 id="estrutura">📊 Estrutura de Dados</h2>
O script espera uma hierarquia específica para atribuir os metadados corretamente:
<ul>
    <li>Estado: Extraído da última parte do nome do arquivo (ex: Relatorio_SP.xlsx → "SP").</li>
    <li>Mês: Extraído da célula B1 (índice 0,1) de cada aba (Formato esperado: Texto, ex: "Janeiro").</li>
    <li>Filial: Extraído automaticamente do nome da aba (sheet name).</li>
    <li>Colunas Métricas: Tempo_h, Km e Custo.</li>
</ul>

<h2 id="arquitetura">⚙️ Arquitetura do Processamento</h2>
O fluxo de tratamento segue estas etapas técnicas:

<h3>Mapeamento de Pastas:</h3> Utiliza os.listdir e glob para localizar arquivos .xlsx em subpastas regionais.

<h3>Limpeza de Cabeçalho (Skip Rows):</h3> O script pula as primeiras 6 linhas e realiza uma transposição (.T), transformando o que eram rótulos de linha em colunas.

<h3>Padronização Numérica:</h3> * Remove símbolos monetários e caracteres especiais via Regex; Converter o padrão brasileiro (vírgula) para o padrão computacional (ponto); Trata valores ausentes (NaN) como 0 para evitar erros de cálculo.

<h3>Agregação: Consolida os dados utilizando .groupby("Equipe").sum(), garantindo que cada equipe tenha apenas uma linha de resumo por filial.</h3>

<h2 id="requisitos">🛠 Requisitos</h2>
Python 3.8+

Pandas: Para manipulação de DataFrames.

Openpyxl: Engine necessária para leitura de arquivos Excel modernos.

Numpy: Para tratamento de valores nulos e operações vetoriais.

Bash
pip install pandas openpyxl numpy
<h2 id="executar">🚀 Como Executar</h2>
Certifique-se de que a pasta Automação Relatório Operacional/operacao_logistica está no mesmo diretório que o script.

Coloque seus arquivos .xlsx dentro das subpastas de regionais.

Execute o comando:

Bash
python nome_do_seu_arquivo.py
O arquivo arquivo_final.xlsx será gerado na raiz do projeto.

<h2 id="tratamento">⚠️ Tratamento de Erros</h2>
O script possui blocos try-except robustos para garantir que, caso uma aba específica ou um arquivo esteja corrompido ou fora do padrão, o processamento não seja interrompido. O erro será logado no console informando o local exato do problema para correção manual posterior.