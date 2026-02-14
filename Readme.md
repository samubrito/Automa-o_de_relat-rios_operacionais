📊 Automação de Consolidação – Operação Logística

Pipeline em Python para consolidação automática de relatórios operacionais de logística, realizando:

Leitura de múltiplas pastas

Processamento de múltiplos arquivos Excel

Tratamento e padronização de dados

Conversão de tipos numéricos

Consolidação final em arquivo único

🎯 Objetivo do Projeto

Automatizar a consolidação de relatórios operacionais distribuídos por:

📂 Estado

📂 Mês

📄 Arquivo Excel

📑 Múltiplas planilhas internas

Eliminando trabalho manual de consolidação e reduzindo risco de erro humano.

🏗️ Estrutura Esperada de Pastas
Automação Relatório Operacional/
└── operacao_logistica/
    ├── janeiro/
    │   ├── Operacao_Logistica_SP.xlsx
    │   └── Operacao_Logistica_MG.xlsx
    ├── fevereiro/
    │   ├── Operacao_Logistica_SP.xlsx
    │   └── Operacao_Logistica_RJ.xlsx


O script percorre automaticamente:

Todas as pastas

Todos os arquivos .xlsx

Todas as planilhas dentro dos arquivos

🔄 Fluxo do Pipeline

🔎 Percorre as pastas de meses

📄 Lê todos os arquivos Excel

📑 Processa cada planilha individualmente

🧹 Realiza:

Remoção de linhas irrelevantes

Transposição de dados

Padronização de colunas

Conversão de valores numéricos

📊 Agrupa dados por equipe

➕ Adiciona colunas:

Estado

Mês

Filial

🧩 Consolida tudo em um único DataFrame

💾 Exporta arquivo_final.xlsx

🧠 Principais Técnicas Utilizadas

pandas para manipulação de dados

glob para busca dinâmica de arquivos

os.path para portabilidade de caminhos

Tratamento de exceções com try/except

Modularização com funções

Estrutura profissional com main()

🚀 Como Executar
1️⃣ Instalar dependências
pip install pandas numpy openpyxl

2️⃣ Executar o script
python script.py


O arquivo final será gerado como:

arquivo_final.xlsx

🧩 Funções do Projeto
Função	Responsabilidade
converteNumero()	Limpeza e conversão de colunas numéricas
trataPlanilha()	Tratamento individual de cada planilha
trataArquivos()	Consolidação das planilhas de um arquivo
trataPastas()	Consolidação geral de todos os arquivos
main()	Execução principal do pipeline
🛡️ Tratamento de Erros

O pipeline possui tratamento de exceções em dois níveis:

🔹 Erro por planilha

🔹 Erro por arquivo

Isso garante que um único erro não interrompa toda a consolidação.

📈 Possíveis Melhorias Futuras

Implementação de logging estruturado

Parametrização via CLI (argparse)

Criação de testes unitários

Dockerização do projeto

Integração com banco de dados (PostgreSQL)

Orquestração via Airflow

💼 Aplicabilidade Profissional

Este projeto simula um cenário real de:

Automação de relatórios

Consolidação de dados operacionais

Preparação de base para BI

Pipeline inicial de engenharia de dados

Pode ser facilmente integrado a:

Power BI

Banco de dados

Pipeline ETL

Sistema de monitoramento

👨‍💻 Autor

Samuel Brito
Engenharia de Controle e Automação
Foco em Dados, Automação e Engenharia de Dados

⭐ Considerações

Este projeto demonstra:

Organização

Estrutura modular

Tratamento de dados reais

Resiliência a falhas

Mentalidade de produção