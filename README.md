# AgenteAI - Automação TMS Datafrete (FOTUS)

Este repositório contém a aplicação **AgenteAI**, um agente de automação voltado para a extração, formatação e geração de relatórios do sistema TMS Datafrete da FOTUS.

## 🗺️ Mapeamento da Aplicação

### Estrutura do Projeto
A aplicação foi construída em Python, focando em automação web, manipulação de dados e uma interface gráfica amigável. Abaixo o mapeamento dos arquivos e diretórios:

- **`Datafrete_agent.py`**: Arquivo principal contendo toda a lógica da aplicação.
  - **Interface Gráfica (GUI)**: Utiliza a biblioteca `customtkinter` para fornecer uma interface ao usuário, solicitando credenciais, tipo de relatório (Arquivos ou Entrega) e modo de operação (Geral ou Detalhado). Possui exibição de logs em tempo real.
  - **Automação Web**: Utiliza o `playwright` (modo assíncrono) para logar no portal do Datafrete, navegar até as pendências, extrair a visão de "Organização + Transportador" e realizar o download dos arquivos em Excel. Suporta extração Macro (Resumo) ou Analítica (Drill-down).
  - **Processamento de Dados**: Utiliza bibliotecas como `pandas` e o motor `calamine` para leitura robusta de arquivos Excel extraídos (que frequentemente têm formatos não padronizados).
  - **Geração de Relatórios**: Formata os dados obtidos e cria novas planilhas estilizadas via `openpyxl`, inserindo as logos da FOTUS e da Datafrete nos cabeçalhos.
- **`requirements.txt`**: Lista as dependências do projeto (Playwright, CustomTkinter, Pandas, Openpyxl, etc.).
- **`exports/`**: Diretório gerado automaticamente onde os arquivos originais baixados ("brutos") e os relatórios gerados ("formatados") são armazenados.
- **Recursos Visuais (`logo_datafrete_excel.png`, `logo_fotus_excel.png`)**: Imagens utilizadas para compor os banners e cabeçalhos nas planilhas do Excel.
- **`.env`**: Arquivo para variáveis de ambiente, utilizado para definir configurações estáticas (ex: `DATAFRETE_URL`).
- **`venv/`**: Ambiente virtual do Python isolando as dependências.

## ⚙️ Funcionalidades Principais

1. **Dois Modos de Extração**:
   - **Geral / Macro**: Traz apenas os totais por Unidade e Transportador.
   - **Detalhado / Analítico**: Realiza o *drill-down* dentro do sistema, baixando arquivo por arquivo para montar listagens detalhadas de CTs e NFs pendentes.
2. **Dois Tipos de Pendências**:
   - Pendências de Arquivos (NF sem CT / CT sem NF).
   - Pendências de Entrega (Pendente / Ocorrência).
3. **Resiliência e Cache**: A aplicação possui mecanismos de `retry` na leitura de arquivos via `calamine`, ideal para superar os travamentos eventuais do OneDrive. Também possui um modo "Cache", que permite reprocessar relatórios já baixados na pasta `exports` sem precisar acessar o Datafrete novamente.
4. **Exportações Múltiplas**: Em modo analítico, o sistema divide os resultados em um arquivo mestre contendo múltiplas abas, além de isolar os resultados por Unidade e por "aba" para distribuição granular.

## 🚀 Como Executar
1. Instale as dependências: `pip install -r requirements.txt`
2. Instale o navegador do Playwright: `playwright install chromium`
3. Execute o script principal: `python Datafrete_agent.py`
