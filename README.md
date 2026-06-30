# Processador de Planilhas - Líderes Cariocas

Este é um software desenvolvido em Python para automação e conciliação de dados cadastrais dos Líderes Cariocas. O sistema compara a planilha minibio de registros com a planilha extraída do sistema Ergon, identifica inconsistências, detecta duplicidades, localiza registros ausentes e permite a atualização de um histórico de divergências.

A aplicação conta com uma interface gráfica amigável desenvolvida em `Tkinter` e tabelas interativas utilizando `pandastable`.

---

## ⚡ Funcionalidades

O programa executa as seguintes operações principais:

1. **Divergência de Dados (Valores Diferentes):**
   - Cruza os dados das planilhas utilizando o **CPF** como chave.
   - Compara **ORGAO_ENTIDADE**, **NOME_SETOR** e **REFERENCIA** em ambas as planilhas.
   - Apresenta em uma tabela todos os registros ativos que possuem dados divergentes entre as duas fontes.
   - Permite exportar e acrescentar essas divergências diretamente em uma planilha de **Histórico** selecionada pelo usuário.

2. **Nomes Duplicados:**
   - Detecta registros que possuem o mesmo **CPF** repetido na planilha Ergon e exibe em formato de tabela para análise manual.

3. **Ausentes no Ergon:**
   - Identifica os registros da planilha Minibio cujos CPFs não constam na planilha do Ergon (ex-líderes ou ausentes do programa) e os exibe em formato de janela para análise manual.

4. **Validação Automática de Colunas:**
   - Valida, no momento do carregamento das planilhas, se todas as colunas necessárias estão presentes.
   - Caso falte alguma coluna, exibe um alerta de erro na tela especificando em qual planilha é e qual coluna está ausente.

5. **Padronização e Normalização:**
   - Normaliza os textos removendo acentos, caracteres especiais e padronizando caixas de texto em maiúsculo para evitar falsos positivos na comparação.
   - Normaliza os CPFs para conter apenas números e preenche com zeros à esquerda até 11 dígitos, eliminando discrepâncias de formatação.
   - Mapeia automaticamente os nomes completos dos órgãos (ex: "Gabinete do Prefeito") para suas respectivas siglas (ex: "GBP").

---

## 📂 Estrutura do Projeto

* **`Interface.py`**: Contém o código da interface gráfica, gerenciamento dos estados da tela, carregamento de arquivos e janelas de visualização de dados.
* **`Planilhas.py`**: Contém as funções lógicas de processamento dos dados via `pandas`, incluindo regras de normalização de strings, mapeamento de siglas para órgãos/entidades e junção (*merge*) das bases de dados.
* **`iconeInterface.ico`**: Arquivo de ícone utilizado na janela do programa.

---

## 🛠️ Pré-requisitos e Instalação

### Requisitos do Sistema
* Python 3.8 ou superior instalado.

### Instalação de Dependências
Para executar o projeto, você precisará instalar as bibliotecas de manipulação de planilhas e visualização de dados. Execute o seguinte comando no seu terminal/prompt de comando:

```bash
pip install pandas openpyxl pandastable
```

> [!NOTE]
> A biblioteca gráfica `tkinter` já vem integrada por padrão na instalação padrão do Python no Windows.

---

## 🚀 Como Executar o Programa

### Criando o executável
1. Abra o terminal na pasta do projeto.
2. Execute "python -m PyInstaller --add-data "iconeRefresh.png;." —add-data “iconeInterface.ico;.” --icon "iconeInterface.ico" --name "Processador de Planilhas Lideres Cariocas" --onefile --windowed Interface.py".