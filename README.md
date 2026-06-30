# Processador de Planilhas - Líderes Cariocas

Este é um utilitário desenvolvido em Python para automação e conciliação de dados cadastrais dos Líderes Cariocas. O sistema compara a planilha mensal de registros com a planilha extraída do sistema Ergon, identifica inconsistências, elimina registros de ex-líderes, detecta duplicidades e permite a atualização de um histórico de divergências.

A aplicação conta com uma interface gráfica amigável desenvolvida em `Tkinter` e tabelas interativas utilizando `pandastable`.

---

## ⚡ Funcionalidades

O programa executa as seguintes operações principais:

1. **Limpeza de Ex-Líderes:** 
   - Identifica e remove os registros da planilha mensal cujos nomes não constam na planilha do Ergon (indicando desligamento ou ausência do programa).
   - Gera automaticamente um novo arquivo atualizado chamado `Planilha Mensal - eliminados registros de ex líderes.xlsx`.

2. **Divergência de Dados (Valores Diferentes):**
   - Compara o **Órgão/Entidade** (usando mapeamento inteligente de siglas), o **Setor** e, caso a coluna **REFERENCIA** esteja presente em ambas as planilhas, a **Referência** dos cargos.
   - Apresenta em uma tabela interativa todos os registros ativos que possuem dados divergentes entre as duas fontes.
   - Permite exportar e acrescentar essas divergências diretamente em uma planilha de **Histórico** selecionada pelo usuário.

3. **Nomes Duplicados:**
   - Detecta registros que possuem o mesmo nome repetido na planilha mensal e exibe em formato de tabela para análise manual (detalhando início de lotação, setor e órgão).

4. **Padronização e Normalização:**
   - Normaliza os textos removendo acentos, caracteres especiais e padronizando caixas de texto (maiúsculo) para evitar falsos positivos na comparação.
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

1. Abra o terminal na pasta do projeto.
2. Execute o script da interface:
   ```bash
   python Interface.py
   ```
3. Na janela do programa:
   * Clique em **Subir Planilha Mensal** e selecione o arquivo correspondente.
   * Clique em **Subir Planilha Ergon** e selecione o respectivo relatório do Ergon.
   * Após o carregamento de ambas as planilhas, os botões **Valores Diferentes** e **Nomes Duplicados** ficarão ativos.
   * Ao clicar em **Valores Diferentes**, o programa fará o cruzamento de dados, salvará a versão limpa na pasta raiz e abrirá a tabela com as divergências. Ao fechar esta tabela de divergências, o sistema perguntará se você deseja salvar estes registros na planilha **Histórico**.
   * Use o botão de recarga (`↻` no canto inferior direito) para redefinir as planilhas e realizar um novo processamento.
