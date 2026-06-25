import os, sys
import pandas as pd
import Planilhas
from openpyxl import load_workbook
from pandastable import Table
from tkinter import *
from tkinter import messagebox
from tkinter.filedialog import askopenfilename

def armazenaImagem(diretorioAtual):
    try:
        diretorioTemporario = sys._MEIPASS
    except Exception:
        diretorioTemporario = os.path.abspath('.')
    
    return os.path.join(diretorioTemporario, diretorioAtual)

def carregaPlanilhas(tipo):
    diretorio = askopenfilename(filetypes=[('Excel files', '*.xlsx *.xls')])

    if not diretorio:
        return
    
    def testaPlanilhasLideresMensaisCarregadas():
        if Planilhas.planilhaMensal is not None and Planilhas.planilhaMinibio is not None:
            buttonProcessaLideres.config(state='normal')

    def testaPlanilhasHistoricoMinibioCarregadas():
        if Planilhas.planilhaErgon is not None and Planilhas.planilhaHistorico is not None:
            buttonProcessaHistorico.config(state='normal')

    def validaTipoPlanilha(tipo):
        tipos = {'mensal': ['mensal'], 'minibio': ['minibio'], 'historico': ['historico', 'histórico'], 'ergon': ['ergon']}

        nomeArquivo = diretorio.lower()

        return any(palavra in nomeArquivo for palavra in tipos.get(tipo, []))

    if not validaTipoPlanilha(tipo):
        labelMensagem['text'] = 'Planilha errada.'
        labelMensagem.after(5000, lambda: labelMensagem.config(text=''))

        return

    if tipo == 'mensal':
        Planilhas.planilhaMensal = pd.read_excel(diretorio)

        global planilhaMensalAtiva, planilhaMensalFormatada

        planilhaMensalFormatada = load_workbook(diretorio)
        planilhaMensalAtiva = planilhaMensalFormatada.active

        labelMensagem['text'] = 'Planilha MENSAL DOS LÍDERES carregada.'

        testaPlanilhasLideresMensaisCarregadas()
    elif tipo == 'minibio':
            Planilhas.planilhaMinibio = pd.read_excel(diretorio)

            labelMensagem['text'] = 'Planilha MINIBIO carregada.'

            testaPlanilhasLideresMensaisCarregadas()
    elif tipo == 'ergon':
        Planilhas.planilhaErgon = pd.read_excel(diretorio)

        labelMensagem['text'] = 'Planilha ERGON carregada.'

        testaPlanilhasHistoricoMinibioCarregadas()
    elif tipo == 'historico':
        Planilhas.planilhaHistorico = pd.read_excel(diretorio)
        Planilhas.diretorioHistorico = diretorio
            
        labelMensagem['text'] = 'Planilha HISTÓRICO DA MINIBIO carregada.'
            
        testaPlanilhasHistoricoMinibioCarregadas()

def mostraPlanilha(planilha, titulo, fechaJanelaDadosDuplicados=None):
    if planilha is None or planilha.empty:
        labelMensagem['text'] = 'Não há dados para exibir.'
    else:
        novaJanela = Toplevel(janela)
        novaJanela.title(titulo)
        novaJanela.geometry('1100x500')
        novaJanela.resizable(False, False)

        framePlanilha = Frame(novaJanela)
        framePlanilha.pack(fill=BOTH, expand=1)

        tabela = Table(framePlanilha, dataframe=planilha)
        tabela.show()

        if fechaJanelaDadosDuplicados:
            novaJanela.protocol("WM_DELETE_WINDOW", lambda:fechaJanelaDadosDuplicados(novaJanela))

def processaPlanilhasLideresMensais():
    nomesDuplicados, planilhasMescladas, planilhaAlterada = Planilhas.processaPlanilhasLideresMensais()

    if nomesDuplicados is None:
        labelMensagem['text'] = 'Carregue as planilhas primeiro!'
        return
    
    buttonPlanilhaMensal.config(state='disabled')
    buttonPlanilhaMinibio.config(state='disabled')

    mostraPlanilha(nomesDuplicados[['NOME','INICIO_LOTACAO','NOMESETOR','ORGAO_ENTIDADE']], 'Nomes Duplicados')
    mostraPlanilha(planilhasMescladas.loc[planilhasMescladas['IGUAIS'] == False, ['NOME', 'ORGAO_ENTIDADE_MINIBIO', 'ORGAO_ENTIDADE_MENSAL', 'SIGLA', 'IGUAIS']], 'Valores Different_col')

    for linha in planilhaMensalAtiva.iter_rows():
        for coluna in linha:
            coluna.value = None
    
    for indiceColuna, coluna in enumerate(planilhaAlterada.columns, start=1):
        planilhaMensalAtiva.cell(row=1, column=indiceColuna, value=coluna)

    for indiceLinha, linha in planilhaAlterada.iterrows():
        for indiceColuna, coluna in enumerate(linha, start=1):
            planilhaMensalAtiva.cell(row=indiceLinha + 2, column=indiceColuna, value=coluna)

    for i in range(planilhaMensalAtiva.max_row, 1, -1):
        if planilhaMensalAtiva.cell(row=i, column=1).value is None:
            planilhaMensalAtiva.delete_rows(i)
        
    planilhaMensalFormatada.save('Planilha Mensal - eliminados registros de ex líderes.xlsx')

def processaPlanilhasHistoricoMinibio():
    duplicados = Planilhas.processaPlanilhasHistoricoMinibio()

    if duplicados is None:
        labelMensagem['text'] = 'Carregue as planilhas do Histórico primeiro!'
        return

    colunasExibicao = ['NOME', 'CPF', 'CARGO', 'FUNCAO', 'NOME_SETOR', 'SIGLA_ORGAO_ENTIDADE']

    def confirmarSalvamento(janelaPopup):
        resposta = messagebox.askyesno("Salvar Alterações", "Deseja salvar esses valores duplicados no histórico da minibio?")
        
        if resposta:  
            wb = load_workbook(Planilhas.diretorioHistorico)
            ws = wb.active

            cabecalho = [cell.value for cell in ws[1]]

            for _, linha in duplicados.iterrows():
                novaLinha = []
                for coluna in cabecalho:
                    novaLinha.append(linha.get(coluna, None))
                ws.append(novaLinha)

            wb.save(Planilhas.diretorioHistorico)
            labelMensagem['text'] = 'Histórico atualizado com sucesso!'
        else:
            labelMensagem['text'] = 'Ação cancelada. Dados não foram salvos.'
        
        janelaPopup.destroy()  

    mostraPlanilha(duplicados[colunasExibicao], 'Valores Duplicados por CPF', fechaJanelaDadosDuplicados=confirmarSalvamento)

def resetaTudo():
    buttonPlanilhaMensal.config(state='normal')
    buttonPlanilhaMinibio.config(state='normal')
    buttonProcessaLideres.config(state='disabled')

    buttonPlanilhaErgon.config(state='normal')
    buttonPlanilhaHistorico.config(state='normal')
    buttonProcessaHistorico.config(state='disabled')

    labelMensagem['text'] = ''

#Interface
janela = Tk()
janela.iconbitmap(armazenaImagem('iconeInterface.ico'))
janela.geometry('1100x650')
janela.resizable(False, False)
janela.title('Processador de Planilhas Lideres Cariocas')

corFundo = '#ECF1F4'
corAzul = '#0085B3'
corAzulEscuro = '#094A75'
corCard = '#FFFFFF'
corTexto = '#1D1D1B'
corBorda = '#D6E6F1'

janela.configure(bg=corFundo)

estiloBotao = {'bg': corAzul, 'fg': 'white', 'activebackground': corAzulEscuro, 'activeforeground': 'white', 'font': ('Segoe UI', 10, 'bold'), 'relief': 'flat', 'bd': 0, 'width': 28, 'cursor': 'hand2', 'pady': 6}

labelTitulo = Label(janela, text='Processador de Planilhas\nLideres Cariocas', bg=corFundo, fg=corAzulEscuro, font=('Segoe UI', 26, 'bold'))
labelTitulo.pack(pady=(25, 20))

infosLideresMensais = """Esta seção faz:

1. Elimina registros da planilha mensal que não estejam na planilha minibio;
2. Salva uma nova planilha sem esses registros eliminados;
3. Exibe registros das duas planilhas que estejam com valores diferentes para ORGAO_ENTIDADE."""

infosHistoricoMinibio = """Esta seção faz:

1. Exibe registros dos CPFs duplicados da planilha do Ergon que possuam diferenças em CARGO, FUNCAO, NOME_SETOR e SIGLA_ORGAO_ENTIDADE;
2. Salva os registros duplicados na planilha do histórico da minibio."""

frameSecoes = Frame(janela, bg=corFundo)
frameSecoes.pack(pady=10)

#Seção Líderes Mensais
secaoLideresMensais = Frame(frameSecoes, bg=corCard, width=460, height=400, highlightbackground=corBorda, highlightthickness=1)
secaoLideresMensais.pack(side='left', padx=15)
secaoLideresMensais.pack_propagate(False)

frameTituloLideres = Frame(secaoLideresMensais, bg=corCard)
frameTituloLideres.pack(fill='x', padx=20, pady=(20, 15))

Label(frameTituloLideres, text='👥', bg=corCard, fg=corAzul, font=('Segoe UI Emoji', 22)).pack(side='left')
Label(frameTituloLideres, text='Líderes Mensais', bg=corCard, fg=corAzulEscuro, font=('Segoe UI', 16, 'bold')).pack(side='left', padx=10)
Frame(frameTituloLideres, bg=corAzul, height=2).pack(side='left', fill='x', expand=True, padx=(15, 0), pady=15)
Label(secaoLideresMensais, text=infosLideresMensais, bg=corCard, fg=corTexto, justify='left', wraplength=400, font=('Segoe UI', 10)).pack(anchor='w', padx=25, pady=(0, 15))

buttonPlanilhaMensal = Button(secaoLideresMensais, text='📄 Planilha Mensal', command=lambda: carregaPlanilhas('mensal'), **estiloBotao)
buttonPlanilhaMensal.pack(pady=5)

buttonPlanilhaMinibio = Button(secaoLideresMensais, text='👥 Planilha Minibio', command=lambda: carregaPlanilhas('minibio'), **estiloBotao)
buttonPlanilhaMinibio.pack(pady=5)

buttonProcessaLideres = Button(secaoLideresMensais, text='⚙️ Processar Planilhas', command=lambda: processaPlanilhasLideresMensais(), **estiloBotao)
buttonProcessaLideres.config(state='disabled')
buttonProcessaLideres.pack(pady=5)

#Seção Histórico Minibio
secaoHistoricoMinibio = Frame(frameSecoes, bg=corCard, width=460, height=400, highlightbackground=corBorda, highlightthickness=1)
secaoHistoricoMinibio.pack(side='left', padx=15)
secaoHistoricoMinibio.pack_propagate(False)

frameTituloHistorico = Frame(secaoHistoricoMinibio, bg=corCard)
frameTituloHistorico.pack(fill='x', padx=20, pady=(20, 15))

Label(frameTituloHistorico, text='🕘', bg=corCard, fg=corAzul, font=('Segoe UI Emoji', 22)).pack(side='left')
Label(frameTituloHistorico, text='Histórico Minibio', bg=corCard, fg=corAzulEscuro, font=('Segoe UI', 16, 'bold')).pack(side='left', padx=10)
Frame(frameTituloHistorico, bg=corAzul, height=2).pack(side='left', fill='x', expand=True, padx=(15, 0), pady=15)
Label(secaoHistoricoMinibio, text=infosHistoricoMinibio, bg=corCard, fg=corTexto, justify='left', wraplength=400, font=('Segoe UI', 10)).pack(anchor='w', padx=25, pady=(0, 15))

buttonPlanilhaErgon = Button(secaoHistoricoMinibio, text='💼 Planilha Ergon', command=lambda: carregaPlanilhas('ergon'), **estiloBotao)
buttonPlanilhaErgon.pack(pady=5)
buttonPlanilhaHistorico = Button(secaoHistoricoMinibio, text='🕘 Planilha Histórico', command=lambda: carregaPlanilhas('historico'), **estiloBotao)
buttonPlanilhaHistorico.pack(pady=5)

buttonProcessaHistorico = Button(secaoHistoricoMinibio, text='⚙️ Processar Planilha', command=lambda: processaPlanilhasHistoricoMinibio(), **estiloBotao)
buttonProcessaHistorico.config(state='disabled')
buttonProcessaHistorico.pack(pady=5)


labelMensagem = Label(janela, bg=corFundo, fg=corAzulEscuro, font=('Segoe UI', 10, 'bold'))
labelMensagem.pack(pady=20)

buttonRefresh = Button(janela, text='↻', command=lambda: resetaTudo(), fg=corAzul, activeforeground=corAzulEscuro, font=('Segoe UI', 25, 'bold'), relief='flat', bd=0, width=3, height=1, cursor='hand2')
buttonRefresh.place(relx=1.0, rely=1.0, x=-25, y=-20, anchor='se')

janela.mainloop()