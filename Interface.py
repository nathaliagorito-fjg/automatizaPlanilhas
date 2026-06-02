import os, sys
import pandas as pd
import Planilhas
from openpyxl import load_workbook
from pandastable import Table
from tkinter import *
from tkinter.filedialog import askopenfilename
from tkinter import messagebox

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
            buttonProcessa.config(state='normal')

    def testaPlanilhasHistoricoMinibioCarregadas():
        if Planilhas.planilhaErgon is not None and Planilhas.planilhaHistorico is not None:
            buttonProcessaHist.config(state='normal')

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
    
    buttonMensal.config(state='disabled')
    buttonMinibio.config(state='disabled')

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

    colunas_exibicao = ['NOME', 'CPF', 'CARGO', 'FUNCAO', 'NOME_SETOR', 'SIGLA_ORGAO_ENTIDADE']

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

    mostraPlanilha(duplicados[colunas_exibicao], 'Valores Duplicados por CPF', fechaJanelaDadosDuplicados=confirmarSalvamento)

def resetaTudo():
    buttonMensal.config(state='normal')
    buttonMinibio.config(state='normal')
    buttonProcessa.config(state='disabled')
    buttonErgon.config(state='normal')
    buttonHist.config(state='normal')
    buttonProcessaHist.config(state='disabled')
    labelMensagem['text'] = ''

# def abrir_modal_info():
#     modal = Toplevel(janela)
#     modal.title('Informações do Sistema')
#     modal.geometry('520x320')
#     modal.resizable(False, False)
#     modal.grab_set()  # Bloqueia a janela principal até fechar o modal
#     modal.config(bg='white')

#     # Alinha as opções de estilo do modal com a identidade do app
#     Label(modal, text='Instruções de Processamento', font=('Arial', 11, 'bold'), bg='white').pack(pady=15)
    
#     lbl_texto = Label(modal, text=infos, justify='left', bg='white', font=('Arial', 9), anchor='w')
#     lbl_texto.pack(padx=25, fill='both', expand=True)
    
#     Button(modal, text='Fechar', command=modal.destroy, width=15).pack(pady=15)

janela = Tk()
janela.iconbitmap(armazenaImagem('iconeInterface.ico'))
janela.geometry('1000x600')
janela.resizable(False, False)
janela.title('Processador de Planilhas Lideres Cariocas')
janela.option_add('*Acivebackground', 'black')
janela.option_add('*Activeforeground', 'white')
janela.option_add('*Background', 'white')
janela.option_add('*Foreground', 'black')
janela.option_add('*Bd', 1)
janela.option_add('*Font', ('Arial', 8))
janela.option_add('*Relief', 'solid')
janela.option_add('*Width', 20)

labelTitulo = Label(janela, text='Processador de Planilhas\nLideres Cariocas')
labelTitulo.config(bg=labelTitulo.master.cget('bg'), bd=0, font=10, relief='flat', width=30)
labelTitulo.pack(pady=10)

infos = """
    1. Elimina registros da planilha mensal que não estejam na planilha minibio;
    2. Salva uma nova planilha sem esses registros eliminados;
    3. Exibe registros de ambas planilhas, mensal e minibio, que estejam com valores diferentes para a coluna ORGAO_ENTIDADE.

    1. Exibe registros dos CPFs duplicados da planilha do Ergon que estejam com valores diferentes para as colunas CARGO, FUNCAO, NOME_SETOR, SIGLA_ORGAO_ENTIDADE;
    2. Salva os registros duplicados na planilha do histórico da minibio.
"""

infosMinibio = """
    1. Elimina registros da planilha mensal que não estejam na planilha minibio;
    2. Salva uma nova planilha sem esses registros eliminados;
    3. Exibe registros de ambas planilhas, mensal e minibio, que estejam com valores diferentes para a coluna ORGAO_ENTIDADE.
"""

infosMensal = """
    1. Exibe registros dos CPFs duplicados da planilha do Ergon que estejam com valores diferentes para as colunas CARGO, FUNCAO, NOME_SETOR, SIGLA_ORGAO_ENTIDADE;
    2. Salva os registros duplicados na planilha do histórico da minibio.
"""

# buttonAjuda = Button(janela, text='Como Funciona?', command=abrir_modal_info, width=15)
# buttonAjuda.pack(pady=5)

# infos = Label(janela, text=infos)
# infos.config(bg=infos.master.cget('bg'), bd=0, justify='left', relief='flat', width=0)
# infos.pack(pady=5)

frameSecoes = Frame(janela, bg='white')
frameSecoes.pack(pady=10)

secaoLideres = LabelFrame(frameSecoes, text="Líderes Mensais", bg='white', padx=5, pady=10, labelanchor='n', width=350, height=300)
# secaoLideres.pack(padx=10, side="left")
secaoLideres.pack_propagate(False)

infosMinibio = Label(secaoLideres, text=infosMinibio)
infosMinibio.config(bg=infosMinibio.master.cget('bg'), bd=0, justify='left', relief='flat', width=0)
infosMinibio.pack(pady=5)

buttonMensal = Button(secaoLideres, text='Carregar Mensal', command=lambda:carregaPlanilhas('mensal'))
buttonMensal.pack(pady=3)

buttonMinibio = Button(secaoLideres, text='Carregar Minibio', command=lambda:carregaPlanilhas('minibio'))
buttonMinibio.pack(pady=3)

buttonProcessa = Button(secaoLideres, text='Processar Planilhas', command=lambda:processaPlanilhasLideresMensais())
buttonProcessa.config(state='disabled')
buttonProcessa.pack(pady=3)

secaoHistorico = LabelFrame(frameSecoes, text="Histórico Minibio", bg='white', padx=5, pady=10, labelanchor='n', width=350, height=300)
# secaoHistorico.pack(padx=10, side="right")
secaoHistorico.pack_propagate(False)

infosMensal = Label(secaoHistorico, text=infosMensal)
infosMensal.config(bg=infosMensal.master.cget('bg'), bd=0, justify='left', relief='flat', width=0)
infosMensal.pack(pady=5)

buttonErgon = Button(secaoHistorico, text='Carregar Ergon', command=lambda:carregaPlanilhas('ergon'))
buttonErgon.pack(pady=3)

buttonHist = Button(secaoHistorico, text='Carregar Histórico', command=lambda:carregaPlanilhas('historico'))
buttonHist.pack(pady=3)

buttonProcessaHist = Button(secaoHistorico, text='Processar Planilha', command=lambda:processaPlanilhasHistoricoMinibio())
buttonProcessaHist.config(state='disabled')
buttonProcessaHist.pack(pady=3)

labelMensagem = Label(janela, text='')
labelMensagem.config(bg=labelMensagem.master.cget('bg'), bd=0, relief='flat', width=40)
labelMensagem.pack(pady=15)

diretorioImagem = armazenaImagem('iconeRefresh.png')
imagemButtonRefresh = PhotoImage(file=diretorioImagem)
buttonRefresh = Button(janela, image=imagemButtonRefresh, command=lambda:resetaTudo())
buttonRefresh.config(bg=buttonRefresh.master.cget('bg'), bd=0, relief='flat', width=30)
buttonRefresh.place(relx=1.0, rely=1.0, anchor='se', x=-15, y=-15)

janela.mainloop()