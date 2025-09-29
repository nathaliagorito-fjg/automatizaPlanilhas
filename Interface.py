import os, sys
import pandas as pd
import Planilhas
from openpyxl import load_workbook
from pandastable import Table
from tkinter import *
from tkinter.filedialog import askopenfilename

def armazenaImagem(diretorioAtual):
    try:
        diretorioTemporario = sys._MEIPASS
    except Exception:
        diretorioTemporario = os.path.abspath('.')
    
    return os.path.join(diretorioTemporario, diretorioAtual)

def carregaPlanilhas(tipo):
    diretorio = askopenfilename(filetypes=[('Excel files', '*.xlsx *.xls')])

    def testaPlanilhasCarregadas():
        if Planilhas.planilhaMensal is not None and Planilhas.planilhaMinibio is not None:
            buttonProcessa.config(state='normal')

    if not diretorio:
        return
    
    if tipo == 'mensal':
        if 'mensal' not in diretorio.lower():
            labelMensagem['text'] = 'Planilha errada.'
            labelMensagem.after(5000, lambda:labelMensagem.config(text=''))
        else:
            Planilhas.planilhaMensal = pd.read_excel(diretorio)

            global planilhaMensalAtiva, planilhaMensalFormatada
            
            planilhaMensalFormatada = load_workbook(diretorio)
            planilhaMensalAtiva = planilhaMensalFormatada.active

            labelMensagem['text'] = 'Planilha MENSAL carregada.'

            testaPlanilhasCarregadas()
    elif tipo == 'minibio':
        if 'minibio' not in diretorio.lower():
            labelMensagem['text'] = 'Planilha errada.'
            labelMensagem.after(5000, lambda:labelMensagem.config(text=''))
        else:
            Planilhas.planilhaMinibio = pd.read_excel(diretorio)

            labelMensagem['text'] = 'Planilha MINIBIO carregada.'

            testaPlanilhasCarregadas()

def mostraPlanilha(planilha, titulo):
    if planilha is None or planilha.empty:
        labelMensagem['text'] = 'Não há planilhas carregadas.'
    else:
        novaJanela = Toplevel(janela)
        novaJanela.title(titulo)
        novaJanela.geometry('1100x500')
        novaJanela.resizable(False, False)

        framePlanilha = Frame(novaJanela)
        framePlanilha.pack(fill=BOTH, expand=1)

        tabela = Table(framePlanilha, dataframe=planilha)
        tabela.show()

        # def fechaJanela():
        #     planilha.loc[:] = tabela.model.df

        #     for indiceLinha, linha in Planilhas.planilhaMensal.iterrows():
        #         for indiceColuna, coluna in enumerate(linha, start=1):
        #             planilhaMensalAtiva.cell(row=indiceLinha + 2, column=indiceColuna, value=coluna)
        #             print('deu certo')

        #     planilhaMensalFormatada.save('planilha mudada teste.xlsx')

        #     novaJanela.destroy()

        # novaJanela.protocol('WM_DELETE_WINDOW', fechaJanela)

def processaPlanilhas():
    nomesDuplicados, planilhasMescladas, planilhaAlterada = Planilhas.processaPlanilhas()

    if nomesDuplicados is None:
        labelMensagem['text'] = 'Carregue as planilhas primeiro!'
        return
    
    buttonMensal.config(state='disabled')
    buttonMinibio.config(state='disabled')

    mostraPlanilha(nomesDuplicados[['NOME','INICIO_LOTACAO','NOMESETOR','ORGAO_ENTIDADE']], 'Nomes Duplicados')
    #mostraPlanilha(datasDuplicadas[['NOME', 'INICIO_LOTACAO', 'NOMESETOR', 'ORGAO_ENTIDADE']], 'Datas Duplicadas')
    mostraPlanilha(planilhasMescladas.loc[planilhasMescladas['IGUAIS'] == False, ['NOME', 'ORGAO_ENTIDADE_MINIBIO', 'ORGAO_ENTIDADE_MENSAL', 'SIGLA', 'IGUAIS']], 'Valores Diferentes')

    
    for row in planilhaMensalAtiva.iter_rows():
        for cell in row:
            cell.value = None
            
    for indiceLinha, linha in planilhaAlterada.iterrows():
        for indiceColuna, coluna in enumerate(linha, start=1):
            planilhaMensalAtiva.cell(row=indiceLinha + 2, column=indiceColuna, value=coluna)
            
    planilhaMensalFormatada.save('planilha mudada teste.xlsx')

def resetaTudo():
    buttonMensal.config(state='normal')
    buttonMinibio.config(state='normal')
    buttonProcessa.config(state='disabled')
    labelMensagem['text'] = ''

    #janela.destroy()

janela = Tk()
janela.iconbitmap(armazenaImagem('iconeInterface.ico'))
janela.geometry('550x400')
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
labelTitulo.pack(pady=15)

infos = """
    Este programa realiza:
    1. Elimina nomes da planilha mensal que não estejam na planilha minibio e salva em uma nova
    2. Exibe registros de ambas planilhas que estejam com valores diferentes para a coluna ORGAO_ENTIDADE
    É necessário inserir as duas planilhas para que o processamento ocorra
"""

labelInfo = Label(janela, text=infos)
labelInfo.config(bg=labelInfo.master.cget('bg'), bd=0, justify='left', relief='flat', width=0)
labelInfo.pack(pady=15)

buttonMensal = Button(janela, text='Carregar Mensal', command=lambda:carregaPlanilhas('mensal'))
buttonMensal.pack(pady=5)

buttonMinibio = Button(janela, text='Carregar Minibio', command=lambda:carregaPlanilhas('minibio'))
buttonMinibio.pack(pady=5)

buttonProcessa = Button(janela, text='Processar Planilhas', command=lambda:processaPlanilhas() if (Planilhas.planilhaMensal is not None and Planilhas.planilhaMinibio is not None) else labelMensagem.config(text='Carregue as duas planilhas primeiro!'))
buttonProcessa.config(state='disabled')
buttonProcessa.pack(pady=5)

labelMensagem = Label(janela, text='')
labelMensagem.config(bg=labelMensagem.master.cget('bg'), bd=0, relief='flat', width=30)
labelMensagem.pack(pady=10)

diretorioImagem = armazenaImagem('iconeRefresh.png')
imagemButtonRefresh = PhotoImage(file=diretorioImagem)
buttonRefresh = Button(janela, image=imagemButtonRefresh, command=lambda:resetaTudo())
buttonRefresh.config(bg=buttonRefresh.master.cget('bg'), bd=0, relief='flat', width=30)
buttonRefresh.pack(side=RIGHT, padx=10, pady=10)

janela.mainloop()