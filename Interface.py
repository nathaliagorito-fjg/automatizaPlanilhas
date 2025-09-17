import pandas as pd
import Planilhas
from pandastable import Table
from tkinter import *
from tkinter.filedialog import askopenfilename

def carregaPlanilhas(tipo):
    diretorio = askopenfilename(filetypes=[('Excel files', '*.xlsx *.xls')])

    if not diretorio:
        return
    
    if tipo == 'mensal':
        if 'mensal' not in diretorio.lower():
            labelMensagem['text'] = 'Planilha errada.'
        else:
            Planilhas.planilhaMensal = pd.read_excel(diretorio)

            labelMensagem['text'] = 'Planilha MENSAL carregada.'
    elif tipo == 'minibio':
        if 'minibio' not in diretorio.lower():
            labelMensagem['text'] = 'Planilha errada.'
        else:
            Planilhas.planilhaMinibio = pd.read_excel(diretorio)

            labelMensagem['text'] = 'Planilha MINIBIO carregada.'

def mostraPlanilha(planilha, titulo):
    if planilha is None or planilha.empty:
        labelMensagem['text'] = 'Nada para mostrar!'
    else:
        novaJanela = Toplevel(janela)
        novaJanela.title(titulo)
        novaJanela.geometry('1095x500')
        novaJanela.resizable(False, False)

        framePlanilha = Frame(novaJanela)
        framePlanilha.pack(fill=BOTH, expand=1)

        tabela = Table(framePlanilha, dataframe=planilha)
        tabela.show()

        #def fechaJanela():
            #planilha.loc[:] = tabela.model.df

            #for idx, linha in planilha.iterrows():
                #for coluna in linha.index:
                    #if Planilhas.planilhaMensal.at[idx, coluna] != linha[coluna]:
                        #Planilhas.planilhaMensal.at[idx, coluna] = linha[coluna]
            
            #Planilhas.planilhaMensal.to_excel('Planilha ALTERADA.xlsx', index=False)

            #novaJanela.destroy()

        #novaJanela.protocol('WM_DELETE_WINDOW', fechaJanela)

def processaPlanilhas():
    nomesDuplicados, planilhasMescladas = Planilhas.processaPlanilhas()
    if nomesDuplicados is None:
        labelMensagem['text'] = 'Carregue as planilhas primeiro!'
        return
    
    buttonMensal.config(state='disabled')
    buttonMinibio.config(state='disabled')

    mostraPlanilha(nomesDuplicados[['NOME','INICIO_LOTACAO','NOMESETOR','ORGAO_ENTIDADE']], 'Nomes Duplicados')
    #mostraPlanilha(datasDuplicadas[['NOME', 'INICIO_LOTACAO', 'NOMESETOR', 'ORGAO_ENTIDADE']], 'Datas Duplicadas')
    mostraPlanilha(planilhasMescladas.loc[planilhasMescladas['IGUAIS'] == False, ['NOME', 'ORGAO_ENTIDADE_MINIBIO', 'ORGAO_ENTIDADE_MENSAL', 'SIGLA', 'IGUAIS']], 'Valores Diferentes')

def resetaTudo():
    buttonMensal.config(state='normal')
    buttonMinibio.config(state='normal')
    labelMensagem['text'] = ''

    janela.protocol('WM_DELETE_WINDOW', janela.destroy())

janela = Tk()
janela.title('Leitor de Planilhas')
janela.geometry('550x400')
janela.resizable(False, False)
janela.option_add('*Font', ('Arial', 8))
janela.option_add('*Foreground', 'black')
janela.option_add('*Background', 'white')
janela.option_add('*Bd', 1)
janela.option_add('*Relief', 'solid')
janela.option_add('*Width', 20)
janela.option_add('*Acivebackground', 'black')
janela.option_add('*Activeforeground', 'white')

geometria = {'padx': 5, 'pady': 5, 'sticky': 'n'}

labelTitulo = Label(janela, text='Processador de Planilhas')
labelTitulo.config(bg=labelTitulo.master.cget('bg'), bd=0, font=15, relief='flat', width=30)
labelTitulo.pack(pady=15)

infos = """
    Este programa realiza:
    1. Elimina nomes da planilha mensal que não estejam na planilha minibio
    2. Elimina da planilha mensal registros antigos de uma mesma pessoa
    3. Exibe registros de ambas planilhas que estejam com valores diferentes para a coluna ORGAO_ENTIDADE
    4. Salva em uma nova planilha as alterações realizadas
    É necessário inserir as duas planilhas para que o processamento ocorra
"""

labelInfo = Label(janela, text=infos)
labelInfo.config(bg=labelInfo.master.cget('bg'), bd=0, justify='left', relief='flat', width=0)
labelInfo.pack(pady=15)

buttonMensal = Button(janela, text='Carregar Mensal', command=lambda: carregaPlanilhas('mensal'))
buttonMensal.pack(pady=5)

buttonMinibio = Button(janela, text='Carregar Minibio', command=lambda: carregaPlanilhas('minibio'))
buttonMinibio.pack(pady=5)

buttonProcessa = Button(janela, text='Processar Planilhas', command=lambda: processaPlanilhas() if (Planilhas.planilhaMensal is not None and Planilhas.planilhaMinibio is not None) else labelMensagem.config(text='Carregue as duas planilhas primeiro!'))
buttonProcessa.pack(pady=5)

imagemButtonRefresh = PhotoImage(file='buttonRefresh.png')
buttonRefresh = Button(janela, image=imagemButtonRefresh, command=lambda:resetaTudo())
buttonRefresh.config(bg=buttonRefresh.master.cget('bg'), bd=0, relief='flat', width=30)
buttonRefresh.pack(side=RIGHT, padx=10)

labelMensagem = Label(janela, text='')
labelMensagem.config(bg=labelMensagem.master.cget('bg'), bd=0, relief='flat', width=30)
labelMensagem.pack(pady=20)

janela.mainloop()