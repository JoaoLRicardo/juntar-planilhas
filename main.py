import PySimpleGUI as sG
import pandas as pd
from glob import glob


sG.theme('Dark Grey 13')


class PrimeiraTela:
    
    def __init__(self):
        layout = [
            [sG.I(), sG.FolderBrowse('Local arquivos', key='diretorio')],
            [sG.Input(), sG.FolderBrowse('Salvar arquivo', key='diretorio_save')],
            [sG.T('Nome do arquivo final:', size=(10, 0)), sG.Input(size=(15, 0), key='nome_arquivo')],
            [sG.OK(), sG.Cancel()]
        ]

        janela = sG.Window('JUNTADOR').layout(layout)
        self.button, self.values = janela.Read()

    def iniciar(self):
        diretorio = self.values['diretorio']        
        diretorio_save = self.values['diretorio_save']        
        nome = self.values['nome_arquivo']

        # JUNTAR TODOS ARQUIVOS EXCEL    
        lista_arquivo = []
        for arquivo in glob(fr'{diretorio}\*xlsx'):
            lista_arquivo.append(arquivo)

        # JUNTAR TODAS TABELAS EXCEL
        tabelas = []
        for arquivo in lista_arquivo:
            tabelas.append(pd.read_excel(arquivo, sheet_name=0))

        # EXTRAIR TABELA FINAL
        tabela_final = pd.concat(tabelas)    
        tabela_final.to_excel(fr'{diretorio_save}\{nome}.xlsx', sheet_name='tabelafinal', header=True, index=False)


tela = PrimeiraTela()
tela.iniciar()
