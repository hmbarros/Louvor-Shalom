import pandas as pd
import json


class deep_magic:
    def __init__(self) -> None:
        # iniciar variáveis globais
        self.musicas = pd.DataFrame()
        self.string = ''

    def tinker_bell(self, file=''):
        self.musicas = pd.read_excel(file)

        colunas = ['Andamento']

        for coluna in colunas:
            if coluna not in self.musicas.columns:
                raise KeyError(
                    f'A coluna {coluna} não foi encontrada no arquivo Excel fornecido')

    def matilda(self, rapidas=1, lentas=2):
        # Selecionar as musicas aleatoriamente com base no andamento e salvar em musicas
        df_lento = self.musicas[self.musicas['Andamento']
                                == 'Lenta'].sample(lentas)
        df_rapido = self.musicas[self.musicas['Andamento']
                                 == 'Rápida'].sample(rapidas)
        # Concatenando os dois DataFrames
        musicas_selecionadas = pd.concat(
            [df_lento, df_rapido]).reset_index()
        
        for idx, row in musicas_selecionadas.iterrows():
            for coluna in musicas_selecionadas.columns.tolist():
                if coluna != 'index':
                    self.string += f'{coluna} - {row[coluna]}\n'
            self.string += '\n'

        self.string = self.string.replace(' nan', '')


if __name__ == '__main__':
    x = deep_magic()
    x.tinker_bell('musicas.xlsx')
    x.matilda()
