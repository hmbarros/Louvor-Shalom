import pandas as pd

class deep_magic:
  def __init__(self) -> None:
    #iniciar variáveis globais
    self.musicas = {'Lentas': [],
               'Rapidas': []}
        

  def tinker_bell(self, file=''):
  
    musicas = pd.read_excel(file)
    #Abrir arquivo excel das músicas e separar por andamento
    musicas_por_andamento = musicas.groupby('Andamento')
    self.musicas['Lentas'] = musicas_por_andamento.get_group('Lenta')
    self.musicas['Rápidas'] = musicas_por_andamento.get_group('Rápida')      


  def matilda(self, rapidas = 1, lentas = 2):
    #Selecionar as musicas aleatoriamente com base no andamento e salvar em musicas
    Musicas_Lentas = self.musicas['Lentas'].sample(n = lentas)
    Musicas_Rapidas = self.musicas['Rápidas'].sample(n = rapidas)
    print(type(Musicas_Rapidas))
    self.musicas_escolhidas = pd.concat({Musicas_Lentas, Musicas_Rapidas})
    print(self.musicas_escolhidas)

  
    # musica_escolhida = {
    #         'Lenta': random.sample(self.musicas['Lenta'], lentas),
    #         'Rapida': random.sample(self.musicas['Rapida'], rapidas)}
    # print("Música Escolhida:")
    # print("Lenta:", musica_escolhida['Lenta'])
    # print("Rápida:", musica_escolhida['Rapida'])
    # pass

x = deep_magic()
x.tinker_bell('musicas.xlsx')
x.matilda()

