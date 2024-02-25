import PySimpleGUI as sg
from tkinter.filedialog import asksaveasfilename
import sys
import os
import pandas as pd
from where_the_magic_happens import *

sg.theme('LightGrey6')

# Verifique se estamos executando como um executável PyInstaller one-file bundle
if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
    bundle_dir = getattr(sys, '_MEIPASS')
else:
    bundle_dir = os.path.dirname(os.path.abspath(__file__))

# Adicione o caminho para o arquivo de imagem do logo
path = getattr(sys, '_MEIPASS', os.getcwd())

logo_Shalom = os.path.join(bundle_dir, path+'\\Logo_Shalom-verde-mini.png')
logo_IPB = os.path.join(bundle_dir, path+'\\Logo-ipb-verde-mini.png')

sg.set_global_icon(os.path.join(bundle_dir, path+'\\IPBfav.ico'))

bf = ('Montserrat', 10, 'bold')  # Fonte do botão
bc = ('#ffffff', '#074e2e')  # Cor do botão
df = pd.DataFrame()

def validar_numero(numero):
    if numero == '' or numero.isdigit():
        return True
    else:
        return False

layout = [
    # Barra de título
    [
        sg.Column([[sg.Image(logo_Shalom)]], justification='left',
                  expand_x=True),  # Logotipo Shalom
        sg.Column([[sg.Text('')]], expand_x=True),  # Espaço em branco
        sg.Column([[sg.Text('IPShalom - Seletor de Musicas para Louvor',
                  font=('Arial', 14, 'bold'), text_color='#074e2e')]], expand_x=True),  # Título
        sg.Column([[sg.Text('')]], justification='right',
                  expand_x=True),  # Espaço em branco
        sg.Column([[sg.Image(logo_IPB)]],
                  justification='right')  # Logotipo IPB
    ],

    # Entrada para músicas rápidas e lentas
    [
        # Rótulo para músicas rápidas
        sg.Text('Musicas Rápidas:', font=bf, text_color='#2E2E2E'),
        sg.InputText('1', key='-Rapidas-', size=(5, 1), background_color='#f0f0f0',
                     text_color='#2E2E2E', enable_events=True),  # Campo de entrada para músicas rápidas
        # Rótulo para músicas lentas
        sg.Text('Musicas Lentas:', font=bf, text_color='#2E2E2E'),
        sg.InputText('2', key='-Lentas-', size=(5, 1), background_color='#f0f0f0',
                     text_color='#2E2E2E', enable_events=True)  # Campo de entrada para músicas lentas
    ],

    # Selecionar arquivo Excel
    [sg.Text('Selecione o arquivo Excel (.xlsx) com os dados da BPA.',
             font=bf, text_color='#2E2E2E')],  # Rótulo para seleção de arquivo
    [sg.Input(key='-FILE-', font=bf, background_color='#f0f0f0', text_color='#2E2E2E', expand_x=True,),  # Campo de entrada para arquivo
     sg.FileBrowse(font=bf, button_color=bc, button_text='Abrir')],  # Botão de navegação para seleção de arquivo

    # Botão OK e exibição de erros
    [sg.Button('OK', key='OK', font=bf, button_color=bc)],  # Botão OK
    # Exibição de erros
    [sg.Text('', key='-ERROR-', font=bf, text_color='red')],

    # Botão Copiar e exibição de texto
    [sg.Button('Copiar', font=bf, button_color=bc)],  # Botão Copiar
    [sg.Multiline("", size=(400, 20), key='-TEXT-',
                  background_color='#074e2e', disabled=True)]  # Exibição de texto
]

window = sg.Window('IPShalom - Seletor de Musicas para Louvor', layout,
                   finalize=True, resizable=True, size=(800, 600))

while True:
    event, values = window.read()

    if event == sg.WINDOW_CLOSED:
        break
    elif event == 'OK':
        window['-ERROR-'].update(text_color='red')
        filename = values['-FILE-']

        if not validar_numero(values['-Lentas-']):
            sg.popup('Por favor, digite uma quantia válida de músicas lentas')
        elif not validar_numero(values['-Rapidas-']):
            sg.popup('Por favor, digite uma quantia válida de músicas rápidas.')

        try:
            if filename and validar_numero(values['-Lentas-']) and validar_numero(values['-Rapidas-']):
                window['-ERROR-'].update('')
                data = deep_magic()
                data.tinker_bell(filename)
                data.matilda(int(values['-Rapidas-']), int(values['-Lentas-']))
                window.TKroot.clipboard_clear()  # Limpa o conteúdo atual da área de transferência
                # Adiciona o texto à área de transferência
                window.TKroot.clipboard_append(data.string)
                window['-TEXT-'].update(value=data.string)

        # Tratando Erros
        except Exception as e:
            window['-ERROR-'].update(f"Erro: {e}")


window.close()
