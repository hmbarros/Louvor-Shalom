import PySimpleGUI as sg
# from edit_data import edit_archive
# from get_data import open_archive
# from open_codes_sigtap import open_codes
from tkinter.filedialog import asksaveasfilename
import sys
import os
import pandas as pd

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
bc = ('#ffffff', '#2690ce')  # Cor do botão
tabelas_sus = {}
bpas = {}

layout = [
    [
        sg.Column([[sg.Image(logo_Shalom)]],
                  justification='left', expand_x=True),
        sg.Column([[sg.Text('')]], expand_x=True),
        sg.Column([[sg.Text('IPShalom - Seletor de Musicas para Louvor',
                  font=('Arial', 14, 'bold'), text_color='#074e2e')]], expand_x=True),
        sg.Column([[sg.Text('')]], justification='right', expand_x=True),
        sg.Column([[sg.Image(logo_IPB)]], justification='right')
    ],
    [sg.Text('Selecione o arquivo Excel (.xlsx) com os dados da BPA.',
             font=bf, text_color='#2E2E2E')],
    [sg.Input(key='-FILE-', font=bf, background_color='#f0f0f0', text_color='#2E2E2E', expand_x=True,),
     sg.FileBrowse(font=bf, button_color=bc, button_text='Abrir')],
    # Elemento de texto para exibir o erro
    [sg.Button('OK',
               key='OK',
               font=bf,
               button_color=bc),
     sg.Button('Salvar',
               key='SAVE',
               font=bf,
               button_color=bc,
               disabled=True),
     sg.Button('Cancel',
               key='Cancel',
               font=bf,
               button_color=bc)],

    [sg.Text('', key='-ERROR-', font=bf, text_color='red')],
    [sg.Text('Selecione o arquivo excel com as tabelas de códigos os procedimentos.', key='-LABEL-PROCEDURE-',
             font=bf, text_color='#CBCCCD')],
    [sg.Input(key='-FILE-PROCEDURE-', font=bf, background_color='#f0f0f0', text_color='#2E2E2E', expand_x=True, disabled=True),
     sg.FileBrowse(key='-BUTTON-PROCEDURE-', font=bf, button_color=bc, button_text='Abrir', disabled=True)],
    [sg.Text('', key='-ERROR2-', font=bf, text_color='red')],
]

window = sg.Window('IPShalom - Seletor de Musicas para Louvor', layout,
                   finalize=True, resizable=True, size=(800, 300))

while True:
    event, values = window.read()

    if event == sg.WINDOW_CLOSED or event == 'Cancel':
        break
    elif event == 'OK':
        window['-ERROR-'].update(text_color='red')
        filename = values['-FILE-']

        try:
            if filename:
                bpa = pd.read_excel(filename)
                window['-ERROR-'].update('')

                if 'CODIGO_SUS' not in bpa.columns:
                    raise ValueError("Coluna CODIGO_SUS não encontrada")

                elif 'QTDE' not in bpa.columns:
                    raise ValueError("Coluna QTDE não encontrada")

                bpa['QTDE'] = bpa['QTDE'].astype(int)

                bpas['BPA_limpo'] = bpa[['CODIGO_SUS', 'QTDE']]
                bpas['BPA_limpo_agrupado'] = bpas['BPA_limpo'].groupby(
                    'CODIGO_SUS')['QTDE'].sum().reset_index()

                window['-ERROR-'].update('')

        # Tratando Erros
        except Exception as e:
            window['-ERROR-'].update(f"Erro: {e}")

        if not window['-ERROR-'].get():
            window['SAVE'].update(disabled=False)
            window['-LABEL-PROCEDURE-'].update(text_color='#000000')
            window['-FILE-PROCEDURE-'].update(disabled=False)
            window['-BUTTON-PROCEDURE-'].update(disabled=False)

        else:
            window['SAVE'].update(disabled=True)
            window['-LABEL-PROCEDURE-'].update(text_color='#CBCCCD')
            window['-FILE-PROCEDURE-'].update(disabled=True)
            window['-BUTTON-PROCEDURE-'].update(disabled=True)

    elif event == 'SAVE':
        window['-ERROR-'].update('')
        window['-ERROR2-'].update('')

        filename_procedure = values['-FILE-PROCEDURE-']

        try:
            tabela_sus_excel = pd.read_excel(
                filename_procedure, sheet_name=None)

            for nome_tabela, df in tabela_sus_excel.items():
                tabelas_sus[nome_tabela] = df

            # Iterar sobre os valores de 'CODIGO_SUS'
            for indice, codigo_sus, qtde in bpas['BPA_limpo'][['CODIGO_SUS', 'QTDE']].itertuples(index=True):
                # Flag para indicar se uma correspondência foi encontrada
                correspondencia_encontrada = False
                # Iterar sobre as tabelas SUS para encontrar correspondência
                for nome_tabela, df in tabelas_sus.items():
                    if codigo_sus in df['CODIGO'].values and not correspondencia_encontrada:
                        # Encontrando a linha correspondente no DataFrame da tabela SUS
                        linha_referencia = df[df['CODIGO']
                                              == codigo_sus].iloc[0]
                        # Adicionar o nome da tabela correspondente na coluna 'REFERÊNCIA'
                        bpas['BPA_limpo'].loc[indice,
                                              'REFERÊNCIA'] = nome_tabela
                        # Adicionar o procedimento correspondente na coluna 'PROCEDIMENTO'
                        bpas['BPA_limpo'].loc[indice,
                                              'PROCEDIMENTO'] = linha_referencia['PROCEDIMENTO']
                        # Multiplicar o VALOR da tabela de REFERÊNCIA com a QTDE do bpas['BPA_limpo']
                        valor_referencia = linha_referencia['VALOR']
                        bpas['BPA_limpo'].loc[indice,
                                              'TOTAL'] = valor_referencia * qtde
                        # Definir a flag como True se uma correspondência for encontrada
                        correspondencia_encontrada = True

                # Verificar se nenhuma correspondência foi encontrada e imprimir uma mensagem apropriada
                if not correspondencia_encontrada:
                    bpas['BPA_limpo'].loc[indice,
                                          'REFERÊNCIA'] = "REFERÊNCIA não localizada"
                    bpas['BPA_limpo'].loc[indice, 'PROCEDIMENTO'] = ""
                    bpas['BPA_limpo'].loc[indice, 'TOTAL'] = ""

            for indice, codigo_sus, qtde in bpas['BPA_limpo_agrupado'][['CODIGO_SUS', 'QTDE']].itertuples(index=True):
                # Flag para indicar se uma correspondência foi encontrada
                correspondencia_encontrada = False
                # Iterar sobre as tabelas SUS para encontrar correspondência
                for nome_tabela, df in tabelas_sus.items():
                    if codigo_sus in df['CODIGO'].values and not correspondencia_encontrada:
                        # Encontrando a linha correspondente no DataFrame da tabela SUS
                        linha_referencia = df[df['CODIGO']
                                              == codigo_sus].iloc[0]
                        # Adicionar o nome da tabela correspondente na coluna 'REFERÊNCIA'
                        bpas['BPA_limpo_agrupado'].loc[indice,
                                                       'REFERÊNCIA'] = nome_tabela
                        # Adicionar o procedimento correspondente na coluna 'PROCEDIMENTO'
                        bpas['BPA_limpo_agrupado'].loc[indice,
                                                       'PROCEDIMENTO'] = linha_referencia['PROCEDIMENTO']
                        # Multiplicar o VALOR da tabela de REFERÊNCIA com a QTDE do bpas['BPA_limpo_agrupado']
                        valor_referencia = linha_referencia['VALOR']
                        bpas['BPA_limpo_agrupado'].loc[indice,
                                                       'TOTAL'] = valor_referencia * qtde
                        # Definir a flag como True se uma correspondência for encontrada
                        correspondencia_encontrada = True

                # Verificar se nenhuma correspondência foi encontrada e imprimir uma mensagem apropriada
                if not correspondencia_encontrada:
                    bpas['BPA_limpo_agrupado'].loc[indice,
                                                   'REFERÊNCIA'] = "Referência não localizada"
                    bpas['BPA_limpo_agrupado'].loc[indice, 'PROCEDIMENTO'] = ""
                    bpas['BPA_limpo_agrupado'].loc[indice, 'TOTAL'] = ""

            window['-ERROR-'].update(text_color='blue')
            window['-ERROR-'].update('BPA salva com sucesso')

            file_path_save = asksaveasfilename(
                defaultextension=".xls",
                filetypes=[("Arquivos Excel", "*.xlsx"),
                           ("Todos os arquivos", "*.*")])

            # Criar um objeto ExcelWriter para escrever em um arquivo Excel
            if file_path_save:
                with pd.ExcelWriter(file_path_save) as writer:
                    # Salvar o DataFrame bpa_limpo ordenado na primeira aba
                    bpas['BPA_limpo'].sort_values(by='REFERÊNCIA').to_excel(
                        writer, sheet_name='BPA Limpo', index=False)
                    # Salvar o DataFrame bpa_agrupado ordenado na segunda aba
                    bpas['BPA_limpo_agrupado'].sort_values(by='REFERÊNCIA').to_excel(
                        writer, sheet_name='BPA Agrupado', index=False)

        except Exception as e:
            window['-ERROR2-'].update(str(e))

window.close()
