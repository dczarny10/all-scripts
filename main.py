import PySimpleGUI as sg
from os import path
import all_files_extract, BOM_extract, HT_WRL, MDP_check, files_extract #, gear_srednice_extract,


sg.theme("DarkTeal12")
# The tab 1, 2, 3 layouts - what goes inside the tab
tab1_layout = [[sg.Text('Skrypt zrzucajacy wyniki pomiarow z plikow PDF od inspektorow jakosci\nW zaleznosci od plikow zrzut moze chwile trwac :)')],
               [sg.Text('Sciezka do folderu z plikami'), sg.Input(size=(12,1), key='-IN-TAB1-'), sg.FolderBrowse()],
               [sg.Button('GO', key='-GO-TAB1-')],
               [sg.Push()],
               [sg.Text('Progress: '), sg.ProgressBar(max_value=100, orientation='h', size=(20, 20), key='-PROGRESS-BAR-')],
               [sg.Text('', key='-PROGRESS-FILE-TAB1-')]]

tab2_layout = [[sg.Text('Skrypt zrzucajacy BOMy Eatona z plikow PDF')],
               [sg.Text('Sciezka do folderu z plikami'), sg.Input(size=(12,1), key='-IN-TAB2-'), sg.FolderBrowse()],
               [sg.Button('GO', key='-GO-TAB2-')],
               [sg.Text('', key='-PROGRESS-FILE-TAB2-')]]

tab3_layout = [[sg.Text('Skrypt zrzucajacy wyniki CHD z plikow PDF')],
               [sg.Text('Sciezka do folderu z plikami'), sg.Input(size=(12,1), key='-IN-TAB3-'), sg.FolderBrowse()],
               [sg.Button('GO', key='-GO-TAB3-')],
               [sg.Text('', key='-PROGRESS-FILE-TAB3-')]]

tab4_layout = [[sg.Text('Skrypt sprawdzajacy wage folderow i ilosci plikow \nPlik z wynikami na pulpicie: Folders_check.csv')],
               [sg.Text('Sciezka do folderu'), sg.Input(size=(12,1), key='-IN-TAB4-'), sg.FolderBrowse()],
               [sg.Button('GO', key='-GO-TAB4-')],
               [sg.Push()],
               [sg.Text('', key='-DIRPATH-')]]

tab5_layout = [[sg.Text('Skrypt tworzący liste plikow i rozszerzen wraz z linkami \nPlik z wynikami na pulpicie: Files_check.csv')],
               [sg.Text('Sciezka do folderu'), sg.Input(size=(12,1), key='-IN-TAB5-'), sg.FolderBrowse()],
               [sg.Checkbox('Sprawdz oraz wypakuj wszystkie pliki ZIP', key='-ZIPEXTRACT-')],
               [sg.Button('GO', key='-GO-TAB5-')],
               [sg.Push()],
               [sg.Text('', key='-FILEPATH-')]]

# The TabgGroup layout - it must contain only Tabs
tab_group_layout = [[sg.Tab('All inspektor', tab1_layout, key='-TAB1-'),
                     sg.Tab('BOM Eaton', tab2_layout, key='-TAB2-'),
                     sg.Tab('CHD', tab3_layout, key='-TAB3-'),
                     sg.Tab('Folders check', tab4_layout, key='-TAB4-'),
                     sg.Tab('Files check', tab5_layout, key='-TAB5-')]]

# The window layout - defines the entire window
layout = [[sg.TabGroup(tab_group_layout,
                       enable_events=True,
                       key='-TABGROUP-', size=(800,300))],
          [sg.Text("Wersja: 1.0"), sg.Push(), sg.Text("W razie problemow: memory->find->Dominik")]]

window = sg.Window('All scripts', layout, no_titlebar=False, font= ('Helvetica', 16), resizable=True, size=(900,400))

tab_keys = ('-TAB1-','-TAB2-','-TAB3-', '-TAB4-', '-TAB5-')         # map from an input value to a key
while True:
    event, values = window.read()       # type: str, dict
    print(event, values)
    if event == sg.WIN_CLOSED:
        break

    elif event == '-TABGROUP-':
        window['-IN-TAB1-'].update([])
        window['-IN-TAB2-'].update([])
        window['-IN-TAB3-'].update([])
        window['-IN-TAB4-'].update([])
        window['-IN-TAB5-'].update([])

    elif event == '-OPERATION DONE-':
        if not values['-OPERATION DONE-']:
            sg.popup_error(f"Zamknij proszę plik Excel z wynikami znajdujacy sie na pulpicie")
        else:
            window['-PROGRESS-FILE-TAB1-'].update('')
            window['-PROGRESS-FILE-TAB2-'].update('')
            window['-PROGRESS-FILE-TAB3-'].update('')
            window['-DIRPATH-'].update('')
            window['-FILEPATH-'].update('')

    elif event == '-GO-TAB1-':
        if not path.exists(values['-IN-TAB1-']):
            sg.Popup('Niepoprawna sciezka')
        else:
            try:
                window.perform_long_operation(lambda: all_files_extract.extract(values['-IN-TAB1-'], window), '-OPERATION DONE-')
            except Exception as e:
                sg.Popup(f"Something is no yes\n\n{e}")
            window['-PROGRESS-BAR-'].update_bar(0)
            window['-PROGRESS-FILE-TAB1-'].update('')

    elif event == '-GO-TAB2-':
        if not path.exists(values['-IN-TAB2-']):
            sg.Popup('Niepoprawna sciezka')
        else:
            try:
                window.perform_long_operation(lambda: BOM_extract.extract(values['-IN-TAB2-'], window), '-OPERATION DONE-')
            except Exception as e:
                sg.Popup(f"Something is no yes\n\n{e}")
            window['-PROGRESS-FILE-TAB2-'].update([])

    elif event == '-GO-TAB3-':
        if not path.exists(values['-IN-TAB3-']):
            sg.Popup('Niepoprawna sciezka')
        else:
            try:
                window.perform_long_operation(lambda: HT_WRL.extract(values['-IN-TAB3-'], window), '-OPERATION DONE-')
            except Exception as e:
                sg.Popup(f"Something is no yes\n\n{e}")
            window['-PROGRESS-FILE-TAB3-'].update('')

    elif event == '-GO-TAB4-':
        if not path.exists(values['-IN-TAB4-']):
            sg.Popup('Niepoprawna sciezka')
        else:
            try:
                window.perform_long_operation(lambda: MDP_check.extract(values['-IN-TAB4-'], window), '-OPERATION DONE-')
                #MDP_check.extract(values['-IN-TAB4-'], window, 'utf-8' if values['-UTF8-'] else 'utf32')
            except Exception as e:
                sg.Popup(f"Something is no yes\n\n{e}")
            window['-DIRPATH-'].update('')


    elif event == '-GO-TAB5-':
        if not path.exists(values['-IN-TAB5-']):
            sg.Popup('Niepoprawna sciezka')
        else:
            try:
                window.perform_long_operation(lambda: files_extract.extract(values['-IN-TAB5-'], window, values['-ZIPEXTRACT-']), '-OPERATION DONE-')
            except Exception as e:
                sg.Popup(f"Something is no yes\n\n{e}")
            window['-FILEPATH-'].update('')

window.close()





