from csv import writer, QUOTE_MINIMAL
from os import path, walk, stat, environ, remove, startfile
import PySimpleGUI as sg

def get_size(windo, start_path = '.'):
    dict = {}
    for dirpath, dirnames, filenames in walk(start_path):
        windo['-DIRPATH-'].update(dirpath)
        windo.refresh()
        total_size = 0
        files = 0
        for f in filenames:
            fp = path.join(dirpath, f)
            # skip if it is symbolic link
            if not path.islink(fp):
                #total_size += os.path.getsize(fp)
                total_size += stat(fp).st_size
        dict.setdefault(dirpath, [])
        #dict.setdefault(dirpath, str(total_size/1048576).replace(".", ","));
        #dict[dirpath].append(str(total_size/1048576).replace(".", ","))
        dict[dirpath].append(total_size / 1048576)
        dict[dirpath].append(len(filenames))

    for g in dict:
        found = [key for key, value in dict.items() if g in key]
        found.remove(g)
        for f in found:
            dict[g][0] += dict[f][0]
            dict[g][1] += dict[f][1]
    return dict

def extract(p, w):
    desktop = path.join(path.join(environ['USERPROFILE']), 'Desktop')
    desktop_onedrive_danfoss = path.join(path.join(environ['USERPROFILE'], 'OneDrive - Danfoss'), 'Desktop')
    desktop_onedrive_white = path.join(path.join(environ['USERPROFILE'], 'OneDrive - White Drive Products'), 'Desktop')
    results_file = desktop + r"\Folders_check.csv"
    results_file_onedrive_danfoss = desktop_onedrive_danfoss + r"\Folders_check.csv"
    results_file_onedrive_white = desktop_onedrive_white + r"\Folders_check.csv"
    if path.exists(results_file):
        try:
            remove(results_file)
        except:
            #sg.popup_error(f"Zamknij proszę plik Excel z wynikami: {results_file}")
            return False
    elif path.exists(results_file_onedrive_danfoss):
        try:
            remove(results_file)
        except:
            #sg.Popup(f"Zamknij proszę plik Excel z wynikami: {results_file_onedrive_danfoss}")
            return False

    elif path.exists(results_file_onedrive_white):
        try:
            remove(results_file)
        except:
            #sg.Popup(f"Zamknij proszę plik Excel z wynikami: {results_file_onedrive_white}")
            return False


    if path.exists(desktop):
        output_file = results_file
    elif path.exists(desktop_onedrive_danfoss):
        output_file = results_file_onedrive_danfoss
    elif path.exists(desktop_onedrive_white):
        output_file = results_file_onedrive_white

    d = get_size(w, p)
    ilosc_plikow = []
    waga_plikow = []
    header = ["Path", "Size [MB]", "Count"]
    for h in list(d.values()):
        waga_plikow.append(str(h[0]).replace('.', ','))
        ilosc_plikow.append(h[1])

    with open(output_file, mode='wb') as file:
        file.write(u'\ufeff'.encode('utf-8'))
    file.close()

    with open(output_file, mode='a', newline='', encoding='utf-8') as file:
        wr = writer(file, delimiter=';', quotechar='"', quoting=QUOTE_MINIMAL)
        wr.writerow(header)
        #p = zip(d.keys(), list(d.values())[0],list(d.values())[1])
        p = zip(d.keys(), waga_plikow, ilosc_plikow)
        for count, row in enumerate(p):
            k = list(row)
            wr.writerow(k)
    startfile(output_file)
    return True