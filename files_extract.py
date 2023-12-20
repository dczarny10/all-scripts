from os import path, walk, stat, environ, remove, startfile
from zipfile import ZipFile
from csv import writer, QUOTE_MINIMAL
import PySimpleGUI as sg


def extract(p, w, z):
    desktop = path.join(path.join(environ['USERPROFILE']), 'Desktop')
    desktop_onedrive_danfoss = path.join(path.join(environ['USERPROFILE'], 'OneDrive - Danfoss'), 'Desktop')
    desktop_onedrive_white = path.join(path.join(environ['USERPROFILE'], 'OneDrive - White Drive Products'), 'Desktop')
    results_file = desktop + r"\Files_check.csv"
    results_file_onedrive_danfoss = desktop_onedrive_danfoss + r"\Files_check.csv"
    results_file_onedrive_white = desktop_onedrive_white + r"\Files_check.csv"
    if path.exists(results_file):
        try:
            remove(results_file)
        except:
            #sg.Popup(f"Zamknij proszę plik Excel z wynikami: {results_file}")
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


    files = []
    extensions = []
    full_paths = []
    folders_paths = []
    bad_zip_files = []
    for dirpath, dirnames, filenames in walk(p):
        w['-FILEPATH-'].update(dirpath)
        w.refresh()
        for f in filenames:
            fp = path.join(dirpath, f)
            split_tup = path.splitext(f)

            if z:
                if split_tup[1] == '.zip':
                    listoffiles = []
                    try:
                        with ZipFile(fp) as zipObj:
                            listoffiles = zipObj.namelist()

                            for member in zipObj.infolist():
                                file_path = path.join(dirpath, member.filename)
                                if not path.exists(file_path):
                                    zipObj.extract(member, dirpath)
                        zipObj.close()

                    except:
                        bad_zip_files.append(path.join(dirpath, f))
                        continue

                    for i in listoffiles:
                        file = path.basename(i)
                        if file != '':
                            splited_zip = path.splitext(file)
                            files.append(splited_zip[0])
                            extensions.append(splited_zip[1])

                            folders_zip = path.dirname(i)
                            folders_paths.append(path.join(dirpath, folders_zip).replace('/', '\\'))
                            full_paths.append(path.join(dirpath, folders_zip, file).replace('/', '\\'))

            files.append(split_tup[0])
            extensions.append(split_tup[1])
            folders_paths.append(dirpath)
            full_paths.append(fp)



    header = ['Full path', 'Folder path', 'File name', 'Extension']


    with open(output_file, mode='wb') as file:
        file.write(u'\ufeff'.encode('utf-8'))
    file.close()

    with open(output_file, mode='a', newline='', encoding='utf-8' ) as file:
        wr = writer(file, delimiter=';', quotechar='"', quoting=QUOTE_MINIMAL)
        wr.writerow(header)

        p = zip(full_paths, folders_paths, files, extensions)
        for count, row in enumerate(p):
            k = list(row)
            wr.writerow(k)

    file.close()
    startfile(output_file)
    if bad_zip_files:
        sg.Popup(f'Uszkodzone pliki zip:\n{bad_zip_files}')
    return True



