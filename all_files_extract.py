from os import chdir, path, remove, listdir, startfile
from pdfplumber import open as pdfplumber_open
from re import sub, compile, DOTALL, findall
from time import strftime, strptime, ctime
from csv import writer, QUOTE_MINIMAL
from natsort import os_sorted
import PySimpleGUI as sg

def load_files(results, p):
    onlyfiles= []
    while not onlyfiles:
        chdir(p)
        if path.exists(results):
            try:
                remove(results)
            except:
                #sg.Popup(f"Zamknij proszę plik Excel z wynikami: {results}")
                return False
        onlyfiles = os_sorted([f for f in listdir(p) if path.isfile(path.join(p, f)) & (f.endswith(".txt") | f.endswith(".pdf"))])
        if not onlyfiles:
            sg.popup_error("Nie znaleziono w folderze plikow")
            p = "x"
    return onlyfiles

def file_to_text(filename):
    text = []
    soft = ""
    gear = ""
    standard = ""
    hyperlinks:bool
    table_settings = {
        "vertical_strategy": "lines",
        "horizontal_strategy": "lines",
        "snap_tolerance": 1,
        "intersection_x_tolerance": 1, }
    if filename.split('.')[-1] == "txt":
        f = open(filename)
        text.append(f.read())
        f.close()
    elif filename.split('.')[-1] == "pdf":
        f = pdfplumber_open(filename)
        try:
            soft = f.metadata["Author"].lower()
            gear = f.metadata["Title"].lower()
        except:
            sg.Popup(f"To nie jest plik z wynikami pomiarow: {filename}")
            return False
        if "zeiss" in soft or 'deprc01 messraum' in soft:
            if "gear" in gear:
                k = f.pages[0]
                standard = 'iso' if 'ISO' in k.extract_text() else 'ansi'
                if 'iso' in standard:
                    cropped_gear = k.crop((0, 356.4, 612, 396)) # dla gear (0 * float(k.width), 0.45 * float(k.height), 1 * float(k.width), 0.5 * float(k.height))
                    try:
                        text.append(sub("[a-z]", "", cropped_gear.extract_text()))
                    except:
                        text.append(None)
                    cropped_helix = k.crop((91.8, 696.96, 563, 752.4 ) )# dla helix (0.15 * float(k.width), 0.88 * float(k.height), 0.92 * float(k.width), 0.95 * float(k.height))
                    try:
                        text.append(sub("[a-z]", "", cropped_helix.extract_text()))
                    except:
                        text.append(None)
                    try:
                        k = f.pages[1]
                        text_standard = k.extract_text()
                        text_tolerance_y5 = k.extract_text(y_tolerance=5)
                        text.append(findall('[Ff]\w \d+[,.]?\d* \d+', text_standard) + findall('[Ff]\w \d+[,.]?\d* \d+', text_tolerance_y5)) #fp, Fp, fu, Fr
                        text.append(findall('(D\w,\w\w\w) (\d+[,.]\d+)', text_standard)) #Da, Df
                        crop_mdk = k.crop((0.7 * float(k.width), 0.875 * float(k.height), 0.87 * float(k.width), 0.92 * float(k.height)))  # crop dla Mdk
                        crop_mdr = k.crop((0.7 * float(k.width), 0.95 * float(k.height), 0.87 * float(k.width), 0.99 * float(k.height)) ) # crop dla Mdr
                        text.append(tuple(zip(('Mdk', 'Mdk_max', 'Mdk_min'), findall('\d+[,.]\d+', crop_mdk.extract_text())))) #Mdk
                        text.append(tuple(zip(('Mdr', 'Mdr_max', 'Mdr_min'), findall('\d+[,.]\d+', crop_mdr.extract_text()))))  # Mdr

                    except:
                        text.append('')
                        text.append('')
                        text.append('')
                        text.append('')

                else: #dla ANSI
                    cropped_gear = k.crop((0, 364, 612,
                                           396))  # dla gear (0 * float(k.width), 0.45 * float(k.height), 1 * float(k.width), 0.5 * float(k.height))
                    try:
                        text.append(sub("[a-z]", "", cropped_gear.extract_text()))
                    except:
                        text.append(None)
                    cropped_helix = k.crop((91.8, 730, 563,
                                            759.4))  # dla helix (0.15 * float(k.width), 0.88 * float(k.height), 0.92 * float(k.width), 0.95 * float(k.height))
                    try:
                        text.append(sub("[a-z]", "", cropped_helix.extract_text()))
                    except:
                        text.append(None)
                    try:
                        k = f.pages[1]
                        text_standard = k.extract_text()
                        text_tolerance_y5 = k.extract_text(y_tolerance=5)
                        text.append(
                            findall('[Ff]\w \d+[,.]?\d* \d+', text_standard) + findall('[Ff]\w \d+[,.]?\d* \d+',
                                                                                       text_tolerance_y5))  # fp, Fp, fu, Fr
                        text.append(findall('(D\w,\w\w\w) (\d+[,.]\d+)', text_standard))  # Da, Df
                        crop_mdk = k.crop((0.7 * float(k.width), 0.89 * float(k.height), 0.87 * float(k.width),
                                           0.93 * float(k.height)))  # crop dla Mdk
                        crop_mdr = k.crop((0.7 * float(k.width), 0.95 * float(k.height), 0.87 * float(k.width),
                                           0.99 * float(k.height)))  # crop dla Mdr
                        text.append(tuple(zip(('Mdk', 'Mdk_min', 'Mdk_max'),
                                              findall('\d+[,.]\d+', crop_mdk.extract_text()))))  # Mdk
                        text.append(tuple(zip(('Mdr', 'Mdr_min', 'Mdr_max'),
                                              findall('\d+[,.]\d+', crop_mdr.extract_text()))))  # Mdr

                    except:
                        text.append('')
                        text.append('')
                        text.append('')
                        text.append('')


            else:
                if "standardprotocol" in gear:
                    for k in f.pages:
                        text.append(k.extract_text(x_tolerance=3, y_tolerance=0))
                else:
                    for k in f.pages:
                        table = k.extract_table(table_settings)
                        if not table or "" in table[0]:
                            text.append(sub("\|-+|-+\|", "", k.extract_text(x_tolerance=3, y_tolerance=0)))
                            soft="old_zeiss"
        else:
            for k in f.pages:
                cropped = k.crop((0 * float(k.width), 0 * float(k.height), 0.5 * float(k.width), 0.95 * float(k.height)) )
                table = cropped.extract_table({"vertical_strategy": "lines",
        "horizontal_strategy": "text",
        "snap_tolerance": 20,
        "intersection_x_tolerance": 15,
        "join_tolerance": 15,
        "min_words_horizontal": 2,
        "keep_blank_chars": True})
                for j in table:
                    #extracted = (j[0] + "  " + j[1]) if j[1] is not None else j[0]
                    extracted = j[0]+ "\n" + j[1]
                    text.append(extracted)
    return text, gear, soft, standard, select_pattern(soft, gear, filename.split('.')[-1])

def select_pattern(soft, gear, type):
    if type == "pdf":
        if "zeiss" in soft or 'deprc01 messraum' in soft:
            ###if "gear" in gear or "helix" in gear:
                ###pattern = "(?: [-±>]?\d{1,3})(?: [-±>]?\d{1,3})( (?:[-±>]?)\d{1,3})( (?:[-±>]?)\d{1,3})( (?:[-±>]?)\d{1,3})( (?:[-±>]?)\d{1,3})" #pattern gear/helix
            if "old_zeiss" in soft:
                pattern = "(.+?)\n(?: +\d{1,4}[,.]\d{1,5}\n)?(?: +\n)? +(-?\d{1,4}[,.]\d{2,5})(?: +-?\d{1,4}[,.]\d{1,5}){2}" #pattern OLD Zeiss
            else:
                pattern = "(.+?) (-?\d{1,4}[,.]\d{1,5})(?: -?\d{1,4}[,.]\d{1,5}){2}"  # pattern NEW Zeiss
        else:
            pattern = compile("^(.+?)  (?:.+?)(\d{1,4}[.]\d{1,5})$", DOTALL)
    else:
        pattern = "(?:.+?);(.+?;.+?);(-?\d{1,4}[,.]\d{1,6});(?:-?\d{1,4}[,.]\d{1,6});(?:-?\d{1,4}[,.]\d{1,6})" #pattern txt Wenzel
    return pattern

def write_dict(d, tuples, count):
    for h in tuples:
        char_name = h[0].strip(' |-')
        d.setdefault(char_name, {})  #### dopisz do dict nazwa + wynik, usun |----
        d[char_name].setdefault(count, h[1].replace(".", ","))

def create_tuples(tuples, values, keys):
    for count, k in enumerate(keys):
        [tuples.append(j) for j in tuple(zip(k, values[count]))]

def from_dict_to_list(d):
    r = []
    if not d:
        return []
    m = max([max(val.keys()) for _, val in d.items()])
    for i in d.values():
        r.append([i[j] if j in i.keys() else "0" for j in range(m+1)])
    return r



def extract(path1, w):
    results = {}
    file_names = []
    time_of_measurement = []
    finded_tuples =[]
    results_file = "wyniki.csv"
    header =["Plik", "Data modyfikacji pliku (pomiaru)"]
    keys_gear_iso = (('Falfa_lewa_1', 'Falfa_lewa_2', 'Falfa_lewa_3', 'Falfa_lewa_4'),
                 ('Falfa_lewa_Qa', 'Falfa_prawa_Qa'),
                 ('Falfa_prawa_1', 'Falfa_prawa_2', 'Falfa_prawa_3', 'Falfa_prawa_4'),
                 ('Ffalfa_lewa_1', 'Ffalfa_lewa_2', 'Ffalfa_lewa_3', 'Ffalfa_lewa_4'),
                 ('Ffalfa_lewa_Qa', 'Ffalfa_prawa_Qa'),
                 ('Ffalfa_prawa_1', 'Ffalfa_prawa_2', 'Ffalfa_prawa_3', 'Ffalfa_prawa_4'),
                 ('fHalfa_lewa_1', 'fHalfa_lewa_2', 'fHalfa_lewa_3', 'fHalfa_lewa_4'),
                 ('FHalfa_lewa_Qa', 'FHalfa_prawa_Qa'),
                 ('fHalfa_prawa_1', 'fHalfa_prawa_2', 'fHalfa_prawa_3', 'fHalfa_prawa_4'))
    keys_helix_iso = (('FB_lewa_1', 'FB_lewa_2', 'FB_lewa_3', 'FB_lewa_4'),
                  ('FB_lewa_Qa', 'FB_prawa_Qa'),
                  ('FB_prawa_1', 'FB_prawa_2', 'FB_prawa_3', 'FB_prawa_4'),
                  ('FfB_lewa_1', 'FfB_lewa_2', 'FfB_lewa_3', 'FfB_lewa_4'),
                  ('FfB_lewa_Qa', 'FfB_prawa_Qa'),
                  ('FfB_prawa_1', 'FfB_prawa_2', 'FfB_prawa_3', 'FfB_prawa_4'),
                  ('fHB_lewa_1', 'fHB_lewa_2', 'fHB_lewa_3', 'fHB_lewa_4'),
                  ('fHB_lewa_Qa', 'fHB_prawa_Qa'),
                  ('fHB_prawa_1', 'fHB_prawa_2', 'fHB_prawa_3', 'fHB_prawa_4'))

    keys_gear_ansi = (('minus_Falfa_lewa_1', 'minus_Falfa_lewa_2', 'minus_Falfa_lewa_3', 'minus_Falfa_lewa_4'),
                      ('minus_Falfa_lewa_Qa', 'minus_Falfa_prawa_Qa'),
                      ('minus_Falfa_prawa_1', 'minus_Falfa_prawa_2', 'minus_Falfa_prawa_3', 'minus_Falfa_prawa_4'),
                      ('plus_Falfa_lewa_1', 'plus_Falfa_lewa_2', 'plus_Falfa_lewa_3', 'plus_Falfa_lewa_4'),
                      ('plus_Falfa_lewa_Qa', 'plus_Falfa_prawa_Qa'),
                      ('plus_Falfa_prawa_1', 'plus_Falfa_prawa_2', 'plus_Falfa_prawa_3', 'plus_Falfa_prawa_4'))

    keys_helix_ansi = (('FB_lewa_1', 'FB_lewa_2', 'FB_lewa_3', 'FB_lewa_4'),
                      ('FB_lewa_Qa', 'FB_prawa_Qa'),
                      ('FB_prawa_1', 'FB_prawa_2', 'FB_prawa_3', 'FB_prawa_4'))


    files = load_files(results_file, path1)
    if not files:
        return False
    w['-PROGRESS-BAR-'].update_bar(1)
    w['-PROGRESS-FILE-TAB1-'].update('Starting...')
    progress = 0
    progress_step = 100/len(files)

    for count_files, i in enumerate(files):
        finded_tuples = []
        text, gear, soft, standard, pattern = file_to_text(i)
        time_of_measurement.append(strftime("%d.%m.%Y;%H:%M", strptime(ctime(path.getmtime(i)), "%a %b  %d %H:%M:%S %Y")))
        file_names.append(i)
        if "gear" in gear:
            if 'iso' in standard:
                if text[0]:
                ##finded_values = re.findall(pattern, text[0])
                ###create_tuples(finded_tuples, finded_values, keys_gear_iso)
                    splitted = text[0].split()
                    create_tuples(finded_tuples, (splitted[4:8], splitted[8:10], splitted[10:14], splitted[21:25], splitted[25:27], splitted[27:31], splitted[36:40], splitted[40:42], splitted[42:46]), keys_gear_iso)
                    write_dict(results, finded_tuples, count_files)
                if text[1]:
                ###finded_values = re.findall(pattern, text[0])
                ###create_tuples(finded_tuples, finded_values, keys_helix_iso)
                    splitted = text[1].split()
                    create_tuples(finded_tuples, (splitted[3:7], splitted[7:9], splitted[9:13], splitted[18:22], splitted[22:24], splitted[24:27], splitted[33:37], splitted[37:39], splitted[39:43]), keys_helix_iso)
                    write_dict(results, finded_tuples, count_files)

            else: #dla ANSI
                if text[0]:
                    ##finded_values = re.findall(pattern, text[0])
                    ###create_tuples(finded_tuples, finded_values, keys_gear_iso)
                    splitted = text[0].split()
                    create_tuples(finded_tuples, (splitted[20:24], splitted[24:26], splitted[26:30], splitted[37:41], splitted[41:43], splitted[43:47]), keys_gear_ansi)
                    write_dict(results, finded_tuples, count_files)
                if text[1]:
                    ###finded_values = re.findall(pattern, text[0])
                    ###create_tuples(finded_tuples, finded_values, keys_helix_iso)
                    splitted = text[1].split()
                    create_tuples(finded_tuples, (splitted[19:23], splitted[23:27:3], splitted[27:31]), keys_helix_ansi)
                    write_dict(results, finded_tuples, count_files)

            if text[2]:
                temp_dict = {}
                temp_tuples = []
                for oo in text[2]:
                    splitted = oo.split()
                    temp_dict.setdefault(splitted[0], [])
                    temp_dict[splitted[0]].append(splitted[1])
                    temp_dict[splitted[0]].append(splitted[2])

                for jj in temp_dict:
                    if len(temp_dict[jj]) == 4:
                        temp_tuples.append((jj + '_lewa', temp_dict[jj][0]))
                        temp_tuples.append((jj + '_lewa_Qa', temp_dict[jj][1]))
                        temp_tuples.append((jj + '_prawa', temp_dict[jj][2]))
                        temp_tuples.append((jj + '_prawa_Qa', temp_dict[jj][3]))
                    else:
                        temp_tuples.append((jj , temp_dict[jj][0]))
                        temp_tuples.append((jj + '_Qa', temp_dict[jj][1]))

                write_dict(results, temp_tuples, count_files)


            if text[3]:
                write_dict(results, text[3], count_files)

            if text[4]:
                write_dict(results, text[4], count_files)

            if text[5]:
                write_dict(results, text[5], count_files)


        else:
            for j in text:
                finded_tuples = findall(pattern, j)
                write_dict(results, finded_tuples, count_files)
        progress += progress_step
        w['-PROGRESS-BAR-'].update_bar(progress)
        w['-PROGRESS-FILE-TAB1-'].update(i)
        w.refresh()
    header.extend(results.keys())

    with open(results_file, mode='w', newline='', encoding='utf-8') as file:
        wr = writer(file, delimiter=';', quotechar='"', quoting=QUOTE_MINIMAL)
        wr.writerow(header)
        p = zip(*from_dict_to_list(results))  #### transpozycja matrycy wynikow (poziom -> pion)
        for count, row in enumerate(p):
            d = list(row)
            d.insert(0, time_of_measurement[count])
            d.insert(0, file_names[count])
            wr.writerow(d)
    startfile(results_file)
    return True