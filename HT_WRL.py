from os import chdir, path, remove, listdir, startfile
from pdfplumber import open as pdfplumber_open
from re import findall

### Wyciagniecie samych plikow z folderow:
#for dirpath, dirnames, filenames in os.walk(r"C:\Users\u331609\Desktop\Projekty\Cardany 30Cr\MINI\CHD"):
#    for j in filenames:
#        if "cardan" in j and "G" in j:
#            copyfile(dirpath+"\\"+j, r"C:\Users\u331609\Desktop\Projekty\Cardany 30Cr\MINI\test"+"\\"+j)
#
###


def extract(p, w):
    chdir(p)
    if path.exists("wyniki.txt"): remove("wyniki.txt")
    onlyfiles = [f for f in listdir(p) if path.isfile(path.join(p, f)) & f.lower().endswith(".pdf")]
    results_file = open("wyniki.txt", "a", encoding="utf8")
    results_file.write("Dystans;HV;Pozycja;Wynik CHD;Kod;Data;Przewodnik;Piec;Spec;Plik")
    results_file.write("\n")


    pattern_kod = "Numer rysunku: (.+)\n"
    pattern_data = "Raport z pomiaru [CI]HD (.+)\n"
    pattern_order = "Materiał : (\d+)"
    pattern_piec = "Numer pieca : (\d+)\n"
    pattern_spec = "IHD :  (?:[a-zA-Z]+)?(\d+)"
    pattern = " (\d\d\d) (\d[.,]\d\d)"
    pattern_core = "([123]) <?(\d\d\d)"
    pattern_position = "Materiał :(?:.*)(R|O|Szb)"
    pattern_CHD_result = " \d\d\d \d[.,]\d\d (\d[.,]\d\d)"
    for i in onlyfiles:
        w['-PROGRESS-FILE-TAB3-'].update(i)
        w.refresh()
        g = []
        wyniki = []
        f = pdfplumber_open(i)
        text = f.pages[0].extract_text()
        if not text:
            continue
        r_kod = findall(pattern_kod, text)
        if r_kod == []:
            r_kod = [""]
        r_data = findall(pattern_data, text)
        if r_data == []:
            r_data = [""]
        r_order = findall(pattern_order, text)
        if r_order == []:
            r_order = [""]

        r_piec = findall(pattern_piec, text)
        if r_piec == []:
            r_piec = [""]

        r_spec = findall(pattern_spec, text)
        if r_spec == []:
            r_spec = [""]

        r_position = findall(pattern_position, text)
        if r_position == []:
            r_position = [""]

        r_CHD_result = findall(pattern_CHD_result, text)
        if r_CHD_result == []:
            r_CHD_result = [""]


        r = findall(pattern, text)

        r_core = findall(pattern_core, text)

        for k in r:
            results_file.write(
                ";".join(k) + ";" + r_position[0] +";"+ r_CHD_result[0] +";" + r_kod[0] + ";" + r_data[0] + ";" + r_order[0] + ";" + r_piec[0] + ";" + r_spec[0] + ";" + i)
            results_file.write("\n")

        if r_position[0] == "R" or r_position[0] == "r":
            for c in r_core:
                results_file.write(";".join(c[::-1]) + ";" + r_position[0] + ";" + "-" + ";" + r_kod[0] + ";" + r_data[0] + ";" + r_order[0] + ";" + r_piec[0] + ";" + "-" + ";" + i)
                results_file.write("\n")



    results_file.close()
    startfile("wyniki.txt")
    return True

