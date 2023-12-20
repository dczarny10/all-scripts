from os import chdir, path, remove, listdir, startfile
from pdfplumber import open as pdfplumber_open
from re import findall, IGNORECASE, MULTILINE

def extract(p, w):
    chdir(p)
    if path.exists("wyniki.txt"): remove("wyniki.txt")
    onlyfiles = [f for f in listdir(p) if path.isfile(path.join(p, f)) & f.lower().endswith(".pdf")]
    print(onlyfiles)
    results_file = open("wyniki.txt", "a", encoding="utf8")

    pattern = "\n *(\d+)* *\| *([^|]+?)\| *([^|]+?)\| *([^|]+)\n"
    pattern = "^ *(\d+)* *\| *([^|]+?)\| *([^|]+?)\| *([^|\\n]+)$"

    for i in onlyfiles:
        w['-PROGRESS-FILE-TAB2-'].update(i)
        w.refresh()
        f = pdfplumber_open(i)
        for j in range(len(f.pages)):
            text = f.pages[j].extract_text()
            if not text:
                continue
            if j ==0:
                pattern_model_number = "PRODUCT NO\.? *\: *(.+?)\s+"
                pattern_model_code = "MODEL CODE *\: *(.+)\n"
                pattern_revision = "REVISION *\: *(.+)\n"
                pattern_product_family = "DESCRIPTION:\n(?: *M    PRODUCT - MOTOR\n)* *(.+)\n"

                r_model_number = findall(pattern_model_number, text)
                if r_model_number == []:
                    r_model_number = [""]
                r_model_code = findall(pattern_model_code, text)
                if r_model_code == []:
                    r_model_code = [""]
                r_revision = findall(pattern_revision, text)
                if r_revision == []:
                    r_revision = [""]
                r_product_family = findall(pattern_product_family, text)
                if r_product_family == []:
                    r_product_family = [""]

            # check = findall("ITEM.*PART NUMBER.*QTY.*DESCRIPTION", text, flags=IGNORECASE)
            check = True

            if check:
                r = findall(pattern, text, flags=IGNORECASE | MULTILINE)
                if not r:
                    print(i)
                    # for j1 in range(len(f.pages)):
                        #text = f.pages[j1].extract_text()
                        #pattern_alternative = "\n(\d+)* *([^ \n]+) *([^\n]+?)? *(.+)"
                        #r1 = re.findall(pattern_alternative, text, flags=re.IGNORECASE)
                    table = f.pages[j].extract_table()
                    if table:
                        text = f.pages[0].extract_text()
                        pattern_model_number = "PRODUCT NUMBER\.? *\: *(.+?)\n"
                        pattern_revision = "REVISION *\: *(.+)\n"
                        pattern_product_family = "DESCRIPTION:\n(?: *M    PRODUCT - MOTOR\n)* *(.+)\n"
                        r_model_code = [""]

                        r_model_number = findall(pattern_model_number, text)
                        if r_model_number == []:
                            r_model_number = [""]
                        r_revision = findall(pattern_revision, text)
                        if r_revision == []:
                            r_revision = [""]
                        r_product_family = findall(pattern_product_family, text)
                        if r_product_family == []:
                            r_product_family = [""]

                        for e in table:
                            if not None in e and not 'REV' in e and len(e) == 4:
                                if not e in ( ['ITEM', 'PART NUMBER', 'QTY', 'DESCRIPTION'], ['', '', '', '', '', ''], ['', '', '', '']):
                                    print(e)
                                    results_file.write(
                                        ";".join(e) + ";" + r_model_number[0] + ";" + r_revision[0] + ";" + r_model_code[0] + ";" + r_product_family[0] + ";" + i)
                                    results_file.write("\n")


                for k in r:
                    if k not in (('', '', '', ''), ('', ' ', ' ', ' '), ('', 'PART NUMBER    ', 'QTY ', 'DESCRIPTION')):
                        k = (k[0], k[1].strip(' '), k[2], k[3])
                        results_file.write(
                            ";".join(k) + ";" + r_model_number[0] + ";" + r_revision[0] + ";" + r_model_code[0] + ";" + r_product_family[0] + ";" + i)
                        results_file.write("\n")



            #r = re.findall(pattern, text, flags=re.IGNORECASE)
            #if not r:
                #pattern_alternative = "\n(\d+)* *([^ \n]+) *([^\n]+?)? *(.+)"
                #pattern_alternative = "\n(\d+)* *(.+) *(.+) *(.+)\n"
                #r = re.findall(pattern_alternative, text, flags=re.IGNORECASE)
                #print(i)
            #for k in r:
                #results_file.write(";".join(k)+";"+r_model_number[0]+";"+r_revision[0]+";"+r_model_code[0]+";"+i)
                #results_file.write("\n")

    results_file.close()
    startfile("wyniki.txt")
    return True

