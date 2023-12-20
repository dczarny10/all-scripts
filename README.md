# Compilation of scripts with simple GUI
During my work as an Engineer there was always needs for tedious tasks, so naturally I got tired of it and decided to write some scripts to make life easier. Later on I added simple GUI so other people in company could use them
![image](https://github.com/dczarny10/all-scripts/assets/105910358/1892dbdc-503d-43ad-b228-5cfc3c3c02cc)

Modules (in each tab):
* PDF scrapper (CMM measurement reports)
* PDF scrapper (Bill of Materials)
* PDF scrapper (hardness measurements)
* Folders check
* Files check

## PDF scrapper (CMM measurement reports)
Original script that I have worked on. Our quality department was wasting 2hrs each day to manualy rewrite measurement results from PDFs to Excel file :)
![image](https://github.com/dczarny10/all-scripts/assets/105910358/b24325ab-bfc9-4039-8845-415eed9ad7f6)


So this script was made to pull all results from PDFs, but challange here was to handle all different formats, as each CMM machine generated slightly different PDFs
Important modules used here are pdfplumber for PDF scrapping and python RegEx for picking measurement results

User need only to specify path with PDF files and then press *GO*

CSV file is generated in same path that consists of measurement characteristics and their results:
![image](https://github.com/dczarny10/all-scripts/assets/105910358/6652b454-ab98-466e-87d7-c09c2275d806)

## PDF scrapper (Bill of Materials)
Another idea came to me when there was a need to get information about product Bill of Materials (BOM) for thousands of different PDF files. It can be imagined that doing this manually would take a lof of work hours

Each product have it's own specification including BOM in separate PDF file. Parts are scrapped from table:
![image](https://github.com/dczarny10/all-scripts/assets/105910358/1e48c27f-8269-4e98-905a-4eaa6b156e00)

## PDF scrapper (hardness measurements)
Simillar to previous two, but designed to work with different files

![image](https://github.com/dczarny10/all-scripts/assets/105910358/bbb433ce-77fd-4ef2-a8a4-0370c5dde60e)

## Folders check
Script that works through os.walk to list all catalogues in specified path and check their content (file count and total size). Used to compare if all data were sucesfully migrated

![image](https://github.com/dczarny10/all-scripts/assets/105910358/a5a8c9ec-3172-4147-b0b9-ef3636a6bbef)

## Files check

Used to make a list of all files in given path (with full path to file and extension). Usefull when a lof of data is stored in remote server where Windows Explorer search is too slow to reasonably find something
![image](https://github.com/dczarny10/all-scripts/assets/105910358/1a3b7786-970b-4ca2-8ed3-690540030985)


