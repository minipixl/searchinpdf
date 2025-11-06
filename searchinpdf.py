# Use: for Windows 11 - opens console to enter path and textstring to search for
# purpose: search in all PDF files within specified path and all subdirectories
# console shows file name, page numbers of the matches and total number of matches
# outputs an excel file 'matches.xlsx' with the number of matches and hyperlinks

import os
from pdfminer.layout import LAParams
from pdfminer.converter import PDFResourceManager, PDFPageAggregator
# from pdfminer.pdfpage import PDFPage
from pdfminer.layout import LTTextBoxHorizontal
# from pdfminer.pdfparser import PDFParser
# from pdfminer.pdfdocument import PDFDocument
from pdfminer.pdfpage import PDFPage
# from pdfminer.pdfpage import PDFTextExtractionNotAllowed
from pdfminer.pdfinterp import PDFResourceManager
from pdfminer.pdfinterp import PDFPageInterpreter
# from pdfminer.pdfdevice import PDFDevice
import pandas as pd


print('Seaches for text string in PDF files in specified path and below. Leave out mutated vowels.')
pfad = input('Enter search path : ')
print(pfad + '\\' + 'matches.xlsx')
suchtext = input('search string: ')

# initialise output file
ausgabe = [] # Initialisierung Ausgabe als Liste

anzahldat = 0
anzmeldtot = 0 # total no of matches

# determine file count for console
for subdir, dirs, files in os.walk(pfad):
    for file in files: # loop over files
        anzahldat+= 1
print('Number of files for search: ', anzahldat)

# loop over dirs and files
for subdir, dirs, files in os.walk(pfad):
    for file in files: # each file
        filepath = subdir + os.sep + file
        anzmelddat = 0 # number of matches per file

        if filepath.endswith(".pdf"):  # each PDF
            # initialise count of matches
            meldung = filepath + " ¦ page(s): "
            treffer = False
            pageno = 0 # initialize page number
            document = open(filepath, 'rb')
            #Create resource manager
            rsrcmgr = PDFResourceManager()
            # Set parameters for analysis.
            laparams = LAParams()
            # Create a PDF page aggregator object.
            device = PDFPageAggregator(rsrcmgr, laparams=laparams)
            interpreter = PDFPageInterpreter(rsrcmgr, device)
            for page in PDFPage.get_pages(document, check_extractable=False): # for each page
                interpreter.process_page(page)
                # receive the LTPage object for the page.
                layout = device.get_result()
                pageno += 1 # page number
                for element in layout:
                    if isinstance(element, LTTextBoxHorizontal):
                        aktuellertext = element.get_text()
                        #print(aktuellertext)
                        if suchtext in aktuellertext:
                            # append page number
                            meldung = meldung + " " + str(pageno)
                            anzmeldtot += 1
                            anzmelddat += 1
                            treffer = True
                        else:
                            # nichts machen
                            continue
            if treffer:
                print(meldung)  # output for each file
                # create hyperlink and append match count
                linkliste = {
                    'anzmelddat': anzmelddat,
                    'link': f'=HYPERLINK("{filepath}", "{os.path.basename(filepath)}")'
                }
                ausgabe.append(linkliste)
            else:
                continue
        else: # jedes PDF
            continue
df = pd.DataFrame(ausgabe)
df.rename(columns={'anzmelddat': 'Count of matches', 'link': 'Link'}, inplace=True)
df.to_excel(pfad + '\\' + 'matches.xlsx', sheet_name='Links', engine='xlsxwriter', index= False)
print("Total number of matches: ", anzmeldtot)
input("press Enter to open 'matches.xlsx'...")
os.startfile(pfad + '\\' + 'matches.xlsx')

