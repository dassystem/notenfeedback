#!/usr/bin/python3

# Benoetigt: ezodf (python) und seine Abhaenigkeiten, Python3, pdflatex, pdfnup

# Lauffaehig unter Linux und Windows.

# Dieses Skript setzt aus einer Tabellenkalkulationsdatei ein LaTeX Dokument fuer Schuelerfeedback.
# Zu benutzten mit "$ python generate-feedback.py [[DATEINAME.ods]]"

# Geschrieben von Adrian Salamon, 2016-08-26.

import os
import sys
import platform
import datetime


# # # # # # # # # # # # # # # # # #
#
# Inputfile verarbeiten und Sheets anlegen
#
# # # # # # # # # # # # # # # # # #

input_datei = sys.argv[1]

from ezodf import opendoc, Sheet
doc = opendoc(input_datei)
#for sheet in doc.sheets:
   #print(sheet.name)
notenuebersicht = doc.sheets['Notenuebersicht']
somiListe       = doc.sheets['SoMi-Liste']
settings        = doc.sheets['Kursliste-Settings']


# # # # # # # # # # # # # # # # # #
#
# Metadaten des Kurses und Abschnittes einlesen
#
# # # # # # # # # # # # # # # # # #

kursname      = notenuebersicht['B1']
lehrkraftName = settings['C1']
zeitraumVon   = somiListe['C1']
zeitraumBis   = somiListe['W1']



latex_name="feedbackausgabe.tex"

if platform.system() == 'Linux': #Datei löschen Befehl Linux
        os.system("rm -f "+latex_name)
if platform.system() == 'Windows': #Datei löschen Befehl Windows
        os.system("del "+latex_name)


latex_pre ="""\\documentclass[a6paper,10pt]{scrartcl}
\\usepackage{tabularx}
\\usepackage[T1]{fontenc}
\\usepackage[utf8]{inputenc}
\\usepackage{lmodern}
\\usepackage[left=1cm,right=1cm,bottom=1cm]{geometry}

\\usepackage[ngerman]{babel}

\\usepackage[automark,headsepline]{scrpage2}
\\pagestyle{scrheadings}
\\cfoot{}
\\ihead{"""

latex_pre=latex_pre+str(kursname.value)+"""}
\\chead{}
\\ohead{"""
latex_pre=latex_pre+datetime.datetime.strptime(\
str(zeitraumVon.value), '%Y-%m-%d').strftime('%d.%m.%Y') + \
" bis " + datetime.datetime.strptime(\
str(zeitraumBis.value), '%Y-%m-%d').strftime('%d.%m.%Y') + """}


\\pagestyle{scrheadings}

\\begin{document}\n\n"""

latex_post ="""\\end{document}"""


# # # # # # # # # # # # # # # # # #
#
# Daten der Lernenden einlesen und verabeiten
#
# # # # # # # # # # # # # # # # # #

druck = ""
for i in range (5,36):
	lernerName = notenuebersicht['B'+str(i)] #Namen der Lernenden speichern
	if str(lernerName.value) != "0.0": #Wenn SuS einen Namen hat, dann mache weiter

    # Klausuren
		klausur1 = notenuebersicht['D'+str(i)] 
		klausur1wert = str(klausur1.value)
		if klausur1wert == "None":
			klausur1wert = "Keine 1. Klausurnote"

		klausur2 = notenuebersicht['E'+str(i)] 
		klausur2wert = str(klausur2.value)
		if klausur2wert == "None":
			klausur2wert = "Keine 2. Klausurnote"

		klausur3 = notenuebersicht['E'+str(i)] 
		klausur3wert = str(klausur3.value)
		if klausur3wert == "None":
			klausur3wert = "Keine 3. Klausurnote"					
					
          
    # SoMi      
		somiQ1 = notenuebersicht['H'+str(i)] 
		somiQ1Wert = str(somiQ1.value)
		if somiQ1Wert == "None":
			somiQ1Wert = "Keine SoMi-Note für Q1"

		somiQ2 = notenuebersicht['I'+str(i)] 
		somiQ2Wert = str(somiQ2.value)
		if somiQ2Wert == "None":
			somiQ2Wert = "Keine SoMi-Note für Q2"
		
    # Kommentar - Feedback		
		kommQ1 = notenuebersicht['R'+str(i)] 
		kommQ1Wert = str(kommQ1.value)
		if kommQ1Wert == "None":
			kommQ1Wert = "Kein Kommentar für Q1"
			
		kommQ2 = notenuebersicht['S'+str(i)]
		kommQ2Wert = str(kommQ2.value)
		if kommQ2Wert == "None":
			kommQ2Wert = "Kein Kommentar für Q2"
			
      
      			
		druck = druck + "\\section*{"+ str(lernerName.value) +"} \\begin{tabularx}{\\textwidth}{lX}\n Klausurnote: &"+str(klausur1wert) + "\\\\\n SoMi-Q1: &" +str(somiQ1Wert) + "\\\\\n Kommentar: &"+kommQ1Wert +"\\end{tabularx}\n\n \\vfill " + str(lehrkraftName.value) + ", \\today\n \clearpage" +"\n \n \n"

# # # # # # # # # # # # # # # # # #
#
# LaTeX aufruf für gesamtes PDF
#
# # # # # # # # # # # # # # # # # #

with open(latex_name, 'a') as latex_file:
		latex_file.write(latex_pre)
		latex_file.write(druck)
		latex_file.write(latex_post)	
		latex_file.close()

os.system ("pdflatex " + latex_name)
os.system ("pdfnup feedbackausgabe.pdf --nup 4x2 --frame true --outfile print-feedback.pdf") #8 auf 1 drucken
if platform.system() == 'Linux':                  #Datei löschen Befehl Linux
        os.system ("rm -f *.aux *.synctex* *.log")#pdflatex hilfsdateien entfernen
if platform.system() == 'Windows':                #Datei löschen Befehl Windows
        os.system ("del *.aux *.synctex* *.log")	#pdflatex hilfsdateien entfernen

