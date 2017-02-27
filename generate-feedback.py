#!/usr/bin/python

# Benoetigt: ezodf (python) und seine Abhaenigkeiten, Python3, pdflatex, pdfnup

# Lauffaehig unter Linux.

#Dieses Skript setzt aus einer Tabellenkalkulationsdatei ein LaTeX Dokument fuer Schuelerfeedback.
#Zu benutzten mit "$ python generate-feedback.py [[DATEINAME.ods]]" Geschrieben von Adrian Salamon, 2016-08-26.

import os
import sys

input_datei = sys.argv[1]

from ezodf import opendoc, Sheet
doc = opendoc(input_datei)
#for sheet in doc.sheets:
   #print(sheet.name)
sheet = doc.sheets['Notenrechnen']

kursname = sheet['B1']

latex_name="feedbackausgabe.tex"
os.system("rm -f "+latex_name)


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
\\pagestyle{scrheadings}

\\begin{document}\n\n"""

latex_post ="""\\end{document}"""

druck = ""
for i in range (5,36):
	susname = sheet['B'+str(i)] #Namen der Lernenden speichern
	if str(susname.value) != "0.0": #Wenn SuS einen Namen hat, dann mache weiter

    # Klausuren
		klausur1 = sheet['D'+str(i)] 
		klausur1wert = str(klausur1.value)
		if klausur1wert == "None":
			klausur1wert = "Keine Klausurnote"

		klausur2 = sheet['E'+str(i)] 
		klausur2wert = str(klausur2.value)
		if klausur2wert == "None":
			klausur2wert = "Keine Klausurnote"

		klausur3 = sheet['E'+str(i)] 
		klausur3wert = str(klausur3.value)
		if klausur3wert == "None":
			klausur3wert = "Keine Klausurnote"					
					
          
    # SoMi      
		somiQ1 = sheet['H'+str(i)] 
		somiQ1Wert = str(somiQ1.value)
		if somiQ1Wert == "#DIV/0!":
			somiQ1Wert = "Keine SoMi-Note"

		somiQ2 = sheet['I'+str(i)] 
		somiQ2Wert = str(somiQ2.value)
		if somiQ2Wert == "#DIV/0!":
			somiQ2Wert = "Keine SoMi-Note"
		
    # Kommentar - Feedback		
		kommQ1 = sheet['R'+str(i)] 
		kommQ1Wert = str(kommQ1.value)
		if kommQ1Wert == "None":
			kommQ1Wert = "Kein Kommentar"
			
		kommQ2 = sheet['S'+str(i)]
		kommQ2Wert = str(kommQ2.value)
		if kommQ2Wert == "None":
			kommQ2Wert = "Kein Kommentar"
			
      
      			
		druck = druck + "\\section*{"+ str(susname.value) +"} \\begin{tabularx}{\\textwidth}{lX}\n Klausurnote: &"+str(klausur1wert) + "\\\\\n SoMi-Q1: &" +str(somiQ1Wert) + "\\\\\n Kommentar: &"+kommQ1Wert +"\\end{tabularx}\n\n \\vfill Salamon, \\today\n \clearpage" +"\n \n \n"


with open(latex_name, 'a') as latex_file:
		latex_file.write(latex_pre)
		latex_file.write(druck)
		latex_file.write(latex_post)	
		latex_file.close()

os.system ("pdflatex " + latex_name)
os.system ("pdfnup feedbackausgabe.pdf --nup 4x2 --frame true --outfile print-feedback.pdf") #8 auf 1 drucken
os.system ("rm -f *.aux *.synctex* *.log")		#pdflatex hilfsdateien entfernen
