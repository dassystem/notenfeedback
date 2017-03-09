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
# String in Excel-Spalte finden
#
# # # # # # # # # # # # # # # # # #

def findString(myString, myTabelle, myXlsSpalte, rangeTo, rangeFrom=0):
        for i in range(rangeFrom, rangeTo):
                if str(myTabelle[str(myXlsSpalte)+str(i)].value)==myString:
                        return i
        return 0
         

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
#notenuebersicht = doc.sheets['Notenuebersicht']
klausurergebnis = doc.sheets['Klausurergebnis']
#somiListe       = doc.sheets['SoMi-Liste']
#settings        = doc.sheets['Kursliste-Settings']


# # # # # # # # # # # # # # # # # #
#
# Metadaten des Kurses und Abschnittes einlesen
#
# # # # # # # # # # # # # # # # # #

kursname        = klausurergebnis['B'+str(findString("Kursname:", klausurergebnis, 'A', 10))]
lehrkraftName   = klausurergebnis['B'+str(findString("Dozent:", klausurergebnis, 'A', 10))]
klausurdatum    = klausurergebnis['B'+str(findString("Datum:", klausurergebnis, 'A', 10))]
lernfeld        = klausurergebnis['B'+str(findString("Lernfeld:", klausurergebnis, 'A', 10))]
teilnehmer      = klausurergebnis['B'+str(findString("Teilnehmer:", klausurergebnis, 'A', 10))]
maximalpunkte   = klausurergebnis['B'+str(findString("Maximalpunkte:", klausurergebnis, 'A', 10))]
durchschnittsnote = klausurergebnis['B'+str(findString("Durchschnittsnote:", klausurergebnis, 'A', 10))]
#zeitraumBis   = somiListe['W1']



latex_name="klausurfeedbackausgabe.tex"

if platform.system() == 'Linux': #Datei löschen Befehl Linux
        os.system("rm -f "+latex_name)
if platform.system() == 'Windows': #Datei löschen Befehl Windows
        os.system("del "+latex_name)


# # # # # # # # # # # # # # # # # #
#
# Latex Präambel
#
# # # # # # # # # # # # # # # # # #

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
\\ohead{Klausur vom """
latex_pre=latex_pre+datetime.datetime.strptime(\
str(klausurdatum.value), '%Y-%m-%d').strftime('%d.%m.%Y') + """}

\\pagestyle{scrheadings}

\\begin{document}\n\n"""


# # # # # # # # # # # # # # # # # #
#
# Notenspiegel
#
# # # # # # # # # # # # # # # # # #

notenspiegelZeile = findString("Notenspiegel:", klausurergebnis, 'A', 10)

notenspiegel = """\n\\subsection*{Notenspiegel}\n
\\begin{center}
\\begin{tabularx}{\\linewidth}{|@{} *6{>{\\centering\\arraybackslash}X|}@{}}
 1 & 2 & 3 & 4 & 5 & 6 \\\\\\hline\n"""
notenspiegel = notenspiegel +  str(klausurergebnis['B'+str(notenspiegelZeile+1)].value) + " & "
notenspiegel = notenspiegel +  str(klausurergebnis['C'+str(notenspiegelZeile+1)].value) + " & "
notenspiegel = notenspiegel +  str(klausurergebnis['D'+str(notenspiegelZeile+1)].value) + " & "
notenspiegel = notenspiegel +  str(klausurergebnis['E'+str(notenspiegelZeile+1)].value) + " & "
notenspiegel = notenspiegel +  str(klausurergebnis['F'+str(notenspiegelZeile+1)].value) + " & "
notenspiegel = notenspiegel +  str(klausurergebnis['G'+str(notenspiegelZeile+1)].value) + """ \\\\
\\end{tabularx}
\\end{center}

\\begin{tabularx}{\\linewidth}{lX}
"""

notenspiegel = notenspiegel + "Teilnehmer: &" + str(teilnehmer.value)+ "\\\\ \n"
notenspiegel = notenspiegel + "Durchschnittsnote: &"+str(durchschnittsnote.value) +"\n"
notenspiegel = notenspiegel + """
\\end{tabularx}\n\n 

"""


# # # # # # # # # # # # # # # # # #
#
# LaTeX-Dokumentabschluss
#
# # # # # # # # # # # # # # # # # #

latex_post ="""\\end{document}"""


# # # # # # # # # # # # # # # # # #
#
# Daten der Lernenden einlesen und verabeiten
#
# # # # # # # # # # # # # # # # # #

druck = ""
for i in range (12,51):
	lernerName = klausurergebnis['A'+str(i)] #Namen der Lernenden speichern
	if str(lernerName.value) != "None": #Wenn SuS einen Namen hat, dann mache weiter

        # Klausurergebnis
		note = klausurergebnis['B'+str(i)] 
		notenWert = str(note.value)
		if notenWert == "None":
			notenWert = "Keine Klausurnote"

		punkte = klausurergebnis['C'+str(i)] 
		punkteWert = str(punkte.value)
		if punkteWert == "None":
			punkteWert = "Keine Punkte"
				
					
          
    
		
        # Kommentar - Feedback		
		komm = klausurergebnis['H'+str(i)] 
		kommWert = str(komm.value)
		if kommWert == "None":
			kommWert = "Kein Kommentar"
	
		druck = druck +"\\section*{"+ str(lernerName.value) +"} \n\\begin{tabularx}{\\linewidth}{lX}\n Lernfeld(er): &" + str(lernfeld.value)+ "\\\\ \n Klausurnote: &"
		druck = druck +str(notenWert) +"\\\\\n Erreichte Punkte: &" +str(punkteWert) +" von " +str(maximalpunkte.value)
		druck = druck +"\\\\\n Kommentar: &"+kommWert
		druck = druck +"\n\\end{tabularx}\n\n \\vfill " 

		druck = druck + notenspiegel + "\n\n \\vfill " + str(lehrkraftName.value) +", \\today\n \clearpage" +"\n \n \n"
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
os.system ("pdfnup klausurfeedbackausgabe.pdf --nup 4x2 --frame true --outfile print-klausurfeedback.pdf") #8 auf 1 drucken
if platform.system() == 'Linux':                  #Datei löschen Befehl Linux
        os.system ("rm -f *.aux *.synctex* *.log")#pdflatex hilfsdateien entfernen
if platform.system() == 'Windows':                #Datei löschen Befehl Windows
        os.system ("del *.aux *.synctex* *.log")	#pdflatex hilfsdateien entfernen

