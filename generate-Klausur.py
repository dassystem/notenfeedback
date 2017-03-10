#!/usr/bin/python3

# Benoetigt: ezodf (python) und seine Abhaenigkeiten, Python3, pdflatex, pdfnup

# Lauffaehig unter Windows.

# Dieses Skript generiert aus der Kursliste und der aktet-Schule-Klausur.tex eine xls Tabelle
# Zu benutzten mit "$ python generate-klausur.py [[DATEINAME.ods]] [[DATEINAME.tex]]"

# Geschrieben von Jörn Hornig, 2017-03-10.

import os
import sys
import platform
import datetime
import mmap


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

input_datei1 = sys.argv[1]
input_datei2 = sys.argv[2]

from ezodf import opendoc, Sheet
doc1 = opendoc(input_datei1)
kursliste        = doc1.sheets['Kursliste-Settings']

odsName="klausur1.ods"

if platform.system() == 'Linux': #Datei löschen Befehl Linux
        os.system("copy blanko-klausur.ods " + odsName)
if platform.system() == 'Windows': #Datei löschen Befehl Windows
        #os.system("del " + odsName)
        os.system("copy blanko-klausur.ods " + odsName)

doc2 = opendoc(odsName)
klausur        = doc2.sheets['Klausur']        


# # # # # # # # # # # # # # # # # #
#
# tex File mit Paket-Schule nach Aufgaben durchsuchen und in xls schreiben
#
# # # # # # # # # # # # # # # # # #

gesamtpunkte = 0

with open(input_datei2) as texfile:
        anzAufgaben = 0
        for line in texfile:
                #print(line.find("\\begin{aufgabe}"))
                if line.find("\\begin{aufgabe}")>-1: #Finde Aufgabe in tex
                        if line.find("%", 0, line.find("\\begin{aufgabe}"))==-1: #überprüfe ob Aufgabe auskommentiert ist
                                line = line.strip()
                                anzAufgaben = anzAufgaben + 1
                                punkte = line[line.find("}{")+2:-1]
                                gesamtpunkte = gesamtpunkte + int(punkte)
                                klausur[chr(anzAufgaben+64+6)+'1'].set_value("A" + str(anzAufgaben))
                                klausur[chr(anzAufgaben+64+6)+'2'].set_value(int(punkte))

print(str(anzAufgaben) + " Aufgaben")
print("Gesamtpunke: "+ str(gesamtpunkte))


# # # # # # # # # # # # # # # # # #
#
# Daten Kursliste einlesen und in xls schreiben
#
# # # # # # # # # # # # # # # # # #

for i in range (5,40):
        vorname = kursliste['B'+str(i)] #Vornamen speichern
        nachname = kursliste['C'+str(i)] #Nachnamen speichern
        if str(vorname.value) != "None" and str(nachname.value) != "None":
                klausur['B'+str(i-1)].set_value((vorname.value) + " " + str(nachname.value))


doc2.saveas(odsName)
