import csv
#from msilib.text import tables
import openpyxl
from openpyxl.styles import Font, Fill, Border, Side


donnees = []
with open('ADECal.csv', newline="") as csvfile:
  reader = csv.reader(csvfile, delimiter = ",") 
  for row in reader:
    donnees.append(row)
del(donnees[0])
#print(donnees[0])

###
### Matières :
###

Module_RT1 = []
for i in range(len(donnees)):
    if "RT1App" in donnees[i][3]:
        if donnees[i][0] not in Module_RT1:
            Module_RT1.append(donnees[i][0])
#print(Module_RT1)
            
Module_RT2 = []
for i in range(len(donnees)):
    if "RT2App" in donnees[i][3]:
        if donnees[i][0] not in Module_RT2:
            Module_RT2.append(donnees[i][0])
#print(Module_RT2)
            
###
### Planning :
###
            
RT1_S1 = []
for i in range(len(donnees)):
    if "RT1-S1" in donnees[i][3]:
        RT1_S1.append([donnees[i][0], donnees[i][1]])
#print(RT1_S1)


RT1_S2 = []
for i in range(len(donnees)):
    if "RT1-S2" in donnees[i][3]:
        RT1_S2.append([donnees[i][0], donnees[i][1]])
#print(RT1_S1)


RT2_S3 = []
for i in range(len(donnees)):
    if "RT2-S3" in donnees[i][3]:
        RT2_S3.append([donnees[i][0], donnees[i][1]])
#print(RT1_S1)
        

RT1_App = []
for i in range(len(donnees)):
    if "RT1App" in donnees[i][3]:
        RT1_App.append([donnees[i][0], donnees[i][1]])
#print(RT1_App)


RT2_App = []
for i in range(len(donnees)):
    if "RT2App" in donnees[i][3]:
        RT2_App.append([donnees[i][0], donnees[i][1]])
#print(RT1_App)


Shannon1 = []
for i in range(len(donnees)):
    if "RT1Shannon1" in donnees[i][3]:
        Shannon1.append([donnees[i][0], donnees[i][1]])
#print(Shannon1)

Shannon2 = []
for i in range(len(donnees)):
    if "RT1Shannon2" in donnees[i][3]:
        Shannon2.append([donnees[i][0], donnees[i][1]])
#print(Shannon2)


Turing = []
for i in range(len(donnees)):
    if "RT1Turing" in donnees[i][3]:
        Turing.append([donnees[i][0], donnees[i][1]])
#print(Turing)


Huffman = []
for i in range(len(donnees)):
    if "RT1Huffman" in donnees[i][3]:
        Huffman.append([donnees[i][0], donnees[i][1]])
#print(Huffman)


Dijkstra = []
for i in range(len(donnees)):
    if "Dijkstra" in donnees[i][3]:
        Dijkstra.append([donnees[i][0], donnees[i][1]])
#print(RT1_S1)


Hamming = []
for i in range(len(donnees)):
    if "Hamming" in donnees[i][3]:
        Dijkstra.append([donnees[i][0], donnees[i][1]])
#print(RT1_S1)
        

Bell = []
for i in range(len(donnees)):
    if "Bell" in donnees[i][3]:
        Bell.append([donnees[i][0], donnees[i][1]])
#print(RT1_S1)
        

Fourier = []
for i in range(len(donnees)):
    if "Fourier" in donnees[i][3]:
        Fourier.append([donnees[i][0], donnees[i][1]])
#print(RT1_S1)

BELL = [Shannon1, Turing]

FOURIER = [Shannon2, Huffman]
  

RT1 = [BELL, FOURIER, RT1_S1, RT1_S2]
        

RT2 = [Dijkstra, Hamming, RT2_S3]

a = True
while a == True:
	Groupe = input("Quelle année souhaitez-vous voir ? RT1 ou RT2 : ")
	a = False

if Groupe == "RT1" : 
    SousGroupe = input("Vous êtes dans la section RT1, dans quel Groupe voulez-vous être dirigé ? Bell (B), Fourier (F), Shannon1 (S1), Shannon2 (S2), Turing (T), Huffman (Hu): ")

if Groupe == "RT2" :
    SousGroupe = input("Vous êtes dans la section RT2, dans quel Groupe voulez-vous être dirigé ? Dijkstra (D), Hamming (Ha) : ")

if SousGroupe == "B":
    SousGroupe = Bell
    #print(Bell)
if SousGroupe == "F":
    SousGroupe = Fourier
    #print(Fourier)
if SousGroupe == "S1":
    SousGroupe = Shannon1
    #print(Shannon1)
if SousGroupe == "S2":
    SousGroupe = Shannon2
    #print(Shannon2)
if SousGroupe == "T":
    SousGroupe = Turing
    #print(Turing)
if SousGroupe == "Hu":
    SousGroupe = Huffman
    #print(Huffman)
if SousGroupe == "D":
    SousGroupe = Dijkstra
    #print(Dijkstra)
if SousGroupe == "Ha":
    SousGroupe = Hamming
    #print(Hamming)


tableau = openpyxl.Workbook()
sheet = tableau.active
for row in SousGroupe :
	sheet.append(row)


tableau.save('calendrier.xlsx')