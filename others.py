Python 3.11.4 (tags/v3.11.4:d2340ef, Jun  7 2023, 05:45:37) [MSC v.1934 64 bit (AMD64)] on win32
Type "help", "copyright", "credits" or "license()" for more information.
>>> import csv
... #from msilib.text import tables
... import openpyxl
... 
... 
... donnees = []
... with open('ADECal.csv', newline="") as csvfile:
...   reader = csv.reader(csvfile, delimiter = ",") 
...   for row in reader:
...     donnees.append(row)
... del(donnees[0])
... #print(donnees[0])
... 
... ###
... ### MatiÃ¨res :
... ###
... 
... Module_RT1 = []
... for i in range(len(donnees)):
...     if "RT1App" in donnees[i][3]:
...         if donnees[i][0] not in Module_RT1:
...             Module_RT1.append(donnees[i][0])
... #print(Module_RT1)
...             
... Module_RT2 = []
... for i in range(len(donnees)):
...     if "RT2App" in donnees[i][3]:
...         if donnees[i][0] not in Module_RT2:
...             Module_RT2.append(donnees[i][0])
... #print(Module_RT2)
...             
... ###
... ### Planning :
... ###
...             
... RT1_S1 = []
... for i in range(len(donnees)):
...     if "RT1-S1" in donnees[i][3]:
...         RT1_S1.append([donnees[i][0], donnees[i][1]])
... #print(RT1_S1)
... 
... 
... RT1_S2 = []
... for i in range(len(donnees)):
...     if "RT1-S2" in donnees[i][3]:
...         RT1_S2.append([donnees[i][0], donnees[i][1]])
... #print(RT1_S1)
... 
... 
... RT2_S3 = []
... for i in range(len(donnees)):
...     if "RT2-S3" in donnees[i][3]:
...         RT2_S3.append([donnees[i][0], donnees[i][1]])
... #print(RT1_S1)
...         
... 
... RT1_App = []
... for i in range(len(donnees)):
...     if "RT1App" in donnees[i][3]:
...         RT1_App.append([donnees[i][0], donnees[i][1]])
... #print(RT1_App)
... 

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


Groupe = input("Quelle annÃ©e souhaitez-vous voir ? RT1 ou RT2 : ")

if Groupe == "RT1" : 
    SousGroupe = input("Vous Ãªtes dans la section RT1, dans quel Groupe voulez-vous Ãªtre dirigÃ© ? Bell (B), Fourier (F), Shannon1 (S1), Shannon2 (S2), Turing (T), Huffman (Hu): ")

if Groupe == "RT2" :
    SousGroupe = input("Vous Ãªtes dans la section RT2, dans quel Groupe voulez-vous Ãªtre dirigÃ© ? Dijkstra (D), Hamming (Ha) : ")

if SousGroupe == "B":
    print(BELL)
if SousGroupe == "F":
    print(FOURIER)
if SousGroupe == "S1":
    print(Shannon1)
if SousGroupe == "S2":
    print(Shannon2)
if SousGroupe == "T":
    print(Turing)
if SousGroupe == "Hu":
    print(Huffman)
if SousGroupe == "D":
    print(Dijkstra)
if SousGroupe == "Ha":
    print(Hamming)


a = True
while a == True:
    Q2 = input("Quel cours voulez-vous voir ?")
    a = False

tableau = openpyxl.Workbook()
sheet = tableau.active
##(numÃ©ro ligne commence Ã  0, numÃ©ro colonne commence Ã  0,'str Ã  Ã©crire'ou int)

sheet.cell.font = openpyxl.styles.Font(name = 'MatiÃ¨res', size = 20, underline = 'single', color = 'FF0000')
for i in range(len(Turing)):
    sheet.cell(i+1, 2).value = str(Turing[i])
    sheet.cell(i+1, 1).value = 'Turing'





tableau.save('calendrier.xlsx')
tableau = openpyxl.load_workbook('calendrier.xlsx')
