import csv

donnees = []
with open('ADECal.csv', newline="") as csvfile:
  reader = csv.reader(csvfile, delimiter = ",") 
  for row in reader:
    donnees.append(row)
del(donnees[0])

print(donnees[0])

donnees_tri_date = []
maxi = donnees[0][1][10]
for i in range(len(donnees)):

    
#print(len(donnees))