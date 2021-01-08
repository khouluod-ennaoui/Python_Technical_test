# -*- coding: utf-8 -*-
"""
Created on Thu Jan  7 22:47:50 2021

@author: ennao
"""

import xlrd
import json


#tri list function
def triCroissant(L):
    for i in range(1, len(L)):  
        temporaire=L[i]
        j = i - 1 
        while   j >=0 and L[j] > temporaire: 
            L[j+1] = L[j]
            j -= 1
        L[j+1] = temporaire
    return L


#get user List from input
List = list(int(num) for num in input("Entrez les numéros de la liste séparés par un espace: ").strip().split())
print("Votre list est: ", List)

print("\n")
#tri list
print("List dans l'ordre croissant: ")
print(triCroissant(List))




#structured excel data 
# Give the location of the file
location = input("Veuillez saisir l'emplacement du fichier excel: ")

#open excel file 
wb = xlrd.open_workbook(location)
sheet = wb.sheet_by_index(0)

#create list data for json file 
data_list=[]


#order variable to increment
k=0

#structure data 
for i in range(4,sheet.nrows) :
    if(sheet.cell_value(i,0)!="" and sheet.cell_value(i,2)=="") :
         data= {}
         data['order'] = k
         data['type'] = 'SECTION'
         data['reference'] = sheet.cell_value(i,0)
         data['level'] = 1
         data['description'] = sheet.cell_value(i,1)
         data_list.append(data)
         k=k+1
    elif sheet.cell_value(i,2)!="" :
         data= {}
         data['order'] = k
         data['type'] = 'CELL'
         data['description'] = sheet.cell_value(i,1)
         data['unity'] = sheet.cell_value(i,2)
         data['quantity'] = sheet.cell_value(i,3)
         data_list.append(data)
         k=k+1

#create json items dict
Jsonitem={"items" : data_list}


#dump data in json 
j = json.dumps(Jsonitem)

# Write to file
with open('data.json', 'w') as f:
    f.write(j)

print("Succès")




#test unitaire
import unittest
class MyTest(unittest.TestCase):
    def test_tri_croissant(self):
        L=[55, 100, 87, 203, 852, 0, 785, 2, 5, 7]
        self.assertTrue(triCroissant(L),L.sort())
      

if __name__ == '__main__':
    unittest.main()
