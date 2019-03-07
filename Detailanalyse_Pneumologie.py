# -*- coding: utf-8 -*-
"""
Created on Thu Jan 31 19:20:12 2019

@author: olivia
"""

#pakete,daten = excelBearbeiten('./Rohdaten/2018.12.05_Q1-3_2018_Pneumo.xls', 'test040219')
# Liste mit immer zwei Eintraegen, Grundkriterium und Nebenbedingungen
bedingungsListe = {
 '008' :  ( ['15.0710'],           ['00.0010','00.0050','00.0055','00.0056','00.0110','00.0131','00.0141','00.0161','00.0136','00.0146', '00.0166','00.0138','00.0148', '00.0168','00.0415','00.0416','00.0417','00.0610','00.0615','00.0616','00.0855','00.2285','00.2310', '04.0100','05.0560','13.0020', '15.0040','15.0060','15.0110','15.0130','15.0160','15.0200','15.0240','15.0270','15.0285', '15.0300','15.0320', '15.0340','15.0410','15.0630','15.0720','15.0730','15.0740','15.0750','16.0010', '17.0010','19.0020','35.0210','39.0020','39.0500','39.3310','39.3420','39.3700', '39.3710']),
 '009' :  (['15.0720'],            ['00.0010','00.0050','00.0055','00.0056','00.0110','00.0131','00.0141','00.0161','00.0136','00.0146', '00.0166','00.0138','00.0148', '00.0168','00.0415','00.0416','00.0417','00.0610','00.0615','00.0616','00.0855','00.2285','00.2310', '04.0100','05.0560','13.0020', '15.0040','15.0060','15.0110','15.0130','15.0160','15.0200','15.0240','15.0270','15.0285', '15.0300','15.0320', '15.0340','15.0410','15.0630','15.0710','15.0730','15.0740','15.0750','16.0010', '17.0010','19.0020','35.0210','39.0020','39.0500','39.3310','39.3420','39.3700', '39.3710']),
 '010' :  (['15.0730'],            ['00.0010','00.0050','00.0055','00.0056','00.0110','00.0131','00.0141','00.0161','00.0136','00.0146', '00.0166','00.0138','00.0148', '00.0168','00.0415','00.0416','00.0417','00.0610','00.0615','00.0616','00.0855','00.2285','00.2310', '04.0100','05.0560','13.0020', '15.0040','15.0060','15.0110','15.0130','15.0160','15.0200','15.0240','15.0270','15.0285', '15.0300','15.0320', '15.0340','15.0410','15.0630','15.0710','15.0720','15.0740','15.0750','16.0010', '17.0010','19.0020','35.0210','39.0020','39.0500','39.3310','39.3420','39.3700', '39.3710']),
 '011' :  (['15.0740'],            ['00.0010','00.0050','00.0055','00.0056','00.0110','00.0131','00.0141','00.0161','00.0136','00.0146', '00.0166','00.0138','00.0148', '00.0168','00.0415','00.0416','00.0417','00.0610','00.0615','00.0616','00.0855','00.2285','00.2310', '04.0100','05.0560','13.0020', '15.0040','15.0060','15.0110','15.0130','15.0160','15.0200','15.0240','15.0270','15.0285', '15.0300','15.0320', '15.0340','15.0410','15.0630','15.0710','15.0720','15.0730','15.0750','16.0010', '17.0010','19.0020','35.0210','39.0020','39.0500','39.3310','39.3420','39.3700', '39.3710']),
}

# Datenliste enthaelt die einzelnen Bedingungen
datenListe = []
for nummer, bedingung in bedingungsListe.items():
    kriterienID = []
    for paket in pakete:
        k = paket.key
        erfuelltBedingung = (
                   all( [(b in k) for b in bedingung[0]] )
               and any( [(b in k) for b in bedingung[1]])
               )
        if erfuelltBedingung:
            kriterienID.append(paket.id)
    zeilen = daten['paketID'].apply(lambda x : x in kriterienID)
    gefilterteDaten = daten.loc[zeilen]
    gefilterteDaten['Regel'] = nummer
    datenListe.append(gefilterteDaten)

for liste in datenListe:
    liste.drop_duplicates(subset='FallDatum',inplace=True)
#Zusammensetzen der Listen zu einem Excel
zusammengesetzt = pd.concat(datenListe)

# Speichern
zusammengesetzt.to_excel('./Pneumo_060319.xlsx',index=False)

# -*- coding: utf-8 -*-
# """
# Created on Thu Jan 31 19:20:12 2019

# @author: simon
# """

# #pakete,daten = excelBearbeiten('./Rohdaten/2018.12.05_Q1-3_2018_Pneumo.xls', 'ollipolli')
# # Liste mit immer zwei Eintraegen, Positiv und Negativ
# bedingungsListe = {
 # '1' :  (['15.0630'],            ['15.0200']),
 # '2' :  (['15.0630'],            ['00.0610','00.0615','00.0616']),
 # '3' :  (['15.0630', '15.0620'], []),
 # '4' :  (['15.0630'],            ['00.0010']),
 # '5' :  (['15.0630'],            ['00.0415']),
# }

# # Datenliste enthaelt die einzelnen Bedingungen
# datenListe = []
# for nummer, bedingung in bedingungsListe.items():
    # kriterienID = []
    # for paket in pakete:
        # k = paket.key
        # erfuelltBedingung = all(
            # [(b in k) for b in bedingung[0]] + 
            # [(b not in k) for b in bedingung[1]]
            # )
        # if erfuelltBedingung:
            # kriterienID.append(paket.id)
    # zeilen = daten['paketID'].apply(lambda x : x in kriterienID)
    # gefilterteDaten = daten.loc[zeilen]
    # gefilterteDaten['Regel'] = nummer
    # datenListe.append(gefilterteDaten)

# for liste in datenListe:
    # liste.drop_duplicates(subset='FallDatum',inplace=True)
# #Zusammensetzen der Listen zu einem Excel
# zusammengesetzt = pd.concat(datenListe)

# # Speichern
# zusammengesetzt.to_excel('./test2.xlsx',index=False)
