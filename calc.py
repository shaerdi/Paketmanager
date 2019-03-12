###############################################################################
# Benoetigte Module
###############################################################################
import numpy as np
import pandas as pd
import pathlib
import xlsxwriter
from collections import defaultdict


###############################################################################
# Hilfsfunktionen
###############################################################################
class idCounter:
    """Zaehler fuer die Paketnummern
    """
    counter = 0
    def __init__(self):
        self.id = idCounter.counter
        idCounter.counter += 1
        self.count = 0
    def increase(self):
        self.count += 1

def convertLeistung(l):
    """Macht aus einer Zahl eine Buchstabenfolge (String)

    :returns: Den string im Format xx.xxxx
    """
    try:
        return '{:07.4f}'.format(float(l))
    except:
        return str(l)

def datenEinlesen(dateiname):
    """Liest ein Excel ein

    :returns: Ein pandas Objekt mit allen Daten im ersten Sheet des Excels und
    eine Liste mit den Kategorien aus dem zweiten Sheet des Excels

    """
    if '.xls' in dateiname:
        daten = pd.read_excel(
                dateiname,
                converters = {'Leistung':convertLeistung},
                )
        try:
            kategorien = pd.read_excel(
                    dateiname,
                    sheet_name=1,
                    converters = {0:convertLeistung},
                    header = None,
                    )
            kategorien = kategorien.values.flatten()
            print("Verwende Kategorien:")
            for k in kategorien:
                print(k)
        except IndexError:
            print("Keine Kategorien gefunden. Gibt es ein zweites Sheet in der Datei {}?".format(dateiname))
            return
        return daten,kategorien
    else:
        print("Die Datei hat nicht die Endung .xls(x)")


def leistungenFiltern(daten):
    """Filtert die Daten, so dass nur noch Zeilen mit der Kategorie Tarmed
    vorhanden sind

    :daten: Pandas Objekt mit Daten
    :returns: Gefilterte Daten

    """
    datenNurTarmed = daten[daten.Leistungskategorie == 'Tarmed']
    leistungen = np.unique(datenNurTarmed.Leistung)
    return leistungen

def getKategorie(group, kategorien):
    """Sucht die Kategorie einer Gruppe
    """
    tarmedgroup = group[group.Leistungskategorie=='Tarmed']
    leistungen = tarmedgroup.Leistung.values
    if len(leistungen) == 0:
        return 'OhneTarmed'

    for k in kategorien:
        if k in leistungen:
            return k
    return 'Restgruppe'


def createKey(group):
    """Baut eine Gruppen ID

    Die ID ist ein langer String mit allen Leistungsnummer sortiert
    aneinandergehaengt.
    """
    tarmedgroup = group[group.Leistungskategorie=='Tarmed']
    key = ''.join([
            str(l) for l in
            sorted(np.unique((tarmedgroup.Leistung.values)))
            #sorted((tarmedgroup.Leistung.values))
            ])
    return key

def sheetSchreiben(sheetname, daten, writer):
    """Schreibt Daten in ein neues sheet in einem Excel
    """
    # Daten schreiben
    daten.to_excel(writer,sheet_name=sheetname, index=False)

    # Zeilen nach paketID abwechselnd faerben
    paketID = daten['paketID'].values
    paketIDwechsel = 1 * (np.absolute(np.diff(paketID)) > 0)
    paketIDwechsel = np.hstack(([0],paketIDwechsel))
    paketIDwechsel[ np.where(paketIDwechsel)[0][1::2] ] = -1
    graueZeilen = np.where(np.cumsum(paketIDwechsel))[0] + 1
    sheet = writer.sheets[sheetname]
    workbook = writer.book
    cell_format_grey = workbook.add_format({'bg_color':'#dddddd'})
    for zeile in graueZeilen:
        sheet.set_row(zeile,None,cell_format_grey)

def getFirstGroup(groups):
    for i,g in groups:
        return g
###############################################################################
# Hauptfunktion, geht alle Leistungen durch und schreibt sie in ein Excel
###############################################################################
def createPakete(daten,kategorien):
    idCounter.counter=0
    pakete = defaultdict(lambda : idCounter())

    # Erster Durchgang:
    # Die Daten werden nach FallDatum gruppiert, und jede Gruppe wird in die
    # Pakete einsortiert.
    for falldatum, group in daten.groupby('FallDatum'):
        key = createKey(group)
        pakete[key].increase()
        daten.loc[group.index,'paketID'] = pakete[key].id
        daten.loc[group.index,'kategorie'] = getKategorie(group,kategorien)

    # Zweiter Durchgang:
    # Die Anzahl jedes Pakets wird in die entsprechende Zeile geschrieben
    for falldatum, group in daten.groupby('FallDatum'):
        key = createKey(group)
        daten.loc[group.index,'Anzahl'] = pakete[key].count


    # Daten absteigend nach Anzahl sortieren
    daten=daten.sort_values(['Anzahl','Leistungskategorie'], ascending=False)

    return daten,pakete


def writePaketeToExcel(daten, kategorien, pakete, filename, ordner = 'Resultate'):
    """Schreibt die Pakete in ein neues Excel

    :daten: Pandas objekt mit allen Daten
    :kategorien: Liste mit den Kategorien
    :filename: Filename der neuen Resultatdatei
    :ordner: Ordner, in dem das Resultat gespeichert wird
    """

    #############################################
    # Ab jetzt der Code um das Excel zu schreiben
    #############################################

    # Pfad zur Resultatdatei
    fname = pathlib.Path('.') / ordner / filename
    fname = fname.with_suffix('.xlsx')

    # Ordner erstellen, wenn es ihn noch nicht gibt
    if not fname.parent.exists():
        fname.parent.mkdir()
        
    # Excel erstellen
    writer = pd.ExcelWriter(str(fname), engine='xlsxwriter')
    workbook = writer.book

    ## Daten Schreiben
    # Rohdaten
    daten.drop(['kategorie'],axis=1).to_excel(
            writer, sheet_name='Rohdaten', index=False
            )

    # Alle Pakete, jeweils die erste Fallnummer im entsprechenden Paket
    allePakete = pd.concat(sorted([
            getFirstGroup(g[1].groupby('FallDatum',sort=False))
            for g in daten.drop('kategorie',axis=1).groupby('paketID',sort=False)
            ],
            key = lambda x : x.Anzahl.max(), reverse = True
            ))
    sheetSchreiben('AllePakete', allePakete ,writer)

    # Pro Kategorie
    for kategorie in kategorien.tolist() + ['Restgruppe','OhneTarmed']:
        katData = daten[daten['kategorie']==kategorie].drop('kategorie',axis=1)
        try:
            katData = pd.concat(sorted([
                getFirstGroup(g[1].groupby('FallDatum',sort=False))
                for g in katData.groupby('paketID',sort=False)
                ],
                key = lambda x : x.Anzahl.max(), reverse = True
                ))
            sheetSchreiben(kategorie,katData,writer)
        except ValueError:
            pass

    workbook.close()
    return pakete,daten


def excelBearbeiten(
        inputDatei,
        resultatDatei,
        ordner = 'Resultate',
        ) :
    result = datenEinlesen(inputDatei)
    if result is None:
        return
    daten,kategorien = result
    pakete,daten = createPakete(daten,kategorien)
    writePaketeToExcel(daten, kategorien, pakete, resultatDatei, ordner)
    pkts = []
    for pid, paket in pakete.items():
        p = paket
        p.key = pid
        pkts.append(p)
    return pkts,daten



# Beispiel fuer das Filtern von Paketen
"""    
pakete = excelBearbeiten('./Rohdaten/2018.12.05_Q1-3_2018_Pneumo.xls', 'ollipolli')

# pakete ist eine Liste mit allen Paketen
#Jedes Paket hat
#    paket.id => Paketnummer im Excel
#    paket.count => Anzahl vorkommen
#    paket.key => String mit allen Leistunen im Paket aneinandergehaengt

# Um pakete zu filtern, Beispiel:
gefilterteListe = [] # Leere Liste zu Beginn
for paket in pakete: # fuer jedes paket in der liste
    if '00.0010' in paket.key and '00.0020' in paket.key and not '00.0050' in paket.key:
        gefilterteListe.append(paket)

# in gefilterter Liste sind jetyt alle pakete die diese bedingung erfuellen
# zum Ausgeben ihrer excel id:
for paket in gefilterteListe:
    print(paket.id)
"""
