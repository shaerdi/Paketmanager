###############################################################################
# Benoetigte Module
###############################################################################
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
        kategorien = pd.read_excel(
                dateiname,
                sheet_name=1,
                converters = {0:convertLeistung},
                )
        return daten,kategorien.values.flatten()
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
def writePaketeToExcel(daten, kategorien, filename, ordner = 'Resultate'):
    """Schreibt die Pakete in ein neues Excel

    :daten: Pandas objekt mit allen Daten
    :kategorien: Liste mit den Kategorien
    :filename: Filename der neuen Resultatdatei
    :ordner: Ordner, in dem das Resultat gespeichert wird
    """

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
        katData = pd.concat(sorted([
            getFirstGroup(g[1].groupby('FallDatum',sort=False))
            for g in katData.groupby('paketID',sort=False)
            ],
            key = lambda x : x.Anzahl.max(), reverse = True
            ))
        sheetSchreiben(kategorie,katData,writer)

    workbook.close()


def excelBearbeiten(
        inputDatei,
        resultatDatei,
        ordner = 'Resultate',
        ) :
    daten,kategorien = datenEinlesen(inputDatei)
    writePaketeToExcel(daten, kategorien, resultatDatei, ordner)

