###############################################################################
# Benoetigte Module
###############################################################################
import pandas as pd
import pathlib
import xlsxwriter
import itertools
from collections import defaultdict


###############################################################################
# Hilfsfunktionen
###############################################################################
counter = itertools.count()
class idCounter:
    """Zaehler fuer die Paketnummern
    """
    def __init__(self):
        self.id = next(counter)
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

    :returns: Ein pandas Objekt mit allen Daten im ersten Sheet des Excels

    """
    if '.xls' in filename:
        daten = pd.read_excel(
                filename,
                converters = {'Leistung':convertLeistung},
                )
        return daten
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

# def setColWidth(sheet, lens):
    # """Setzt die Spaltenbreite fuer alle Spalten in einem Excel
    # """
    # for i, w  in enumerate(lens):
        # sheet.set_column(i,i,w)

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

    # Pfad zur Resultatdatei
    fname = pathlib.Path('.') / ordner / filename
    fname = fname.with_suffix('.xlsx')

    # Ordner erstellen, wenn es ihn noch nicht gibt
    if not fname.parent.exists():
        fname.parent.mkdir()


    
    # lens = [
            # 1+max([len(str(s)) for s in daten[col].values] + [len(col)]) 
            # for col in daten.columns
            # ]

    ## Hilfsfunktionen
    def getKatNum(group):
        """Sucht die Nummer einer Gruppe
        """
        tarmedgroup = group[group.Leistungskategorie=='Tarmed']
        leistungen = tarmedgroup.Leistung.values
        if len(leistungen) == 0:
            return len(kategorien)+1

        for i,k in enumerate(kategorien):
            if k in leistungen:
                return i
        return i+1


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
        
    pakete = defaultdict(lambda : idCounter())

    # Erster Durchgang:
    # Die Daten werden nach FallDatum gruppiert, und jede Gruppe wird in die
    # Pakete einsortiert.
    for falldatum, group in daten.groupby('FallDatum'):
        key = createKey(group)
        pakete[key].increase()
        daten.loc[group.index,'paketID'] = pakete[key].id

    # Zweiter Durchgang:
    # Die Anzahl jedes Pakets wird in die entsprechende Zeile geschrieben
    for falldatum, group in daten.groupby('FallDatum'):
        key = createKey(group)
        daten.loc[group.index,'Anzahl'] = pakete[key].count

    #############################################
    # Ab jetzt der Code um das Excel zu schreiben
    #############################################
    # Excel erstellen
    writer = pd.ExcelWriter(str(fname), engine='xlsxwriter')
    # Rohdaten schreiben
    daten.to_excel(writer, sheet_name='Rohdaten', index=False)
    workbook = writer.book
    sheet1 = writer.sheets['Rohdaten']

    # Excel formatierungen
    # Titel fett (bold) und unterstrichen
    title_format = workbook.add_format()
    title_format.set_bold()
    title_format.set_bottom()

    # Zwei Zellenformate, eines Standard und eines mit grauem (dddddd)
    # Hintergrund
    cell_format1 = workbook.add_format()
    cell_format2 = workbook.add_format()
    cell_format2.set_bg_color('#dddddd')

    # Neues sheet mit Namen AllePakete erstellen
    sheet_alle = workbook.add_worksheet('AllePakete')
    counter_alle = 1

    # Neues sheet fuer jede Kategorie erstellen
    sheets = [workbook.add_worksheet(l) for l in kategorien]
    # Zwei neue sheets fuer den Rest
    sheets.append(workbook.add_worksheet('Restgruppe'))
    sheets.append(workbook.add_worksheet('OhneTarmed'))

    # Zaehler zum Festhalten der aktuellen Zeile
    counters = [1 for l in kategorien] + 2*[1]
    cycler_alle = itertools.cycle( (cell_format1, cell_format2) )
    cycler = [
                itertools.cycle( (cell_format1, cell_format2) )
                for i in range(len(kategorien)+2)
             ] 



    # Die Daten werden an den entsprechenden Ort im Excel geschrieben
    for sheet in [sheet_alle] + sheets:
        # Titel schreiben
        sheet.write_row(
                'A1', 
                daten.columns.values,
                title_format,
                )

    # Daten absteigend sortieren nach Anzahl und Leistungskategorie
    daten=daten.sort_values(['Anzahl','Leistungskategorie'], ascending=False)

    for j,distGroup in daten.groupby('paketID',sort=False):
        for k,fallGroup in distGroup.groupby('FallDatum'):
            katnum = getKatNum(fallGroup)
            clr = next(cycler[katnum])
            clr_alle = next(cycler_alle)
            sheet = sheets[katnum]
            currentRow = counters[katnum]
            for colNum,colName in enumerate(distGroup.columns):
                sheet.write_column(
                        currentRow,
                        colNum,
                        fallGroup[colName],
                        clr
                        )
                sheet_alle.write_column(
                        counter_alle,
                        colNum,
                        fallGroup[colName],
                        clr
                        )
            break
        counters[katnum] += len(fallGroup)
        counter_alle += len(fallGroup)
    
    workbook.close()

def doCalcs():
    files = [
            './Rohdaten/Q1-Q3_2018_Angio_alle.xlsx',
            './Rohdaten/Q1-Q3_2018_Endo_alle.xlsx',
            './Rohdaten/Q1-Q3_2018_Gastro_alle.xlsx',
            './Rohdaten/Q1-Q3_2018_RangesPneumo_alle.xlsx',
            ]

    resultFilenames = [
            l.replace('./Rohdaten/','').replace('.xlsx','') + '_Eingeteilt'
            for l in files
            ]

    for f,r in zip(files,resultFilenames):
        c = Categorize(f)
        c.doCalc(20)
        c.writeExcel(r)

