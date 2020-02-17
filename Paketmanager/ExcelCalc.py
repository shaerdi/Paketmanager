import pickle
from collections import defaultdict
import numpy as np
import pandas as pd
import pathlib
import xlsxwriter

class UIError(Exception):
    pass

def convertLeistung(leistung):
    """Macht aus einer Zahl eine Buchstabenfolge (String)

    :returns: Den string im Format xx.xxxx
    """
    try:
        return '{:07.4f}'.format(float(leistung))
    except ValueError:
        return str(leistung)

def datenEinlesen(dateiname):
    """Liest ein Excel ein

    :returns: Ein pandas Objekt mit allen Daten im ersten Sheet des Excels und
    eine Liste mit den Kategorien aus dem zweiten Sheet des Excels

    """
    if '.xls' in dateiname:
        daten = pd.read_excel(
            dateiname,
            converters={'Leistung':convertLeistung},
        )
        try:
            kategorien = pd.read_excel(
                dateiname,
                sheet_name=1,
                converters={0:convertLeistung},
                header=None,
            )
            kategorien = kategorien.values.flatten()
        except IndexError:
            kategorien = None
    elif '.csv' in dateiname:
        daten = pd.read_csv(
            dateiname,
            converters={'Leistung':convertLeistung},
        )
        kategorien = None
    else:
        raise UIError("Datei hat nicht die Endung '.xls','.xlsx' oder '.csv'")

    benoetigteSpalten = ['FallNr', 'Datumsfeld', 'Tarifgruppe', 'Leistung']
    fehlerMeldung = "Die Spalte {} muss in den Rohdaten vorhanden sein"
    for spalte in benoetigteSpalten:
        if not spalte in daten.columns:
            raise UIError(fehlerMeldung.format(spalte))

    # Serial Date Format von Excel sind Tage seit dem 01.01.1900
    startDate = pd.datetime(1900,1,1) 
    serialDate = (daten['Datumsfeld'] - startDate).dt.days
    if not 'FallDatum' in daten.columns:
        daten['FallDatum'] = pd.to_numeric(
            daten['FallNr'].astype(str) + serialDate.astype(str)
            )
    return daten, kategorien

def sheetSchreiben(sheetname, daten, writer):
    """Schreibt Daten in ein neues sheet in einem Excel
    """
    # Daten schreiben
    daten.to_excel(writer, sheet_name=sheetname, index=False)

    # Zeilen nach paketID abwechselnd faerben
    paketID = daten['paketID'].values
    paketIDwechsel = 1 * (np.absolute(np.diff(paketID)) > 0)
    paketIDwechsel = np.hstack(([0], paketIDwechsel))
    paketIDwechsel[np.where(paketIDwechsel)[0][1::2]] = -1
    graueZeilen = np.where(np.cumsum(paketIDwechsel))[0] + 1
    sheet = writer.sheets[sheetname]
    workbook = writer.book
    cellFormatGrey = workbook.add_format({'bg_color':'#dddddd'})
    for zeile in graueZeilen:
        sheet.set_row(zeile, None, cellFormatGrey)

def getKategorie(key, kategorien):
    if len(key) == 0:
        return 'OhneTarmed'
    for k in kategorien:
        if k in key:
            return k
    return 'Restgruppe'

###############################################################################
# Hauptfunktion, geht alle Leistungen durch und schreibt sie in ein Excel
###############################################################################
def createPakete(daten, kategorien):
    """
    :daten: Pandas objekt mit allen Daten
    :kategorien: Liste mit den Kategorien
    """
    buildKey = lambda s: ','.join(set(s))

    leistungen = daten[daten['Tarifgruppe'].str.contains('TARMED')][['FallDatum', 'Leistung']]
    keys = leistungen.groupby('FallDatum').aggregate(buildKey)
    keys.rename({'Leistung': 'key'}, axis=1, inplace=True)

    alleLeistungen = daten[['FallDatum', 'Leistung']]
    alleKeys = alleLeistungen.groupby('FallDatum').aggregate(buildKey)
    alleKeys.rename({'Leistung': 'keyAlle'}, axis=1, inplace=True)

    daten = daten.join(keys, on='FallDatum')
    daten = daten.join(alleKeys, on='FallDatum')
    daten.fillna({'key':'', 'keyAlle':''}, inplace=True)

    for i, (_, group) in enumerate(daten.groupby('key')):
        daten.loc[group.index, 'paketID'] = int(i)
        daten.loc[group.index, 'Anzahl'] = group['FallDatum'].drop_duplicates().shape[0]

    return daten

def getFirstGroup(groups):
    """Gibt die erste Gruppe eines Groupby Objektes zurueck"""
    for i, g in groups:
        return g

def writePaketeToExcel(daten, kategorien, filename):
    """ Schreibt die Daten in ein Excel, nach kategorien sortiert"""

    fname = pathlib.Path(filename).with_suffix('.xlsx')

    if not fname.parent.exists():
        fname.parent.mkdir()

    writer = pd.ExcelWriter(str(fname), engine='xlsxwriter')
    workbook = writer.book

    if kategorien is not None:
        daten['Kategorie'] = daten['key'].apply(
            lambda k: getKategorie(k, kategorien))

        daten.drop(['Kategorie'], axis=1).to_excel(
            writer, sheet_name='Rohdaten', index=False
            )

        # Alle Pakete, jeweils die erste Fallnummer im entsprechenden Paket
        allePakete = pd.concat(sorted([
            getFirstGroup(g[1].groupby('FallDatum', sort=False))
            for g in daten.drop('Kategorie', axis=1).groupby('paketID', sort=False)
            ], key=lambda x : x.Anzahl.max(), reverse=True))
        sheetSchreiben('AllePakete', allePakete, writer)

        # Pro Kategorie
        for kategorie in kategorien + ['Restgruppe', 'OhneTarmed']:
            katData = daten[daten['Kategorie'] == kategorie].drop('Kategorie', axis=1)
            try:
                katData = pd.concat(sorted([
                    getFirstGroup(g[1].groupby('FallDatum', sort=False))
                    for g in katData.groupby('paketID', sort=False)
                    ], key=lambda x: x.Anzahl.max(), reverse=True))
                sheetSchreiben(kategorie, katData, writer)
            except ValueError:
                pass
    else:
        daten.to_excel(writer, sheet_name='Rohdaten', index=False)
        allePakete = pd.concat(sorted([
            getFirstGroup(g[1].groupby('FallDatum', sort=False))
            for g in daten.groupby('paketID', sort=False)
            ], key=lambda x: x.Anzahl.max(), reverse=True))
        sheetSchreiben('AllePakete', allePakete, writer)

    workbook.close()

class ObserverSubject:
    """Klasse, die eine Liste von Observern hat und diese updaten kann"""

    def __init__(self):
        self._observer = []

    def registerObserver(self, observer):
        """Registriert ein Observerobjekt, das per Aufrufen der Funktion update
        auf Aenderungen aufmerksam gemacht wird.

        :observer: Observer objekt. Muss die Funktion update haben

        """
        self._observer.append(observer)

    def notifyObserver(self):
        """Ruft die Methode update fuer alle Observer auf

        """
        for observer in self._observer:
            observer.update()

class Regel:
    """Stellt eine Regel dar, die ein Paket erfuellen kann oder nicht"""

    UND = 0
    ODER = 1
    NICHT = 2

    def __init__(self, name, daten):
        self.name = name
        self._bedingungen = {
            Regel.UND: [],
            Regel.ODER : [],
            Regel.NICHT : [],
            }
        self.anzahl = '-'
        self._daten = daten
        self._erfuellt = None

    def validateTyp(self, typ):
        """Ueberprueft, ob der Typ ein gueltiger Regel-Typ ist"""
        if not typ in [Regel.UND, Regel.ODER, Regel.NICHT]:
            raise RuntimeError("Unbekannte Bedingung")

    def getDict(self):
        """Erstellt ein Dict aus den Bedingungen dieser Regel

        :returns: Dictionary mit den Eintraegen UND, ODER und NICHT

        """
        return self._bedingungen

    def addLeistung(self, newItem, typ):
        """Fuegt eine neue Leistung zur einer Liste hinzu

        :new_item: Neue Leistung
        :bedingungs_art: Regel.UND, ODER oder NICHT
        """

        self.validateTyp(typ)
        newItem = convertLeistung(newItem)
        self._bedingungen[typ].append(newItem)
        self.update()

    def removeLeistung(self, index, typ):
        """Loescht eine Leistung aus einer Liste

        :index: Index der zu loeschenden Leistung
        :bedingungs_art: Regel.UND, ODER oder NICHT
        """

        self.validateTyp(typ)
        try: 
            self._bedingungen[typ] = [
                    x for i,x in enumerate(self._bedingungen[typ])
                    if i not in index
                    ]
        except TypeError: # index nicht iterierbar
            del self._bedingungen[typ][index]
        self.update()

    def clearItems(self, typ):
        """Loescht die Leistungen aus einer Liste

        :index: Index der zu loeschenden Leistung
        :bedingungs_art: Regel.UND, ODER oder NICHT
        """

        self.validateTyp(typ)
        self._bedingungen[typ] = []
        self.update()

    def update(self):
        """Berechnet die Pakete, die diese Regel erfuellen"""

        if self._daten.dataframe is None:
            self.anzahl = '-'
            return

        def erfuellt(key):
            """Checkt, ob ein Key diese Regel erfuellt"""
            erfuelltalle = all([(k in key) for k in self._bedingungen[Regel.UND]])
            erfuelltoder = len(self._bedingungen[Regel.ODER]) == 0 or \
                           any([(k in key) for k in self._bedingungen[Regel.ODER]])
            erfuelltnot = all([(k not in key) for k in self._bedingungen[Regel.NICHT]])
            return  erfuelltalle and erfuelltoder and erfuelltnot

        ind = self._daten.dataframe.keyAlle.apply(erfuellt)
        self._erfuellt = self._daten.dataframe[ind]
        self.anzahl = str(len(self._erfuellt.groupby('FallDatum')))

    def getAnzahlErfuellt(self):
        """Gibt die Anzahl der Falldaten zurueck, die diese Regel erfuellen
        :returns: Anzahl der Falldaten

        """
        return self.anzahl

    def moveUNDBedingungToTop(self, dataframe):
        """Wenn die Regel eine UND Bedingung enhaelt, wird eine Zeile die diese
        Bedingung erfuellt an die erste Stelle des dataframe geschoben"""
        if self._bedingungen[Regel.UND]:
            bedingung = self._bedingungen[Regel.UND][0]
            inds = dataframe.keyAlle.str.contains(bedingung)
            inds = np.array(inds).nonzero()[0]
            if inds.size > 0:
                swap0, swap1 = dataframe.iloc[0].copy(), dataframe.iloc[inds[0]].copy()
                dataframe.iloc[0], dataframe.iloc[inds[0]] = swap1, swap0
        return dataframe

    def getErfuellt(self):
        """Gibt ein Dataframe zurueck, das alle Falldaten enthaelt, die diese
        Regel erfuellen.

        :return: Pandas DataFrame
        """
        try:
            kopie = self._erfuellt.copy()
            kopie['Regel'] = self.name
            kopie = self.moveUNDBedingungToTop(kopie)
            return kopie
        except AttributeError:
            spalten = list(self._daten.dataframe.columns)
            spalten.append('Regel')
            return pd.DataFrame(columns=spalten)

    def getLeistungen(self, typ):
        """Gibt die Leistungen im Typ der Regel zurueck

        :typ: Regel.UND, Regel.ODER oder Regel.NICHT
        :returns: Liste mit Leistungen
        """
        self.validateTyp(typ)
        return self._bedingungen[typ]

class Regeln(ObserverSubject):
    """Klasse, die die Regeln speichert"""

    def __init__(self, excelDaten):
        super().__init__()
        self.regeln = []
        self._aktiveRegel = None
        self._excelDaten = excelDaten
        self._excelDaten.registerObserver(self)

    def update(self):
        """Wird aufgerufen, wenn die ExcelDaten sich aendern"""
        self.updateRegel()
        self.notifyObserver()

    def updateRegel(self, index=None):
        """Berechnet fuer die Regel die Anzahl der Pakete, die die Regel
        erfuellen.

        :index: Index der zu updatenden Regel. Alle, wenn None
        """
        if index is None:
            for regel in self.regeln:
                regel.update()
        else:
            self.regeln[index].update()

    def addRegel(self, name):
        """Fuegt eine neu Regel hinzu

        :name: Name der neuen Regel
        """
        neueRegel = Regel(name, self._excelDaten)
        self.regeln.append(neueRegel)
        self.notifyObserver()

    def renameRegel(self, index, neuerName):
        """Benennt eine Regel um.

        :index: Index der Regel, die umbenannt wird
        :neuer_name: Neuer Name
        """
        self.regeln[index].name = neuerName
        self.notifyObserver()

    def removeRegel(self, index):
        """Loescht eine Regel.

        :index: Index, der geloscht wird.
        """
        del self.regeln[index]
        self.notifyObserver()

    def clearRegeln(self):
        """Loescht alle Regeln"""
        self.regeln = []
        self._aktiveRegel = None
        self.notifyObserver()

    def getBedingungsliste(self):
        """Gibt eine Liste von Falldaten zurueck, die Bedingungen erfuellen

        :returns: Pandas Dataframe
        """

        if not self.regeln:
            raise UIError("Keine Regeln definiert")
        if self._excelDaten.dataframe is None:
            raise UIError("Noch keine Daten vorhanden")
        datenListe = [regel.getErfuellt() for regel in self.regeln]
        datenListe = [l.drop_duplicates(subset='FallDatum') for l in datenListe]
        return pd.concat(datenListe)

    def saveToFile(self, filename):
        """Speichert die enthaltenen Regeln in ein File

        :filename: Filename
        """
        path = pathlib.Path(filename)
        head = {Regel.UND: 'UND', Regel.ODER: 'ODER', Regel.NICHT: 'NICHT'}
        regelDataFrames = []
        for regel in self.regeln:
            regelDict = regel.getDict()
            regelDF = pd.DataFrame(
                {head[i]: pd.Series(regelDict[i]) for i in head}
                )
            regelDF['Name'] = regel.name
            regelDataFrames.append(regelDF)
        regeln = pd.concat(regelDataFrames)
        columns = ['Name'] + list(head.values())
        regeln.to_excel(path, index=False, columns=columns)

    def setAktiv(self, index):
        """Setzt die momentan aktive Regel

        :index: Index der neuen aktiven Regel
        """
        if index is None:
            self._aktiveRegel = None
        if 0 <= index < len(self.regeln):
            self._aktiveRegel = self.regeln[index]
        self.notifyObserver()

    def getAktiv(self):
        """Gibt die momentan aktive Regel zurueck """
        return self._aktiveRegel

    def addLeistungToAktiverRegel(self, name, typ):
        if self._aktiveRegel:
            self._aktiveRegel.addLeistung(name, typ)
            self.notifyObserver()

    def removeLeistungenFromAktiverRegel(self, indices, typ):
        if self._aktiveRegel:
            self._aktiveRegel.removeLeistung(indices, typ)
            self.notifyObserver()

    def getErfuelltAktiveRegel(self):
        """Gibt die Anzahl der Pakete zurueck, die die aktive Regel erfuellen
        :returns: Anzahl Pakete

        """
        if self._aktiveRegel:
            return self._aktiveRegel.getAnzahlErfuellt()
        else:
            return '-'


class ExcelDaten(ObserverSubject):
    """Objekt, das die Excel Daten enthaelt"""

    def __init__(self):
        super().__init__()
        self._dataframe = None
        self._kategorien = []
        self._leistungen = None

    @property
    def dataframe(self):
        """Getter dataframe"""
        return self._dataframe

    @dataframe.setter
    def dataframe(self, daten):
        """Setter dataframe"""
        self._dataframe = daten
        self.calcUniqueLeistungen()
        self.notifyObserver()

    def addKategorie(self, kategorie):
        """Fuegt eine Kategorie hinzu"""
        if not kategorie in self._kategorien:
            self._kategorien.append(kategorie)
            self.notifyObserver()

    def calcUniqueLeistungen(self):
        """Berechnet eine Liste mit allen Leistungen im Excel"""
        leistungen = self._dataframe['Leistung']
        self._leistungen = leistungen.drop_duplicates()

    def getLeistungen(self, filterLeistung=None):
        """Gibt die Unique Leistungen zurueck

        :returns: Die Leistungen

        """
        if self._leistungen is None:
            return []
        if filterLeistung:
            ind = self._leistungen.str.contains(filterLeistung)
            return self._leistungen[ind].values
        return self._leistungen

    def getAnzahlFalldaten(self):
        """Gibt die Anzahl Falldaten zurueck

        :return: Anzahl Falldaten
        """

        if self._dataframe is not None:
            return self.dataframe.FallDatum.drop_duplicates().shape[0]
        return 0

    def checkItem(self, label):
        """Prueft, ob eine Leistung in den Daten vorhanden ist

        :label: Name der Leistung
        :returns: True, wenn die Bedingung vorhanden ist
        """
        if self._dataframe is None:
            return False
        return label in self._leistungen.values

    def clearKategorien(self):
        """Loescht alle Kategorien"""
        self._kategorien = []
        self.notifyObserver()

    def getKategorien(self):
        """Gibt die Kategorien als Liste zurueck
        :returns: Kategorien
        """
        return self._kategorien

    def removeKategorien(self, rows):
        """Loescht die Kategorien"""
        self._kategorien = [
            k for i, k in enumerate(self._kategorien)
            if i not in rows
            ]
        self.notifyObserver()
