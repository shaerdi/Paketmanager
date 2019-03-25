###############################################################################
# Benoetigte Module
###############################################################################
import numpy as np
import pandas as pd
import pathlib
import xlsxwriter
from collections import defaultdict

class UIError(Exception):
    pass

def convertLeistung(leistung):
    """Macht aus einer Zahl eine Buchstabenfolge (String)

    :returns: Den string im Format xx.xxxx
    """
    try:
        return '{:07.4f}'.format(float(leistung))
    except:
        return str(leistung)

def getKategorie(group, kategorien):
    """Sucht die Kategorie einer Gruppe
    """
    tarmedgroup = group[group.Leistungskategorie == 'Tarmed']
    leistungen = tarmedgroup.Leistung.values
    if not leistungen:
        return 'OhneTarmed'

    for k in kategorien:
        if k in leistungen:
            return k
    return 'Restgruppe'


def datenEinlesen(dateiname):
    """Liest ein Excel ein

    :returns: Ein pandas Objekt mit allen Daten im ersten Sheet des Excels und
    eine Liste mit den Kategorien aus dem zweiten Sheet des Excels

    """
    if not '.xls' in dateiname:
        raise IOError("Datei hat nicht die Endung '.xls' oder '.xlsx'")
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
        raise IOError(
                "Keine Kategorien gefunden. Gibt es ein zweites "
              + "Sheet in der Datei {}?".format(dateiname)
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

###############################################################################
# Hauptfunktion, geht alle Leistungen durch und schreibt sie in ein Excel
###############################################################################
def createPakete(daten, kategorien):
    """
    :daten: Pandas objekt mit allen Daten
    :kategorien: Liste mit den Kategorien
    """

    # Erster Durchgang:
    # Die Daten werden nach FallDatum gruppiert, und jede Gruppe wird in die
    # Pakete einsortiert.

    def buildKey(s):
        return ','.join(set(s))

    Leistungen = daten[daten.Leistungskategorie == 'Tarmed'][['FallDatum','Leistung']]
    keys = Leistungen.groupby('FallDatum').aggregate(buildKey)
    keys.rename({'Leistung': 'key'}, axis=1, inplace=True)

    daten = daten.join(keys, on='FallDatum')
    daten.fillna({'key':''}, inplace=True)

    # for fd,g in daten.groupby('FallDatum'): 
        # key = set(g[g.Leistungskategorie=='Tarmed'].Leistung)  
        # daten.loc[g.index,'key'] = ",".join(key)

    for i,(fd,g) in enumerate(daten.groupby('key')):
        daten.loc[g.index, 'paketID'] = int(i)
        daten.loc[g.index,'Anzahl'] = g.shape[0]


    def getKat(key):
        if len(key) == 0:
            return 'OhneTarmed'
        for k in kategorien:
            if k in key:
                return k
        return 'Restgruppe'

    daten['Kategorie'] = daten['key'].apply(getKat)

    return daten

def getFirstGroup(groups):
    for i,g in groups:
        return g

def writePaketeToExcel(daten, kategorien, filename):
    #############################################
    # Ab jetzt der Code um das Excel zu schreiben
    #############################################

    # Pfad zur Resultatdatei
    fname = pathlib.Path(filename).with_suffix('.xlsx')

    # Ordner erstellen, wenn es ihn noch nicht gibt
    if not fname.parent.exists():
        fname.parent.mkdir()
        
    # Excel erstellen
    writer = pd.ExcelWriter(str(fname), engine='xlsxwriter')
    workbook = writer.book

    ## Daten Schreiben
    # Rohdaten
    daten.drop(['Kategorie'],axis=1).to_excel(
            writer, sheet_name='Rohdaten', index=False
            )

    # Alle Pakete, jeweils die erste Fallnummer im entsprechenden Paket
    allePakete = pd.concat(sorted([
            getFirstGroup(g[1].groupby('FallDatum',sort=False))
            for g in daten.drop('Kategorie',axis=1).groupby('paketID',sort=False)
            ],
            key = lambda x : x.Anzahl.max(), reverse = True
            ))
    sheetSchreiben('AllePakete', allePakete, writer)

    # Pro Kategorie
    for kategorie in kategorien + ['Restgruppe','OhneTarmed']:
        katData = daten[daten['Kategorie']==kategorie].drop('Kategorie',axis=1)
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

    def __init__(self, name, daten, notifyFunc):
        self.name = name
        self._bedingungUnd = []
        self._bedingungOder = []
        self._bedingungNicht = []
        self.anzahl = '-'
        self._daten = daten
        self._erfuellt = None
        self._notifyFunc = notifyFunc

    def getDict(self):
        """Erstellt ein Dict aus den Bedingungen dieser Regel

        :returns: Dictionary mit den Eintraegen UND, ODER und NICHT

        """
        return {
            'UND' : self._bedingungUnd,
            'ODER' : self._bedingungOder,
            'NICHT' : self._bedingungNicht,
        }

    def addLeistung(self, newItem, typ):
        """Fuegt eine neue Leistung zur einer Liste hinzu

        :new_item: Neue Leistung
        :bedingungs_art: Regel.UND, ODER oder NICHT
        """

        if typ == Regel.UND:
            self._bedingungUnd.append(newItem)
        elif typ == Regel.ODER:
            self._bedingungOder.append(newItem)
        elif typ == Regel.NICHT:
            self._bedingungNicht.append(newItem)
        else:
            raise RuntimeError("Unbekannte Bedingung")
        self.update()
        self._notifyFunc()

    def removeLeistung(self, index, typ):
        """Loescht eine Leistung aus einer Liste

        :index: Index der zu loeschenden Leistung
        :bedingungs_art: Regel.UND, ODER oder NICHT
        """

        if typ == Regel.UND:
            del self._bedingungUnd[index]
        elif typ == Regel.ODER:
            del self._bedingungOder[index]
        elif typ == Regel.NICHT:
            del self._bedingungNicht[index]
        else:
            raise RuntimeError("Unbekannte Bedingung")
        self.update()
        self._notifyFunc()

    def clearItems(self, typ):
        """Loescht die Leistungen aus einer Liste

        :index: Index der zu loeschenden Leistung
        :bedingungs_art: Regel.UND, ODER oder NICHT
        """

        if typ == Regel.UND:
            self._bedingungUnd = []
        elif typ == Regel.ODER:
            self._bedingungOder = []
        elif typ == Regel.NICHT:
            self._bedingungNicht = []
        else:
            raise RuntimeError("Unbekannte Bedingung")
        self.update()
        self._notifyFunc()

    def update(self):
        """Berechnet die Pakete, die diese Regel erfuellen"""

        if self._daten.dataframe is None:
            self.anzahl = '-'
            return

        def erfuellt(key):
            """Checkt, ob ein Key diese Regel erfuellt"""
            erfuelltalle = all([(k in key) for k in self._bedingungUnd])
            erfuelltoder = len(self._bedingungOder) == 0 or \
                           any([(k in key) for k in self._bedingungOder])
            erfuelltnot = all([(k not in key) for k in self._bedingungNicht])
            return  erfuelltalle and erfuelltoder and erfuelltnot

        ind = self._daten.dataframe.key.apply(erfuellt)
        self._erfuellt = self._daten.dataframe[ind]
        self.anzahl = str(ind.sum)

    def getErfuellt(self):
        """Gibt ein Dataframe zurueck, das alle Falldaten enthaelt, die diese
        Regel erfuellen.

        :return: Pandas DataFrame
        """
        try:
            kopie = self._erfuellt.copy()
            kopie['Regel'] = self.name
            return kopie
        except AttributeError:
            spalten = self._daten.dataframe.columns
            spalten.append('Regel')
            return pd.DataFrame(columns=spalten)

    def getLeistungen(self, typ):
        """Gibt die Leistungen im Typ der Regel zurueck

        :typ: Regel.UND, Regel.ODER oder Regel.NICHT
        :returns: Liste mit Leistungen
        """
        if typ == Regel.UND:
            return self._bedingungUnd
        elif typ == Regel.ODER:
            return self._bedingungOder
        elif typ == Regel.NICHT:
            return self._bedingungNicht
        else:
            raise RuntimeError("Unbekannte Bedingung")

class Regeln(ObserverSubject):
    """Klasse, die die Regeln speichert"""

    def __init__(self, excelDaten):
        super().__init__()
        self.regeln = []
        self.aktiveRegel = None
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
        neueRegel = Regel(name, self._excelDaten, self.notifyObserver)
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
        self.aktiveRegel = None
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
        regelDict = {}
        for regel in self.regeln:
            regelDict[regel.name] = regel.getDict()

        with path.with_suffix('.tpf').open('wb') as f:
            pickle.dump(regelDict, f)

    def loadFromFile(self, filename):
        """Laedt die Regeln aus einem File

        :filename: Filename
        """
        path = pathlib.Path(filename)
        with path.with_suffix('.tpf').open('rb') as f:
            regelnDict = pickle.load(f)

        if (
                not isinstance(regelnDict, dict)
                or ({'UND', 'ODER', 'NICHT'} - regelnDict.keys())
        ):
            raise UIError("Fehler beim Laden der Regeln, ungültiges File")

        self.regeln = []
        for name, bedingungen in regelnDict.items():
            neueRegel = Regel(name, self._excelDaten, self.notifyObserver)
            for leistung in bedingungen['UND']:
                neueRegel.addLeistung(leistung, Regel.UND)
                print(leistung)
            for leistung in bedingungen['ODER']:
                neueRegel.addLeistung(leistung, Regel.ODER)
            for leistung in bedingungen['NICHT']:
                neueRegel.addLeistung(leistung, Regel.NICHT)
            self.regeln.append(neueRegel)
        self.notifyObserver()
        self.updateRegel()

    def setAktiv(self, index):
        """Setzt die momentan aktive Regel

        :index: Index der neuen aktiven Regel
        """
        if 0 <= index < len(self.regeln):
            self.aktiveRegel = self.regeln[index]
        self.notifyObserver()


class ExcelDaten(ObserverSubject):
    """Objekt, das die Excel Daten enthaelt"""

    def __init__(self):
        super().__init__()
        self._dataframe = None
        self._kategorien = set()
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
        self._kategorien.add(kategorie)

    def calcUniqueLeistungen(self):
        """Berechnet eine Liste mit allen Leistungen im Excel"""
        leistungen = self._dataframe[
            self._dataframe['Leistungskategorie'] == 'Tarmed'
            ]['Leistung']
        self._leistungen = leistungen.drop_duplicates()

    def getLeistungen(self, filterLeistung=None):
        """Gibt die Unique Leistungen zurueck

        :returns: Die Leistungen

        """
        if self._leistungen is None:
            return []
        ind = self._leistungen.str.contains(filterLeistung)
        return self._leistungen[ind].values

    def getAnzahlFalldaten(self):
        """Gibt die Anzahl Falldaten zurueck

        :return: Anzahl Falldaten
        """

        if self._dataframe is not None:
            return self.dataframe.FallDatum.drop_duplicates().shape[0]
        return 0

    def checkItem(self, label):
        """Prueft, ob eine Bedingung in den Daten vorhanden ist

        :label: Name der Bedingung
        :returns: True, wenn die Bedingung vorhanden ist
        """
        if self._dataframe is None:
            return False
        return label in self._leistungen.values

    def getKategorien(self):
        """Gibt die Kategorien als Liste zurueck
        :returns: Kategorien
        """
        return list(self._kategorien)
