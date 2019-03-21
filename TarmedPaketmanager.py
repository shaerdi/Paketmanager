"""GUI Modul des TarmedPaketmanagers"""

import pickle
import pathlib
import threading
import wx
import wx.lib.mixins.listctrl
from wx.lib.pubsub import pub

import pandas as pd
from ExcelCalc import datenEinlesen, createPakete, writePaketeToExcel


class ExcelReader(threading.Thread):
    """Thread, um ein Excel einzulesen"""

    def __init__(self, parent, fname):
        threading.Thread.__init__(self)
        self._parent = parent
        self._fname = fname
        self.start()

    def run(self):
        try:
            result = datenEinlesen(self._fname)
            if result is not None:
                daten, kategorien = result
                daten = createPakete(daten, kategorien)
                success = True
                result = (daten, kategorien)
            else:
                success = False

            wx.CallAfter(
                pub.sendMessage,
                'excel.read',
                success=success,
                data=result,
            )
        except Exception as error:
            errMsg = '{}'.format(error)
            wx.CallAfter(pub.sendMessage,
                'excel.read',
                success=False,
                msg=errMsg,
            )

class ExcelPaketWriter(threading.Thread):
    """Thread, um ein Excel zu speichern"""
    def __init__(self, parent, fname, daten, kategorien):
        threading.Thread.__init__(self)
        self._parent = parent
        self._fname = fname
        self._kategorien = kategorien
        self._daten = daten
        self.start()

    def run(self):
        try:
            writePaketeToExcel(self._daten, self._kategorien, self._fname)
            wx.CallAfter(pub.sendMessage, 'excel.write', success=True)
        except Exception as error:
            errMsg = '{}'.format(error)
            wx.CallAfter(pub.sendMessage,
                'excel.write',
                success=False,
                msg=errMsg,
            )


class ExcelDataFrameWriter(threading.Thread):
    """Thread, um ein Dataframe in ein Excel zu speichern"""
    def __init__(self, parent, fname, dataframe):
        threading.Thread.__init__(self)
        self._parent = parent
        self._fname = fname
        self._dataframe = dataframe
        self.start()

    def run(self):
        try:
            self._dataframe.to_excel(self._fname, index=False)
            wx.CallAfter(pub.sendMessage, 'excel.write', success=True)
        except Exception as error:
            errMsg = '{}'.format(error)
            wx.CallAfter(pub.sendMessage,
                'excel.write',
                success=False,
                msg=errMsg,
            )

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

LIST_TOOLTIPS = {
    'regel' : 'Definierte Regeln',
    'Kategorien' : 'Kategorien, in die die Pakete eingeordnet werden',
    Regel.UND : 'Alle Leistungen müssen im Paket vorkommen',
    Regel.ODER: 'Mindestens eine Leistung muss im Paket vorkommen',
    Regel.NICHT : 'Keine der Leistungen darf im Paket vorkommen',
}
LIST_HEADER = {
    'RegelListe' : 'Regelname',
    'KategorieListe' : 'Leistung',
    Regel.UND : 'Und-Bedingung',
    Regel.ODER: 'Oder-Bedingung',
    Regel.NICHT : 'Nicht-Bedingung',
}

class UIError(Exception):
    pass

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


class BedingungswahlDialog(wx.Dialog):
    def __init__(self, parent, id, title, daten):
        wx.Dialog.__init__(self, parent, id, title)
        self.daten = daten
        self.InitUI()
        self.SetList()

    def InitUI(self):
        vbox = wx.BoxSizer(wx.VERTICAL) 
        hbox1 = wx.BoxSizer(wx.HORIZONTAL)
        txt = wx.StaticText(self, label = "Neue Bedingung")
        hbox1.Add(txt,proportion=0,flag=wx.ALL,border=5)
        self.insertTxt = wx.TextCtrl(self)
        hbox1.Add(self.insertTxt,proportion=1,flag=wx.EXPAND)
        vbox.Add(hbox1, proportion=0, flag = wx.EXPAND|wx.TOP, border = 4) 

        self.helpList = wx.ListBox(self,
                style= wx.LB_SINGLE|wx.LB_NEEDED_SB,
                )
        vbox.Add(self.helpList, proportion=1, flag = wx.EXPAND|wx.ALL,border=5)

        sizer =  self.CreateButtonSizer(wx.OK|wx.CANCEL)
        vbox.Add(sizer, proportion=0, flag=wx.EXPAND|wx.ALL, border=5)
        self.SetSizer(vbox)
        
        self.Bind(wx.EVT_LISTBOX, self.OnListboxClicked, self.helpList)
        self.Bind(wx.EVT_LISTBOX_DCLICK, self.OnListboxDoubleClicked, self.helpList)
        self.Bind(wx.EVT_TEXT, self.OnTextChanged, self.insertTxt)

    def OnTextChanged(self,event):
        self.SetList()

    def OnListboxDoubleClicked(self,event):
        self.OnListboxClicked(event)
        self.EndModal(wx.ID_OK)

    def OnListboxClicked(self,event):
        string = event.GetEventObject().GetStringSelection()
        self.insertTxt.ChangeValue(string)
        
    def SetList(self):
        self.helpList.Clear()
        leistungen = self.daten.getLeistungen(self.insertTxt.GetValue())
        if not leistungen is None and not len(leistungen)==0:
            self.helpList.InsertItems(leistungen,0)

    def GetValue(self):
        return self.insertTxt.GetValue()


class SummaryPanel(wx.Panel):
    def __init__(self, *args, **kwargs):
        super().__init__(*args,**kwargs)
        self.InitUI()

    def InitUI(self):
        sizer = wx.FlexGridSizer(3, 2, 10,50)

        txt = wx.StaticText(self, label="Infos", style=wx.ALIGN_CENTRE_HORIZONTAL)
        sizer.Add(txt)
        sizer.Add(wx.StaticText(self))

        txt = wx.StaticText(self, label="Anzahl Falldaten:", style=wx.ALIGN_LEFT)
        sizer.Add(txt)

        self.anzahlPaketeTotal = wx.StaticText(self, label="0", style=wx.ALIGN_LEFT)
        sizer.Add(self.anzahlPaketeTotal)

        txt = wx.StaticText(self, label="Falldaten mit Bedingung:", style=wx.ALIGN_LEFT)
        sizer.Add(txt)

        self.anzahlPaketeBedingung = wx.StaticText(self, label="0", style=wx.ALIGN_LEFT)
        sizer.Add(self.anzahlPaketeBedingung)

        self.SetSizer(sizer)

    def updateTotal(self, num):
        self.anzahlPaketeTotal.SetLabel( str(num) )

    def updateBedingung(self, num):
        self.anzahlPaketeBedingung.SetLabel( str(num) )


class AnzeigeListe(wx.ListCtrl, wx.lib.mixins.listctrl.ListCtrlAutoWidthMixin):

    def __init__(self, parent, regeln, daten, *args, **kw):
        self._parent = parent
        self.regeln = regeln
        self.daten = daten
        titel = kw.pop('titel','')

        if 'style' not in kw:
            kw['style'] = wx.LC_REPORT|wx.LC_HRULES|wx.LC_VIRTUAL

        wx.ListCtrl.__init__(self, parent, *args, **kw)
        wx.lib.mixins.listctrl.ListCtrlAutoWidthMixin.__init__(self)

        self.InsertColumn(0, titel)

    def OnGetItemText(self, item, col):
        return self.items[item]


class RegelListe(AnzeigeListe):
    def __init__(self, parent, regeln, daten, *args, **kw):
        style = wx.LC_REPORT|wx.LC_HRULES|wx.LC_VIRTUAL|wx.LC_SINGLE_SEL
        AnzeigeListe.__init__(self, parent, regeln, daten, *args, style=style, **kw)
        self.Bind(wx.EVT_LIST_ITEM_ACTIVATED, self.onDoubleClick)
        self.Bind(wx.EVT_LIST_END_LABEL_EDIT, self.labelEdit)
        self.Bind(wx.EVT_LIST_ITEM_SELECTED, self.onSetFocus)
        regeln.registerObserver(self)
        self.update()

    def onSetFocus(self, event):
        index = event.GetIndex()
        self.regeln.setAktiv(index)

    def labelEdit(self, event):
        """Methode, die nach dem Editieren eines Labels aufgerufen wird"""
        newLabel = event.GetLabel()
        oldLabel = self.items[event.GetIndex()]
        self.regeln.rename_regel(oldLabel, newLabel)
        self.update()

    def onDoubleClick(self,event):
        """Methode, die bei einem Doppelklick aufgerufen wird"""
        self.EditLabel(event.GetIndex())

    def update(self):
        """Liest die Items neu ein"""
        if self.regeln is not None:
            self.items = [r.name for r in self.regeln.regeln]
            self.SetItemCount(len(self.regeln.regeln))

    def deleteSelected(self):
        """Loescht die selektierten Items """
        index = self.GetFirstSelected()
        self.regeln.removeRegel(index)


class BedingungsListe(AnzeigeListe):
    def __init__(self, parent, regeln, daten, *args, **kw):

        self._typ = kw.pop('typ')
        kw['titel'] = LIST_HEADER[self._typ]

        AnzeigeListe.__init__(self, parent, regeln, daten, *args, **kw)

        self.normalItem = wx.ListItemAttr()
        self.redItem = wx.ListItemAttr()
        self.redItem.SetBackgroundColour(wx.Colour(255,204,204))

        regeln.registerObserver(self)
        self.update()

    def update(self):
        """Setzt Listenitems neu
        """
        aktiveRegel = self.regeln.aktiveRegel
        if aktiveRegel is not None:
            self.items = aktiveRegel.getLeistungen(self._typ)
            self.SetItemCount(len(self.items))
        else:
            self.items = []
            self.SetItemCount(0)

    def OnGetItemAttr(self, item):
        """Prueft, ob ein Item in den Daten vorhanden ist

        :item: Index des zu pruefenden Items
        """
        if self.daten.checkItem(self.items[item]):
            return self.normalItem
        else:
            return self.redItem

    def deleteSelected(self):
        """Loescht die selektierten Items """
        index = self.GetFirstSelected()
        itemsToDelete = []
        while index >= 0:
            itemsToDelete.append(index)
            index = self.GetNextSelected(index)

        aktiveRegel = self.regeln.aktiveRegel
        if aktiveRegel is not None:
            for item in itemsToDelete:
                aktiveRegel.removeLeistung(index, self._typ)


    def newItem(self, event):
        with BedingungswahlDialog(self, wx.ID_ANY, "Neue Bedingung", self.daten) as dlg:
            if dlg.ShowModal() == wx.ID_OK:
                text = dlg.GetValue()
                aktiveRegel = self.regeln.aktiveRegel
                if text != '' and aktiveRegel is not None:
                    aktiveRegel.addLeistung(text, self.typ)


    def clrItem(self, event):
        aktiveRegel = self.regeln.aktiveRegel
        if aktiveRegel is not None:
            aktiveRegel.clearItems(self.typ)


class TarmedpaketGUI(wx.Frame):
    name = "Tarmed Pakete"
    windowSize = (1300, 800)

    def __init__(self, parent):
        super().__init__(parent, 
                title=self.name,
                size=self.windowSize,
                )

        self._currentWorker = None

        self.daten = ExcelDaten()
        self.regeln = Regeln(self.daten)

        self.initUI()
        self.setupEvents()
        self.Centre()

    def setupEvents(self):
        self.Bind(wx.EVT_CLOSE, self.onCloseFrame)
        self.Bind(wx.EVT_MENU, self.onCloseFrame, self.fileMenuExitItem)
        self.Bind(wx.EVT_MENU, self.onSaveRule, self.fileMenuSaveRule)
        self.Bind(wx.EVT_MENU, self.onLoadRule, self.fileMenuLoadRule)
        self.Bind(wx.EVT_MENU, self.onSaveExcel, self.fileMenuExportExcel)
        self.Bind(wx.EVT_MENU, self.onSaveRegelExcel, self.fileMenuExportRegelExcel)
        self.Bind(wx.EVT_BUTTON, self.openExcel, self.excelOpenButton)

        pub.subscribe(self.finishWrite, 'excel.write')
        pub.subscribe(self.onFinishExcelCalc, 'excel.read')

    def finishWrite(self, success, msg=None):
        """Funktion, die nach dem Schreiben eines Excels aufgerufen wird

        :success: Bool ob erfolgreich
        :errMsg: Fehlermeldung
        """
        if not success and msg is not None:
            wx.MessageBox(
                msg,
                'Fehler',
                wx.OK | wx.ICON_ERROR,
            )

        self._currentWorker = None
        self.enableWindow()

    def onSaveRegelExcel(self,event):
        saveFileDialog = wx.FileDialog(
            self,
            "Speichern unter", 
            "", 
            "", 
            "Excel files (*.xlsx; *.xls)|*.xlsx;*.xls", 
            wx.FD_SAVE,
        )
        saveFileDialog.ShowModal()
        filePath = pathlib.Path(saveFileDialog.GetPath())
        saveFileDialog.Destroy()
        try:
            dataframe = self.regeln.getBedingungsliste()
            self.disableWindow()
            self._currentWorker = ExcelDataFrameWriter(self, filePath, dataframe)
        except UIError as error:
            wx.MessageBox(
                "{}".format(error),
                'Fehler',
                wx.OK | wx.ICON_ERROR,
            )


    def onFinishExcelCalc(self, success, data=None, msg=None):
        """ Funktion, die nach dem Lesen eines Excels aufgerufen wird

        :data: Tuple mit Dataframe und Kategorienliste
        :success: Bool ob erfolgreich
        :errMsg: Fehlermeldung:
        """
        if success:
            self.daten.dataframe = data[0]
            kategorien = data[1]
            if kategorien is not None:
                for kategorie in kategorien:
                    self.daten.addKategorie(kategorie)
            # TODO
            # self.summaryPanel.updateTotal( self.regeln.get_anzahl_falldaten() )
            # self.daten.updateSummaryPanel()
        else:
            if errMsg:
                wx.MessageBox(
                    message=errMsg,
                    caption='Fehler',
                    style=wx.OK | wx.ICON_INFORMATION,
                )

        self._currentWorker = None
        self.enableWindow()

    def onExitApp(self, event):
        self.Destroy()

    def onSaveRule(self, event):
        saveFileDialog = wx.FileDialog(
            self, 
            "Speichern unter", "", "", 
            "TarmedPaketGUI files (*.tpf)|*.tpf", 
            wx.FD_SAVE,
           )
        saveFileDialog.ShowModal()
        file_path = pathlib.Path(saveFileDialog.GetPath())
        saveFileDialog.Destroy()
        self.regeln.saveToFile(file_path)

    def onSaveExcel(self, event):
        saveFileDialog = wx.FileDialog(
            self, 
            "Speichern unter", "", "",
            "Excel files (*.xlsx; *.xls)|*.xlsx;*.xls",
            wx.FD_SAVE,
           )
        saveFileDialog.ShowModal()
        filePath = pathlib.Path(saveFileDialog.GetPath())
        saveFileDialog.Destroy()

        if self.daten.dataframe is None:
            wx.MessageBox(
                'Noch keine Daten vorhanden',
                'Info',
                wx.OK | wx.ICON_INFORMATION,
            )
            return

        if filePath:
            self.disableWindow()
            self._currentWorker = ExcelPaketWriter(
                self,
                filePath,
                self.daten.dataframe,
                self.daten.getKategorien(),
            )


    def onLoadRule(self, event):
        openFileDialog = wx.FileDialog(self, "Öffnen", "", "", 
                                      "TarmedPaketGUI files (*.tpf)|*.tpf", 
                                       wx.FD_OPEN | wx.FD_FILE_MUST_EXIST,
                                       )
        openFileDialog.ShowModal()
        filePath = pathlib.Path(openFileDialog.GetPath())
        openFileDialog.Destroy()
        try:
            self.regeln.loadFromFile(filePath)
        except UIError as error:
            wx.MessageBox(
                message='{}'.format(error),
                caption='Fehler',
                style=wx.OK | wx.ICON_ERROR,
            )

    def onCloseFrame(self, event):
        dialog = wx.MessageDialog(self, message="Programm wirklich Schliessen?", caption="", style=wx.YES_NO, pos=wx.DefaultPosition)
        response = dialog.ShowModal()

        if response == wx.ID_YES:
            self.onExitApp(event)
        else:
            event.StopPropagation()

    def initMenuBar(self):
        menubar = wx.MenuBar()

        fileMenu = wx.Menu()
        self.fileMenuExportExcel = wx.MenuItem(
            fileMenu,
            wx.ID_ANY,
            text="Export Excel",
        )
        fileMenu.Append(self.fileMenuExportExcel)

        self.fileMenuExportRegelExcel = wx.MenuItem(
            fileMenu,
            wx.ID_ANY,
            text="Export Bedingungen in Excel",
        )
        fileMenu.Append(self.fileMenuExportRegelExcel)

        fileMenu.AppendSeparator()

        self.fileMenuSaveRule = wx.MenuItem(
            fileMenu,
            wx.ID_SAVE,
            text="Regeln speichern",
        )
        self.fileMenuLoadRule = wx.MenuItem(
            fileMenu,
            wx.ID_OPEN,
            text="Regeln laden",
        )

        fileMenu.Append(self.fileMenuSaveRule)
        fileMenu.Append(self.fileMenuLoadRule)

        fileMenu.AppendSeparator()

        self.fileMenuExitItem = wx.MenuItem(fileMenu, wx.ID_EXIT, '&Quit\tCtrl+Q')
        fileMenu.Append(self.fileMenuExitItem)

        menubar.Append(fileMenu, '&Datei')
        self.SetMenuBar(menubar)



    def initUI(self):
        """Initialisert die UI Elemente"""
        self.initMenuBar()

        panel = wx.Panel(self)
        self.panel = panel

        mainBox = wx.BoxSizer(wx.VERTICAL)

        hBox1 = wx.BoxSizer(wx.HORIZONTAL)
        text = wx.StaticText(panel, label='Aktuelles Excel')
        hBox1.Add(text, flag=wx.RIGHT|wx.TOP, border=5)

        self.excelPath = wx.TextCtrl(panel)
        hBox1.Add(self.excelPath, proportion=1, flag=wx.LEFT, border=10)

        self.excelOpenButton = wx.Button(panel, label='Öffnen')
        hBox1.Add(self.excelOpenButton, flag=wx.LEFT|wx.RIGHT, border=15)

        mainBox.Add(hBox1, flag=wx.TOP|wx.LEFT|wx.RIGHT|wx.EXPAND, border=15)

        # -----------------------------------
        line = wx.StaticLine(panel)
        mainBox.Add(line, flag=wx.EXPAND|wx.BOTTOM|wx.TOP, border=15)

        hBox2 = wx.BoxSizer(wx.HORIZONTAL)

        ruleBox = wx.StaticBox(panel, label="Infos")
        ruleBoxSizer = wx.StaticBoxSizer(ruleBox, wx.HORIZONTAL)

        summaryPanel = SummaryPanel(panel)
        ruleBoxSizer.Add(summaryPanel, flag=wx.EXPAND|wx.ALL, border=15)

        hBox2.Add(ruleBoxSizer, proportion = 1, flag=wx.EXPAND|wx.ALL, border=15)

        ruleBox = wx.StaticBox(panel, label="Kategorien")
        ruleBoxSizer = wx.StaticBoxSizer(ruleBox, wx.HORIZONTAL)

        kategorieListe = RegelListe(
            panel,
            regeln=self.regeln,
            daten=self.daten,
            size=(150, -1),
            titel=LIST_HEADER['KategorieListe'],
        )
        ruleBoxSizer.Add(kategorieListe, flag=wx.EXPAND|wx.ALL, border=10)

        hBox2.Add(ruleBoxSizer, proportion=0, flag=wx.EXPAND|wx.ALL, border=15)

        ruleBox = wx.StaticBox(panel, label="Regeln")
        ruleBoxSizer = wx.StaticBoxSizer(ruleBox, wx.HORIZONTAL)


        regelPanel = RegelListe(
            panel,
            regeln=self.regeln,
            daten=self.daten,
            size=(150, -1),
            titel=LIST_HEADER['RegelListe'],
        )
            
        bPanel1 = BedingungsListe(
            panel,
            regeln=self.regeln,
            daten=self.daten,
            size=(150, -1),
            typ=Regel.UND,
        )
        bPanel2 = BedingungsListe(
            panel,
            regeln=self.regeln,
            daten=self.daten,
            size=(150, -1),
            typ=Regel.ODER,
        )
        bPanel3 = BedingungsListe(
            panel,
            regeln=self.regeln,
            daten=self.daten,
            size=(150, -1),
            typ=Regel.UND,
        )

        flags = wx.EXPAND|wx.BOTTOM|wx.TOP|wx.LEFT
        border = 10
        ruleBoxSizer.Add(regelPanel, proportion=0, flag=flags|wx.RIGHT, border=border)
        ruleBoxSizer.Add(bPanel1, proportion=0, flag=flags, border=border)
        ruleBoxSizer.Add(bPanel2, proportion=0, flag=flags, border=border)
        ruleBoxSizer.Add(bPanel3, proportion=0, flag=flags|wx.RIGHT, border=border)

        hBox2.Add(ruleBoxSizer, proportion=0, flag=wx.EXPAND|wx.ALL, border=15)

        mainBox.Add(hBox2, proportion=1, flag=wx.LEFT|wx.RIGHT|wx.EXPAND, border=15)

        panel.SetSizer(mainBox)

    def openExcel(self, *_):
        """Oeffnet ein Excel mit Falldaten"""
        openFileDialog = wx.FileDialog(
            self,
            "Wählen",
            "",
            "Excel files (*.xlsx;*.xls)|*.xlsx;*.xls",
            "",
            wx.FD_OPEN | wx.FD_FILE_MUST_EXIST,
        )
        if not openFileDialog.ShowModal() == wx.ID_OK:
            return
        filePath = openFileDialog.GetPath()
        openFileDialog.Destroy()
        self.disableWindow()
        self.excelPath.SetValue(filePath)
        self._currentWorker = ExcelReader(self, filePath)

    def disableWindow(self):
        """Schaltet das Fenster in den Wartemodus"""
        self.SetCursor(wx.Cursor(wx.CURSOR_WAIT))
        self.Disable()

    def enableWindow(self):
        """Schaltet den Wartemodus aus"""
        self.SetCursor(wx.Cursor(wx.CURSOR_ARROW))
        self.Enable()

    def menuhandler(self, event):
        """Funktion, die ein MenuEvent handled"""
        eventID = event.GetId()
        if eventID == wx.ID_EXIT:
            self.Close()


def main():
    """Main GUI Loop"""
    app = wx.App()
    ex = TarmedpaketGUI(None)
    ex.Show()
    app.MainLoop()


if __name__ == '__main__':
    main()
