"""GUI Modul des TarmedPaketmanagers"""

import pickle
import pathlib
import threading
import wx
import wx.lib.mixins.listctrl

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
        evt = ResultEvent(success=False)
        if self._fname is not None:
            try:
                result = datenEinlesen(self._fname)
                if result is not None:
                    daten, kategorien = result
                    daten = createPakete(daten, kategorien)
                    evt.success = True
                    evt.data = (daten, kategorien)
            except IOError as error:
                evt.errMsg = '{}'.format(error)

        wx.PostEvent(self._parent, evt)

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
        evt = ResultEvent(success=False)

        if self._fname is not None:
            try:
                writePaketeToExcel(self._daten, self._kategorien, self._fname)
                evt.success = True
            except IOError as error:
                evt.errMsg = '{}'.format(error)

        wx.PostEvent(self._parent, evt)


EVT_RESULT_ID = 1001

TOOLTIPS = {
    'regel' : 'Definierte Regeln',
    'and' : 'Alle Leistungen müssen im Paket vorkommen',
    'or' : 'Mindestens eine Leistung muss im Paket vorkommen',
    'not' : 'Keine der Leistungen darf im Paket vorkommen',
}


def EVT_RESULT(win, func):
    win.Connect(-1, -1, EVT_RESULT_ID, func)


class ResultEvent(wx.PyEvent):
    """Event, der von einem Thread zurueck gegeben wird."""
    def __init__(self,
                 data=None,
                 success=True,
                 errMsg='',
                ):
        wx.PyEvent.__init__(self)
        self.SetEventType(EVT_RESULT_ID)
        self.data = data
        self.success = success
        self.errMsg = errMsg


class Regel(object):
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
        self._notifyFunc()

    def update(self):
        """Berechnet die Pakete, die diese Regel erfuellen"""

        if self._daten is None:
            self.anzahl = '-'
            return

        def erfuellt(key):
            """Checkt, ob ein Key diese Regel erfuellt"""
            erfuelltalle = all([(k in key) for k in self._bedingungUnd])
            erfuelltoder = len(self._bedingungOder) == 0 or \
                           any([(k in key) for k in self._bedingungOder])
            erfuelltnot = all([(k not in key) for k in self._bedingungNicht])
            return  erfuelltalle and erfuelltoder and erfuelltnot

        ind = self._daten.key.apply(erfuellt)
        self._erfuellt = self._daten[ind]
        self.anzahl = str(ind.sum)

    def getErfuellt(self):
        """Gibt ein Dataframe zurueck, das alle Falldaten enthaelt, die diese
        Regel erfuellen.

        :return: Pandas DataFrame
        """
        return self._erfuellt.copy()

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


class Regeln(object):
    """Klasse, die die Regeln speichert"""

    def __init__(self):
        self.observers = []
        self.regeln = []
        self.aktiveRegel = None
        self._excelDaten = None

    def registerObserver(self, observer):
        """Registriert ein Observerobjekt, das per Aufrufen der Funktion update
        auf Aenderungen aufmerksam gemacht wird.

        :observer: Observer objekt. Muss die Funktion update haben

        """
        observer.regeln = self
        self.observers.append(observer)

    def notifyObserver(self):
        """Ruft die Methode update fuer alle Observer auf

        """
        for observer in self.observers:
            observer.update()

    @property
    def excelDaten(self):
        """Getter ecxel_daten"""
        return self._excelDaten

    @excelDaten.setter
    def excelDaten(self, daten):
        """Setter ecxel_daten"""
        self._excelDaten = daten
        self.updateRegel()

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

    def getBedingungsliste(self):
        """Gibt eine Liste von Falldaten zurueck, die Bedingungen erfuellen

        :returns: Pandas Dataframe
        """

        datenListe = [regel.get_erfuellt() for regel in self.regeln]
        datenListe = [l.drop_duplicates(subset='FallDatum') for l in datenListe]
        return pd.concat(datenListe)

    def saveToFile(self, filename):
        """Speichert die enthaltenen Regeln in ein File

        :filename: Filename
        """
        path = pathlib.Path(filename)
        with path.with_suffix('.tpf').open('wb') as f:
            pickle.dump(self.regeln, f)

    def loadFromFile(self, filename):
        """Laedt die Regeln aus einem File

        :filename: Filename
        """
        path = pathlib.Path(filename)
        with path.with_suffix('.tpf').open('rb') as f:
            self.regeln = pickle.load(f)
        self.updateRegel()

    def setAktiv(self, index):
        """Setzt die momentan aktive Regel

        :index: Index der neuen aktiven Regel
        """
        if 0 <= index < len(self.regeln):
            self.aktiveRegel = self.regeln[index]


class ExcelDaten(object):
    """Objekt, das die Excel Daten enthaelt"""

    def __init__(self):
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
        return label in self._leistungen







# class DatenStruktur:
    # Listen = []
    # kategorien = None
    # daten = None

    # def __init__(self):
        # self.regeln = OrderedDict()
        # self.aktiv = ''

    # def saveRegelnToExcel(self, filePath):
        # if self.daten is None:
            # return
        # datenListe = [self.applyRegelToData(regel) for regel in self.regeln]
        # datenListe = [l.drop_duplicates(subset='FallDatum') for l in datenListe]
        # daten = pd.concat(datenListe)
        # print('hallo')
        # daten.to_excel(filePath, index=False)
        # print('test')

    # def applyRegelToData(self, regel=None):
        # if self.daten is None:
            # return
        # if regel is None:
            # regel = self.aktiv
        # aktiveRegel = self.regeln[regel or self.aktiv]
        # def erfuellt(key):
            # erfuelltAlle = all([ (    k in key) for k in aktiveRegel['and']])
            # erfuelltOder = len(aktiveRegel['or']) == 0 or \
                           # any([ (    k in key) for k in aktiveRegel['or']])
            # erfuelltNot  = all([ (not k in key) for k in aktiveRegel['not']])
            # return  erfuelltAlle and erfuelltOder and erfuelltNot

        # ind = self.daten.key.apply(erfuellt)
        # kopie = self.daten[ind].copy()
        # kopie.drop_duplicates(subset='FallDatum',inplace=True)
        # kopie['Regel'] = regel
        # return kopie

    # def getAnzahlFalldaten(self):
        # if not self.daten is None:
            # return self.daten.FallDatum.drop_duplicates().shape[0]
        # else:
            # return 0

    # def writeDatenToExcel(self,filePath):
        # if not self.daten is None:
            # writePaketeToExcel(self.daten, self.kategorien, filePath)
            # return True
        # else:
            # return False

    # def saveToFile(self,path):
        # with path.with_suffix('.tpf').open('wb') as f:
            # pickle.dump(self.regeln, f)

    # def renameRegel(self, from_, to_):
        # self.regeln = OrderedDict(
                # (to_ if k == from_ else k, v) 
                # for k, v in self.regeln.items()
                # )

    # def openFromFile(self,path):
        # with path.with_suffix('.tpf').open('rb') as f:
            # self.regeln = pickle.load(f)
        # self.updateListen()

    # def setAktiv(self,name):
        # self.aktiv=name

    # def CheckItem(self, item):
        # if self.daten is None:
            # return False
        # else:
            # return item in self.daten.Leistung.values

    # def getLeistungen(self, filter_ = ''):
        # if self.daten is None:
            # return
        # leistungen = self.daten[self.daten['Leistungskategorie'] == 'Tarmed']['Leistung']
        # leistungen = leistungen.drop_duplicates()
        # ind = leistungen.str.contains(filter_)
        # return leistungen[ind].values

    # def getRegeln(self):
        # return list(self.regeln.keys())

    # def deleteRegel(self,name):
        # self.regeln.pop(name)

    # def clearRegeln(self):
        # self.regeln = OrderedDict()

    # def addRegel(self,name):
        # if name in self.regeln:
            # return False
        # else:
            # self.regeln[name] = { 'and' : [], 'or' : [], 'not' : [] }
            # return True

    # def deleteItem(self,titel,item):
        # if not titel in ['and','or','not']:
            # return
        # self.regeln[self.aktiv][titel].remove(item)

    # def clearItems(self, titel):
        # if not titel in ['and','or','not']:
            # return
        # self.regeln[self.aktiv][titel] = []


    # def addItem(self,titel,item):
        # if not titel in ['and','or','not']:
            # return
        # self.regeln[self.aktiv].append(item)

    # def getAktiveRegel(self,titel):
        # if titel in ['and','or','not'] and self.aktiv in self.regeln:
            # return self.regeln[self.aktiv][titel]
        # else:
            # return []

    # def updateListen(self):
        # for l in self.Listen:
            # l.update()
        # self.updateSummaryPanel()

    # def updateSummaryPanel(self):
        # try:
            # self.summaryPanel.updateBedingung( self.applyRegelToData().shape[0] )
        # except:
            # pass


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

        if not 'style' in kw:
            kw['style'] = wx.LC_REPORT|wx.LC_NO_HEADER|wx.LC_HRULES|wx.LC_VIRTUAL

        wx.ListCtrl.__init__(self, parent, **kw)
        wx.lib.mixins.listctrl.ListCtrlAutoWidthMixin.__init__(self)

        self.InsertColumn(0, '')

    def OnGetItemText(self, item, col):
        return self.items[item]


class RegelListe(AnzeigeListe):
    def __init__(self, parent, regeln, daten, *args, **kw):
        style = wx.LC_REPORT|wx.LC_NO_HEADER|wx.LC_HRULES|wx.LC_VIRTUAL|wx.LC_SINGLE_SEL
        AnzeigeListe.__init__(self, parent, regeln, daten, *args, style=style, **kw)
        self.Bind(wx.EVT_LIST_ITEM_ACTIVATED, self.OnDoubleClick)
        self.Bind(wx.EVT_LIST_END_LABEL_EDIT, self.LabelEdit)
        self.Bind(wx.EVT_LIST_ITEM_SELECTED, self.onSetFocus)
        regeln.registerObserver(self)
        self.update()

    def onSetFocus(self, event):
        index = event.GetIndex()
        self.regeln.setAktiv(index)

    def LabelEdit(self,event):
        """Methode, die nach dem Editieren eines Labels aufgerufen wird"""
        newLabel = event.GetLabel()
        oldLabel = self.items[event.GetIndex()]
        self.regeln.rename_regel(oldLabel, newLabel)
        self.update()

    def OnDoubleClick(self,event):
        """Methode, die bei einem Doppelklick aufgerufen wird"""
        self.EditLabel(event.GetIndex())

    def update(self):
        """Liest die Items neu ein"""
        if self.regeln is not None:
            self.items = [r.name for r in self.regeln.regeln]
            self.SetItemCount(len(self.regeln.regeln))

    def deleteSelection(self):
        """Loescht die aktuell selektierten Items
        """
        pass


class BedingungsListe(AnzeigeListe):
    def __init__(self, parent, regeln, daten, *args, **kw):
        AnzeigeListe.__init__(self, parent, regeln, daten, *args, **kw)

        self.normalItem = wx.ListItemAttr()
        self.redItem = wx.ListItemAttr()
        self.redItem.SetBackgroundColour(wx.Colour(255,204,204))

        regeln.registerObserver(self)
        self.update()
        self._typ = None

    def setType(self, typ):
        """Setzt den Typ der Liste

        :typ: Regel.UND, Regel.ODER oder Regel.NICHT

        """
        self._typ = typ

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


class ListePanel(wx.Panel):
    """Abstrakte Basisklasse fuer ein Panel mit einer Liste"""
    def __init__(self, *args, **kwargs):

        titel = kwargs.pop('titel', '')
        self.regeln = kwargs.pop('regeln', {})
        self.daten = kwargs.pop('daten', {})

        super(ListePanel, self).__init__(*args, **kwargs)

        self.initUI(titel)
        self.setupEvents()

    def setupEvents(self):
        """Funktion, die alle wx Events bindet"""
        raise NotImplementedError()

    def getCtrlList(self):
        """Funktion, die eine Instanz der ListCtrl Klasse zurueckgibt"""
        raise NotImplementedError()

    def initUI(self, titel):
        """Setup des UI"""
        sizer = wx.GridBagSizer(5, 5)

        txt = wx.StaticText(self, label=titel, style=wx.ALIGN_CENTRE_HORIZONTAL)
        txt.SetToolTip(wx.ToolTip(TOOLTIPS[titel.lower()]))
        sizer.Add(txt, pos=(0,0), span=(1,4), flag=wx.EXPAND, border=15)

        self.listbox = self.getCtrlList()
        sizer.Add(self.listbox, pos=(1,0), span=(1,4), flag=wx.EXPAND|wx.BOTTOM, border=15)

        def create_button(symbol):
            btn = wx.Button(self, label=symbol, size=(50,30))
            font = wx.Font(15, wx.DEFAULT, wx.NORMAL, wx.BOLD)
            btn.SetFont(font)
            return btn

        self.newBtn = create_button('+')
        self.delBtn = create_button('-')
        self.clrBtn = create_button('X')

        sizer.Add(self.newBtn, pos=(2,0))
        sizer.Add(self.delBtn, pos=(2,1))
        sizer.Add(self.clrBtn, pos=(2,3))

        sizer.AddGrowableRow(1)
        sizer.AddGrowableCol(2)

        self.SetSizer(sizer)


class RegelPanel(ListePanel):
    def __init__(self, *args, **kwargs):
        super(RegelPanel, self).__init__(*args, **kwargs)
        self.setFocus()

    def setFocus(self, index = 0):
        listLen = len(self.listbox.items)
        if listLen == 0: return
        index = max(min( listLen-1, index), 0)
        self.listbox.Select(index)
        self.listbox.Focus(index)

    def getCtrlList(self):
        return RegelListe(self, self.regeln, self.daten, size=(70,-1))

    def setupEvents(self):
        self.Bind(wx.EVT_BUTTON, self.newItem, id=self.newBtn.GetId())
        self.Bind(wx.EVT_BUTTON, self.delItem, id=self.delBtn.GetId())
        self.Bind(wx.EVT_BUTTON, self.clrItem, id=self.clrBtn.GetId())

    def newItem(self, event):
        text = wx.GetTextFromUser('Enter a new item', 'Insert dialog')
        if text != '':
            self.regeln.addRegel(text)
            self.listbox.Select(len(self.regeln.regeln)-1)
            self.listbox.Focus(len(self.regeln.regeln)-1)

    def delItem(self, event):
        index = self.listbox.GetFirstSelected()
        if index >= 0:
            item = self.listbox.GetItem(index).GetText()
            self.daten.deleteRegel(item)
            self.daten.updateListen()
            self.setFocus(index)

    def clrItem(self, event):
        self.daten.clearRegeln()
        self.updateAktiv()


class BedingungsPanel(ListePanel):
    def __init__(self, *args, **kwargs):
        self.titel = kwargs.get('titel', '').lower()
        self.typ = {
            "and" : Regel.UND,
            "or" : Regel.ODER,
            "not" : Regel.NICHT,
           }[self.titel]

        super().__init__(*args,**kwargs)

    def getCtrlList(self):
        liste = BedingungsListe(self, self.regeln, self.daten, size=(100, -1))
        liste.setType(self.typ)
        return liste

    def setupEvents(self):
        self.Bind(wx.EVT_BUTTON, self.newItem, id=self.newBtn.GetId())
        self.Bind(wx.EVT_BUTTON, self.delItem, id=self.delBtn.GetId())
        self.Bind(wx.EVT_BUTTON, self.clrItem, id=self.clrBtn.GetId())
        self.Bind(wx.EVT_KEY_DOWN, self.onKeyPress)

    def onKeyPress(self, event):
        keycode = event.GetKeyCode()

        index = self.listbox.GetFocusedItem()
        if index < 0:
            return

        if keycode == wx.WXK_UP and index > 0:
            self.listbox.Focus(index-1)
        elif keycode == wx.WXK_DOWN and index < self.listbox.GetItemCount()-1:
            self.listbox.Focus(index+1)
        elif keycode == wx.WXK_DELETE or keycode == wx.WXK_NUMPAD_DELETE:
            self.DelItem(event)

    def newItem(self, event):
        with BedingungswahlDialog(self, wx.ID_ANY, "Neue Bedingung", self.daten) as dlg:
            if dlg.ShowModal() == wx.ID_OK:
                text = dlg.GetValue()
                aktiveRegel = self.regeln.aktiveRegel
                if text != '' and aktiveRegel is not None:
                    aktiveRegel.addLeistung(text, self.typ)

    def delItem(self, event):
        self.listbox.deleteSelected()

    def clrItem(self, event):
        aktiveRegel = self.regeln.aktiveRegel
        if aktiveRegel is not None:
            aktiveRegel.clearItems(self.typ)


class TarmedpaketGUI(wx.Frame):
    name = "Tarmed Pakete"
    windowSize = (1300, 800)

    panels = {}

    def __init__(self, parent):
        super().__init__(parent, 
                title=self.name,
                size=self.windowSize,
                )

        self.daten = ExcelDaten()
        self.regeln = Regeln()

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

        EVT_RESULT(self, self.finishExcelCalc)

    def onSaveRegelExcel(self,event):
        saveFileDialog = wx.FileDialog(
                self,
                "Speichern unter", "", "", 
                "Excel files (*.xlsx; *.xls)|*.xlsx;*.xls", 
                wx.FD_SAVE,
               )
        saveFileDialog.ShowModal()
        filePath = pathlib.Path(saveFileDialog.GetPath())
        saveFileDialog.Destroy()
        # TODO:
        # self.regeln.save_to_file(filePath)

    def finishExcelCalc(self, event):
        if event.success:
            self.regeln.excelDaten = event.data[0]
            # TODO
            # self.daten.kategorien = event.data[1]
            # self.summaryPanel.updateTotal( self.regeln.get_anzahl_falldaten() )
            # self.daten.updateSummaryPanel()
        else:
            if event.errMsg:
                wx.MessageBox(
                    message=event.errMsg,
                    caption='Fehler',
                    style=wx.OK | wx.ICON_INFORMATION,
                   )
            self.regeln.excelDaten = None
            # self.daten.kategorien = None,None

        self.excelWorker = None
        self.Enable()
        self.SetCursor(wx.Cursor(wx.CURSOR_ARROW))

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
        if not self.regeln.writeDatenToExcel(filePath):
            wx.MessageBox('Noch keine Daten vorhanden', 'Info',
                    wx.OK | wx.ICON_INFORMATION,
                    )
        saveFileDialog.Destroy()


    def onLoadRule(self, event):
        openFileDialog = wx.FileDialog(self, "Öffnen", "", "", 
                                      "TarmedPaketGUI files (*.tpf)|*.tpf", 
                                       wx.FD_OPEN | wx.FD_FILE_MUST_EXIST,
                                       )
        openFileDialog.ShowModal()
        filePath = pathlib.Path(openFileDialog.GetPath())
        openFileDialog.Destroy()
        self.regeln.loadFromFile(filePath)

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

        ruleBox = wx.StaticBox(panel, label="Regeln")
        ruleBoxSizer = wx.StaticBoxSizer(ruleBox, wx.HORIZONTAL)

        regelPanel = RegelPanel(panel, titel='Regel', regeln=self.regeln, daten=self.daten)
        bPanel1 = BedingungsPanel(panel, titel='AND', regeln=self.regeln, daten=self.daten)
        bPanel2 = BedingungsPanel(panel, titel='OR', regeln=self.regeln, daten=self.daten)
        bPanel3 = BedingungsPanel(panel, titel='NOT', regeln=self.regeln, daten=self.daten)

        bPanel1.listbox.setType(Regel.UND)
        bPanel2.listbox.setType(Regel.ODER)
        bPanel3.listbox.setType(Regel.NICHT)

        ruleBoxSizer.Add(regelPanel, proportion=1, flag=wx.EXPAND|wx.ALL, border=15)
        ruleBoxSizer.Add(bPanel1, proportion=1, flag=wx.EXPAND|wx.ALL, border=15)
        ruleBoxSizer.Add(bPanel2, proportion=1, flag=wx.EXPAND|wx.ALL, border=15)
        ruleBoxSizer.Add(bPanel3, proportion=1, flag=wx.EXPAND|wx.ALL, border=15)

        hBox2.Add(ruleBoxSizer, proportion=2, flag=wx.EXPAND|wx.ALL, border=15)

        # self.summaryPanel = SummaryPanel(panel)
        # self.daten.summaryPanel = self.summaryPanel
        # hBox2.Add(self.summaryPanel, proportion = 1, flag=wx.EXPAND|wx.ALL, border=15)

        mainBox.Add(hBox2, proportion=1, flag=wx.LEFT|wx.RIGHT|wx.EXPAND, border=15)

        panel.SetSizer(mainBox)

    def openExcel(self, event):
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

        self.SetCursor(wx.Cursor(wx.CURSOR_WAIT))
        self.excelPath.SetValue(filePath)
        self.Disable()
        self.excelWorker = ExcelReader(self, filePath)

    def menuhandler(self, event):
        id_ = event.GetId()
        if id_ == wx.ID_EXIT:
            self.Close()


def main():
    """Main GUI Loop"""
    app = wx.App()
    ex = TarmedpaketGUI(None)
    ex.Show()
    app.MainLoop()


if __name__ == '__main__':
    main()
