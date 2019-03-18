from collections import OrderedDict
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


class Regel:
    """Stellt eine Regel dar, die ein Paket erfuellen kann oder nicht"""

    UND = 0
    ODER = 1
    NICHT = 2

    def __init__(self, name, daten):
        self.name = name
        self._bedingung_und = []
        self._bedingung_oder = []
        self._bedingung_nicht = []
        self.anzahl = '-'
        self._daten = daten
        self._erfuellt = None

    def add_leistung(self, new_item, bedingungs_art):
        """Fuegt eine neue Leistung zur einer Liste hinzu

        :new_item: Neue Leistung
        :bedingungs_art: Regel.AND, OR oder NOT
        """

        if bedingungs_art == Regel.UND:
            self._bedingung_und.append(new_item)
        elif bedingungs_art == Regel.ODER:
            self._bedingung_oder.append(new_item)
        elif bedingungs_art == Regel.NICHT:
            self._bedingung_nicht.append(new_item)
        else:
            raise RuntimeError("Unbekannte Bedingung")

    def remove_leistung(self, index, bedingungs_art):
        """Loescht eine Leistung aus einer Liste

        :index: Index der zu loeschenden Leistung
        :bedingungs_art: Regel.AND, OR oder NOT
        """

        if bedingungs_art == Regel.UND:
            del self._bedingung_und[index]
        elif bedingungs_art == Regel.ODER:
            del self._bedingung_oder[index]
        elif bedingungs_art == Regel.NICHT:
            del self._bedingung_nicht[index]
        else:
            raise RuntimeError("Unbekannte Bedingung")

    def update(self):
        """Berechnet die Pakete, die diese Regel erfuellen"""

        if self._daten is None:
            self.anzahl = '-'
            return

        def erfuellt(key):
            """Checkt, ob ein Key diese Regel erfuellt"""
            erfuelltalle = all([(k in key) for k in self._bedingung_und])
            erfuelltoder = len(self._bedingung_oder) == 0 or \
                           any([(k in key) for k in self._bedingung_oder])
            erfuelltnot  = all([(k not in key) for k in self._bedingung_nicht])
            return  erfuelltalle and erfuelltoder and erfuelltnot

        ind = self._daten.key.apply(erfuellt)
        self._erfuellt = self._daten[ind]
        self.anzahl = str(ind.sum)

    def get_erfuellt(self):
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
        if bedingungs_art == Regel.UND:
            return self._bedingung_und
        elif bedingungs_art == Regel.ODER:
            return self._bedingung_oder
        elif bedingungs_art == Regel.NICHT:
            return self._bedingung_nicht
        else:
            raise RuntimeError("Unbekannte Bedingung")


class Regeln:
    """Klasse, die die Regeln speichert"""

    def __init__(self):
        self.observers = []
        self.regeln = []
        self.aktiveRegel = None
        self._excel_daten = None

    def register_observer(self, observer):
        """Registriert ein Observerobjekt, das per Aufrufen der Funktion update
        auf Aenderungen aufmerksam gemacht wird.

        :observer: Observer objekt. Muss die Funktion update haben

        """
        observer.regeln = self
        self.observers.append(observer)

    def notify_observers(self):
        """Ruft die Methode update fuer alle Observer auf

        """
        for observer in self.observers:
            observer.update()

    @property
    def excel_daten(self):
        """Getter ecxel_daten"""
        return self._excel_daten

    @excel_daten.setter
    def excel_daten(self, daten):
        """Setter ecxel_daten"""
        self._excel_daten = daten
        self.update_regel()

    def update_regel(self, index=None):
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
        neue_regel = Regel(name, self._excel_daten)
        self.regeln.append(neue_regel)
        self.notify_observers()

    def rename_regel(self, index, neuer_name):
        """Benennt eine Regel um.

        :index: Index der Regel, die umbenannt wird
        :neuer_name: Neuer Name
        """
        self.regeln[index].name = neuer_name
        self.notify_observers()

    def remove_regel(self, index):
        """Loescht eine Regel.

        :index: Index, der geloscht wird.
        """
        del self.regeln[index]
        self.notify_observers()

    def checkItem(self, label):
        """Prueft, ob eine Bedingung in den Daten vorhanden ist

        :label: Name der Bedingung
        :returns: True, wenn die Bedingung vorhanden ist
        """
        if self._excel_daten is None:
            return False
        else:
            return item in self._excel_daten.Leistung.values

    def get_bedingungsliste(self, filename):
        """Speichert ein Excel, in dem fuer jedes Falldatum eine Zeile fuer
        jede Regel steht, die dieses Falldatum erfuellt

        :filename: Pfad zur Datei
        """

        datenListe = [regel.get_erfuellt() for regel in self.regeln]
        datenListe = [l.drop_duplicates(subset='FallDatum') for l in datenListe]
        return pd.concat(datenListe)

    def get_anzahl_falldaten(self):
        """Gibt die Anzahl Falldaten zurueck

        :return: Anzahl Falldaten
        """

        if self._excel_daten is not None:
            return self.excel_daten.FallDatum.drop_duplicates().shape[0]
        else:
            return 0

    def save_to_file(self,filename):
        """Speichert die enthaltenen Regeln in ein File

        :filename: Filename
        """
        path = pathlib.Path(filename)
        with path.with_suffix('.tpf').open('wb') as f:
            pickle.dump(self.regeln, f)

    def loadFromFile(self,filename):
        """Laedt die Regeln aus einem File

        :filename: Filename
        """
        path = pathlib.Path(filename)
        with path.with_suffix('.tpf').open('rb') as f:
            self.regeln = pickle.load(f)
        self.update_regel()

    def set_aktiv(self, index):
        """Setzt die momentan aktive Regel

        :index: Index der neuen aktiven Regel
        """
        if 0 <= index < len(self.regeln):
            self.aktiveRegel = self.regeln[index]


class ExcelDaten:
    """Objekt, das die Excel Daten enthaelt"""

    def __init__(self, daten, kategorien=None):
        self._dataframe = daten
        self._kategorien = set(kategorien)

    def addKategorie(self, kategorie):
        """Fuegt eine Kategorie hinzu"""
        self._kategorien.add(kategorie)

    def calcUniqueLeistungen(self):
        """Berechnet eine Liste mit allen Leistungen im Excel"""
        leistungen = self.dataframe[self.daten['Leistungskategorie'] == 'Tarmed']['Leistung']
        leistungen = leistungen.drop_duplicates()
        ind = leistungen.str.contains(filter_)
        self._leistungen = leistungen[ind].values
            





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
        leistungen = self.regeln.getLeistungen(self.insertTxt.GetValue())
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

    def __init__(self, parent, regeln, *args, **kw):
        self._parent = parent
        self.regeln = regeln

        if not 'style' in kw:
            kw['style'] = wx.LC_REPORT|wx.LC_NO_HEADER|wx.LC_HRULES|wx.LC_VIRTUAL

        wx.ListCtrl.__init__(self, parent, **kw)
        wx.lib.mixins.listctrl.ListCtrlAutoWidthMixin.__init__(self)

        self.InsertColumn(0, '')
        # self.Bind(wx.EVT_KEY_DOWN, lambda e:wx.PostEvent(self._parent,e))

    # def OnItemSelected(self, event):
        # self.currentItem = event.m_itemIndex

    # def OnItemActivated(self, event):
        # self.currentItem = event.m_itemIndex

    def OnGetItemText(self, item, col):
        return self.items[item]


class RegelListe(AnzeigeListe):
    def __init__(self, parent, regeln, *args, **kw):
        style = wx.LC_REPORT|wx.LC_NO_HEADER|wx.LC_HRULES|wx.LC_VIRTUAL|wx.LC_SINGLE_SEL
        AnzeigeListe.__init__(self, parent, regeln, *args, style=style, **kw)
        self.Bind(wx.EVT_LIST_ITEM_ACTIVATED, self.OnDoubleClick)
        self.Bind(wx.EVT_LIST_ITEM_ACTIVATED, self.OnDoubleClick)
        self.Bind(wx.EVT_LIST_END_LABEL_EDIT, self.LabelEdit)
        regeln.register_observer(self)
        self.update()

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
    def __init__(self, parent, regeln, *args, **kw):
        AnzeigeListe.__init__(self, parent, regeln, *args, **kw)

        self.normalItem = wx.ListItemAttr()
        self.redItem = wx.ListItemAttr()
        self.redItem.SetBackgroundColour(wx.Colour(255,204,204))

        regeln.register_observer(self)
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
            self.items = aktiveRegel.getLeistungen(self.typ)
            self.SetItemCount(len(self.items))
        else:
            self.items = []
            self.SetItemCount(0)

    def OnGetItemAttr(self, item):
        """Prueft, ob ein Item in den Daten vorhanden ist

        :item: Index des zu pruefenden Items
        """
        if self.regeln.checkItem(self.items[item]):
            return self.normalItem
        else:
            return self.redItem

class ListePanel(wx.Panel):
    def __init__(self, *args, **kwargs):

        titel = kwargs.pop('titel', '')
        self.regeln = kwargs.pop('regeln', {})

        super(ListePanel, self).__init__(*args, **kwargs)

        self.InitUI(titel)
        self.SetupEvents()

    def SetupEvents(self):
        raise NotImplementedError()

    def getCtrlList(self):
        raise NotImplementedError()

    def InitUI(self,titel):
        sizer = wx.GridBagSizer(5,5)

        txt = wx.StaticText(self, label=titel, style=wx.ALIGN_CENTRE_HORIZONTAL)
        txt.SetToolTip(wx.ToolTip(TOOLTIPS[titel.lower()]))
        sizer.Add( txt, pos=(0,0), span=(1,4), flag=wx.EXPAND,border=15)

        self.listbox = self.getCtrlList()
        sizer.Add( self.listbox, pos=(1,0), span=(1,4), flag=wx.EXPAND|wx.BOTTOM,border=15)

        def create_button(symbol):
            btn = wx.Button(self, label=symbol, size=(50,30))
            font = wx.Font(15, wx.DEFAULT, wx.NORMAL, wx.BOLD)
            btn.SetFont(font)
            return btn

        self.newBtn = create_button('+')
        self.delBtn = create_button('-')
        self.clrBtn = create_button('X')

        sizer.Add( self.newBtn, pos=(2,0))
        sizer.Add( self.delBtn, pos=(2,1))
        sizer.Add( self.clrBtn, pos=(2,3))

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
        self.updateAktiv(index)

    def getCtrlList(self):
        return RegelListe(self, self.regeln, size=(70,-1))

    def SetupEvents(self):
        self.Bind(wx.EVT_BUTTON, self.NewItem, id=self.newBtn.GetId())
        self.Bind(wx.EVT_BUTTON, self.DelItem, id=self.delBtn.GetId())
        self.Bind(wx.EVT_BUTTON, self.ClrItem, id=self.clrBtn.GetId())

    def NewItem(self, event):
        text = wx.GetTextFromUser('Enter a new item', 'Insert dialog')
        if text != '':
            self.regeln.addRegel(text)
            self.listbox.Select( len(self.regeln.regeln) -1 )
            self.listbox.Focus( len(self.regeln.regeln) -1 )

    def DelItem(self, event):
        index = self.listbox.GetFirstSelected()
        if index >= 0:
            item = self.listbox.GetItem(index).GetText()
            self.daten.deleteRegel(item)
            self.daten.updateListen()
            self.setFocus(index)

    def ClrItem(self, event):
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
        liste = BedingungsListe(self, self.regeln, size=(100, -1))
        liste.setType(self.typ)
        return liste

    def SetupEvents(self):
        self.Bind(wx.EVT_BUTTON, self.NewItem, id=self.newBtn.GetId())
        self.Bind(wx.EVT_BUTTON, self.DelItem, id=self.delBtn.GetId())
        self.Bind(wx.EVT_BUTTON, self.ClrItem, id=self.clrBtn.GetId())
        self.Bind(wx.EVT_KEY_DOWN, self.OnKeyPress)

    def OnKeyPress(self, event):
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

    def NewItem(self, event):
        with BedingungswahlDialog(self, wx.ID_ANY, "Neue Bedingung", self.regeln) as dlg:
            if dlg.ShowModal() == wx.ID_OK:
                text = dlg.GetValue()
                aktiveRegel = self.regeln.aktiveRegel
                if text != '' and aktiveRegel is not None:
                    aktiveRegel.add_leistung(text, self.typ)

    def DelItem(self, event):
        index = self.listbox.GetFirstSelected()
        itemsToDelete = []
        while index >= 0:
            itemsToDelete.append(self.listbox.GetItem(index).GetText())
            index = self.listbox.GetNextSelected(index)

        for item in itemsToDelete:
            self.regeln.deleteItem(self.titel, item)

    def ClrItem(self, event):
        self.regeln.clearItems(self.titel)
        self.regeln.updateListen()


class TarmedpaketGUI(wx.Frame):
    name = "Tarmed Pakete"
    windowSize = (1300,800)

    panels = {}

    def __init__(self, parent):
        super().__init__(parent, 
                title=self.name,
                size=self.windowSize,
                )

        self.regeln = Regeln()

        self.InitUI()
        self.SetupEvents()
        self.Centre()

    def SetupEvents(self):
        self.Bind(wx.EVT_CLOSE, self.OnCloseFrame)
        self.Bind(wx.EVT_MENU, self.OnCloseFrame, self.fileMenuExitItem)
        self.Bind(wx.EVT_MENU, self.OnSaveRule, self.fileMenuSaveRule)
        self.Bind(wx.EVT_MENU, self.OnLoadRule, self.fileMenuLoadRule)
        self.Bind(wx.EVT_MENU, self.OnSaveExcel, self.fileMenuExportExcel)
        self.Bind(wx.EVT_MENU, self.OnSaveRegelExcel, self.fileMenuExportRegelExcel)
        self.Bind(wx.EVT_BUTTON, self.OpenExcel, self.excelOpenButton)

        EVT_RESULT(self, self.FinishExcelCalc)

    def OnSaveRegelExcel(self,event):
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

    def FinishExcelCalc(self, event):
        if event.success:
            self.regeln.excel_daten = event.data[0]
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
            self.regeln.excel_daten = None
            # self.daten.kategorien = None,None

        self.excelWorker = None
        self.Enable()
        self.SetCursor(wx.Cursor(wx.CURSOR_ARROW))

    def OnExitApp(self, event):
        self.Destroy()

    def OnSaveRule(self, event):
        saveFileDialog = wx.FileDialog(
            self, 
            "Speichern unter", "", "", 
            "TarmedPaketGUI files (*.tpf)|*.tpf", 
            wx.FD_SAVE,
           )
        saveFileDialog.ShowModal()
        file_path = pathlib.Path(saveFileDialog.GetPath())
        saveFileDialog.Destroy()
        self.regeln.save_to_file(file_path)

    def OnSaveExcel(self, event):
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


    def OnLoadRule(self, event):
        openFileDialog = wx.FileDialog(self, "Öffnen", "", "", 
                                      "TarmedPaketGUI files (*.tpf)|*.tpf", 
                                       wx.FD_OPEN | wx.FD_FILE_MUST_EXIST,
                                       )
        openFileDialog.ShowModal()
        filePath = pathlib.Path(openFileDialog.GetPath())
        openFileDialog.Destroy()
        self.regeln.loadFromFile(filePath)

    def OnCloseFrame(self, event):
        self.OnExitApp(event)
        dialog = wx.MessageDialog(self, message="Programm wirklich Schliessen?", caption="", style=wx.YES_NO, pos=wx.DefaultPosition)
        response = dialog.ShowModal()

        if response == wx.ID_YES:
            self.OnExitApp(event)
        else:
            event.StopPropagation()

    def InitMenuBar(self):
        menubar = wx.MenuBar()

        fileMenu = wx.Menu()
        self.fileMenuExportExcel = wx.MenuItem(fileMenu, wx.ID_ANY,
                text = "Export Excel",
                )
        fileMenu.Append(self.fileMenuExportExcel)

        self.fileMenuExportRegelExcel = wx.MenuItem(fileMenu, wx.ID_ANY,
                text = "Export Bedingungen in Excel",
                )
        fileMenu.Append(self.fileMenuExportRegelExcel)

        fileMenu.AppendSeparator()

        self.fileMenuSaveRule = wx.MenuItem(fileMenu, wx.ID_SAVE,
                text = "Regeln speichern",
                )
        self.fileMenuLoadRule = wx.MenuItem(fileMenu, wx.ID_OPEN,
                text = "Regeln laden",
                )

        fileMenu.Append(self.fileMenuSaveRule)
        fileMenu.Append(self.fileMenuLoadRule)

        fileMenu.AppendSeparator()

        self.fileMenuExitItem = wx.MenuItem(fileMenu, wx.ID_EXIT, '&Quit\tCtrl+Q') 
        fileMenu.Append(self.fileMenuExitItem)
        
        menubar.Append(fileMenu, '&Datei')
        self.SetMenuBar(menubar)



    def InitUI(self):
        self.InitMenuBar()

        panel = wx.Panel(self)
        self.panel = panel

        mainBox = wx.BoxSizer(wx.VERTICAL)

        hBox1 = wx.BoxSizer(wx.HORIZONTAL)
        text = wx.StaticText(panel, label='Aktuelles Excel')
        hBox1.Add(text, flag=wx.RIGHT|wx.TOP,border=5)

        self.excelPath = wx.TextCtrl(panel)
        hBox1.Add(self.excelPath, proportion=1, flag=wx.LEFT, border=10)

        self.excelOpenButton = wx.Button(panel, label='Öffnen')
        hBox1.Add(self.excelOpenButton, flag=wx.LEFT|wx.RIGHT, border=15)

        mainBox.Add(hBox1,flag=wx.TOP|wx.LEFT|wx.RIGHT|wx.EXPAND, border=15)

        # -----------------------------------
        line = wx.StaticLine(panel)
        mainBox.Add(line, flag=wx.EXPAND|wx.BOTTOM|wx.TOP, border=15)

        hBox2 = wx.BoxSizer(wx.HORIZONTAL)

        ruleBox = wx.StaticBox(panel, label= "Regeln")
        ruleBoxSizer = wx.StaticBoxSizer(ruleBox, wx.HORIZONTAL)
        
        pl = RegelPanel(panel, titel='Regel', regeln=self.regeln)
        pl1 = BedingungsPanel(panel, titel='AND', regeln=self.regeln)
        pl2 = BedingungsPanel(panel, titel='OR', regeln=self.regeln)
        pl3 = BedingungsPanel(panel, titel='NOT', regeln=self.regeln)
        ruleBoxSizer.Add(pl, proportion=1,flag=wx.EXPAND|wx.ALL, border=15)
        ruleBoxSizer.Add(pl1, proportion=1,flag=wx.EXPAND|wx.ALL, border=15)
        ruleBoxSizer.Add(pl2, proportion=1,flag=wx.EXPAND|wx.ALL, border=15)
        ruleBoxSizer.Add(pl3, proportion=1,flag=wx.EXPAND|wx.ALL, border=15)

        hBox2.Add(ruleBoxSizer, proportion = 2, flag=wx.EXPAND|wx.ALL, border=15)

        # self.summaryPanel = SummaryPanel(panel)
        # self.daten.summaryPanel = self.summaryPanel
        # hBox2.Add(self.summaryPanel, proportion = 1, flag=wx.EXPAND|wx.ALL, border=15)

        mainBox.Add(hBox2, proportion=1,flag=wx.LEFT|wx.RIGHT|wx.EXPAND, border=15)

        panel.SetSizer(mainBox)
        
    def OpenExcel(self,e):
        openFileDialog = wx.FileDialog(self, "Wählen", "", "Excel files (*.xlsx;*.xls)|*.xlsx;*.xls", "",
                                        wx.FD_OPEN | wx.FD_FILE_MUST_EXIST
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

    app = wx.App()
    ex = TarmedpaketGUI(None)
    ex.Show()
    app.MainLoop()


if __name__ == '__main__':
    main()
