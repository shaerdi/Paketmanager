import wx
import wx.lib.mixins.listctrl
from collections import OrderedDict
from ExcelCalc import datenEinlesen, createPakete, writePaketeToExcel
import pickle
import pathlib
import threading
import pandas as pd


class ExcelReader(threading.Thread):
    def __init__(self, parent, fname):
        threading.Thread.__init__(self)
        self._parent = parent
        self._fname = fname
        self.start()

    def run(self):
        evt = ResultEvent(success=False)
        if not self._fname is None:
            try:
                result = datenEinlesen(self._fname)
                if not result is None:
                    daten, kategorien = result
                    daten = createPakete(daten, kategorien)
                    evt.success = True
                    evt.data = (daten, kategorien)
            except IOError as error:
                evt.errMsg = '{}'.format(error)

        wx.PostEvent(self._parent, evt)

class ExcelWriter(threading.Thread):
    def __init__(self, parent, fname, daten):
        threading.Thread.__init__(self)
        self._parent = parent
        self._fname = fname
        self._daten = daten
        self.start()

    def run(self):
        evt = ResultEvent(success=False)

        if not self._fname is None:
            try:
                result = datenEinlesen(self._fname)
                if not result is None:
                    daten, kategorien = result
                    daten = createPakete(daten, kategorien)
                    evt.success = True
                    evt.data = (daten, kategorien)
            except IOError as error:
                evt.errMsg = '{}'.format(error)

        wx.PostEvent(self._parent, evt)


EVT_RESULT_ID = 1001

tooltips = {
        'regel' : 'Definierte Regeln',
        'and' : 'Alle Leistungen müssen im Paket vorkommen',
        'or' : 'Mindestens eine Leistung muss im Paket vorkommen',
        'not' : 'Keine der Leistungen darf im Paket vorkommen',
}


def EVT_RESULT(win, func):
    win.Connect(-1, -1, EVT_RESULT_ID, func)


class ResultEvent(wx.PyEvent):
    def __init__(self,
            data = None,
            success = True,
            errMsg = '',
            ):
        wx.PyEvent.__init__(self)
        self.SetEventType(EVT_RESULT_ID)
        self.data=data
        self.success = success
        self.errMsg = ''


class Log:
    r"""\brief Needed by the wxdemos.
    The log output is redirected to the status bar of the containing frame.
    """

    def WriteText(self,text_string):
        self.write(text_string)

    def write(self,text_string):
        wx.GetApp().GetTopWindow().SetStatusText(text_string)


class DatenStruktur:
    Listen = []
    kategorien = None
    daten = None

    def __init__(self):
        self.regeln = OrderedDict()
        self.aktiv = ''

    def saveRegelnToExcel(self, filePath):
        if self.daten is None:
            return
        datenListe = [self.applyRegelToData(regel) for regel in self.regeln]
        datenListe = [l.drop_duplicates(subset='FallDatum') for l in datenListe]
        daten = pd.concat(datenListe)
        print('hallo')
        daten.to_excel(filePath, index=False)
        print('test')

    def applyRegelToData(self, regel=None):
        if self.daten is None:
            return
        if regel is None:
            regel = self.aktiv
        aktiveRegel = self.regeln[regel or self.aktiv]
        def erfuellt(key):
            erfuelltAlle = all([ (    k in key) for k in aktiveRegel['and']])
            erfuelltOder = len(aktiveRegel['or']) == 0 or \
                           any([ (    k in key) for k in aktiveRegel['or']])
            erfuelltNot  = all([ (not k in key) for k in aktiveRegel['not']])
            return  erfuelltAlle and erfuelltOder and erfuelltNot

        ind = self.daten.key.apply(erfuellt)
        kopie = self.daten[ind].copy()
        kopie.drop_duplicates(subset='FallDatum',inplace=True)
        kopie['Regel'] = regel
        return kopie

    def getAnzahlFalldaten(self):
        if not self.daten is None:
            return self.daten.FallDatum.drop_duplicates().shape[0]
        else:
            return 0

    def writeDatenToExcel(self,filePath):
        if not self.daten is None:
            writePaketeToExcel(self.daten, self.kategorien, filePath)
            return True
        else:
            return False

    def saveToFile(self,path):
        with path.with_suffix('.tpf').open('wb') as f:
            pickle.dump(self.regeln, f)

    def renameRegel(self, from_, to_):
        self.regeln = OrderedDict(
                (to_ if k == from_ else k, v) 
                for k, v in self.regeln.items()
                )

    def openFromFile(self,path):
        with path.with_suffix('.tpf').open('rb') as f:
            self.regeln = pickle.load(f)
        self.updateListen()

    def setAktiv(self,name):
        self.aktiv=name

    def CheckItem(self, item):
        if self.daten is None:
            return False
        else:
            return item in self.daten.Leistung.values

    def getLeistungen(self, filter_ = ''):
        if self.daten is None:
            return
        leistungen = self.daten[self.daten['Leistungskategorie'] == 'Tarmed']['Leistung']
        leistungen = leistungen.drop_duplicates()
        ind = leistungen.str.contains(filter_)
        return leistungen[ind].values

    def getRegeln(self):
        return list(self.regeln.keys())

    def deleteRegel(self,name):
        self.regeln.pop(name)

    def clearRegeln(self):
        self.regeln = OrderedDict()

    def addRegel(self,name):
        if name in self.regeln:
            return False
        else:
            self.regeln[name] = { 'and' : [], 'or' : [], 'not' : [] }
            return True

    def deleteItem(self,titel,item):
        if not titel in ['and','or','not']:
            return
        self.regeln[self.aktiv][titel].remove(item)

    def clearItems(self, titel):
        if not titel in ['and','or','not']:
            return
        self.regeln[self.aktiv][titel] = []


    def addItem(self,titel,item):
        if not titel in ['and','or','not']:
            return
        self.regeln[self.aktiv].append(item)

    def getAktiveRegel(self,titel):
        if titel in ['and','or','not'] and self.aktiv in self.regeln:
            return self.regeln[self.aktiv][titel]
        else:
            return []

    def updateListen(self):
        for l in self.Listen:
            l.update()
        self.updateSummaryPanel()

    def updateSummaryPanel(self):
        try:
            self.summaryPanel.updateBedingung( self.applyRegelToData().shape[0] )
        except:
            pass


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

    def __init__(self, parent,daten, *args, **kw):
        self._parent = parent

        if not 'style' in kw:
            kw['style'] = wx.LC_REPORT|wx.LC_NO_HEADER|wx.LC_HRULES|wx.LC_VIRTUAL

        wx.ListCtrl.__init__(self, parent, **kw)
        wx.lib.mixins.listctrl.ListCtrlAutoWidthMixin.__init__(self)

        self.log = Log()

        self.daten = daten

        self.daten.Listen.append(self)
        
        self.InsertColumn(0,'')

        self.Bind(wx.EVT_KEY_DOWN, lambda e:wx.PostEvent(self._parent,e) )

    def OnItemSelected(self, event):
        self.currentItem = event.m_itemIndex
        self.log.WriteText('OnItemSelected: "%s", "%s", "%s", "%s"\n' %
                           (self.currentItem,
                            self.GetItemText(self.currentItem),
                            self.getColumnText(self.currentItem, 1),
                            self.getColumnText(self.currentItem, 2)))

    def OnItemActivated(self, event):
        self.currentItem = event.m_itemIndex
        self.log.WriteText("OnItemActivated: %s\nTopItem: %s\n" %
                           (self.GetItemText(self.currentItem), self.GetTopItem()))

    def OnItemDeselected(self, evt):
        self.log.WriteText("OnItemDeselected: %s" % evt.m_itemIndex)

    def OnGetItemText(self, item, col):
        return self.items[item]


class RegelListe(AnzeigeListe):
    def __init__(self, parent,daten, *args, **kw):
        style = wx.LC_REPORT|wx.LC_NO_HEADER|wx.LC_HRULES|wx.LC_VIRTUAL|wx.LC_SINGLE_SEL
        AnzeigeListe.__init__(self, parent,daten, *args, style=style, **kw)
        self.Bind(wx.EVT_LIST_ITEM_ACTIVATED, self.OnDoubleClick)
        self.Bind(wx.EVT_LIST_ITEM_ACTIVATED, self.OnDoubleClick)
        self.Bind(wx.EVT_LIST_END_LABEL_EDIT, self.LabelEdit)
        self.update()

    def LabelEdit(self,event):
        newLabel = event.GetLabel()
        oldLabel = self.items[event.GetIndex()]
        self.daten.renameRegel(oldLabel, newLabel)
        self.update()

    def OnDoubleClick(self,event):
        self.EditLabel(event.GetIndex())

    def update(self):
        self.items = self.daten.getRegeln()
        self.SetItemCount(len(self.items))

    def OnGetItemAttr(self, item):
        return None


class BedingungsListe(AnzeigeListe):
    def __init__(self, parent,daten,titel, *args, **kw):
        AnzeigeListe.__init__(self, parent,daten, *args, **kw)

        self.normalItem = wx.ListItemAttr()
        self.redItem = wx.ListItemAttr()
        self.redItem.SetBackgroundColour(wx.Colour(255,204,204))

        self.titel = titel.lower()
        self.update()

    def update(self):
        self.items = self.daten.getAktiveRegel(self.titel)
        self.SetItemCount(len(self.items))

    def OnGetItemAttr(self, item):
        if self.daten.CheckItem(self.items[item]):
            return self.normalItem
        else:
            return self.redItem


class ListePanel(wx.Panel):
    def __init__(self, *args, **kwargs):

        titel = kwargs.pop('titel', '')
        self.daten = kwargs.pop('daten', {})

        super(ListePanel,self).__init__(*args,**kwargs)

        self.InitUI(titel)
        self.SetupEvents()

    def InitUI(self,titel):
        sizer = wx.GridBagSizer(5,5)

        txt = wx.StaticText(self, label=titel, style=wx.ALIGN_CENTRE_HORIZONTAL)
        txt.SetToolTip(wx.ToolTip(tooltips[titel.lower()]))
        sizer.Add( txt, pos=(0,0), span=(1,4), flag=wx.EXPAND,border=15)

        self.listbox = self.getCtrlList()
        sizer.Add( self.listbox, pos=(1,0), span=(1,4), flag=wx.EXPAND|wx.BOTTOM,border=15)

        def createBitmapButton(pfad, symbol):
            # bmp = wx.Bitmap(pfad, wx.BITMAP_TYPE_ICO) 
            # btn = wx.BitmapButton(self, bitmap = bmp, size=(30,30))
            btn = wx.Button(self, label=symbol, size=(50,30))
            font = wx.Font(15, wx.DEFAULT, wx.NORMAL, wx.BOLD)
            btn.SetFont(font)
            return btn

        self.newBtn = createBitmapButton('./Bilder/Plus.ico', '+')
        self.delBtn = createBitmapButton('./Bilder/Minus.ico', '-')
        self.clrBtn = createBitmapButton('./Bilder/Clear.ico', 'X')

        sizer.Add( self.newBtn, pos=(2,0) )
        sizer.Add( self.delBtn, pos=(2,1) )
        sizer.Add( self.clrBtn, pos=(2,3) )

        sizer.AddGrowableRow(1)
        sizer.AddGrowableCol(2)

        self.SetSizer(sizer)


class RegelPanel(ListePanel):
    def __init__(self, *args, **kwargs):
        super(RegelPanel,self).__init__(*args,**kwargs)
        self.setFocus()

    def setFocus(self, index = 0):
        listLen = len(self.listbox.items)
        if listLen == 0: return
        index = max(min( listLen-1, index), 0)
        self.listbox.Select(index)
        self.listbox.Focus(index)
        self.updateAktiv(index)

    def getCtrlList(self):
        return RegelListe(self, self.daten, size=(70,-1))

    def SetupEvents(self):
        self.Bind(wx.EVT_BUTTON, self.NewItem, id=self.newBtn.GetId())
        self.Bind(wx.EVT_BUTTON, self.DelItem, id=self.delBtn.GetId())
        self.Bind(wx.EVT_BUTTON, self.ClrItem, id=self.clrBtn.GetId())
        self.Bind(wx.EVT_LIST_ITEM_SELECTED, self.OnClickItem, self.listbox)
        self.Bind(wx.EVT_KEY_DOWN, self.OnKeyPress)

    def OnKeyPress(self,event):
        keycode = event.GetKeyCode()

        index = self.listbox.GetFocusedItem()
        if index < 0:
            return

        if keycode == wx.WXK_UP and index > 0:
            self.listbox.Select(index-1)
            self.listbox.Focus(index-1)
        elif keycode == wx.WXK_DOWN and index < self.listbox.GetItemCount()-1:
            self.listbox.Select(index+1)
            self.listbox.Focus(index+1)
        elif keycode == wx.WXK_DELETE or keycode == wx.WXK_NUMPAD_DELETE:
            self.DelItem(event)

    def OnClickItem(self, event):
        ind = event.GetIndex()
        self.updateAktiv(ind)

    def updateAktiv(self,ind=0):
        items  = self.listbox.items
        if len(items) > ind:
            self.daten.setAktiv(items[ind])
        else:
            self.daten.setAktiv('')

        self.daten.updateSummaryPanel()
        self.daten.updateListen()
        
    def NewItem(self, event):
        text = wx.GetTextFromUser('Enter a new item', 'Insert dialog')
        if text != '':
            self.daten.addRegel(text)
            self.daten.updateListen()
            self.listbox.Select( len(self.daten.regeln) -1 )
            self.listbox.Focus( len(self.daten.regeln) -1 )

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
        self.titel = kwargs.get('titel','').lower()
        super().__init__(*args,**kwargs)

    def getCtrlList(self):
        return BedingungsListe(self, self.daten, self.titel, size=(100,-1))

    def SetupEvents(self):
        self.Bind(wx.EVT_BUTTON, self.NewItem, id=self.newBtn.GetId())
        self.Bind(wx.EVT_BUTTON, self.DelItem, id=self.delBtn.GetId())
        self.Bind(wx.EVT_BUTTON, self.ClrItem, id=self.clrBtn.GetId())
        self.Bind(wx.EVT_KEY_DOWN, self.OnKeyPress)

    def OnKeyPress(self,event):
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
        with BedingungswahlDialog(self,wx.ID_ANY, "Neue Bedingung", self.daten) as dlg:
            if dlg.ShowModal() == wx.ID_OK:
                text = dlg.GetValue()
                if text != '' and self.daten.aktiv in self.daten.regeln:
                    aktiveRegel = self.daten.regeln[self.daten.aktiv]
                    self.daten.regeln[self.daten.aktiv][self.titel].append(
                            text
                            )
                    self.listbox.update()
                    self.daten.updateSummaryPanel()

    def DelItem(self, event):
        index = self.listbox.GetFirstSelected()
        itemsToDelete = []
        while index >= 0:
            itemsToDelete.append(self.listbox.GetItem(index).GetText())
            index = self.listbox.GetNextSelected(index)

        for item in itemsToDelete:
            self.daten.deleteItem(self.titel, item)
        self.daten.updateListen()

    def ClrItem(self, event):
        self.daten.clearItems(self.titel)
        self.daten.updateListen()


class TarmedpaketGUI(wx.Frame):
    name = "Tarmed Pakete"
    windowSize = (1300,800)

    panels = {}

    def __init__(self, parent):
        super().__init__(parent, 
                title=self.name,
                size=self.windowSize,
                )

        self.daten = DatenStruktur()

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
        saveFileDialog = wx.FileDialog(self, "Speichern unter", "", "", 
                                      "Excel files (*.xlsx; *.xls)|*.xlsx;*.xls", 
                                       wx.FD_SAVE,
                                       )
        saveFileDialog.ShowModal()
        filePath = pathlib.Path(saveFileDialog.GetPath())
        saveFileDialog.Destroy()
        self.daten.saveRegelnToExcel(filePath)

    def FinishExcelCalc(self, event):
        if event.success:
            self.daten.daten = event.data[0]
            self.daten.kategorien = event.data[1]
            self.summaryPanel.updateTotal( self.daten.getAnzahlFalldaten() )
            self.daten.updateSummaryPanel()
        else:
            if event.errMsg:
                wx.MessageBox(
                        message=event.errMsg,
                        caption='Fehler',
                        style=wx.OK | wx.ICON_INFORMATION,
                        )
            self.daten.daten, self.daten.kategorien = None,None

        self.excelWorker = None
        self.Enable()
        self.SetCursor(wx.Cursor(wx.CURSOR_ARROW))

    def OnExitApp(self, event):
        self.Destroy()

    def OnSaveRule(self, event):
        saveFileDialog = wx.FileDialog(self, "Speichern unter", "", "", 
                                      "TarmedPaketGUI files (*.tpf)|*.tpf", 
                                       wx.FD_SAVE,
                                       )
        saveFileDialog.ShowModal()
        filePath = pathlib.Path(saveFileDialog.GetPath())
        saveFileDialog.Destroy()
        self.daten.saveToFile(filePath)

    def OnSaveExcel(self, event):
        saveFileDialog = wx.FileDialog(self, "Speichern unter", "", "", 
                                      "Excel files (*.xlsx; *.xls)|*.xlsx;*.xls", 
                                       wx.FD_SAVE,
                                       )
        saveFileDialog.ShowModal()
        filePath = pathlib.Path(saveFileDialog.GetPath())
        if not self.daten.writeDatenToExcel(filePath):
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
        self.daten.openFromFile(filePath)

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
        
        # tc = wx.TextCtrl(panel)
        pl = RegelPanel(panel, titel='Regel', daten=self.daten)
        pl1 = BedingungsPanel(panel, titel='AND', daten=self.daten)
        pl2 = BedingungsPanel(panel, titel='OR', daten=self.daten)
        pl3 = BedingungsPanel(panel, titel='NOT', daten=self.daten)
        ruleBoxSizer.Add(pl, proportion=1,flag=wx.EXPAND|wx.ALL, border=15)
        ruleBoxSizer.Add(pl1, proportion=1,flag=wx.EXPAND|wx.ALL, border=15)
        ruleBoxSizer.Add(pl2, proportion=1,flag=wx.EXPAND|wx.ALL, border=15)
        ruleBoxSizer.Add(pl3, proportion=1,flag=wx.EXPAND|wx.ALL, border=15)

        hBox2.Add(ruleBoxSizer, proportion = 2, flag=wx.EXPAND|wx.ALL, border=15)

        self.summaryPanel = SummaryPanel(panel)
        self.daten.summaryPanel = self.summaryPanel
        hBox2.Add(self.summaryPanel, proportion = 1, flag=wx.EXPAND|wx.ALL, border=15)

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
