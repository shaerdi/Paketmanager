"""
ZetCode wxPython tutorial

In this example we create a new class layout
with wx.GridBagSizer.

author: Jan Bodnar
website: www.zetcode.com
last modified: April 2018
"""

import wx
import wx.lib.mixins.listctrl
from collections import OrderedDict
from ExcelCalc import datenEinlesen, createPakete
import pickle
import pathlib

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

    def __init__(self):
        self.regeln = OrderedDict()
        self.regeln['abc'] = { 
                'and' : ['1','2','3'],
                'or' : ['4','5'],
                'not' : ['6'],
                }
        self.regeln['neueRegel'] = { 
                'and' : ['1','2','3'],
                'or' : ['1','2','3'],
                'not' : [],
                }
        self.regeln['neueRegel2'] = { 
                'and' : ['1','2','3'],
                'or' : ['1','2','3'],
                'not' : ['5'],
                }
        self.aktiv = 'abc'

    def saveToFile(self,path):
        with path.with_suffix('.tpf').open('wb') as f:
            pickle.dump(self.regeln, f)

    def openFromFile(self,path):
        with path.with_suffix('.tpf').open('rb') as f:
            self.regeln = pickle.load(f)
        self.updateListen()

    def setAktiv(self,name):
        self.aktiv=name

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






class SummaryPanel(wx.Panel):
    def __init__(self, *args, **kwargs):
        super().__init__(*args,**kwargs)
        self.InitUI()

    def InitUI(self):
        sizer = wx.GridBagSizer(5,5)

        txt = wx.StaticText(self, label="infos", style=wx.ALIGN_CENTRE_HORIZONTAL)
        sizer.Add(txt, pos=(0,0), flag=wx.EXPAND)

        sizer.AddGrowableRow(0)
        sizer.AddGrowableCol(0)

        self.SetSizer(sizer)

class AnzeigeListe(wx.ListCtrl, wx.lib.mixins.listctrl.ListCtrlAutoWidthMixin):
    def __init__(self, parent,daten, *args, **kw):
        if not 'style' in kw:
            kw['style'] = wx.LC_REPORT|wx.LC_NO_HEADER|wx.LC_HRULES|wx.LC_VIRTUAL

        wx.ListCtrl.__init__(self, parent, **kw)
        wx.lib.mixins.listctrl.ListCtrlAutoWidthMixin.__init__(self)

        self.log = Log()

        self.daten = daten

        self.daten.Listen.append(self)
        
        self.InsertColumn(0,'')

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
        self.update()

    def update(self):
        self.items = self.daten.getRegeln()
        self.SetItemCount(len(self.items))

    def OnGetItemAttr(self, item):
        return None

class BedingungsListe(AnzeigeListe):
    def __init__(self, parent,daten,titel, *args, **kw):
        AnzeigeListe.__init__(self, parent,daten, *args, **kw)

        self.titel = titel.lower()
        self.update()

    def update(self):
        self.items = self.daten.getAktiveRegel(self.titel)
        self.SetItemCount(len(self.items))

    def OnGetItemAttr(self, item):
        return None


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
        sizer.Add( txt, pos=(0,0), span=(1,4), flag=wx.EXPAND,border=15)

        self.listbox = self.getCtrlList()
        sizer.Add( self.listbox, pos=(1,0), span=(1,4), flag=wx.EXPAND|wx.BOTTOM,border=15)

        def createBitmapButton(pfad, symbol):
            bmp = wx.Bitmap(pfad, wx.BITMAP_TYPE_ICO) 
            # btn = wx.BitmapButton(self, bitmap = bmp, size=(30,30))
            btn = wx.Button(self, label=symbol, size=(30,30))
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
        return RegelListe(self, self.daten, size=(100,-1))

    def SetupEvents(self):
        self.Bind(wx.EVT_BUTTON, self.NewItem, id=self.newBtn.GetId())
        self.Bind(wx.EVT_BUTTON, self.DelItem, id=self.delBtn.GetId())
        self.Bind(wx.EVT_BUTTON, self.ClrItem, id=self.clrBtn.GetId())
        self.Bind(wx.EVT_LIST_ITEM_SELECTED, self.OnClickItem, self.listbox)

    def OnClickItem(self, event):
        ind = event.GetIndex()
        self.updateAktiv(ind)

    def updateAktiv(self,ind=0):
        items  = self.listbox.items
        if len(items) > ind:
            self.daten.setAktiv(items[ind])
        else:
            self.daten.setAktiv('')

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
            self.daten.deleteItem(item)
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

    def NewItem(self, event):
        text = wx.GetTextFromUser('Enter a new item', 'Insert dialog')
        if text != '' and self.daten.aktiv in self.daten.regeln:
            aktiveRegel = self.daten.regeln[self.daten.aktiv]
            self.daten.regeln[self.daten.aktiv][self.titel].append(
                    text
                    )
            self.listbox.update()

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
    windowSize = (800,600)

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
        return
        dialog = wx.MessageDialog(self, message = "Programm wirklich Schliessen?", caption = "", style = wx.YES_NO, pos = wx.DefaultPosition)
        response = dialog.ShowModal()

        if (response == wx.ID_YES):
            self.OnExitApp(event)
        else:
            event.StopPropagation()

    def InitMenuBar(self):
        menubar = wx.MenuBar()

        fileMenu = wx.Menu()
        self.fileMenuExportItem = wx.MenuItem(fileMenu, wx.ID_ANY,
                text = "Export Excel",
                )
        fileMenu.Append(self.fileMenuExportItem)

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

        mainBox = wx.BoxSizer(wx.VERTICAL)

        hBox1 = wx.BoxSizer(wx.HORIZONTAL)
        text = wx.StaticText(panel, label='Aktuelles Excel')
        hBox1.Add(text, flag=wx.RIGHT|wx.TOP,border=5)

        path = wx.TextCtrl(panel)
        hBox1.Add(path, proportion=1, flag=wx.LEFT, border=10)

        excelOpenButton = wx.Button(panel, label='Öffnen')
        hBox1.Add(excelOpenButton, flag=wx.LEFT|wx.RIGHT, border=15)

        mainBox.Add(hBox1,flag=wx.TOP|wx.LEFT|wx.RIGHT|wx.EXPAND, border=15)

        #-----------------------------------
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

        hBox2.Add(ruleBoxSizer, proportion = 4, flag=wx.EXPAND|wx.ALL, border=15)

        hBox2.Add(SummaryPanel(panel), proportion = 1, flag=wx.EXPAND|wx.ALL, border=15)

        mainBox.Add(hBox2, proportion=1,flag=wx.LEFT|wx.RIGHT|wx.EXPAND, border=15)

        panel.SetSizer(mainBox)
        
    def openExcel(self,e):
        openFileDialog = wx.FileDialog(self, "Wählen", "", "", 
                                      "Excel files (*.xlsx)|*.xlsx", 
                                       wx.FD_OPEN | wx.FD_FILE_MUST_EXIST)
        openFileDialog.ShowModal()
        filePath = openFileDialog.GetPath()
        openFileDialog.Destroy()

        wait = wx.BusyCursor()
        # daten,kategorie = datenEinlesen(filePath)
        excelPanel = self.panels['excelPath']
        excelPanel.SetValue(filePath)

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
