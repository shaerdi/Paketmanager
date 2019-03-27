"""GUI Modul des TarmedPaketmanagers"""

import sys
import pathlib
from PyQt5 import QtCore, QtGui, QtWidgets
from ExcelCalc import datenEinlesen, createPakete, writePaketeToExcel
from ExcelCalc import Regeln, ExcelDaten, Regel
from UI import MainWindow, LeistungswahldialogUI

class ExcelReader(QtCore.QThread):
    """Thread, um ein Excel einzulesen"""

    signal = QtCore.pyqtSignal(dict)

    def __init__(self, parent, fname):
        super().__init__()
        self._fname = fname
        self.start()

    def run(self):
        returnValue = {}
        try:
            result = datenEinlesen(self._fname)
            if result is not None:
                daten, kategorien = result
                daten = createPakete(daten, kategorien)
                returnValue['success'] = True
                returnValue['data'] = (daten, kategorien)
            else:
                returnValue['success'] = False

        except Exception as error:
            returnValue['success'] = False
            returnValue['errMsg'] = '{}'.format(error)

        self.signal.emit(returnValue)

class ExcelPaketWriter(QtCore.QThread):
    """Thread, um ein Excel zu speichern"""
    def __init__(self, parent, fname, daten, kategorien):
        super().__init__()
        self._parent = parent
        self._fname = fname
        self._kategorien = kategorien
        self._daten = daten
        self.start()

    def run(self):
        try:
            writePaketeToExcel(self._daten, self._kategorien, self._fname)
            # wx.CallAfter(pub.sendMessage, 'excel.write', success=True)
        except Exception as error:
            errMsg = '{}'.format(error)
            # wx.CallAfter(pub.sendMessage,
                # 'excel.write',
                # success=False,
                # msg=errMsg,
            # )


# class BedingungswahlDialog(wx.Dialog):
    # def __init__(self, parent, id, title, daten):
        # wx.Dialog.__init__(self, parent, id, title)
        # self.daten = daten
        # self.InitUI()
        # self.SetList()

    # def InitUI(self):
        # vbox = wx.BoxSizer(wx.VERTICAL) 
        # hbox1 = wx.BoxSizer(wx.HORIZONTAL)
        # txt = wx.StaticText(self, label = "Neue Bedingung")
        # hbox1.Add(txt,proportion=0,flag=wx.ALL,border=5)
        # self.insertTxt = wx.TextCtrl(self)
        # hbox1.Add(self.insertTxt,proportion=1,flag=wx.EXPAND)
        # vbox.Add(hbox1, proportion=0, flag = wx.EXPAND|wx.TOP, border = 4) 

        # self.helpList = wx.ListBox(self,
                # style= wx.LB_SINGLE|wx.LB_NEEDED_SB,
                # )
        # vbox.Add(self.helpList, proportion=1, flag = wx.EXPAND|wx.ALL,border=5)

        # sizer =  self.CreateButtonSizer(wx.OK|wx.CANCEL)
        # vbox.Add(sizer, proportion=0, flag=wx.EXPAND|wx.ALL, border=5)
        # self.SetSizer(vbox)
        
        # self.Bind(wx.EVT_LISTBOX, self.OnListboxClicked, self.helpList)
        # self.Bind(wx.EVT_LISTBOX_DCLICK, self.OnListboxDoubleClicked, self.helpList)
        # self.Bind(wx.EVT_TEXT, self.OnTextChanged, self.insertTxt)

    # def OnTextChanged(self,event):
        # self.SetList()

    # def OnListboxDoubleClicked(self,event):
        # self.OnListboxClicked(event)
        # self.EndModal(wx.ID_OK)

    # def OnListboxClicked(self,event):
        # string = event.GetEventObject().GetStringSelection()
        # self.insertTxt.ChangeValue(string)
        
    # def SetList(self):
        # self.helpList.Clear()
        # leistungen = self.daten.getLeistungen(self.insertTxt.GetValue())
        # if not leistungen is None and not len(leistungen)==0:
            # self.helpList.InsertItems(leistungen,0)

    # def GetValue(self):
        # return self.insertTxt.GetValue()


# class SummaryPanel(wx.Panel):
    # def __init__(self, *args, **kwargs):
        # super().__init__(*args,**kwargs)
        # self.InitUI()

    # def InitUI(self):
        # sizer = wx.FlexGridSizer(3, 2, 10,50)

        # txt = wx.StaticText(self, label="Infos", style=wx.ALIGN_CENTRE_HORIZONTAL)
        # sizer.Add(txt)
        # sizer.Add(wx.StaticText(self))

        # txt = wx.StaticText(self, label="Anzahl Falldaten:", style=wx.ALIGN_LEFT)
        # sizer.Add(txt)

        # self.anzahlPaketeTotal = wx.StaticText(self, label="0", style=wx.ALIGN_LEFT)
        # sizer.Add(self.anzahlPaketeTotal)

        # txt = wx.StaticText(self, label="Falldaten mit Bedingung:", style=wx.ALIGN_LEFT)
        # sizer.Add(txt)

        # self.anzahlPaketeBedingung = wx.StaticText(self, label="0", style=wx.ALIGN_LEFT)
        # sizer.Add(self.anzahlPaketeBedingung)

        # self.SetSizer(sizer)

    # def updateTotal(self, num):
        # self.anzahlPaketeTotal.SetLabel( str(num) )

    # def updateBedingung(self, num):
        # self.anzahlPaketeBedingung.SetLabel( str(num) )


# class AnzeigeListe(wx.ListCtrl, wx.lib.mixins.listctrl.ListCtrlAutoWidthMixin):

    # def __init__(self, parent, regeln, daten, *args, **kw):
        # self._parent = parent
        # self.regeln = regeln
        # self.daten = daten
        # titel = kw.pop('titel','')

        # if 'style' not in kw:
            # kw['style'] = wx.LC_REPORT|wx.LC_HRULES|wx.LC_VIRTUAL

        # wx.ListCtrl.__init__(self, parent, *args, **kw)
        # wx.lib.mixins.listctrl.ListCtrlAutoWidthMixin.__init__(self)

        # self.InsertColumn(0, titel)

    # def OnGetItemText(self, item, col):
        # return self.items[item]


# class RegelListe(AnzeigeListe):
    # def __init__(self, parent, regeln, daten, *args, **kw):
        # style = wx.LC_REPORT|wx.LC_HRULES|wx.LC_VIRTUAL|wx.LC_SINGLE_SEL
        # AnzeigeListe.__init__(self, parent, regeln, daten, *args, style=style, **kw)
        # self.Bind(wx.EVT_LIST_ITEM_ACTIVATED, self.onDoubleClick)
        # self.Bind(wx.EVT_LIST_END_LABEL_EDIT, self.labelEdit)
        # self.Bind(wx.EVT_LIST_ITEM_SELECTED, self.onSetFocus)
        # regeln.registerObserver(self)
        # self.update()

    # def onSetFocus(self, event):
        # index = event.GetIndex()
        # self.regeln.setAktiv(index)

    # def labelEdit(self, event):
        # """Methode, die nach dem Editieren eines Labels aufgerufen wird"""
        # newLabel = event.GetLabel()
        # oldLabel = self.items[event.GetIndex()]
        # self.regeln.rename_regel(oldLabel, newLabel)
        # self.update()

    # def onDoubleClick(self,event):
        # """Methode, die bei einem Doppelklick aufgerufen wird"""
        # self.EditLabel(event.GetIndex())

    # def update(self):
        # """Liest die Items neu ein"""
        # if self.regeln is not None:
            # self.items = [r.name for r in self.regeln.regeln]
            # self.SetItemCount(len(self.regeln.regeln))

    # def deleteSelected(self):
        # """Loescht die selektierten Items """
        # index = self.GetFirstSelected()
        # self.regeln.removeRegel(index)


# class BedingungsListe(AnzeigeListe):
    # def __init__(self, parent, regeln, daten, *args, **kw):

        # self._typ = kw.pop('typ')
        # kw['titel'] = LIST_HEADER[self._typ]

        # AnzeigeListe.__init__(self, parent, regeln, daten, *args, **kw)

        # self.normalItem = wx.ListItemAttr()
        # self.redItem = wx.ListItemAttr()
        # self.redItem.SetBackgroundColour(wx.Colour(255,204,204))

        # regeln.registerObserver(self)
        # self.update()

    # def update(self):
        # """Setzt Listenitems neu
        # """
        # aktiveRegel = self.regeln.aktiveRegel
        # if aktiveRegel is not None:
            # self.items = aktiveRegel.getLeistungen(self._typ)
            # self.SetItemCount(len(self.items))
        # else:
            # self.items = []
            # self.SetItemCount(0)

    # def OnGetItemAttr(self, item):
        # """Prueft, ob ein Item in den Daten vorhanden ist

        # :item: Index des zu pruefenden Items
        # """
        # if self.daten.checkItem(self.items[item]):
            # return self.normalItem
        # else:
            # return self.redItem

    # def deleteSelected(self):
        # """Loescht die selektierten Items """
        # index = self.GetFirstSelected()
        # itemsToDelete = []
        # while index >= 0:
            # itemsToDelete.append(index)
            # index = self.GetNextSelected(index)

        # aktiveRegel = self.regeln.aktiveRegel
        # if aktiveRegel is not None:
            # for item in itemsToDelete:
                # aktiveRegel.removeLeistung(index, self._typ)


    # def newItem(self, event):
        # with BedingungswahlDialog(self, wx.ID_ANY, "Neue Bedingung", self.daten) as dlg:
            # if dlg.ShowModal() == wx.ID_OK:
                # text = dlg.GetValue()
                # aktiveRegel = self.regeln.aktiveRegel
                # if text != '' and aktiveRegel is not None:
                    # aktiveRegel.addLeistung(text, self.typ)


    # def clrItem(self, event):
        # aktiveRegel = self.regeln.aktiveRegel
        # if aktiveRegel is not None:
            # aktiveRegel.clearItems(self.typ)

class Leistungswahldialog(QtWidgets.QDialog):
    def __init__(self, parent, excelDaten, typ):
        super().__init__(parent)
        self._uInterface = LeistungswahldialogUI.Ui_Dialog()
        self._uInterface.setupUi(self)
        self._excelDaten = excelDaten
        self._neueLeistung = self._uInterface.NeueLeistung
        self._radioButtons = {
                Regel.UND : self._uInterface.radioButton_UND,
                Regel.ODER : self._uInterface.radioButton_ODER,
                Regel.NICHT : self._uInterface.radioButton_NICHT,
        }
        typ = typ or Regel.UND
        self._radioButtons[typ].setChecked(True)

        self.ok = False
        self.setupSlots()

    def setupListView(self):
        filterLeistung = self._neueLeistung.text()
        leistungen = self._excelDaten.getLeistungen(filterLeistung)
        model = QtCore.QStandardItemModel()
        model.appendRow(leistungen)
        self._uInterface.listView_Vorschlaege.setModel(model)

    def setupSlots(self):
        uInter = self._uInterface
        uInter.buttonBox.accepted.connect(self.okClicked)
        uInter.buttonBox.rejected.connect(self.cancelClicked)

    def okClicked(self):
        self.ok = True
        self.close()

    def cancelClicked(self):
        self.ok = False
        self.close()

    def getValue(self):
        typ = Regel.UND
        for t, button in self._radioButtons.items():
            if button.isChecked():
                typ = t
        return self._neueLeistung.text(), typ, self.ok


class RegelListe(QtCore.QAbstractListModel):
    neueRegel = QtCore.pyqtSignal()
    neueLeistung = QtCore.pyqtSignal(int)
    deleteLeistung = QtCore.pyqtSignal(int)

    def __init__(self, regeln, listViews):
        super().__init__()
        self._regeln = regeln
        self._regeln.registerObserver(self)
        self._bedingungsListViews = listViews
        for t, view in listViews.items():
            view.installEventFilter(self)

    def rowCount(self, parent=None):
        """Ueberschrieben von QAbstractListModel"""
        return len(self._regeln.regeln)

    def data(self, index, role):
        """Ueberschrieben von QAbstractListModel"""
        row = index.row()
        if role == QtCore.Qt.DisplayRole:
            return self._regeln.regeln[row].name
        return QtCore.QVariant()

    def selectionChanged(self, current, previous):
        self._regeln.setAktiv(current.row())

    def update(self):
        if self._regeln._aktiveRegel:
            bedingungen = self._regeln._aktiveRegel.getDict()
            for typ in [Regel.UND, Regel.ODER, Regel.NICHT]:
                model = QtGui.QStandardItemModel()
                for leistung in bedingungen[typ]:
                    model.appendRow(QtGui.QStandardItem(leistung))
                self._bedingungsListViews[typ].setModel(model)

    def addRegel(self, name):
        nItems = self.rowCount()
        self.beginInsertRows(QtCore.QModelIndex(), nItems, nItems+1)
        self._regeln.addRegel(name)
        self.endInsertRows()

    def deleteSelected(self):
        # TODO
        pass

    def eventFilter(self, watched, event):
        typ = None
        for t, view in self._bedingungsListViews.items():
            if view == watched:
                typ = t
        if event.type() == QtCore.QEvent.KeyPress:
            if event.key() == QtCore.Qt.Key_Delete:
                if typ == None:
                    self.deleteSelected()
                else:
                    self.deleteLeistung.emit(typ)
                return True
        elif event.type() == QtCore.QEvent.ContextMenu:
            if typ == None:
                self.neueRegel.emit()
            else:
                self.neueLeistung.emit(typ)
            return True
        return False



class TarmedPaketManagerApp(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()

        self._workerThread = None

        self.uInterface = MainWindow.Ui_MainWindow()
        self.uInterface.setupUi(self)

        self._excelDaten = ExcelDaten()
        self._regeln = Regeln(self._excelDaten)

        listViews = {
            Regel.UND : self.uInterface.listView_regel_und,
            Regel.ODER : self.uInterface.listView_regeln_oder,
            Regel.NICHT : self.uInterface.listView_regeln_nicht,
        }

        self._regelListe = RegelListe(self._regeln, listViews)
        self.uInterface.listView_regeln.setModel(self._regelListe)
        self.uInterface.listView_regeln.installEventFilter(self._regelListe)
        self.uInterface.listView_regeln.selectionModel().currentChanged.connect(
            self._regelListe.selectionChanged)

        self.setupSlots()


    def setupSlots(self):
        """Definiert die slot Funktionen der Menu Eintraege"""
        uInter = self.uInterface
        uInter.actionRohdaten_laden.triggered.connect(self.openExcel)
        uInter.actionNeue_Kategorie.triggered.connect(self.addKategorie)
        uInter.actionNeue_Regel.triggered.connect(self.addRegel)
        uInter.actionExcel_exportieren.triggered.connect(self.writeExcel)
        uInter.actionRegel_laden.triggered.connect(self.loadRegeln)
        uInter.actionRegeln_speichern.triggered.connect(self.writeRegeln)
        uInter.action_Exit.triggered.connect(self.quitApp)

        self._regelListe.neueRegel.connect(self.addRegel)
        self._regelListe.neueLeistung.connect(self.addLeistungToRegel)

    def addKategorie(self):
        """Fragt den Benutzer nach einer neuen Kategorie und fuegt sie dem
        Daten modell hinzu
        """
        pass

    def addRegel(self):
        """Fragt den Benutzer nach einer neuen Regel und fuegt sie hinzu"""
        name, ok = QtWidgets.QInputDialog.getText(self, "Neue Regel",
                "Name der neuen Regel:")
        if ok:
            self._regelListe.addRegel(name)


    def writeExcel(self):
        """Schreibt die Pakete in ein Excel"""
        pass

    def finishWrite(self, result):
        """Funktion, die nach dem Schreiben einer Datei aufgerufen wird

        :result: Dict, das vom Writer zurueck gegeben wird
        """
        if not result['success']:
            errMsg = result.get('errMsg','Es ist ein Fehler aufgetreten')
            box = QtWidgets.QMessageBox.warning(self, "Warnung", errMsg,
                QtWidgets.QMessageBox.Ok,)

        self._workerThread = None
        self.enableWindow()

    def writeRegeln(self):
        """Schreibt die Regeln"""
        options = QtWidgets.QFileDialog.Options()
        options |= QtWidgets.QFileDialog.DontUseNativeDialog
        fileName, _ = QtWidgets.QFileDialog.getSaveFileName(
            self,
            "Speichern unter",
            "","TPF file (*.tpf)",
            options=options
        )
        if fileName:
            path = pathlib.Path(fileName)
            if path.exists():
                reply = QtWidgets.QMessageBox.question(self, "Datei überschreiben",
                        "Die Datei existiert bereits. Überschreiben?",
                    QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No ,)
                if reply == QtWidgets.QMessageBox.No:
                    return
            self._regeln.saveToFile(path)

    def loadRegeln(self):
        """Laedt ein Regelfile"""
        options = QtWidgets.QFileDialog.Options()
        options |= QtWidgets.QFileDialog.DontUseNativeDialog
        fileName, _ = QtWidgets.QFileDialog.getOpenFileName(
            self,
            "Regeln laden",
            "","TPF file (*.tpf)",
            options=options
        )
        if fileName:
            self._regeln.loadFromFile(fileName)

    def addLeistungToRegel(self, typ=None):
        if self._regeln.getAktiv():
            dialog = Leistungswahldialog(self, self._excelDaten, typ)
            dialog.exec_()
            dialog.show()
            name, typ, ok = dialog.getValue()
            if ok:
                self._regeln.addLeistungToAktiverRegel(name, typ)


    # def onSaveRegelExcel(self,event):
        # saveFileDialog = wx.FileDialog(
            # self,
            # "Speichern unter", 
            # "", 
            # "", 
            # "Excel files (*.xlsx; *.xls)|*.xlsx;*.xls", 
            # wx.FD_SAVE,
        # )
        # saveFileDialog.ShowModal()
        # filePath = pathlib.Path(saveFileDialog.GetPath())
        # saveFileDialog.Destroy()
        # try:
            # dataframe = self.regeln.getBedingungsliste()
            # self.disableWindow()
            # self._currentWorker = ExcelDataFrameWriter(self, filePath, dataframe)
        # except UIError as error:
            # wx.MessageBox(
                # "{}".format(error),
                # 'Fehler',
                # wx.OK | wx.ICON_ERROR,
            # )


    # def onExitApp(self, event):
        # self.Destroy()

    # def onSaveExcel(self, event):
        # saveFileDialog = wx.FileDialog(
            # self, 
            # "Speichern unter", "", "",
            # "Excel files (*.xlsx; *.xls)|*.xlsx;*.xls",
            # wx.FD_SAVE,
           # )
        # saveFileDialog.ShowModal()
        # filePath = pathlib.Path(saveFileDialog.GetPath())
        # saveFileDialog.Destroy()

        # if self.daten.dataframe is None:
            # wx.MessageBox(
                # 'Noch keine Daten vorhanden',
                # 'Info',
                # wx.OK | wx.ICON_INFORMATION,
            # )
            # return

        # if filePath:
            # self.disableWindow()
            # self._currentWorker = ExcelPaketWriter(
                # self,
                # filePath,
                # self.daten.dataframe,
                # self.daten.getKategorien(),
            # )


    # def onLoadRule(self, event):
        # openFileDialog = wx.FileDialog(self, "Öffnen", "", "", 
                                      # "TarmedPaketGUI files (*.tpf)|*.tpf", 
                                       # wx.FD_OPEN | wx.FD_FILE_MUST_EXIST,
                                       # )
        # openFileDialog.ShowModal()
        # filePath = pathlib.Path(openFileDialog.GetPath())
        # openFileDialog.Destroy()
        # try:
            # self.regeln.loadFromFile(filePath)
        # except UIError as error:
            # wx.MessageBox(
                # message='{}'.format(error),
                # caption='Fehler',
                # style=wx.OK | wx.ICON_ERROR,
            # )

    def quitApp(self):
        print('bla')
        reply = QtWidgets.QMessageBox.question(self, "Beenden",
                "Programm wirklich beenden?",
            QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No ,)
        if reply == QtWidgets.QMessageBox.No:
            return
        QtGui.QGuiApplication.quit()

    def openExcel(self):
        """Laedt die Rohdaten"""
        options = QtWidgets.QFileDialog.Options()
        options |= QtWidgets.QFileDialog.DontUseNativeDialog
        fileName, _ = QtWidgets.QFileDialog.getOpenFileName(
            self,
            "Rohdaten laden",
            "","Excel Files (*.xlsx *.xls);;CSV Files (*.csv)",
            options=options
        )
        if fileName:
            self._workerThread = ExcelReader(self, fileName)
            self._workerThread.signal.connect(self.finishReadExcel)
            self.disableWindow()

    def finishReadExcel(self, result):
        """ Funktion, die nach dem Lesen eines Excels aufgerufen wird

        :result: Dict mit dem Signal des Thread
        """
        if result['success']:
            self._excelDaten.dataframe = result['data'][0]
            kategorien = result['data'][1]
            if kategorien is not None:
                for kategorie in kategorien:
                    self._excelDaten.addKategorie(kategorie)
        else:
            errMsg = result.get('errMsg', '')
            if errMsg:
                box = QtWidgets.QMessageBox.warning(
                    self,
                    "Warnung",
                    errMsg,
                    QtWidgets.QMessageBox.Ok,
                )

        self.enableWindow()

    def disableWindow(self):
        """Schaltet das Fenster in den Wartemodus"""
        QtWidgets.QApplication.setOverrideCursor(QtGui.QCursor(QtCore.Qt.WaitCursor))
        self.setEnabled(False)

    def enableWindow(self):
        """Schaltet den Wartemodus aus"""
        self.setEnabled(True)
        QtWidgets.QApplication.restoreOverrideCursor()


if __name__ == '__main__':
    app = QtWidgets.QApplication([])
    application = TarmedPaketManagerApp()
    application.show()
    sys.exit(app.exec())
