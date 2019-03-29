"""GUI Modul des TarmedPaketmanagers"""

import sys
import pathlib
import pickle
from PyQt5 import QtCore, QtGui, QtWidgets
from ExcelCalc import datenEinlesen, createPakete, writePaketeToExcel
from ExcelCalc import Regeln, ExcelDaten, Regel, UIError
import MainWindow, LeistungswahldialogUI, Ueber

VERSION = "0.9.0"
BESCHREIBUNG = """
Tarmed Paketmanager Version {}

Copyright (c) 2019 Simon Härdi

shaerdi@protonmail.ch

Hiermit wird unentgeltlich jeder Person, die eine Kopie der Software und der zugehörigen Dokumentationen (die "Software") erhält, die Erlaubnis erteilt, sie uneingeschränkt zu nutzen, inklusive und ohne Ausnahme mit dem Recht, sie zu verwenden, zu kopieren, zu verändern, zusammenzufügen, zu veröffentlichen, zu verbreiten, zu unterlizenzieren und/oder zu verkaufen, und Personen, denen diese Software überlassen wird, diese Rechte zu verschaffen, unter den folgenden Bedingungen:

Der obige Urheberrechtsvermerk und dieser Erlaubnisvermerk sind in allen Kopien oder Teilkopien der Software beizulegen.

DIE SOFTWARE WIRD OHNE JEDE AUSDRÜCKLICHE ODER IMPLIZIERTE GARANTIE BEREITGESTELLT, EINSCHLIESSLICH DER GARANTIE ZUR BENUTZUNG FÜR DEN VORGESEHENEN ODER EINEM BESTIMMTEN ZWECK SOWIE JEGLICHER RECHTSVERLETZUNG, JEDOCH NICHT DARAUF BESCHRÄNKT. IN KEINEM FALL SIND DIE AUTOREN ODER COPYRIGHTINHABER FÜR JEGLICHEN SCHADEN ODER SONSTIGE ANSPRÜCHE HAFTBAR ZU MACHEN, OB INFOLGE DER ERFÜLLUNG EINES VERTRAGES, EINES DELIKTES ODER ANDERS IM ZUSAMMENHANG MIT DER SOFTWARE ODER SONSTIGER VERWENDUNG DER SOFTWARE ENTSTANDEN.
""".format(VERSION)

class UeberDialog(QtWidgets.QDialog):
    def __init__(self, parent):
        super().__init__(parent)
        self._uInterface = Ueber.Ui_Dialog()
        self._uInterface.setupUi(self)
        self._uInterface.text.setPlainText(BESCHREIBUNG)

    @classmethod
    def show(cls, parent):
        dialog = cls(parent)
        dialog.open()


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

        except UIError as error:
            returnValue['success'] = False
            returnValue['errMsg'] = '{}'.format(error)

        self.signal.emit(returnValue)

class ExcelPaketWriter(QtCore.QThread):
    """Thread, um ein Excel zu speichern"""

    signal = QtCore.pyqtSignal(dict)

    def __init__(self, parent, fname, excelDaten):
        super().__init__()
        self._parent = parent
        self._fname = fname
        self._kategorien = excelDaten.getKategorien()
        self._daten = excelDaten.dataframe
        self.start()

    def run(self):
        returnValue = {'success':False}
        try:
            writePaketeToExcel(self._daten, self._kategorien, self._fname)
        except UIError as error:
            returnValue['success'] = False
            returnValue['errMsg'] = str(error)
        self.signal.emit(returnValue)

class ExcelRegelWriter(QtCore.QThread):
    """Thread, um ein Excel zu speichern"""

    signal = QtCore.pyqtSignal(dict)

    def __init__(self, parent, fname, regeln):
        super().__init__()
        self._parent = parent
        self._fname = fname
        self._regeln = regeln
        self.start()

    def run(self):
        returnValue = {'success':False}
        try:
            bedingungen = self._regeln.getBedingungsliste()
            bedingungen.to_excel(self._fname, index=False)
            returnValue['success'] = True
        except UIError as error:
            returnValue['errMsg'] = str(error)
        self.signal.emit(returnValue)

class InfoTable:
    def __init__(self):
        self._getFuncs = []
        self.model = QtGui.QStandardItemModel(0,2)

        header0 = QtGui.QStandardItem("Name")
        header1 = QtGui.QStandardItem("Wert")
        self.model.setHorizontalHeaderItem(0, header0)
        self.model.setHorizontalHeaderItem(1, header1)


    def addInfo(self, name, valueFunc):
        """Fuegt ein InfoItem hinzu"""
        self._getFuncs.append(valueFunc)

        item0 = QtGui.QStandardItem(name)
        item1 = QtGui.QStandardItem(valueFunc())
        item1.setTextAlignment(QtCore.Qt.AlignRight | QtCore.Qt.AlignVCenter)
        self.model.appendRow([item0,item1])

    def update(self):
        for i, func in enumerate(self._getFuncs):
            self.model.item(i,1).setText(str(func()))


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
        if typ < 0:
            for _,b in self._radioButtons.items():
                b.hide()
            self._uInterface.label_3.hide()
        else:
            self._radioButtons[typ].setChecked(True)

        self.ok = False
        self.setupSlots()
        self.setupListView()
        self._neueLeistung.setFocus()

    def clickOnLeistung(self, index):
        """Fuegt die angeklickte Leistung in das Textfeld ein"""
        self._neueLeistung.setText(index.data())

    def doubleClickOnLeistung(self, index):
        """Fuegt die angeklickte Leistung in das Textfeld ein und schliesst das
        Widget
        """
        self.okClicked()

    def setupListView(self):
        filterLeistung = self._neueLeistung.text()
        leistungen = self._excelDaten.getLeistungen(filterLeistung)
        model = QtGui.QStandardItemModel()
        for leistung in leistungen:
            model.appendRow(QtGui.QStandardItem(leistung))
        self._uInterface.listView_Vorschlaege.setModel(model)

    def setupSlots(self):
        uInter = self._uInterface
        uInter.buttonBox.accepted.connect(self.okClicked)
        uInter.buttonBox.rejected.connect(self.cancelClicked)
        self._neueLeistung.textEdited.connect(self.setupListView)
        uInter.listView_Vorschlaege.clicked.connect(self.clickOnLeistung)
        uInter.listView_Vorschlaege.doubleClicked.connect(self.doubleClickOnLeistung)

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

class KategorieModel(QtCore.QObject):
    neueKategorie = QtCore.pyqtSignal()
    """Schreibt die Kategorien in die Liste"""
    def __init__(self, excelDaten, listView):
        super().__init__()
        self._excelDaten = excelDaten
        self._listView = listView

        self._excelDaten.registerObserver(self)

        self._listView.installEventFilter(self)

    def update(self):
        kategorien = self._excelDaten.getKategorien()
        model = QtGui.QStandardItemModel()
        for kategorie in kategorien:
            model.appendRow(QtGui.QStandardItem(kategorie))
        self._listView.setModel(model)

    def deleteSelected(self):
        """Loescht die selektierten Kategorien"""
        liste = self._listView
        selection = liste.selectionModel().selectedIndexes()
        rows = [index.row() for index in selection]
        self._excelDaten.removeKategorien(rows)
        self.update()

    def eventFilter(self, watched, event):
        """Wird aufgerufen, wenn eine Taste gedrueckt wird"""
        if event.type() == QtCore.QEvent.KeyPress:
            if event.key() == QtCore.Qt.Key_Delete:
                self.deleteSelected()
                return True
        elif event.type() == QtCore.QEvent.ContextMenu:
            self.neueKategorie.emit()
            return True
        return False


class RegelListe(QtCore.QAbstractListModel):
    """Model der Regelliste"""
    neueRegel = QtCore.pyqtSignal()
    neueLeistung = QtCore.pyqtSignal(int)


    def __init__(self, regelListView, excelDaten, listViews):
        super().__init__()
        self._regeln = Regeln(excelDaten)
        self._regeln.registerObserver(self)
        self._bedingungsListViews = listViews
        for t, view in listViews.items():
            view.installEventFilter(self)
        
        self._regelListView = regelListView
        self._excelDaten = excelDaten

        self._redBrush = QtGui.QBrush()
        self._redBrush.setColor(QtGui.QColor(255,150,150))

        self._regelListView.setModel(self)
        self._regelListView.installEventFilter(self)
        self._regelListView.selectionModel().currentChanged.connect(
            self.selectionChanged)

    def registerObserver(self, observer):
        """Registiert einen Observer der Regeln"""
        self._regeln.registerObserver(observer)

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
        """Setzt die aktuelle Regel"""
        self._regeln.setAktiv(current.row())

    def update(self):
        """Updated die Bedingungslisten"""
        if self._regeln._aktiveRegel:
            bedingungen = self._regeln._aktiveRegel.getDict()
            for typ in [Regel.UND, Regel.ODER, Regel.NICHT]:
                model = QtGui.QStandardItemModel()
                for leistung in bedingungen[typ]:
                    item = QtGui.QStandardItem(leistung)
                    if not self._excelDaten.checkItem(leistung):
                        item.setForeground(self._redBrush)
                    model.appendRow(item)
                self._bedingungsListViews[typ].setModel(model)

    def addRegel(self, name):
        """Fuegt eine neue Regel hinzu"""
        nItems = self.rowCount()
        self.beginInsertRows(QtCore.QModelIndex(), nItems, nItems+1)
        self._regeln.addRegel(name)
        self.endInsertRows()
        self._regelListView.setCurrentIndex(self.createIndex(nItems, 0))

    def deleteSelected(self):
        """Loescht das selektierte Item"""
        row = self._regelListView.currentIndex().row()
        self.beginRemoveRows(QtCore.QModelIndex(), row, row+1)
        self._regeln.removeRegel(row)
        self.endRemoveRows()

        row = max(row,0)
        row = min(row,self.rowCount()-1)
        self._regelListView.setCurrentIndex( self.createIndex(row,0) )

    def deleteSelectedLeistungen(self, typ):
        """Loescht die selektierten Leistungen der aktiven Regel"""
        liste = self._bedingungsListViews[typ]
        selection = liste.selectionModel().selectedIndexes()
        rows = [index.row() for index in selection]
        self._regeln.removeLeistungenFromAktiverRegel(rows, typ)
        self.update()

    def regelIstAktiv(self):
        """Gibt an, ob eine Regel selektiert ist"""
        return not self._regeln.getAktiv() is None

    def eventFilter(self, watched, event):
        """Wird aufgerufen, wenn eine Taste gedrueckt wird"""
        typ = None
        for t, view in self._bedingungsListViews.items():
            if view == watched:
                typ = t
        if event.type() == QtCore.QEvent.KeyPress:
            if event.key() == QtCore.Qt.Key_Delete:
                if typ == None:
                    self.deleteSelected()
                else:
                    self.deleteSelectedLeistungen(typ)
                return True
        elif event.type() == QtCore.QEvent.ContextMenu:
            if typ == None:
                self.neueRegel.emit()
            else:
                self.neueLeistung.emit(typ)
            return True
        return False

    def clearRegeln(self):
        nItems = self.rowCount()
        self.beginRemoveRows(QtCore.QModelIndex(), 0, nItems)
        self._regeln.clearRegeln()
        self.endRemoveRows()

    def loadRegelnFromFile(self, filename):
        self.clearRegeln()

        path = pathlib.Path(filename)
        with path.with_suffix('.tpf').open('rb') as f:
            regelnDict = pickle.load(f)

        try:
            regeln = []
            for name, bedingungen in regelnDict.items():
                neueRegel = Regel(name, self._regeln._excelDaten)
                for leistung in bedingungen[Regel.UND]:
                    neueRegel.addLeistung(leistung, Regel.UND)
                for leistung in bedingungen[Regel.ODER]:
                    neueRegel.addLeistung(leistung, Regel.ODER)
                for leistung in bedingungen[Regel.NICHT]:
                    neueRegel.addLeistung(leistung, Regel.NICHT)
                regeln.append(neueRegel)

            self.beginInsertRows(QtCore.QModelIndex(), 0, len(regeln))
            self._regeln.regeln = regeln
            self.endInsertRows()
        except AttributeError:
            raise UIError("Fehler beim Laden der Regeln, ungültiges File")
        except KeyError:
            raise UIError("Fehler beim Laden der Regeln, ungültiges File")

    def saveRegelnToFile(self, fileName):
        """Speichert die Regeln in ein File"""
        self._regeln.saveToFile(fileName)

    def getErfuelltAktiveRegel(self):
        """Gibt die Anzahl der Pakete zurueck, die die aktive Regel erfuellen
        :returns: Anzahl Pakete

        """
        return self._regeln.getErfuelltAktiveRegel()

    def addLeistungToAktiverRegel(self, name, typ):
        """Fuegt der aktiven Regel eine Leistung hinzu"""
        self._regeln.addLeistungToAktiverRegel(name, typ)

    def getBedingungsliste(self):
        """Gibt die Bedingungsliste zurueck"""
        return self._regeln.getBedingungsliste()


class TarmedPaketManagerApp(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()

        self._workerThread = None

        self.uInterface = MainWindow.Ui_MainWindow()
        self.uInterface.setupUi(self)

        self._excelDaten = ExcelDaten()

        self._excelName = ''

        listViews = {
            Regel.UND : self.uInterface.listView_regel_und,
            Regel.ODER : self.uInterface.listView_regeln_oder,
            Regel.NICHT : self.uInterface.listView_regeln_nicht,
        }

        self._regelListe = RegelListe(self.uInterface.listView_regeln, 
            self._excelDaten, listViews)

        self._kategorieModel = KategorieModel(self._excelDaten, 
                self.uInterface.listView_kategorien)

        self._infoTable = InfoTable()
        self.uInterface.infoTableView.setModel(self._infoTable.model)

        self.setupSlots()
        self.setupInfoTable()



    def setupInfoTable(self):
        """Baut die InfoTable auf"""

        # Styling
        table = self.uInterface.infoTableView
        header = table.horizontalHeader()
        header.setSectionResizeMode(0, QtWidgets.QHeaderView.ResizeToContents)
        header.setSectionResizeMode(1, QtWidgets.QHeaderView.Stretch)

        # Rows
        tableInfo = self._infoTable
        tableInfo.addInfo('Aktuelles Excel', self.getExcelName)
        tableInfo.addInfo('Anzahl Falldaten',
                lambda : self._excelDaten.getAnzahlFalldaten() or '-')
        tableInfo.addInfo('Anzahl verschiedene Leistungen', 
                lambda : len(self._excelDaten.getLeistungen()) or '-')
        tableInfo.addInfo('Anzahl Falldaten in aktiver Regel', 
                self._regelListe.getErfuelltAktiveRegel)

        self._regelListe.registerObserver(tableInfo)


    def setupSlots(self):
        """Definiert die slot Funktionen der Menu Eintraege"""
        uInter = self.uInterface
        uInter.actionRohdaten_laden.triggered.connect(self.openExcel)
        uInter.actionNeue_Kategorie.triggered.connect(self.addKategorie)
        uInter.actionKategorien_l_schen.triggered.connect(self._excelDaten.clearKategorien)
        uInter.actionNeue_Regel.triggered.connect(self.addRegel)
        uInter.actionExcel_exportieren.triggered.connect(self.writeExcel)
        uInter.actionRegel_laden.triggered.connect(self.loadRegeln)
        uInter.actionRegeln_speichern.triggered.connect(self.writeRegeln)
        uInter.actionRegeln_loeschen.triggered.connect(self._regelListe.clearRegeln)
        uInter.actionNeue_Bedingung.triggered.connect(self.addLeistungToRegel)
        self._regelListe.neueLeistung.connect(self.addLeistungToRegel)
        uInter.action_Exit.triggered.connect(self.quitApp)
        uInter.actionRegelExcel_exportieren.triggered.connect(self.writeRegelExcel)
        self._regelListe.neueRegel.connect(self.addRegel)

        self._kategorieModel.neueKategorie.connect(self.addKategorie)
        uInter.action_uber.triggered.connect(self.showUeber)
        
    def showUeber(self):
        """Oeffnet den UeberDialog"""
        UeberDialog.show(self)

    def getExcelName(self):
        """Gibt den Namen des aktuellen Excels zurueck
        :returns: Name des aktuellen Excels
        """
        return self._excelName

    def addRegel(self):
        """Fragt den Benutzer nach einer neuen Regel und fuegt sie hinzu"""
        name, ok = QtWidgets.QInputDialog.getText(self, "Neue Regel",
                "Name der neuen Regel:")
        if ok:
            self._regelListe.addRegel(name)

    def addKategorie(self):
        dialog = Leistungswahldialog(self, self._excelDaten, -1)
        dialog.exec_()
        dialog.show()
        name, typ, ok = dialog.getValue()
        if ok:
            self._excelDaten.addKategorie(name)

    def writeExcel(self):
        """Schreibt die Pakete in ein Excel"""
        if self._excelDaten.dataframe is None:
            errMsg = "Keine Daten vorhanden"
            box = QtWidgets.QMessageBox.warning(self, "Warnung", errMsg,
                QtWidgets.QMessageBox.Ok,)
            return

        options = QtWidgets.QFileDialog.Options()
        options |= QtWidgets.QFileDialog.DontUseNativeDialog
        fileName, _ = QtWidgets.QFileDialog.getSaveFileName(
            self,
            "Excel speichern",
            "","Excel Files (*.xlsx *.xls)",
            options=options
        )
        if fileName:
            self._workerThread = ExcelPaketWriter(self, fileName, self._excelDaten)
            self._workerThread.signal.connect(self.finishReadExcel)
            self.disableWindow()

    def writeRegelExcel(self):
        """Schreibt die Bedingungen in ein Excel"""
        if self._excelDaten.dataframe is None:
            errMsg = "Keine Daten vorhanden"
            box = QtWidgets.QMessageBox.warning(self, "Warnung", errMsg,
                QtWidgets.QMessageBox.Ok,)
            return

        options = QtWidgets.QFileDialog.Options()
        options |= QtWidgets.QFileDialog.DontUseNativeDialog
        fileName, _ = QtWidgets.QFileDialog.getSaveFileName(
            self,
            "Excel speichern",
            "","Excel Files (*.xlsx *.xls)",
            options=options
        )
        if fileName:
            path = pathlib.Path(fileName)
            if not path.suffix in ['.xls', 'xlsx']:
                path = path.with_suffix('.xlsx')
            self._workerThread = ExcelRegelWriter(self, path, self._regelListe)
            self._workerThread.signal.connect(self.finishWrite)
            self.disableWindow()

    def finishWrite(self, result):
        """Funktion, die nach dem Schreiben einer Datei aufgerufen wird

        :result: Dict, das vom Writer zurueck gegeben wird
        """
        if not result['success']:
            errMsg = result.get('errMsg', 'Es ist ein Fehler aufgetreten')
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
            "", "TPF file (*.tpf)",
            options=options
        )
        if fileName:
            path = pathlib.Path(fileName).with_suffix('.tpf')
            self._regelListe.saveRegelnToFile(path)

    def loadRegeln(self):
        """Laedt ein Regelfile"""
        options = QtWidgets.QFileDialog.Options()
        options |= QtWidgets.QFileDialog.DontUseNativeDialog
        fileName, _ = QtWidgets.QFileDialog.getOpenFileName(
            self,
            "Regeln laden",
            "", "TPF file (*.tpf)",
            options=options
        )
        if fileName:
            try:
                self._regelListe.loadRegelnFromFile(fileName)
            except UIError as error:
                box = QtWidgets.QMessageBox.warning(self, "Warnung", 
                        str(error), QtWidgets.QMessageBox.Ok,)


    def addLeistungToRegel(self, typ=None):
        if self._regelListe.regelIstAktiv():
            dialog = Leistungswahldialog(self, self._excelDaten, typ)
            dialog.exec_()
            dialog.show()
            name, typ, ok = dialog.getValue()
            if ok:
                self._regelListe.addLeistungToAktiverRegel(name, typ)

    def quitApp(self):
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
            "","Excel oder CSV Files (*.xlsx *.xls *.csv)",
            options=options
        )
        if fileName:
            self._workerThread = ExcelReader(self, fileName)
            self._workerThread.signal.connect(self.finishReadExcel)
            self.disableWindow()
            self._excelName = pathlib.Path(fileName).stem

    def finishReadExcel(self, result):
        """ Funktion, die nach dem Lesen eines Excels aufgerufen wird

        :result: Dict mit dem Signal des Thread
        """
        if result['success']:
            self._excelDaten.dataframe = result['data'][0]
            self._excelDaten.clearKategorien()
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

        self._infoTable.update()
        self.enableWindow()

    def disableWindow(self):
        """Schaltet das Fenster in den Wartemodus"""
        QtWidgets.QApplication.setOverrideCursor(QtGui.QCursor(QtCore.Qt.WaitCursor))
        self.setEnabled(False)

    def enableWindow(self):
        """Schaltet den Wartemodus aus"""
        self.setEnabled(True)
        QtWidgets.QApplication.restoreOverrideCursor()

    # def saveRegelnToExcel(self, filePath):
        # if self.daten is None:
            # return
        # datenListe = [self.applyRegelToData(regel) for regel in self.regeln]
        # datenListe = [l.drop_duplicates(subset='FallDatum') for l in datenListe]
        # daten = pd.concat(datenListe)
        # daten.to_excel(filePath, index=False)

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


if __name__ == '__main__':
    appExists = False
    try:
        app
        appExists = True
    except NameError:
        app = QtWidgets.QApplication([])
    application = TarmedPaketManagerApp()
    application.show()
    app.exec_()
