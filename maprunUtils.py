from PyQt5.QtWidgets import (
    QMainWindow,
    QApplication,
    QPushButton,
    QFileDialog,
    QMessageBox,
    QTableView,
    QTableWidgetItem,
    QMenuBar,
    QAction,
    QMenu,
    QAbstractScrollArea,
)
from PyQt5.QtGui import QIcon, QFont
from PyQt5 import QtCore
from PyQt5.QtCore import Qt
import xml.etree.ElementTree as et
import pandas as pd
import sys, os, csv


def resourcePath(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


def secs2hrs(seconds):
    mins, secs = divmod(seconds, 60)
    hrs, mins = divmod(mins, 60)
    return "{:02d}:{:02d}:{:02}".format(hrs, mins, secs)


def time2sec(timestr):
    groups = [0, 0, 0] + timestr.split(":")
    s, m, h = groups[-1], groups[-2], groups[-3]
    return int(h) * 3600 + int(m) * 60 + int(s)


def getXMLResults(xmlfilelist, header):
    results = []
    ns = {"ol": "http://www.orienteering.org/datastandard/3.0"}
    et.register_namespace("", ns["ol"])

    for filename in xmlfilelist:
        gathered = []
        tree = et.parse(filename)
        root = tree.getroot()

        for course in root.findall(
            "./ol:ClassResult/ol:Class/ol:ShortName", namespaces=ns
        ):
            courseName = course.text
        for person in root.findall("./ol:ClassResult/ol:PersonResult", namespaces=ns):
            surname = person.find("./ol:Person/ol:Name/ol:Given", namespaces=ns).text
            lastname = person.find("./ol:Person/ol:Name/ol:Family", namespaces=ns).text
            time = person.find("./ol:Result/ol:Time", namespaces=ns).text
            status = person.find("./ol:Result/ol:Status", namespaces=ns).text
            if status == "OK":
                position = person.find("./ol:Result/ol:Position", namespaces=ns).text
            else:
                position = "MP"
            try:
                organisation = person.find(
                    "./ol:Organisation/ol:Name", namespaces=ns
                ).text
            except AttributeError:
                organisation = "-"
            gathered.append(
                [
                    courseName,
                    position,
                    lastname,
                    surname,
                    organisation,
                    secs2hrs(int(time)),
                ]
            )

        results += gathered + [["", "", "", "", "", ""]]
    del results[-1]

    results = pd.DataFrame(results)
    results.drop_duplicates(inplace=True)
    results.columns = header
    return results


def getCSVResults(csvfilelist, header):
    results = []
    for filename in csvfilelist:
        gathered = []
        rest = []
        with open(filename, "r", encoding="utf8") as f:
            csvRaw = list(csv.reader(f))
            f.close()
            csvRaw = csvRaw[1:]
        for i in range(len(csvRaw)):
            if csvRaw[i][12] == "0":
                gathered.append(
                    [
                        "nicht in Datei",
                        csvRaw[i][3],
                        csvRaw[i][4],
                        csvRaw[i][14],
                        csvRaw[i][24],
                        time2sec(csvRaw[i][11]),
                    ]
                )

            else:
                rest.append(
                    [
                        "nicht in Datei",
                        "MP",
                        csvRaw[i][3],
                        csvRaw[i][4],
                        csvRaw[i][14],
                        csvRaw[i][24],
                        secs2hrs(time2sec(csvRaw[i][11])),
                    ]
                )
        gathered.sort(key=lambda i: i[5])
        for i in range(len(gathered)):
            gathered[i][5] = secs2hrs(gathered[i][5])
            gathered[i].insert(1, (i + 1))

        results += gathered + rest + [["", "", "", "", "", ""]]
    del results[-1]

    results = pd.DataFrame(results)
    results.drop_duplicates(inplace=True)
    results.columns = header
    return results


def data2xlsx(dataFrame, filename, severalSheets=True):
    writer = pd.ExcelWriter(filename, engine="xlsxwriter")
    dataFrame.to_excel(
        writer, sheet_name=language[win.standardLang]["allcourses"], index=False
    )

    if severalSheets:
        breakRows = []
        singleDFs = []

        for row in dataFrame.itertuples():
            if list(row)[1:] == ["", "", "", "", "", ""]:
                breakRows.append(row.Index)
        noSpaces = dataFrame.drop(breakRows)
        breakRows = [-1] + breakRows + [len(dataFrame)]

        for i in range(len(breakRows) - 1):
            single = dataFrame.loc[breakRows[i] + 1 : breakRows[i + 1] - 1]
            name = single[language[win.standardLang]["course"]].iloc[0]
            print(name)
            single.to_excel(writer, sheet_name=name, index=False)

    writer.save()


class PandasModel(
    QtCore.QAbstractTableModel
):  # https://github.com/eyllanesc/stackoverflow/blob/master/questions/44603119/PandasModel.py
    def __init__(self, df=pd.DataFrame(), parent=None):
        QtCore.QAbstractTableModel.__init__(self, parent=parent)
        self._df = df

    def headerData(self, section, orientation, role=QtCore.Qt.DisplayRole):
        if role != QtCore.Qt.DisplayRole:
            return QtCore.QVariant()

        if orientation == QtCore.Qt.Horizontal:
            try:
                return self._df.columns.tolist()[section]
            except (IndexError,):
                return QtCore.QVariant()
        elif orientation == QtCore.Qt.Vertical:
            try:
                # return self.df.index.tolist()
                return self._df.index.tolist()[section]
            except (IndexError,):
                return QtCore.QVariant()

    def data(self, index, role=QtCore.Qt.DisplayRole):
        if role != QtCore.Qt.DisplayRole:
            return QtCore.QVariant()

        if not index.isValid():
            return QtCore.QVariant()

        return QtCore.QVariant(str(self._df.iloc[index.row(), index.column()]))

    def setData(self, index, value, role):
        row = self._df.index[index.row()]
        col = self._df.columns[index.column()]
        if hasattr(value, "toPyObject"):
            # PyQt4 gets a QVariant
            value = value.toPyObject()
        else:
            # PySide gets an unicode
            dtype = self._df[col].dtype
            if dtype != object:
                value = None if value == "" else dtype.type(value)
        self._df.set_value(row, col, value)
        return True

    def rowCount(self, parent=QtCore.QModelIndex()):
        return len(self._df.index)

    def columnCount(self, parent=QtCore.QModelIndex()):
        return len(self._df.columns)

    def sort(self, column, order):
        colname = self._df.columns.tolist()[column]
        self.layoutAboutToBeChanged.emit()
        self._df.sort_values(
            colname, ascending=order == QtCore.Qt.AscendingOrder, inplace=True
        )
        self._df.reset_index(inplace=True, drop=True)
        self.layoutChanged.emit()


class window(QMainWindow):
    def __init__(self):
        super().__init__()

        # self.setGeometry(100,100,270,265)
        self.fileList = None
        self.standardLang = "german"
        self.severalSheets = True

        self.xmlheader = [
            language[self.standardLang]["course"],
            language[self.standardLang]["placement"],
            language[self.standardLang]["lastname"],
            language[self.standardLang]["surname"],
            language[self.standardLang]["organisation"],
            language[self.standardLang]["time"],
        ]
        self.csvheader = [
            language[self.standardLang]["course"],
            language[self.standardLang]["placement"],
            language[self.standardLang]["lastname"],
            language[self.standardLang]["surname"],
            language[self.standardLang]["organisation"],
            language[self.standardLang]["rundate"],
            language[self.standardLang]["time"],
        ]
        self.initUI()

    def initUI(self):
        self.setGeometry(0, 0, 750, 750)
        self.setWindowTitle("MapRun Result Utilities")
        self.setWindowIcon(QIcon(resourcePath("icon.png")))

        self.fileMenuOpenXML = QAction(
            "&" + language[self.standardLang]["importXML"], self
        )
        self.fileMenuOpenXML.setStatusTip(language[self.standardLang]["importXMLdesc"])
        self.fileMenuOpenXML.triggered.connect(self.openXMLFilesAction)
        self.fileMenuOpenXML.setShortcut("Ctrl+O")

        self.fileMenuOpenCSV = QAction(
            "&" + language[self.standardLang]["importCSV"], self
        )
        self.fileMenuOpenCSV.setStatusTip(language[self.standardLang]["importCSVdesc"])
        self.fileMenuOpenCSV.triggered.connect(self.openCSVFilesAction)
        self.fileMenuOpenCSV.setShortcut("Ctrl+P")

        self.fileMenuSaveAs = QAction("&" + language[self.standardLang]["save"], self)
        self.fileMenuSaveAs.setStatusTip(language[self.standardLang]["savedesc"])
        self.fileMenuSaveAs.triggered.connect(self.saveXLSXAction)
        self.fileMenuSaveAs.setShortcut("Ctrl+S")

        self.fileMenuExit = QAction("&" + language[self.standardLang]["exit"], self)
        self.fileMenuExit.setStatusTip(language[self.standardLang]["exitdesc"])
        self.fileMenuExit.triggered.connect(self.close)
        self.fileMenuExit.setShortcut("Ctrl+Q")

        self.englishAction = QAction("&English")
        self.englishAction.triggered.connect(lambda: self.changeLanguage("english"))

        self.germanAction = QAction("&German")
        self.germanAction.triggered.connect(lambda: self.changeLanguage("german"))

        self.helpSubMenuLanguage = QMenu("&Change Language", self)
        self.helpSubMenuLanguage.addAction(self.englishAction)
        self.helpSubMenuLanguage.addAction(self.germanAction)

        self.editMenuPlacementAction = QAction(
            "&" + language[self.standardLang]["placement"], self, checkable=True
        )
        self.editMenuPlacementAction.triggered.connect(self.placementSubmenuAction)
        self.editMenuPlacementAction.setChecked(True)

        self.editMenuSurnameAction = QAction(
            "&" + language[self.standardLang]["surname"], self, checkable=True
        )
        self.editMenuSurnameAction.triggered.connect(self.surnameSubmenuAction)
        self.editMenuSurnameAction.setChecked(True)

        self.editMenuLastnameAction = QAction(
            "&" + language[self.standardLang]["lastname"], self, checkable=True
        )
        self.editMenuLastnameAction.triggered.connect(self.lastnameSubmenuAction)
        self.editMenuLastnameAction.setChecked(True)

        self.editMenuTimeAction = QAction(
            "&" + language[self.standardLang]["time"], self, checkable=True
        )
        self.editMenuTimeAction.triggered.connect(self.timeSubmenuAction)
        self.editMenuTimeAction.setChecked(True)

        self.editMenuOrganisationAction = QAction(
            "&" + language[self.standardLang]["organisation"], self, checkable=True
        )
        self.editMenuOrganisationAction.triggered.connect(
            self.organisationSubmenuAction
        )
        self.editMenuOrganisationAction.setChecked(True)

        self.editMenuCourseAction = QAction(
            "&" + language[self.standardLang]["course"], self, checkable=True
        )
        self.editMenuCourseAction.triggered.connect(self.courseSubmenuAction)
        self.editMenuCourseAction.setChecked(True)

        self.editMenuDateAction = QAction(
            "&" + language[self.standardLang]["rundate"], self, checkable=True
        )
        self.editMenuDateAction.triggered.connect(self.rundateSubmenuAction)
        self.editMenuDateAction.setChecked(True)

        self.editMenuSplittimeAction = QAction(
            "&" + language[self.standardLang]["splittime"], self, checkable=True
        )
        self.editMenuSplittimeAction.triggered.connect(self.splittimeSubmenuAction)
        self.editMenuSplittimeAction.setChecked(False)

        self.helpMenuAbout = QAction("&" + language[self.standardLang]["about"], self)
        self.helpMenuAbout.triggered.connect(self.about)

        self.editMenuSubSelect = QMenu(
            "&" + language[self.standardLang]["columnSelect"], self
        )
        self.editMenuSubSelect.addAction(self.editMenuPlacementAction)
        self.editMenuSubSelect.addAction(self.editMenuSurnameAction)
        self.editMenuSubSelect.addAction(self.editMenuLastnameAction)
        self.editMenuSubSelect.addAction(self.editMenuTimeAction)
        self.editMenuSubSelect.addAction(self.editMenuOrganisationAction)
        self.editMenuSubSelect.addAction(self.editMenuCourseAction)
        self.editMenuSubSelect.addAction(self.editMenuDateAction)
        self.editMenuSubSelect.addAction(self.editMenuSplittimeAction)

        self.menu = self.menuBar()
        self.fileMenu = self.menu.addMenu("&" + language[self.standardLang]["file"])
        self.fileMenu.addAction(self.fileMenuOpenXML)
        self.fileMenu.addAction(self.fileMenuOpenCSV)
        self.fileMenu.addAction(self.fileMenuSaveAs)
        self.fileMenu.addAction(self.fileMenuExit)
        self.fileMenu.addSeparator()

        self.editMenu = self.menu.addMenu("&" + language[self.standardLang]["edit"])
        self.editMenu.addMenu(self.editMenuSubSelect)

        self.helpMenu = self.menu.addMenu("&" + language[self.standardLang]["help"])
        self.helpMenu.addMenu(self.helpSubMenuLanguage)
        self.helpMenu.addAction(self.helpMenuAbout)

        self.debugMenuAction = QAction("&Debug", self)
        self.debugMenuAction.triggered.connect(self.button3Clicked)

        self.debugMenu = self.menu.addMenu("&Debug")
        self.debugMenu.addAction(self.debugMenuAction)

        self.table = QTableView(self)
        self.table.setGeometry(
            3, 28, self.geometry().width() - 6, self.geometry().height() - 32
        )
        self.table.verticalHeader().setVisible(False)

        self.vertTableHeaders = self.table.horizontalHeader()
        self.vertTableHeaders.setContextMenuPolicy(Qt.CustomContextMenu)
        self.vertTableHeaders.customContextMenuRequested.connect(
            self.tableContextMenuRequest
        )

        self.tableMenuMoveRightAction = QAction(
            "&" + language[self.standardLang]["moveright"]
        )
        # self.tableMenuMoveRightAction.triggered.connect()

        self.tableMenuMoveLeftAction = QAction(
            "&" + language[self.standardLang]["moveleft"]
        )

        self.tableContextMenu = QMenu()
        self.tableContextMenu.addAction(self.tableMenuMoveLeftAction)
        self.tableContextMenu.addAction(self.tableMenuMoveRightAction)

        # self.showMaximized()
        self.show()

    def tableContextMenuRequest(self, point):
        self.row = self.table.rowAt(point.y())
        self.col = self.table.columnAt(point.x())

        # print(self.cell.text)
        self.tableContextMenu.exec_(self.table.viewport().mapToGlobal(point))
        return

    def resizeEvent(self, event):
        self.table.setGeometry(
            3, 28, self.geometry().width() - 6, self.geometry().height() - 32
        )

    def openFileNamesDialog(self, filetype, title):
        fileName, _ = QFileDialog.getOpenFileNames(self, title, "", filetype)
        if fileName:
            return fileName

    def saveFileNameDialog(self, filetype, title):
        filename, _ = QFileDialog.getSaveFileName(self, title, "", filetype)
        if filename:
            return filename

    def showMessageBox(self, title, text):
        self.msg = QMessageBox(self)
        self.msg.setWindowTitle(title)
        self.msg.setText(text)
        self.msg.setFont(QFont("Arial", 10))
        self.msg.exec()

    def openXMLFilesAction(self):
        self.fileList = self.openFileNamesDialog(
            "Extensible Markup Language (*.xml)", language[self.standardLang]["select"]
        )
        if self.fileList == None:
            return
        else:
            self.wholeData = getXMLResults(self.fileList, self.xmlheader)
            self.wholeData.columns = [
                language[self.standardLang]["course"],
                language[self.standardLang]["placement"],
                language[self.standardLang]["lastname"],
                language[self.standardLang]["surname"],
                language[self.standardLang]["organisation"],
                language[self.standardLang]["time"],
            ]
            self.tableData = self.wholeData
            self.updateTable(self.tableData)

    def openCSVFilesAction(self):
        self.fileList = self.openFileNamesDialog(
            "Comma Separated Values (*.csv)", language[self.standardLang]["select"]
        )
        if self.fileList == None:
            return
        else:
            self.wholeData = getCSVResults(self.fileList, self.csvheader)
            self.tableData = self.wholeData
            self.updateTable(self.tableData)

    def saveXLSXAction(self):
        self.fileName = self.saveFileNameDialog(
            "Excel (*.xlsx)", language[self.standardLang]["save"]
        )
        if self.fileName == None:
            return

        if self.fileList == None:
            self.showMessageBox(
                language[self.standardLang]["error"],
                language[self.standardLang]["noFiles"],
            )
        else:
            if not self.fileName == None:
                # xml2excel(self.fileList, self.fileName)
                data2xlsx(self.tableData, self.fileName, self.severalSheets)
                self.showMessageBox(
                    language[self.standardLang]["done"],
                    language[self.standardLang]["exportSuccess"],
                )

    def updateTable(self, data):
        self.model = PandasModel(data)
        self.table.setModel(self.model)

    def swapLeft(dataFrame, index):
        return

    def swapRight(dataFrame, index):
        return

    def placementSubmenuAction(self):
        print("placement")

    def surnameSubmenuAction(self):
        print("surname")

    def lastnameSubmenuAction(self):
        print("lastname")

    def timeSubmenuAction(self):
        print("time")

    def organisationSubmenuAction(self):
        print("organisation")

    def courseSubmenuAction(self):
        print("course")

    def rundateSubmenuAction(self):
        print("rundate")

    def splittimeSubmenuAction(self):
        print("splittime")

    def about(self):
        print("about")

    def changeLanguage(self, lang):
        self.standardLang = lang
        self.fileMenuOpenXML.setText("&" + language[lang]["importXML"])
        self.fileMenuSaveAs.setText("&" + language[lang]["save"])
        self.fileMenuExit.setText("&" + language[lang]["exit"])
        self.editMenuPlacementAction.setText("&" + language[lang]["placement"])
        self.editMenuSurnameAction.setText("&" + language[lang]["surname"])
        self.editMenuLastnameAction.setText("&" + language[lang]["lastname"])
        self.editMenuTimeAction.setText("&" + language[lang]["time"])
        self.editMenuOrganisationAction.setText("&" + language[lang]["organisation"])
        self.editMenuCourseAction.setText("&" + language[lang]["course"])
        self.editMenuDateAction.setText("&" + language[lang]["rundate"])
        self.editMenuSplittimeAction.setText("&" + language[lang]["splittime"])
        self.helpMenuAbout.setText("&" + language[lang]["about"])
        self.editMenuSubSelect.setTitle("&" + language[lang]["columnSelect"])
        self.fileMenu.setTitle("&" + language[lang]["file"])
        self.editMenu.setTitle("&" + language[lang]["edit"])
        self.helpMenu.setTitle("&" + language[lang]["help"])
        self.tableMenuMoveLeftAction.setText("&" + language[lang]["moveleft"])
        # add translate table + data headers (difference between csv and xml!)

    def button3Clicked(self):
        print("debug")
        # self.whole = getXMLResults(["LS.xml", "ML.xml"])
        # self.whole.columns = [language[self.standardLang]["course"], language[self.standardLang]["placement"], language[self.standardLang]["lastname"], language[self.standardLang]["surname"], language[self.standardLang]["organisation"], language[self.standardLang]["time"]]
        # self.tableData = self.whole
        # self.whole
        # self.updateTable(self.tableData)


if __name__ == "__main__":

    language = {
        "german": {
            "file": "Datei",
            "edit": "Werkzeuge",
            "help": "Hilfe",
            "save": "Speichern",
            "savedesc": ".xlsx Datei exportieren",
            "exit": "Programm schließen",
            "exitdesc": "Programm schließen",
            "placement": "Platzierung",
            "surname": "Vorname",
            "lastname": "Nachname",
            "time": "Zeit",
            "organisation": "Verein",
            "course": "Bahn",
            "splittime": "Zwischenzeiten",
            "importXML": ".xml importieren",
            "importXMLdesc": "Ergebnisse von .xml Dateien laden",
            "importCSV": ".csv importieren",
            "importCSVdesc": "Ergebnisse von .csv Dateien laden",
            "columnSelect": "Spalten wählen",
            "select": "Datei auswählen",
            "error": "Fehler",
            "noFiles": "Keine Dateien importiert!",
            "done": "Fertig",
            "exportSuccess": "Datei erfolgreich gespeichert!",
            "about": "Über",
            "moveright": "rechts",
            "moveleft": "links",
            "allcourses": "Alle Bahnen",
            "rundate": "Laufdatum",
        },
        "english": {
            "file": "File",
            "edit": "Tools",
            "help": "Help",
            "save": "Save",
            "savedesc": "Export .xlsx file",
            "exit": "Exit application",
            "exitdesc": "Exit application",
            "placement": "Placement",
            "surname": "Surname",
            "lastname": "Lastname",
            "time": "Time",
            "course": "Course",
            "splittime": "Split Times",
            "organisation": "Organisation",
            "importXML": "Import .xml",
            "importXMLdesc": "Load results from .xml file",
            "importCSV": "Import .csv",
            "importCSVdesc": "Load results from .csv file",
            "columnSelect": "Select columns",
            "select": "Select file",
            "error": "Error",
            "noFiles": "No files imported!",
            "done": "Success",
            "exportSuccess": "Exported file successfully!",
            "about": "About",
            "moveright": "right",
            "moveleft": "left",
            "allcourses": "Combined",
            "rundate": "Date",
        },
    }

    app = QApplication(sys.argv)
    win = window()
    sys.exit(app.exec_())