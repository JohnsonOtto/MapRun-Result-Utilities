from PyQt5.QtWidgets import QMainWindow, QApplication, QPushButton, QFileDialog, QMessageBox
from PyQt5.QtGui import QIcon, QFont
import xml.etree.ElementTree as et
import pandas as pd
import sys, os

ns = {"ol": "http://www.orienteering.org/datastandard/3.0"}
et.register_namespace("", ns["ol"])

def resourcePath(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

def secs2hrs(seconds):
    mins, secs = divmod(seconds, 60)
    hrs, mins = divmod(mins, 60)
    return '{:02d}:{:02d}:{:02}'.format(hrs, mins, secs)

def gatherResults(xmlFile):
    global root
    resultList = []
    for course in root.findall('./ol:ClassResult/ol:Class/ol:ShortName',namespaces=ns):
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
            organisation = person.find("./ol:Organisation/ol:Name", namespaces=ns).text
        except AttributeError:
            organisation = "kein Verein"
        resultList.append([courseName, position, lastname, surname, organisation, secs2hrs(int(time))]) 
    return resultList

def xml2excel(filelist):
    global root
    writer = pd.ExcelWriter("results.xlsx", engine="xlsxwriter")

    singles = []
    whole = []

    for filename in filelist:
        tree = et.parse(filename)
        root = tree.getroot()
        gathered = gatherResults(filename)
        singles.append(gathered)
        whole += gathered + [""]

    wholeDF = pd.DataFrame(whole)
    wholeDF.columns = ["Bahn", "Platzierung", "Name", "Vorname", "Verein", "Zeit"]
    wholeDF.to_excel(writer, sheet_name="Alle Bahnen", index=False)

    for i in range(len(singles)):
        singleDF = pd.DataFrame(singles[i])
        singleDF.columns = ["Bahn", "Platzierung", "Name", "Vorname", "Verein", "Zeit"]
        singleDF.to_excel(writer, sheet_name=singles[i][0][0], index=False)

    writer.save()

    msg = QMessageBox(win)
    msg.setStandardButtons(QMessageBox.Ok)
    msg.setWindowTitle("Fertig!")
    msg.setText("Erstellung erfolgreich!")
    msg.setFont(QFont("Arial", 10))
    msg.exec()

class window(QMainWindow):
    
    def __init__(self):
        super().__init__()

        self.setGeometry(100,100,270,180)
        self.setWindowTitle(" ")
        self.setWindowIcon(QIcon(resourcePath("icon.png")))
        self.initUI()

    def initUI(self):
        self.button1 = QPushButton(self)
        self.button1.setGeometry(10,10,250,75)
        self.button1.setText("Datei(en) auswählen")
        self.button1.setFont(QFont("Arial", 11))
        self.button1.clicked.connect(self.button1Clicked)

        self.button2 = QPushButton(self)
        self.button2.setGeometry(10,95,250,75)
        self.button2.setText("Schließen")
        self.button2.setFont(QFont("Arial", 11))
        self.button2.clicked.connect(self.close)

        self.show()

    def openFileNamesDialog(self):
        fileName, _ = QFileDialog.getOpenFileNames(self,"Select .xml file(s)", "","Extensible Markup Language (*.xml)")
        if fileName:
            return fileName

    def button1Clicked(self):
        fileList = self.openFileNamesDialog()
        if not fileList == None:
            xml2excel(fileList)
    
if __name__ == '__main__':
    app = QApplication(sys.argv)
    win = window()
    sys.exit(app.exec_())