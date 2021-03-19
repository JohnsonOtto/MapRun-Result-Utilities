import xml.etree.ElementTree as et
import pandas as pd

def secs2hrs(seconds):
    mins, secs = divmod(seconds, 60)
    hrs, mins = divmod(mins, 60)
    return '{:02d}:{:02d}:{:02}'.format(hrs, mins, secs)

def gatherResults(xmlFile):
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

ns = {"ol": "http://www.orienteering.org/datastandard/3.0"}
et.register_namespace("", ns["ol"])

try:
    xmlFiles = input("Enter filenames: ").split()

    singles = []
    whole = []

    for filename in xmlFiles:
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

    #print(df)
    print("Conversion completed successfully!")
    input("Press any key to close program.")
except FileNotFoundError:
    print("File not found!")
    input("Press any key to close program.")