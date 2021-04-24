import xml.etree.ElementTree as et
import pandas as pd
import os


def getXMLResultList(xmlfilelist):
    results = []
    ns = {"ol": "http://www.orienteering.org/datastandard/3.0"}
    et.register_namespace("", ns["ol"])
    for filename in xmlfilelist:
        gathered = []
        tree = et.parse(filename)
        root = tree.getroot()
        for event in root.findall("./ol:Event/ol:Name", namespaces=ns):
            eventName = event.text
            eventName = formatEventName(eventName)
        for course in root.findall(
            "./ol:ClassResult/ol:Class/ol:ShortName", namespaces=ns
        ):
            courseName = course.text
        for person in root.findall("./ol:ClassResult/ol:PersonResult", namespaces=ns):
            status = person.find("./ol:Result/ol:Status", namespaces=ns).text
            if status == "OK":
                position = person.find("./ol:Result/ol:Position", namespaces=ns).text
                surname = person.find(
                    "./ol:Person/ol:Name/ol:Given", namespaces=ns
                ).text
                lastname = person.find(
                    "./ol:Person/ol:Name/ol:Family", namespaces=ns
                ).text
                time = person.find("./ol:Result/ol:Time", namespaces=ns).text
                try:
                    organisation = person.find(
                        "./ol:Organisation/ol:Name", namespaces=ns
                    ).text
                except AttributeError:
                    organisation = "-"
                gathered.append(
                    [
                        eventName,
                        courseName,
                        position,
                        lastname,
                        surname,
                        organisation,
                        int(time),
                    ]
                )
        results.append(gathered)
    return results


def courseScore(resultlist):
    for i in range(len(resultlist)):
        currentResultList = resultlist[i]
        bestTime = resultlist[i][0][6]
        for j in range(len(resultlist[i])):
            currentPerson = resultlist[i][j]
            score = bestTime / resultlist[i][j][6] * 100
            score = round(score, 2)
            resultlist[i][j].append(score)
    return resultlist


def courseResults(procdata):
    out = []
    deletable = procdata
    while len(deletable) > 0:
        name = deletable[0][4] + " " + deletable[0][3]
        toappend = [
            deletable[0][1],
            deletable[0][3],
            deletable[0][4],
            deletable[0][5],
            [deletable[0][0], deletable[0][7]],
        ]
        event = deletable[0][0]
        for i in range(1, len(deletable) - 1):
            if name == deletable[i][4] + " " + deletable[i][3]:
                if event == deletable[i][0]:
                    del deletable[i]
                else:
                    toappend.append([deletable[i][0], deletable[i][7]])
                    del deletable[i]
        # print(toappend)
        out.append(toappend)
        del deletable[0]
    # print("")
    return out


def sortEventAddScore(courseresult, order):
    out = []
    for i in range(len(courseresult)):
        toappend = [
            courseresult[i][0],
            courseresult[i][1],
            courseresult[i][2],
            courseresult[i][3],
        ]
        for event in order:
            check = False
            for j in range(4, len(courseresult[i])):
                if courseresult[i][j][0] == event:
                    toappend.append(courseresult[i][j][1])
                    check = True
            if not check:
                toappend.append(0.0)

        scores = toappend[4 : 4 + len(order)]
        score = sum(findNmax(scores, 4))
        toappend.append(score)

        out.append(toappend)
    return out


def formatEventName(event):
    groups = event.split(" ")
    return " ".join(groups[len(groups) - 3 : len(groups) - 2])


def time2sec(timestr):
    h, m, s = timestr.split(":")
    return int(h) * 3600 + int(m) * 60 + int(s)


def secs2hrs(seconds):
    mins, secs = divmod(int(seconds), 60)
    hrs, mins = divmod(mins, 60)
    return "{:02d}:{:02d}:{:02}".format(hrs, mins, secs)


def findNmax(array, n):
    if n > len(array):
        return array
    tmp = array
    out = []
    for i in range(n):
        m = max(tmp)
        tmp.remove(m)
        out.append(m)
    return out


if __name__ == "__main__":
    fileList = os.listdir("input-xml")
    eventOrder = ["MOL2021", "USC2021"]
    eventNames = ["Magdeburger OL", "Ottos Wald-OL"]

    dataFrameHeaders = ["Bahn", "Name", "Vorname", "Verein"]
    for event in eventNames:
        dataFrameHeaders.append(event)
    dataFrameHeaders.append("Gesamt")

    for i in range(len(fileList)):
        fileList[i] = "input-xml/" + fileList[i]

    resultList = getXMLResultList(fileList)
    resultList = courseScore(resultList)

    kl = []
    ks = []
    ml = []
    ms = []
    ll = []
    ls = []

    for i in range(len(resultList)):
        for j in range(len(resultList[i])):
            if resultList[i][j][1] == "KL":
                kl.append(resultList[i][j])
            elif resultList[i][j][1] == "KS":
                ks.append(resultList[i][j])
            elif resultList[i][j][1] == "ML":
                ml.append(resultList[i][j])
            elif resultList[i][j][1] == "MS":
                ms.append(resultList[i][j])
            elif resultList[i][j][1] == "LL":
                ll.append(resultList[i][j])
            elif resultList[i][j][1] == "LS":
                ls.append(resultList[i][j])

    # processing data
    kl = courseResults(kl)
    kl = sortEventAddScore(kl, eventOrder)
    kl = pd.DataFrame(kl)
    # kl.columns = dataFrameHeaders
    # kl = kl.sort_values(by="Gesamt", ascending=False)
    # kl.insert(loc=0, column="Platzierung", value=range(1, len(kl) + 1))

    ks = courseResults(ks)
    # ks = sortEventAddScore(ks, eventOrder)
    ks = pd.DataFrame(ks)
    print(ks)

    # print(dataFrameHeaders)
    # writer = pd.ExcelWriter("regiocup/Auswertung Regiocup.xlsx", engine="xlsxwriter")
    # ks.to_excel(writer, index=False, header=False)
    # writer.save()

    print("Done! Press any button to exit...")
    # input()
