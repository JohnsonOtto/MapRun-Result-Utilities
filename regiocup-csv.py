import csv
import pandas as pd
import os
import datetime as dt
from time import perf_counter


def readCSV(path):
    with open(path, "r", encoding="utf8") as f:
        tmp = list(csv.reader(f))
        f.close()
        return tmp[1:]


def formatEventCourse(filename):
    groups = filename.split(" ")
    course = groups[1][:2]
    groups = groups[0].split("/")
    event = groups[2]
    return event, course


def formatDate(datestr):
    day, month, year = datestr.split("-")
    return dt.date(int(year), int(month), int(day)).timetuple().tm_yday


def time2sec(timestr):
    groups = [0, 0, 0] + timestr.split(":")
    s, m, h = groups[-1], groups[-2], groups[-3]
    return int(h) * 3600 + int(m) * 60 + int(s)


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


# read files; filter mp, time = 0, not in timeframe; sort for first run
def readFilesCleanData(files):
    results = []
    rest = []
    out = []
    for f in files:
        event, course = formatEventCourse(f)
        tmp = readCSV(f)
        for i in range(len(tmp)):
            # print(tmp[i])
            personDict = {
                "info": {
                    "surname": tmp[i][4],
                    "lastname": tmp[i][3],
                    "organisation": tmp[i][14],
                },
                "run": {
                    "event": event,
                    "course": course,
                    "status": tmp[i][12],
                    "starttime": 24 * 3600 * formatDate(tmp[i][24])
                    + time2sec(tmp[i][9]),
                    "startday": tmp[i][24],
                    "time": tmp[i][11],
                    "note": "",
                },
            }
            if personDict in results:
                continue
            if personDict["run"]["time"] == "00:00":
                continue

            eventDict = list(
                filter(lambda a: a["shortname"] == personDict["run"]["event"], events)
            )
            firstday = 24 * 3600 * formatDate(eventDict[0]["firstday"])
            lastday = 24 * 3600 * (formatDate(eventDict[0]["lastday"]) + 1)
            status = personDict["run"]["status"]
            if not firstday < personDict["run"]["starttime"] < lastday:
                personDict["run"]["score"] = 0.0
                rest.append(personDict)
            elif status == "0":
                results.append(personDict)
            elif status == "3":
                personDict["run"]["score"] = 0.0
                personDict["run"]["note"] = "Fehlstempel"
                rest.append(personDict)
            elif status == "4":
                personDict["run"]["score"] = 0.0
                personDict["run"]["note"] = "a.K."
                rest.append(personDict)
            else:
                personDict["run"]["score"] = 0.0
                rest.append(personDict)

    while len(results) > 0:
        currentPerson = [results[0]]
        identifier = (
            results[0]["info"]["surname"]
            + results[0]["info"]["lastname"]
            + results[0]["info"]["organisation"]
            + results[0]["run"]["event"]
        )
        for j in range(1, len(results)):
            jdentifier = (
                results[j]["info"]["surname"]
                + results[j]["info"]["lastname"]
                + results[j]["info"]["organisation"]
                + results[j]["run"]["event"]
            )
            if identifier == jdentifier:
                currentPerson.append(results[j])

        if len(currentPerson) == 1:
            out.append(currentPerson[0])
            del results[0]
        else:
            earliest = 24 * 3600 * formatDate("31-12-4444") + time2sec("23:59:59")
            toappend = "hee"
            for person in currentPerson:
                if person["run"]["starttime"] < earliest:
                    if not toappend == "hee":
                        results.remove(toappend)
                        toappend["run"]["score"] = 0.0
                        toappend["run"]["note"] = "Zweitlauf"
                        rest.append(toappend)
                    earliest = person["run"]["starttime"]
                    toappend = person
                else:
                    results.remove(person)
                    person["run"]["score"] = 0.0
                    person["run"]["note"] = "Zweitlauf"
                    rest.append(person)
            out.append(toappend)
            results.remove(toappend)
    return out, rest


# add score
def calcScore(cup):
    out = []
    for event in events:
        course = []
        bestTime = time2sec("24:00:00")

        for person in cup:
            if event["shortname"] == person["run"]["event"]:
                course.append(person)
                # cup.remove(person)

                if time2sec(person["run"]["time"]) < bestTime:
                    bestTime = time2sec(person["run"]["time"])

        for person in course:
            score = bestTime / time2sec(person["run"]["time"]) * 100
            person["run"]["score"] = round(score, 2)
        out += course

    return out


# for results and rest, returns df
def compactPersons(courseresult):
    compacted = []
    while len(courseresult) > 0:
        toappend = [
            courseresult[0]["run"]["course"],
            courseresult[0]["info"]["lastname"],
            courseresult[0]["info"]["surname"],
            courseresult[0]["info"]["organisation"],
            [
                courseresult[0]["run"]["event"],
                courseresult[0]["run"]["startday"],
                courseresult[0]["run"]["time"],
                courseresult[0]["run"]["score"],
                courseresult[0]["run"]["note"],
            ],
        ]

        i = 1
        while True:
            if i >= len(courseresult) or len(courseresult) == 1:
                break
            if courseresult[0]["info"] == courseresult[i]["info"]:
                toappend.append(
                    [
                        courseresult[i]["run"]["event"],
                        courseresult[i]["run"]["startday"],
                        courseresult[i]["run"]["time"],
                        courseresult[i]["run"]["score"],
                        courseresult[i]["run"]["note"],
                    ]
                )
                del courseresult[i]
            else:
                i += 1

        compacted.append(toappend)
        del courseresult[0]

    out = []
    for i in range(len(compacted)):
        toappend = [compacted[i][0], compacted[i][1], compacted[i][2], compacted[i][3]]
        misPunch = False
        # print(toappend)
        for event in events:
            check = False
            for j in range(4, len(compacted[i])):
                if compacted[i][j][0] == event["shortname"]:
                    check = True
                    if compacted[i][j][4] == "Fehlstempel":
                        misPunch = True
                        toappend += [
                            compacted[i][j][1],
                            compacted[i][j][2],
                            "0",
                            compacted[i][j][4],
                        ]
                    else:
                        toappend += [
                            compacted[i][j][1],
                            compacted[i][j][2],
                            compacted[i][j][3],
                            compacted[i][j][4],
                        ]
            if not check:
                toappend += ["----", "----", 0.0, ""]
        if not misPunch:
            scores = [toappend[6], toappend[10], toappend[14], toappend[18]]
            score = round(sum(findNmax(scores, 4)), 2)
        else:
            score = 0.0
        toappend.append(score)
        out.append(toappend)
    return pd.DataFrame(out)


if __name__ == "__main__":
    print("Start!")
    t1 = perf_counter()

    inputDir = "regiocup/csv"
    outputDir = "regiocup"
    events = [
        {
            "name": "Magdeburger OL",
            "shortname": "MOL2021",
            "firstday": "06-03-2021",
            "lastday": "14-03-2021",
        },
        {
            "name": "Ottos Wald-OL",
            "shortname": "USC2021",
            "firstday": "20-03-2021",
            "lastday": "05-04-2021",
        },
        {
            "name": "64. Kreismeisterschaft Quedlinburg",
            "shortname": "KM64",
            "firstday": "27-03-2021",
            "lastday": "02-05-2021",
        },
        {
            "name": "Wolfsburg",
            "shortname": "WOB",
            "firstday": "24-04-2021",
            "lastday": "09-05-2021",
        },
    ]

    headers = [
        "Bahn",
        "Platzierung",
        "Nachname",
        "Vorname",
        "Verein",
        "ESV Laufdatum",
        "ESV Laufzeit",
        "ESV Punkte",
        "ESV Bemerkung",
        "USC Laufdatum",
        "USC Laufzeit",
        "USC Punkte",
        "USC Bemerkung",
        "QLB Laufdatum",
        "QLB Laufzeit",
        "QLB Punkte",
        "QLB Bemerkung",
        "WOB Laufdatum",
        "WOB Laufzeit",
        "WOB Punkte",
        "WOB Bemerkung",
        "Punkte gesamt",
    ]

    fileList = os.listdir(inputDir)
    for i in range(len(fileList)):
        fileList[i] = inputDir + "/" + fileList[i]

    cupResults, restResults = readFilesCleanData(fileList)

    coursesCup = [[], [], [], [], [], []]
    coursesRest = [[], [], [], [], [], []]

    for i in range(len(cupResults)):
        if cupResults[i]["run"]["course"] == "KL":
            coursesCup[0].append(cupResults[i])
        elif cupResults[i]["run"]["course"] == "KS":
            coursesCup[1].append(cupResults[i])
        elif cupResults[i]["run"]["course"] == "ML":
            coursesCup[2].append(cupResults[i])
        elif cupResults[i]["run"]["course"] == "MS":
            coursesCup[3].append(cupResults[i])
        elif cupResults[i]["run"]["course"] == "LL":
            coursesCup[4].append(cupResults[i])
        elif cupResults[i]["run"]["course"] == "LS":
            coursesCup[5].append(cupResults[i])

    for i in range(len(restResults)):
        if restResults[i]["run"]["course"] == "KL":
            coursesRest[0].append(restResults[i])
        elif restResults[i]["run"]["course"] == "KS":
            coursesRest[1].append(restResults[i])
        elif restResults[i]["run"]["course"] == "ML":
            coursesRest[2].append(restResults[i])
        elif restResults[i]["run"]["course"] == "MS":
            coursesRest[3].append(restResults[i])
        elif restResults[i]["run"]["course"] == "LL":
            coursesRest[4].append(restResults[i])
        elif restResults[i]["run"]["course"] == "LS":
            coursesRest[5].append(restResults[i])

    writer = pd.ExcelWriter("regiocup/Auswertung Regiocup.xlsx", engine="xlsxwriter")

    total = []
    for i in range(len(coursesCup)):
        coursesCup[i] = calcScore(coursesCup[i])
        coursesCup[i] = compactPersons(coursesCup[i])
        coursesCup[i].sort_values(by=20, ascending=False, inplace=True)
        coursesCup[i].reset_index(inplace=True, drop=True)
        coursesCup[i].insert(
            loc=1, value=range(1, len(coursesCup[i]) + 1), column="Platzierung"
        )
        coursesCup[i].replace(to_replace=0.0, value="----", inplace=True)
        coursesCup[i].columns = headers

        coursesRest[i] = compactPersons(coursesRest[i])
        coursesRest[i].sort_values(by=1, inplace=True)
        coursesRest[i].reset_index(inplace=True, drop=True)
        coursesRest[i].insert(loc=1, value="----", column="Platzierung")
        coursesRest[i].replace(to_replace=0.0, value="----", inplace=True)
        coursesRest[i].columns = headers

        final = (
            [headers]
            + coursesCup[i].values.tolist()
            + [" "]
            + coursesRest[i].values.tolist()
        )
        final = pd.DataFrame(final)
        final.columns = headers
        sheetName = final.iloc[-1, 0]
        final.to_excel(writer, sheet_name=sheetName, index=False, header=False)

        worksheet = writer.sheets[sheetName]
        worksheet.set_column(1, 3, 10)
        worksheet.set_column(4, 4, 20)
        worksheet.set_column(5, 21, 14)

        total += coursesCup[i].values.tolist() + [" "]

    total = pd.DataFrame([headers] + total)
    total.to_excel(writer, sheet_name="Gesamt", index=False, header=False)
    writer.save()

    t2 = perf_counter()
    t = round(t2 - t1, 3)
    print("Done! " + str(t) + "s")
    # input("Press Enter to exit...")
