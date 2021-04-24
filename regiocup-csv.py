import csv
import pandas as pd
import os


def readCSV(path):
    with open(path, "r", encoding="utf8") as f:
        tmp = list(csv.reader(f))
        f.close()
        return tmp[1:]


def formatEventName(event):
    groups = event.split(" ")
    return " ".join(groups[len(groups) - 3 : len(groups) - 2])


def formatCourseName(course):
    groups = course.split(" ")
    return " ".join(groups[len(groups) - 2 : len(groups) - 1])


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


def getFilesResults(files):
    results = []
    for f in files:
        event = [formatEventName(f)]
        course = [formatCourseName(f)]
        tmp = readCSV(f)

        bestTime = time2sec("24:00:00")
        for i in range(len(tmp)):
            if time2sec(tmp[i][11]) < bestTime and tmp[i][12] == "0":
                bestTime = time2sec(tmp[i][11])
        for i in range(len(tmp)):
            time = time2sec(tmp[i][11])
            if time == 0 or tmp[i][12] == "0":
                score = 0.0
            else:
                score = bestTime / time * 100
                score = round(score, 2)
            tmp[i].insert(25, score)
            tmp[i] = course + event + tmp[i]

        results = results + tmp
        df = pd.DataFrame(results)

    todrop = [2, 3, 4, 7, 8, 9, 10, 11, 12, 15, 17, 18, 19, 21, 22, 23, 24, 25] + list(
        range(28, len(df.columns))
    )
    df.drop(columns=df.columns[todrop], inplace=True)
    df.columns = [
        "coursefile",
        "event",
        "name",
        "surname",
        "time",
        "status",
        "organisation",
        "shortname",
        "date",
        "score",
    ]
    return df


def cleanResults(df):
    print(df)
    df.drop_duplicates(inplace=True)
    df.drop_duplicates(
        subset=["surname", "name", "event", "coursefile"], keep="last", inplace=True
    )
    return df.values.tolist()


def compactPersons(courseresult):
    out = []
    while len(courseresult) > 0:
        name = courseresult[0][3] + " " + courseresult[0][2]
        toappend = [
            courseresult[0][0],
            courseresult[0][2],
            courseresult[0][3],
            courseresult[0][6],
            [courseresult[0][1], courseresult[0][8], courseresult[0][9]],
        ]

        for i in range(1, len(courseresult) - 1):
            if name == courseresult[i][3] + " " + courseresult[i][2]:
                toappend.append(
                    [courseresult[i][1], courseresult[i][8], courseresult[i][9]]
                )
                del courseresult[i]

        out.append(toappend)
        del courseresult[0]
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
                    toappend.append(courseresult[i][j][2])
                    check = True
            if not check:
                toappend.append(0.0)

        scores = toappend[4 : 4 + len(order)]
        score = round(sum(findNmax(scores, 4)), 2)
        toappend.append(score)

        out.append(toappend)
    return out


if __name__ == "__main__":
    inputDir = "input-csv"
    eventOrder = ["MOL2021", "USC2021"]
    eventNames = ["Magdeburger OL", "Ottos Wald-OL"]
    headers = ["Bahn", "Nachname", "Vorname", "Verein"] + eventNames + ["Gesamt"]
    finalHeaders = (
        ["Bahn", "Platzierung", "Nachname", "Vorname", "Verein"]
        + eventNames
        + ["Gesamt"]
    )

    fileList = os.listdir(inputDir)
    for i in range(len(fileList)):
        fileList[i] = inputDir + "/" + fileList[i]

    rawResults = getFilesResults(fileList)
    cleanedList = cleanResults(rawResults)

    courses = [[], [], [], [], [], []]

    for i in range(len(cleanedList)):
        if cleanedList[i][0] == "KL" or :
            courses[0].append(cleanedList[i])
        if cleanedList[i][0] == "KS":
            courses[1].append(cleanedList[i])
        if cleanedList[i][0] == "ML":
            courses[2].append(cleanedList[i])
        if cleanedList[i][0] == "MS":
            courses[3].append(cleanedList[i])
        if cleanedList[i][0] == "LL":
            courses[4].append(cleanedList[i])
        if cleanedList[i][0] == "LS":
            courses[5].append(cleanedList[i])

    total = []
    for i in range(len(courses)):
        courses[i] = compactPersons(courses[i])
        courses[i] = sortEventAddScore(courses[i], eventOrder)
        courses[i] = pd.DataFrame(
            courses[i],
            columns=headers,
        )
        courses[i].sort_values(by="Gesamt", ascending=False, inplace=True)
        courses[i].reset_index(inplace=True, drop=True)
        courses[i].insert(
            loc=1, column="Platzierung", value=range(1, len(courses[i]) + 1)
        )
        courses[i] = pd.DataFrame(
            [finalHeaders] + courses[i].values.tolist(), columns=finalHeaders
        )
        total += courses[i].values.tolist() + [" "]

    writer = pd.ExcelWriter("regiocup/Auswertung Regiocup.xlsx", engine="xlsxwriter")

    pd.DataFrame(total).to_excel(writer, sheet_name="Gesamt", index=False, header=False)
    for i in range(len(courses)):
        courses[i].to_excel(
            writer, sheet_name=courses[i].iloc[-1, 0], index=False, header=False
        )

    writer.save()
    print("Done!")