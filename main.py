import numpy as np
import pandas as pd
import requests
import re
from bs4 import BeautifulSoup
from pathlib import Path
from datetime import datetime
from itertools import chain
from zipfile import ZipFile


URL = "https://sfedu.ru/www/stat_pages22.show?p=ELs/sotr/D&x=ELS/9500000000000"
columns = ["Дата", "Аудитория", "Время начала", "Время конца", "Название",
           "Тип занятия", "Преподаватель", "Почта преподавалетя", "Группы", "Курс"]
months = {
    'января': 'January', 'февраля': 'February', 'марта': 'March',
    'апреля': 'April', 'мая': 'May', 'июня': 'June',
    'июля': 'July', 'августа': 'August', 'сентября': 'September',
    'октября': 'October', 'ноября': 'November', 'декабря': 'December'
}


def pandasTeacherTable():
    content = requests.get(URL).content
    df = pd.read_html(content, flavor="lxml")[0]
    df = df.drop([1, 2], axis=1)
    df.columns = ["Teacher", "Email"]
    df["Email"] = df["Email"] + "@sfedu.ru"
    return df


def getTeacherEmailByName(teacherName):
    teacherName = teacherName.split()
    fullLastName = teacherName[0]
    try:
        initials = [teacherName[1][0], teacherName[2][0]]
        regex_pattern = rf"{fullLastName}\s+{initials[0]}\w*\s+{initials[1]}\w*"
    except Exception:
        regex_pattern = rf"{fullLastName.replace('е', '(е|ё)')}"

    id = teachersEmail["Teacher"].str.contains(regex_pattern, regex=True).idxmax()
    return teachersEmail.loc[id, "Email"]


def unzipDirectory():
    zip_path = 'Archive.zip'
    extract_path = 'unzipped'

    with ZipFile(zip_path, 'r') as zip:
        zip.extractall(extract_path)


def parseTime(time):
    time = time.split("-")
    timeStart = time[0]
    timeEnd = time[1]
    return timeStart, timeEnd


def strToDate(dateStr):
    dateList = dateStr.split(",")[1].split()
    dateMonth = months.get(dateList[1])
    dateDay = dateList[0]
    dateRes = datetime.strptime(f'{dateDay} {dateMonth} {str(datetime.now().year)}', '%d %B %Y').date()
    return dateRes.strftime('%d.%m.%Y')


def splitCell(cell):
    if ' лек.' in cell:
        delimiter = ' лек.'
    elif ' пр.' in cell:
        delimiter = ' пр.'
    else:
        raise ValueError("Неизвестный формат строки")

    groups, rest = cell.split(delimiter, 1)
    session, teacher = rest.rsplit(" ", 3)[0], " ".join(rest.rsplit(" ", 3)[1:])
    session = delimiter.strip() + session
    return groups, session, teacher


def parseCell(cell):
    if cell in ["<> ИКТИБ", "<> Кумов А. М.", "_"]:
        return np.nan
    try:
        cell = splitCell(cell)
    except ValueError:
        print("Value error", cell)
        return np.nan


    groups = cell[0]
    session = cell[1].split('.')
    sessionType = "лекция" if session[0] == "лек" else "практика"
    sessionName = session[1]
    teacher = cell[-1]
    if "2" in teacher: # hardcoding(((
        teacher = "Пленкин"
    teacherEmail = getTeacherEmailByName(teacher)


    groupNumber = re.search(r"\d", groups)
    if not groupNumber or "ВПК" in groups:
        groupNumber = "ВПК"
    else:
        groupNumber = groupNumber.group()


    cell = {"Группы": groups, "Тип занятия": sessionType, "Название": sessionName,
            "Преподаватель": teacher, "Почта преподавалетя": teacherEmail, "Курс": groupNumber}
    return cell


def updateCellWithTime(d, time):
    if pd.isna(d):
        return d
    timeStart, timeEnd = parseTime(time)
    d["Время начала"] = timeStart
    d["Время конца"] = timeEnd
    return d


def updateRowWithDate(d):
    date = d.iloc[0]
    for i in range(1, len(d)):
        if isinstance(d.iloc[i], dict):
            d.iloc[i]["Дата"] = date
    return d


def updateCellWithAudience(d, audience):
    if isinstance(d, dict):
        d["Аудитория"] = audience
        return d
    return d


def changeToValid(table, audience):
    timeRow = table.loc[1]
    t = table.copy()[2:-1]
    t.rename(columns={0: "Дата"}, inplace=True)
    t["Дата"] = t["Дата"].apply(strToDate)
    for col in t.columns[1:]:
        t[col] = t[col].apply(parseCell)
        t[col] = t[col].apply(lambda d: updateCellWithTime(d, timeRow[col]))
        t[col] = t[col].apply(lambda d: updateCellWithAudience(d, audience))
    t = t.apply(updateRowWithDate, axis=1)
    return t


def parseHTML(path):
    with open(path, "r", encoding="windows-1251") as file:
        html_content = file.read()

    soup = BeautifulSoup(html_content, "html.parser")

    audience = soup.find('font').next_sibling.get_text(separator="|", strip=True).split("|")[0]

    html = html_content.replace("</br>", "<br>")

    with open(path, "w", encoding="windows-1251") as file:
        file.write(html)

    return audience


teachersEmail = pandasTeacherTable()

docs = [str(file) for file in Path("unzipped").iterdir() if Path(file).is_file()]
audiences = [parseHTML(doc) for doc in docs]

tables = [pd.read_html(doc, flavor="lxml") for doc in docs]
tables = dict(zip(audiences, tables))

for audience, table in tables.items():
    tables[audience] = list(map(lambda el: changeToValid(el, audience), table))
table = chain.from_iterable(tables.values())
table = pd.concat(table).drop(columns=["Дата"])

resultTables = [pd.json_normalize(table[i]) for i in range(1, 8)]
resultTable = pd.concat(resultTables)
resultTable = resultTable.dropna(how="all")

resultTable.to_excel("result.xlsx", index=False)

