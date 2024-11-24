import os
import re
from datetime import datetime

import zipfile
import requests
from bs4 import BeautifulSoup
import pandas as pd

URL = "https://sfedu.ru/www/stat_pages22.show?p=ELs/sotr/D&x=ELS/9500000000000"
PATH = "a9.html"
months = {
    'января': 'January', 'февраля': 'February', 'марта': 'March',
    'апреля': 'April', 'мая': 'May', 'июня': 'June',
    'июля': 'July', 'августа': 'August', 'сентября': 'September',
    'октября': 'October', 'ноября': 'November', 'декабря': 'December'
}


def createTeachersDict():
    content = requests.get(URL).content
    teachersTable = BeautifulSoup(content, "html.parser").find('table')
    dictPairs = []
    for row in teachersTable.find_all('tr'):
        items = [item.getText() for item in row.find_all('td')]
        pair = (items[0], items[-1] + '@sfedu.ru')
        dictPairs.append(pair)
    teachersDict = dict(dictPairs)
    return teachersDict


def getTeacherEmailByName(teacherName):
    teacherName = teacherName.split()
    # print(teacherName)
    fullLastName = teacherName[0]
    try:
        initials = [teacherName[1][0], teacherName[2][0]]
        regex_pattern = rf"{fullLastName}\s+{initials[0]}\w*\s+{initials[1]}\w*"
    except Exception:
        regex_pattern = rf"{fullLastName.replace('е', '(е|ё)')}"
    teachersNames = list(teachersDict.keys())
    teacher = [key for key in teachersNames if re.match(regex_pattern, key)][0]
    return teachersDict.get(teacher)


def strToDate(dateStr):
    dateList = dateStr.split(",")[1].split()
    dateMonth = months.get(dateList[1])
    dateDay = dateList[0]
    dateRes = datetime.strptime(f'{dateDay} {dateMonth} {str(datetime.now().year)}', '%d %B %Y').date()
    return dateRes.strftime('%d.%m.%Y')


def parseRow(timeRow, testRow, audience):
    dateSession = strToDate(testRow[0][0])

    for i, cell in enumerate(testRow[1:]):
        # print(cell)
        if cell != ['_'] and cell != ['<>', 'ИКТИБ'] and cell != ['<>', 'Кумов А. М.']:
            # print(cell)
            time = timeRow[i][0].split("-")
            timeStart = time[0]
            timeEnd = time[1]
            group = cell[0]
            session = cell[1].split('.')
            # print(session)
            sessionType = "лекция" if session[0] == "лек" else "практика"
            sessionName = session[1]
            teacher = cell[-1]
            teacherEmail = getTeacherEmailByName(teacher)
            groupNumber = re.search(r"\d", group)
            if not groupNumber or "ВПК" in group:
                groupNumber = "ВПК"
            else:
                groupNumber = groupNumber.group()
            resultRow = [dateSession, audience, timeStart,
                         timeEnd, sessionName, sessionType,
                         teacher, teacherEmail, group, groupNumber]

            parseResult.append(resultRow)



def parseTable(table, audience):
    rows = table.find_all("tr")
    tableList = []

    for row in rows:
        rowData = [cell.get_text(separator="|", strip=True).split("|") for cell in row.find_all("td")]

        tableList.append(rowData)

    timeRow = tableList[1][1:]
    for resultRow in tableList[2:-1]:
        parseRow(timeRow, resultRow, audience)



def parseDocument(path):
    with open(path, "r", encoding="windows-1251") as file:
        html_content = file.read()

    soup = BeautifulSoup(html_content, "html.parser")
    tables = soup.find_all("table")

    audience = soup.find('font').next_sibling.get_text(separator="|", strip=True).split("|")[0]
    # print(audience)
    # audienceText = soup.find_all('font')[-1].get_text(strip=True)
    # audience = audienceText.split('<br>')[0].strip()
    # print(audience)

    for table in tables:
        parseTable(table, audience)

    # return tablesList


def unzipDirectory():
    zip_path = 'Archive.zip'
    extract_path = 'unzipped'

    with zipfile.ZipFile(zip_path, 'r') as zip:
        zip.extractall(extract_path)


def cleanHTML(path):
    with open(path, "r", encoding="windows-1251") as file:
        html = file.read()

    html = html.replace("</br>", "<br>")

    with open(path, "w", encoding="windows-1251") as file:
        file.write(html)

teachersDict = createTeachersDict()
unzipDirectory()


parseResult = []
documents = [f"unzipped\\{file}" for file in os.listdir("unzipped") if file.endswith(".html")]
# print(documents)
for document in documents:
    cleanHTML(document)
for document in documents:
    parseDocument(document)

columns = ["Дата", "Аудитория", "Время начала", "Время конца", "Название",
           "Тип занятия", "Преподаватель", "Почта преподавалетя", "Группы", "Курс"]
dateframe = pd.DataFrame(parseResult, columns=columns)

dateframe.to_excel("result.xlsx", index=False)
