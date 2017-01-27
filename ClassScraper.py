import requests, openpyxl, bs4, re
from pprint import pprint

filename = "CS Threads.xlsx"

def getClasses(url):
    wb = openpyxl.load_workbook(filename)
    sheet = wb.worksheets[0]
    maxRow = sheet.max_row

    res = requests.get(url)
    res.raise_for_status()

    soup = bs4.BeautifulSoup(res.text, "lxml")
    list_elem = soup.select('.body > ul > li')

    class_num_regex = re.compile(r'^(\D+).(\d{4})(.+), (\d)')
    for course in list_elem:
        try:
            match = class_num_regex.findall(course.getText())[0]
            dept = match[0]
            class_num = match[1]
            class_desc = match[2]
            credits = match[3]
            sheet.cell(column = 1, row = maxRow + 1).value = "{} {}".format(dept, class_num)
            sheet.cell(column = 2, row = maxRow + 1).value = class_desc
            sheet.cell(column = 3, row = maxRow + 1).value = credits
            maxRow += 1
        except:
            pass

    sheet.cell(column = 1, row = maxRow + 1).value = "first"
    wb.save(filename)

getClasses("http://www.cc.gatech.edu/information-internetworks")
getClasses("http://www.cc.gatech.edu/intelligence")
getClasses("http://www.cc.gatech.edu/modeling-simulation")
