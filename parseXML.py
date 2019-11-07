import xml.etree.ElementTree as ET
import urllib
import urllib.request
import openpyxl
import datetime

def GetSuggestWord(keyword):
    keyword = urllib.parse.quote(keyword,encoding='utf-8')
    url = "https://www.google.com/complete/search?hl=jp&output=toolbar&ie-utf_8&oe=utf_8&q=%22" + keyword + "%22"
    req = urllib.request.Request(url)
    print(url)

    with urllib.request.urlopen(req) as response:
        XmlData = response.read()


    root = ET.fromstring(XmlData)

    ret = []
    for i in range(4):
        ret.append(root[i][0].attrib['data'])
    print(ret)
    return ret

def WriteExcel(keyword):

    try:
        wb = openpyxl.load_workbook('Ironword.xlsx')
    except FileNotFoundError:
        wb = openpyxl.Workbook()
        sheet = wb.active
        sheet.title = 'Sheet1'
        wb.save('Ironword.xlsx')

    sheet = wb['Sheet1']
    cnt = 1
    while sheet["A"+str(cnt)].value != None:
        cnt += 1

    sheet["A"+str(cnt)] = datetime.date.today()
    sheet["B"+str(cnt)] = "word: " + keyword
    suggests = GetSuggestWord(keyword)

    uni = 67
    for w in suggests:
        sheet[str(chr(uni))+str(cnt)] = w
        uni += 1

    wb.save('Ironword.xlsx')


if __name__ == "__main__":
    # %20 = SPACE
    # %22 = "
    WriteExcel("アイアンFX")
    WriteExcel("アイアン FX")
    WriteExcel("ironfx")
    WriteExcel("iron fx")
