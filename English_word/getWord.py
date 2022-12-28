import requests
import re
from bs4 import BeautifulSoup
import xlwt
import xlrd

def getword():

        wb = xlrd.open_workbook("word.xlsx")
        sheet1 = wb.sheet_by_index(0)
        WorkBook = xlwt.Workbook(encoding="utf-8")
        WorkSheet = WorkBook.add_sheet("words")
        size=sheet1.nrows
        print('size='+str(size))
        for i in range(size):
            try:
                print(i)
                word=sheet1.cell(i,0).value
                print(word)
                result = main(word)
                soundmark = str(result[0])
                translation = str(result[1])
                for j in range(len(soundmark)-1):
                    if(soundmark[j]=='\''):
                        soundmark=soundmark.split('\'')[0]+'ˈ'+soundmark.split('\'')[1][1:]
                print(soundmark)
                WorkSheet.write(i,1,soundmark)
                WorkSheet.write(i,0,word)
                WorkSheet.write(i,2,translation)
            except BaseException as e:
                WorkSheet.write(i, 1, 'soundmark')
                WorkSheet.write(i, 0, word)
                WorkSheet.write(i, 2, 'translation')
            WorkBook.save('word01.xls')

def getHtmlText(url):
    try:
        kv = {'User-agent': 'Mozilla/5.0'}
        r = requests.get(url, timeout=30, headers=kv)
        r.raise_for_status()
        r.encoding = r.apparent_encoding
        return r.text
    except BaseException as e:
        print(str(e))
def getSoundmark(HtmlText):
    try:
        SM = re.findall(r'\"phonetic\"\>\[.*?\]', HtmlText)
        SoundMark = str(SM[0].split('>')[1])
        #print(SoundMark)
        return SoundMark
    except BaseException as e:
        return "123123"
        print(str(e))
def getTranslation(HtmlText):
    try:
        SM = re.findall(r'<li>([a-z])([\s\S]*?)<\/li>', HtmlText)
        result=''
        for i in range(len(SM)):

            translation = str(SM[i][0]+SM[i][1])
            #print(SoundMark)
            result=result+"  "+translation
            print(translation)
        return result
    except BaseException as e:
        return "123123"
        print(str(e))

def main(word):
    url = "http://www.youdao.com/w/eng/"
    url=url+word
    HtmlText=getHtmlText(url)
    result=(getSoundmark(HtmlText),getTranslation(HtmlText))
    print(result)
    #result.add(getTranslation(HtmlText))
    return result


getword()
