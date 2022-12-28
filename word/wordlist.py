
import datetime
import xlrd
import yagmail
import openpyxl


def getword():
    daynumber = 100
    wordslist =[[]]

    rb = xlrd.open_workbook("wordsr.xlsx")
    tabler = rb.sheets()[1]
    today = tabler.row_values(rowx=0, start_colx=11, end_colx=12)[0]  # 0.11记录的是距离第一天的时间
    wb = openpyxl.load_workbook("wordsw.xlsx")
    sheet = wb['Sheet2']

    sheet['L1']=sheet['L1'].value+1
    for i in range(1, tabler.nrows):#遍历单词表
        datevalue = tabler.row_values(rowx=i, start_colx=11, end_colx=13)#
        word=tabler.row_values(rowx=i,start_colx=0,end_colx=10)
        if isrember(datevalue[1], today,datevalue[0]):#判断是否为今天所记单词
            for x in range(10):#遍历单词内容
                if word[x]!="":#删除空格
                    1+1
                else:
                    break
            word=word[0:x]
            wordslist.append(word)

            if datevalue[1]==0:#第一次记
                sheet['M'+str(i+1)]=sheet['M'+str(i+1)].value+1
                sheet['L' + str(i + 1)] = sheet['L1'].value
            else:
                sheet['M' + str(i + 1)] = sheet['M' + str(i + 1)].value + 1
            daynumber -= 1
            if daynumber == 0:
                break
    wb.save('wordsw.xlsx')
    return wordslist

def isrember(number, today, firstday):
    if number == 0:
        return True
    elif number == 1:
        if today - firstday == 1:
            return True
    elif number == 2:
        if today - firstday == 2:
            return True
    elif number == 3:
        if today - firstday == 4:
            return True
    elif number == 4:
        if today - firstday == 7:
            return True
    elif number == 5:
        if today - firstday == 15:
            return True
    else:
        False
def getsubhtml(wordlist):
    str='''<div>'''
    for i in range(3):
        if i==0:
            str=str+'''<div class="word">
                    <texe>'''+wordlist[0]+'''\n</texe>
                </div>'''
        elif i==1:
            str=str+'''<div class="mark">
                    <texe>'''+wordlist[1]+''' \n</texe>   
                </div>'''
        elif i==2:
            str = str + '''<div class="translate">\n'''
            for x in range(2,len(wordlist)):
                if wordlist[x]!="":
                    str=str+''' <text class="tetext">\n'''+wordlist[x]+''' \n</text>'''
                else:
                    break
            str=str+'''</div>'''
    str=str+'''</div>'''

    return str

def gethtml(wordslist):
    file = open("wordlist.html","wb")
    text1='''<!DOCTYPE html>
<html>

<head>

<meta charset="ANSI">

<title>Document</title>

<style>
.parent {

display: table;

width: 100%;

table-layout: fixed;

}



.left3 {

display: table-cell;

height: 297mm;

width: 210mm;

background-color: pink;

}



.right3 {

display: table-cell;

height: 297mm;
width: 105mm;

background-color: purple;

}
.word{
    display: table-cell;

    font-size: 50px;
}
.mark{
    display: table-cell;

    font-size: 30px;
}
.translate{
    display: table;

    font-size: 30px;

}
.tetext{
    display: table;
    width: 210mm;
}

</style>

</head>

<body class="parent">

    <div>

        <div class="left3">'''
    text2='''        </div>
    </div>
</body>
</html>'''
    file.write(bytes(text1, 'UTF-8'))
    for i in range(len(wordslist)):
        substr=getsubhtml(wordslist[i])
        file.write(bytes(substr,'UTF-8'))
    file.write(bytes(text2, 'UTF-8'))
    file.close

def sendmail():
    yag = yagmail.SMTP(host='smtp.qq.com', user='704657714@qq.com',password='ulwkydutivzsbbic', smtp_ssl=True).send('704657714@qq.com', 'hello', 'hello','wordlist.html')

def getnewfile():
    file1=open("wordsr.xlsx",'wb')
    file2=open("wordsw.xlsx",'rb')
    file1.write(file2.read())
    file1.close()
    file2.close()



def main():
    gethtml(getword()[1:])
    sendmail()
    getnewfile()
main()
