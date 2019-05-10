import requests
import re
import xlwt
import urllib3
from io import BytesIO
from bs4 import BeautifulSoup  
import bs4

'''
By Chr_
Email: chr@chrxw.com
'''

def download_exam():
    print('爬爬爬，开始咯')
    s=requests.session()
    workbook = xlwt.Workbook(encoding = 'utf-8')
    for j in range(1,102,1):
        txt=[]
        url='http://lib.nbdp.net/paper/%d.html'  % j
        html=s.get(url).content
        html=str(html,encoding='utf-8',errors='ignore')
        soup = BeautifulSoup(html,'lxml')  
        exams=soup.find_all(name='div',attrs={'s':'math3'})
        title=str(j)+soup.title.get_text()
        worksheet = workbook.add_sheet(title)
        exam=[(url,title)]
        print('当前进度[%d/100]：%s [%s]' % (j,exam[0][1],exam[0][0]))
        for x in exams:
            out=analyzesoup(x)
            exam.append(out)
        sheetwriter(exam,worksheet)
    workbook.save('dump.xls')
    print('保存为./dump.xls')
    pass

def analyzesoup(soupobj:bs4.element.NavigableString):
    def notNone(obj):
        if obj is None:
            return(False)
        return(True)
    result=soupobj.find(name='p',attrs={'class':'pt1'})#题干
    if notNone(result):
        tigan=result.get_text().strip()

        result=soupobj.find(name='img')#有无图片
        if (notNone(result)):
            hasimg=True
            tupian=[]
            for i in soupobj.find_all(name='img'):
                tupian.append(i.attrs['src'])
        else:
            hasimg=False

        result=soupobj.find(name='li')#如果是选择题
        if notNone(result):
            xuanxiang=[]
            for i in soupobj.find_all(name='li'):
                xuanxiang.append(i.get_text())
            result=soupobj.find(attrs={'class':'col-md-3 column xz'})
            if notNone(result):
                daan=result.get_text()
            else:
                daan='未找到答案'
            
            if hasimg:
                out={'tg':tigan,'xx':xuanxiang,'da':daan,'tp':tupian}
            else:
                out={'tg':tigan,'xx':xuanxiang,'da':daan}

            return(out)

        result=soupobj.find(name='pre')#是否存在代码块
        if notNone(result):
            daima=result.get_text()
            if hasimg:
                out={'tg':tigan,'dm':daima,'tp':tupian}
            else:
                out={'tg':tigan,'dm':daima}       
            return(out)
        else:
            result=soupobj.find(name='span',attrs={'class':'tiankong'})
            if notNone(result):
                jianda=result.get_text()

            if hasimg:
                out={'tg':tigan,'jd':jianda,'tp':tupian}
            else:
                out={'tg':tigan,'jd':jianda}
            return(out)

        return(None)

def sheetwriter(list,sheetobj):
    sheetobj.write(0,1, xlwt.Formula('HYPERLINK("%s";"%s")' % list[0]))
    sheetobj.write(0,2, xlwt.Formula('HYPERLINK("https://blog.chrxw.com";"Generate By Chr_")'))
    sheetobj.col(0).width=0
    sheetobj.col(1).width=80*256
    sheetobj.col(2).width=40*256
    sheetobj.col(3).width=40*256
    sheetobj.col(4).width=40*256
    sheetobj.col(5).width=40*256
    _row=1
    for item in list:
        #print(item)
        if 'tp' in item:#插图片
            _col=7
            for tp in item['tp']:
                url='http://lib.nbdp.net/'+tp
                sheetobj.write(_row,_col, xlwt.Formula('HYPERLINK("%s";"查看图片")' % url))
                _col+=1
            pass

        if 'xx' in item:#选择题
            sheetobj.write(_row,1, label =item['tg'])
            sheetobj.write(_row,0, label =item['da'])
            col=2
            for i in item['xx']:
                sheetobj.write(_row,col, label =i)
                col+=1
            _row+=1
            continue

        if 'dm' in item:#填空题
            lines=item['tg'].splitlines(False)
            _row+=1
            for line in lines:
                sheetobj.write(_row,1, label =line)
                _row+=1
            lines=item['dm'].splitlines(False)
            for line in lines:
                sheetobj.write(_row,1, label =line)
                _row+=1
            continue

        if 'jd' in item:#简答题
            row1=_row
            row2=_row
            lines=item['tg'].splitlines(False)
            for line in lines:
                sheetobj.write(row1,1, label =line)
                row1+=1
            lines=item['jd'].splitlines(False)
            for line in lines:
                sheetobj.write(row2,0, label =line)
                row2+=1
            _row=(row1 if row1>row2 else row2)+1
    pass

if __name__=="__main__":
    download_exam()
