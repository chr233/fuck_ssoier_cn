import requests
import re
import xlwt
import urllib3
from bs4 import BeautifulSoup
import bs4

'''
By Chr_
Email: chr@chrxw.com
'''
S = requests.session()
BASE_URL = 'http://ybt.ssoier.cn:8088/'

def main():
    analyze_indexpage(generate_pagetree())

#解析首页，返回章节大纲列表 格式：[(章名,(篇名,链接)]
def generate_pagetree():
    html = S.get(BASE_URL + 'index.php').content
    html = str(html,encoding='utf-8',errors='ignore')
    soup = BeautifulSoup(html,'lxml')
    x = soup.find(name='div',attrs={'class':'menuDiv'})
    pagetree = []
    for child in x.find_all(attrs={'href': "#"}):
        pianming = child.h3.string.strip()
        zhanglist = []
        for li in child.parent.ul.find_all(name='li'):
            zhanglist.append((li.string,li.a.attrs['href']))
        pagetree.append((pianming,zhanglist))
    return(pagetree)

#解析章节大纲，返回题目列表
def analyze_indexpage(pagetree:list):
    print('爬爬爬，开始咯')
    workbook = xlwt.Workbook(encoding = 'utf-8')
    for pian in pagetree:
        for url in pian[1]:
            S.get(BASE_URL + url[1])
            header = {
                'Referer': 'http://ybt.ssoier.cn:8088/' + url[1]
                }
            html = S.get(BASE_URL,headers=header).content
            html = str(html,encoding='utf-8',errors='ignore')
            soup = BeautifulSoup(html,'lxml')
            
            table = soup.find(name='table',attrs={'class':'plist'})
            lieshu = len(table.find_all(name='th'))
            test = []
            if lieshu == 4:
                for td in table.find_all(name='td',attrs={'class':'xlist'}):
                    tm = td.string
                    pid = td.previous_sibling.string
                    test.append((pid,tm))
            elif lieshu == 8:
                i = 0
                def specifictd(tag):
                    if(tag.name == 'td'):
                        dict = tag.attrs
                        if(('class' in dict) and (dict['class'][0] == 'xlist')):
                            return(True)
                        elif(tag.find(name='font',attrs={'color':'#001290'})):
                            return(True)
                    return(False)
                tds = table.find_all(specifictd)
                a=[tds[i] for i in range(0,len(tds),2)]
                b=[tds[i] for i in range(1,len(tds),2)]
                tds=a+b
                for td in tds:
                    tm = td.string
                    pid = td.previous_sibling.string
                    test.append((pid,tm))

            sheetobj = workbook.add_sheet(pian[0])
            analyze_testpagelist(test,sheetobj)
            



#解析题目页，返回题目详情，
def analyze_testpagelist(testlist,sheetobj):
     def notNone(obj):
        if obj is None:
            return(False)
        return(True)
     for item in testlist:
        if(item[0]):#有pid
            html = S.get(BASE_URL + 'problem_show.php?pid=' + pid).content
            html = str(html,encoding='utf-8',errors='ignore')
            soup = BeautifulSoup(html,'lxml')
            form=soup.find(name='td',attrs={'class':'pcontent'})

        else:#无pid
            pass


def download_exam():
    for j in range(1000,1671,1):
        txt = []
        url = 'http://ybt.ssoier.cn:8088/index.php#'
        html = s.get(url).content
        html = str(html,encoding='utf-8',errors='ignore')
        soup = BeautifulSoup(html,'lxml')

        title = str(j) + soup.find(attrs={'class':'pcontent'}).get_test()
        worksheet = workbook.add_sheet(title)
        exam = [(url,title)]

        print('当前进度[%d/100]：%s [%s]' % (j,exam[0][1],exam[0][0]))

        sheetwriter(exam,worksheet)
    workbook.save('dump.xls')
    print('保存为./dump.xls')
    pass

def analyzesoup(soupobj:bs4.element.NavigableString):
    

    result = soupobj.find(name='p',attrs={'class':'pt1'})#题干
    if notNone(result):
        tigan = result.get_text().strip()

        result = soupobj.find(name='img')#有无图片
        if (notNone(result)):
            hasimg = True
            tupian = []
            for i in soupobj.find_all(name='img'):
                tupian.append(i.attrs['src'])
        else:
            hasimg = False

        result = soupobj.find(name='li')#如果是选择题
        if notNone(result):
            xuanxiang = []
            for i in soupobj.find_all(name='li'):
                xuanxiang.append(i.get_text())
            result = soupobj.find(attrs={'class':'col-md-3 column xz'})
            if notNone(result):
                daan = result.get_text()
            else:
                daan = '未找到答案'

            if hasimg:
                out = {'tg':tigan,'xx':xuanxiang,'da':daan,'tp':tupian}
            else:
                out = {'tg':tigan,'xx':xuanxiang,'da':daan}

            return(out)

        result = soupobj.find(name='pre')#是否存在代码块
        if notNone(result):
            daima = result.get_text()
            if hasimg:
                out = {'tg':tigan,'dm':daima,'tp':tupian}
            else:
                out = {'tg':tigan,'dm':daima}
            return(out)
        else:
            result = soupobj.find(name='span',attrs={'class':'tiankong'})
            if notNone(result):
                jianda = result.get_text()

            if hasimg:
                out = {'tg':tigan,'jd':jianda,'tp':tupian}
            else:
                out = {'tg':tigan,'jd':jianda}
            return(out)

        return(None)

def sheetwriter(list,sheetobj):
    sheetobj.write(0,1, xlwt.Formula('HYPERLINK("%s";"%s")' % list[0]))
    sheetobj.write(0,2, xlwt.Formula('HYPERLINK("https://blog.chrxw.com";"Generate By Chr_")'))
    sheetobj.col(0).width = 0
    sheetobj.col(1).width = 80 * 256
    sheetobj.col(2).width = 40 * 256
    sheetobj.col(3).width = 40 * 256
    sheetobj.col(4).width = 40 * 256
    sheetobj.col(5).width = 40 * 256
    _row = 1
    for item in list:
        #print(item)
        if 'tp' in item:#插图片
            _col = 7
            for tp in item['tp']:
                url = 'http://lib.nbdp.net/' + tp
                sheetobj.write(_row,_col, xlwt.Formula('HYPERLINK("%s";"查看图片")' % url))
                _col+=1
            pass

        if 'xx' in item:#选择题
            sheetobj.write(_row,1, label =item['tg'])
            sheetobj.write(_row,0, label =item['da'])
            col = 2
            for i in item['xx']:
                sheetobj.write(_row,col, label =i)
                col+=1
            _row+=1
            continue

        if 'dm' in item:#填空题
            lines = item['tg'].splitlines(False)
            _row+=1
            for line in lines:
                sheetobj.write(_row,1, label =line)
                _row+=1
            lines = item['dm'].splitlines(False)
            for line in lines:
                sheetobj.write(_row,1, label =line)
                _row+=1
            continue

        if 'jd' in item:#简答题
            row1 = _row
            row2 = _row
            lines = item['tg'].splitlines(False)
            for line in lines:
                sheetobj.write(row1,1, label =line)
                row1+=1
            lines = item['jd'].splitlines(False)
            for line in lines:
                sheetobj.write(row2,0, label =line)
                row2+=1
            _row = (row1 if row1 > row2 else row2) + 1
    pass

if __name__ == "__main__":
    main()
