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
INDEX_URL = 'http://ybt.ssoier.cn:8088/index.php'
TEST_URL = 'http://ybt.ssoier.cn:8088/problem_show.php?pid='
SUBMIT_URL = 'http://ybt.ssoier.cn:8088/submit.php?pid='

#解析首页
def analyze_index():
    print('开始咯')
    workbook = xlwt.Workbook(encoding = 'utf-8')
    html = S.get(INDEX_URL).content
    html = str(html,encoding='utf-8',errors='ignore')
    soup = BeautifulSoup(html,'lxml')
    menu = soup.find(name='div',attrs={'class':'menuDiv'})
    for child in menu.find_all(attrs={'href': "#"}):
        pianming = fstr(child.h3.string.strip())
        print('爬取 %s' % pianming)
        zhanglist = []
        for li in child.parent.ul.find_all(name='li'):
            timulist = analyze_zhangindex(li.a.attrs['href'])
            zhangming = li.string
            print(' 爬取 %s' % zhangming)
            zhanglist.append((zhangming,timulist))
        sheetobj = workbook.add_sheet(pianming)
        print('保存数据……')
        write_to_sheet(pianming,zhanglist,sheetobj)
        workbook.save('dump.xls')
        print('写入完成')
    print('数据保存在 ./dump.xls 中')

#解析指定章，返回题目列表 格式：[(pid,题目名,题目详情),(pid,题目名,题目详情),……]
def analyze_zhangindex(url):
    S.get(BASE_URL + url)
    header = {
        'Referer': 'http://ybt.ssoier.cn:8088/' + url
        }
    html = S.get(BASE_URL,headers=header).content
    html = str(html,encoding='utf-8',errors='ignore')
    soup = BeautifulSoup(html,'lxml')
    
    table = soup.find(name='table',attrs={'class':'plist'})
    lieshu = len(table.find_all(name='th'))
    testlist = []
    if lieshu == 4:
        for td in table.find_all(name='td',attrs={'class':'xlist'}):
            pid = td.previous_sibling.string
            tm = td.string
            xq = analyze_testpage(pid)
            print('   爬取 [%s] %s' % (str(pid),tm))
            testlist.append((pid,tm,xq))
    elif lieshu == 8:
        def specifictd(tag): 
            if(tag.name == 'td'):
                dict = tag.attrs
                if(('class' in dict) and (dict['class'][0] == 'xlist')):
                    return(True)
                elif(tag.find(name='font',attrs={'color':'#001290'})):
                    return(True)
            return(False)
        tds = table.find_all(specifictd)
        a = [tds[i] for i in range(0,len(tds),2)]
        b = [tds[i] for i in range(1,len(tds),2)]
        tds = a + b
        for td in tds:
            pid = td.previous_sibling.string
            tm = td.string
            if(pid):
                xq = analyze_testpage(pid)
                print('   爬取 [%s] %s' % (str(pid),tm))
                testlist.append((pid,tm,xq))
            else:
                print('  爬取 %s' % tm)
                testlist.append((tm,))
    return(testlist)

#解析题目页，返回题目详情，格式：{'tm','i','o','ie','oe',['t'],['s']}
def analyze_testpage(pid):
    html = S.get(TEST_URL + str(pid)).content
    html = str(html,encoding='utf-8',errors='ignore')
    soup = BeautifulSoup(html,'lxml')

    try:
        td = soup.find(name='td',attrs={'class':'pcontent'})
        font = td.find(name='font',attrs={'size':'2'},recursive=False)
        tp = []
        for img in td.find_all(name='img'):
            tp.append(BASE_URL + img.attrs['src'])
    except AttributeError:
        print('***无法读取题目，题号：%s' % str(pid))
        return({'error':'题目不正常，或者是权限类题目!'})

    tm = []
    input = []
    output = []
    inputexp = []
    outputexp = []
    tip = []
    source = None

    flag = 0

    for tag in font.find_all(True):
        if(tag.name == 'font' or tag.name == 'div' or tag.name == 'br'):
            continue
        elif(tag.name == 'h3'):
            if(tag.string != '【来源】'):
                flag+=1
            else:
                flag = -1
            continue

        if(flag == 1):
            tm.append(fstr(tag.string))
        elif(flag == 2):
            input.append(fstr(tag.string))
        elif(flag == 3):
            output.append(fstr(tag.string))
        elif(flag == 4):
            try:
                inputexp.extend(fstr(tag.string,'^$^$^$').split('^$^$^$'))
            except:
                print('***读取题目信息遇到错误，题号 %s' %  str(pid))
        elif(flag == 5):
            try:
                outputexp.extend(fstr(tag.string,'^$^$^$').split('^$^$^$'))
            except:
                print('***读取题目信息遇到错误，题号 %s' %  str(pid))
        elif(flag == 6):
            tip.append(fstr(tag.string))
        elif(flag == -1):
            try:
                if((tag.string).upper() != 'NO' and tag.string != '无'):
                    source = (BASE_URL + tag.attrs['href'],tag.string)
                    print('!!!!!!!!!!%s,%s' % (source[0],source[1]))
            finally:
                break
    if(not tm):
        tm='无'
    if(not input):
        input=['无']
    if(not output):
        output=['无']
    if(not inputexp):
        inputexp=['无']
    if(not outputexp):
        outputexp=['无']
    test = {'tm':tm,'i':input,'o':output,'ie':inputexp,'oe':outputexp}
    if(tp):
        test['tp'] = tp
    if(tip):
        test['t'] = tip
    if(source):
        test['s'] = source
    return(test)

#写出数据到工作簿中
def write_to_sheet(title,list,sheetobj):
    style = xlwt.XFStyle()
    style.alignment.horz = 0x03#HORZ_RIGHT
    sheetobj.write(0,1, label = '篇')
    sheetobj.write(0,2, label = title)
    sheetobj.write(0,3, xlwt.Formula('HYPERLINK("https://blog.chrxw.com";"Generate By Chr_")'))
    sheetobj.col(0).width = 6 * 256   #PID
    sheetobj.col(1).width = 10 * 256 #标记
    sheetobj.col(2).width = 200 * 256 #题目描述

    _row = 1
    for zhang in list:
        _row+=1
        sheetobj.write(_row,1, label ='章')
        sheetobj.write(_row,2, label =zhang[0])
        _row+=1

        for test in zhang[1]:
            #print(test)
            if (len(test) == 1):
                _row+=1
                sheetobj.write(_row,1, label ='节')
                sheetobj.write(_row,2, label =test[0])
                _row+=1
            else:
                _row+=1
                sheetobj.write(_row,0, label =test[0])
                sheetobj.write(_row,1, label ='题目')
                sheetobj.write(_row,2, label =test[1])
                sheetobj.write(_row + 1,0, xlwt.Formula('HYPERLINK("%s";"原题")' % (TEST_URL + test[0])))
                sheetobj.write(_row + 2,0, xlwt.Formula('HYPERLINK("%s";"提交")' % (SUBMIT_URL + test[0])))

                if 'tp' in test[2]:#有无图片
                    _col = 3
                    for i in test[2]['tp']:
                        sheetobj.write(_row,_col,  xlwt.Formula('HYPERLINK("%s";"查看图片")' % i))
                        _col+=1

                _row+=1
                
                if 'error' in test[2]:
                    sheetobj.write(_row,1, label ='题目描述')
                    sheetobj.write(_row,2, label =test[2]['error'])
                    _row+=1
                    continue

                sheetobj.write(_row,1, label ='题目描述')
                for i in test[2]['tm']:
                    sheetobj.write(_row,2, label =i)
                    _row+=1

                sheetobj.write(_row,1, label ='输入')
                for i in test[2]['i']:
                    sheetobj.write(_row,2, label =i)
                    _row+=1

                sheetobj.write(_row,1, label ='输出')
                for i in test[2]['o']:
                    sheetobj.write(_row,2, label =i)
                    _row+=1       

                sheetobj.write(_row,1, label ='输入示例')
                for i in test[2]['ie']:
                    sheetobj.write(_row,2, label =i)
                    _row+=1

                sheetobj.write(_row,1, label ='输出示例')
                for i in test[2]['oe']:
                    sheetobj.write(_row,2, label =i)
                    _row+=1

                if 't' in test[2]:#有无提示
                    sheetobj.write(_row,1, label ='提示')
                    for i in test[2]['t']:
                        sheetobj.write(_row,2, label =i)
                        _row+=1

                if 's' in test[2]:#有无来源
                    sheetobj.write(_row,1, label ='来源')
                    sheetobj.write(_row,2, xlwt.Formula('HYPERLINK("%s";"%s")' % test[2]['s']))
                    _row+=1
    pass

def fstr(str,replace=''):
    if(str):
        r1 = re.compile(r'&nbsp|\$|_|\xa0|\\xa0')
        r2 = re.compile(r'\t|\r\n|\r|\n')
        str = r1.sub('',str)
        str = r2.sub(replace,str)
    return(str)

if __name__ == "__main__":
    analyze_index()