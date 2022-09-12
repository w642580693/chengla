from distutils.command.build_scripts import first_line_re
import requests
import xlrd
from re import A, S
from traceback import print_tb
import re

print('-----')
print('目前只支持excel的格式为xls')
print('粘贴请用鼠标右键')
print('如需在标记过程终止，请键入ctrl+c')
print('版本号：1.2')
print('-----')


headers ={
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/105.0.0.0 Safari/537.36'
}
#构造post发送方法
def postto(phone):
    post_url='https://clapp.orangevip.com/otm/web/activity/doActivity'
    data = {
        'mobilePhone':phone,
        'uniCode':activity,
        'vCode':'1234',
    }
    requests.post(url=post_url,headers=headers,data=data)



#构造获取excel内容
def ExceL_phone(num):
    n=sheet_names.nrows
    return sheet_names.row_values(num)[trann-1]

#构造(需要设置)电话号码在哪一列函数
def lieerr():
    try:
        global trann
        trann=int(input('电话号码在哪一列:'))
        print('-----')
    except ValueError:
        print('请输入数字！')
        print('-----')
        lieerr()
#构造(需要设置)xls里的电话号码一共有多少行函数
def finaerr():
    try:
        global SuM
        SuM=int(input('标记到哪一行结束？'))
        print('-----')
    except ValueError:
        print('请输入数字！')
        print('-----')
        finaerr()

#构造从哪一行开始标记？函数
def frierr():
    try:
        global fri
        fri=int(input('哪一行开始标记？'))
        print('-----')
    except ValueError:
        print('请输入数字！')
        print('-----')
        frierr()
#获取表格基本信息
def getex():
    try:
        global SheeT
        global eclname
        eclname=str(input('请输入excel表格文件位置:'))
        print('-----')
        SheeT=str(input('请输入表格名（如Sheet1,需区分大小写）:'))
        global worksheet
        global sheet_names
        worksheet = xlrd.open_workbook(eclname)
        sheet_names = worksheet.sheet_by_name(SheeT)
        print('-----')
    except FileNotFoundError:
        print('未检测到表，请检查文件名是否存在，或者表名是否正确（注意区分大小写）')
        print('-----')
        getex()



# 号段列表正则表达式
pattern = re.compile(r'^(13[0-9]|14[0|5|6|7|9]|15[0|1|2|3|5|6|7|8|9]|'
                         r'16[2|5|6|7]|17[0|1|2|3|5|6|7|8]|18[0-9]|'
                         r'19[1|3|5|6|7|8|9])\d{8}$')
#计数器
summ=0
hansum=0
while 1 :
    
    # #(需要设置)读取的excel文件名
    # eclname=str(input('请输入excel表格文件位置:'))
    # print('-----')
    
    # #(需要设置)表格的序列号
    # SheeT=str(input('请输入表格名（如Sheet1,需区分大小写）:'))
    # print('-----')
    # #获取表格基本信息
    # worksheet = xlrd.open_workbook(eclname)
    # sheet_names = worksheet.sheet_by_name(SheeT)
    getex()
    #(需要设置)电话号码在哪一列
    lieerr()

    #从哪一行开始标记？
    frierr()

    #(需要设置)标记到哪一行结束？
    finaerr()
    # SuM=int(input('xls里的电话号码一共有多少行：'))
    # print('-----')

    #需要设置)提交的网站地址
    activity = str(input('请输入activity参数:'))
    print('-----')



    #判断是否为浮点数并转化为字符串 
    for i in range(fri-1,SuM):
        phone=ExceL_phone(i)
        if type(phone) == float:
            aphone=str(phone)[:11]

        else:
            aphone=str(phone)
    #判断是否为电话号码并标记
        if pattern.search(aphone):
            print('已获取到号码'+aphone)
            postto(aphone)
            summ+=1
            hansum+=1
            print(aphone+'标记成功')
            print('已读取'+str(hansum)+'行，'+'共'+str(SuM-fri+1)+'行，已标记'+str(summ)+'个')
            print('-----')
        else:
            print('"'+str(phone)+'"'+"不是号码！")
            hansum+=1
            print('已读取'+str(hansum)+'行，'+'共'+str(SuM-fri+1)+'行，已标记'+str(summ)+'个')
            print('-----')

    print('本次标记已读取'+str(hansum)+'行，'+'共'+str(SuM-fri+1)+'行，已标记'+str(summ)+'个')
    print('本次标记结束')
    print('-----')
    summ=0
    hansum=0
        