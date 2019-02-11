import time
import tkinter
from datetime import datetime
import os

import json
from tkinter import scrolledtext
from tkinter import messagebox

import xlwt
from bs4 import BeautifulSoup
from selenium import webdriver

from config import SERVICE_ARGS

# browser = webdriver.PhantomJS(executable_path='C:\\Users\\kenmeon\\Desktop\\phantomjs-2.1.1-windows\\bin\\phantomjs.exe')

browser = webdriver.PhantomJS(service_args=SERVICE_ARGS)

# browser = webdriver.Chrome()

root = tkinter.Tk()


# 导出的数组信息
exports = []

def get_screen_size(window):
    return window.winfo_screenwidth(), window.winfo_screenheight()


def get_window_size(window):
    return window.winfo_reqwidth(), window.winfo_reqheight()


def center_window(root, width, height):
    screenwidth = root.winfo_screenwidth()
    screenheight = root.winfo_screenheight()
    size = '%dx%d+%d+%d' % (width, height, (screenwidth - width) / 2, (screenheight - height) / 2)
    print(size)
    root.geometry(size)

def get_page_data(url):
    browser.get(url)
    html = browser.page_source

    soup = BeautifulSoup(html, 'lxml')
    codes = soup.find_all('code')
    code = codes[len(codes) - 3]
    return code.string
def parse_page(dataJson):
    toalPage = 0
    if dataJson and 'data' in dataJson.keys():
        data = dataJson.get('data')
        if data and 'paging' in data.keys():
            paging = data.get('paging')
            if paging:
                total = paging.get('total')
                count = paging.get('count')
                if total and count:
                    e = total / count
                    n = total % count
                    if not n == 0:
                        e = e + 1
                    toalPage = int(e)
    return toalPage

def parse_position(dataJson):
    # 解析出所有的职位数组
    if dataJson and 'data' in dataJson.keys():
        data = dataJson.get('data')
        if data and 'elements' in data.keys():
            pelements = data.get('elements')
            if pelements and len(pelements) > 0:
                for pelement in pelements:
                    if pelement and 'elements' in pelement.keys():
                        elements = pelement.get('elements')
                        if elements and len(elements) > 0:
                            for element in elements:
                                product = {
                                    'name': '',
                                    'position': '',
                                    'company': '',
                                    'address': ''
                                }

                                title = element.get('title')
                                if title:
                                    product['name'] = title.get('text')
                                headline = element.get('headline')
                                if headline:
                                    product['position'] = headline.get('text')
                                snippetText = element.get('snippetText')
                                if snippetText:
                                    product['company'] = snippetText.get('text')
                                subline = element.get('subline')
                                if subline:
                                    product['address'] = subline.get('text')

                                exports.append(product)
                                lb.insert(tkinter.END,'已完成' + str(len(exports)) + '\n')
                                lb.update()

                        else:
                            continue
                    else:
                        print('不包含elements的key')
        else:
            print('不包含pelements的key')
    # 解析职位完成

def write_to_excel():

    if companyIn.get() == None or companyIn.get() == '':
        messagebox.showinfo('提示', '请输入公司编号\n')
        lb.see(tkinter.END)
        return
    if keywordIn.get() == None or keywordIn.get() == '':
        messagebox.showinfo('提示', '请输入关键词\n')
        return

    lb.insert(tkinter.END,'开始导出Excel\n')
    facetCurrentCompany = companyIn.get()
    keyWord = keywordIn.get()

    titles = ['姓名','职位','公司','地址']
    wb = xlwt.Workbook(encoding='utf-8')
    ws = wb.add_sheet('Sheet1')

    tall_style = xlwt.easyxf('font:height 360;')  # 36pt,类型小初的字号
    # 设置title标题
    for index,title in enumerate(titles):
        ws.write(0, index, title)
        # 设置列宽
        ws.col(index).width = 7000
        # 设置行高
        ws.row(0).set_style(tall_style)
    # 设置内容
    for index,person in enumerate(exports):
        ws.write(index + 1, 0, person.get('name'))
        ws.write(index + 1, 1, person.get('position'))
        ws.write(index + 1, 2, person.get('company'))
        ws.write(index + 1, 3, person.get('address'))
        # 设置行高
        ws.row(index + 1).set_style(tall_style)

    dir = datetime.now().strftime('%Y%m%d')
    if not os.path.exists(dir):
        os.mkdir(dir)
    fileName = ''
    if facetCurrentCompany and not facetCurrentCompany == '':
        fileName = facetCurrentCompany
    else:
        fileName = 'company'

    if keyWord and not keyWord == '':
        fileName = fileName + keyWord
    else:
        fileName = 'company' + keyWord
    fileName.replace('\\n','')
    filePath = dir + '/' + fileName + '-' + dir + '.xls'
    print(filePath)
    wb.save(filePath)

    lb.insert(tkinter.END, '导出Excel完成！！！\n')

def executeSearch():



    if companyIn.get() == None or companyIn.get() == '':
        messagebox.showinfo('提示', '请输入公司编号\n')
        lb.see(tkinter.END)
        return
    if keywordIn.get() == None or keywordIn.get() == '':
        messagebox.showinfo('提示', '请输入关键词\n')
        return
    facetCurrentCompany = companyIn.get()
    keyWord = keywordIn.get()

    companyIds = facetCurrentCompany.split('+')
    companySearch = '%5B'
    for index,companyId in enumerate(companyIds):
        if index == 0:
            companySearch = companySearch + '"' + companyId + '"'

        else:
            companySearch = companySearch + '%2C' + '"' + companyId + '"'
    companySearch = companySearch + '%5D'


    url = 'https://www.linkedin.com/search/results/people/?facetCurrentCompany='+companySearch+'&keywords='+keyWord+'&origin=FACETED_SEARCH'
    code = get_page_data(url)

    print(url)
    dataJson = json.loads(code)
    # 解析计算总页数
    toalPage = parse_page(dataJson)

    print(toalPage)
    # 解析职位完成
    parse_position(dataJson)
    print('当前职位总数')
    print(len(exports))
    # //执行翻页查询数据

    if toalPage > 1:
        for i in range(toalPage):
            if (i + 1) > 1:
                pageTo = i + 1
                moreUrl = url + '&page=' + str(pageTo)
                print('请求地址')
                print(moreUrl)
                code = get_page_data(moreUrl)
                dataJson = json.loads(code)
                # 解析职位完成
                parse_position(dataJson)
                print('当前职位总数')
                print(len(exports))

    # 所有数据检索完成
    print('所有数据检索完成,执行导出操作')
    lb.insert(tkinter.END,'所有数据检索完成\n')

def loginYouUsername():
    # print('正在打开登录页面...\n')

    lb.insert(tkinter.END, '正在打开登录页面...\n')
    lb.see(tkinter.END)
    try:

        # 模拟登陆的操作
        browser.get('https://www.linkedin.com')
        time.sleep(1)
        lb.insert(tkinter.END, '正在登录...\n')
        print('正在登录...\n')
        browser.find_element_by_id('login-email').send_keys("*****")
        browser.find_element_by_id('login-password').send_keys("****")
        browser.find_element_by_id('login-submit').click()
        lb.insert(tkinter.END, '登录成功\n')

        # 进行搜索的操作
    except Exception as e:
        lb.insert(tkinter.END, '登陆失败' + str(e) + '\n')

if __name__ == '__main__':

    # 初始化界面
    root.columnconfigure(0, weight=1)
    root.columnconfigure(1, weight=4)

    group = tkinter.Frame(root).grid(row=0, column=0, columnspan=4, sticky=tkinter.W)

    company = tkinter.Label(group, text='公司编号:', anchor='c').grid(row=0, column=0, sticky=tkinter.S)

    companyIn = tkinter.Entry(group, width=60)
    companyIn.grid(row=0, column=1, columnspan=3, sticky=tkinter.W)

    keyword = tkinter.Label(group, text='关键词:', anchor='c').grid(row=1, column=0, sticky=tkinter.S)
    keywordIn = tkinter.Entry(group, width=60)
    keywordIn.grid(row=1, column=1, columnspan=3, sticky=tkinter.W)

    btnnGroup = tkinter.LabelFrame(root, bg='black', pady=15,padx = 15).grid(row=2)
    btnLogin = tkinter.Button(root, text="登陆账户", command=loginYouUsername).grid(row=2, column=0)
    btnSearch = tkinter.Button(root, text="搜索数据", command=executeSearch).grid(row=2, column=1)
    btn = tkinter.Button(root, text="导出Excel", command=write_to_excel).grid(row=2, column=2)
    margineright = tkinter.Frame(root, height=10,width =30)
    margineright.grid(row=2, column=3)

    print(type(btn))

    margine = tkinter.LabelFrame(root, height=10)
    margine.grid(row=3)

    lb = scrolledtext.ScrolledText(root, width=600, height=46)
    lb.grid(row=4, column=0, columnspan=4, sticky=tkinter.W + tkinter.S + tkinter.N)
    lb.insert('1.0', '请输入公司编码和关键词\n')
    # lb.configure(bg = 'green')
    # lb.configure(state=tkinter.DISABLED)
    # print(type(lb))

    # root的设置
    root.title('搜索并导出Excel')
    center_window(root, 600, 700)
    lb.see(tkinter.END)
    root.mainloop()









