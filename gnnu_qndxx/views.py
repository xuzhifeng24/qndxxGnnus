from django.shortcuts import render
from gnnu_qndxx.models import User
from django.shortcuts import HttpResponse
from .models import *
import tkinter
import tkinter as tk
from tkinter import *
import tkinter.messagebox
import threading
from tkinter import filedialog
from tkinter import ttk
import pandas as pd
import tkinter.font as tkFont
import requests
import re
import time
import os
import random
import base64
import json
import math
import sys
import xlwings as xw
import urllib
import pythoncom
from browsermobproxy import Server
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from urllib.parse import urlencode
from bs4 import BeautifulSoup
from tkinter import messagebox
from gnnu_qndxx import frozen_dir

import warnings
warnings.filterwarnings("ignore")

headers = {
    'User-Agent':'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/65.0.3298.4 Safari/537.36'
}

# txt = None
# usr_name = "dlyhjgcxytw_101726"
# usr_pwd = "dlxy9116"
# qishu = '第八期'
# # course = None
# # accessToken = None
# filePath = r'C:\Users\66483\Desktop\全院学生信息.xlsx'
# xy_name = '地理与环境工程学院'
# sys_path = os.path.split(os.path.realpath(__file__))[0]
# accessToken = 'D2D4BE1E-4896-4229-89BA-9D4FDB3E6412'
# course = 'C0045'

txt = None
usr_name = None
usr_pwd = None
qishu = None
course = None
accessToken = None
filePath = None
xy_name = None
sys_path = os.path.split(os.path.realpath(__file__))[0]



def login(request):
    return render(request, "loginMain.html")

def userLogin(request):
    global xy_name
    global usr_name
    global usr_pwd
    xy_name = request.POST.get("college")
    usr_name = request.POST.get("user_name")
    usr_pwd = request.POST.get("password")
    user = User.objects.filter(xy_name=xy_name, use_name=usr_name, password=usr_pwd)
    if user:
        return render(request, "useInfo.html")
    else:
        return HttpResponse("请输入正确的信息")

def get_xy_info():
    qishu_course = {"第一期": "C0038", "第二期": "C0039", "第三期": "C0040", "第四期": "C0041",
                    "第五期": "C0042", "第六期": "C0043", "第七期": "C0044", "第八期": "C0045",
                    "“青年大学习”特辑": "C0046", "第九期": "C0047", "第十期": "C0048", '第十一期': "C0049",
                    "第十二期": "C0050", "第十三期": "C0051", "第十四期": "C0052", "第十五期": "C0053"}
    return qishu_course

def userdownload(request):
    global filePath
    global qishu
    global course
    qishu = request.POST.get("period")
    qishu_course = get_xy_info()
    course = qishu_course[qishu]
    filePath = request.POST.get('file_path')
    index_main()
    return render(request, "download.html")

def index_main():
    def auto_get_info():
        server = Server(frozen_dir.app_path()+r"/browsermob-proxy-2.1.4-bin/browsermob-proxy-2.1.4/bin/browsermob-proxy.bat")
        server.start()
        proxy = server.create_proxy()
        chrome_options = Options()

        chrome_options.add_argument('--proxy-server={0}'.format(proxy.proxy))
        proxy.new_har("ht_list2", options={'captureContent': True})
        dcap = dict(DesiredCapabilities.PHANTOMJS)  # 设置useragent
        dcap['phantomjs.page.settings.userAgent'] = (
            'Mozilla/5.0 (Macintosh; Intel Mac OS X 10.9; rv:25.0) Gecko/20100101 Firefox/25.0 ')  # 根据需要设置具体的浏览器信息c
        url = r'https://jxtw.h5yunban.cn/jxtw-qndxx/admin/login.php'
        browser = webdriver.Chrome(chrome_options=chrome_options)
        browser.get(url)
        usename = browser.find_element_by_xpath('//*[@id="LAY-user-login-username"]')
        usename.clear()
        usename.send_keys(usr_name)
        password = browser.find_element_by_xpath('//*[@id="LAY-user-login-password"]')
        password.clear()
        password.send_keys(usr_pwd)
        form = browser.find_element_by_xpath('//*[@id="LAY-user-login"]/div[1]/div[2]/div[3]/button')
        form.click()

        time.sleep(10)
        result = proxy.har

        url_list = []
        for entry in result['log']['entries']:
            _url = entry['request']['url']
            if "accessToken" and 'course' and 'records' in _url:
                url_list.append(_url)

        url_ = url_list[0]
        content = json.loads(requests.get(url_).text)
        qishu = content['result']['list'][0]['course'][-3:]
        # print('期数', qishu)
        # course = url_.split('&')[4].split('=')[1]
        accessToken = url_.split('&')[5].split('=')[1]
        # with open(sys_path + os.path.sep + "qishu_course.pickle", "wb") as usr_file:
        #     qishu_course = {"第九期": "C0024", "第十期": "C0026", "第十一期": "C0027", "第十二期": "C0028",
        #                     "第十三期": "C0029", "第十四期": "C0030", "第十五期": "C0031", "第十六期": "C0032",
        #                     "第十七期": "C0033", "第十八期": "C0034", "第十九期": "C0035", '第二十期': "C0036",
        #                     "第二十一期": "C0037", "第二十二期": "C0038"}
        #     qishu_course.update({qishu: course})
        #     pickle.dump(qishu_course, usr_file)
        browser.close()
        return accessToken

    def get_content():
        nw_path = os.getcwd()
        os.chdir(frozen_dir.app_path() +  r'/data')
        if os.path.exists(xy_name) is False:
            os.makedirs(xy_name)
        os.chdir(nw_path)
        option = webdriver.ChromeOptions()
        dcap = dict(DesiredCapabilities.PHANTOMJS)  # 设置useragent
        dcap['phantomjs.page.settings.userAgent'] = ('Mozilla/5.0 (Macintosh; Intel Mac OS X 10.9; rv:25.0) Gecko/20100101 Firefox/25.0 ')  # 根据需要设置具体的浏览器信息c
        url = r'https://jxtw.h5yunban.cn/jxtw-qndxx/admin/login.php'
        browser = webdriver.Chrome(options=option)
        browser.maximize_window()
        browser.get(url)
        usename = browser.find_element_by_xpath('//*[@id="LAY-user-login-username"]')
        usename.clear()
        usename.send_keys(usr_name)
        password = browser.find_element_by_xpath('//*[@id="LAY-user-login-password"]')
        password.clear()
        password.send_keys(usr_pwd)
        form = browser.find_element_by_xpath('//*[@id="LAY-user-login"]/div[1]/div[2]/div[3]/button')
        form.click()
        html = requests.get(r'https://jxtw.h5yunban.cn/jxtw-qndxx/cgi-bin/branch-api/course/records?pageSize=20&pageNum=1&'
                    r'desc=createTime&nid=N001300131017&course={0}&accessToken={1}'.format(course, accessToken), verify=False)

        content = BeautifulSoup(html.text, 'lxml').text
        data = json.loads(content)
        qishu = data['result']['list'][0]['course'][-3:]
        now_path = os.getcwd()
        os.chdir(frozen_dir.app_path() +  r'/data/{}'.format(xy_name))
        if os.path.exists(qishu) is False:
            os.makedirs(qishu)
        os.chdir(os.getcwd() + os.path.sep + qishu)
        all_student = int(data['result']['pagedInfo']['total'])
        all_page = math.ceil(all_student / 20)
        out_data = pd.DataFrame(data=[[0, 0, 0, 0, 0]], index=[0], columns=['序号', '时间', '姓名', '学院', '班级'])

        index = 0
        for percent in range(1, all_page + 1):
            link = r'https://jxtw.h5yunban.cn/jxtw-qndxx/cgi-bin/branch-api/course/records?pageSize=20&pageNum={0}&desc=createTime&nid=N001300131017&course={1}&accessToken={2}'.format(
                percent, course, accessToken)
            cont = BeautifulSoup(requests.get(link, verify=False).text, 'lxml').text
            dt = json.loads(cont)
            for i in range(20):
                try:
                    tt = dt['result']['list'][i]
                    out_data.loc[index, '序号'] = index + 1
                    out_data.loc[index, '时间'] = tt['createTime']
                    out_data.loc[index, '姓名'] = tt['cardNo']
                    out_data.loc[index, '学院'] = tt['branchs'][2]
                    out_data.loc[index, '班级'] = tt['branchs'][3]
                    index += 1
                    # print('正在读取第{}个学生的青年大学习数据'.format(index))
                except:
                    pass
        print(out_data)
        out_data.to_excel(os.getcwd() + os.path.sep + '{}.xlsx'.format(qishu), index=False)
        browser.close()
        os.chdir(now_path)

    def process_data():
        student_data = pd.read_excel(filePath)
        pro_data = pd.read_excel(frozen_dir.app_path() +  r'/data/{0}/{1}/{2}.xlsx'.format(xy_name, qishu, qishu))
        pythoncom.CoInitialize()
        xh_xm_dict = {}
        for i, j in zip(student_data['学号'], student_data['姓名']):
            xh_xm_dict[i] = j
        bj_xm_dict = {}
        for b, dt in student_data.groupby(['班级']):
            student_list = list(dt['姓名'].values)
            bj_xm_dict[b] = student_list
        bj_xm_dh_dict = {}
        for y, dk in student_data.groupby(['班级']):
            bj_xm_dh_dict[y] = {}
            for q, p in zip(dk['姓名'], dk['学号']):
                bj_xm_dh_dict[y][q] = p

        pro_data.dropna(inplace=True)
        sum = 0
        su = 0
        un_qus_list = []
        yy = str(time.strftime('%Y_%m_%d_%H_%M_%S', time.localtime()))
        month = str(yy).split('_')[1]
        day = str(yy).split('_')[2]
        hour = str(yy).split('_')[3]
        second = str(yy).split('_')[4]
        writer = pd.ExcelWriter(frozen_dir.app_path() +  r'/data/{0}/{1}/{2}未学习学生_{3}月{4}日{5}时{6}分.xlsx'.format(
            xy_name, qishu, qishu, month, day, hour, second, mode='a'))

        for i, df in pro_data.groupby(['班级']):
            student_list = list(df['姓名'].values)
            out_list = []
            for j in student_list:
                k = re.findall('[\u4e00-\u9fa5]+', j)
                t = re.findall('\d+', j)
                if t:
                    try:
                        out_list.append(xh_xm_dict[int(t[0])])
                    except:
                        pass
                if k:
                    out_list.append(k[0])
            inx = 0
            un_que_student = list(set(bj_xm_dict[i]).difference(set(out_list)))
            output_data = pd.DataFrame(data=[[0, 0, 0]], index=[0], columns=['班级', '未学习团课学生', '学号'])
            if un_que_student:
                un_qus_list.append(i)
                sum += len(un_que_student)
                for s in un_que_student:
                    output_data.loc[inx, '班级'] = i
                    output_data.loc[inx, '未学习团课学生'] = s
                    output_data.loc[inx, '学号'] = bj_xm_dh_dict[i][s]
                    inx += 1

                if su == 0:
                    output_data.to_excel(excel_writer=writer, sheet_name=i, index=False)
                    # worksheet = writer.sheets[i]
                    # worksheet.column_dimensions['A'].width = 33
                    # worksheet.column_dimensions['B'].width = 20
                    # worksheet.column_dimensions['C'].width = 20
                    writer.save()

                else:
                    output_data.to_excel(excel_writer=writer, sheet_name=i, index=False)
                    # worksheet = writer.sheets[i]
                    # worksheet.column_dimensions['A'].width = 33
                    # worksheet.column_dimensions['B'].width = 20
                    # worksheet.column_dimensions['C'].width = 20
                    writer.save()
                su += 1

        writer.save()
        writer.close()

        get_ria()

    def get_ria():
        html = requests.get(r'https://jxtw.h5yunban.cn/jxtw-qndxx/cgi-bin/branch-api/course/statis?nid=N001300131017&course={0}&accessToken={1}'.format(
                course, accessToken), verify=False)
        content = BeautifulSoup(html.text, 'lxml').text
        data = json.loads(content)
        output_data = pd.DataFrame(data=[[0, 0, 0, 0]], index=[0], columns=['团组织', '团组织人数', '学习人数', '学习率'])
        inx = 0
        all_page = len(data['result'])
        for i in range(all_page):
            output_data.loc[inx, '团组织'] = data['result'][i]['title']
            output_data.loc[inx, '团组织人数'] = data['result'][i]['memberCnt']
            output_data.loc[inx, '学习人数'] = data['result'][i]['users']
            output_data.loc[inx, '学习率'] = float(str(data['result'][i]['rate']).strip('%'))
            inx += 1

        output_data.loc[inx+1, '学习率'] = output_data['学习率'].mean()
        output_data.to_excel(frozen_dir.app_path() +  r'/data/{0}/{1}/{2}学习率.xlsx'.format(xy_name, qishu, qishu),
                             index=False)

        writer = pd.ExcelWriter(frozen_dir.app_path() +  r'/data/{0}/{1}/{2}学习率.xlsx'.format(xy_name, qishu, qishu))
        output_data.to_excel(excel_writer=writer, sheet_name='覆盖率', index=False)
        # worksheet = writer.sheets['覆盖率']
        # worksheet.column_dimensions['A'].width = 33
        # worksheet.column_dimensions['B'].width = 15
        # worksheet.column_dimensions['C'].width = 15
        # worksheet.column_dimensions['D'].width = 15
        writer.save()
        writer.close()

    def start_download():
        global accessToken
        accessToken = auto_get_info()
        get_content()

    def start_sava():
        run_sava()

    def run_sava():
        th = threading.Thread(target=process_data)
        th.setDaemon(False)
        th.start()

    start_download()
    start_sava()

# index_main()