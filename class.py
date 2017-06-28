# -*- coding:utf-8 -*-
from selenium import webdriver
from PIL import Image
import re,urllib2, os, cookielib, threading, time,xlwt
from bs4 import BeautifulSoup
from wxbot import *

class Grade(object):
    check_code=''
    __myid=''
    __mypassword=''
    check_code_path=''
    html_cj=''
    grade_lxs_path=''
    login_flag=False
    heads = {'Connection': 'keep-alive',
             'User-Agent': 'Mozilla/5.0 (Windows NT 6.3; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.96 Safari/537.36',
             'Accept-Encoding': 'gzip, deflate, sdch', 'Accept-Language': 'zh-CN,zh;q=0.8,en-US;q=0.6,en;q=0.4'}
    def __init__(self,id,ps):
        self.__myid=id
        self.__mypassword=ps
    def keep_cookies(self):
        cookies1 = cookielib.LWPCookieJar()
        cookies1_support = urllib2.HTTPCookieProcessor(cookies1)
        openner = urllib2.build_opener(cookies1_support, urllib2.HTTPHandler)
        urllib2.install_opener(openner)
    def get_yzm(self):
        driver = webdriver.PhantomJS(
        executable_path=os.getcwd()+r'\phantomjs.exe')
        driver.get('http://my.hlju.edu.cn')
        data = driver.page_source
        flag_yzm1 = re.search('<img id="captchaImg" src="', data).end()
        flag_yzm2 = re.search('" title="" class="ipt1">', data).start()
        yzm_url = data[flag_yzm1:flag_yzm2]
        yzm_url = 'http://my.hlju.edu.cn/captchaGenerate.portal?' + yzm_url
        method = 'GET'
        respone = urllib2.Request(yzm_url, headers=self.heads)
        imadata = urllib2.urlopen(respone)
        ima_data = imadata.read()
        ima_path = os.getcwd()
        ima_wenjian = open(ima_path + r'\yzm1.jpg', 'wb')
        ima_wenjian.write(ima_data)
        ima_wenjian.close()
        ima_path = ima_path + r'\yzm1.jpg'
        time.sleep(1)
        driver.quit()
        self.check_code_path=ima_path
    def login(self,yzm_text):
        user = self.__myid
        password = self.__mypassword
        login_data = 'Login.Token1=' + user + '&Login.Token2=' + password + '&captcha=' + yzm_text + '&goto=http%3A%2F%2Fmy.hlju.edu.cn%2FloginSuccess.portal&gotoOnFail=http%3A%2F%2Fmy.hlju.edu.cn%2FloginFailure.portal'
        host1 = 'http://my.hlju.edu.cn/userPasswordValidate.portal'
        method = 'POST'
        check_yzmr = urllib2.Request(host1, data=login_data, headers=self.heads)
        result_yzm = urllib2.urlopen(check_yzmr)
        print result_yzm.read()
        host = 'http://my.hlju.edu.cn/index.portal'
        index_data = urllib2.urlopen(host)
        host2 = 'http://ssfw3.hlju.edu.cn/ssfw/j_spring_ids_security_check'
        security_check = urllib2.urlopen(host2)
        heads1 = {'Connection': 'keep-alive',
                  'User-Agent': 'Mozilla/5.0 (Windows NT 6.3; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.96 Safari/537.36',
                  'Accept-Encoding': 'gzip, deflate, sdch', 'Accept-Language': 'zh-CN,zh;q=0.8,en-US;q=0.6,en;q=0.4',
                  'Referer': 'http://ssfw3.hlju.edu.cn/ssfw/j_spring_ids_security_check'}
        host3 = 'http://ssfw3.hlju.edu.cn/ssfw/index.do?from='
        cxcj1_pequest = urllib2.Request(host3, headers=heads1)
        cxcj1 = urllib2.urlopen(cxcj1_pequest)
        host_cj = 'http://ssfw3.hlju.edu.cn/ssfw/zhcx/cjxx.do'
        cj_request = urllib2.Request(host_cj, headers=self.heads)
        cj = urllib2.urlopen(cj_request)
        self.html_cj = cj.read()
    def save(self):
        html = str(self.html_cj)
        head_list = ['序号', '学年学期', '课程号', '课程名称', '课程类别', '课程性质', '学分', '成绩', '修读方式', '备注']
        workbook = xlwt.Workbook(encoding='utf-8')
        worksheet = workbook.add_sheet(u'成绩结果', cell_overwrite_ok=True)
        count = -1
        for i in head_list:
            count = count + 1
            worksheet.write(0, count, i)
        kc_start1 = re.search('<td>合计：</td>', html).end()
        kc_start1 = re.search('<td><strong>', html).end()
        kc_end1 = re.search('</strong></td>', html).start()
        kc_count = html[kc_start1:kc_end1]
        list_information = {}
        soup = BeautifulSoup(self.html_cj, "html.parser")
        list_information = list(soup.stripped_strings)
        count = 0
        list_count = []
        for count1 in range(1, int(kc_count) + 1):
            list_count.append(list_information.index(str(count1)))
        for ii in range(0, len(list_count) - 1):
            count2 = -1
            for iii in list_information[list_count[ii]:list_count[ii + 1]]:
                count2 += 1
                worksheet.write(ii + 1, count2, list_information[list_count[ii]:list_count[ii + 1]][count2])

        list_information[list_count[int(kc_count) - 1]:list_information.index(
            u'\u6ce8\uff1a"\'\u2014\u2014\'"\u6807\u6ce8\u7684\u8bfe\u7a0b\u53f7\u53ca\u8bfe\u7a0b\u540d\u4e3a\u7ecf\u8fc7\u66ff\u4ee3\u5904\u7406\u540e\u7684\u539f\u59cb\u8bfe\u7a0b\u3002')]
        count2 = -1
        for ii in list_information[list_count[int(kc_count) - 1]:list_information.index(
                u'\u6ce8\uff1a"\'\u2014\u2014\'"\u6807\u6ce8\u7684\u8bfe\u7a0b\u53f7\u53ca\u8bfe\u7a0b\u540d\u4e3a\u7ecf\u8fc7\u66ff\u4ee3\u5904\u7406\u540e\u7684\u539f\u59cb\u8bfe\u7a0b\u3002')]:
            count2 += 1
            worksheet.write(len(list_count), count2, ii)
        workbook.save(os.getcwd() + r"\grade.xls")
        self.grade_lxs_path = os.getcwd() + r"\grade.xls"
a= Grade('20153916','wangjian687897')
a.keep_cookies()
class MyWXBot(WXBot):
    def handle_msg_all(self, msg):
        global send_yzm_flag
        if msg['msg_type_id'] == 4 and msg['content']['type'] == 0:
            if msg['content']['data']==u'查询成绩':
                self.send_msg_by_uid(u'稍等！',msg['user']['id'])
                self.send_msg_by_uid(u'请输入验证码:', msg['user']['id'])
                a.get_yzm()
                yzm=a.check_code_path
                self.send_img_msg_by_uid(yzm,msg['user']['id'])
                send_yzm_flag=True
            elif (send_yzm_flag==True):
                yzm_text=msg['content']['data']
                a.login(yzm_text)
                a.save()
                self.send_file_msg_by_uid(a.grade_lxs_path,msg['user']['id'])
                send_yzm_flag=False
bot=MyWXBot()
bot.DEBUG=False
bot.conf['qr']='png'
bot.run()