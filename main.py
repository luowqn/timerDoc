#!/usr/bin/python
# -*- coding:utf-8 -*-
# Author:luowq
# Data: 2020/3/24 21:57
from selenium import webdriver
from chinese_calendar import is_holiday
from apscheduler.schedulers.background import BackgroundScheduler
from imapclient import IMAPClient
import datetime
import yaml
import time
import os
import email
import re
from selenium.webdriver.chrome.options import Options
import traceback
import logging.handlers

gConf = None
gLog = None

LOG_TYPE={
    "warn":logging.WARNING,
    "error":logging.ERROR,
    "info":logging.INFO,
    "debug":logging.DEBUG
}

class SwLog(object):
    '''
    默认按天生成日志，保留3天的日志
    初始化入参：
    （1）when：S：秒  M：分 H：小时 D：天 W：星期
    （2）backCount：保留日志个数
    '''
    def __init__(self, filename, level, when='D', backCount=3):
        self.logger = logging.getLogger(filename)
        format = logging.Formatter('%(asctime)s [%(filename)s:%(lineno)d] %(levelname)s %(message)s')
        self.logger.setLevel(LOG_TYPE.get(level.lower()))

        stream = logging.StreamHandler()
        stream.setFormatter(format)

        handler = logging.handlers.TimedRotatingFileHandler(filename=filename, when=when, backupCount=backCount,
                                                            encoding='utf-8')
        handler.setFormatter(format)

        self.logger.addHandler(stream)
        self.logger.addHandler(handler)

    def warn(self, msg, *args, **kwargs):
        self.logger.warning(msg, *args, **kwargs)

    def info(self, msg, *args, **kwargs):
        self.logger.info(msg, *args, **kwargs)

    def error(self, msg, *args, **kwargs):
        self.logger.error(msg, *args, **kwargs)

    def debug(self, msg, *args, **kwargs):
        self.logger.debug(msg, *args, **kwargs)

class basePage(object):

    def __init__(self,driver):
        self.driver = driver

    def wait(self,seconds):
        self.driver.implicitly_wait(seconds)

    def clear(self,xpath):
        self.driver.find_element_by_xpath(xpath).clear()

    def quit(self):
        try:
            self.driver.quit()
        except NameError as e:
            raise e

    def switchToFrame(self,frame):
        try:
            self.driver.switch_to_frame('login_frame')
        except Exception as ex:
            raise  ex

    def switchToParentFrame(self):
        try:
            self.driver.switch_to.parent_frame()
        except Exception as ex:
            raise  ex

    def click(self, xpath):
        try:
            self.driver.implicitly_wait(3)
            self.driver.find_element_by_xpath(xpath).click()
            time.sleep(2)
        except Exception as ex:
            raise ex

    def inputData(self,xpath,data):
        try:
            self.driver.implicitly_wait(2)
            self.driver.find_element_by_xpath(xpath).send_keys(data)
        except Exception as ex:
            raise ex

    def selectData(self,xpath,data):
        try:
            eles = self.driver.find_elements_by_xpath(xpath)
            for ele in eles:
                if data == ele.text:
                    ele.click()
                    break
            time.sleep(1)
        except Exception as ex:
            raise ex

    def getEles(self,xpath):
        try:
            self.driver.implicitly_wait(3)
            return self.driver.find_elements_by_xpath(xpath)
        except Exception as ex:
            raise ex

class SwTimer(object):

    def __init__(self):
        self.scheduler = BackgroundScheduler()

    def getJob(self,jobId):
        return self.scheduler.get_job(jobId)

    def stop_job(self,jobId):
        self.scheduler.remove_job(jobId)

    def add_crond(self,jobId,jobFunc,hour,minute,*args):
        self.scheduler.add_job(func=jobFunc,trigger='cron',id=jobId,day_of_week='mon-sun',hour=hour,minute=minute)

    def add_interval(self,jobId,jobFunc,days=1,*args):
        self.scheduler.add_job(jobFunc, 'interval', days=days,id=jobId, args=args)

    def run(self):
        self.scheduler.start()

def WriteMsg(fPage,name,message):
    global gLog
    try:
        fPage.switchToParentFrame()
        eles = fPage.getEles("//div[@class='submit-suc-icon']")
        if len(eles)  >0:
            gLog.info("今天填写过健康表啦")
            return
        fPage.click('//span[@class="rc-select-arrow"]')
        fPage.wait(5)
        fPage.click('//ul[@role="listbox"]/li[contains(text(),"平台中心")]')
        fPage.wait(2)
        fPage.inputData("//input[@aria-label='请填写姓名']",name)
        fPage.wait(2)
        elesInput  = fPage.getEles('//input[@tabindex="0"]')
        for ele in elesInput[1:]:
            ele.send_keys('否')
        elesSelect = fPage.getEles('//div[@class="question-content-radio-normal"]')
        for ele in elesSelect:
            ele.click()
        fPage.inputData('//div[contains(text(),"接触史")]/../../../div[@class="question-content"]/div[@class="question-content-rich"]/textarea[@placeholder="请填写"]',message)
        fPage.click('//button[contains(text(),"提交")]')
        fPage.wait(5)
        fPage.click('//button[contains(text(),"确认")]')
    except Exception:
        gLog.error(traceback.format_exc())
        pass

def Login(fPage,driver, user,pwd):
    global  gLog
    try:
        fPage.click('//button[contains(text(),"登录")]')
        fPage.switchToFrame('login_frame')
        fPage.click('//*[@id="switcher_plogin"]')
        fPage.inputData('//*[@id="u"]',user)
        fPage.inputData('//*[@id="p"]', pwd)
        fPage.click('//*[@id="login_button"]')
    except Exception:
        gLog.error(traceback.format_exc())
        pass

def openBrowser(url):
    global gLog
    try:
        browserPath = os.path.abspath(os.path.dirname(__file__))
        chrome_options = Options()
        chrome_options.add_argument('--headless')
        chrome_options.add_argument('--disable-gpu')
        driver = webdriver.Chrome(os.path.join(browserPath, './drivers/chromedriver'),chrome_options=chrome_options)
        driver.get(url)
        driver.maximize_window()
        driver.implicitly_wait(10)
        return driver
    except Exception:
        gLog.error(traceback.format_exc())

def get_email_content(message):
    # 获取邮件的所有文本内容
    contents = []
    maintype = message.get_content_maintype()

    msgs = message.get_payload()
    if maintype == 'multipart':
        for msg in msgs:
            msg_type = msg.get_content_maintype()
            if msg_type == 'text':
                msg_con_type = msg.get('Content-Type')
                # 并不是所有的文本信息都在text/plain中，所以这里取text/html
                if msg_con_type and 'text/html' in msg_con_type:
                    con1 = msg.get_payload(decode=True).strip()
                    contents.append(str(con1, encoding='utf-8'))
            elif msg_type == 'multipart':
                for item in msg.get_payload():
                    item_tyep = item.get_content_maintype()
                    if item_tyep == 'text':
                        item_con_type = item.get('Content-Type')
                        if item_con_type and 'text/html' in item_con_type:
                            con2 = item.get_payload(decode=True).strip()
                            contents.append(str(con2, encoding='utf-8'))
                    else:
                        print('wrong main type: %s' % item_tyep)
            else:
                pass
    elif maintype == 'text':
        m_type = message.get('Content-Type')
        if m_type and 'text/html' in m_type:
            con3 = message.get_payload(decode=True).strip()
            contents.append(str(con3, encoding='utf-8'))
    else:
        pass

    conts = '\n'.join(contents)

    return conts

def getAddr():
    addr = None
    try:
        with IMAPClient('imap.exmail.qq.com', use_uid=True) as server:
            # 输入用户名和密码
            server.login("luowq@signalway.com.cn", "Bjt11.28")
            select_info = server.select_folder('其他文件夹/健康表')
            # 查询来至指定账号的邮件
            email_uids = server.search(['SINCE', '05-March-2020'])
            msgdict = server.fetch(email_uids[-1], ['BODY[]'])
            mailbody = msgdict[email_uids[-1]][b'BODY[]']
            message = email.message_from_string(str(mailbody, encoding='utf-8'))
            contents = get_email_content(message)
            matchObj = re.search(r'(https://docs.qq.com/form/fill/?.*?_w_tencentdocx_form=1)', contents, re.M | re.I)
            if matchObj:
                addr = matchObj.group(1)
            server.close_folder()
            server.logout()
    except Exception as ex:
        raise ex
    finally:
        return addr

def initLog(file,loglevel):
    global gLog
    try:
        gLog = SwLog(file,loglevel)
    except Exception as ex:
        raise ex

def initConf(file):
    global gConf
    with open(file,encoding='utf-8') as f:
        gConf = yaml.safe_load(f)

def jobFunc():
    global  gLog
    try:
        gLog.info("任务开始执行")
        addr = getAddr()
        gLog.info("获取健康表地址:%s"%addr)
        if addr:
            driver = openBrowser(addr)
            fPage = basePage(driver)
            Login(fPage, driver, gConf['user'], gConf['passwd'])
            gLog.info("今天登录成功啦")
            if is_holiday(datetime.datetime.now()):
                WriteMsg(fPage, gConf['name'], '在家')
            else:
                WriteMsg(fPage, gConf['name'], '上班')
            driver.quit()
        gLog.info("今天填表任务完成啦，Yeah!")
    except Exception:
        gLog.error(traceback.format_exc())


if __name__ == '__main__':
    initConf('config.yaml')
    initLog('timerDoc.log','INFO')
    gLog.info("定时任务开启，每天晚上七点半填写健康表")
    t = SwTimer()
    t.add_crond('1',jobFunc,'19','30',[gConf])
    t.run()
    while True:
        time.sleep(1)


