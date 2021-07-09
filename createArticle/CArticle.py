import datetime
import os
import traceback
from collections import namedtuple
from pathlib import Path

import xlrd
import xlwt
from selenium import webdriver
from xlutils.copy import copy
import time
import json
from dateutil import parser


'获取创建好的直播间Id，直播开始时间，结束时间'

class Zhweb:

    # hasNext = True
    nextStartRowKey = 1
    mylist = []
    loginflag=True
    driver=None

    # 登录页面
    def loginWeb(self,email_account,pwd_account):

        # 定义chromedriver驱动的位置
        # baigalama
        chromedriver = r"C:\Program Files (x86)\Google\Chrome\Application\chromedriver.exe"

        # mita
        # chromedriver = r"C:\Program Files\Google\Chrome\Application\chromedriver.exe"

        # nancy
        # chromedriver = r"C:\Users\ww\AppData\Local\Google\Chrome\Application\chromedriver.exe"


        # 高妞
        # chromedriver = r"C:\Program Files (x86)\Google\Chrome\Application\chromedriver.exe"

        # 打开一个浏览器窗口
        driver = webdriver.Chrome(chromedriver)

        url = "https://we.aliexpress.com/post/home.htm"
        # url = "https://login.aliexpress.com/buyer.htm"

        # 发送请求
        driver.get(url)
        time.sleep(4)

        driver.find_element_by_id("fm-login-id").send_keys(email_account)
        time.sleep(2)
        driver.find_element_by_id("fm-login-password").send_keys(pwd_account)
        time.sleep(2)
        driver.find_element_by_class_name('fm-button').click()
        time.sleep(10)
        print("登录成功")
        Zhweb.nextStartRowKey = 1
        # print("============================================================")
        return driver

    # 获取数据
    def write_file(self):
        try:

            mydriver=Zhweb.loginWeb(Zhweb,"ropa_es_oficial@service.aliyun.com","gret5632hjy")
            time.sleep(10)

            createArtile=mydriver.find_element_by_xpath("/html/body/div[4]/div/div/div/div[2]/div/div/div[1]/div/ul/li[1]/a")
            createArtile.click()
            time.sleep(3)

            mydriver.find_element_by_id("post-title").send_keys("Lo clásico vuelve")
            time.sleep(1)
            mydriver.find_element_by_id("post-summary").send_keys("Este outfit bastante clasico y elegante , el cual nos sirve para ir a una cena formal o para ir a tomar algo por la noche consta de un abrigo bastante grues con un diseño de pelos  con distintos tonos claros y oscuros ; con un toque sensual gracias a la parte de la cintura que no tiene pelos sino que tiene como un cinturon que bordea esa parte ajustandola a la cintura. De camiseta observamos una camisa blanca con un escote en pico bastante pronunciado que consigue estilizar y a la vez dar un toque sensual. En la parte inferior vemos una falda oscura con algo de vuelo y volumen de triro alto por lo que la camisa va por dentro de la falda para marcar la cintura y vemos que el largo llega hasta por encima de las rodillas. Por ultimo tenemos unos botines de ante con una abertura al principio del botin , el tacon es de aguaja y van adornados con unos detalles plateados en el lado exterior del botin")
            time.sleep(3)

            # addPhoto=mydriver.find_element_by_xpath("//*[@id="image-selected"]/div/a")
            addPhoto=mydriver.find_element_by_id("image-selected")
            time.sleep(2)
            addPhoto.click()
            # mydriver.find_element_by_id("ksu-html5-kkun223d").send_keys(r"E:\abc.jpg")
            mydriver.find_element_by_css_selector("ksu-html5-kkun223d").send_keys(r"E:\abc.jpg")

            time.sleep(50)







        except Exception as e:
            print("====================异常信息=============================")
            print(traceback.format_exc())
            return 'fail'
        except EOFError as eof:
            print("=====================错误信息============================")
            print(traceback.format_exc())
            return 'fail'
        Zhweb.loginflag = True

        return 'success'

if __name__=='__main__':
    Zhweb.write_file(Zhweb)

