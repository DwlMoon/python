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

        url = "https://live.aliexpress.com/livebackend/liveRoomList.htm?spm=a2g0t.10516576.0.0.9de8450c7WhQZ3"
        # url = "https://login.aliexpress.com/buyer.htm"

        # 发送请求
        driver.get(url)
        time.sleep(4)

        driver.find_element_by_id("fm-login-id").send_keys(email_account)
        time.sleep(2)
        driver.find_element_by_id("fm-login-password").send_keys(pwd_account)
        time.sleep(2)
        driver.find_element_by_class_name('fm-button').click()
        time.sleep(3)
        print("登录成功")
        Zhweb.nextStartRowKey = 1
        # print("============================================================")
        return driver

    # 获取数据
    def write_file(self,email_account,pwd_account, mylist, mydriver):
        try:
            if Zhweb.loginflag==True:
                # mylist=Zhweb.mylist
                # nextStartRowKey=Zhweb.nextStartRowKey
                mydriver = Zhweb.loginWeb(None,email_account,pwd_account)
                Zhweb.loginflag=False
            sChildPath=None
            time.sleep(3)
            mydriver.get("https://live.aliexpress.com/livebackend/AjaxGetLiveList.do?nextStartRowKey=" + str(Zhweb.nextStartRowKey))
            time.sleep(5)
            json_text = mydriver.find_element_by_tag_name("pre").text
            # print(json_text)
            val_dict = json.loads(json_text)
            # print("================================")


            # find_index = email_account.find("@")

            # account_name = email_account[0:find_index]

            a=email_account.split("@")
            print(a)
            account_name = a[0].replace('\n', '').replace('\r', '').strip()

            # print(account_name)
            base_path = r"C:/findRoomId"
            a=[account_name,".xls"]
            filename ="".join(a)
            print(filename)
            sChildPath = os.path.join(base_path,filename)
            # sChildPath = r"C:/findRoomId/"+ account_name + ".xls"
            print("文件路径为 : %s "%sChildPath)


            if len(val_dict['body']['list']) > 0:
                json_str = json.dumps(val_dict)
                data = json.loads(json_str, object_hook=lambda d: namedtuple('data', d.keys())(*d.values()))
                # print(data)
                ownlist = data.body.list
                mylist += ownlist

                # 保存数据
                Zhweb.nextStartRowKey += 1
                mydriver.get("https://live.aliexpress.com/livebackend/AjaxGetLiveList.do?nextStartRowKey=" + str(Zhweb.nextStartRowKey))
                Zhweb.write_file(Zhweb,email_account,pwd_account,mylist,mydriver)
            else:
                print("已加载全部")
                print("===================================")
                workbook = xlwt.Workbook(encoding='utf-8')
                # 创建一个worksheet
                worksheet = workbook.add_sheet('Worksheet')
                worksheet.write(0, 0, label='直播间ID')
                worksheet.col(0).width = 256 * 20
                worksheet.write(0, 1, label='标题')
                worksheet.col(1).width = 256 * 30
                worksheet.write(0, 2, label='开始时间(美西时间)')
                worksheet.col(2).width = 256 * 30
                worksheet.write(0, 3, label='开始时间(北京时间)')
                worksheet.col(3).width = 256 * 30
                worksheet.write(0, 4, label='结束时间(美西时间)')
                worksheet.col(4).width = 256 * 30
                worksheet.write(0, 5, label='结束时间(北京时间)')
                worksheet.col(5).width = 256 * 30
                worksheet.write(0, 6, label='直播链接')
                worksheet.col(6).width = 256 * 60
                worksheet.write(0, 7, label='审核状态')
                worksheet.col(7).width = 256 * 15
                worksheet.write(0, 8, label='是否直播过')
                worksheet.col(8).width = 256 * 15
                workbook.save(sChildPath)

                data = xlrd.open_workbook(sChildPath, formatting_info=True)
                excel = copy(wb=data)  # 完成xlrd对象向xlwt对象转换
                excel_table = excel.get_sheet(0)

                writelist=[]
                for i in range(0, len(mylist)):


                    # if mylist[i].status<18:
                    #     writelist.append(mylist[i])
                    # 获取直播过的状态
                    if mylist[i].status==18:
                        writelist.append(mylist[i])


                for i in range(0, len(writelist)):

                        line = 0
                        print("第%s个房间"%i)
                        excel_table.write(i+1, int(line),str(writelist[i].liveId) )
                        print("房间号为 : %s"%writelist[i].liveId)
                        line += 1

                        excel_table.write(i+1, int(line),writelist[i].title )
                        print("标题为 : %s"%writelist[i].title)
                        line += 1

                        excel_table.write(i+1, int(line),writelist[i].startTime )
                        print("开始时间(美西时间) : %s"%writelist[i].startTime)
                        line += 1

                        time1 = parser.parse(writelist[i].startTime)
                        time2 = time1 + datetime.timedelta(hours=16)
                        excel_table.write(i+1, int(line),time2.strftime("%Y-%m-%d %H:%M:%S"))
                        print("开始时间(北京时间) : %s"%writelist[i].startTime)
                        line += 1

                        excel_table.write(i+1, int(line),writelist[i].endTime)
                        print("结束时间(美西时间) : %s"%writelist[i].endTime)
                        line += 1

                        time1 = parser.parse(writelist[i].endTime)
                        time2 = time1 + datetime.timedelta(hours=16)
                        excel_table.write(i+1, int(line),time2.strftime("%Y-%m-%d %H:%M:%S"))
                        print("结束时间(北京时间) : %s"%writelist[i].endTime)
                        line += 1

                        excel_table.write(i + 1, int(line), writelist[i].liveUrl)
                        print("直播链接 : %s" % writelist[i].liveUrl)

                        audit_str = ""
                        if writelist[i].audit == 0:
                            audit_str = "not apply audit"
                        elif writelist[i].audit == 1:
                            audit_str = "in review"
                        elif writelist[i].audit == 2:
                            audit_str = "reviewed"
                        line += 1
                        excel_table.write(i + 1, int(line), audit_str)
                        print("审核状态 : %s" % audit_str)

                        live_Status = ""
                        if writelist[i].status == 16:
                            live_Status = "not live"
                        elif writelist[i].status == 18:
                            live_Status= "live"
                        else:
                            live_Status=writelist[i].status
                        line += 1
                        excel_table.write(i + 1, int(line), live_Status)
                        print("审核状态 : %s" % live_Status)
                        line = 0

                excel.save(sChildPath)
                print("===================================")
                print("excel数据追加保存成功")

                mydriver.close()
                mydriver.quit()


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
    # Zhweb.write_file(Zhweb,'jocafariasl@gmail.com','826010',[],None)
    # haha="D:"
    base_path = r"D:/findRoomId"
    a=["456123",".xls"]
    filename = "".join(a)
    print(filename)
    sChildPath = Path(os.path.join(base_path, filename)).as_posix()
    print(sChildPath)

