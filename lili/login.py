
import _thread
import xlrd
from selenium import webdriver
import time
from selenium.webdriver.common.keys import Keys

from selenium.webdriver.chrome.options import Options


class login:

    # 均分数组
    def list_split(items, n):
        return [items[i:i + n] for i in range(0, len(items), n)]


    # 登录页面
    def loginWeb(self,accountlist):

        # 定义chromedriver驱动的位置
        # baigalama
        chromedriver = r"C:\Program Files (x86)\Google\Chrome\Application\chromedriver.exe"

        options = Options()
        options.add_argument("--incognito")

        # 打开一个浏览器窗口
        driver = webdriver.Chrome(chrome_options=options)

        url = "http://jinanzyk.36ve.com/login/login"

        password = 'znzd@123456'

        list2 = login.list_split(accountlist,10)

        for i in list2:

            for ac in i:

                try:
                    # 发送请求
                    driver.get(url)
                    time.sleep(4)

                    driver.find_element_by_id("loginform-username").send_keys(ac)
                    time.sleep(2)

                    driver.find_element_by_id("loginform-password").send_keys(password)
                    time.sleep(2)

                    driver.find_element_by_xpath('//*[@id="login-form-1"]/div[3]/button').click()
                    time.sleep(4)
                    nowurl = driver.current_url

                    if nowurl == url:
                        continue
                    else:
                        print("登录成功")
                        js = "window.open('http://www.baidu.com')"
                        driver.execute_script(js)
                        driver.switch_to.window(driver.window_handles[-1])
                        driver.delete_all_cookies()

                except Exception as e:
                    continue

            time.sleep(300)






    # 读数据
    def read_xlrd(excelFile):
        data = xlrd.open_workbook('E:\\456.xls')

        print("工作表为：" + str(data.sheet_names()))
        list=data.sheet_names()

        table = data.sheet_by_name(list[0])

        print("总行数：" + str(table.nrows))
        print("总列数：" + str(table.ncols))
        line = int(table.nrows)

        countList=[]

        for l in range(1, line):
            account = str(table.cell(l, 0).value)
            print("第%s行第一列的值为: %s" % (l, account))
            countList.append(account)

        login.loginWeb(login,countList)

if __name__=='__main__':
    # login.loginWeb(login,'16638335117','Long6518132')
    login.read_xlrd(None)

