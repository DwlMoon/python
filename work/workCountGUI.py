import os
import threading
import time
import traceback
from tkinter import *
import tkinter.messagebox as messagebox
from work import zhweb
from PIL import Image, ImageTk, ImageSequence



class Application(Frame):
    try:
        def __init__(self, master=None):
            Frame.__init__(self, master)
            self.pack()
            self.createWidgets()

        def createWidgets(self):
            account_name = Label(self, text="账号 :")
            account_name.pack()
            self.account = Entry(self, bd=5, width=30)
            self.account.pack()
            pass_name = Label(self, text="密码 :")
            pass_name.pack()
            self.password = Entry(self, bd=5, width=30)
            self.password.pack()
            self.alertButton = Button(self, text='查找', command=self.hello ,width=10)
            self.alertButton.pack()

        def hello(self):


                acount = self.account.get()
                password = self.password.get()
                print("获取的账号为:%s"% acount)
                print("获取的密码为:%s"% password)
                result=zhweb.Zhweb.write_file(zhweb.Zhweb,acount,password,[],None)

                t = threading.Thread(target=zhweb.Zhweb.write_file, args=(zhweb.Zhweb,acount,password,[],None))
                # 守护线程
                t.setDaemon(True)
                # 启动线程
                t.start()


                # result=zhweb.Zhweb.loginWeb(zhweb.Zhweb,acount,password)
                # if 'success' is result:
                #     messagebox.showinfo('Message', '您输入的账号： %s 直播间查找成功 !' % acount)
                # elif "fail" is result:
                #     messagebox.showinfo('Message', '您输入的账号： %s 直播间查找失败 !' % acount)

    except Exception as e:
        print("====================异常信息=============================")
        print(traceback.format_exc())
    except EOFError as eof:
        print("=====================错误信息============================")
        print(traceback.format_exc())


app = Application()
# 设置窗口标题:
app.master.title('查找速卖通直播间（直播过）')
app.master.geometry('400x400')
# logo = PhotoImage(file='E:/pythonProject/work/Ironman.gif')
# Label(app.master,compound=CENTER,image=logo).pack(side="left")

root=app.master

canvas = Canvas(root,width=400, height=300,bg='black')
canvas.pack()
img=[]

# tmp = open('tmp.gif', 'wb+')  # 临时文件用来保存gif文件
# tmp.write(base64.b64decode(logo))
# tmp.close()

def resource_path(relative_path):
    if getattr(sys, 'frozen', False): #是否Bundle Resource
        base_path = sys._MEIPASS
    else:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


#分解gif并逐帧显示
def pick(event):
    global a,flag
    while 1:
        # bundle_dir = getattr(sys, '_MEIPASS', os.path.abspath(os.path.dirname(__file__)))
        # path = os.path.join(bundle_dir, 'tony.gif')
        filename = resource_path(os.path.join("images", "tmp.gif"))
        # filename = resource_path("tmp.gif")
        print("*" * 10)
        print(filename)
        im = Image.open(filename)
        # im = Image.open(path)
        # im = Image.open('tmp.gif')
        # GIF图片流的迭代器
        iter = ImageSequence.Iterator(im)
        #frame就是gif的每一帧，转换一下格式就能显示了
        for frame in iter:
            pic=ImageTk.PhotoImage(frame)
            canvas.create_image((200,150), image=pic)
            time.sleep(0.1)
            root.update_idletasks()  #刷新
            root.update()

canvas.bind("<Enter>",pick)

# 主消息循环:
app.mainloop()


if __name__ == '__main__':
   Application.hello(None)
