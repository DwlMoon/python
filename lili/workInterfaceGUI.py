import threading
import traceback
from tkinter import *
from lili import login
import encodings.idna


class Application(Frame):
    try:
        def __init__(self, master=None):
            Frame.__init__(self, master)
            self.pack()
            self.createWidgets()

        def createWidgets(self):
            account_name = Label(self, text="Excel路径 :")
            account_name.pack()
            self.myExcel = Entry(self, bd=5, width=40)
            self.myExcel.pack()
            pass_name = Label(self, text="学校 :")
            pass_name.pack()
            self.school = Entry(self, bd=5, width=40)
            self.school.pack()
            pass_name = Label(self, text="专业 :")
            pass_name.pack()
            self.professional = Entry(self, bd=5, width=40)
            self.professional.pack()
            self.alertButton = Button(self, text='开始登陆', command=self.hello, width=15)
            self.alertButton.pack()

        def hello(self):

                myExcel = self.myExcel.get()
                # myExcel = 'D:\lili.xlsx'
                school = self.school.get()
                professional = self.professional.get()

                # login.login.read_xlrd(login,myExcel,school,professional)

                t = threading.Thread(target=login.login.read_xlrd, args=(login,myExcel,school,professional))
                # 守护线程
                t.setDaemon(True)
                # 启动线程
                t.start()

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
app.master.title('Login')
app.master.geometry('400x250')
# logo = PhotoImage(file='E:/pythonProject/work/Ironman.gif')
# Label(app.master,compound=CENTER,image=logo).pack(side="left")

root=app.master

# canvas = Canvas(root,width=400, height=300,bg='black')
# canvas.pack()
# img=[]

# tmp = open('tmp.gif', 'wb+')  # 临时文件用来保存gif文件
# tmp.write(base64.b64decode(logo))
# tmp.close()

# def resource_path(relative_path):
#     if getattr(sys, 'frozen', False): #是否Bundle Resource
#         base_path = sys._MEIPASS
#     else:
#         base_path = os.path.abspath(".")
#     return os.path.join(base_path, relative_path)


#分解gif并逐帧显示
# def pick(event):
#     global a,flag
#     while 1:
#         # bundle_dir = getattr(sys, '_MEIPASS', os.path.abspath(os.path.dirname(__file__)))
#         # path = os.path.join(bundle_dir, 'tony.gif')
#         filename = resource_path(os.path.join("images", "tmp.gif"))
#         # filename = resource_path("tmp.gif")
#         print("*" * 10)
#         print(filename)
#         im = Image.open(filename)
#         # im = Image.open(path)
#         # im = Image.open('tmp.gif')
#         # GIF图片流的迭代器
#         iter = ImageSequence.Iterator(im)
#         #frame就是gif的每一帧，转换一下格式就能显示了
#         for frame in iter:
#             pic=ImageTk.PhotoImage(frame)
#             canvas.create_image((200,150), image=pic)
#             time.sleep(0.1)
#             root.update_idletasks()  #刷新
#             root.update()
#
# canvas.bind("<Enter>",pick)

# 主消息循环:
app.mainloop()


if __name__ == '__main__':
    Application.hello(None)
