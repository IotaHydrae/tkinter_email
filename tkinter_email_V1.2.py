# -*- coding:utf-8 -*-
from tkinter import *
from email.header import Header
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import smtplib
from tkinter import filedialog
from tkinter import messagebox

class Application(Frame):
    """Build the basic window frame template"""

    def __init__(self, master):
        super(Application, self).__init__(master)
        # 使用网格布局
        self.grid()
        self.create_widgets()


    # 清空邮件内容事件
    def clear(self):
        self.message.delete('1.0', END)
    def clear_attchments_method(self):
        need_Del = self.label_attachments.curselection()
        self.label_attachments.delete(ACTIVE)
        attach_list.remove(attach_list[need_Del[0]])
    # 窗口实现方法
    def create_widgets(self):
        # 菜单
        Menubar = Menu(self)
        filemenu = Menu(self)
        filemenu.add_command(label='打开', command=self.openfile)
        Menubar.add_cascade(label='文件', menu=filemenu)
        Menubar.add_command(label='帮助', command=self.showhelp)
        Menubar.add_command(label='测试', command=self.test)
        Menubar.add_command(label='退出', command=self.quit)
        # 主窗体中加载菜单
        root.config(menu=Menubar)

        # 程序名
        self.label_title = Label(self, text='快速邮件')
        self.label_title.grid(row=0, columnspan=3)
        
        # 收件人
        self.label_to = Label(self, text='输入收件人邮箱地址：')
        self.label_to.grid(row=2, column=0)
        self.to = Entry(self)
        # self.to.insert(END, "1657802074@qq.com")
        self.to.grid(row=2, column=1, sticky=W)
        # self.to.focus_set()

        # 标题
        self.label_subject = Label(self, text='输入标题：')
        self.label_subject.grid(row=3, column=0)
        self.subject = Entry(self)
        # self.subject.insert(END, "This is a test email")
        self.subject.grid(row=3, column=1, sticky=W)
        self.label_subject_info = Label(self, text='(至少输入5个字)')
        self.label_subject_info.grid(row=3, column=2)

        # 内容
        self.label_message = Label(self, text='在下方输入要发送的消息内容:')
        self.label_message.grid(row=4, column=0)
        self.message = Text(self, width=50, height=10)
        # self.message.insert(END, "快速邮件程序测试内容")
        self.message.grid(row=5, column=0, columnspan=2)
        self.empty = Button(self,text='清空消息', command=self.clear)
        self.empty.grid(row=4, column=2)

        # 添加附件按钮
        self.packbutton = Button(self, text='添加附件', command=self.add_attach)
        self.packbutton.grid(row=6, column=2, sticky=E)

        # 发送按钮
        self.button_send = Button(self, text='立即发送', command=self.send)
        self.button_send.grid(row=6, column=0, sticky=W)

        # 附件列表
        self.label_attachment = Label(self, text='附件列表')
        self.label_attachment.grid(row=8, column=0)
        self.label_attachments = Listbox(self)
        self.label_attachments.grid(row=9, column=0)
        self.clear_attchments = Button(self, text='删除指定附件', command=self.clear_attchments_method)
        self.clear_attchments.grid(row=8, column=1)

    # 打开文件，读取文件内容到窗口的邮件内容
    def openfile(self):
        print('opening')
        filename = filedialog.askopenfilename(title='选择txt文件', filetypes=[('Text', '*.txt')])
        file = open(filename, 'r')
        text = file.read()
        self.message.insert(END, text)
    
    def showhelp(self):
        print('showing')

    def add_attach(self):
        print('adding')
        abs_filename = filedialog.askopenfilename(title='选择文件', filetypes=[('All Files', '*')])
        try:
            temp = abs_filename.split("/")
            filename = temp[len(temp)-1:]
        except:
            pass
        if type(abs_filename) != type(()) and abs_filename != '':
            self.label_attachments.insert(END, filename)
            attach_list.append(abs_filename)

    def test(self):
        print(attach_list)

    def send(self):
        """处理文本，构建信息并发送"""
        if self.to.get() == "" or self.subject.get() == "" or self.message.get("1.0",END) == "":
            messagebox.showinfo("内容不能为空", "收件人，标题以及邮件内容")
        else:
            sender = '【发件人邮箱】'
            receiver = self.to.get()

            # 主题
            """**主题如果是纯中文或纯英文则字符数必须大于等于5个，
            不然会报错554 SPM被认为是垃圾邮件或者病毒** """
            subject = self.subject.get()
            # 文章
            body = self.message.get("0.0", END)
            # 服务器地址
            smtpserver = 'smtp.163.com'
            # 用户名（不是邮箱）
            username = '【你的用户名】'
            # 163授权码
            password = '【你的163授权码】'
            message_all = MIMEMultipart()
            message_all['Subject'] = Header(subject, 'utf-8')
            message_all['From'] = sender
            message_all['To'] = receiver

            self.msg = MIMEText(body, 'plain', 'utf-8')
            message_all.attach(self.msg)
            for attachment in attach_list:
                attach_file = MIMEText(open(attachment, 'rb').read(), 'plain', 'utf-8')
                attach_file["Content-Type"] = 'application/octet-stream'
                temp = attachment.split("/")
                filename = temp[len(temp)-1:]
                realname = filename[0]
                realname = realname.encode('unicode_escape')
                attach_file["Content-Disposition"] = "attachment; filename={0}".format(realname)
                message_all.attach(attach_file)

            # 服务器地址和端口25
            smtp = smtplib.SMTP(smtpserver, 25)
            smtp.set_debuglevel(1)
            smtp.login(username, password)
            smtp.sendmail(sender, receiver, message_all.as_string())
            self.message.delete('1.0', END)
            self.message.insert(END, '邮件已发送！')
            smtp.quit()

if __name__ == "__main__":
    attach_list = []
    root = Tk()
    root.title('快速邮件')
    root.geometry('500x550')
    app = Application(root)
    app.mainloop()
