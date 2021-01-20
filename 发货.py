import win32api
from tkinter import *
from tkinter import messagebox
import yagmail
from pyperclip import copy

yag = yagmail.SMTP(user="2728098382@qq.com",host="smtp.qq.com")
contents = ["D:\\实用小工具\\Fish-v3212\\Fish-v3212\\kpdf\\"]
# 这里填写的是你要发送的人的扣扣邮箱


def open_fish():
    try:
        copy(entry1.get())
        win32api.ShellExecute(0, 'open', "D:\实用小工具\Fish-v3212\Fish-v3212\Fish.exe", '', entry1.get(), 1)
    except Exception as e:
        messagebox.showerror("打开Fish软件错误",e)

def open_conv():
    try:
        win32api.ShellExecute(0, 'open', "D:\Program Files (x86)\SolidDocuments\Solid Converter v10\SolidConverterPDFv10.exe", entry2.get()+'.pdf', 'D:\实用小工具\Fish-v3212\Fish-v3212\kpdf', 1)
    except Exception as e:
       messagebox.showerror("打开Converter软件错误",e) 

def email():
    try:
        addr = entry3.get()
        file_name=entry2.get()+"."+file_ext
        yag.send(addr,"百度文库文档",[contents[0]+file_name])
        messagebox.showinfo('恭喜', "发送成功！")
    except Exception as e:
        messagebox.showerror('发送邮件错误', e)

def file_selection():
    global file_ext
    file_ext=option.get()


def gui(root):
    global option
    fishbtn = Button(root, text='打开Fish',command=open_fish,width=10)
    fishbtn.place(x=0, y=80)

    convbtn = Button(root, text='打开conv',command=open_conv,width=10)
    convbtn.place(x=100, y=80)

    emailbtn = Button(root, text='发送邮件',command=email,width=10)
    emailbtn.place(x=200, y=80)

    r1 = Radiobutton(root, text='docx', variable=option, value='docx', command=file_selection)
    r1.place(x=225, y=0)
    r2 = Radiobutton(root, text='pptx', variable=option, value='pptx', command=file_selection)
    r2.place(x=225, y=20)
    r3 = Radiobutton(root, text='xlsx', variable=option, value='xlsx', command=file_selection)
    r3.place(x=225, y=40)

    r4 = Radiobutton(root, text='pdf', variable=option, value='pdf', command=file_selection)
    r4.place(x=225, y=59)

    Label(root,text="链接地址:").grid(row=0,column=1)
    Label(root,text="文件名:").grid(row=1,column=1)
    Label(root,text="邮箱地址:").grid(row=2,column=1)
        

if __name__ == '__main__':  
    root = Tk()
    sw = root.winfo_screenwidth() # 获取电脑屏幕宽度
    sh = root.winfo_screenheight() # 获取电脑屏幕宽度
    ww = 300 #设置窗口宽度
    wh = 110 #设置窗口高度

# 窗口宽高为100

    x = int((sw-ww) / 2)
    y = int((sh-wh) / 2)
    root.geometry("{}x{}+{}+{}".format(ww, wh, x, y))
    file_ext="docx"
    option = StringVar()
    option.set('docx') 

    gui(root)
    root.title("百度文档发货")
    entry1 = Entry(root)
    entry1.grid(row=0,column=2)
    entry2 = Entry(root)
    entry2.grid(row=1,column=2)
    entry3 = Entry(root)
    entry3.grid(row=2,column=2)
    root.mainloop()