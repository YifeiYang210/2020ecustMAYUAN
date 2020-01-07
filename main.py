# -*- coding:utf8 -*-
from tkinter import *
from pandas import read_excel
import xlrd # 打包用
import pyexpat # 打包用

root = Tk()
root.title("2020马原选择题刷题系统")
root.geometry("800x600")

# 读取文件数据
sq = read_excel('single_question.xlsx',sheet_name='Sheet1',header=0)
sa = read_excel('single_answer.xlsx',sheet_name='Sheet1',header=0)
sp = read_excel('single_pro.xlsx',sheet_name='Sheet1',header=0)
mq = read_excel('multiple_question.xlsx',sheet_name='Sheet1',header=0)
ma = read_excel('multiple_answer.xlsx',sheet_name='Sheet1',header=0)

# 欢迎信息
welcome_01 = Label(root,text="\n\n\n\n\n\n欢迎使用本刷题系统",fg='red',font=("黑体",15))
welcome_01.pack(side=TOP)
welcome_02 = Label(root,text="\n\n点击“选择题”开始刷题",font=("黑体",10))
welcome_02.pack()

# 清屏函数
def clean():    
    Label(root,text='',width=200,height=200).place(x=0,y=0)
    # 此处非常感谢曹旭同学的帮助和指点！
    
# 开发者信息
def about():
    clean()
    info_01 = Label(root,text="开发者：机设173丁泓宇\n",fg='red',font=("黑体",10))
    info_01.pack(side=TOP)
    info_02 = Label(root,text="题目来源：2019-2020学年马原选择题题库")
    info_02.pack()
    info_02 = Label(root,text="作者：国贸171周子寒，校对：梁天一，指导老师：汪帮琼")
    info_02.pack()
    info_03 = Label(root,text="特别鸣谢：曹旭",fg='red',font=("黑体",10))
    info_03.pack()
    
# 设置全局变量
i = 0
j = 0 

# 单选题函数
def single():
    clean()
    
    # 验证答案正误
    def single_check():
        if RadioVar.get() == 0:
            answer_right = Label(root,text="请选择！",fg='red')
            answer_right.place(x=110,y=242.5)
        elif RadioVar.get() == sa.iloc[i,5]:
            answer_right = Label(root,text="正确！ ",fg='red')
            answer_right.place(x=110,y=242.5)
        else:
            answer_wrong = Label(root,text="错误！ ",fg='red')
            answer_wrong.place(x=110,y=242.5)

    # 单选题提示内容
    def single_pro():
        p1 = Label(root,text=sp.iloc[i,1],fg='blue')
        p1.place(x=110,y=282.5)
        p2 = Label(root,text=sp.iloc[i,2],fg='blue')
        p2.place(x=110,y=302.5)
        p3 = Label(root,text=sp.iloc[i,3],fg='blue')
        p3.place(x=110,y=322.5)
        p4 = Label(root,text=sp.iloc[i,4],fg='blue')
        p4.place(x=110,y=342.5)

    # 单选题答案
    def single_answer():
        answer = Label(root,text=sa.iloc[i,4],fg='red')
        answer.place(x=250,y=242.5)

    # 下一题 循环嵌套
    def single_next():
        global i
        i += 1
        if i <= 81:
            clean()
            single()     
        else:
            info_single_next = Label(root,text="已经是最后一题了哦！",fg='red')
            info_single_next.place(x=300,y=282.5)
            
    # 显示下一题 循环嵌套
    def single_last():
        global i
        i -= 1
        if i >= 0:
            clean()
            single()
        else:
            info_single_last = Label(root,text="已经是第一题了哦！",fg='red')
            info_single_last.place(x=300,y=282.5)

    # 单选题题目
    L1 = Label(root,text=sq.iloc[i,1])
    L1.place(x=15,y=10)
    L2 = Label(root,text='   '+sq.iloc[i,2])
    L2.place(x=15,y=30)
    L3 = Label(root,text='   '+sq.iloc[i,3])
    L3.place(x=15,y=50)
    L4 = Label(root,text='   '+sq.iloc[i,4])
    L4.place(x=15,y=70)
    L5 = Label(root,text='   '+sq.iloc[i,5])
    L5.place(x=15,y=90)
    L6 = Label(root,text='   '+sq.iloc[i,6])
    L6.place(x=15,y=110)
    L7 = Label(root,text='   '+sq.iloc[i,7])
    L7.place(x=15,y=130)

    # 单选框
    RadioVar = IntVar() 
    R1 = Radiobutton(root,text=sa.iloc[i,0],variable=RadioVar,value=1)
    R1.place(x=25,y=150)
    R2 = Radiobutton(root,text=sa.iloc[i,1],variable=RadioVar,value=2)
    R2.place(x=25,y=170)
    R3 = Radiobutton(root,text=sa.iloc[i,2],variable=RadioVar,value=3)
    R3.place(x=25,y=190)
    R4 = Radiobutton(root,text=sa.iloc[i,3],variable=RadioVar,value=4)
    R4.place(x=25,y=210)

    # 单选题按钮
    B1 = Button(root, text="确定",width=8,command=single_check)
    B1.place(x=30,y=240)
    B2 = Button(root, text="提示",width=8,command=single_pro)
    B2.place(x=30,y=280)
    B3 = Button(root, text="答案",width=8,command=single_answer)
    B3.place(x=170,y=240)
    B4 = Button(root, text="下一题",width=8,command=single_next)
    B4.place(x=280,y=240)
    B5 = Button(root, text="上一题",width=8,command=single_last)
    B5.place(x=360,y=240)

    mainloop()

# 多选题函数
def multiple():
    clean()
    
    # 验证答案正误
    def multiple_check():
        if CheckVar1.get()==CheckVar2.get()==CheckVar3.get()==CheckVar4.get()==0:
            answer_right = Label(root,text="请选择！",fg='red')
            answer_right.place(x=105,y=262.5)
        elif CheckVar1.get()==ma.iloc[j,5] and CheckVar2.get()==ma.iloc[j,6] and CheckVar3.get()==ma.iloc[j,7] and CheckVar4.get()==ma.iloc[j,8]:
            answer_right = Label(root,text="正确！ ",fg='red')
            answer_right.place(x=105,y=262.5)
        else:
            answer_wrong = Label(root,text="错误！ ",fg='red')
            answer_wrong.place(x=105,y=262.5)

    # 多选题答案
    def multiple_answer():
        answer = Label(root,text=ma.iloc[j,4],fg='red')
        answer.place(x=235,y=262.5)

    # 下一题 循环嵌套
    def multiple_next():
        global j
        j += 1
        if j <= 86:
            clean()
            multiple()     
        else:
            info_multiple_next = Label(root,text="已经是最后一题了哦！",fg='red')
            info_multiple_next.place(x=300,y=302.5)
            
    # 下一题 循环嵌套
    def multiple_last():
        global j
        j -= 1
        if j >= 0:
            clean()
            multiple()
        else:
            info_multiple_last = Label(root,text="已经是第一题了哦！",fg='red')
            info_multiple_last.place(x=300,y=302.5)

    # 多选题题目
    L1 = Label(root,text=mq.iloc[j,1])
    L1.place(x=15,y=10)
    L2 = Label(root,text='   '+mq.iloc[j,2])
    L2.place(x=15,y=30)
    L3 = Label(root,text='   '+mq.iloc[j,3])
    L3.place(x=15,y=50)
    L4 = Label(root,text='   '+mq.iloc[j,4])
    L4.place(x=15,y=70)
    L5 = Label(root,text='   '+mq.iloc[j,5])
    L5.place(x=15,y=90)
    L6 = Label(root,text='   '+mq.iloc[j,6])
    L6.place(x=15,y=110)
    L7 = Label(root,text='   '+mq.iloc[j,7])
    L7.place(x=15,y=130)
    L8 = Label(root,text='   '+mq.iloc[j,8])
    L8.place(x=15,y=150)

    # 复选框
    CheckVar1 = IntVar()
    CheckVar2 = IntVar()
    CheckVar3 = IntVar()
    CheckVar4 = IntVar()
    C1 = Checkbutton(root,text=ma.iloc[j,0],variable=CheckVar1,onvalue=1,offvalue=0)
    C1.place(x=25,y=170)
    C2 = Checkbutton(root,text=ma.iloc[j,1],variable=CheckVar2,onvalue=1,offvalue=0)
    C2.place(x=25,y=190)
    C3 = Checkbutton(root,text=ma.iloc[j,2],variable=CheckVar3,onvalue=1,offvalue=0)
    C3.place(x=25,y=210)
    C4 = Checkbutton(root,text=ma.iloc[j,3],variable=CheckVar4,onvalue=1,offvalue=0)
    C4.place(x=25,y=230)

    # 多选题按钮
    B1 = Button(root, text="确定",width=8,command=multiple_check)
    B1.place(x=30,y=260)
    B3 = Button(root, text="答案",width=8,command=multiple_answer)
    B3.place(x=160,y=260)
    B4 = Button(root, text="下一题",width=8,command=multiple_next)
    B4.place(x=280,y=260)
    B5 = Button(root, text="上一题",width=8,command=multiple_last)
    B5.place(x=360,y=260)

    mainloop()

# 菜单栏
def menu():
    menubar=Menu(root)
    mathmenu=Menu(menubar,tearoff=0)
    mathmenu.add_command(label="单选题 (共82题)",command=single)
    mathmenu.add_command(label="多选题 (共87题)",command=multiple)
    menubar.add_cascade(label="选择题",menu=mathmenu)

    helpmenu=Menu(menubar,tearoff=0)
    helpmenu.add_command(label="关于",command=about)
    helpmenu.add_separator()
    helpmenu.add_command(label="关闭",command=root.destroy)
    menubar.add_cascade(label="帮助",menu=helpmenu)

    root.config(menu=menubar)
    root.mainloop()

if __name__ == '__main__':
    menu()
