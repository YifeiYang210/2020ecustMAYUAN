# -*- coding:utf8 -*-
from tkinter import *
from pandas import read_excel
import random

root = Tk()
root.title("2020马原选择题刷题系统v1.2")
root.geometry("450x375")

# 读取文件数据
sq = read_excel('single_question.xlsx',sheet_name='Sheet1',header=0)
sa = read_excel('single_answer.xlsx',sheet_name='Sheet1',header=0)
sp = read_excel('single_pro.xlsx',sheet_name='Sheet1',header=0)
mq = read_excel('multiple_question.xlsx',sheet_name='Sheet1',header=0)
ma = read_excel('multiple_answer.xlsx',sheet_name='Sheet1',header=0)

# 欢迎信息
welcome_01 = Label(root,text="\n\n\n\n\n\n欢迎使用本刷题系统",fg='red',font=("黑体",15))
welcome_01.pack(side=TOP)
welcome_02 = Label(root,text="\n\n点击左上角的菜单栏开始刷题",font=("黑体",10))
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
    info_03 = Label(root,text="作者：国贸171周子寒，校对：梁天一，指导老师：汪帮琼")
    info_03.pack()
    info_04 = Label(root,text="特别鸣谢：曹旭，杨逸飞",fg='red',font=("黑体",10))
    info_04.pack()
    info_05 = Label(root,text="测试：陈烨",fg='red',font=("黑体",10))
    info_05.pack()
    
# 设置全局变量
i = 0
j = 0
m = 0
n = 0
single_test_right = 0
multiple_test_right = 0

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

    # 下一题 
    def single_next():
        global i
        i += 1
        if i <= 81:
            clean()
            single()  # 循环嵌套
        else:
            info_single_next = Label(root,text="已经是最后一题了哦！",fg='red')
            info_single_next.place(x=300,y=282.5)
            B4['state'] = DISABLED  # 禁用按钮(下同) 此处感谢陈烨同学提出的建议！
            
    # 上一题 循环嵌套
    def single_last():
        global i
        i -= 1
        if i >= 0:
            clean()
            single()  # 循环嵌套
        else:
            info_single_last = Label(root,text="已经是第一题了哦！",fg='red')
            info_single_last.place(x=300,y=282.5)
            B5['state'] = DISABLED

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

    # 下一题 
    def multiple_next():
        global j
        j += 1
        if j <= 86:
            clean()
            multiple()  # 循环嵌套
        else:
            info_multiple_next = Label(root,text="已经是最后一题了哦！",fg='red')
            info_multiple_next.place(x=300,y=302.5)
            B3['state'] = DISABLED
            
    # 上一题 
    def multiple_last():
        global j
        j -= 1
        if j >= 0:
            clean()
            multiple()  # 循环嵌套
        else:
            info_multiple_last = Label(root,text="已经是第一题了哦！",fg='red')
            info_multiple_last.place(x=300,y=302.5)
            B4['state'] = DISABLED

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
    B2 = Button(root, text="答案",width=8,command=multiple_answer)
    B2.place(x=160,y=260)
    B3 = Button(root, text="下一题",width=8,command=multiple_next)
    B3.place(x=280,y=260)
    B4 = Button(root, text="上一题",width=8,command=multiple_last)
    B4.place(x=360,y=260)

    mainloop()

# 单选题自测
def single_self_test():
    global m
    m = 0
    clean()
    single_list = []
    single_wrong = []
    for a in range(1,83):
        single_list.append(a)

    # 单选题自测刷题
    def single_test():
        single_choice = random.sample(single_list,int(E1.get())) # 生成不重复的随机数

        # 显示成绩结果
        def submit():
            result_01 = Label(root,text='你一共做了%d题，做对%d题，正确率为：%.2f%%' %(int(E1.get()),single_test_right,single_test_right*100/int(E1.get())),fg='red')
            result_01.place(x=40,y=300)
            single_quit = Button(root, text="退出自测",width=8,command=clean)
            single_quit.place(x=240,y=250)
            global m
            m = 0
            if len(single_wrong) != 0:
                result_02 = Label(root,text='做错的题号为：'+str(single_wrong),fg='red')
                result_02.place(x=40,y=320)
            
        # 下一题     
        def single_test_next():
            global m
            global single_test_right

            if RadioVar.get() == sa.iloc[single_choice[m],5]:
                    single_test_right += 1  # 答案正确
            else:
                single_wrong.append(single_choice[m]+1)  # 答案错误
                
            m += 1          
            if m <= int(E1.get())-1:
                clean()
                single_test()  # 循环嵌套
                  
            elif m == int(E1.get()):
                B2 = Button(root, text="提交",width=8,command=submit)
                B2.place(x=140,y=250)       
                B1['state'] = DISABLED
        
        # 自测单选题题目
        L1 = Label(root,text=sq.iloc[single_choice[m],1])
        L1.place(x=15,y=10)
        L2 = Label(root,text='   '+sq.iloc[single_choice[m],2])
        L2.place(x=15,y=30)
        L3 = Label(root,text='   '+sq.iloc[single_choice[m],3])
        L3.place(x=15,y=50)
        L4 = Label(root,text='   '+sq.iloc[single_choice[m],4])
        L4.place(x=15,y=70)
        L5 = Label(root,text='   '+sq.iloc[single_choice[m],5])
        L5.place(x=15,y=90)
        L6 = Label(root,text='   '+sq.iloc[single_choice[m],6])
        L6.place(x=15,y=110)
        L7 = Label(root,text='   '+sq.iloc[single_choice[m],7])
        L7.place(x=15,y=130)

        # 自测题单选框
        RadioVar = IntVar() 
        R1 = Radiobutton(root,text=sa.iloc[single_choice[m],0],variable=RadioVar,value=1)
        R1.place(x=25,y=150)
        R2 = Radiobutton(root,text=sa.iloc[single_choice[m],1],variable=RadioVar,value=2)
        R2.place(x=25,y=170)
        R3 = Radiobutton(root,text=sa.iloc[single_choice[m],2],variable=RadioVar,value=3)
        R3.place(x=25,y=190)
        R4 = Radiobutton(root,text=sa.iloc[single_choice[m],3],variable=RadioVar,value=4)
        R4.place(x=25,y=210)

        # 自测单选题按钮
        B1 = Button(root, text="下一题",width=8,command=single_test_next)
        B1.place(x=40,y=250)

    # 验证单选题自测开始
    def single_begin():
        if len(E1.get().strip())!=0 and E1.get().isdigit()==1 and float(E1.get())%1==0 and 1<=int(E1.get())<=82:        
            clean()
            single_test()            
        else:
            L4 = Label(root,text='请正确输入[1,82]范围内的整数',fg='red')
            L4.place(x=135,y=260)            

    # 单选题自测UI界面
    L1 = Label(root,text='单选题自测(总共82题)',fg='red',font=("黑体",12))
    L1.place(x=140,y=100)
    E1 = Entry(root,width=5,text='single_num')
    E1.place(x=280,y=142.5)
    L2 = Label(root,text='请输入你要做的题目数量：')
    L2.place(x=120,y=140)
    L3 = Label(root,text='（系统将随机生成相应数量的单选题）')
    L3.place(x=115,y=180)
    B1 = Button(root,text="开始刷题",width=8,command=single_begin)
    B1.place(x=185,y=220)

    mainloop()

# 多选题自测
def multiple_self_test():
    global n
    n = 0
    clean()
    multiple_list = []
    multiple_wrong = []
    for b in range(1,88):
        multiple_list.append(b)

    # 多选题自测刷题
    def multiple_test():
        multiple_choice = random.sample(multiple_list,int(E1.get())) # 生成不重复的随机数

        # 显示成绩结果
        def submit():
            result_01 = Label(root,text='你一共做了%d题，做对%d题，正确率为：%.2f%%' %(int(E1.get()),multiple_test_right,multiple_test_right*100/int(E1.get())),fg='red')
            result_01.place(x=40,y=310)
            multiple_quit = Button(root, text="退出自测",width=8,command=clean)
            multiple_quit.place(x=240,y=270)
            global n
            n = 0
            if len(multiple_wrong) != 0:
                result_02 = Label(root,text='做错的题号为：'+str(multiple_wrong),fg='red')
                result_02.place(x=40,y=330)
            
        # 下一题     
        def multiple_test_next():
            global n
            global multiple_test_right

            if CheckVar1.get()==ma.iloc[multiple_choice[n],5] and CheckVar2.get()==ma.iloc[multiple_choice[n],6] and CheckVar3.get()==ma.iloc[multiple_choice[n],7] and CheckVar4.get()==ma.iloc[multiple_choice[n],8]:
                    multiple_test_right += 1  # 答案正确
            else:
                multiple_wrong.append(multiple_choice[n]+1)  # 答案错误
                
            n += 1          
            if n <= int(E1.get())-1:
                clean()
                multiple_test()  # 循环嵌套
                  
            elif n == int(E1.get()):
                B2 = Button(root, text="提交",width=8,command=submit)
                B2.place(x=140,y=270)       
                B1['state'] = DISABLED
        
        # 自测多选题题目
        L1 = Label(root,text=sq.iloc[multiple_choice[n],1])
        L1.place(x=15,y=10)
        L2 = Label(root,text='   '+sq.iloc[multiple_choice[n],2])
        L2.place(x=15,y=30)
        L3 = Label(root,text='   '+sq.iloc[multiple_choice[n],3])
        L3.place(x=15,y=50)
        L4 = Label(root,text='   '+sq.iloc[multiple_choice[n],4])
        L4.place(x=15,y=70)
        L5 = Label(root,text='   '+sq.iloc[multiple_choice[n],5])
        L5.place(x=15,y=90)
        L6 = Label(root,text='   '+sq.iloc[multiple_choice[n],6])
        L6.place(x=15,y=110)
        L7 = Label(root,text='   '+sq.iloc[multiple_choice[n],7])
        L7.place(x=15,y=130)
        L8 = Label(root,text='   '+mq.iloc[multiple_choice[n],8])
        L8.place(x=15,y=150)

        # 自测题复选框
        CheckVar1 = IntVar()
        CheckVar2 = IntVar()
        CheckVar3 = IntVar()
        CheckVar4 = IntVar() 
        C1 = Checkbutton(root,text=ma.iloc[multiple_choice[n],0],variable=CheckVar1,onvalue=1,offvalue=0)
        C1.place(x=25,y=170)
        C2 = Checkbutton(root,text=ma.iloc[multiple_choice[n],1],variable=CheckVar2,onvalue=1,offvalue=0)
        C2.place(x=25,y=190)
        C3 = Checkbutton(root,text=ma.iloc[multiple_choice[n],2],variable=CheckVar3,onvalue=1,offvalue=0)
        C3.place(x=25,y=210)
        C4 = Checkbutton(root,text=ma.iloc[multiple_choice[n],3],variable=CheckVar4,onvalue=1,offvalue=0)
        C4.place(x=25,y=230)

        # 自测多选题按钮
        B1 = Button(root, text="下一题",width=8,command=multiple_test_next)
        B1.place(x=40,y=270)

    # 验证多选题自测开始
    def multiple_begin():
        if len(E1.get().strip())!=0 and E1.get().isdigit()==1 and float(E1.get())%1==0 and 1<=int(E1.get())<=87:         
            clean()
            multiple_test()
        else:
            L4 = Label(root,text='请正确输入[1,87]范围内的整数',fg='red')
            L4.place(x=135,y=260)            

    # 多选题自测UI界面
    L1 = Label(root,text='多选题自测(总共87题)',fg='red',font=("黑体",12))
    L1.place(x=140,y=100)
    E1 = Entry(root,width=5,text='multiple_num')
    E1.place(x=280,y=142.5)
    L2 = Label(root,text='请输入你要做的题目数量：')
    L2.place(x=120,y=140)
    L3 = Label(root,text='（系统将随机生成相应数量的多选题）')
    L3.place(x=115,y=180)
    B1 = Button(root,text="开始刷题",width=8,command=multiple_begin)
    B1.place(x=185,y=220)

    mainloop()

# 菜单栏
def menu():
    menubar=Menu(root)
    
    mathmenu=Menu(menubar,tearoff=0)
    mathmenu.add_command(label="单选题 (共82题)",command=single)
    mathmenu.add_command(label="多选题 (共87题)",command=multiple)
    menubar.add_cascade(label="题库",menu=mathmenu)

    testmenu=Menu(menubar,tearoff=0)
    testmenu.add_command(label="单选题",command=single_self_test)
    testmenu.add_command(label="多选题",command=multiple_self_test)
    menubar.add_cascade(label="自测",menu=testmenu)

    helpmenu=Menu(menubar,tearoff=0)
    helpmenu.add_command(label="关于",command=about)
    helpmenu.add_separator()
    helpmenu.add_command(label="退出",command=root.destroy)
    menubar.add_cascade(label="帮助",menu=helpmenu)

    root.config(menu=menubar)
    root.mainloop()

if __name__ == '__main__':
    menu()
