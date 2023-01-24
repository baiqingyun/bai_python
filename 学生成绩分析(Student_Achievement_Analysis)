from tkinter import *
from tkinter import filedialog
import tkinter.messagebox
import pandas as pd


window = Tk()
window.title("学生成绩分析程序 By BaiQing")
window.geometry('643x388')
window.resizable(0,0)   #窗口不允许用户调整大小

path = "示例.xlsx"
manfen = 100
manfen_float = float(manfen)
work = "pass"
use_work = "new"


def choose_file():
    global path
    path = filedialog.askopenfilename(filetypes=[('Excel', '.xlsx',)], title="选择xlsx格式文件")
    return path


def manfenpanduan():
    global manfen
    global manfen_float
    manfen = input_manfen.get()
    if manfen.isdigit():
        print_info.config(state=NORMAL)
        print_info.delete("1.0", "end")
        print_info.insert(1.0, "                    --------------------设置成功--------------------            ")
        print_info.config(state=DISABLED)
        return
    else:
        tkinter.messagebox.showerror('提示', '输入的满分数值错误')


def choose_work1():
    global use_work
    work = "replace"
    use_work = work
    input_work.config(state=NORMAL)
    input_work.delete("1.0", "end")
    input_work.insert(1.0, "当前工作方式：覆盖")
    input_work.config(state=DISABLED)


def choose_work2():
    global work
    global use_work
    work = "new"
    use_work = work
    input_work.config(state=NORMAL)
    input_work.delete('1.0', 'end')
    input_work.insert(1.0, "当前工作方式：新建")
    input_work.config(state=DISABLED)


def pdfile():
    global ask
    if path == "示例.xlsx":
        ask = tkinter.messagebox.askyesno("提示", "当前未选择Excel文件，是否使用示例文件运行")
    else:
        ask = "pass"


def duqu():
    global ask
    global ask_str
    ask_str = str(ask)
    if ask_str == "True" or ask_str == "pass":
        global path
        global work
        global manfen_float
        global use_work

        data = pd.read_excel(path, index_col=0)
        jige = manfen_float * 0.6
        lianghao = manfen_float * 0.75
        youxiu = manfen_float * 0.85

        cuowu1 = data.loc[(data["成绩"] > manfen_float) | (data["成绩"] < 0)]
        if cuowu1.empty:
            pass
        else:
            writer = pd.ExcelWriter(path, mode='a', engine='openpyxl', if_sheet_exists=use_work)
            cuowu1.to_excel(writer, sheet_name='成绩错误')
            # writer.save()
            writer.close()

        bujige_student = data.loc[(data["成绩"] < jige) & (data["成绩"] >= 0)]
        if bujige_student.empty:
            pass
        else:
            writer = pd.ExcelWriter(path, mode='a', engine='openpyxl', if_sheet_exists=use_work)
            bujige_student.to_excel(writer, sheet_name='不及格')
            # writer.save()
            writer.close()

        jige_student = data.loc[(data["成绩"] < lianghao) & (data["成绩"] >= jige)]
        if jige_student.empty:
            pass
        else:
            writer = pd.ExcelWriter(path, mode='a', engine='openpyxl', if_sheet_exists=use_work)
            jige_student.to_excel(writer, sheet_name='及格')
            # writer.save()
            writer.close()

        lianghao_student = data.loc[(data["成绩"] >= lianghao) & (data["成绩"] < youxiu)]
        if lianghao_student.empty:
            pass

        else:
            writer = pd.ExcelWriter(path, mode='a', engine='openpyxl', if_sheet_exists=use_work)
            lianghao_student.to_excel(writer, sheet_name='良好')
            # writer.save()
            writer.close()

        youxiu_student = data.loc[(data["成绩"] >= youxiu) & (data["成绩"] < manfen_float)]
        if youxiu_student.empty:
            pass


        else:
            writer = pd.ExcelWriter(path, mode='a', engine='openpyxl', if_sheet_exists=use_work)
            youxiu_student.to_excel(writer, sheet_name='优秀')
            # writer.save()
            writer.close()

        manfen_student = data.loc[data["成绩"] == manfen_float]
        if manfen_student.empty:
            print_info.config(state=NORMAL)
            print_info.delete("1.0", "end")
            print_info.insert(1.0,"----------------------------------详细信息已写入原文件----------------------------------")
            print_info.config(state=DISABLED)
        else:
            writer = pd.ExcelWriter(path, mode='a', engine='openpyxl', if_sheet_exists=use_work)
            manfen_student.to_excel(writer, sheet_name='满分')
            #    writer.save()
            writer.close()
            print_info.config(state=NORMAL)
            print_info.delete("1.0", "end")
            print_info.insert(1.0,"----------------------------------详细信息已写入原文件----------------------------------")
            print_info.config(state=DISABLED)
    else:
        return


def print_text():  # 显示选择的文件
    input_file.config(state=NORMAL)
    input_file.delete('1.0', 'end')
    input_file.insert(1.0, choose_file())
    input_file.config(state=DISABLED)






l0 = Label(window, text="一、请选择工作模式:", anchor="center")
l0.place(x=-1, y=20, width=142, height=30)

btn00 = Button(window, text="覆盖原工作表", anchor="center", command=choose_work1)
btn00.place(x=160, y=20, width=89, height=32)

btn01 = Button(window, text="新建工作表", anchor="center", command=choose_work2)
btn01.place(x=260, y=20, width=89, height=32)

l1 = Label(window, text="二、请输入满分数值:", anchor="center")
l1.place(x=0, y=90, width=138, height=34)

input_manfen = Entry(window)
input_manfen.place(x=150, y=90, width=357, height=30)

l2 = Label(window, text="三、请选择Excel文件：", anchor="center")
l2.place(x=0, y=160, width=146, height=30)

input_file = Text(window)
input_file.config(state=DISABLED)
input_file.place(x=150, y=160, width=361, height=30)

btn02 = Button(window, text="浏览...", anchor="center", command=print_text, fg="blue")
btn02.place(x=530, y=160, width=86, height=30)

print_info = Text(window, fg="red")
print_info.config(state=DISABLED)
print_info.place(x=10, y=290, width=621, height=48)

def print_tips():
    print_info.config(state=NORMAL)
    print_info.insert(0.0, "程序运行提示：如未选择（输入）工作模式、满分数值、Excel文件，本程序将使用默认设置运行", '\n',"                --------------------设置完成后不点确定设置无效--------------------")
    print_info.config(state=DISABLED)
    return


if path == "示例.xlsx":
    print_tips()


input_work = Text(window)
input_work.config(state=DISABLED)
input_work.place(x=360, y=20, width=136, height=29)

queding_manfen = Button(window, text="确定", command=manfenpanduan, fg="blue")
queding_manfen.place(x=530, y=90, width=83, height=32)

btn03 = Button(window, text="开始分析", command=lambda: [pdfile(), duqu()])
btn03.place(x=250, y=210, width=140, height=42)

window.mainloop()
