# -*- coding:utf-8 -*-
import sys
import xlrd
import tkinter as tk
import tkinter.messagebox

# 生成物编
def function_read(sheetDaoju):
    __ItemInfo = ""
    lists = [[] for i in range(sheetDaoju.ncols)]
    lists2 = []

    for i in range(2, sheetDaoju.nrows):
        for j in range(sheetDaoju.ncols):
            lists[j].append(sheetDaoju.cell_value(i, j))

    for i in range(1, sheetDaoju.nrows-3):
        for j in range(sheetDaoju.ncols):
            lists2.append(sheetDaoju.cell_value(i, j))


    return lists2, lists
def function_lua(sheetDaoju, sheetinfo, sheetDefine, isshow):
    temp_array = ["装备_01主手", "装备_02副手", "装备_03肩部", "装备_04上身", "装备_05腰部", "装备_06腿部", "装备_07脚部", "装备_08颈部", "装备_09手腕", "装备_10手指", "装备_11称号"]

    temp_number1,temp_number2,temp_number3,temp_number4,temp_number5,temp_number6,temp_number7,temp_number8,temp_number9,temp_number10,temp_number11 = 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0
    temp_number_array = [temp_number1, temp_number2, temp_number3, temp_number4, temp_number5, temp_number6, temp_number7, temp_number8, temp_number9, temp_number10, temp_number11]

    _sheetName = sheetName.get(index1=1.0, index2=tk.END)[:-1]
    __Lua_Title, __Lua_Item = function_read(sheetDaoju)
    LuaCode = luaCodeStart + luaCodeFunction
    LuaCode1 = luaCodeStart + luaCodeFunction1
    iteminfo = data_handle(sheetinfo, isshow)
    iteminfo_array = iteminfo.split(',')
    if '道具' in _sheetName:#如果表格名含有道具那么生成的jass函数名为 test1
        tempItem = LuaCode
    else:#如果表格名不含有道具那么生成的jass函数名为 test2
        tempItem = LuaCode1
    for i in range(len(__Lua_Item[0])):
        for j in range(len(__Lua_Item)):
            if '道具' in _sheetName:
                if j == 0:
                    tempItem = tempItem + luaCodeID.format(__Lua_Item[0][i])
                else:
                    if __Lua_Item[j][i] == "":
                        pass
                    else:
                        if __Lua_Title[j] == 'Description':
                            tempItem = tempItem + luaCodeItem2.format(__Lua_Title[j], '"', iteminfo_array[i], '"')
                        elif __Lua_Title[j] == 'Tip':
                            tempItem = tempItem + luaCodeItem2.format(__Lua_Title[j], '"', iteminfo_array[i], '"')
                        elif __Lua_Title[j] == 'Ubertip':
                            tempItem = tempItem + luaCodeItem2.format(__Lua_Title[j], '"', iteminfo_array[i], '"')
                        else:
                            tempItem = tempItem + luaCodeItem.format(__Lua_Title[j], __Lua_Item[j][i])


            else:
                if j == 0:
                    tempItem = tempItem + luaCodeID2.format(_sheetName[-4:], __Lua_Item[0][i])
                else:
                    if __Lua_Item[j][i] == "":
                        pass
                    else:
                        tempItem = tempItem + luaCodeItem.format(__Lua_Title[j], __Lua_Item[j][i])

        for number in range(len(temp_array)):#处理道具注册
            if(sheetDefine.cell_value(i+2, 2) == temp_array[number][-2:]):
                temp_number_array[number] = temp_number_array[number] + 1
                jass_define = jassCodeDefine.format(temp_array[number], temp_number_array[number], sheetDefine.cell_value(i+2, 0))
            else:
                pass
        tempItem = tempItem + "\t?>\n" + jass_define#每生成一个道具，在下面加上一个定义
    tempItem = tempItem + luaCodeEnd
    print(tempItem)
    f = open(path + "\{0}.txt".format(_sheetName), 'w')
    f.write(tempItem)
    f.close()

# 生成预览
def data_handle(sheetinfo, isshow):
    title_array = []
    for i in range(1, len(sheetinfo.row_values(0))):
        title_array.append(sheetinfo.row_values(0)[i])

    temp_1, temp_2, temp_3, temp_4, temp_5, temp_6 = "", "", "", "", "", ""
    temp_array = ["", "", "", "", "", ""]
    temp = ""

    if isshow:
        for i in range(1, len(sheetinfo.col_values(0))):
            for j in range(1, len(title_array)+1):
                if sheetinfo.cell_value(i, j) == '':
                    pass
                else:
                    num = j-1
                    if j == 1:
                        temp_1 = title_array[num] + ":" + sheetinfo.cell_value(i, j) + '\n'
                    elif j == 2:
                        temp_2 = title_array[num] + ":" + sheetinfo.cell_value(i, j) + '\n'
                    elif j == 3:
                        temp_3 = title_array[num] + ":" + sheetinfo.cell_value(i, j) + '\n'
                    elif j == 4:
                        temp_4 = title_array[num] + ":" + sheetinfo.cell_value(i, j) + '\n'
                    elif j == 5:
                        tx_str = sheetinfo.cell_value(i, j)
                        tx_array1 = tx_str.split('|')
                        tx = ''
                        if len(tx_array1)>1:
                            for c in range(len(tx_array1)):
                                for a in range(1, len(tx_sheet.col_values(0))):
                                    for b in range(1, len(tx_sheet.row_values(0))):
                                        if tx_array1[c][1:len(sheetinfo.cell_value(i, j))] == tx_sheet.cell_value(a, b):
                                            tx = tx + tx_array1[c] + tx_sheet.cell_value(a, b + 1) + '\n'
                                        else:
                                            pass
                            temp_5 = title_array[num] + ":\n" + tx
                            pass
                        for a in range(1, len(tx_sheet.col_values(0))):
                            for b in range(1, len(tx_sheet.row_values(0))):
                                if sheetinfo.cell_value(i, j)[1:len(sheetinfo.cell_value(i, j))] == tx_sheet.cell_value(a, b):
                                    temp_5 = title_array[num] + ":\n" + sheetinfo.cell_value(i, j) + ":" + tx_sheet.cell_value(a, b+1) + '\n'
                                else:
                                    pass
                    elif j == 6:
                        temp_6 = title_array[num] + ":" + sheetinfo.cell_value(i, j) + '\n'
                    else:
                        pass
            temp = temp + temp_1 + temp_2 + temp_3 + temp_4 + temp_5 + temp_6 + '\n'
        f = open(path + "\{0}.txt".format('效果预览'), 'w')
        f.write(temp)
        f.close()
        return
    else:
        for i in range(1, len(sheetinfo.col_values(0))):
            for j in range(1, len(title_array)+1):
                if sheetinfo.cell_value(i, j) == '':
                    pass
                else:
                    num = j-1
                    if j == 1:
                        pass
                    elif j == 2:
                        temp_2 = title_array[num] + ":" + sheetinfo.cell_value(i, j) + '|n'
                    elif j == 3:
                        temp_3 = title_array[num] + ":" + sheetinfo.cell_value(i, j) + '|n'
                    elif j == 4:
                        temp_4 = title_array[num] + ":" + sheetinfo.cell_value(i, j) + '|n'
                    elif j == 5:
                        tx_str = sheetinfo.cell_value(i, j)
                        tx_array1 = tx_str.split('|')
                        tx = ''
                        if len(tx_array1) > 1:
                            for c in range(len(tx_array1)):
                                for a in range(1, len(tx_sheet.col_values(0))):
                                    for b in range(1, len(tx_sheet.row_values(0))):
                                        if tx_array1[c][1:len(sheetinfo.cell_value(i, j))] == tx_sheet.cell_value(a, b):
                                            tx = tx + tx_array1[c] + ":" + tx_sheet.cell_value(a, b + 1) + '|n'
                                        else:
                                            pass
                            temp_5 = title_array[num] + ":|n" + tx
                            pass
                        for a in range(1, len(tx_sheet.col_values(0))):
                            for b in range(1, len(tx_sheet.row_values(0))):
                                if sheetinfo.cell_value(i, j)[1:len(sheetinfo.cell_value(i, j))] == tx_sheet.cell_value(a, b):
                                    temp_5 = title_array[num] + ":|n" + sheetinfo.cell_value(i, j) + ":" + tx_sheet.cell_value(a, b + 1) + '|n'
                                else:
                                    pass
                    elif j == 6:
                        temp_6 = title_array[num] + ":" + sheetinfo.cell_value(i, j)
                    else:
                        pass
            temp = temp + temp_1 + temp_2 + temp_3 + temp_4 + temp_5 + temp_6 + ','
        return temp
def data_write():
    pass


def t1():

    Btn1.place(x=30, y=500)
    Btn2.place(x=30, y=500)
    Btn3.place(x=30, y=500)
    Btn4.place(x=30, y=500)

    Btn5.place(x=30, y=60)
    Btn6.place(x=260, y=60)
    Btn7.place(x=340, y=20)

    PathTip.place(x=5, y=5)
    FilePath.place(x=80, y=5)
    SkillTip.place(x=5, y=30)
    sheetName.place(x=80, y=35)

def t2():
    Btn1.place(x=30, y=60)
    Btn2.place(x=260, y=60)
    Btn3.place(x=260, y=300)
    Btn4.place(x=340, y=300)

    Btn5.place(x=30, y=400)
    Btn6.place(x=260, y=400)
    Btn7.place(x=340, y=400)

    PathTip.place(x=5, y=400)
    FilePath.place(x=80, y=400)
    SkillTip.place(x=5, y=300)
    sheetName.place(x=80, y=140)
    sheetName.delete(1.0, tk.END)


def T1():
    # 根据表格名称获取表格
    # 打开文件
    # path + '/' +file 是文件的完整路径
    path = FilePath.get(index1=1.0, index2=tk.END)[:-1]
    name = sheetName.get(index1=1.0, index2=tk.END)[:-1]
    isshow = False
    data = xlrd.open_workbook(path)
    _sheetDaoju = data.sheet_by_name(name)

    _sheetinfo = data.sheet_by_name('道具信息')
    _sheetDefine = data.sheet_by_name('道具注册')
    function_lua(_sheetDaoju, _sheetinfo, _sheetDefine, isshow)

def T2():
    PathTip.place(x=5, y=5)
    FilePath.place(x=80, y=5)
    SkillTip.place(x=5, y=30)
    sheetName.place(x=80, y=35)
    Btn1.place(x=30, y=100)
    Btn2.place(x=260, y=100)
    Btn3.place(x=150, y=65)
    Btn4.place(x=340, y=20)

def T3():
    # 根据表格名称获取表格
    # 打开文件
    # path + '/' +file 是文件的完整路径
    path = FilePath.get(index1=1.0, index2=tk.END)[:-1]
    data = xlrd.open_workbook(path)
    _sheetName = sheetName.get(index1=1.0, index2=tk.END)[:-1]
    _sheetDaoju = data.sheet_by_name(_sheetName)
    function_lua(_sheetDaoju)

def T4():
    Btn1.place(x=30, y=60)
    Btn2.place(x=260, y=60)
    Btn3.place(x=150, y=300)
    Btn4.place(x=150, y=300)

    PathTip.place(x=5, y=400)
    FilePath.place(x=80, y=400)
    SkillTip.place(x=5, y=300)
    sheetName.place(x=80, y=140)
    sheetName.delete(1.0, tk.END)

def T5():
    # 根据表格名称获取表格
    # 打开文件
    # path + '/' +file 是文件的完整路径
    path = FilePath.get(index1=1.0, index2=tk.END)[:-1]
    name = sheetName.get(index1=1.0, index2=tk.END)[:-1]
    data = xlrd.open_workbook(path)
    _sheetinfo = data.sheet_by_name(name)
    isshow = True
    data_handle(_sheetinfo, isshow)

if __name__ == '__main__':

    # 字符串定义
    luaCodeStart = "<?local slk = require 'slk' ?>\n"
    luaCodeFunction = "function Test1 takes nothing returns nothing\n"
    luaCodeFunction1 = "function Test2 takes nothing returns nothing\n"
    luaCodeID    = "\t<? local obj = slk.item.afac:new '{0}'?>\n\t<?\n"
    luaCodeID2   = "\t<? local obj = slk.ability.{0}:new '{1}'?>\n\t<?\n"
    luaCodeItem  = "\t\tobj.{0} = {1}\n"
    luaCodeItem2 = "\t\tobj.{0} = {1}{2}{3}\n"
    luaCodeEnd = "endfunction"
    jassCodeDefine = "\tset udg_{0}[{1}] = '{2}'\n"# 0 变量名 1 数字 2道具ID

    window = tk.Tk()
    window.title('物编生成器 龙吟水上 1243391642')
    window.geometry('400x100')
    window.resizable(width=False, height=False)

    PathTip = tk.Label(window, text='文件路径', font='微软雅黑')
    FilePath = tk.Text(window, width=35, height=1, )
    SkillTip = tk.Label(window, text='表格名称', font='微软雅黑')
    sheetName = tk.Text(window, width=35, height=1, )

    # path = sys.path[0]
    path = sys.argv[0]
    # 程序运行的话，用12，打包成exe的话，改成13
    path = path[:-12]
    # path = path[:-13]
    FilePath.insert(index=tk.INSERT, chars=path + "物编表.xlsx")

    Btn1 = tk.Button(window, text='道具', width=15, height=1, command=t1)
    Btn2 = tk.Button(window, text='技能', width=15, height=1, command=T2)
    Btn3 = tk.Button(window, text='技能物编', width=15, height=1, command=T3)
    Btn4 = tk.Button(window, text='返回', width=5, height=1, command=T4)

    Btn5 = tk.Button(window, text='描述效果', width=15, height=1, command=T5)
    Btn6 = tk.Button(window, text='道具物编', width=15, height=1, command=T1)
    Btn7 = tk.Button(window, text='返回', width=5, height=1, command=t2)

    PathTip.place(x=5, y=400)
    FilePath.place(x=80, y=400)
    SkillTip.place(x=5, y=300)
    sheetName.place(x=80, y=140)

    Btn1.place(x=30, y=60)
    Btn2.place(x=260, y=60)
    Btn3.place(x=260, y=300)
    Btn4.place(x=340, y=300)

    Btn5.place(x=30,  y=400)
    Btn6.place(x=260, y=400)
    Btn7.place(x=340, y=400)



    # 特性表
    tx_path = FilePath.get(index1=1.0, index2=tk.END)[:-1]
    tx_data = xlrd.open_workbook(tx_path)
    tx_sheet = tx_data.sheet_by_name('特性表')

    window.mainloop()