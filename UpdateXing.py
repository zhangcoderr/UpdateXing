# coding=utf-8
from pykeyboard import PyKeyboard
from pymouse import PyMouse
import time
import pyHook
import pythoncom
import xlrd
import xlwt
import pyperclip
from pynput import mouse, keyboard
import threading
import sys
from openpyxl import Workbook, load_workbook
import os


def copy():
    k.press_key(k.control_key)
    k.tap_key("c")  # 改小写！！！！ 大写的话由于单进程会触发shift键 ctrl键就失效了
    k.release_key(k.control_key)


def getCopy(maxTime=2):
    # maxTime = 3  # 3秒复制 调用copy() 不管结果对错
    while (maxTime > 0):
        maxTime = maxTime - 0.5
        time.sleep(0.5)
        # print('doing')
        copy()

    result = pyperclip.paste()
    return result


def tapkey(key, count=1):
    for i in range(0, count):
        k.tap_key(key)
        time.sleep(0.2)


def Do():
    global start
    global curExcelUrls
    global end
    if start:
        update_workbook = xlrd.open_workbook(updateExcelUrl)
        u_table = update_workbook.sheets()[6]#最后一个sheet
        u_rowCount = u_table.nrows
        u_colCount = u_table.ncols

        wb = load_workbook(filename=updateExcelUrl)
        load_worksheet = wb.active
        load_worksheet = wb['UpdateSheet']

        # 主代码---------------
        for curExcelUrl in curExcelUrls:
            workbook = xlrd.open_workbook(curExcelUrl)
            table = workbook.sheets()[0]
            rowCount = table.nrows
            colCount = table.ncols
            for i in range(rowCount):
                name=str(table.cell_value(i,0))#名称
                features=str(table.cell_value(i,1))#特征
                value=str(table.cell_value(i,3))#值

                for u_row in range(u_rowCount):
                    u_name = str(u_table.cell_value(i, 0))  # 需更新的名称
                    u_features = str(u_table.cell_value(i, 1))  # 需更新的特征
                    u_value = str(u_table.cell_value(i, 3))  # 需更新的值
                    if(u_name==name and u_features==features):
                        load_worksheet.cell(row=u_row, column=3, value=value)

        print('存储中……')
        wb.save(updateExcelUrl)
        print('存储完毕')
    start=False



def saveToExcel(name, type, unit, value):
    # saveworkbook = xlrd.open_workbook(saveExcelUrl)
    # wb = excel_copy(saveworkbook)  # 利用xlutils.copy下的copy函数复制
    wb = load_workbook(filename=saveExcelUrl)
    worksheet = wb.active
    worksheet = wb['Sheet1']
    global rowMaxCount
    # print(rowMaxCount)
    worksheet.cell(row=rowMaxCount + 1, column=1, value=name)
    worksheet.cell(row=rowMaxCount + 1, column=2, value=type)
    worksheet.cell(row=rowMaxCount + 1, column=3, value=unit)

    worksheet.cell(row=rowMaxCount + 1, column=4, value=value)

    wb.save(saveExcelUrl)
    rowMaxCount = rowMaxCount + 1


# 我的代码
def onpressed(Key):
    while True:
        # print(Key)
        if (Key == keyboard.Key.caps_lock):  # 开始
            global start
            start = True
            print('go')
        if (Key == keyboard.Key.f3):  # 结束
            if (saving):
                print('it''s saving !! not yet')
            else:
                sys.exit()
        # print(Key)
        return True


def main():
    while True:
        # 主程序在这
        Do()

def GetExcelUrls(curDirUrl):
    list=[]
    dirlist=os.listdir(curDirUrl)
    for url in dirlist:
        if('.xlsx') in url:
            list.append(curDirUrl+'\\'+url)
    return list
if __name__ == '__main__':
    k = PyKeyboard()
    m = PyMouse()
    scoll_count = 1
    start = False
    saving = False
    curCount = 0
    # saveExcelUrl = r"C:\Users\123\Desktop\广联达\安装\save.xlsx"  # to do-------------
    curExcelUrls=[]#当月的信息价的所有excel
    updateExcelUrl = r"C:\Users\Administrator\Desktop\save.xlsx"  # to do-------------
    curDirUrl=r'C:\Users\123\Desktop\广联达\安装'
    curExcelUrls =GetExcelUrls(curDirUrl)
    print(curExcelUrls)

    threads = []
    t2 = threading.Thread(target=main, args=())
    threads.append(t2)
    for t in threads:
        t.setDaemon(True)
        t.start()
    print('press Capital to start,记得手动关')

    with keyboard.Listener(on_press=onpressed) as listener:
        listener.join()

