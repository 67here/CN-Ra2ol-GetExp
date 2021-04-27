import os
import time
import win32api
from win32com.client import Dispatch
from noreg import nreg
from dpi import checkdpi

#免注册 DPI检测
nreg()
checkdpi()

#准备调用
dm = Dispatch('dm.dmsoft')
dm.setpath(os.getcwd())
dm.SetMouseDelay("normal",30)
dm.SetDict(0,"dm_soft.txt")
find = ['-1','-1','-1']
exp = i = 0
x = y = 1

#1.启动程序
print(r'请输入新平台快捷方式的路径,格式如：C:\\yy\\ra2.lnk 注：后缀为lnk ')
#xx = input(':',)
win32api.ShellExecute(0, 'open','c:\\1.lnk', '', '', 1)
time.sleep(3)

#2.点击登录
hwnd1 = dm.FindWindowByProcess('Client.exe','HwndWrapper','公测平台')
dm.SetWindowState(hwnd1,1)
ret = dm.BindWindow(hwnd1,"gdi","dx","windows",0)
if ret == 1 :
    pass
else:
    print('窗口绑定失败')

print('查找\"登录\"按钮中',end='')

while find[0] != 0 :
    find = dm.FindStrFast(245,251,292,272,"登录","#255-20|#209-0",1.0,x,y)
    time.sleep(1)
print('\t查找完毕')
find = ['-1','-1','-1']
dm.moveto(268,262)
dm.LeftClick()
time.sleep(1)
ret = dm.UnBindWindow()

#3 关闭公告栏
print('关闭公告栏')
hwnd2 = dm.FindWindowByProcess('Client.exe','HwndWrapper','MainWindow')
time.sleep(2)
dm.MoveWindow(hwnd2,0,0)
dm.moveto(487,513)
time.sleep(1)
dm.LeftClick()
time.sleep(2)

#4 进入频道
print('查找\"频道2\"按钮',end='')

while find[0] != 0 :
    find = dm.FindStrFast(250,101,306,131,"频道2","#85-30|#35-0",1.0,x,y)
    time.sleep(1)
print('\t查找完毕')
find = ['-1','-1','-1']
dm.moveto(446,118)
dm.LeftClick()
time.sleep(1)
dm.LeftDoubleClick()
time.sleep(2)

while 1 :
    #5 建房
    find = ['-1', '-1', '-1']
    print('查找\"创建\"按钮',end='')
    while find[0] != 0 :
        find = dm.FindStrFast(596,639,652,658,"创建房间","#0-0|#67-35",1.0,x,y)
        time.sleep(1)
    print('\t查找完毕')
    find = ['-1','-1','-1']
    dm.moveto(620,649)
    time.sleep(1)
    dm.LeftClick()
    time.sleep(1)
    hwnd3 = dm.FindWindowByProcess('Client.exe','HwndWrapper','创建房间')
    dm.MoveWindow(hwnd3,0,0)
    time.sleep(1)
    find = dm.FindStrFast(8,205,76,237,"测试未开启","#238-100",1.0,x,y)
    if find[0] == -1 :
        print('不点击测试')
    else:
        print('点击测试')
        dm.moveto(32,221)
        time.sleep(1)
        dm.LeftClick()
    find = ['-1','-1','-1']
    time.sleep(1)
    dm.moveto(152,327)
    dm.LeftClick()
    time.sleep(2)

    #6 开始游戏
    print('查找\"开始游戏\"',end='')
    while find[0] != 0 :
        find = dm.FindStrFast(599,644,686,668,"开始游戏","#226-30",1.0,x,y)
        time.sleep(1)
    print('\t查找完毕')
    os.system("time 1:00:00")
    dm.moveto(641,657)
    time.sleep(1)
    dm.LeftClick()
    time.sleep(4)
    os.system("time 2:00:00")
    os.system("taskkill /im gamemd-spawn.exe /f")

    #7 结算
    time.sleep(4)
    exp1 = dm.ocr(694,170,731,206,'ffff00-000000',1.0)
    if int(exp1) == 87 :
        exp = int(exp1) + int(exp)
        i += 1
    elif exp1 == 35 :
        exp = int(exp1) + int(exp)
        i = i + 1
    else:
        print('当前无经验，本局结束')
        print('目前总共获得经验:', exp)
        loop = input('是否结束游戏? y/n \n')
        if loop == 1 :
            break

    print('当前是第',i,'局')
    print('本局获得经验:',exp1)
    print('目前总共获得经验:',exp)
    dm.moveto(831,434)
    time.sleep(1)
    dm.LeftClick()

