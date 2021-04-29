import os
import time
import win32api
from win32com.client import Dispatch
from noreg import nreg
from dpi import fixdpi
import sys

#解压文件
def get_resource_path(relative_path):
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)


#免注册 DPI更改
nreg()
userdpi = fixdpi()

#准备调用
dm = Dispatch('dm.dmsoft')
dm.setpath(os.getcwd())
dm.SetMouseDelay("normal",30)
dm.SetDict(0,get_resource_path('.\dm_soft.txt'))
find = ['-1','-1','-1']
print('操作开始后 不要随意动鼠标和键盘~')
print('请把此脚本和新平台的快捷方式放在一起 才能运行')
ju = input('请设定循环的局数 \'即整数数字\' \n')
ju = int(ju)
x = y = i = 1

#1.启动程序

xx = os.getcwd()+'\新平台.lnk'
win32api.ShellExecute(0, 'open',xx, '', '', 1)
time.sleep(3)

#2.点击登录
hwnd1 = dm.FindWindowByProcess('Client.exe','HwndWrapper','公测平台')
dm.SetWindowState(hwnd1,1)
ret = dm.BindWindow(hwnd1,"gdi","dx","windows",0)
if ret == 1 :
    pass
else:
    print('登录窗口绑定失败')

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

for i in range(ju) :
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
    print('当前是第', i+1, '局')
    time.sleep(1)
    print('\n')
    dm.moveto(831,434)
    time.sleep(1)
    dm.LeftClick()
print('您设定的循环',ju,'局 已经完成')
print('缩放即将还原，请手动更正系统时间~')
print('本脚本是免费的，作者b站id是67here~')
time.sleep(10)
userdpi = str(userdpi)
win32api.ShellExecute(0, 'open', get_resource_path('.\SetDpi.exe'), userdpi, '', 1)
