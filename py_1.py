import os
import time
import win32api
from Reg import rg
from win32com.client import Dispatch
#注册大漠
rg()
#准备调用
dm = Dispatch('dm.dmsoft')
ret = '-1|-1|-1'
dm.setpath(os.getcwd())
#1.启动程序
print(r'请输入新平台的路径,格式比如：C:\\yy\\ra2.lnk 注：快捷方式后缀为lnk ')
xx = input(':',)
win32api.ShellExecute(0, 'open',xx, '', '', 1)
#2.点击登录
print('平台已启动')
while ret == '-1|-1|-1' :
    ret = dm.FindpicE(0, 0, 1920, 1080, '01login.bmp', '050505', 0.9, 0)
    time.sleep(1)
pos = ret.split('|')
dm.moveto(pos[1], pos[2])
time.sleep(2)
dm.LeftClick()
ret = '-1|-1|-1'
#3 关闭公告栏
print('登录完成')
while ret == '-1|-1|-1' :
    ret = dm.FindpicE(0, 0, 1920, 1080, '02guanbi.bmp', '050505', 0.9, 0)
    time.sleep(1)
pos = ret.split('|')
dm.moveto(pos[1], pos[2])
time.sleep(2)
dm.LeftClick()
time.sleep(2)
ret = '-1|-1|-1'
#4 进入频道
while ret == '-1|-1|-1' :
    ret = dm.FindpicE(0, 0, 1920, 1080, '03pindao.bmp', '050505', 0.9, 0)
    time.sleep(1)
pos = ret.split('|')
dm.moveto(pos[1], pos[2])
time.sleep(1)
dm.MoveR(100,0)
dm.LeftDoubleClick()
for i in range(1,10,1) :
    #5 创建房间
    print('频道进入成功')
    ret = '-1|-1|-1'
    while ret == '-1|-1|-1':
        ret = dm.FindpicE(0, 0, 1920, 1080, '04jianfang.bmp', '050505', 0.9, 0)
        time.sleep(1)
    pos = ret.split('|')
    dm.moveto(pos[1], pos[2])
    time.sleep(2)
    dm.LeftClick()
    ret = '-1|-1|-1'
    #6 密码
    time.sleep(1)
    ret = dm.FindpicE(0, 0, 1920, 1080, '05pwd.bmp', '050505', 0.9, 0)
    if ret != '-1|-1|-1' :
        pos = ret.split('|')
        dm.moveto(pos[1], pos[2])
        time.sleep(1)
        dm.MoveR(10,10)
        dm.LeftClick()
        print('测试已开启')
    else:
        print('不点击测试')
    ret = '-1|-1|-1'
    #7 创建成功
    while ret == '-1|-1|-1':
        ret = dm.FindpicE(0, 0, 1920, 1080, '06creat.bmp', '050505', 0.9, 0)
        time.sleep(1)
    pos = ret.split('|')
    dm.moveto(pos[1], pos[2])
    time.sleep(1)
    dm.LeftClick()
    #8 修改时间 1次
    print("建房成功")
    os.system("time 1:00:00")
    time.sleep(1)
    ret = '-1|-1|-1'
    #9 开始游戏
    while ret == '-1|-1|-1':
        ret = dm.FindpicE(0, 0, 1920, 1080, '07start.bmp', '050505', 0.9, 0)
        time.sleep(1)
    pos = ret.split('|')
    dm.moveto(pos[1], pos[2])
    time.sleep(1)
    dm.LeftClick()
    time.sleep(1)
    dm.LeftClick()
    time.sleep(4)
    #10 改时间2次
    os.system("time 2:00:00")
    #11 杀进程
    os.system("taskkill /im gamemd-spawn.exe /f")
    ret = '-1|-1|-1'
    #12 返回大厅
    while ret == '-1|-1|-1':
        ret = dm.FindpicE(0, 0, 1920, 1080, '08return.bmp', '050505', 0.9, 0)
        time.sleep(1)
    pos = ret.split('|')
    dm.moveto(pos[1], pos[2])
    time.sleep(1)
    dm.LeftClick()
    print('已经循环执行',i,'次')
    print('\n')
    time.sleep(1)

