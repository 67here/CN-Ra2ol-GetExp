from pywinauto import application
import os
from time import sleep
import win32api

print('欢迎使用 请使用管理员身份运行')
app = application.Application(backend='uia')
xx = os.getcwd()+'\新平台.lnk'
win32api.ShellExecute(0, 'open',xx, '', '', 1)
sleep(3)
app.connect(title_re='公测平台', class_name='Window')
app.window(title_re="公测平台")['登录'].click()
sleep(2)
app.connect(title='MainWindow', class_name='Window')
app['MainWindow']['关闭'].click()
sleep(2)
app['MainWindow'].listbox['Client.ListItemProp_Lobby'].select().click_input(button='left', double=True)
sleep(2)
i = 1
allexp = ['0','0']
while i :
    i += 1
    app['MainWindow']['创建房间'].click()
    sleep(2)
    app.connect(title='创建房间', class_name='Window')
    app['创建房间'].child_window(auto_id="pwd", control_type="Edit").type_keys('1')
    sleep(1)
    app['创建房间']['创建房间（Enter）'].click()
    os.system("time 1:00:00")
    app.connect(title='MainWindow', class_name='Window')
    app['MainWindow']['开始游戏'].click()
    sleep(4)
    os.system("time 2:00:00")
    os.system("taskkill /im gamemd-spawn.exe /f")
    sleep(4)
    lst = app['MainWindow'].child_window(auto_id="Frame_Right_Lobby", control_type="Pane").Static.texts()
    list1 = list(lst[0])
    everyexp = list1[-5:-3]
    print('第', i - 1, '局')
    if list1[-4] == '经':
        loop = input('本局经验已上限 是否停止 y/n \n')
        if loop == 'y':
            print('脚本结束,当前总经验为:',''.join(allexp))
            print('本脚本是免费的，作者b站ID:67here,记得修正系统时间')
            sleep(5)
            break
        else:
            app['MainWindow']['返回大厅'].click()
    else:
        allexp = list1[19:-7]
        print('当前总经验为:',''.join(allexp))
        print('本局经验增加:',''.join(everyexp))
        app['MainWindow']['返回大厅'].click()