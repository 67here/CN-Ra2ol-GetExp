from win32com.client import Dispatch
import time
def rg():
    try:
        dm = Dispatch('dm.dmsoft')
        print('本机已经安装大漠插件，版本为:', dm.ver())
        reg = '0'
    except:
        print('调用出错 您还没有安装大漠插件')
        reg = '1'

    if reg == '1' :
        x = input('是否需要注册大漠插件？ y/n \n')
        if x == 'y':
            import os
            localtion1 = 'regsvr32 '+os.getcwd()+'\dm.dll'
            print(localtion1)
            os.system(localtion1)
            dm = Dispatch('dm.dmsoft')
            if dm.ver() == '7.2116' :
                print('插件注册成功！版本号:',dm.ver())
            else:
                print('版本号为',dm.ver())
        elif x == 'n':
                print('再见!')
                time.sleep(2)
        else:
            print('不要输入其它字符哦')
            time.sleep(2)