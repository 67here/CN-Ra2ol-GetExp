import ctypes
import os
from comtypes.client import CreateObject
from win32com.client import Dispatch
import sys
#解压文件
def get_resource_path(relative_path):
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)

def nreg():
    try:
        dm = Dispatch('dm.dmsoft')
        print('本机系统中已经安装大漠插件，版本为:', dm.ver())
    except:
        dms = ctypes.windll.LoadLibrary(get_resource_path('.\DmReg.dll'))
        dms.SetDllPathW(get_resource_path('.\dm.dll'), 0)
        dm = CreateObject('dm.dmsoft')
        print('免注册成功 版本号为:',dm.Ver())

    s = dm.Reg('jv965720b239b8396b1b7df8b768c919e86e10f', 'ziucqz60')
    if s == 1:
        print('高级功能启用')
    elif s == -2:
        print('程序未使用管理员身份运行')
    else:
        print('会员功能异常，异常值为:',s)
