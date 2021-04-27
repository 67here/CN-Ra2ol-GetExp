import ctypes
import os
from comtypes.client import CreateObject
from win32com.client import Dispatch
def nreg():
    try:
        dm = Dispatch('dm.dmsoft')
        print('本机系统中已经安装大漠插件，版本为:', dm.ver())
    except:
        dms = ctypes.windll.LoadLibrary('DmReg.dll')
        location_dmreg = os.getcwd()+'\dm.dll'
        dms.SetDllPathW(location_dmreg, 0)
        dm = CreateObject('dm.dmsoft')
        print('免注册成功 版本号为:',dm.Ver())

    s = dm.Reg('jv965720b239b8396b1b7df8b768c919e86e10f', 'ziucqz60')
    if s == 1:
        print('高级功能启用')
    else:
        print('会员功能异常，异常值为:',s)
