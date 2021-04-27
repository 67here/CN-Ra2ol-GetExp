from win32 import win32api, win32gui, win32print
from win32.lib import win32con
from win32.win32api import GetSystemMetrics
import time

def get_real_resolution():
    """获取真实的分辨率"""
    hDC = win32gui.GetDC(0)
    # 横向分辨率
    w = win32print.GetDeviceCaps(hDC, win32con.DESKTOPHORZRES)
    # 纵向分辨率
    h = win32print.GetDeviceCaps(hDC, win32con.DESKTOPVERTRES)
    return w, h

def get_screen_size():
    """获取缩放后的分辨率"""
    w = GetSystemMetrics (0)
    h = GetSystemMetrics (1)
    return w, h

def getapi():

    real_resolution = get_real_resolution()
    screen_size = get_screen_size()

    screen_scale_rate = round(real_resolution[0] / screen_size[0], 2)
    screen_scale_rate = screen_scale_rate * 100
    return screen_scale_rate

def checkdpi():
    dpi = getapi()
    print('当前系统缩放率为:',int(dpi),'%')
    if dpi == 100 :
        pass
    else:
        print('请将win系统缩放率设置到100%')
        print('程序即将关闭')
        time.sleep(5)
        os.exit(0)
