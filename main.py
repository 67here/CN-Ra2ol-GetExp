import os
import time
import win32api
from win32com.client import Dispatch
from noreg import nreg
from dpi import fixdpi
import sys

#ocr的准备工作！
import json
import base64
from urllib.request import urlopen
from urllib.request import Request
from urllib.error import URLError
from urllib.parse import urlencode
import ssl


def fetch_token():
    params = {'grant_type': 'client_credentials',
              'client_id': 'XDOIwHjkjrMjyBs2lstmEKYx',
              'client_secret': 'Ec7ZNzja6GMRzlSoTcON8sTGxR35F4n7'}
    post_data = urlencode(params)
    post_data = post_data.encode('utf-8')
    req = Request(TOKEN_URL, post_data)
    try:
        f = urlopen(req, timeout=5)
        result_str = f.read()
    except URLError as err:
        print(err)
        result_str = result_str.decode()
    result = json.loads(result_str)

    if ('access_token' in result.keys() and 'scope' in result.keys()):
        if not 'brain_all_scope' in result['scope'].split(' '):
            print ('please ensure has check the  ability')
            exit()
        return result['access_token']
    else:
        print ('please overwrite the correct API_KEY and SECRET_KEY')
        exit()
def read_file(image_path):
    f = None
    try:
        f = open(image_path, 'rb')
        return f.read()
    except:
        print('read image file fail')
        return None
    finally:
        if f:
            f.close()

def request(url, data):
    req = Request(url, data.encode('utf-8'))
    has_error = False
    try:
        f = urlopen(req)
        result_str = f.read()
        result_str = result_str.decode()
        return result_str
    except  URLError as err:
        print(err)

def baiduocr():
    if __name__ == '__main__':
        token = fetch_token()
        image_url = OCR_URL + "?access_token=" + token
        text = ""
        file_content = read_file('.\score.bmp')
        result = request(image_url, urlencode({'image': base64.b64encode(file_content)}))
        result_json = json.loads(result)
        for words_result in result_json["words_result"]:
            text = text + words_result["words"]
        print(text)

ssl._create_default_https_context = ssl._create_unverified_context
API_KEY = 'GmhC18eVP1Fo1ECX911dtOzw'
SECRET_KEY = 'PQ2ukO4Aec2PTsgQU9UkiEKYciavlZk8'
OCR_URL = "https://aip.baidubce.com/rest/2.0/ocr/v1/accurate_basic"
TOKEN_URL = 'https://aip.baidubce.com/oauth/2.0/token'


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
    dm.Capture(374,169,785,233,".\score.bmp")
    baiduocr()
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
