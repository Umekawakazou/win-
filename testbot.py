'''
author：tongyao.zhang
只需要用pip3装两个库，cmd命令pip3 install pywin32 apscheduler   
其中pywin32是python在windows系统下针对所有基本操作的api包，可以慢慢研究，如果在linux或者mac上使用就需要用pyautoui这个库
apscheduler是个定时任务常用框架，跨系统的。
用这个脚本时候，要把微信电脑客户端打开，同时单独打开目标的聊天窗口，可能会有冗余的设计，不过不影响性能。
'''

import win32api, win32gui, win32con
import win32clipboard as clipboard
import time
from apscheduler.schedulers.blocking import BlockingScheduler
'''
微信自动发送功能

'''
def send_m(win):
    # 以下为“CTRL+V”组合键,粘贴文本内容后回车发送
    win32api.keybd_event(17, 0, 0, 0)  # 模拟按下CTRL
    time.sleep(0.5)  
    win32gui.SendMessage(win, win32con.WM_KEYDOWN, 86, 0)  # 模拟按下V
    win32api.keybd_event(17, 0, win32con.KEYEVENTF_KEYUP, 0)  # 模拟放开CTRL
    time.sleep(1)  
    win32gui.SendMessage(win, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)  # 模拟回车
    return
def txt_ctrl_v(txt_str):
    #自动化操作剪切板,将信息缓存入剪贴板
    clipboard.OpenClipboard()
    clipboard.EmptyClipboard()
    clipboard.SetClipboardData(win32con.CF_UNICODETEXT, txt_str)
    clipboard.CloseClipboard()
    return

def get_window(className, titleName):
    title_name = className  # 单独打开，好友名称
    win = win32gui.FindWindow(className, titleName)
    '''
    窗口显示到最前面
     win32gui.SetForegroundWindow(win)
     使窗体最大化

    '''
    win32gui.ShowWindow(win, win32con.SW_MAXIMIZE)
    win = win32gui.FindWindow(className, titleName)
    print("找到句柄：%x" % win)
    if win != 0:
        left, top, right, bottom = win32gui.GetWindowRect(win) #从左上到右下，最大化聊天窗口
        print(left, top, right, bottom)  # 执行上面的命令
        win32gui.SetForegroundWindow(win)  # 获取页面控制权
        time.sleep(0.5)
    else:
        print('请注意：找不到[%s]这个人（或群），请单独打开你们的聊天窗口！' % title_name)
    return win
#######################发送过程##########################
def sendTaskLog():
    # 查找微信小窗口
    # win = get_window('ChatWnd', '微信好友或者群名称')，备注什么名字就写什么名字
    win = get_window('ChatWnd', 'xxx') 
    # 读取文本，r表示read，小引号里表示你要发送的文本内容所在的路径
    file = open(r'C:\Users\zty22\Desktop\tasklog.txt', mode='r', encoding='UTF-8')
    str = file.read()
    print(str)
    txt_ctrl_v(str)
    send_m(win)
scheduler = BlockingScheduler()
'''
scheduler.add_job(sendTaskLog, 'interval', seconds=3)，想每各一段时间就发送，用interval，间隔的意思
misfire_grace_time是scheduler的一个子应用功能，是错过执行持续性的任务后，再过多长时间执行任务，写0就是放弃，单位都是默认秒。
下面cron是种日期格式，从周一到周日任意设定一段时间，最大mon-sun。默认简写，mon tus wen thu fri sat sun
'''
scheduler.add_job(sendTaskLog, 'cron', day_of_week='mon-fri', hour=23, minute=55, second=30, misfire_grace_time=30)

try:
    scheduler.start()
except (KeyboardInterrupt, SystemExit):
    pass
