#!/usr/bin/python3
# -*- coding:utf-8 -*-

# build ver2.0

import ctypes, sys
from tkinter import *
import pywintypes
import win32gui, win32con
import time, threading

# get admin
def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

if is_admin():
    # create window
    root = Tk()
    root.title('Auto PCR Sweeper')
    width = 280
    height = 70
    screenwidth = root.winfo_screenwidth()
    screenheight = root.winfo_screenheight()
    size = '%dx%d+%d+%d' % (width, height, (screenwidth - width)/2, (screenheight - height)/2)
    root.geometry(size)

    root.resizable(0, 0)

    # window frame
    frame1 = Frame()
    frame1.pack(side = TOP)
    frame2 = Frame()
    frame2.pack(side = TOP)

    # frame1
    Label(frame1, text = '扫荡次数：').grid(row = 0, column = 0, padx = 2)

    timein = Entry(frame1, width = 15)
    timein.grid(row = 0, column = 1)

    # frame2
    mestr = StringVar()
    mestr.set('输入扫荡次数后按开始或回车。')

    Label(frame2, textvariable = mestr).grid(row = 3, column =0, pady = 3)

    def main():
        mestr.set('')
        count = timein.get()
        
        if(count == ''):
            mestr.set('未输入扫荡次数。')
        else:
            # check & get hwnd of PCR in simulator
            hwnd = win32gui.FindWindow('Qt5QWindowIcon', '公主连结R - MuMu模拟器')
            time.sleep(0.1)

            # circulate PCR sweep process
            now = 0
            while(now < int(count)):
                win32gui.PostMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_NUMPAD1, 0)
                win32gui.PostMessage(hwnd, win32con.WM_KEYUP, win32con.VK_NUMPAD1, 0)
                time.sleep(0.6)

                win32gui.PostMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_NUMPAD2, 0)
                win32gui.PostMessage(hwnd, win32con.WM_KEYUP, win32con.VK_NUMPAD2, 0)
                time.sleep(3.2)

                win32gui.PostMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_NUMPAD3, 0)
                win32gui.PostMessage(hwnd, win32con.WM_KEYUP, win32con.VK_NUMPAD3, 0)
                time.sleep(0.8)

                win32gui.PostMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_NUMPAD4, 0)
                win32gui.PostMessage(hwnd, win32con.WM_KEYUP, win32con.VK_NUMPAD4, 0)
                time.sleep(0.6)

                now += 1

            mestr.set('本次扫荡完毕。')

    # multi-threading
    def thread(func):
        th = threading.Thread(target = func)
        th.setDaemon(True)
        th.start()

    def touchmain(self):
        thread(main)

    btn = Button(frame1, text = '开始', width = 10, command = lambda: thread(main))
    btn.grid(row = 0, column = 2, padx=15, pady =5)

    root.bind("<Return>", touchmain)

    root.mainloop()

else:
    ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, __file__, None, 1)
