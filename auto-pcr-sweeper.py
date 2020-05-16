#!/usr/bin/python3
# -*- coding:utf-8 -*-

# build ver1.3

from tkinter import *
from tkinter import ttk
import win32api, win32gui, win32con
import time
import pywintypes
import ctypes, sys

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
    height = 100
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
    Label(frame1, text = '模拟器：').grid(row = 0, column = 0)
    Label(frame1, text = '扫荡次数：').grid(row = 1, column = 0)

    timein = Entry(frame1, width = 15)
    timein.grid(row = 1, column = 1)

    smltin = ttk.Combobox(frame1, width = 12)
    smltin['values'] = ('MuMu', '雷电')
    smltin.grid(row = 0, column = 1)

    # frame2
    mestr = StringVar()
    mestr.set('选择模拟器并输入扫荡次数后按开始或回车。')

    Label(frame2, textvariable = mestr).grid(row = 3, column =0, pady = 5)

    def main():
        mestr.set('确保模拟器的游戏已经开启。')

        smltc = smltin.get()
        count = timein.get()
        
        if(smltc == '' or count == ''):
            if(smltc == ''):
                mestr.set('未选择模拟器。')
            else:
                mestr.set('未输入扫荡次数。')
        else:
            # check & get hwnd of PCR in simulator
            smltc = smltin.get()
            if(smltc == 'MuMu'):
                hwnd = win32gui.FindWindow('Qt5QWindowIcon', '公主连结R - MuMu模拟器')
            elif(smltc == '雷电'):
                hwnd = win32gui.FindWindow('LDPlayerMainFrame', '雷电模拟器')

            # switch to game window
            win32gui.SendMessage(hwnd, win32con.WM_SYSCOMMAND, win32con.SC_RESTORE, 0)
            win32gui.SetForegroundWindow(hwnd)

            time.sleep(0.1)

            # circulate PCR sweep process
            now = 0
            while(now < int(count)):
                win32api.keybd_event(65, win32api.MapVirtualKey(65, 0), 0, 0) # simulate type A
                win32api.keybd_event(65, win32api.MapVirtualKey(65, 0), win32con.KEYEVENTF_KEYUP, 0)
                time.sleep(0.6)

                win32api.keybd_event(66, win32api.MapVirtualKey(66, 0), 0, 0) # simulate type B
                win32api.keybd_event(66, win32api.MapVirtualKey(66, 0), win32con.KEYEVENTF_KEYUP, 0)
                time.sleep(3.3)

                win32api.keybd_event(67, win32api.MapVirtualKey(67, 0), 0, 0) # simulate type C
                win32api.keybd_event(67, win32api.MapVirtualKey(67, 0), win32con.KEYEVENTF_KEYUP, 0)
                time.sleep(0.75)

                win32api.keybd_event(68, win32api.MapVirtualKey(68, 0), 0, 0) # simulate type D
                win32api.keybd_event(68, win32api.MapVirtualKey(68, 0), win32con.KEYEVENTF_KEYUP, 0)
                time.sleep(0.55)

                now += 1

            mestr.set('本次扫荡完毕。')

    def touchmain(self):
        main()

    btn = Button(frame1, text = '开始', width = 10, command = main)
    btn.grid(row = 0, column = 2, padx=15, pady =5)

    root.bind('<Return>', touchmain)

    root.mainloop()

else:
    ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, __file__, None, 1)
