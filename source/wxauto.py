# -*- coding: utf-8 -*- 
""" ++++++++++++++++++++++++++++++++++++++
@product->name PyCharm
@project->name WeChatAuto
@editor->name Sanliy
@file->name wxauto.py
@create->time 2023/1/31-10:14
@desc->
++++++++++++++++++++++++++++++++++++++ """
import subprocess
import time

import win32gui, win32con
import uiautomation as auto


class WxAuto:
    def __init__(self, window_class_name, window_name):
        self.handle = None
        self.search = None
        self.win32_handle = None
        self.window_class_name = window_class_name
        self.window_name = window_name

    def open_wx_window(self):
        handle = win32gui.FindWindow(self.window_class_name, None)
        if handle != 0:
            # 如果窗口被最小化了，就恢复原先的尺寸
            if win32gui.IsIconic(handle):
                win32gui.ShowWindow(handle, win32con.SW_SHOWNORMAL)
            # 将窗口置顶
            win32gui.SetForegroundWindow(handle)
        self.win32_handle = handle

    def catch_wx_window(self):
        handle = auto.WindowControl(searchDepth=1, className=self.window_class_name, Name=self.window_name)
        self.handle = handle

    def search_contact(self, search_name):
        if self.handle == 0:
            return 0
        search = self.handle.EditControl(Name='搜索')
        search.Click()
        search.GetParentControl().GetChildren()[1].SendKeys(search_name)
        search_result = self.handle.ListControl(Name='搜索结果').GetChildren()
        for sear in search_result:
            if sear.Name == search_name:
                sear.Click()
                break

    def open_chat_window(self, search_name):
        self.handle.ListControl(Name="会话")
        self.handle.ButtonControl(Name=search_name)

    def send_message(self, message_content):
        self.handle.ButtonControl(Name='聊天信息')

        self.handle.EditControl(Name="输入").SendKeys(message_content)
        send_but = self.handle.ButtonControl(Name='sendBtn')
        send_but.Click()

    def get_contacts(self):
        button = self.handle.ButtonControl(Name='通讯录')
        button.Click()

    def send_pic(self, pic_path):
        # image = Image.open(pic_path)
        self.handle.ButtonControl(Name='聊天信息')
        args = ['powershell', f'Get-Item {pic_path} | Set-Clipboard']
        subprocess.Popen(args=args)
        time.sleep(0.5)
        self.handle.EditControl(Name="输入").SendKeys(text='{Ctrl}v', waitTime=0.2)
        send_but = self.handle.ButtonControl(Name='sendBtn')
        send_but.Click()

    def easy_chat_flow(self, search_name, message_content):
        self.open_wx_window()
        self.catch_wx_window()
        self.search_contact(search_name)
        self.open_chat_window(search_name)
        self.send_message(message_content)

    def send_pic_flow(self, search_name, pic_path):
        self.open_wx_window()
        self.catch_wx_window()
        self.search_contact(search_name)
        self.open_chat_window(search_name)
        self.send_pic(pic_path)

    def get_contacts_all(self):
        self.open_wx_window()
        self.catch_wx_window()
        self.get_contacts()


if __name__ == '__main__':
    wx = WxAuto(
        'WeChatMainWndForPC',
        '微信'
    )
    # wx.easy_chat_flow(
    #     '文件传输助手',
    #     'test'
    # )
    wx.send_pic_flow(
        '文件传输助手',
        r"D:\Win版 AE 2021（64bit）.rar",
    )
