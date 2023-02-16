# -*- coding: utf-8 -*- 
""" ++++++++++++++++++++++++++++++++++++++
@product->name PyCharm
@project->name WeChatAuto
@editor->name Sanliy
@file->name wxauto.py
@create->time 2023/1/31-10:14
@desc->
++++++++++++++++++++++++++++++++++++++ """
import math
import operator
import os
import re
import subprocess
import time
from functools import reduce

import pyautogui
import uiautomation as auto
from PIL import Image, ImageChops
from pyautogui import hotkey


class WxAuto:
    def __init__(self, window_class_name, window_name):
        self.handle = None
        self.search = None
        self.window_class_name = window_class_name
        self.window_name = window_name
        self.black_name = ["新的朋友", "公众号", "世纪空间", ""]

        self.up_pic_path1 = '../images/up_image1.png'
        self.up_pic_path2 = '../images/up_image2.png'

        self.down_pic_path1 = '../images/down_image1.png'
        self.down_pic_path2 = '../images/down_image2.png'

        self.contacts_list = []

    def catch_wx_window(self):
        handle = auto.WindowControl(searchDepth=1, className=self.window_class_name, Name=self.window_name)
        auto.SendKeys(text='{Alt}{Ctrl}w')
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
        self.handle.ButtonControl(Name='聊天信息')
        args = ['powershell', f'Get-Item {pic_path} | Set-Clipboard']
        subprocess.Popen(args=args)
        time.sleep(0.5)
        self.handle.EditControl(Name="输入").SendKeys(text='{Ctrl}v', waitTime=0.2)
        send_but = self.handle.ButtonControl(Name='sendBtn')
        send_but.Click()

    def easy_chat_flow(self, search_name, message_content):
        self.catch_wx_window()
        self.search_contact(search_name)
        self.open_chat_window(search_name)
        self.send_message(message_content)

    def send_pic_flow(self, search_name, pic_path):
        self.catch_wx_window()
        self.search_contact(search_name)
        self.open_chat_window(search_name)
        self.send_pic(pic_path)

    def get_all_contacts(self):
        up_flag = True
        down_flag = True
        pyautogui.moveRel(80, 0)
        currentMouseX, currentMouseY = pyautogui.position()
        while down_flag:
            contacts = self.handle.ListControl(Name="联系人")
            if os.path.exists(self.down_pic_path1) and os.path.exists(self.down_pic_path2):
                os.remove(self.down_pic_path1)
                os.remove(self.down_pic_path2)
            while up_flag:
                if os.path.exists(self.up_pic_path1) and os.path.exists(self.up_pic_path2):
                    os.remove(self.up_pic_path1)
                    os.remove(self.up_pic_path2)
                # 移动鼠标到联系人所在的位置，然后开始向上滚动
                self.handle.CaptureToImage(self.up_pic_path1)
                auto.WheelUp(waitTime=0.01)
                self.handle.CaptureToImage(self.up_pic_path2)
                up_img_one = Image.open(self.up_pic_path1)
                up_img_two = Image.open(self.up_pic_path2)

                h1 = up_img_one.histogram()
                h2 = up_img_two.histogram()

                result = math.sqrt(reduce(operator.add, list(map(lambda a, b: (a - b) ** 2, h1, h2))) / len(h1))
                if result == 0:
                    up_flag = False
                os.remove(self.up_pic_path1)
                os.remove(self.up_pic_path2)
            flag = 0
            for contact in contacts.GetChildren():
                contact_name = contact.Name
                print("+++++++++++++++++++++++++++++")
                print(contact_name)
                print("+++++++++++++++++++++++++++++")
                detail_msg = {}
                if contact_name not in self.black_name and not re.search(contact_name, str(self.contacts_list)):
                    try:
                        name_but = contact.GetChildren()[0].ButtonControl(Name=contact_name)
                        name_but.Click()
                    except Exception as error:
                        print("当前联系人获取失败[{0}]".format(contact_name))
                        continue
                    try:
                        nc_control = self.handle.TextControl(Name="昵称：")
                        name_control = nc_control.GetNextSiblingControl()

                        detail_msg[nc_control.Name] = name_control.Name
                    except Exception as error:
                        print("当前联系人不存在{0}".format("昵称属性"))
                    try:
                        wxh_control = self.handle.TextControl(Name="微信号：")
                        number_control = wxh_control.GetNextSiblingControl()

                        detail_msg[wxh_control.Name] = number_control.Name
                    except Exception as error:
                        print("当前联系人不存在{0}".format("微信号"))
                    try:
                        dq_control = self.handle.TextControl(Name="地区：")
                        area_control = dq_control.GetNextSiblingControl()
                        detail_msg[dq_control.Name] = area_control.Name
                    except Exception as error:
                        print("当前联系人不存在{0}".format("地区属性"))

                    try:
                        bzm_control = self.handle.TextControl(Name="备注名")

                        bzm_name_control = bzm_control.GetNextSiblingControl().GetChildren()[0]

                        detail_msg[bzm_control.Name] = bzm_name_control.Name
                    except Exception as error:
                        print("当前联系人不存在{0}".format("备注名属性"))

                    try:
                        ly_control = self.handle.TextControl(Name="来源")
                        source_control = ly_control.GetNextSiblingControl()

                        detail_msg[ly_control.Name] = source_control.Name
                    except Exception as error:
                        print("当前联系人不存在{0}".format("来源属性"))
                    if detail_msg != {}:
                        self.contacts_list.append(detail_msg)
            # 复原鼠标位置
            pyautogui.moveTo(currentMouseX, currentMouseY)
            # 移动鼠标到联系人所在的位置，然后开始向下滚动
            self.handle.CaptureToImage(self.down_pic_path1)
            auto.WheelDown(waitTime=1)
            print("+++++++++++++++++++++++++++++++++++")
            print("休息10秒")
            time.sleep(10)
            print("+++++++++++++++++++++++++++++++++++")
            self.handle.CaptureToImage(self.down_pic_path2)
            down_img_one = Image.open(self.down_pic_path1)
            down_img_two = Image.open(self.down_pic_path2)

            h1 = down_img_one.histogram()
            h2 = down_img_two.histogram()

            result = math.sqrt(reduce(operator.add, list(map(lambda a, b: (a - b) ** 2, h1, h2))) / len(h1))
            if result == 0:
                down_flag = False
            os.remove(self.down_pic_path1)
            os.remove(self.down_pic_path2)
        print(self.contacts_list)

    def get_contacts_name_flow(self):
        self.catch_wx_window()
        self.get_contacts()
        self.get_all_contacts()


if __name__ == '__main__':
    wx = WxAuto(
        'WeChatMainWndForPC',
        '微信'
    )
    wx.easy_chat_flow(
        'xxx',
        'test'
    )
    # wx.send_pic_flow(
    #     '文件传输助手',
    #     r"D:\test\metadata.21at",
    # )
    # wx.get_contacts_name_flow()
