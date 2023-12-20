# 使用提醒:
# 1. xbot包提供软件自动化、数据表格、Excel、日志、AI等功能
# 2. package包提供访问当前应用数据的功能，如获取元素、访问全局变量、获取资源文件等功能
# 3. 当此模块作为流程独立运行时执行main函数
# 4. 可视化流程中可以通过"调用模块"的指令使用此模块

import threading

import subprocess
import time
import winreg
import os
import sys
import win32com.client
import win32gui
import pythoncom

from functools import wraps
from win32com.client import GetObject
from win32gui import IsWindowEnabled, IsWindowVisible
from win32gui import GetWindowText, IsWindow


class SapGui:
    def __init__(self) -> None:
        self.session = None
        self.connection = None
        self.application = None
        self.SapGui = None
        self.get_object()

    def get_object_wrap(func):
        @wraps(func)
        def inner(self, *args, **kwargs):
            self.get_object()
            res = func(self, *args, **kwargs)
            return res

        return inner

    @get_object_wrap
    def login(self, connection, username, password, crop_id, after_login="2"):
        """登录SAP
        :param connection: str, 连接名
        :param username: str, 账号
        :param password: srr, 密码
        """
        assert (
            type(self.application) == win32com.client.CDispatch
        ), "请先打开 SAP 软件或检测SAP的配置"
        after_login_dict = {
            "1": "wnd[1]/usr/radMULTI_LOGON_OPT1",
            "2": "wnd[1]/usr/radMULTI_LOGON_OPT2",
            "3": "wnd[1]/usr/radMULTI_LOGON_OPT3",
        }

        if not self.connection:
            self.connection = self.application.OpenConnection(connection, True)
        else:
            self.connection = self.application.Children(0)
        # self.connection = self.application.OpenConnection(connection, True)

        time.sleep(2)

        try:
            self.session = self.connection.Children(0)
        except:
            raise Exception("请检查SAP配置是否正常")

        self.session.findById("wnd[0]").maximize()
        if crop_id:
            self.session.findById("wnd[0]/usr/txtRSYST-MANDT").text = crop_id
        self.session.findById("wnd[0]/usr/txtRSYST-BNAME").text = username
        self.session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = password
        self.session.findById("wnd[0]").sendVKey(0)

        assert self.get_status_bar() != "E", "请检查账号密码是否正常"

        # 处理不同的弹窗类型
        # 判断登录后是否有弹窗
        try:
            wnd1 = self.session.findById("wnd[1]")
        except:
            wnd1 = None

        # 1. 是否单点登录弹窗
        if wnd1 and (wnd1.Text == "多次登录许可信息" or wnd1.Text == "多次登录许可证信息"):
            login_after_path = after_login_dict.get(after_login)
            self.session.findById(login_after_path).Select()
            self.session.findById("wnd[1]/tbar[0]/btn[0]").Press()

        # 判断登录后是否有弹窗
        try:
            wnd1 = self.session.findById("wnd[1]")
        except:
            wnd1 = None

        # 2. 密码错误弹窗
        if wnd1 and wnd1.Text == "信息":
            info_confrim = self.TopWnd.findById("tbar[0]/btn[0]")
            info_confrim.Press()

    @get_object_wrap
    def logout(self, flag):
        """退出登录
        :flag: 是否关闭客户端
        """
        if self.connection:
            self.connection.CloseConnection()
        if flag:
            p = subprocess.Popen("taskkill /im saplogon.exe /f")
            p.wait(3)

    @get_object_wrap
    def open_transaction(self, code):
        """打开事务码
        :param code: str, 待操作的事务码
        """
        self.session.findById("/app/con[0]/ses[0]/wnd[0]/tbar[0]/okcd").text = code
        self.session.findById("/app/con[0]/ses[0]/wnd[0]/tbar[0]/btn[0]").Press()

    def launch_app(self):
        """利用注册表打开 SAP 软件"""
        sub_key = r"SOFTWARE\WOW6432Node\SAP\SAP Shared"
        key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, sub_key)
        sap_sys_dir, _ = winreg.QueryValueEx(key, "SAPsysdir")
        sap_path = os.path.join(sap_sys_dir, "saplogon.exe")
        p = subprocess.Popen(sap_path, creationflags=0x01000000)
        while True:
            try:
                self.SapGui = win32com.client.GetObject("SAPGUI").GetScriptingEngine
                break
            except:
                time.sleep(1)
        self.wait_sap_gui()

    def get_status_bar(self):
        """获取状态栏的状态"""
        return self.session.findById("wnd[0]/sbar").messageType

    def get_status_bar_text(self):
        """获取状态栏的状态"""
        return self.session.findById("wnd[0]/sbar").Text

    def wait_sap_gui(self):
        """等待SAP的窗口出现"""
        while True:
            hWnd_list = []
            win32gui.EnumWindows(lambda hWnd, param: param.append(hWnd), hWnd_list)
            for hwnd in hWnd_list:
                if (
                    IsWindow(hwnd)
                    and IsWindowEnabled(hwnd)
                    and IsWindowVisible(hwnd)
                    and "SAP Logon" in GetWindowText(hwnd)
                ):
                    time.sleep(1)
                    return

    @get_object_wrap
    def set_checkbox(self, path, select):
        """设置复现框
        :param path:
        :param select: 状态
        """
        if select == "TRUE":
            self.TopWnd.findById(path).selected = True
        if select == "FALSE":
            self.TopWnd.findById(path).selected = False

    @get_object_wrap
    def click(self, path):
        """点击元素"""
        self.TopWnd.findById(path).Press

    @get_object_wrap
    def set_combobox(self, path, value, click_info, select_mode):
        """设置下拉框
        :param path:
        :param value: 下拉框值
        :param click_info: bool, 自动点击对话框
        :param select_mode: str, 选择方式
        """
        entries = self.TopWnd.findById(path).Entries
        key_dict = {}

        if select_mode == "CONTENT":
            for item in entries:
                tmp = {}
                tmp["key"] = item.Key
                tmp["pos"] = item.Pos
                key_dict[item.Value] = tmp
        if select_mode == "INDEX":
            for item in entries:
                tmp = {}
                tmp["key"] = item.Key
                tmp["value"] = item.Value
                key_dict[str(item.Pos)] = tmp

        try:
            # 获取下拉框的每一条内容
            entry = key_dict.get(str(value))
            assert entry is not None, "请检查输入内容"
            self.TopWnd.findById(path).key = entry.get("key")
        except:
            raise ValueError("检查下拉框的状态能否选择选项")

        try:
            info_dialog = self.TopWnd
        except:
            info_dialog = None

        if info_dialog and click_info:
            self.TopWnd.findById("tbar[0]/btn[0]").Press()

    @get_object_wrap
    def get_fext_field_tool_tip(self, path):
        return self.session.findById(path).Tooltip

    @get_object_wrap
    def wait_query(self, timeout, error_deal):
        """等待查询结束
        :param timeout: int, 超时时间
        :param error_deal: enum, 1, 继续执行, 2, 报错
        """
        count = 0
        while self.session.Busy:
            time.sleep(2)
            count += 2
            if timeout and count >= int(timeout):
                if error_deal == "1":
                    return False
                if error_deal == "2":
                    raise TimeoutError("查询超时")

        if self.session.findById("wnd[0]/sbar").messageType == "E":
            raise ValueError("查询异常, 请查询输入内容")
        return True

    @get_object_wrap
    def multi_input(self, sap_path, guitab, values):
        guitab_dict = {
            "选择单值": "wnd[1]/usr/tabsTAB_STRIP/tabpSIVA",
            "排除单值": "wnd[1]/usr/tabsTAB_STRIP/tabpNOSV",
        }
        guitab = guitab_dict.get(guitab)

        self.TopWnd.findById(sap_path).press()
        self.TopWnd.findById(guitab).select()
        self.TopWnd.findById("wnd[1]/tbar[0]/btn[16]").press()
        # xbot.win32.clipboard.clear()
        # xbot.win32.clipboard.set_text("\r\n".join(values))
        self.TopWnd.findById("wnd[1]/tbar[0]/btn[24]").press()
        self.TopWnd.findById("wnd[1]/tbar[0]/btn[0]").press()
        self.TopWnd.findById("wnd[1]/tbar[0]/btn[8]").press()

    def get_object(self):
        try:
            self.SapGui = win32com.client.GetObject("SAPGUI")
            self.application = self.SapGui.GetScriptingEngine
            self.connection = self.application.Children(0)
            self.session = self.connection.Children(0)
        except Exception as e:
            # print(e)
            pass

    @property
    def TopWnd(self):
        """获取顶层的窗口"""
        wnd_count = self.session.Children.Count
        assert wnd_count >= 1, "请检查是打开 SAP 软件"
        return self.session.Children(wnd_count - 1)

    @get_object_wrap
    def get_table_data(self, ele_path, use_titles=True):
        table_ele = self.TopWnd.findById(ele_path)

        # Grid or Table
        table_type = table_ele.Type

        if table_type == "GuiShell":
            data = self.get_grid_shell_data(table_ele, use_titles)

        if table_type == "GuiTableControl":
            # print(111111)
            data = self.get_table_shell_data(table_ele, ele_path, use_titles)
        return data

    def get_active_session(self):

        # 目前是通过计算得到的session, 其他方式激活窗口, 通过 application.ActiveSesiion
        conn_cout = self.application.Connections.Count
        latest_conn = self.application.Connections(conn_cout - 1)
        session_count = latest_conn.Sessions.Count
        active_session = latest_conn.Sessions(session_count - 1)
        return active_session

    def end_transaction(self):
        self.session.EndTransaction()

    def select_all(self, path):
        self.session.findById(path).SelectAll()

    def set_input(self, path, value):
        self.session.findById(path).text = value
