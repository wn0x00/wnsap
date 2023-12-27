# from wnsap import SapGui

# sap = SapGui()
# import pythoncom

# pythoncom.CoInitialize()

# from win32com.client import DispatchEx, GetActiveObject, GetObject, Dispatch

# # sapgui = Dispatch("Sapgui.Application")
# # app = Dispatch("Sapgui.ScriptingCtrl.1")

# Dispatch("Sapgui.Application", clsctx=pythoncom.CLSCTX_LOCAL_SERVER)

# GetObject("SAPGUI")

# sapgui.OpenConnection("jx", False)

# sap.login("jx", "xxxxxxxxxxxxx", "xxxxxxxxxxxxx")


import win32com.client

# for key in win32com.client.gencache.GetModuleForProgID(
#     "Sapgui.Application"
# ).dict.keys():
#     print(key)


res = win32com.client.gencache.GetModuleForProgID("SAPGUI")
print(res)
