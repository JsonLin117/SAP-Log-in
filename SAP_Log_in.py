import win32com.client
import sys
import subprocess
import time
import os


class ReadTXT():
    def __init__(self, filepath):
        self.filepath = filepath

    def read_credentials(self):
        try:
            with open(self.filepath, 'r') as file:
                lines = file.readlines()
                username = lines[0].strip()  # first line is username
                password = lines[1].strip()  # second line is password
                return username, password
        except Exception as e:
            print(f"读取文件时发生错误: {e}")
            return None, None


class SapGui():
    def __init__(self):
        self.path = r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"
        subprocess.Popen(self.path)
        time.sleep(5)

        self.SapGuiAuto = win32com.client.GetObject("SAPGUI")
        if not type(self.SapGuiAuto) == win32com.client.CDispatch:
            return

        application = self.SapGuiAuto.GetScriptingEngine
        self.connection = application.OpenConnection(
            "S4 PRD [Wiwynn S4 PRD]", True)
        self.session = self.connection.Children(0)
        self.session.findById("wnd[0]").maximize()

    def saplogin(self, username, password):

        try:
            self.session.findById("wnd[0]/usr/txtRSYST-MANDT").text = "999"
            self.session.findById("wnd[0]/usr/txtRSYST-BNAME").text = username
            self.session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = password
            self.session.findById("wnd[0]").sendVKey(0)

        except Exception as e:
            print(f"登入SAP時發生錯誤: {e}")
            sys.exit(1)


if __name__ == "__main__":
    downloads_path = os.path.join(os.path.expanduser(
        '~'), 'Downloads')
    credentials_file = os.path.join(downloads_path, "SAP Login.txt")
    reader = ReadTXT(credentials_file)
    username, password = reader.read_credentials()
    sap = SapGui()
    sap.saplogin(username, password)
