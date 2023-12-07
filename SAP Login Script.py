# SAP Login Script
import win32com.client
import subprocess
import time
import os
import tkinter as tk
from tkinter import filedialog, messagebox


def get_sap_client_session(my_username, my_password, my_sap_sid="SB2", my_sap_sysname='"S4 PRD [Wiwynn S4 PRD]"', my_sap_client="999"):
    try:
        sap_gui = win32com.client.GetObject("SAPGUI")
        application = sap_gui.GetScriptingEngine
        connection = application.Children(0)

    except:
        # Can not find the GUI window, so utilize sapshcut to open
        subprocess.check_call('start "C:/Program Files (x86)/SAP/FrontEnd/SAPgui/" sapshcut.exe -system=' +
                              my_sap_sid + ' -sysname=' + my_sap_sysname + ' -user=' + my_username +
                              ' -pw=' + my_password + ' -client=' + my_sap_client, shell=True)

        # Make sure the login window is open
        time.sleep(10)
        sap_gui = win32com.client.GetObject("SAPGUI")
        if not type(sap_gui) == win32com.client.CDispatch:
            return

        application = sap_gui.GetScriptingEngine
        if not type(application) == win32com.client.CDispatch:
            sap_gui = None
            return None

        connection = application.Children(0)
        if not type(connection) == win32com.client.CDispatch:
            application = None
            sap_gui = None
            return None

        session = connection.Children(0)
        if not type(session) == win32com.client.CDispatch:
            connection = None
            application = None
            sap_gui = None

    # Get session
    try:
        session = connection.Children(0)
        # If password and account is not corrrect
        if ("Name or password is incorrect" in session.findById("wnd[0]/sbar").Text):
            print('Name or password is incorrect')
            return None

        # Handle duplicated login
        if (session.children.count > 1):
            try:
                session.findById("wnd[1]/usr/radMULTI_LOGON_OPT2").select()
                session.findById("wnd[1]/usr/radMULTI_LOGON_OPT2").setFocus()
                session.findById("wnd[0]").sendVKey(0)
            except:
                # print('Can not login to SAP successfully.')
                # session = None
                pass
        if (session.children.count > 1):
            try:
                session.findById("wnd[1]").close()
            except:
                session = None
        return session
    except:
        print('Cannot get SAP GUI client sucessfully')
        return None


def choose_file():
    # 隐藏主窗口
    root = tk.Tk()
    root.withdraw()

    file_path = filedialog.askopenfilename(filetypes=[("Text files", "*.txt")])

    return file_path


def read_credentials(file_path):
    # Read credentials from file
    try:
        with open(file_path, "r") as f:
            credentials = f.readlines()
            username = credentials[0].strip()
            password = credentials[1].strip()
    except Exception as e:
        print("Error reading credentials: ", e)
        return None, None
    return username, password


def main():
    credentials_file = choose_file()
    if credentials_file:
        username, password = read_credentials(credentials_file)
    get_sap_client_session(username, password)


if __name__ == "__main__":
    main()
