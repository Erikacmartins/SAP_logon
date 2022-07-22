

import click
import win32com.client
import subprocess
import sys
import time
from tkinter import *
from tkinter import messagebox

class SapGui():
    def __init__(self):
        self.path = r"C:\Program Files (x86)\SAP\FrontEnd\SapGui\saplogon.exe"
        subprocess.Popen(self.path)
        time.sleep(3)

        self.SapGuiAuto = win32com.client.GetObject("SAPGUI")
        if not type(self.SapGuiAuto)== win32com.client.CDispatch:
            return

        application = self.SapGuiAuto.GetScriptingEngine

        self.connection = application.OpenConnection("NOME DA CONEXAO", True)
        time.sleep(3)
        self.session = self.connection.Children(0)


    def sapLogin(self):

        try:
            self.session.findById("wnd[0]/usr/txtRSYST-MANDT").text = "300"
            self.session.findById("wnd[0]/usr/txtRSYST-BNAME").text = "USUARIO"
            self.session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = "SENHA"
            self.session.findById("wnd[0]/usr/txtRSYST-LANGU").text = "PT"
            self.session.findById("wnd[0]").sendVKey(0)

        except:
            print(sys.exc_info()[0])
        messagebox.showinfo("showinfo", "Login realizado com sucesso")

if __name__ == '__main__':
    window = Tk()
    window.geometry("200x50")
    botao = Button(window, text = "Login SAP", command= lambda :SapGui().sapLogin())
    botao.pack()
    mainloop()
