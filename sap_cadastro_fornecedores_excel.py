

import click
import win32com.client
import subprocess
import sys
import time
from tkinter import *
from tkinter import messagebox
import pandas as pd

class SapGui():
    def __init__(self):
        self.path = r"C:\Program Files (x86)\SAP\FrontEnd\SapGui\saplogon.exe"    #Caminho do executor SAP
        subprocess.Popen(self.path)
        time.sleep(3)

        self.SapGuiAuto = win32com.client.GetObject("SAPGUI")
        if not type(self.SapGuiAuto)== win32com.client.CDispatch:
            return

        application = self.SapGuiAuto.GetScriptingEngine

        self.connection = application.OpenConnection("Curso Grandes Projetos", True)
        time.sleep(3)
        self.session = self.connection.Children(0)


    def sapLogin(self):

        try:
            self.session.findById("wnd[0]/usr/txtRSYST-MANDT").text = "800"     # mandante SAP
            self.session.findById("wnd[0]/usr/txtRSYST-BNAME").text = "USUARIO" #usuário SAP
            self.session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = "senha"   #senha SAP
            self.session.findById("wnd[0]/usr/txtRSYST-LANGU").text = "PT"
            self.session.findById("wnd[0]").sendVKey(0)

        except:
            print(sys.exc_info()[0])

        time.sleep(3)
        self.register_supplier()

    def register_supplier(self):

        data = pd.read_excel(r"C:\Users\Usuario\Application Data\Desktop\fornec.xlsx", sheet_name="fornec").astype(str) # caminho do arquivo xlsx, nome da aba onde os dados estão
        data.columns = data.columns.str.replace(' ','_')
        time.sleep(5)

        self.session.findById("wnd[0]").maximize
        self.session.findById("wnd[0]/tbar[0]/okcd").text = "FK01"
        self.session.findById("wnd[0]").sendVKey(0)
        self.session.findById("wnd[0]/usr/ctxtRF02K-BUKRS").text = "0001"
        self.session.findById("wnd[0]/usr/ctxtRF02K-KTOKK").text = "0001"
        self.session.findById("wnd[0]/usr/ctxtRF02K-KTOKK").setFocus
        self.session.findById("wnd[0]/usr/ctxtRF02K-KTOKK").caretPosition = 4

        for index, row in data.iterrows():
            print(index, row.NOME)
            #script abaixo: via script SAP
            self.session.findById("wnd[0]").sendVKey(0)
            self.session.findById("wnd[0]/usr/subADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/cmbSZA1_D0100-TITLE_MEDI").key = "Empresa"
            self.session.findById("wnd[0]/usr/subADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/txtADDR1_DATA-NAME1").text = row.NOME
            self.session.findById("wnd[0]/usr/subADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/txtADDR1_DATA-SORT1").text = row.TERMO_DE_PESQUISA
            self.session.findById("wnd[0]/usr/subADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/txtADDR1_DATA-SORT2").text = row.TERMO_DE_PESQUISA
            self.session.findById("wnd[0]/usr/subADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/txtADDR1_DATA-STREET").text = row.RUA
            self.session.findById("wnd[0]/usr/subADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/txtADDR1_DATA-HOUSE_NUM1").text = row.NUMERO
            self.session.findById("wnd[0]/usr/subADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/txtADDR1_DATA-POST_CODE1").text = row.COD_POSTA
            self.session.findById("wnd[0]/usr/subADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/ctxtADDR1_DATA-COUNTRY").text = row.PAIS
            self.session.findById("wnd[0]/usr/subADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/ctxtADDR1_DATA-REGION").text = row.REGIAO
            self.session.findById("wnd[0]/usr/subADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/txtADDR1_DATA-PO_BOX").text = row.COD_POSTAL
            self.session.findById("wnd[0]/usr/subADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/txtADDR1_DATA-POST_CODE2").text = row.COD_POSTAL
            self.session.findById("wnd[0]/usr/subADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/txtADDR1_DATA-POST_CODE2").setFocus
            self.session.findById("wnd[0]/usr/subADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/txtADDR1_DATA-POST_CODE2").caretPosition = 9
            self.session.findById("wnd[0]/tbar[1]/btn[8]").press()
            self.session.findById("wnd[0]/tbar[1]/btn[8]").press()
            self.session.findById("wnd[0]/tbar[1]/btn[8]").press()
            self.session.findById("wnd[0]/tbar[1]/btn[8]").press()
            self.session.findById("wnd[0]/usr/ctxtLFB1-AKONT").text = row.CTA_CONCIL
            self.session.findById("wnd[0]/usr/ctxtLFB1-FDGRV").text = row.FDGRV
            self.session.findById("wnd[0]/usr/ctxtLFB1-FDGRV").setFocus
            self.session.findById("wnd[0]/usr/ctxtLFB1-FDGRV").caretPosition = 2
            self.session.findById("wnd[0]").sendVKey(0)
            self.session.findById("wnd[0]/tbar[0]/btn[11]").press()
            self.session.findById("wnd[1]/tbar[0]/btn[0]").press()

        messagebox.showinfo("showinfo", "Cadastro realizado com sucesso!")

if __name__ == '__main__':
    window = Tk()
    window.geometry("200x50")
    botao = Button(window, text = "Login SAP", command= lambda :SapGui().sapLogin())
    botao.pack()
    mainloop()

