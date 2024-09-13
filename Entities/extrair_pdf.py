from datetime import datetime
from Entities.dependencies.config import Config
from Entities.dependencies.credenciais import Credential
from Entities.dependencies.sap import SAPManipulation
import pygetwindow
import pyautogui
from pygetwindow._pygetwindow_win import Win32Window
import os


class ExtrairPDFs(SAPManipulation):
    def __init__(self) -> None:
        credential_seleted = credential_seleted['crd'] if (credential_seleted:=Config()['credential']) else "None"
        crd:dict = Credential(credential_seleted).load()
        super().__init__(user=crd.get('user'), password=crd.get('password'), ambiente=crd.get('ambiente'))
    
    @SAPManipulation.start_SAP    
    def vtin_mde(self, *, date:datetime=datetime.now()):
        
        
        self.session.findById("wnd[0]/tbar[0]/okcd").text = "/n/VTIN/MDE"
        self.session.findById("wnd[0]").sendVKey(0)
        
        #02.09.2021
        self.session.findById("wnd[0]/usr/ctxtS_CREDAT-LOW").text = "02.09.2021"#date.strftime('%d.%m.%Y')
        self.session.findById("wnd[0]/usr/ctxtS_CREDAT-HIGH").text = "02.09.2021"#date.strftime('%d.%m.%Y')
        
        self.session.findById("wnd[0]/usr/ctxtS_BUKRS-LOW").text = "*"
        
        self.session.findById("wnd[0]/tbar[1]/btn[8]").press()
        
        self.session.findById("wnd[0]/usr/shell/shellcont[0]/shell").setCurrentCell(-1,"BELNR")
        self.session.findById("wnd[0]/usr/shell/shellcont[0]/shell").selectColumn("BELNR")
        self.session.findById("wnd[0]/usr/shell/shellcont[0]/shell").contextMenu()
        self.session.findById("wnd[0]/usr/shell/shellcont[0]/shell").selectContextMenuItem("&FILTER")
        
        self.session.findById("wnd[1]/tbar[0]/btn[2]").press()
        self.session.findById("wnd[2]/usr/cntlOPTION_CONTAINER/shellcont/shell").setCurrentCell(5,"TEXT")
        self.session.findById("wnd[2]/usr/cntlOPTION_CONTAINER/shellcont/shell").selectedRows = "5"
        self.session.findById("wnd[2]/usr/cntlOPTION_CONTAINER/shellcont/shell").doubleClickCurrentCell()
        
        self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
        
        import pdb; pdb.set_trace()
        
        download:dict
        contador:int = 0
        while True:
            try:
                self.session.findById("wnd[0]/usr/shell/shellcont[0]/shell").currentCellRow = contador
                self.session.findById("wnd[0]/usr/shell/shellcont[0]/shell").selectedRows = str(contador)
                
                janela:Win32Window = ExtrairPDFs.get_window("Monitor DF-e Estaduais")
                janela.minimize()
                janela.restore()
                janela.moveTo(0,0)
                pyautogui.hotkey('ctrl' , 'f2')
                pyautogui.press('enter')
                
                self.session.findById("wnd[0]/sbar").text
                
            except:
                break
            contador += 1
        
       
        
        # session.findById("wnd[0]").resizeWorkingPane 150,24,false
        # session.findById("wnd[0]/usr/shell/shellcont[0]/shell").currentCellColumn = ""
        # session.findById("wnd[0]/usr/shell/shellcont[0]/shell").selectedRows = "0"
        # self.session.findById("wnd[0]/mbar/menu[0]/menu[4]").select()
        # self.session.findById("wnd[0]/usr/shell/shellcont[0]/shell").currentCellRow = 0
        # self.session.findById("wnd[0]/usr/shell/shellcont[0]/shell").selectedRows = "0"
        # session.findById("wnd[0]/mbar/menu[0]/menu[4]").select        
        
        
        
        self._listar('/wnd[0]')   
    
    @staticmethod
    def get_window(target:str) -> Win32Window:
        for _window in pygetwindow.getAllTitles():
            if target in _window:
                return pygetwindow.getWindowsWithTitle(_window)[0]
        raise Exception("nao encontrado")    