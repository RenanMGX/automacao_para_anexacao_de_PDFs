from datetime import datetime
from Entities.dependencies.config import Config
from Entities.dependencies.credenciais import Credential
from Entities.dependencies.sap import SAPManipulation
from Entities import exceptions
import pygetwindow
import pyautogui
pyautogui.FAILSAFE = False
from pygetwindow._pygetwindow_win import Win32Window
import os
import re
from Entities.dependencies.logs import Logs
import traceback
from typing import List, Dict
import shutil
from time import sleep
from Entities.dependencies.functions import P
import sys


class AnexarPDF(SAPManipulation):
    download_path:str = os.path.join(os.getcwd(), 'downloads/pdf')
    if not os.path.exists(download_path):
        os.makedirs(download_path)
    
    def __init__(self) -> None:
        credential_seleted = credential_seleted['crd'] if (credential_seleted:=Config()['credential']) else "None"
        crd:dict = Credential(credential_seleted).load()
        super().__init__(user=crd.get('user'), password=crd.get('password'), ambiente=crd.get('ambiente'))
        
    
    @SAPManipulation.start_SAP    
    def extrair_pdf_vtin_mde(self, *, date:datetime, fechar_sap_no_final:bool=False):
        self.session.findById("wnd[0]/tbar[0]/okcd").text = "/n/VTIN/MDE"
        self.session.findById("wnd[0]").sendVKey(0)
        
        #02.09.2021
        self.session.findById("wnd[0]/usr/ctxtS_CREDAT-LOW").text = date.strftime('%d.%m.%Y')# <---------------------------
        self.session.findById("wnd[0]/usr/ctxtS_CREDAT-HIGH").text = date.strftime('%d.%m.%Y')#<---------------------------
        
        self.session.findById("wnd[0]/usr/ctxtS_BUKRS-LOW").text = "*"
        
        self.session.findById("wnd[0]/usr/ctxtP_VARI").text = "/SEIDOR"
        
        self.session.findById("wnd[0]/tbar[1]/btn[8]").press()
        
        titl:str = self.session.findById("wnd[0]/titl").text
        if (num_documento:=re.search(r'[\d+]+(?= documento\(s\) selecionado\(s\))', titl)):
            num_documento = int(num_documento.group())
        else:
            raise exceptions.VerificQuantDocumentsError(f"não foi possivel identificar quandos documentos tem disponivel\n ")

        if num_documento <= 0:
            raise exceptions.NoDocuments("Sem documentos para anexar pdf")
        
        self.session.findById("wnd[0]/usr/shell/shellcont[0]/shell").setCurrentCell(-1,"BELNR")
        self.session.findById("wnd[0]/usr/shell/shellcont[0]/shell").selectColumn("BELNR")
        self.session.findById("wnd[0]/usr/shell/shellcont[0]/shell").contextMenu()
        self.session.findById("wnd[0]/usr/shell/shellcont[0]/shell").selectContextMenuItem("&FILTER")
        
        self.session.findById("wnd[1]/tbar[0]/btn[2]").press()
        self.session.findById("wnd[2]/usr/cntlOPTION_CONTAINER/shellcont/shell").setCurrentCell(5,"TEXT")
        self.session.findById("wnd[2]/usr/cntlOPTION_CONTAINER/shellcont/shell").selectedRows = "5"
        self.session.findById("wnd[2]/usr/cntlOPTION_CONTAINER/shellcont/shell").doubleClickCurrentCell()
        
        self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
        
        downloads:list = []
        contador:int = 0
        
        #import pdb; pdb.set_trace()
        
        self.limpar_download_path() 
        while True:
            try:
                self.session.findById("wnd[0]/usr/shell/shellcont[0]/shell").currentCellRow = contador
                self.session.findById("wnd[0]/usr/shell/shellcont[0]/shell").selectedRows = str(contador)
                
                janela:Win32Window = self.get_window("Monitor DF-e Estaduais")
                janela.minimize()
                janela.restore()
                janela.moveTo(0,0)
                pyautogui.hotkey('ctrl' , 'f2')
                janela.minimize()
                while True:
                    #print(P('esperando a tela Procurar Arquivos ou Pastas'))
                    if 'Procurar Arquivos ou Pastas' in pygetwindow.getAllTitles():
                        janela_procurar = self.get_window('Procurar Arquivos ou Pastas')
                        janela.minimize()
                        janela_procurar.minimize()
                        janela_procurar.restore()
                        janela_procurar.moveTo(0,0)
                        janela_procurar.resize(300,300)
                        pyautogui.press('enter')
                        break
                    sleep(1)
                janela.restore()
                tbar:str = self.session.findById("wnd[0]/sbar").text
                if 'Download:' in tbar:
                    if (msg:=re.search(r'C:\\[\d\D]+', tbar)):
                        #downloads.append(msg.group())
                        shutil.move(msg.group(), self.download_path)
                    else:
                        print(f"não foi possivel baixar o pdf - {msg}")
                        Logs().register(status='Report', description="erro ao fazer download do pdf", exception=msg)
                else:
                    print(f"não foi possivel baixar o pd - {tbar}")
                    Logs().register(status='Report', description="erro ao fazer download do pdf", exception=tbar)
                        
            except Exception as error:
                Logs().register(status='Error', description=str(error), exception=traceback.format_exc())
                sys.exit()
            contador += 1
        
        #import pdb; pdb.set_trace()
        # self.limpar_download_path() 
        # for download in downloads:
            # shutil.move(download, self.download_path)
            #os.unlink(download)

        #re.search(r'(?<=NFe)[\d]+(?=.pdf)', msg.group())
    
    @SAPManipulation.start_SAP            
    def anexar_pdf_miro(self, *, chave_acesso:str|None, caminho_arquivo:str|None):
        if not (chave_acesso and caminho_arquivo):
            print(P(f'{chave_acesso=} ou {caminho_arquivo=} não pode estar vazio'))
            Logs().register(status='Report', description=f'{chave_acesso=} ou {caminho_arquivo=} não pode estar vazio')
            return False
        
        print(P(os.path.basename(caminho_arquivo)), end=" ")
        try:
            self.session.findById("wnd[0]/tbar[0]/okcd").text = "/n/vtin/inb"
            self.session.findById("wnd[0]").sendVKey(0)
            self.session.findById("wnd[0]/usr/ctxtW_CHAVE_ACESSO").text = chave_acesso
            self.session.findById("wnd[0]").sendVKey(0)
            
            if (sbar:=self.session.findById("wnd[0]/sbar").text):
                raise exceptions.ValidarChaveAcessoError(sbar)
            
            try:
                self.session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
            except:
                pass
            
            try:
                if not self.session.findById("wnd[0]/usr/shell/shellcont[1]/shell").GetCellValue(1,"LINK_DOC"):
                    raise exceptions.MiroNotFoundError("o campo da MIRO está vazio")
            except:
                raise exceptions.MiroNotFoundError("o campo da MIRO está vazio")
            
            self.session.findById("wnd[0]/usr/shell/shellcont[1]/shell").setCurrentCell(1,"LINK_DOC")
            self.session.findById("wnd[0]/usr/shell/shellcont[1]/shell").clickCurrentCell()
            
            self.session.findById("wnd[0]/titl/shellcont/shell").pressButton("%GOS_TOOLBOX")
            janela = self.get_window("Fatura recebida")
            janela.minimize()
            janela.restore()
            self.session.findById("wnd[0]/shellcont/shell").pressContextButton("CREATE_ATTA")
            self.session.findById("wnd[0]/shellcont/shell").selectContextMenuItem("PCATTA_CREA")
            
            self.session.findById("wnd[1]/usr/ctxtDY_PATH").text = os.path.dirname(caminho_arquivo)
            self.session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = os.path.basename(caminho_arquivo)
            
            self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
            
            self.session.findById("wnd[0]/shellcont").close()
            
            print("finalizado!")
        
            return True
        except Exception as error:
            print("error!")
            print(P((chave_acesso, str(error)),title='ERROR',color='red'))
            Logs().register(status='Report', description=f"{chave_acesso=} - {str(error)}", exception=traceback.format_exc())
            return False
        
        
    @staticmethod
    def get_window(target:str) -> Win32Window:
        for _window in pygetwindow.getAllTitles():
            if target in _window:
                return pygetwindow.getWindowsWithTitle(_window)[0]
        raise Exception("nao encontrado")
    
    @staticmethod
    def limpar_download_path():
        for file in os.listdir(AnexarPDF.download_path):
            file = os.path.join(AnexarPDF.download_path, file)
            if os.path.isfile(file):
                try:
                    os.unlink(file)
                except:
                    pass
                
    def _listar_arquivos(self) -> List[Dict[str,str]]:
        lista:List[Dict[str,str]] = []
        for file in os.listdir(self.download_path):
            file = os.path.join(self.download_path, file)
            if (chave_acesso:=re.search(r'(?<=NFe)[\d]+(?=.pdf)', os.path.basename(file))):
                lista.append({
                    "chave_de_acesso":chave_acesso.group(),
                    "endereço": file
                })
            else:
                print(P(f"não foi possivel extrair a cheve de acesso do caminho '{file}'", title='REPORT', color='red'))
                Logs().register(status='Report', description=f"não foi possivel extrair a cheve de acesso do caminho '{file}'")
        return lista
                
if __name__ == "__main__":
    pass
