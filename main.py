import os
from patrimar_dependencies.sharepointfolder import SharePointFolders
sharepoint_path = SharePointFolders(r'RPA - Dados\Configs\SAP - anexar_pdf_MIRO\json').value
if not os.path.exists(sharepoint_path):
    raise FileNotFoundError(f"pasta no sharepoint '{sharepoint_path}' não foi encontrada")
os.environ['reg_path'] = os.path.join(sharepoint_path, 'registro_chave_acesso.json')

from Entities.anexar_pdf import AnexarPDF
from patrimar_dependencies.functions import P
import Entities
from datetime import datetime
import sys
from typing import Dict, List
import Entities.exceptions
import traceback
from botcity.maestro import * # type:ignore
from Entities import utils

class ExecuteAPP:    
    @staticmethod
    def start(*, user:str, password:str, ambiente:str, maestro:BotMaestroSDK|None = None, date: datetime, range_dias:int=0):
        
        try:
            bot = AnexarPDF(
                user=user,
                password=password,
                ambiente=ambiente,
                maestro=maestro
            )   
            
            bot.extrair_pdf_vtin_mde(date=date, range_dias=range_dias)
            bot.fechar_sap()
            
            lista_pdfs:list = bot._listar_arquivos()
            
            if not lista_pdfs:
                raise Entities.exceptions.NoDocuments()
            
            for pdf in lista_pdfs:
                pdf:Dict[str,str]
                for _ in range(5):
                    result = bot.anexar_pdf_miro(chave_acesso=pdf.get('chave_de_acesso'), caminho_arquivo=pdf.get('endereço'))
                    #bot.fechar_sap()
                    if result:
                        try:
                            utils.RegistroNFe().add(pdf['chave_de_acesso'])
                        except Exception as err:
                            print(pdf['chave_de_acesso'], type(err), err)
                        break
            
            utils.RegistroNFe().clear_per_date(2, per='years')    
            if not maestro is None: 
                maestro.new_log_entry(
                    activity_label="SAP-anexar_pdf_MIRO",
                    values={
                        "texto": f"Trabalho concluido com sucesso"
                    }
                ) 
                
            print("Automação Concluida!")   
                
                        
        except Entities.exceptions.NoDocuments:
            print(P("Execução concluida sem decumentos para esta data", color='cyan'))
            if not maestro is None: 
                maestro.new_log_entry(
                    activity_label="SAP-anexar_pdf_MIRO",
                    values={
                        "texto": "Execução concluida sem decumentos para esta data"
                    }
                ) 

        except Exception as error:
            raise error
        finally:
            bot.fechar_sap()
            
        

# def start_date(args):
#     if isinstance(args, list):
#         date = args[0]
#     elif isinstance(args, str):
#         date = args
#     else:
#         raise Exception("Tipo inconpativel!")
    
#     Execute.date = datetime.strptime(date, "%d.%m.%Y") #"02.09.2021"
#     Execute.start()
#     Logs().register(status='Concluido', description=f"Trabalho com a data {date} concluido com sucesso")
    

if __name__ == "__main__":
    from patrimar_dependencies.credenciais import Credential
    from patrimar_dependencies.sharepointfolder import SharePointFolders
    
    crd:dict = Credential(
        path_raiz=SharePointFolders(r'RPA - Dados\CRD\.patrimar_rpa\credenciais').value,
        name_file="SAP_PRD_NF"
    ).load()
    
    print(crd)
    
    date = datetime.now()
    ExecuteAPP.start(
        user=crd['user'],
        password=crd['password'],
        ambiente=crd['ambiente'],
        date=date
    )
    
        
