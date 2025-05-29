from Entities.anexar_pdf import AnexarPDF
from Entities.dependencies.functions import P
from Entities.dependencies.logs import Logs
import Entities
from datetime import datetime
import sys
from typing import Dict, List
from Entities.dependencies.arguments import Arguments
import Entities.exceptions
import traceback

class Execute:
    date:datetime = datetime.now()
    
    @staticmethod
    def start():
        try:
            bot = AnexarPDF()
            
            bot.extrair_pdf_vtin_mde(date=Execute.date)
            bot.fechar_sap()
            
            for pdf in bot._listar_arquivos():
                pdf:Dict[str,str]
                for _ in range(5):
                    result = bot.anexar_pdf_miro(chave_acesso=pdf.get('chave_de_acesso'), caminho_arquivo=pdf.get('endereço'))
                    bot.fechar_sap()
                    if result:
                        break
                    
                    
            
            Logs().register(status='Concluido', description=f"Trabalho concluido com sucesso")                
        except Entities.exceptions.NoDocuments:
            print(P("Execução concluida sem decumentos para esta data", color='cyan'))
            Logs().register(status='Concluido', description="Execução concluida sem decumentos para esta data")
        except Exception as error:
            print(P((type(error), str(error)), color='red'))
            Logs().register(status='Error', description=str(error), exception=traceback.format_exc())
        finally:
            bot.fechar_sap()

def start_date(args):
    if isinstance(args, list):
        date = args[0]
    elif isinstance(args, str):
        date = args
    else:
        raise Exception("Tipo inconpativel!")
    
    Execute.date = datetime.strptime(date, "%d.%m.%Y") #"02.09.2021"
    Execute.start()
    Logs().register(status='Concluido', description=f"Trabalho com a data {date} concluido com sucesso")
    

if __name__ == "__main__":
    Arguments({
        "start" : Execute.start,
        "start_date": start_date
    })
        
