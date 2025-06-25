import os
from pywinauto.application import Application
import json
import pandas as pd
from datetime import datetime
from dateutil.relativedelta import relativedelta
from typing import Literal

#Procurar Arquivos ou Pastas
def click_butto_per_windows(*, window_name:str, button_name:str, raise_exception:bool=True) -> bool:
    try:
        app = Application(backend="uia").connect(title_re=f".*{window_name}.*")
        janela = app.window(title_re=f".*{window_name}.*")
        
        botao = janela.child_window(title=button_name, control_type="Button")
        
        botao.invoke()
        return True
    except Exception as err:
        if raise_exception:
            raise err
        return False
    
class RegistroNFe:
    def __init__(self) -> None:
        file_path = os.environ['reg_path']
        
        file_path = os.path.normpath(file_path)
        
        if not file_path.lower().endswith('.json'):
            file_path += ".json"
        if not os.path.exists(os.path.dirname(file_path)):
            os.makedirs(os.path.dirname(file_path))
            
        self.__file_path:str = file_path
        self.load()
    
    def load(self) -> None:
        if os.path.exists(self.__file_path):
            try:
                with open(self.__file_path, 'r', encoding='utf-8')as _file:
                    data = json.load(_file)
                
                self.df = pd.DataFrame(data)    
                #self.df = pd.read_json(self.__file_path)
                if 'chave_acesso' in self.df.columns:
                    self.df['chave_acesso'] = self.df['chave_acesso'].astype(str)
            except:
                self.df = pd.DataFrame()
                self.__save()
        else:
            self.df = pd.DataFrame()
            self.__save()
    
    def __save(self):
        if 'chave_acesso' in self.df.columns:
            self.df['chave_acesso'] = self.df['chave_acesso'].astype(str)
        with open(self.__file_path, 'w', encoding='utf-8') as _file:
            json.dump(self.df.to_dict(orient='records'), _file, indent = 4 , ensure_ascii=False)
        #self.df.to_json(self.__file_path, orient='records', date_format='iso', indent=4, default_handler=str)
    
    def add(self, chave_acesso:str, *, raise_exeption:bool=True) -> None:
        try:
            check_chave:pd.DataFrame = self.df[
                self.df['chave_acesso'] == chave_acesso
            ]
            if not check_chave.empty:
                if raise_exeption:
                    raise ValueError(f"a chave '{chave_acesso}' ja foi cadastrada anteriormente!")
        except KeyError:
            pass
        
        new_line = pd.DataFrame({'chave_acesso': [str(chave_acesso)], 'date': [datetime.now().isoformat()]})
        new_line['chave_acesso'] = new_line['chave_acesso'].astype(str)
        
        self.df = pd.concat([self.df, new_line], ignore_index=True)
        
        self.__save()
        self.load()
        
    def delete(self, chave_acesso:str) -> None:
        self.df = self.df[
            self.df['chave_acesso'] != chave_acesso
        ]
        
        self.__save()
        self.load()   
        
    def exists(self, chave_acesso:str) -> bool:
        try:
            df = self.df[
                self.df['chave_acesso'] == chave_acesso
            ]
        except KeyError:
            return False
        if not df.empty:
            return True
        return False
    
    def clear_per_date(self, value:int, *, per:Literal['years', 'months', 'days']):
        if not 'date' in self.df.columns:
            return False
        
        ranged_date = datetime.now().replace(hour=0, minute=0, second=0,microsecond=0)
        if per == 'years':
            ranged_date = ranged_date - relativedelta(years=value)
        elif per == 'months':
            ranged_date = ranged_date - relativedelta(months=value)
        elif per == 'days':
            ranged_date = ranged_date - relativedelta(days=value)
        
        
        df = self.df
        df['date'] = df['date'].astype(str)

 
        df = df[
            pd.to_datetime(df['date']) > ranged_date
        ]
        
        self.df = df
        
        self.__save()
        self.load()

    
    
if __name__ == "__main__":
    from patrimar_dependencies.sharepointfolder import SharePointFolders
    sharepoint_path = SharePointFolders(r'RPA - Dados\Configs\SAP - anexar_pdf_MIRO\json').value
    if not os.path.exists(sharepoint_path):
        raise FileNotFoundError(f"pasta no sharepoint '{sharepoint_path}' não foi encontrada")
    os.environ['reg_path'] = os.path.join(sharepoint_path, 'registro_chave_acesso.json')
    
    reg = RegistroNFe()
    
    print(reg.exists('31250500830498000543550010000031221290322465'))
    
    print(reg.clear_per_date(2, per='years') )
    print(reg.df)
