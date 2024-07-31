import os
import pandas as pd
import xlwings as xw
from xlwings.main import Sheet #for typing
from dependencies.functions import Functions, _print
import shutil

class PathNotFound(Exception):
    def __init__(self, *args: object) -> None:
        super().__init__(*args)

class FilesJoined:
    @property
    def file_path(self) -> str:
        return self.__file_path
    
    def __init__(self, file_path:str) -> None:
        if os.path.exists(file_path):
            if file_path.endswith('.xlsx'):
                self.__file_path:str = file_path
            else:
                raise TypeError("Apenas arquivos .xlsx")
        else:
            raise FileNotFoundError("Arquivo não encontrado!")
        
    def copyTo(self,destiny:str):
        if not os.path.exists(os.path.dirname(destiny)):
            os.makedirs(os.path.dirname(destiny))
            
        shutil.copy2(self.file_path, destiny)
        _print(f"{self.file_path} foi copiado para {destiny}")

class FilesManipulation:
    @property
    def path_base(self) -> str:
        return self.__path_base
    
    @property
    def files(self) -> list:
        return [os.path.join(self.path_base, value) for value in os.listdir(self.path_base)]
    
    def __init__(self, path_base:str) -> None:
        if os.path.exists(path_base):
            self.__path_base:str = path_base
        else:
            raise PathNotFound("Arquivo não encontrado")
        
    def unify(self) -> FilesJoined:
        _print("Iniciando Unificação de dados")
        
        df = pd.DataFrame()
        for file in self.files:
            file = os.path.join(self.path_base, file)
            _print(f"copiando dados do arquivo {file}")
            apps = xw.App(visible=False)
            with apps.books.open(file)as wb:
                ws:Sheet = wb.sheets[0]
                df_temp:pd.DataFrame = ws.range('A1').expand().options(pd.DataFrame, index=False).value
                df = pd.concat([df, df_temp], ignore_index=True)

            Functions.fechar_excel(file)
        
        path_destiny:str = os.path.join(self.path_base, "unificado")
        if not os.path.exists(path_destiny):
            os.makedirs(path_destiny)
        
        file_path = os.path.join(path_destiny, "unificado.xlsx")
        _print(f"unificação realizada salvando no caminho {file_path}")
        df.to_excel(file_path, index=False)
        return FilesJoined(file_path)


if __name__ == "__main__":
    pass