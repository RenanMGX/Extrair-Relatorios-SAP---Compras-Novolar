import os
import pandas as pd
import xlwings as xw
from xlwings.main import Sheet #for typing
from dependencies.functions import Functions, _print
import shutil
from typing import List,Dict
from .dependencies.logs import Logs
import traceback
import zipfile

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
        
    def copyTo(self, destiny:str, file_name:str="") -> None:
        if os.path.exists(destiny):
            #os.makedirs(destiny)
        
            if file_name:
                destiny = os.path.join(destiny, file_name)
                if not os.path.exists(os.path.dirname(destiny)):
                    os.makedirs(os.path.dirname(destiny))
                            
            for _ in range(2):
                try:
                    shutil.copy2(self.file_path, destiny)
                    _print(f"{self.file_path} foi copiado para {destiny}")
                    return
                except shutil.Error:
                    os.unlink(destiny)
                    continue
        else:
            raise FileNotFoundError(f"pasta destino não foi encontrada '{destiny}'")

class FilesManipulation:
    @property
    def path_base(self) -> str:
        return self.__path_base
    
    @property
    def files(self) -> List[str]:
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
            if os.path.basename(file).startswith('~$'):
                continue
            if (file.endswith('.xlsx')) or (file.endswith('.xls')):
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
    
    def tratar_arquivos_me3n(self):
        download_path:str = self.__path_base
        if not os.path.exists(download_path):
            raise FileNotFoundError(f"Arquivo '{download_path}' não existe!")
        
        def separar_documentos(df:pd.DataFrame) -> pd.DataFrame:
            documentos: Dict[str, List[pd.Series]] = {}
            ultimo_documento = ""
            for row, value in df.iterrows():
                if value['Item'] == value['Item']:
                    if "Documento de compras" in value['Item']:
                        ultimo_documento = value['Item']
                        documentos[value['Item']] = []
                    else:
                        documentos[ultimo_documento].append(value)
            return tratar_documentos(documentos)
        
        def tratar_documentos(documents:Dict[str, List[pd.Series]]) -> pd.DataFrame:
            result:list = []
            for key, values in documents.items():
                values:List[pd.Series]
                
                result.append(
                    {
                        "Item":key,
                        "Tipo doc.compras": values[0]['Tipo doc.compras'],
                        "Ctg.doc.compras": values[0]['Ctg.doc.compras'],
                        "Grupo de compradores": values[0]['Grupo de compradores'],
                        "Data do documento" : values[0]['Data do documento'],
                        "Fornecedor/centro fornecedor": values[0]['Fornecedor/centro fornecedor'],
                        "Material": values[0]['Material'],
                        "Texto breve": values[0]['Texto breve'],
                        "Início per.validade": values[0]['Início per.validade'],
                        "Fim da validade": values[0]['Fim da validade'],
                        "Grupo de mercadorias": values[0]['Grupo de mercadorias'],
                        "Ctg.class.cont.": values[0]['Ctg.class.cont.'],
                        "Centro": values[0]['Centro'],
                        "Qtd.do pedido": sum([float(valor['Qtd.do pedido']) for valor in values]),
                        "UM pedido": values[0]['UM pedido'],
                        "Qtd.na UnidGestEstoq": sum([float(valor['Qtd.na UnidGestEstoq']) for valor in values]),
                        "Preço líquido": sum([float(valor['Preço líquido']) for valor in values]),
                        "Quantidade prevista": sum([float(valor['Quantidade prevista']) for valor in values]),
                        "Qtd.prev.pendente": sum([float(valor['Qtd.prev.pendente']) for valor in values]),
                        "Valor pendente": sum([float(valor['Valor pendente']) for valor in values]),
                    }
                )
            
            df = pd.DataFrame(result)
            return df            
                 
        for file in os.listdir(download_path):
            file = os.path.join(download_path, file)
            
            if os.path.isfile(file):
                try:
                    if file.endswith('.xlsx'):
                        df:pd.DataFrame = pd.read_excel(file, dtype=str, engine='openpyxl')
                        df['Data do documento'] = pd.to_datetime(df['Data do documento'], format='%Y-%m-%d %H:%M:%S')
                        df['Início per.validade'] = pd.to_datetime(df['Início per.validade'], format='%Y-%m-%d %H:%M:%S')
                        df['Fim da validade'] = pd.to_datetime(df['Fim da validade'], format='%Y-%m-%d %H:%M:%S')
                        
                        df = separar_documentos(df)
                        
                        df.to_excel(file, index=False)
                
                except zipfile.BadZipFile:
                    continue                
                except Exception:
                    Logs(name=str(self.__class__.__name__)).register(status='Error', description='erro ao tratar arquivos me3n', exception=traceback.format_exc())
        
        return self          
                    
                    
                    
                    


if __name__ == "__main__":
    pass