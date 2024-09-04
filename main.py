
from Entities.dependencies.functions import _print
import sys


class Execute:
    @property
    def lista_obras_path(self) -> str:
        return f"C:\\Users\\{getuser()}\\PATRIMAR ENGENHARIA S A\\RPA - Documentos\\RPA - Dados\\RPA - Suprimentos Novolar\\obras_novolar.xlsx"
    
    @property
    def path_destiny_zmm019_compras(self) -> str:
        return self.__path_destiny_zmm019_compras
    @path_destiny_zmm019_compras.setter
    def path_destiny_zmm019_compras(self, value:str) -> None:
        if not isinstance(value, str):
            raise TypeError("apenas strings")
        self.__path_destiny_zmm019_compras = value
        
    @property
    def path_destiny_contratos(self) -> str:
        return self.__path_destiny_contratos
    @path_destiny_contratos.setter
    def path_destiny_contratos(self, value:str) -> None:
        if not isinstance(value, str):
            raise TypeError("apenas strings")
        self.__path_destiny_contratos = value
    
    def __init__(self) -> None:
        self.__extrair_relat:ExtrairRelatorio = ExtrairRelatorio(choicer='SAP_PRD')
        self.__lista_obras:list = pd.read_excel(self.lista_obras_path)['Obras Novolar'].unique().tolist()
        self.__log:Logs = Logs()
        self.__path_destiny_zmm019_compras:str = r'\\server008\g\ARQ_PATRIMAR\Setores\dpt_tecnico\Suprimentos_Novolar\RELATÓRIOS\VOLUME DE COMPRAS'
        self.__path_destiny_contratos:str = r'\\server008\g\ARQ_PATRIMAR\Setores\dpt_tecnico\Suprimentos_Novolar\RELATÓRIOS\CONTRATOS'
        
    def start(self) -> None:
        agora = datetime.now()
        self.start_zmm019()        
        
        # try:
        #     self.__extrair_relat.extrair_rel_me5a(self.__lista_obras)
        #     FilesManipulation(self.__extrair_relat.download_path_me5a).unify().copyTo(self.path_destiny_contratos, file_name=self.__extrair_relat.file_name_contratos)
        # except Exception as error:
        #     _print(f"Erro: {error}")
        #     self.__log.register(status='Error', description="Erro ao executar me5a")
        
        self.__extrair_relat.finalizar(fechar_sap_no_final=True)
        self.__log.register(status='Report', description=f"tempo de execução do scrip '{datetime.now() - agora}'")
        
    def start_zmm019(self):
        try:
            self.__extrair_relat.extrair_rel_zmm019(empreendimentos=self.__lista_obras)
            FilesManipulation(self.__extrair_relat.download_path_zmm019).unify().copyTo(self.path_destiny_zmm019_compras, file_name=self.__extrair_relat.file_name_zmm019_compras)
        except Exception as error:
            _print(f"Erro: {error}")
            self.__log.register(status='Error', description="erro ao executar zmm019")
        
if __name__ == "__main__":
    argv:list = sys.argv
    if len(argv) > 1:
        from Entities.extrair_relatorios import ExtrairRelatorio
        from Entities.files_manipulation import FilesManipulation
        import pandas as pd
        from getpass import getuser
        from Entities.dependencies.logs import Logs
        from datetime import datetime
        
        if argv[1] == "start":
            Execute().start()
        elif argv[1] == "zmm019":
            Execute().start_zmm019()
        else:
            _print(f"Argumento invalido: '{argv[1]}'")
    else:
        _print("obrigatorio uso de argumento")
        _print("[start, zmm019]")
        
        