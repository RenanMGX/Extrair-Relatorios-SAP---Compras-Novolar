from .logs import Logs
import win32com.client
from datetime import datetime
import psutil
import subprocess
from time import sleep
import traceback
from typing import Literal

def _print(*args, end="\n"):
    if not end.endswith("\n"):
        end += "\n"
    value = ""
    for arg in args:
        value += f"{arg} " 
    
    print(datetime.now().strftime(f"[%d/%m/%Y - %H:%M:%S] - {value}"), end=end)


class SAPManipulation():
    @property
    def ambiente(self) -> str:
        return self.__ambiente
    
    @property
    def session(self) -> win32com.client.CDispatch:
        return self.__session
    
    @property
    def log(self) -> Logs:
        return Logs(name=self.__class__.__name__)
    
    @property
    def using_active_conection(self) -> bool:
        return self.__using_active_conection
    
    def __init__(self, *, user:str="", password:str="", ambiente:str="", using_active_conection:Literal[False, True]=False) -> None:
        self.__using_active_conection = using_active_conection
        self.__user:str = user
        self.__password:str = password
        self.__ambiente:str = ambiente
         
    #decorador
    @staticmethod
    def start_SAP(f):
        def wrap(self, *args, **kwargs):
            try:
                self.session
            except AttributeError:
                self.__conectar_sap()
            try:
                result =  f(self, *args, **kwargs)
            finally:
                sleep(5)
                try:
                    if kwargs['fechar_sap_no_final']:
                        self.fechar_sap()
                except:
                    pass
            return result
                #raise Exception("o sap precisa ser conectado primeiro!")
        return wrap
    
    def __conectar_sap(self) -> None:
        self.__session: win32com.client.CDispatch
        if not self.using_active_conection:
            try:
                if not self._verificar_sap_aberto():
                    subprocess.Popen(r"C:\Program Files (x86)\SAP\FrontEnd\SapGui\saplogon.exe")
                    sleep(5)
                
                SapGuiAuto: win32com.client.CDispatch = win32com.client.GetObject("SAPGUI")# type: ignore
                application: win32com.client.CDispatch = SapGuiAuto.GetScriptingEngine# type: ignore
                connection = application.OpenConnection(self.__ambiente, True) # type: ignore
                self.__session = connection.Children(0)# type: ignore
                
                
                self.session.findById("wnd[0]/usr/txtRSYST-BNAME").text = self.__user # Usuario
                self.session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = self.__password # Senha
                self.session.findById("wnd[0]").sendVKey(0)
                
                return 
            except Exception as error:
                if "connection = application.OpenConnection(self.__ambiente, True)" in traceback.format_exc():
                    raise Exception("SAP está fechado!")
                else:
                    self.log.register(status='Error')
                    raise ConnectionError(f"não foi possivel se conectar ao SAP motivo: {type(error).__class__} -> {error}")
        else:
            try:
                if not self._verificar_sap_aberto():
                    raise Exception("SAP está fechado!")
                
                self.SapGuiAuto: win32com.client.CDispatch = win32com.client.GetObject("SAPGUI")
                self.application: win32com.client.CDispatch = self.SapGuiAuto.GetScriptingEngine
                self.connection: win32com.client.CDispatch = self.application.Children(0)
                self.__session = self.connection.Children(0)
                
            except Exception as error:
                if "self.connection: win32com.client.CDispatch = self.application.Children(0)" in traceback.format_exc():
                    raise Exception("SAP está fechado!")
                elif "SAP está fechado!" in traceback.format_exc():
                    raise Exception("SAP está fechado!")
                else:
                    self.log.register(status='Error')

    #@usar_sap
    def fechar_sap(self):
        _print("fechando SAP!")
        try:
            sleep(1)
            self.session.findById("wnd[0]").close()
            sleep(1)
            try:
                self.session.findById('wnd[1]/usr/btnSPOP-OPTION1').press()
            except:
                self.session.findById('wnd[2]/usr/btnSPOP-OPTION1').press()
        except Exception as error:
            if not "(-2147417848, 'O objeto chamado foi desconectado de seus clientes.', None, None)" in str(error):
                _print(f"não foi possivel fechar o SAP {type(error)} | {error}")

    @start_SAP
    def _listar(self, campo):
        cont = 0
        for child_object in self.session.findById(campo).Children:
            _print(f"{cont}: ","ID:", child_object.Id, "| Type:", child_object.Type, "| Text:", child_object.Text)
            cont += 1

    def _verificar_sap_aberto(self) -> bool:
        for process in psutil.process_iter(['name']):
            if "saplogon" in process.name().lower():
                return True
        return False    
    
    @start_SAP
    def _teste(self):
        _print("testado")
  
if __name__ == "__main__":
    pass
    #crd = Credential("SAP_QAS").load()
    
    #bot = SAPManipulation(user=crd['user'], password=crd['password'], ambiente="S4Q")
    #bot.conectar_sap()
    #bot._teste()
    
    #import pdb;pdb.set_trace()
    #bot.fechar_sap()
