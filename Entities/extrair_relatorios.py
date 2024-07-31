from typing import Literal
from dependencies.sap import SAPManipulation
from dependencies.credenciais import Credential
import pygetwindow
import pyautogui
import re
from pygetwindow._pygetwindow_win import Win32Window
import os
import shutil
from datetime import datetime
from dateutil.relativedelta import relativedelta
from typing import List, Dict
from dependencies.logs import Logs
from dependencies.functions import Functions, _print
import traceback

class ExtrairRelatorio(SAPManipulation):
    @property
    def download_path_zmm030(self) -> str:
        return os.path.join(os.getcwd(), r'downloads\zmm030')
    
    @property
    def download_path_zmm019(self) -> str:
        return os.path.join(os.getcwd(), r'downloads\zmm019')
    
    def __init__(self, choicer:Literal['SAP_PRD', 'SAP_QAS', 'SAP_QAS-Renan']) -> None:
        crd:dict = Credential(choicer).load() #type: ignore
        super().__init__(user=crd['user'], password=crd['password'], ambiente=crd['ambiente'])
    
    @SAPManipulation.start_SAP
    def extrair_rel_zmm030(self, empreendimento:list|str, *, fechar_sap_no_final):
        _print(f"Iniciando extração das planilhas da transação ZMM030")
        download_path:str = self.download_path_zmm030
        if not os.path.exists(download_path):
            os.makedirs(download_path)
        else:
            for _ in range(60):
                try:
                    shutil.rmtree(download_path)
                    break
                except PermissionError as error:
                    arquivo_aberto = re.search(r"(?<=')[\D\d]+(?=')", str(error)).group() #type: ignore
                    Functions.fechar_excel(arquivo_aberto)
                pyautogui.sleep(1)
                
            for _ in range(60*2):
                try:
                    os.makedirs(download_path)
                    break
                except FileExistsError:
                    pyautogui.sleep(.5)
        
        if isinstance(empreendimento, str):
            if not self.__zmm030(centro=empreendimento, download_path=download_path):
                _print(f"error ao gerar relatorio do zmm030 '{empreendimento}' vide log")
        elif isinstance(empreendimento, list):
            for emp in empreendimento:
                if not self.__zmm030(centro=emp, download_path=download_path):
                    _print(f"error ao gerar relatorio do zmm030 '{emp}' vide log")
    
    @SAPManipulation.start_SAP           
    def extrair_rel_zmm019(self, *,fechar_sap_no_final):
        _print(f"Iniciando extração das planilhas da transação ZMM019")
        download_path:str = self.download_path_zmm019
        if not os.path.exists(download_path):
            os.makedirs(download_path)
        else:
            for _ in range(60):
                try:
                    shutil.rmtree(download_path)
                    break
                except PermissionError as error:
                    arquivo_aberto = re.search(r"(?<=')[\D\d]+(?=')", str(error)).group() #type: ignore
                    Functions.fechar_excel(arquivo_aberto)
                pyautogui.sleep(1)

            for _ in range(60*2):
                try:
                    os.makedirs(download_path)
                    break
                except FileExistsError:
                    pyautogui.sleep(.5)
        
        data_inicial:str = "01/01/2023"
        padrao_str:str = '%d.%m.%Y'
        for value in self.obter_datas(data_inicial):
            inicio = value['inicio'].strftime(padrao_str)
            fim = value['fim'].strftime(padrao_str)
            
            if not self.__zmm019(
                data_inicial=inicio,
                data_final=fim,
                download_path=download_path,
                file_name=f"de_{inicio}-ate_{fim}.xlsx"
            ):
                _print(f"error ao gerar relatorio do zmm019 'de_{inicio}-ate_{fim}' vide log")
        
        
    
    @SAPManipulation.start_SAP
    def __zmm030(self, *, centro:str, download_path:str) -> bool:
        _print(f"Iniciando extração do relatorio {centro.upper()} da tranzação ZMM030")
        try:
            self.session.findById("wnd[0]/tbar[0]/okcd").text = "/n zmm030"
            self.session.findById("wnd[0]").sendVKey(0)
            self.session.findById("wnd[0]/usr/ctxtS_WERKS-LOW").text = centro.upper()
            self.session.findById("wnd[0]/usr/chkP_SALDO").setFocus()
            self.session.findById("wnd[0]/usr/chkP_SALDO").selected = 'false'
            self.session.findById("wnd[0]/usr/chkP_VIGEN").selected = 'false'
            self.session.findById("wnd[0]/usr/chkP_VIGEN").setFocus()
            self.session.findById("wnd[0]/tbar[0]/btn[0]").press()
            self.session.findById("wnd[0]/tbar[1]/btn[8]").press()
            #self.session.findById("wnd[0]/tbar[1]/btn[2]").press()
            if self.session.findById("wnd[0]/sbar").text == 'Nenhum dado encontrado':
                _print(f"nenhum dado encontrado para o empreendimento {centro.upper()}")
                return True            
        except:
            Logs(name=f"{self.__class__.__name__}.__download_autoGui").register(status='Error', description=f"error ao gerar relatorio do zmm030 '{centro.upper()}' vide log", exception=traceback.format_exc())
            return False
        
        try: 
            self.__download_autoGui(download_path)
        except Exception as error:
            _print(f"error no __download_autoGui do empreendimento {centro.upper()}")
            Logs(name=f"{self.__class__.__name__}.__download_autoGui").register(status='Error', description=f"error no __download_autoGui do empreendimento {centro.upper()}", exception=traceback.format_exc())
            return False
        
        return True
    
    @SAPManipulation.start_SAP
    def __zmm019(
        self, *,
        data_inicial:str,
        data_final:str,
        download_path:str,
        file_name:str,
        variante:str = "VOL. COMPRAS",        
        ) -> bool:
        
        _print(f"Iniciando extração do relatorio 'de_{data_inicial}-ate_{data_final}' da tranzação ZMM019")
        
        if not file_name.endswith('.xlsx'):
            file_name += '.xlsx'
        
        try:
            self.session.findById("wnd[0]/tbar[0]/okcd").text = "/n zmm019"
            self.session.findById("wnd[0]").sendVKey(0)
            self.session.findById("wnd[0]/tbar[1]/btn[17]").press()
            self.session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").currentCellColumn = "TEXT"
            self.session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").selectedRows = "0"
            self.session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").contextMenu()
            self.session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").selectContextMenuItem("&FILTER")
            self.session.findById("wnd[2]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").text = variante
            self.session.findById("wnd[2]/tbar[0]/btn[0]").press()
            self.session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").selectedRows = "0"
            self.session.findById("wnd[1]/tbar[0]/btn[2]").press()
            self.session.findById("wnd[0]/usr/ctxtSO_AEDAT-LOW").text = data_inicial
            self.session.findById("wnd[0]/usr/ctxtSO_AEDAT-HIGH").text = data_final
            self.session.findById("wnd[0]/usr/ctxtSO_AEDAT-HIGH").caretPosition = 10
            self.session.findById("wnd[0]/tbar[1]/btn[8]").press()
            self.session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").setCurrentCell(3,"ERNAM")
            self.session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectedRows = "3"
            self.session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").contextMenu()
            self.session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectContextMenuItem("&XXL")
            self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
            self.session.findById("wnd[1]/usr/ctxtDY_PATH").text = download_path
            self.session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = file_name
            self.session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 13
            self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
            
            Functions.fechar_excel(os.path.join(download_path, file_name), wait=3)
            
            return True    
        except Exception as error:
            Logs(name=f"{self.__class__.__name__}.__zmm019").register(status='Error', description=f"error ao gerar relatorio do zmm019 'de_{data_inicial}-ate_{data_final}' vide log", exception=traceback.format_exc())
            return False
        
    def __download_autoGui(self, download_path:str) -> bool:
        try:
            janela_sap:Win32Window = pygetwindow.getWindowsWithTitle('Relatório de Valores de Contratos')[0]
            if not janela_sap.isActive:
                janela_sap.minimize()
                janela_sap.restore()
            janela_sap.resizeTo(900,600)
            janela_sap.moveTo(0,0)
            janela_sap.moveRel(600,100)         
   
            pyautogui.FAILSAFE = False
            pyautogui.sleep(1)

            bt_download = self.__procurar_imagem(r'Entities\images\download_sap\01-bt_download.png', confidence=0.9)
            self.__segurar_ponteiro(bt_download)
            pyautogui.moveRel(-60,15)
            pyautogui.click()

            pyautogui.sleep(2)
            janela_download:Win32Window = pygetwindow.getWindowsWithTitle('Procurar Arquivos ou Pastas')[0]
            if not janela_download.isActive:
                janela_download.minimize()
                janela_download.restore()
            janela_download.resizeTo(400,400)
            janela_download.moveTo(0,0)
            janela_download.moveRel(600,100)


            pyautogui.sleep(1)
            local_caminho = self.__procurar_imagem(r'Entities\images\download_sap\02-local_caminho.png', confidence=0.6)
            self.__segurar_ponteiro(local_caminho)
            pyautogui.moveRel(20,-22)
            pyautogui.doubleClick()
            pyautogui.hotkey('ctrl', 'a')
            pyautogui.press('backspace')
            pyautogui.typewrite(download_path)

            pyautogui.sleep(1)
            bt_ok = self.__procurar_imagem(r'Entities\images\download_sap\03-bt_ok.png', confidence=0.6)
            self.__segurar_ponteiro(bt_ok)
            pyautogui.moveRel(-40,0)
            pyautogui.click()
            
            return True
        except Exception as error:
            print(traceback.format_exc())
            Logs(name=f"{self.__class__.__name__}.__download_autoGui").register(status='Error', exception=traceback.format_exc())
            return False 
       
    def obter_datas(self, inicial_date:datetime|str, *, agora:datetime=datetime.now(), mes_atual:bool=True) -> List[Dict[str,datetime]]:
        result:List[Dict[str,datetime]] = []
        
        inicial:datetime
        if isinstance(inicial_date, str):
            inicial = datetime.strptime(inicial_date, '%d/%m/%Y')
        elif isinstance(inicial_date, datetime):
            inicial = inicial_date
        
        date:datetime = datetime(inicial.year, inicial.month, 1)
        
        while (date.month < agora.month) or (date.year < agora.year):
            fim_do_mes = (date + relativedelta(months=1)) - relativedelta(days=1)
            result.append({'inicio': date, 'fim': fim_do_mes})
            date = date + relativedelta(months=1)
        
        if mes_atual:
            comeco_mes_atual = datetime(agora.year, agora.month, 1)
            result.append({'inicio': comeco_mes_atual, 'fim': agora})

        return result
    
    def __segurar_ponteiro(self, target:pyautogui.pyscreeze.Box, *, tempo_espera:int=1.5, timeout:int=2*60) -> None: #type: ignore
        pyautogui.moveTo(target)
        lock = pyautogui.position()
        tempo_ultima_vez_que_moveu = datetime.now()
        pyautogui.FAILSAFE = True
        while True:
            posicao_atual = pyautogui.position()
            regra = (lock.x,lock.y) == (posicao_atual.x, posicao_atual.y)
            
            if regra:
                if datetime.now() >= (tempo_ultima_vez_que_moveu + relativedelta(seconds=tempo_espera)):
                    pyautogui.FAILSAFE = False
                    return
                #else:
                #    _print(f"Restando -> {(tempo_ultima_vez_que_moveu + relativedelta(seconds=tempo_espera)) - datetime.now()} segundos")
            else:
                tempo_ultima_vez_que_moveu = datetime.now()
                pyautogui.moveTo(target)
                lock = pyautogui.position()
            pyautogui.sleep(.1)
                    
    def __procurar_imagem(self, image_path:str, *, confidence:float=1, timeout:int=2*60) -> pyautogui.pyscreeze.Box: #type: ignore
        if (confidence < 0) and (confidence > 1):
            raise ValueError("valor do confidence incorreto apenas valores entre '0' e '1'")
        for _ in range(timeout*2):
            try:
                return pyautogui.locateOnScreen(os.path.join(os.getcwd(), image_path), confidence=confidence)
            except:
                pyautogui.sleep(0.5)            
       
    
if __name__ == "__main__":
    pass