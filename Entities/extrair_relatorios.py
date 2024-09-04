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
import pyperclip
import locale
locale.setlocale(locale.LC_ALL, "Portuguese_Brazil.1252")

class ExtrairRelatorio(SAPManipulation):
    @property
    def download_path_zmm030(self) -> str:
        return os.path.join(os.getcwd(), r'downloads\zmm030')
    
    @property
    def download_path_zmm019(self) -> str:
        return os.path.join(os.getcwd(), r'downloads\zmm019')
    
    @property
    def download_path_me5a(self) -> str:
        return os.path.join(os.getcwd(), r'downloads\me5a')
    
    @property
    def file_name_zmm019_compras(self) -> str:
        agora = self.__date
        nome = f"{agora.strftime('%Y')}\\{agora.strftime('%m%B').upper()}\\Relatorio_Vol.Compras {agora.strftime('%m')} ({agora.strftime('%d')}{agora.strftime('%B').title()}{agora.strftime('%Y')}).xlsx"
        return nome
    
    @property
    def file_name_contratos(self) -> str: # me5a, zmm030
        agora = self.__date
        nome = f"{agora.strftime('%Y')}\\{agora.strftime('%m%B').upper()}\\Relatorio_Contratos {agora.strftime('%m')} ({agora.strftime('%d')}{agora.strftime('%B').title()}{agora.strftime('%Y')}).xlsx"
        return nome
    
    def __init__(self, *, choicer:Literal['SAP_PRD', 'SAP_QAS'], date:datetime=datetime.now()) -> None:
        crd:dict = Credential(choicer).load() #type: ignore
        super().__init__(user=crd['user'], password=crd['password'], ambiente=crd['ambiente'])
        self.__date: datetime = date
        
    def __preparar_download_path(self, download_path) -> str:
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
        return download_path
    
    @SAPManipulation.start_SAP
    def extrair_rel_zmm030(self, empreendimento:list|str, *, fechar_sap_no_final:bool=False):
        _print(f"Iniciando extração das planilhas da transação ZMM030")
        download_path:str = self.__preparar_download_path(self.download_path_zmm030)

        if isinstance(empreendimento, str):
            if not self.__zmm030(centro=empreendimento, download_path=download_path):
                _print(f"error ao gerar relatorio do zmm030 '{empreendimento}' vide log")
        elif isinstance(empreendimento, list):
            for emp in empreendimento:
                if not self.__zmm030(centro=emp, download_path=download_path):
                    _print(f"error ao gerar relatorio do zmm030 '{emp}' vide log")

    @SAPManipulation.start_SAP
    def extrair_rel_me5a(self, empreendimento:list|str, *, fechar_sap_no_final:bool=False):
        _print(f"Iniciando extração das planilhas da transação ME5A")
        download_path:str = self.__preparar_download_path(self.download_path_me5a)
        
        if isinstance(empreendimento, str):
            if not self.__me5a(centro=empreendimento, download_path=download_path):
                _print(f"error ao gerar relatorio do me5a '{empreendimento}' vide log")
        elif isinstance(empreendimento, list):
            for emp in empreendimento:
                if not self.__me5a(centro=emp, download_path=download_path):
                    _print(f"error ao gerar relatorio do me5a '{emp}' vide log")

    
    #@SAPManipulation.start_SAP           
    def extrair_rel_zmm019(self, *,fechar_sap_no_final=False,   data_atual:datetime=datetime.now(), empreendimentos:list=[], data_inicial:str = "01/01/2023"):
        _print(f"Iniciando extração das planilhas da transação ZMM019")
        download_path:str = self.__preparar_download_path(self.download_path_zmm019)
      
        
        padrao_str:str = '%d.%m.%Y'
        for value in self.obter_datas(data_inicial, agora=data_atual):
            inicio = value['inicio'].strftime(padrao_str)
            fim = value['fim'].strftime(padrao_str)
            
            if not self.__zmm019(
                data_inicial=inicio,
                data_final=fim,
                download_path=download_path,
                empreendimentos=empreendimentos,
                fechar_sap_no_final=False,
                file_name=f"de_{inicio}-ate_{fim}.xlsx"
            ):
                _print(f"error ao gerar relatorio do zmm019 'de_{inicio}-ate_{fim}' vide log")
            #pyautogui.sleep(60)
        
        
    
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
    def __me5a(self, *, centro:str, download_path:str) -> bool:
        _print(f"Iniciando extração do relatorio {centro.upper()} da tranzação me5a")
        file_name = f"me5a_{centro}.xlsx"
        
        try:
            self.session.findById("wnd[0]/tbar[0]/okcd").text = "/n me5a"
            self.session.findById("wnd[0]").sendVKey(0)
            self.session.findById("wnd[0]/usr/ctxtS_WERKS-LOW").text = "E038"
            self.session.findById("wnd[0]/usr/ctxtS_WERKS-LOW").setFocus()
            self.session.findById("wnd[0]/tbar[1]/btn[8]").press()
            self.session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").setCurrentCell(2,"TXZ01")
            self.session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectedRows = "2"
            self.session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").contextMenu()
            self.session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectContextMenuItem("&XXL")
            self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
            self.session.findById("wnd[1]/usr/ctxtDY_PATH").text = download_path
            self.session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = file_name
            self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
            
            Functions.fechar_excel(os.path.join(download_path, file_name), wait=3)
            return True


        except:
            Logs(name=f"{self.__class__.__name__}").register(status='Error', description=f"error ao gerar relatorio do me5a '{centro.upper()}' vide log", exception=traceback.format_exc())
            return False
        
    
    @SAPManipulation.start_SAP
    def __zmm019(
        self, *,
        data_inicial:str,
        data_final:str,
        download_path:str,
        file_name:str,
        empreendimentos:list|str=[],
        fechar_sap_no_final=False,
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
            
            # self.session.findById("wnd[0]/usr/btn%_SO_WERKS_%_APP_%-VALU_PUSH").press()
            
            # texto_para_copiar = '\r\n'.join(empreendimentos)
            # pyperclip.copy(texto_para_copiar)
            # self.session.findById("wnd[1]/tbar[0]/btn[24]").press()
            
            # self.session.findById("wnd[1]/tbar[0]/btn[8]").press()  
            # self.session.findById("wnd[0]/usr/txtP_TOTAL").text = "999999999"
            self.session.findById("wnd[0]/usr/ctxtP_LAYOUT").text = "NOVOLAR"


            self.session.findById("wnd[0]/tbar[1]/btn[8]").press()
                
            
            
            self.session.findById("wnd[0]/tbar[1]/btn[33]").press()
            self.session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").setCurrentCell(0,"TEXT")
            self.session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").selectedRows = "0"
            self.session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").contextMenu()
            self.session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").selectContextMenuItem("&FILTER")
            self.session.findById("wnd[2]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").text = "RPA - VOL DE COMPRAS"
            self.session.findById("wnd[2]/tbar[0]/btn[0]").press()
            self.session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").currentCellColumn = "TEXT"
            self.session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").selectedRows = "0"
            self.session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").clickCurrentCell()
            
            #import pdb; pdb.set_trace()
                
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
                try:
                    janela_sap.activate()
                except:
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
                try:
                    janela_download.activate()
                except:
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
        """
        Retorna uma lista de dicionários com as datas de início e fim de cada mês, incluindo o mês atual se especificado.

        Args:
            incluir_mes_atual (bool): Se True, inclui o mês atual na lista. Padrão é True.

        Returns:
            list: Lista de dicionários com as datas de início e fim de cada mês.
        """        
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
    
    @SAPManipulation.start_SAP
    def finalizar(self, *, fechar_sap_no_final:Literal[True]):
        _print("Finalizando SAP")
    
    def __segurar_ponteiro(self, target:pyautogui.pyscreeze.Box, *, tempo_espera:int=1.5, timeout:int=2*60) -> None: #type: ignore
        """
        Mantém o ponteiro do mouse em uma posição específica por um determinado tempo.

        Args:
            x (int): A coordenada x da posição desejada.
            y (int): A coordenada y da posição desejada.
            duration (int): O tempo em segundos durante o qual o ponteiro deve permanecer na posição.
        """
        pyautogui.moveTo(target)

        lock = pyautogui.position()
        tempo_ultima_vez_que_moveu = datetime.now()
        pyautogui.FAILSAFE = True
        while True:
            posicao_atual = pyautogui.position()
            regra = (lock.x,lock.y) == (posicao_atual.x, posicao_atual.y)
            
            if regra:
                if datetime.now() >= (tempo_ultima_vez_que_moveu + relativedelta(seconds=tempo_espera)):
                    pyautogui.moveTo(target)
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
        """
        Tenta localizar uma imagem na tela dentro de um tempo limite especificado.

        Args:
            image_path (str): O caminho para a imagem a ser localizada.
            confidence (float, opcional): O nível de confiança para a correspondência da imagem. Deve estar entre 0 e 1. Padrão é 1.
            timeout (int, opcional): O tempo limite em segundos para tentar localizar a imagem. Padrão é 120 segundos.

        Raises:
            ValueError: Se o valor de confidence estiver fora do intervalo [0, 1].

        Returns:
            pyautogui.pyscreeze.Box: As coordenadas da caixa delimitadora da imagem localizada.
        """        
        # Verifica se o valor de confidence está fora do intervalo [0, 1]
        if (confidence < 0) or (confidence > 1):
            raise ValueError("valor do confidence incorreto apenas valores entre '0' e '1'")
        for _ in range(timeout*2):
            try:
                return pyautogui.locateOnScreen(os.path.join(os.getcwd(), image_path), confidence=confidence)
            except:
                pyautogui.sleep(0.5)           
       
    
if __name__ == "__main__":
    pass