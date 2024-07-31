import xlwings as xw
from xlwings.main import Book
from time import sleep
from datetime import datetime
import re

class Functions:
    @staticmethod
    def fechar_excel(path:str, *, timeout:int=5, wait:int=0) -> bool:
        if wait > 0:
            sleep(wait)
        try:
            achou:bool = False
            for _ in range(timeout):
                for app in xw.apps:
                    for open_app in app.books:
                        open_app:Book
                        if open_app.name in path:
                            open_app.close()
                            if len(xw.apps) <= 0:
                                app.kill()                        
                            achou = True
                        # if not re.search(r'Pasta[0-9]+', open_app.name) is None:
                        #     open_app.close()
                        #     if len(xw.apps) <= 0:
                        #         app.kill()                        
                sleep(1)
            if achou:
                return True
            return False
        except:
            return False
    
    @staticmethod
    def excel_open() -> list:
        open_excel:list = []
        for app in xw.apps:
            for open_app in app.books:
                open_app:Book
                open_excel.append(open_app.name)
        return open_excel
    
    @staticmethod    
    def tratar_caminho(path:str) -> str:
        if (path.endswith("\\")) or (path.endswith("/")):
            path = path[0:-1]
        return path
    
def _print(*args, end="\n"):
    if not end.endswith("\n"):
        end += "\n"
    value = ""
    for arg in args:
        value += f"{arg} " 
    
    print(datetime.now().strftime(f"[%d/%m/%Y - %H:%M:%S] - {value}"), end=end)

        
if __name__ == "__main__":
    
    _print("ola")
    