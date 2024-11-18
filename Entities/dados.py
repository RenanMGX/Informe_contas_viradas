import pandas as pd
import os
from typing import List, Dict
import numpy as nb
from Entities.dependencies.config import Config    

class ContaStatus:
    @property
    def contas_viradas(self) -> List[pd.Series]:
        return self.__contas_viradas
    
    @property
    def contas_nao_encontradas(self) -> List[str]:
        return self.__contas_nao_encontradas
    
    def __init__(self) -> None:
        path:str = Config()['path']['plano_contas']
        path = ContaStatus.verificar_caminho(path)
        
        self.__df_base:pd.DataFrame = pd.read_excel(path, dtype=str)
        self.__contas_viradas: List[pd.Series] = []
        self.__contas_nao_encontradas:List[str] = []
    
    @staticmethod
    def verificar_caminho(path:str) -> str:
        if not os.path.exists(path):
            raise FileNotFoundError(f'o arquivo não foi encontrado "{path}"')
        if not path.lower().endswith('.xlsx'):
            raise Exception("é aceito apenas arquivos .xlsx")
        return path
        
    def identify(self, conta:pd.Series) -> None:        
        if conta['Nº conta'] is nb.nan:
            return
        
        num_conta:str = conta['Nº conta']
        
        result = self.__df_base[
            self.__df_base['Conta do Razão'] == conta['Nº conta']
        ]['Natureza Contábil']
        
        if not result.empty:
            natureza:str = str(result.values[0])
            try:
                valor:float = float(conta['Total período relatório']) #type: ignore
            except:
                print(conta['Nº conta'], conta['Total período relatório'],'Valor invalido')
                return
            
            if num_conta.startswith('120902'):
                if valor > 0:
                    self.__contas_viradas.append(conta)               
            elif natureza == 'Devedora':
                if valor < 0:
                    self.__contas_viradas.append(conta)
            elif natureza == 'Credora':
                if valor > 0:
                    self.__contas_viradas.append(conta)
            else:
                print(conta['Nº conta'], "natureza não encontrada")
        else:
            if (not num_conta.startswith('9')) and (not num_conta.startswith('6')):
                self.__contas_nao_encontradas.append(str(conta['Nº conta']))

if __name__ == "__main__":
    pass