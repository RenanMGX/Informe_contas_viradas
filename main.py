import sys
from Entities.f_01 import F_01, datetime
from Entities.dados import ContaStatus, pd, os
from Entities.dependencies.arguments import Arguments
from Entities.email import Email, Config, Credential

class Execute:
    @staticmethod
    def start():
        file_path:str = F_01().coletar()
        
        file:pd.DataFrame = pd.read_excel(file_path, dtype=str)
        del file['Total do período de comparação']
        del file['Desvio absoluto']
        
        file = file.dropna(subset=['Nº conta'])
        
        file = file[
           (~file['Nº conta'].str.startswith('1202')) &
           (~file['Nº conta'].str.startswith('1208')) &
           (~file['Nº conta'].str.startswith('2402')) &
           (~file['Nº conta'].str.startswith('49')) &
           (~file['Nº conta'].str.startswith('39')) 
        ]
        
        
        if file.empty:
            print("sem contas para informar")
            return
        
        conta_status = ContaStatus()
        
        for row, value in file.iterrows():
            conta_status.identify(value)
        
         
        try:
            os.unlink(file_path)
        except:
            pass
        
        email = Email().mensagem(
            Destino=Config()['email']['destino'],
            Assunto="Informe de Contas Contabeis Viradas",
            Corpo_email=""
        )
        
        anexo = False
        
        
        contas_viradas_path:str = os.path.join(r'C:\Users\renan.oliveira\Downloads', datetime.now().strftime('contas_viradas-%d%m%Y-.xlsx'))
        df_viradas = pd.DataFrame(conta_status.contas_viradas)
        if not df_viradas.empty:
            anexo = True
            df_viradas['Total período relatório'] = df_viradas['Total período relatório'].astype(float)
            df_viradas.to_excel(contas_viradas_path, index=False)
            email.Anexo(contas_viradas_path)
            try:
                os.unlink(contas_viradas_path)
            except:
                pass
        
        
        contas_nao_encontradas_path:str = os.path.join(r'C:\Users\renan.oliveira\Downloads', datetime.now().strftime('contas_nao_encontradas-%d%m%Y-.xlsx'))
        df_nao_encontradas = pd.DataFrame(conta_status.contas_nao_encontradas)
        if not df_nao_encontradas.empty:
            anexo = True
            df_viradas['Total período relatório'] = df_viradas['Total período relatório'].astype(float)
            df_nao_encontradas.to_excel(contas_nao_encontradas_path, index=False)
            email.Anexo(contas_nao_encontradas_path)
            try:
                os.unlink(contas_nao_encontradas_path)
            except:
                pass
            
        if anexo:
            email.send()
        else:
           Email().mensagem(
            Destino=Config()['email']['destino'],
            Assunto="Informe de Contas Contabeis Viradas",
            Corpo_email="Nenhuma Conta Virada encontrada"
        ) 
                
        
if __name__ == "__main__":
    Arguments({
        "start": Execute.start
    })
    