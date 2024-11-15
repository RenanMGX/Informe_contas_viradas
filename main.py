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
        
        conta_status = ContaStatus()
        
        for row, value in file.iterrows():
            conta_status.identify(value)
            
        try:
            os.unlink(file_path)
        except:
            pass
        
        contas_viradas_path:str = os.path.join(r'C:\Users\renan.oliveira\Downloads', datetime.now().strftime('contas_viradas-%d%m%Y-.xlsx'))
        pd.DataFrame(conta_status.contas_viradas).to_excel(contas_viradas_path, index=False)
        
        contas_nao_encontradas_path:str = os.path.join(r'C:\Users\renan.oliveira\Downloads', datetime.now().strftime('contas_nao_encontradas-%d%m%Y-.xlsx'))
        pd.DataFrame(conta_status.contas_nao_encontradas).to_excel(contas_nao_encontradas_path, index=False)
        
        Email().mensagem(
            Destino=Config()['email']['destino'],
            Assunto="Informe de Contas Contabeis Viradas",
            Corpo_email=""
        ).Anexo(
            contas_viradas_path
        ).Anexo(
            contas_nao_encontradas_path
        ).send()
        
        try:
            os.unlink(contas_viradas_path)
        except:
            pass
        try:
            os.unlink(contas_nao_encontradas_path)
        except:
            pass

if __name__ == "__main__":
    Arguments({
        "start": Execute.start
    })
    