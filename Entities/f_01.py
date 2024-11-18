from .dependencies.sap import SAPManipulation
from .dependencies.credenciais import Credential
from .dependencies.config import Config
from .dependencies.functions import Functions, sleep
from datetime import datetime
from dateutil.relativedelta import relativedelta
import os
from getpass import getuser



class F_01(SAPManipulation):
    def __init__(self):
        crd:dict = Credential(Config()['credential']['crd']).load()
        super().__init__(user=crd['user'], password=crd['password'], ambiente=crd['ambiente'])
        
    
    @SAPManipulation.start_SAP
    def coletar(self) -> str:
        date = datetime.now() - relativedelta(months=1)
        
        self.session.findById("wnd[0]/tbar[0]/okcd").text = "/n f.01"
        self.session.findById("wnd[0]").sendVKey(0)
        self.session.findById("wnd[0]/usr/tabsTABSTRIP_TABBL1/tabpUCOM1/ssub%_SUBSCREEN_TABBL1:RFBILA00:0001/radBILAGRID").select()
        self.session.findById("wnd[0]/usr/ctxtSD_KTOPL-LOW").text = "GPN"
        self.session.findById("wnd[0]/usr/ctxtSD_RLDNR-LOW").text = "0L"
        self.session.findById("wnd[0]/usr/tabsTABSTRIP_TABBL1/tabpUCOM1/ssub%_SUBSCREEN_TABBL1:RFBILA00:0001/ctxtBILAVERS").text = "BAPN"
        self.session.findById("wnd[0]/usr/tabsTABSTRIP_TABBL1/tabpUCOM1/ssub%_SUBSCREEN_TABBL1:RFBILA00:0001/txtBILBJAHR").text = date.strftime('%Y')
        self.session.findById("wnd[0]/usr/tabsTABSTRIP_TABBL1/tabpUCOM1/ssub%_SUBSCREEN_TABBL1:RFBILA00:0001/txtB-MONATE-LOW").text = "1"
        self.session.findById("wnd[0]/usr/tabsTABSTRIP_TABBL1/tabpUCOM1/ssub%_SUBSCREEN_TABBL1:RFBILA00:0001/txtB-MONATE-HIGH").text = date.strftime('%m')
        self.session.findById("wnd[0]/usr/tabsTABSTRIP_TABBL1/tabpUCOM1/ssub%_SUBSCREEN_TABBL1:RFBILA00:0001/txtBILVJAHR").text = date.strftime('%Y')
        self.session.findById("wnd[0]/usr/tabsTABSTRIP_TABBL1/tabpUCOM1/ssub%_SUBSCREEN_TABBL1:RFBILA00:0001/txtV-MONATE-LOW").text = "1"
        self.session.findById("wnd[0]/usr/tabsTABSTRIP_TABBL1/tabpUCOM1/ssub%_SUBSCREEN_TABBL1:RFBILA00:0001/txtV-MONATE-HIGH").text = date.strftime('%m')
        self.session.findById("wnd[0]/usr/tabsTABSTRIP_TABBL1/tabpUCOM1/ssub%_SUBSCREEN_TABBL1:RFBILA00:0001/radBILAGRID").setFocus()
        self.session.findById("wnd[0]/tbar[1]/btn[8]").press()
        self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
        file_temp_path = os.path.join(f'C:\\Users\\{getuser()}\\Downloads', date.strftime('contas_viradas_%d%m%Y%H%M%S_.xlsx'))
        self.session.findById("wnd[0]/tbar[1]/btn[26]").press()
        self.session.findById("wnd[0]/tbar[1]/btn[19]").press()
        self.session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").contextMenu()
        self.session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectContextMenuItem("&XXL")
        self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
        self.session.findById("wnd[1]/usr/ctxtDY_PATH").text = os.path.dirname(file_temp_path)
        self.session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = os.path.basename(file_temp_path)
        self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
        sleep(3)
        for _ in range(3):
            try:
                Functions.fechar_excel(file_temp_path)
            except:
                pass
        
        self.fechar_sap()
        
        return file_temp_path
        
    @SAPManipulation.start_SAP
    def test(self):
        self.session.findById("wnd[0]/tbar[0]/okcd").text = "/n f.01"
        self.session.findById("wnd[0]").sendVKey(0)

if __name__ == "__main__":
    pass
