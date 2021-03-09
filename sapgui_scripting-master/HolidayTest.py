from sapgui_script.framework import Runnable
from sapgui_script.framework import Transaction


def main():    
    class ExecuteScript(Runnable):
        def run(self, ses, tcode, row, utils):
            ses.FindById("wnd[0]").resizeWorkingPane(231,39,False)
            ses.StartTransaction(tcode)                        
            ses.FindById("wnd[0]").sendVKey(0)        
            ses.FindById("wnd[0]/usr/subSUBSCR_PERNR:SAPMP50A:0120/ctxtRP50G-PERSONID_EXT").text = row["SAPID"]
            ses.FindById("wnd[0]/usr/subSUBSCR_PERNR:SAPMP50A:0120/ctxtRP50G-PERSONID_EXT").caretPosition = 7
            ses.FindById("wnd[0]").sendVKey(0)        
            ses.FindById("wnd[0]/usr/ctxtRP50G-EINDA").text = row["BEGDA"]
            ses.FindById("wnd[0]/usr/ctxtRP50G-EINDA").setFocus()
            ses.FindById("wnd[0]/usr/ctxtRP50G-EINDA").caretPosition = 8
            ses.FindById("wnd[0]").sendVKey(0)        
            ses.FindById("wnd[0]/usr/tblSAPMP50ATC_MENU_EVENT/txtT529T-MNTXT[0,9]").setFocus()
            ses.FindById("wnd[0]/usr/tblSAPMP50ATC_MENU_EVENT/txtT529T-MNTXT[0,9]").caretPosition = 24
            ses.FindById("wnd[0]/usr/tblSAPMP50ATC_MENU_EVENT").verticalScrollbar.position = 3
            ses.FindById("wnd[0]/usr/tblSAPMP50ATC_MENU_EVENT").verticalScrollbar.position = 6
            ses.FindById("wnd[0]/usr/tblSAPMP50ATC_MENU_EVENT").verticalScrollbar.position = 9
            ses.FindById("wnd[0]/usr/tblSAPMP50ATC_MENU_EVENT").verticalScrollbar.position = 12
            ses.FindById("wnd[0]/usr/tblSAPMP50ATC_MENU_EVENT").getAbsoluteRow(34).selected = True
            ses.FindById("wnd[0]/usr/tblSAPMP50ATC_MENU_EVENT/txtT529T-MNTXT[0,22]").setFocus()
            ses.FindById("wnd[0]/usr/tblSAPMP50ATC_MENU_EVENT/txtT529T-MNTXT[0,22]").caretPosition = 0
            ses.FindById("wnd[0]/tbar[1]/btn[8]").press()
            ses.FindById("wnd[0]/usr/ctxtP0000-MASSG").text = row["MASSG"]
            ses.FindById("wnd[0]/usr/ctxtP0000-MASSG").setFocus()
            ses.FindById("wnd[0]/usr/ctxtP0000-MASSG").caretPosition = 2
            utils.log.add(ses)
            ses.FindById("wnd[0]/tbar[0]/btn[11]").press()
            ses.FindById("wnd[0]/usr/txtP2001-STDAZ").text = row["HOURS"]
            ses.FindById("wnd[0]/usr/txtP2001-STDAZ").setFocus()
            ses.FindById("wnd[0]/usr/txtP2001-STDAZ").caretPosition = 7
            utils.log.add(ses)
            ses.FindById("wnd[0]/tbar[0]/btn[11]").press()
            ses.FindById("wnd[0]/tbar[0]/btn[11]").press()
            utils.log.add(ses)
            ses.FindById("wnd[0]/tbar[1]/btn[21]").press()
            utils.log.add(ses)
            ses.FindById("wnd[0]/tbar[0]/btn[11]").press()        
            

    tr = Transaction('PA40', 'C:\\yuar\\testData.csv')                
    tr.runScript(ExecuteScript())


if __name__ == "__main__":
    main()
