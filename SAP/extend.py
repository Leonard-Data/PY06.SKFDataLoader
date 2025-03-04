import os
import time
from SAP.scripts import SAPAutomation
from utils import setup_logger
logger = setup_logger(os.path.basename(__file__))

class ExtendFunctions(SAPAutomation):
    def __init__(self , path, connection_name):
        SAPAutomation.__init__(self, path, connection_name)
    
    def FI_posting_R013(self, 
                        path: str, 
                        session: str,
                        export_result: bool = False):
        """Opens the specified SAP transaction."""
        logger.info("start FI posting on SAP on T-Code: R013")
        if not self.session:
            raise Exception("SAP session is not selected. Call select_session() first.")

        try: 
            # 1. maximum window:
            self.maximize_window()
            # 2. call transactions:
            self.run_transaction("YFI_OC_GEN_R013") 
            # 3. send F4 Key -> open export files.
            self.press_f4()
            # 4. input folder and filename:
            self.focus_and_input_text(element_id="wnd[1]/usr/ctxtDY_PATH",
                                      text=os.path.dirname(path))
            self.focus_and_input_text(element_id="wnd[1]/usr/ctxtDY_FILENAME",
                                      text=os.path.basename(path))
            # 5. click OK
            self.click_button(button_id="wnd[1]/tbar[0]/btn[0]")
            # 6. Enter Prefix session name:
            self.focus_and_input_text(element_id="wnd[0]/usr/txtP_GROUP",
                                      text=session)
            # 7. Execute click
            self.click_button(button_id="wnd[0]/tbar[1]/btn[8]")
            type, message = self.check_transaction_status()
            if type == 'E':
                logger.error(f"FI posting upload program failed. Details: {message}")
                return False
            
            # 8. Click back
            self.click_button(button_id="wnd[0]/tbar[0]/btn[3]")
            type, message = self.check_transaction_status()
            return True          
        except Exception as e:
            logger.error(f"Failed to open transaction. Details: {e}")
            return False
    
    def read_log_SM35P(self,path, session):
        """Opens the specified SAP transaction."""
        logger.info("Reading processes log on SAP on T-Code: SM35P")
        user = self.get_user()
        if not self.session:
            raise Exception("SAP session is not selected. Call select_session() first.")
        try: 
            # 1. maximum window:
            self.maximize_window()
            # 2. call transactions:
            self.run_transaction("SM35P")
            # 3. get sessions overview:
            self.click_button(button_id="wnd[0]/tbar[1]/btn[17]")
            # 4.1. filter by Sess.name:
            self.focus_and_input_text(element_id="wnd[0]/usr/subD1000_HEADER:SAPMSBDC_CC:1005/txtD0100-MAPN",
                                      text=session)
            # 4.2. filter by User:
            self.focus_and_input_text(element_id="wnd[0]/usr/subD1000_HEADER:SAPMSBDC_CC:1005/txtD0100-CREATOR",
                                      text=user)
            # 5. click/ enter to apply filter:
            self.send_Vkey("ENTER")
            is_existed = self.element_should_be_present(
                element_id="wnd[0]/usr/tabsD1000_TABSTRIP/tabpALLE/ssubD1000_SUBSCREEN:SAPMSBDC_CC:1010/tblSAPMSBDC_CCTC_APQI/txtITAB_APQI-GROUPID[0,0]"
            )
            if is_existed:
                self.caret_position(
                    element_id="wnd[0]/usr/tabsD1000_TABSTRIP/tabpALLE/ssubD1000_SUBSCREEN:SAPMSBDC_CC:1010/tblSAPMSBDC_CCTC_APQI/txtITAB_APQI-GROUPID[0,0]",
                    pos=0
                )
            else:
                logger.error("cannot founded any sessions to be processed!")
                return False
            #6. click to process selected session
            self.click_button(button_id="wnd[0]/tbar[1]/btn[8]")
            #7. select process in background options
            self.select_radio_button(radio_button_id="wnd[1]/usr/radD0300-BATCH")
            #8. click to applied:
            self.click_button(button_id="wnd[1]/tbar[0]/btn[0]")
            #9. refresh transactions:
            self.run_transaction("SM35P")
            # 10.1 filter by Sess.name:
            self.focus_and_input_text(element_id="wnd[0]/usr/subSCR_INFO:RSBDC_PROTOCOL:0201/txtD0100-MAPN",
                                      text=session)
            # 10.2 filter by User:
            self.focus_and_input_text(element_id="wnd[0]/usr/subSCR_INFO:RSBDC_PROTOCOL:0201/txtD0100-CREATOR",
                                      text=user)
            # 11. click/ enter to apply filter:
            self.send_Vkey(vkey='ENTER')
            # 12. get 1st row status ! = in processing
            isExisted = self.element_should_be_present("wnd[0]/usr/tabsTAB_PROTOCOL/tabpALL_PROT/ssubSCR_CONTENT:RSBDC_PROTOCOL:0210/tblRSBDC_PROTOCOLTC_PROTOCOL/txtLIST_APQLI-MAPPENSTATE[3,0]")
            if isExisted:
                retry = 0
                is_completed = False
                while not is_completed and retry < 150:
                    status = self.get_value(
                        element_id="wnd[0]/usr/tabsTAB_PROTOCOL/tabpALL_PROT/ssubSCR_CONTENT:RSBDC_PROTOCOL:0210/tblRSBDC_PROTOCOLTC_PROTOCOL/txtLIST_APQLI-MAPPENSTATE[3,0]"
                        )
                    if status == "In Processing":
                        # 12.1 refresh current view:
                        self.click_button(button_id="wnd[0]/tbar[1]/btn[13]")
                        time.sleep(3)
                        is_completed = False
                        retry +=1
                    else:
                        is_completed = True
                # 13. select the 1st row:
                self.set_focus(
                    element_id="wnd[0]/usr/tabsTAB_PROTOCOL/tabpALL_PROT/ssubSCR_CONTENT:RSBDC_PROTOCOL:0210/tblRSBDC_PROTOCOLTC_PROTOCOL/txtLIST_APQLI-CRETIME[1,0]")

                # 14. click to print log:
                self.send_Vkey(vkey='F2')
                # 15. select export to excel menu:
                self.click_button(button_id="wnd[0]/tbar[0]/btn[86]")
                self.select_radio_button(radio_button_id="wnd[0]/mbar/menu[0]/menu[1]/menu[1]")
                self.click_button(button_id="wnd[1]/tbar[0]/btn[0]")
                # 16. input folder and filename:
                self.focus_and_input_text(element_id="wnd[1]/usr/ctxtDY_PATH",
                                        text=os.path.dirname(path))
                self.focus_and_input_text(element_id="wnd[1]/usr/ctxtDY_FILENAME",
                                        text=os.path.basename(path))
                # 17. click OK
                self.click_button(button_id="wnd[1]/tbar[0]/btn[0]")
                _type,message = self.check_transaction_status()
                if "already exists" in message:
                    self.click_button(button_id="wnd[1]/tbar[0]/btn[11]")
                logger.info("Reading processes log on SAP completed.")
                time.sleep(2)
                return path
            else:
                logger.error(f"Failed to read log from SM35P: No log sessions generated!")
                return None
        except Exception as e:
            logger.error(f"Failed to read log from SM35P: {e}")
            return None