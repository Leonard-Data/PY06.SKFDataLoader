import subprocess
import time
import win32com.client
import os
from utils import setup_logger
logger = setup_logger(os.path.basename(__file__))
from typing import Literal
# Define key mappings (key names to ID keys)
KEY_MAPPINGS: dict = {
    'ENTER': '0',
    'F1': '1',
    'F2': '2',
    'F3': '3',
    'F4': '4',
    'F5': '5',
    'F6': '6',
    'F7': '7',
    'F8': '8',
    'F9': '9',
    'F10': '10',
    'F11': '11',
    'F12': '12',
    'BACKSPACE': '13',
    'TAB': '14',
    'ESC': '15',
    # Add other key mappings as needed
}



class SAPAutomation:
    def __init__(self, path: str, connection_name=None):
        '''
        Initializes the SAP Automation object with an optional connection name.
        path: SAP GUI in system - 'C:\\Program Files (x86)\SAP\\FrontEnd\SAPgui\saplogon.exe'\n
        object: SAPGUI
        '''
        self.path =path
        self.sap_gui = None
        self.application = None
        self.connection = None
        self.sessions = {}
        self.connection_name = connection_name

        # Attach to the SAP GUI

        if not self._attach_sap_gui():
            self._init_sap_gui()

    def _attach_sap_gui(self):
        """
        Private method to attach to the SAP GUI scripting interface.
        """
        try:
            self.sap_gui = win32com.client.GetObject("SAPGUI")  # Get the SAP GUI Scripting Engine
            if not self.sap_gui:
                raise Exception("SAP GUI is not running.")
            
            self.application = self.sap_gui.GetScriptingEngine
            logger.info("SAP GUI attached successfully.")
            return True

        except Exception as e:
            logger.error(f"Failed to attach to SAP GUI: {e}")
            return False

    def _init_sap_gui(self):
        """
        Private method to attach to the SAP GUI scripting interface.
        """
        try:
            subprocess.Popen(self.path)
            time.sleep(5)

            self.sap_gui = win32com.client.GetObject('SAPGUI')
            if not type(self.sap_gui) == win32com.client.CDispatch:
                return

            self.application = self.sap_gui.GetScriptingEngine
            if not type(self.application) == win32com.client.CDispatch:
                self.sap_gui = None
                return

        except Exception as e:
            logger.error(f"Failed to attach to SAP GUI: {e}")

    def connect(self):
        """
        Connects to the SAP system using the specified connection name.
        """
        if not self.application:
            raise Exception("SAP GUI Application is not initialized.")
        
        try:
            # Find the desired connection by name
            self.connection = self.application.OpenConnection(self.connection_name, True)

            if not type(self.connection) == win32com.client.CDispatch:
                self.application = None
                self.sap_gui = None
                return
            
            if not self.connection:
                raise Exception(f"Connection with name '{self.connection_name}' not found.")

            # Retrieve all sessions for the connection
            self._refresh_sessions()
            logger.info(f"Connected to SAP connection '{self.connection_name}' with {len(self.sessions)} sessions available.")
        
        except Exception as e:
            logger.error(f"Failed to connect to SAP session: {e}")

    def _refresh_sessions(self):
        """
        Refreshes the list of sessions available in the connection.
        """
        self.sessions.clear()
        for i in range(self.connection.Children.Count):
            self.sessions[i] = self.connection.Children(i)

    def select_session(self, session_id):
        """
        Selects a specific SAP session by its session ID.
        """
        if session_id not in self.sessions:
            raise Exception(f"Session with ID {session_id} does not exist.")
        
        self.session = self.sessions[session_id]
        logger.info(f"Selected SAP session {session_id}.")

    def create_new_session(self):
        """
        Creates a new SAP session within the current connection.
        """
        if not self.connection:
            raise Exception("SAP connection is not initialized. Call connect() first.")
        
        try:
            # Get the current active session
            current_session = self.connection.Children(0)  # Using the first session to open a new one
            if not current_session:
                # Use SAP scripting command to open a new session
                current_session.findById("wnd[0]/tbar[0]/okcd").text = "/oSAP Easy Access"
                current_session.findById("wnd[0]").sendVKey(0)
                
                # Wait briefly to allow the session to initialize
                time.sleep(2)  # Adjust sleep time as necessary

                # Refresh the session list after creating a new session
                self._refresh_sessions()
                logger.info(f"New SAP session created successfully. Total sessions now: {len(self.sessions)}")

        except Exception as e:
            logger.error(f"Failed to create a new SAP session: {e}")

    def login(self, username, password, client, language="EN",force_closed: bool=True)-> bool:
        """
        Logs into SAP using the provided credentials.
        """
        if not self.session:
            raise Exception("SAP session is not selected. Call select_session() first.")
        
        try:
            self.session.findById("wnd[0]/usr/txtRSYST-MANDT").text = client
            self.session.findById("wnd[0]/usr/txtRSYST-BNAME").text = username
            self.session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = password
            self.session.findById("wnd[0]/usr/txtRSYST-LANGU").text = language
            self.session.findById("wnd[0]").sendVKey(0)
            _type, message = self.check_transaction_status()
            if _type =='E':
                logger.error(message)
                return False
            else:
                logger.info("Logged into SAP successfully.")
                try:
                    if force_closed:
                        self.session.findById("wnd[1]/usr/radMULTI_LOGON_OPT1").select()
                        logger.info("Continue with this logon and end any other logons in the system")
                        self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
                        return True
                    else:
                        self.session.findById("wnd[1]/usr/radMULTI_LOGON_OPT3").select()
                        logger.info("Terminate this logon")
                        self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
                        return False
                except:
                    logger.error("Not found information for Multiples logons")
                return True
            
        
        except Exception as e:
            logger.error(f"Failed to log into SAP: {e}")
            return False

    def get_user(self):
        """
        Get User login information in current session
        """
        if not self.session:
            raise Exception("SAP session is not selected. Call select_session() first.")
        
        try:
            logger.info(f"Current user: {self.session.Info.User}")
            return self.session.Info.User
        
        except Exception as e:
            logger.error(f"Failed to get user login: {e}")
    

    def open_transaction(self, transaction_code):
        """
        Opens the specified SAP transaction.
        """
        if not self.session:
            raise Exception("SAP session is not selected. Call select_session() first.")
        
        try:
            self.session.StartTransaction(transaction_code)
            logger.info(f"Transaction '{transaction_code}' opened successfully.")
        
        except Exception as e:
            logger.error(f"Failed to open transaction '{transaction_code}': {e}")
    
    def check_transaction_status(self):
        """
        Checks the status of the transaction for any error messages.
        """
        if not self.session:
            raise Exception("SAP session is not selected. Call select_session() first.")
        
        try:
            status_bar = self.session.findById("wnd[0]/sbar")
            message_type = status_bar.MessageType
            message_text = status_bar.Text
            # Interpretation of message type
            if message_type == "I":
                logger.info(f"({message_text}")
            elif message_type == "W":
                logger.warning(f"{message_text}")
            elif message_type == "E":
                logger.error(f"{message_text}")
            elif message_type == "S":
                logger.info(f"{message_text}")
            else:
                logger.debug("Unknown message type.")
                return False, "Unknown message type."
            return message_type, message_text
        except Exception as e:
            print(f"Failed to check transaction status: {e}")
            return False, str(e)

    def read_log(self):
        """
        Reads messages from the SAP GUI status bar.
        """
        if not self.session:
            raise Exception("SAP session is not selected. Call select_session() first.")
        
        try:
            status_bar = self.session.findById("wnd[0]/sbar")
            log_message = status_bar.Text
            print(f"Read log message: {log_message}")
            return log_message
        
        except Exception as e:
            print(f"Failed to read log message: {e}")
            return None
        
    def execute_command(self, command, element_id):
        """
        Executes a specific command on the specified GUI element.
        """
        if not self.session:
            raise Exception("SAP session is not selected. Call select_session() first.")
        
        try:
            gui_element = self.session.findById(element_id)
            gui_element.text = command
            gui_element.sendVKey(0)
            logger.info(f"Executed command on element '{element_id}'.")
        
        except Exception as e:
            logger.error(f"Failed to execute command on element '{element_id}': {e}")
    
    def run_transaction(self, transaction_code):
        """Runs an SAP transaction."""
        if not self.session:
            raise Exception("SAP session is not selected. Call select_session() first.")
        
        try:
            self.session.findById("wnd[0]/tbar[0]/okcd").Text = f"/n{transaction_code}"
            self.session.findById("wnd[0]").sendVKey(0)
            logger.info(f"Executed transaction: {transaction_code}.")
        
        except Exception as e:
            logger.error(f"Failed to executed transaction: {transaction_code}: {e}")

    def open_window_and_run_transaction(self, transaction_code):
        """Runs an SAP transaction. on a new SAP window"""
        if not self.session:
            raise Exception("SAP session is not selected. Call select_session() first.")
        
        try:
            self.session.findById("wnd[0]/tbar[0]/okcd").Text = f"/o{transaction_code}"
            self.session.findById("wnd[0]").sendVKey(0)
            logger.info(f"Executed transaction: {transaction_code}.")
        
        except Exception as e:
            logger.error(f"Failed to executed transaction: {transaction_code}: {e}")

    def click_button(self, button_id):
        """
        Clicks a specific button in the SAP GUI.
        """
        if not self.session:
            raise Exception("SAP session is not selected. Call select_session() first.")
        
        try:
            button = self.session.findById(button_id)
            button.press()
            logger.info(f"Clicked button: '{button_id}'.")
        
        except Exception as e:
            logger.error(f"Failed to click button '{button_id}': {e}")

    def caret_position(self, element_id, pos = 0):
        """
        set focus and caret to element position
        """
        if not self.session:
            raise Exception("SAP session is not selected. Call select_session() first.")
        
        try:
            element = self.session.findById(element_id)
            element.setFocus()
            element.caretPosition = pos
            logger.info(f"caret position {pos} for element: '{element_id}'.")
        
        except Exception as e:
            logger.error(f"Failed to caret position '{element_id}': {e}")

    def click_toolbar_button(self, button_id):
        """Clicks a toolbar button in SAP GUI."""
        if not self.session:
            raise Exception("SAP session is not selected. Call select_session() first.")
        
        try:
            button = self.session.findById(button_id)
            button.press()
            logger.info(f"Clicked button: '{button_id}'.")
        
        except Exception as e:
            logger.error(f"Failed to click button '{button_id}': {e}")

    def disable_screenshots_on_error(self):
        """Disables screenshots on error."""
        if not self.session:
            raise Exception("SAP session is not selected. Call select_session() first.")
        try:
            self.session.scriptingSettings.screenshotOnError = False
        except Exception as e:
            logger.error(f"Failed to disable screenshot option: {e}")

    def doubleClick_element(self, element_id):
        """Double-clicks an element in the SAP GUI."""
        if not self.session:
            raise Exception("SAP session is not selected. Call select_session() first.")
        try:
            self.session.findById(element_id).doubleClick()
        except Exception as e:
            logger.error(f"Failed to double click element {element_id}: {e}")
        
    def element_should_be_present(self, element_id):
        """Checks if an element is present."""
        if not self.session:
            raise Exception("SAP session is not selected. Call select_session() first.")
        try:
            self.session.findById(element_id)
            return True
        except Exception as e:
            logger.error(f"Failed to FIND element {element_id}: {e}")
            return False

    def element_value_should_be(self, element_id, expected_value):
        """Asserts that the element's value matches the expected value."""
        if not self.session:
            raise Exception("SAP session is not selected. Call select_session() first.")
        try:
            actual_value = self.session.findById(element_id).Text
            assert actual_value == expected_value, f"Expected {expected_value}, but got {actual_value}."
        except:
            pass

    def element_value_should_contain(self, element_id, expected_value):
        """Asserts that the element's value contains the expected value."""
        actual_value = self.session.findById(element_id).Text
        assert expected_value in actual_value, f"Expected {actual_value} to contain {expected_value}."

    def enable_screenshots_on_error(self):
        """Enables screenshots on error."""
        self.session.scriptingSettings.screenshotOnError = True

    def focus_and_click(self, element_id):
        """Focuses on an element and clicks it."""
        if not self.session:
            raise Exception("SAP session is not selected. Call select_session() first.")
        try:
            element = self.session.findById(element_id)
            element.setFocus()
            element.press()
        except Exception as e:
            logger.error(f"Error when click element {element_id}: {e}")
            return False

    def focus_and_input_text(self, element_id, text):
        """Focuses on an element and inputs text."""
        if not self.session:
            raise Exception("SAP session is not selected. Call select_session() first.")
        try:
            element = self.session.findById(element_id)
            element.setFocus()
            element.text = text
        except Exception as e:
            logger.error(f"Failed to input text on element {element_id}: {e}")
            return False

    def generic_click_element(self, element_id):
        """Performs a generic click on an element."""
        self.click_element(element_id)

    def generic_input_password(self, element_id, password):
        """Inputs a password into a password field."""
        self.session.findById(element_id).Text = password

    def generic_input_text(self, element_id, text):
        """Inputs text into a text field."""
        self.focus_and_input_text(element_id, text)

    def get_cell_value(self, table_id, row, column):
        """Gets the value from a table cell."""
        if not self.session:
            raise Exception("SAP session is not selected. Call select_session() first.")
        try:
            return self.session.findById(table_id).getCell(row, column)
        except Exception as e:
            logger.error(f"Failed to get cell value on table '{table_id}': {e}")

    def get_element_location(self, element_id):
        """Gets the location of an element."""
        if not self.session:
            raise Exception("SAP session is not selected. Call select_session() first.")
        try:
            element = self.session.findById(element_id)
            return element.abs_x, element.abs_y
        except Exception as e:
            logger.error(f"Failed to get element locations': {e}")

    def get_element_type(self, element_id):
        """Gets the type of an element."""
        if not self.session:
            raise Exception("SAP session is not selected. Call select_session() first.")
        try:
            element = self.session.findById(element_id)
            return element.Type
        except Exception as e:
            logger.error(f"Failed to get element type': {e}")

    def get_element_type_of_object(self, element_id):
        """Gets the object type of an element."""
        element = self.session.findById(element_id)
        return element.Type

    def get_row_count(self, table_id):
        """Gets the number of rows in a table."""
        return self.session.findById(table_id).RowCount

    def get_scroll_position(self, element_id):
        """Gets the scroll position of an element."""
        element = self.session.findById(element_id)
        return element.VerticalScrollbar.position

    def get_statusbar_type(self):
        """Gets the type of the status bar message."""
        return self.session.findById("wnd[0]/sbar").Type

    def get_value(self, element_id):
        """Gets the value of an element."""
        return self.session.findById(element_id).Text

    def get_window_title(self):
        """Gets the title of the current SAP window."""
        return self.session.findById("wnd[0]").Text

    def input_password(self, element_id, password):
        """Inputs a password into a password field."""
        self.generic_input_password(element_id, password)

    def input_text(self, element_id, text):
        """Inputs text into a text field."""
        self.generic_input_text(element_id, text)

    def maximize_window(self):
        """Maximizes the SAP window."""
        self.session.findById("wnd[0]").maximize()

    def open_connection(self, connection_string):
        """Opens a new SAP connection."""
        self.connection = self.sap_gui.OpenConnection(connection_string)

    def press_f1(self):
        """Presses the F1 key."""
        self.session.findById("wnd[0]").sendVKey(1)

    def press_f4(self):
        """Presses the F4 key."""
        self.session.findById("wnd[0]").sendVKey(4)

    def scroll(self, element_id, position):
        """Scrolls an element to a specific position."""
        element = self.session.findById(element_id)
        element.VerticalScrollbar.Position = position

    def select_checkbox(self, checkbox_id):
        """Selects a checkbox."""
        self.session.findById(checkbox_id).Selected = True

    def select_context_menu_item(self, element_id, item_text):
        """Selects an item from a context menu."""
        element = self.session.findById(element_id)
        element.selectContextMenuItem(item_text)

    def select_from_list_by_label(self, element_id, label):
        """Selects an item from a list by label."""
        element = self.session.findById(element_id)
        element.key = label

    def select_node(self, tree_id, node_key):
        """Selects a node in a tree."""
        tree = self.session.findById(tree_id)
        tree.selectedNode = node_key

    def select_node_link(self, tree_id, node_key):
        """Selects a node link in a tree."""
        tree = self.session.findById(tree_id)
        tree.nodeContextMenu(node_key).select()

    def select_radio_button(self, radio_button_id):
        """Selects a radio button."""
        self.session.findById(radio_button_id).Select()

    def select_table_column(self, table_id, column):
        """Selects a column in a table."""
        self.session.findById(table_id).columnSelected = column

    def select_table_row(self, table_id, row):
        """Selects a row in a table."""
        self.session.findById(table_id).rowSelected = row

    def send_Vkey(self, vkey: Literal['ENTER','F1','F2','F3','F4','F5','F6','F7','F8','F9','F10','F11','F12','BACKSPACE','TAB','ESC']):
        """Sends a virtual key."""
        if not self.session:
            raise Exception("SAP session is not selected. Call select_session() first.")

        try: 
            self.session.findById("wnd[0]").sendVKey(KEY_MAPPINGS[vkey])
        except Exception as e:
            logger.error(f"Failed to send {vkey} key. {e}")
    

    def set_cell_value(self, table_id, row, column, value):
        """Sets the value of a table cell."""
        self.session.findById(table_id).setCell(row, column, value)

    def set_explicit_wait(self, seconds):
        """Sets an explicit wait time."""
        time.sleep(seconds)

    def set_focus(self, element_id):
        """Sets focus on an element."""
        self.session.findById(element_id).setFocus()

    def take_screenshot(self, file_path):
        """Takes a screenshot of the SAP window."""
        self.session.findById("wnd[0]").HardCopy(file_path)

    def unselect_checkbox(self, checkbox_id):
        """Unselects a checkbox."""
        self.session.findById(checkbox_id).Selected = False
    
    def close_session(self, session_id=None):
        """
        Closes the specified SAP session. If session_id is None, closes the current session.
        """
        if session_id is None:
            session_to_close = self.session
        else:
            if session_id not in self.sessions:
                raise Exception(f"Session with ID {session_id} does not exist.")
            session_to_close = self.sessions[session_id]

        if session_to_close:
            try:
                session_to_close.findById("wnd[0]").close()
                logger(f"SAP session {session_id} closed.")
                del self.sessions[session_id]
            except Exception as e:
                print(f"Failed to close SAP session {session_id}: {e}")