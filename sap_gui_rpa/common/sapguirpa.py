import win32com.client
import pywintypes

'''
NOTE: You can see several functions except the class which I wrote at the very
beginning of my days with SAP GUI scripting. Those can be omitted except 
two of them, gui_crash_report and gui_dropdown_selection as those two are 
used in method attach_to_session.  
'''

class SapGuiRpa:
    '''
    Wrapper around GuiApplication object to simplify script development.
    Code should be mostly self-explanatory, but in case it's not, please use
    SAP official documentation for more information:
    https://help.sap.com/viewer/b47d018c3b9b45e897faf66a6c0885a8/760.05/en-US/babdf65f4d0a4bd8b40f5ff132cb12fa.html

    Supports attaching to the running instance of SAP GUI through the Running
    Object Table only. This wrapper doesn't support automation of logging in.
    '''
    def __init__(self):
        self.sap_gui_object = None
        self.application = None
        self.connection = None
        self.session = None


    def _get_available_sessions(self):
        '''scans for all connections and their sessions of SAP Gui instance.

        returns dictionary with session titles and indexes for connection
        '''
        available_sessions = dict()
        for conn_index, connection in enumerate(self.application.Connections):
            for sess_index, session in enumerate(connection.Sessions):
                if session.Busy:
                    continue
                # get title of the main window - only child of Session object
                title = session.Children(0).Text
                session_details = {
                    title: {
                        "conn_index": conn_index,
                        "sess_index": sess_index
                    }
                }
                available_sessions.update(session_details)
        
        return available_sessions


    def attach_to_session(self, mode="cli"):
        ''' 
        Lists all available, non-busy sessions and prompts for selection 
        through command line interface by default, or through GUI window (PySimpleGui)

        Throws SapLogonNotStarted if SAP Logon is not running.

        Updates class instance attributes.
        '''
        try:
            self.sap_gui_object = win32com.client.GetObject("SAPGUI")
        except pywintypes.com_error as com_error:
            if com_error.args[0] == -2147221020:
                raise SapLogonNotStarted

        self.application = self.sap_gui_object.GetScriptingEngine

        available_sessions = self._get_available_sessions()
        if not available_sessions:
            raise NoAvailableSession

        # select one of available session
        selected_session = select_session(available_sessions, mode)
        conn_index = available_sessions[selected_session]['conn_index']
        sess_index = available_sessions[selected_session]['sess_index']
        # update attributes
        self.connection = self.application.Children(conn_index)
        self.session = self.connection.Children(sess_index)


    def start_transaction(self, transaction):
        self.session.StartTransaction(transaction)


    def end_transaction(self):
        self.session.EndTransaction()


    def gui_maximize(self):
        self.get_element_by_id("wnd[0]").Maximize()


    def gui_restore_size(self):
        self.get_element_by_id("wnd[0]").Restore()


    def send_vkey(self, vkey, window=None):
        '''executes virtual key as per below - please add if missing
            0 -> Enter
            2 -> F2
            3 -> F3
            8 -> F8
            11 -> Save
            81 -> PageUp
            82 -> PageDown

        window is by default 'wnd[0]' if not provided
            '''
        if window is None:
            window = "wnd[0]"

        if vkey not in (0, 2, 3, 8, 11, 81, 82):
            raise AssertionError(f"Vkey {vkey} is not supported!")
        else:
            self.get_element_by_id(window).sendVKey(vkey)


    def insert_value(self, element_id, value):
        '''takes element ID path and value to be inserted
        Inserts the value to the field. 
        
        Returns nothing'''
        
        element = self.get_element_by_id(element_id)

        if element.type in ("GuiTextField", "GuiCTextField"):
            element.text = value
        elif element.type == "GuiComboBox":
            element.key = value
        else:
            raise AssertionError(f"{element_id} cannot be filled with value/key {value}")


    def press_or_select(self, element_id, check=True):
        '''takes element ID path, optionally check=False to indicate
        desire for un-checking a checkbox.
        Action press or select is based on .type property of the element:
         - GuiButton
         - GuiCheckBox
         - GuiRadioButton
         - GuiTab
         - GuiMenu
         - GuiLabel - in search results or just simple text label
 
        Returns nothing'''
        element = self.get_element_by_id(element_id)
        
        if element.type == 'GuiButton':
            element.setFocus()
            element.press()

        elif element.type == 'GuiCheckBox':
            if check:
                element.selected = -1
            else:
                element.selected = 0

        elif element.type == 'GuiRadioButton':
            if element.selected == False:
                element.select()

        elif element.type in ('GuiTab', 'GuiMenu'):
            element.select()
        
        elif element.type == 'GuiLabel':
            element.setFocus()
            
        else:
            raise AssertionError(f'''{element_id} is not button, checkbox, , radiobutton, tab, or menu.''')


    def get_element_by_id(self, element_id):
        ''' takes element id
        , returns element as an object -> we can use properties and methods
        from GuiAplication object (SAPGUI) when needed'''
        return self.session.findById(element_id)


    def get_element_text(self, element_id):
        ''' takes element ID,
        returns value of the element - if string value'''
        return self.get_element_by_id(element_id).text


    def get_element_type(self, element_id):
        '''takes in element_id
        returns type of the element'''
        return self.get_element_by_id(element_id).type


    def get_status_bar(self):
        '''returns status bar data in format tuple(message_type, text)
         message types are
            - S success
            - W warning
            - E error
            - A abort
            - I information
        '''
        status_bar = self.get_element_by_id("wnd[0]/sbar")
        return (status_bar.MessageType, status_bar.text)


    def verify_element(self, element_id):
        '''returns true if element is found on a screen'''
        try:
            self.get_element_by_id(element_id)
            return True
        except pywintypes.com_error:
            return False


    def insert_row_gridview(self, gridview_id, row_index, tech_name, value):
        '''inserst values into gridview table in to given index'''
        gridview = self.get_element_by_id(gridview_id)
        gridview.modifyCell(row_index, tech_name, value)
    

    def grid_view_get_cell_value(self, element_id, cell_name, row_index):
        ''' takes element id of a GridViewCtrl.1 object and technical name
        of a table cell, and index of row being read

        returns value from the cell in a string format'''
        grid_view_shellcont = self.get_element_by_id(element_id)
        return grid_view_shellcont.GetCellValue(row_index, cell_name)


    def grid_view_scrape_rows(self, element_id, cells):
        ''' takes element id of a grid view and list of cells to be fetched
        from each row

        returns list of lists of cell values '''

        grid_view_shellcont = self.get_element_by_id(element_id)
        total_row_count = grid_view_shellcont.RowCount
        rows_to_scroll = grid_view_shellcont.VisibleRowCount
        scrapped_rows = list()

        for row in range(0, total_row_count):
            shall_we_scroll = (
                row % rows_to_scroll == 0 
                and row + rows_to_scroll <= total_row_count
                and rows_to_scroll != total_row_count
            )
            if shall_we_scroll:
                grid_view_shellcont.currentCellRow = row + rows_to_scroll - 1
            elif total_row_count - row < rows_to_scroll:
                grid_view_shellcont.currentCellRow = total_row_count - 1 
            
            row_content = list()
            for cell in cells:
                cell_value = self.grid_view_get_cell_value(element_id,
                                                           cell,
                                                           row)
                row_content.append(cell_value)

            scrapped_rows.append(row_content)
        return scrapped_rows


    def table_select_absolute_row(self, element_id, index):
        table_control = self.get_element_by_id(element_id)
        table_control.GetAbsoluteRow(index).selected = True


    def disconnect(self):
        self.session = None
        self.connection = None
        self.application = None
        self.sap_gui_object = None


def select_session(available_sessions, mode="cli"):
    '''Prompts user to select one entry from available_sessions either via 
    command line interface or GUI window created by PySimpleGui library'''
    if mode == "cli":
        choices = {index + 1: key for index, key in enumerate(available_sessions.keys())}
        print("These are available sessions:")
        for index,key in choices.items():
            print(f"\t{index}.\t{key}")
        
        user_prompted = True
        while user_prompted:
            user_choice = int(input("Please provide corresponding number one of the above items:"))
            if user_choice not in choices:
                print("...invalid input, please repeat:")
            else:
                session_title = choices[user_choice]
                user_prompted = False
    
    elif mode == "gui":
        session_titles = list(available_sessions.keys())
        session_title = gui_dropdown_selection(
            title='Select SAP session for scripting',
            dropdown_list=session_titles
        )

    return session_title


def gui_dropdown_selection(title='Select one option', dropdown_list=None):
    '''takes title of the gui and list of items for selection
    creates gui window with dropdown selection
    returns value of selected item or None if closed'''
    import PySimpleGUI as sg
    
    if dropdown_list is None:
        dropdown_list = list()
    assert len(dropdown_list) > 0, "provided list is empty!"
    layout = [
        [sg.Text('Please choose one of below options:')],
        [sg.InputCombo(dropdown_list, size=(40, 10))],
        [sg.Submit(), sg.Text("", size=(16,1)), sg.Exit()]
    ]
    window = sg.Window(title, layout)
    while True:
        event, value = window.Read()
        if event is None or event == 'Exit':
            window.Close()
            return None

        elif event == 'Submit':
            window.Close()
            return value[0]


class SapLogonNotStarted(Exception):
    def __init__(self):
        message = "SAP Logon instance has not been found."\
                  "Please ensure you've opened SAP Logon and it's running."
        super().__init__(message)


class NoAvailableSession(Exception):
    def __init__(self):
        message = "Either all sessions are busy or no session is opened."
        super().__init__(message)


if __name__ == "__main__":
    pass
