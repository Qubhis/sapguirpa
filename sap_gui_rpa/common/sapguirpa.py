import win32com.client
import pywintypes


class SapGuiRpa:
    
    """Wrapper around GuiApplication object to simplify script development.

    Code should be mostly self-explanatory, but in case it's not, please use
    SAP official documentation for more information:
    https://help.sap.com/viewer/b47d018c3b9b45e897faf66a6c0885a8/760.05/en-US/babdf65f4d0a4bd8b40f5ff132cb12fa.html

    Supports attaching to the running instance of SAP GUI through the Running
    Object Table only. This wrapper doesn't support automation of logging in.

    I highly recommend to use scripting tracker along when creating your own
    scripts. It will help you to identify correct element's id on a screen,
    and support recording in python code -> https://tracker.stschnell.de/ 

    Attributes
    ----------
    sap_gui_object : win32com.clientCDispatch
        running object entry of SAPGUI

    application : win32com.clientCDispatch 
        running SAP Logon process - GuiApplication

    connection : win32com.clientCDispatch
        connection between SAP GUI and an application server - GuiConnection

    session : win32com.clientCDispatch
        point for performing specific actions by users (scripts)- GuiSession
    """

    def __init__(self):
        self.sap_gui_object = None
        self.application = None
        self.connection = None
        self.session = None


    def _get_available_sessions(self):
        """internal method for getting a list of all available session
        which are not in busy state

        Returns
        -------
        {dict}
            key is title(str) of Session, value is dict containing child's 
            indexes(int) for connection and session
            {
                title: {
                    "conn_index": int, "sess_index": int
                }
            }
        """
        available_sessions = dict()
        for conn_index, connection in enumerate(self.application.Connections):
            for sess_index, session in enumerate(connection.Sessions):
                if session.Busy:
                    continue
                # get title of the main window - the only child of Session obj
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
        """Prompts user to select one of available sessions for attaching to
        and updates instance attributes

        Parameters
        ----------
        mode : str, optional
            "cli" by default - prompts via command line interface;
            "gui" - prompts via GUI window

        Raises
        ------
        SapLogonNotStarted
            if SAP Logon instance is not running

        NoAvailableSession
            when no session is running or all of them are in busy state
        """
        # Get the application object from the running object table (ROT)
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
        conn_index = available_sessions[selected_session]["conn_index"]
        sess_index = available_sessions[selected_session]["sess_index"]
        # update remaining attributes
        self.connection = self.application.Children(conn_index)
        self.session = self.connection.Children(sess_index)


    def start_transaction(self, transaction):
        """Executes transaction code

        Parameters
        ----------
        transaction : str
            transaction code

        Examples
        ----------
        >>> SapGuiRpa.start_transaction("ME21N")

        has the same effect as you would type "/nME21N" in SAP GUI's command
        line
        """
        self.session.StartTransaction(transaction)


    def end_transaction(self):
        """returns to SAP Easy Access screen - same effect as you would type
        "/n" in SAP GUI's command line
        """
        self.session.EndTransaction()


    def gui_maximize(self):
        """Maximazes the main window """
        self.get_element_by_id("wnd[0]").Maximize()


    def gui_restore_size(self):
        """restores the main window and its modal windows(if exists)"""
        self.get_element_by_id("wnd[0]").Restore()


    def send_vkey(self, vkey, window="wnd[0]"):
        """Executes virtual key on the window

        Parameters
        ----------
        vkey : int
            number of the virtual keys (list below)
        window : str, optional
            id of the window, by default "wnd[0]" - main

        Raises
        ------
        AssertionError
            only if not implemented in this class

        List of implemented virtual keys
        --------------------------------
            0 -> Enter
            2 -> F2
            3 -> F3
            8 -> F8
            11 -> Save
            81 -> PageUp
            82 -> PageDown
        
        More keys can be implemented if needed, full list in SAP help
        """
        if vkey not in (0, 2, 3, 8, 11, 81, 82):
            raise AssertionError(f"Vkey {vkey} is not supported!")
        else:
            self.get_element_by_id(window).sendVKey(vkey)


    def insert_value(self, element_id, value):
        """takes element id and inserts provided value to its respective
        attribute (text or key) based on the element's type

        Supported types are GuiTextField, GuiCTextField, GuiComboBox
        
        Parameters
        ----------
        element_id : str
            e.g. "wnd[0]/usr/ctxtEM_MATNR-LOW"

        value : str
            value to be passed to the attribute

        Raises
        ------
        AssertionError
            if the type of provided element is not compatible
        """        
        element = self.get_element_by_id(element_id)

        if element.type in ("GuiTextField", "GuiCTextField"):
            element.text = value
        elif element.type == "GuiComboBox":
            element.key = value
        else:
            raise AssertionError(f"{element_id} cannot be filled with value/key {value}")


    def press_or_select(self, element_id, check=True):
        """Executes press, select, or setFocus method based on element's type

        Modifies attribute 'selected' if the element is a checkbox

        Supported elements:
         - GuiButton
         - GuiCheckBox
         - GuiRadioButton
         - GuiTab
         - GuiMenu
         - GuiLabel - in search results or just simple text label


        Parameters
        ----------
        element_id : str
            e.g. "wnd[0]/usr/ctxtEM_MATNR-LOW"

        check : bool, optional
            only for checkbox (GuiCheckBox), by default True

        Raises
        ------
        AssertionError
            if the type of provided element is not compatible
        """
        element = self.get_element_by_id(element_id)
        
        if element.type == "GuiButton":
            element.setFocus()
            element.press()

        elif element.type == "GuiCheckBox":
            if check:
                element.selected = -1
            else:
                element.selected = 0

        elif element.type == "GuiRadioButton":
            if element.selected == False:
                element.select()

        elif element.type in ("GuiTab", "GuiMenu"):
            element.select()
        
        elif element.type == "GuiLabel":
            element.setFocus()
            
        else:
            raise AssertionError(f"{element_id} is not label, button, checkbox, radiobutton, tab, or menu.")


    def get_element_by_id(self, element_id):
        """takes element id and returns its instance for further manipulation

        Parameters
        ----------
        element_id : str
            e.g. "wnd[0]/usr/ctxtEM_MATNR-LOW"

        Returns
        -------
        {object}
            instance of the element
        """
        return self.session.findById(element_id)


    def get_element_text(self, element_id):
        """returns text property of an element

        Parameters
        ----------
        element_id : str
            e.g. "wnd[0]/usr/ctxtEM_MATNR-LOW"

        Returns
        -------
        {str}
            value of element's text property
        """
        return self.get_element_by_id(element_id).text


    def get_element_type(self, element_id):
        """returns element type property of an element

        Parameters
        ----------
        element_id : str
            e.g. "wnd[0]/usr/ctxtEM_MATNR-LOW"

        Returns
        -------
        {str}
            value of element's type property
        """
        return self.get_element_by_id(element_id).type


    def get_status_bar(self):
        """returns message type and text properties of status bar object
        (GuiStatusbar)

        Returns
        -------
        {tuple}
            tuple containing message type and text, e.g. (message_type, text)
        
        message types are:
            S = success,
            W = warning,
            E = error,
            A = abort,
            I = information
        """
        status_bar = self.get_element_by_id("wnd[0]/sbar")
        return (status_bar.MessageType, status_bar.text)


    def verify_element(self, element_id):
        """returns True if element is found on a screen, else False

        Parameters
        ----------
        element_id : str
            e.g. "wnd[0]/usr/ctxtEM_MATNR-LOW"

        Returns
        -------
        {bool}
        """
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
        """to be used in the end of your script to stop active scripting
        """
        self.session = None
        self.connection = None
        self.application = None
        self.sap_gui_object = None


def select_session(available_sessions, mode="cli"):
    '''Prompts user to select one entry from available_sessions either via 
    command line interface or GUI window created by PySimpleGui library'''
    if mode == "cli":
        choices = {index + 1: key for index, key in enumerate(available_sessions.keys())}
        print("These sessions are available:")
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
            title="Select SAP session for scripting",
            dropdown_list=session_titles
        )

    return session_title


def gui_dropdown_selection(title, dropdown_list):
    '''takes title of the gui and list of items for selection
    creates gui window with dropdown selection
    returns value of selected item or None if closed'''  
    import PySimpleGUI as sg
    
    layout = [
        [sg.Text("Please choose one of below options:")],
        [sg.InputCombo(dropdown_list, size=(40, 10))],
        [sg.Submit(), sg.Text("", size=(16,1)), sg.Exit()]
    ]
    window = sg.Window(title, layout)
    while True:
        event, value = window.Read()
        if event is None or event == "Exit":
            window.Close()
            return None

        elif event == "Submit":
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
