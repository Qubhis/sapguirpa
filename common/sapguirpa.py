import sys
import traceback
import re

import PySimpleGUI as sg
import win32com.client

from pythoncom import com_error


## below snippet is for fast connection to first active session from cmd ##
# import win32com.client
# sap_gui_auto = win32com.client.GetObject("SAPGUI")
# application = sap_gui_auto.GetScriptingEngine
# connection = application.Children(0)
# session = connection.Children(0)

'''
TODO:
1. read SAP GUI scripting API documentation of application and session object
    - application:
        - .Utils.ShowMessageBox(
                            title, 
                            text,
                            .Utils.<MSG_TYPE>,
                            .Utils.<MSG_OPTION>) - SAP GUI popups

    - session:
        - .SendCommand("/nend") simulates inserting into command field 
          and pressing enter

NOTE: 
This wrapper doesn't serve for automatic login as storing passwords 
inside a script is not safe. You can write your own bindings in case 
you use a password manager. 
'''
class SAPGUIRPA:

    def __init__(self):
        self.__sap_gui_auto = None
        self.__application = None
        self.__connection = None
        self.__session = None


    def attach_to_session(self):
        ''' 
        Simply gets SAPGUI object and scripting engine, scans for all
        connections and their sessions and prompts user to select session.
        Changes attributes of SAPGUIRPA instance.
        '''
        while True:
            try:
                self.__sap_gui_auto = win32com.client.GetObject("SAPGUI")
                self.__application = self.__sap_gui_auto.GetScriptingEngine

                session_indexes = dict()
                session_titles = list()
                
                for i, conn in enumerate(self.__application.Children):
                    for j, sess in enumerate(conn.Children):
                        # child is GuiMainWindow obj
                        title = sess.Children(0).Text
                        session_titles.append(title)
                        session_indexes.update(
                            {title: {
                                'conn_idx': i,
                                'sess_idx': j,
                            }}
                        )
                break

            except com_error:
                if gui_crash_report(title="Couldn't find any session...",
                                    button_layout="try_again"):
                    continue
                else:
                    break

        title = gui_dropdown_selection('Select SAP session for scripting',
                                       session_titles)
        if title is None:
            sg.Popup("Program ended. Press OK.")
            sys.exit()
        
        conn_idx = session_indexes[title]['conn_idx']
        sess_idx = session_indexes[title]['sess_idx']
        
        self.__connection = self.__application.Children(conn_idx)
        self.__session = self.__connection.Children(sess_idx)

    def start_transaction(self, transaction):
        self.__session.StartTransaction(transaction)

    def end_transaction(self, transaction):
        self.__session.EndTransaction(transaction)

    def lock_session_ui(self):
        self.__session.LockSessionUI()

    def unlock_session_ui(self):
        self.__session.UnlockSessionUI()

    def gui_maximize(self):
        self.__session.findById("wnd[0]").Maximize()

    def gui_restore_size(self):
        self.__session.findById("wnd[0]").Restore()

    def send_vkey(self, vkey, window="wnd[0]"):
        '''executes virtual key as per below
            0 -> Enter
            3 -> F3
            8 -> F8
            11 -> Save
            81 -> PageUp
            82 -> PageDown
            
            "main" means MainWindow - "wnd[0]"
            "modal" means ModalWindow - "wnd[1]"'''

        if vkey not in (0, 3, 8, 11, 81, 82):
            raise AssertionError(f"Vkey {vkey} is not supported!")
        else:
            self.__session.findById(window).sendVKey(vkey)

    def insert_value(self, element_id, value):
        '''takes element ID path and value to be inserted
        Inserts the value to the field. 
        
        Returns nothing'''
        field = self.__session.findById(element_id)
        type_check = field.type.find('TextField')

        assert type_check != -1, f"{element_id} is not text field"

        field.text = value

    def press_or_select(self, element_id, check=True):
        '''takes element ID path, optionally check=False to indicate
        desire for un-checking a checkbox.
        Acts accordingly based on .type property:
         - GuiButton
         - GuiCheckBox
         - GuiRadioButton
         - GuiTab
         - GuiMenu
 
        Returns nothing'''
        element = self.__session.findById(element_id)
        
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
            
        else:
            raise AssertionError(f'''{element_id} is not button, checkbox, , radiobutton, tab, or GuiMenu''')

    def select_tab(self, element_id):
        ''' takes element ID path
        returns nothing'''
        raise AssertionError("this method is not completed")

    def insert_values_standard(self, inputs=dict()):
        '''takes inputs in format {'type_of_field': [element_id, value]}
        supported types:
        - 'button' -> value is ""
        - 'checkbox' -> value is True for checked and False for opposite
        - 'text_field' -> desired value, "" for empty field 
        '''
        assert len(inputs) > 0, "provided inputs are empty!"

        for key, value in inputs.items():
            element_id = value[0]
            input_data = value[1]

            if key == 'text_field':
                self.insert_value(element_id, input_data)
            else:
                self.press_or_select(element_id, input_data)

    def get_element_by_id(self, element_id):
        ''' takes element id
        , returns element as an object -> we can use properties and methods'''
        return self.__session.findById(element_id)

    def get_element_text(self, element_id):
        ''' takes element ID,
        returns value of the element - if string value'''
        return self.__session.findById(element_id).text
    
    def get_screen_title(self, element_id):
        '''returns title of a current window'''
        assert len(element_id) == 6, "id is too long"
        return self.__session.findById(element_id).text

    def verify_element(self, element_id):
        '''returns true if element is found on a screen'''
        try:
            element = self.__session.findById(element_id)
            return True
        except com_error:
            return False

    def grid_view_get_cell_value(self, element_id, cell_name, row_index):
        ''' takes element id of a GridViewCtrl.1 object and technical name
        of a table cell, and index of row being read

        returns value from the cell in a string format'''
        grid_view_shellcont = self.__session.findById(element_id)
        return grid_view_shellcont.GetCellValue(row_index, cell_name)

    def grid_view_scrape_rows(self, element_id, cells):
        ''' takes element id of a grid view and list of cells to be fetched
        from each row

        returns list of lists of cell values '''

        grid_view_shellcont = self.__session.findById(element_id)
        total_row_count = grid_view_shellcont.RowCount
        scrapped_rows = list()
        rows_to_scroll = 23

        for row in range(0, total_row_count):

            shall_we_scroll = (
                row % rows_to_scroll == 0 
                and row + rows_to_scroll <= total_row_count
            )

            if shall_we_scroll:
                grid_view_shellcont.currentCellRow = row + rows_to_scroll
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
            
    def confirm_screen(self, element_id, vkey=0):
        ''' Takes element id of a confirmation button
        and presses the button, or window's id and presses enter.
        Additionally, it can take vkey value (in case we need to save transaction before continue to next screen) - by default is 0 (enter)

        Stores current name of a screen or a dialog for futher comparison,
        and confirms user's actions.
        Afterwards, it gets title again and compares titles. 

        returns true if we got to next screen or dialog'''
        
        if len(element_id) > 6: # single window has only 6 chars
            # extract window id
            window_id = re.search(r"wnd\[[0-9]\]", element_id).group()
        else:
            window_id = element_id

        curr_title = self.get_screen_title(window_id)

        if len(element_id) > 6:
            self.press_or_select(element_id)
        else:
            self.send_vkey(vkey, window_id)

        try:
            next_title = self.get_screen_title(window_id)
        except com_error:
    # TODO: here we are getting issue with window number
    #       we need to somehow identified wnd[] index for proper
    #       functioning
            next_title = self.get_screen_title("wnd[0]")

        if curr_title != next_title:
            return True
        else:
            return False

    def disconnect(self):
        self.__session = None
        self.__connection = None
        self.__application = None
        self.__sap_gui_auto = None


                
def gui_dropdown_selection(title='Select one option', dropdown_list=list()):
    
    '''takes title of the gui and list of items for selection
    creates gui window with dropdown selection
    returns value of selected item or None if closed'''

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


def gui_crash_report(title='Crash report',button_layout='just_ok'):
    '''takes text for title of gui and button layout
    - displays traceback info in gui window with few button layouts:
        - 'try_again' -> shows buttons 'Try again' and 'Exit'
        - 'just_ok' -> shows only 'OK' button
    
    returns None if closed or canceled by user
    returns True if Try again
    '''
    
    # extract exception traceback info
    exc_type, exc_value, exc_traceback = sys.exc_info()
    # get formated list of it
    traceback_string = traceback.format_exception(exc_type,
                                                  exc_value,
                                                  exc_traceback)
    
    # button and gui layout preparation
    buttons = list()
    if button_layout == 'just_ok':
        buttons.append(sg.Button('OK'))
    elif button_layout == 'try_again':
        buttons.append(sg.Button('Try again'), sg.Button('Exit'))
    
    gui_layout = [
        [sg.Multiline(default_text=(traceback_string), size=(70, 25))],
        [sg.Text('')],
        buttons
    ]

    window = sg.Window(title, gui_layout)
    # read users action and act accordingly
    while True:
        event, __ = window.Read()
        
        if event is None or event == 'Exit' or event == 'OK':
            window.Close()
            return None
        elif event == 'Try again':
            window.Close()
            return True
