import sys
import traceback
import re

import PySimpleGUI as sg
import win32com.client
import pywintypes
import openpyxl


## below snippet is for fast connection to first active session from command line ##
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
class SapGuiRpa:

    def __init__(self):
        self.sap_gui_auto = None
        self.application = None
        self.connection = None
        self.session = None


    def attach_to_session(self):
        ''' 
        Simply gets SAPGUI object and scripting engine, scans for all
        connections and their sessions and prompts user to select session.
        Changes attributes of SAPGUIRPA instance.
        '''
        while True:
            try:
                self.sap_gui_auto = win32com.client.GetObject("SAPGUI")
                self.application = self.sap_gui_auto.GetScriptingEngine

                session_indexes = dict()
                session_titles = list()
                
                for i, conn in enumerate(self.application.Connections):
                    for j, sess in enumerate(conn.Sessions):
                        # if session is Busy, we won't get anything out of it
                        #  neither Text property. Therefore we need to skip it
                        if sess.Busy:
                            continue
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

            except pywintypes.com_error:
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
        
        self.connection = self.application.Children(conn_idx)
        self.session = self.connection.Children(sess_idx)

    def start_transaction(self, transaction):
        self.session.StartTransaction(transaction)

    def end_transaction(self):
        self.session.EndTransaction()

    def lock_session_ui(self):
        self.session.LockSessionUI()

    def unlock_session_ui(self):
        self.session.UnlockSessionUI()

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
        , returns element as an object -> we can use properties and methods'''
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
        # visible_row_count = grid_view_shellcont.VisibleRowCount
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
        self.sap_gui_auto = None


def gui_dropdown_selection(title='Select one option', dropdown_list=None):
    
    '''takes title of the gui and list of items for selection
    creates gui window with dropdown selection
    returns value of selected item or None if closed'''

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
    buttons = []
    if button_layout == 'just_ok':
        buttons.extend([sg.Button('OK')])
    elif button_layout == 'try_again':
        buttons.extend([sg.Button('Try again'), sg.Button('Exit')])
    
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


def gui_repeat_or_continue(title="Human action needed!", info_text=""):
    '''
    GUI window to prompt user to take action when some command
    couldn't be executed. There are just two options:
     - continue
     - repeat last command

    Takes in info about command which couldn't be executed and optionally takes info about next step in the the program flow
    (recommended)
    
    My recommendation is using this within while loop to control the program flow. This function returns "repeat" or "next" string.
    ''' 

    gui_layout = [
        [sg.Multiline(default_text=info_text, size=(70, 25))],
        [sg.Text("")],
        [sg.Button("Repeat last step"), sg.Button("Next step")]
    ]

    window = sg.Window(title, gui_layout)

    while True:
        event, __ = window.Read()
        window.Close()
        if event is None:
            sys.exit()
        else:
            return event

def read_excel_file(path_to_excel_file):
    ''' 
    reads excel and returns tuple of tuples where first tuple
    is header line and others are each row
    NOTE: values to be loaded must be in sheet named 'INPUTS' 
    '''
    workbook = openpyxl.load_workbook(filename=path_to_excel_file,
                                      data_only=True)
    input_data = workbook["INPUTS"]
    all_rows = input_data.rows
    return_list = [tuple(cell.value for cell in row) for row in all_rows]
    return tuple(return_list)    


if __name__ == "__main__":
    pass
