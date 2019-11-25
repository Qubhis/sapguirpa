import logging
import sys
import traceback

import PySimpleGUI as sg
import win32com.client

from pytwintypes import com_error

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
                if gui_crash_report("Couldn't find any session..."):
                    continue
                else:
                    break

        title = gui_dropdown_selection('Select SAP session for scripting')
        if title is None:
            sg.Popup("Program ended. Press OK.")
            sys.exit()
        
        conn_idx = session_indexes[title]['conn_idx']
        sess_idx = session_indexes[title]['sess_idx']
        
        self.__connection = self.__application.Children(conn_idx)
        self.__session = self.__connection.Children(sess_idx)
    
        logging.debug(f"Attached to {title}")

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

    def send_vkey(self, vkey):
        '''executes virtual key as per below
            0 -> Enter
            3 -> F3
            8 -> F8
            11 -> Save
            81 -> PageUp
            82 -> PageDown'''

        if vkey not in (0, 3, 8, 11, 81, 82):
            raise AssertionError(f"Vkey {vkey} is not supported!")

        self.__session.findById("wnd[0]").sendVKey(vkey)

    def insert_value(self, element_id, value):
        '''takes element ID path and value to be inserted
        Inserts the value to the field. Returns nothing'''
        field = self.__session.findById(element_id)
        type_check = field.type.find('TextField')

        assert type_check != -1, f"{element_id} is not text field"

        field.text = value

    def press_or_select(self, element_id, check=True):
        '''takes element ID path, optionally check=False to indicate
        desire for un-checking a checkbox.
        acts accordingly based on .type property. 
        Returns nothing'''
        element = self.__session.findById(element_id)
        
        if element.type == 'GuiButton':
            element.press()

        elif element.type == 'GuiCheckBox':
            if check:
                element.selected = -1
            else:
                element.selected = 0

        elif element.type in ('GuiRadionButton', 'GuiTab'):
            element.select()
            
        else:
            raise AssertionError(f'''{element_id} is not button, checkbox, or 
                                 radiobutton''')

    def select_tab(self, element_id):
        ''' takes element ID path
        returns nothing'''
        self.__session.find


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

    def get_element_value(self, element_id):
        ''' takes element ID,
        returns value of the element - if string value'''
        return self.__session.findById(element_id).text


                
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

    window = sg.Window(title).Layout(layout)

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
        button_layout,
    ]

    window = sg.Window(title).Layout(gui_layout)
    # read users action and act accordingly
    while True:
        event, __ = window.Read()
        
        if event is None or event == 'Exit' or event == 'OK':
            window.Close()
            return None
        elif event == 'Try again':
            window.Close()
            return True
