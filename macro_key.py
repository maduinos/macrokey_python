from pynput.keyboard import Listener, Key, KeyCode, Controller
from io import BytesIO
import win32clipboard
from PIL import Image
import time
import sys
from datetime import datetime
import logging

logging.basicConfig(
        filename = 'macrokey.log',
        format = '%(asctime)s:[%(levelname)s]%(message)s',
        datefmt = '%m/%d/%Y %I:%M:%S %p',
        level=logging.DEBUG
        )

cur_time = datetime.now()

keyboard = Controller()
store = set()
state = 0
quit_flag = 0

HOT_KEYS = {
    #'macro_1': set([ KeyCode(char='1'), KeyCode(char='0')] ),
    #'macro_1': set([ Key.ctrl_l, KeyCode(char='1')] ),
    'macro_1': set([ Key.tab, KeyCode(char='1')] ),
    'macro_2': set([ Key.tab, KeyCode(char='2')] ),
    'macro_3': set([ Key.tab, KeyCode(char='3')] ),
    'macro_4': set([ Key.tab, KeyCode(char='4')] ),
    'macro_5': set([ Key.tab, KeyCode(char='5')] ),
    'macro_6': set([ Key.tab, KeyCode(char='6')] ),
    'macro_7': set([ Key.tab, KeyCode(char='7')] ),
    'macro_8': set([ Key.tab, KeyCode(char='8')] ),
    'macro_9': set([ Key.tab, KeyCode(char='9')] ),
    'macro_10': set([ Key.tab, Key.esc] )
}

def send_to_clipboard(clip_type, data):
    win32clipboard.OpenClipboard()
    win32clipboard.EmptyClipboard()
    win32clipboard.SetClipboardData(clip_type, data)
    win32clipboard.CloseClipboard()

def file_copy(filepath):
    image = Image.open(filepath)
    output = BytesIO()
    image.convert("RGB").save(output, "BMP")
    data = output.getvalue()[14:]
    output.close()
    send_to_clipboard(win32clipboard.CF_DIB, data)

def macro_1():
    global state
    filepath = "macro_1.png"
    file_copy(filepath)
    #print(f'{cur_time} Execute Macro_1 Success')
    logging.info('Execute Macro_1 Success')
    state = 1

def macro_2():
    global state
    filepath = "macro_2.png"
    file_copy(filepath)
    logging.info('Execute Macro_2 Success')
    state = 1

def macro_3():
    global state
    filepath = "macro_3.png"
    file_copy(filepath)
    logging.info('Execute Macro_3 Success')
    state = 1

def macro_4():
    global state
    filepath = "macro_4.png"
    file_copy(filepath)
    logging.info('Execute Macro_4 Success')
    state = 1

def macro_5():
    global state
    filepath = "macro_5.png"
    file_copy(filepath)
    logging.info('Execute Macro_5 Success')
    state = 1

def macro_6():
    global state
    filepath = "macro_6.png"
    file_copy(filepath)
    logging.info('Execute Macro_6 Success')
    state = 1

def macro_7():
    global state
    filepath = "macro_7.png"
    file_copy(filepath)
    logging.info('Execute Macro_7 Success')
    state = 1

def macro_8():
    global state
    filepath = "macro_8.png"
    file_copy(filepath)
    logging.info('Execute Macro_8 Success')
    state = 1

def macro_9():
    pass

def macro_10():
    global quit_flag
    logging.info('Quit Flag Success')
    quit_flag = 1

def key_paste():
    global state
    keyboard.press(Key.ctrl)
    keyboard.press('v')
    keyboard.release('v')
    keyboard.release(Key.ctrl)
    keyboard.press(Key.enter)
    keyboard.release(Key.enter)
    logging.info('Paste Success')
    state = 0
 
def handleKeyPress( key ):
    global state
    if state == 0:
        store.add( key )
        logging.debug('key %s', key)
        for action, trigger in HOT_KEYS.items():
            CHECK = all([ True if triggerKey in store else False for triggerKey in trigger ])
            if CHECK:
                try:
                    func = eval( action )
                    if callable( func ):
                        func() 
                except NameError as err:
                    logging.error('Error %s', err)

 
def handleKeyRelease( key ):
    global state
    global quit_flag
    if key in store:
        store.remove( key )
    #if key == Key.esc
    if quit_flag == 1:
        logging.info('END')
        return False
    if state == 1:
        key_paste()
 
with Listener(on_press=handleKeyPress, on_release=handleKeyRelease) as listener:
    listener.join()

