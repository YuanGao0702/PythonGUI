import win32gui
import re
import pyautogui
from time import sleep
import winsound
import json
import win32con    

class WindowMgr:
    """Encapsulates some calls to the winapi for window management"""

    def __init__ (self):
        """Constructor"""
        self._handle = None

    def find_window(self, class_name, window_name=None):
        """find a window by its class_name"""
        self._handle = win32gui.FindWindow(class_name, window_name)        
        return self._handle

    def _window_enum_callback(self, hwnd, wildcard):
        """Pass to win32gui.EnumWindows() to check all the opened windows"""
        if re.match(wildcard, str(win32gui.GetWindowText(hwnd))) is not None:
            self._handle = hwnd
        return self._handle

    def find_window_wildcard(self, wildcard):
        """find a window whose title matches the wildcard regex"""
        self._handle = None
        win32gui.EnumWindows(self._window_enum_callback, wildcard)
        return self._handle

    def set_foreground(self):
        """put the window in the foreground"""
        win32gui.SetForegroundWindow(self._handle)
        return self._handle

def beep():
    frequency = 600  # Set Frequency To 2500 Hertz
    duration = 250  # Set Duration To 1000 ms == 1 second    
    winsound.Beep(frequency, duration)

def typewritestrs(strs):
    pyautogui.typewrite(strs)
    pyautogui.press('enter')

def scancsn(csn):
    typewritestrs(csn)
    sleep(0.1)

def test1():
    w = WindowMgr()   
    handle = w.find_window_wildcard("Lync*")    
    print(handle)
    w.set_foreground()
    
MAX_WAIT_SSN = 5
MAX_WAIT_CSN = 10

def main():
    
    print('*************START*************')
    w = WindowMgr()    
    data = None    
    with open('data.json') as f:
        data = json.load(f)
    for ssnele in data['ssns']:
        pastseconds = 0.0
        ssn = ssnele['ssn']
        print('SSN: ' + ssn)

        handle = None
        while (handle is None):
            print('Waiting for the CHASSIS window...')        
            handle = w.find_window_wildcard(".*CHASSIS.*")        
            sleep(1.0)            

        print('Found the window with handler ' + str(handle))
        w.set_foreground()

        sleep(500.0/1000.0)
        typewritestrs(ssn)

        handle = None
        while (handle is None):
            print('Waiting for the CSN window...')        
            handle = w.find_window_wildcard(".*Scan Part.*")
            sleep(0.5)
            pastseconds = pastseconds + 0.5
            if (pastseconds >= MAX_WAIT_SSN):        
                break

        if (handle is None):
            print('Skip this SSN. Trying next one...')
            continue

        print('Found the CSN window with handler ' + str(handle))
        for csnele in ssnele['csns']:
            csn = csnele['csn']
            print('CSN: ' + csn)
            scancsn(csn) 
            
        pastseconds = 0
        while (handle is not None):
            print('Waiting for quiting the CSN window...')
            handle = w.find_window_wildcard(".*Scan Part.*")
            sleep(1.0)
            pastseconds = pastseconds + 1
            if (pastseconds >= MAX_WAIT_CSN and handle is not None):        
                print('Timeout. Force to quit this form.')
                win32gui.PostMessage(handle,win32con.WM_CLOSE,0,0)
                break
    print('*************DONE*************')


def test():
    w = WindowMgr()    
    handle = None
    while (handle is None):
        print('Waiting for the CHASSIS window...')        
        handle = w.find_window_wildcard(".*CHASSIS.*")        
        sleep(1.0)
    
    print('Found the window with handler ' + str(handle))
    w.set_foreground()

    sleep(500.0/1000.0)
    typewritestrs('2UA9041YPW')    

    handle = None
    while (handle is None):
        print('Waiting for the CSN window...')        
        handle = w.find_window_wildcard(".*Scan Part.*")        
        sleep(0.5)

    print('Found the CSN window with handler ' + str(handle))
    scancsn('447087-001')
    scancsn('0FZJR005YB5157')
    scancsn('907322-016')
    scancsn('937045-001')
    scancsn('0GVLV8078A3C63')
    scancsn('0HBKK16AB7E941')    

def closewin():
    w = WindowMgr()    
    handle = None
    while (handle is None):
        print('Waiting for the CHASSIS window...')        
        handle = w.find_window_wildcard(".*CHASSIS.*")        
        sleep(1.0)
    
    print('Found the window with handler ' + str(handle))
    w.set_foreground()
    sleep(1.0)
    
    print('Closing ' + str(handle))
    win32gui.PostMessage(handle,win32con.WM_CLOSE,0,0)

#main()
#closewin()
test1()
