import os
import signal
import sys
import threading
import time
import traceback
from datetime import datetime

from win32com.client import constants
from win32com.client.gencache import EnsureDispatch


def getdocsfolder():
    # Gets local user document folder and appends 'Autosaves'
    oshell = EnsureDispatch("Wscript.Shell")

    docs = oshell.SpecialFolders("MyDocuments")
    directory = os.path.join(docs, "autosaves")
    os.makedirs(directory, exist_ok=True)

    return directory


def clearoutput():
    os.system('cls')
    print('Press CTRL-C to exit or change autosave interval')


def savecurrentopen(savedirectory):
    # main function to save currently open pdfs

    acrobat = EnsureDispatch('AcroExch.App')  # access acrobat COM server
    num = acrobat.GetNumAVDocs()  # Get number of open PDFs
    i = 0
    now = datetime.now()
    timestr = now.strftime("On %m/%d/%Y, at %H:%M:%S")
    filelist = []
    print(timestr)
    while i < num:
        doc = acrobat.GetAVDoc(i)  # gets acrobats open windows
        pd = doc.GetPDDoc()  # gets underlying pdfs
        name = pd.GetFileName()
        if name in filelist: #bruteforce for saving docs with identical file names. need to update.
            now = datetime.now()
            name = name[:(len(name) - 4)] + now.strftime('%H-%M-%S.pdf')

        filelist.append(name)
        time.sleep(1)
        pd.Save(constants.PDSaveCopy | constants.PDSaveFull, os.path.join(savedirectory, name))
        print("Saved " + str(os.path.join(savedirectory, name)))
        i += 1


class SignalHandler:
    def __init__(self):
        self.event = threading.Event()

    def sig_handler(self, signal, frame):
        self.event.set() #break old loop
        response = input('Enter Autosave interval in seconds or press enter to exit: ')
        if response.isnumeric():
            newinterval = int(response)
            print('Changed interval to ' + response)
            cleanold()
            self.event.clear()
            while not s.event.isSet(): #new loop with user input interval
                mainloop()
                cleanold(age=newinterval*5)
                print('')
                print(str(int(round(newinterval / 60))) + ' minute(s) until next save', end='\r')
                timer = 0
                while timer < newinterval:
                    s.event.wait(1)
                    timer += 1
                    if (newinterval - timer) % 60 == 0:
                        print(str(int((newinterval - timer) / 60)) + ' minute(s) until next save    ', end='\r')


def mainloop():
    directory = getdocsfolder()
    clearoutput()
    savecurrentopen(directory)


def cleanold(age=3000):
    directory = getdocsfolder()
    onlyfiles = [f for f in os.listdir(directory) if os.path.isfile(os.path.join(directory, f))]
    for file in onlyfiles:
        path = os.path.join(directory, file)
        x = os.stat(path)
        Result = (time.time() - x.st_mtime)
        if Result > age:
            os.remove(path)


def show_exception(exc_type, exc_value, tb):
    traceback.print_exception(exc_type, exc_value, tb)
    input("Press key to exit.")
    sys.exit(-1)


if __name__ == '__main__':
    s = SignalHandler()
    signal.signal(signal.SIGINT, s.sig_handler)
    sys.excepthook = show_exception
    while not s.event.isSet():
        mainloop()
        cleanold()
        timer = 0
        print('')
        print('10 minute(s) until next save', end='\r')
        while timer < 600:
            s.event.wait(1)
            timer += 1
            if timer % 60 == 0:
                print(str(int((600 - timer) / 60)) + ' minute(s) until next save    ', end='\r')
    os.system('cls')
    print("Exiting...")
    time.sleep(5)
