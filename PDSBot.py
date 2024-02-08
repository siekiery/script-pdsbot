import pyautogui as pag
#import pywinauto as pwa
import time, os, sys, warnings
from datetime import datetime

__VERSION__ = '2.1.1'
DEVMODE = False

pag.PAUSE = 1.0

WELCOME = f"""

PDS Bot
v{__VERSION__}

Automated bulk PDS converter.
Developer: Jakub Pitera
__________________________________________________

In order to convert make sure:
 - "Schlumberger PDSView" is installed and running.
 - in 'File' > 'Options', "Display message box after PDS file has been loaded" is ticked.
 
Specific for 'Print to PDF' conversion:
 - in 'File' > 'Print Setup', "Microsoft Print to PDF" is selected
 - in 'File' > 'Print', "Show print status" is unticked

In PDSView select desired DPI of .TIF and any other PDF printing options manually before converting!

There are two ways of converting to PDF: 'File > Save as...' and 'File > Print'. The latter is using 'Microsoft Print to PDF'.
Use it if the 'File' > 'Save  as...' > 'Save as type:' > 'PDF Files (*.PDF)' is unavailable in PDSView.

To start bulk conversion, please first select MODE using the prompt below. 
Secondly, enter the directory of .PDS files.
All the .PDS files inside that directory tree will be converted (including the subfolders).

Important note: This program will take control of the mouse and keyboard until the conversion is finished. 
To interrupt, please hold "CTRL"+"C" while this window is active. 
You can return to the interrupted conversion by selecting "4 CONTINUE from saved log" MODE and then providing 
a full path of a saved log created by POT Bot.

__________________________________________________
"""

MODE_MSG = """
Select MODE:
1   Save as TIF
2   Save as PDF
3   Print to PDF
4   CONTINUE from saved log
"""


class PDSBot:
    def __init__(self):
        self.pds_files = []
        self.pdsview = None

        print(WELCOME)

        self.__main__()

    def __main__(self):
        self.mode = self.ask_mode()

        if self.mode in (1, 2, 3):
            self.directory = self.ask_directory()
            self.pds_files = self.read_pds_files(self.directory)

        elif self.mode == 4:
            self.logpath = self.ask_logpath()
            self.mode, self.directory, self.pds_files = self.read_log(self.logpath)

        print(f"\nTotal of {len(self.pds_files)} .PDS files to be converted across directory tree.\n")

        self.pdsview = self.prepare_pdsview()

        self.starttime = datetime.now()

        self.log = self.start_log()

        self.count = 0

        self.bulk_convert()

        self.endtime = datetime.now()

        self.runtime = self.timer(self.starttime, self.endtime)

        print(f"\nSuccessfully converted {self.count} .PDS files! It took {self.runtime}.")
        input()

    @staticmethod
    def ask_mode():
        mode = None
        while mode not in (1, 2, 3, 4):
            try:
                mode = int(input(MODE_MSG))
            except:
                continue
        return mode

    @staticmethod
    def ask_directory():
        directory = ''
        while not os.path.isdir(directory):
            directory = input("\nEnter directory: ")

        return directory

    @staticmethod
    def ask_logpath():
        logpath = ''
        while not (os.path.isfile(logpath) and logpath.upper().endswith('.TXT')):
            logpath = input("\nEnter saved log filepath: ")

        return logpath

    @staticmethod
    def read_pds_files(directory):
        # Lists absolute paths of all files to be converted
        pds_files = []
        for (root, _, files) in os.walk(directory):
            for name in files:
                if name.upper().endswith(('.PDS')):
                    pds_files.append(os.path.join(root, name))

        return pds_files

    def read_log(self, logpath):
        with open(logpath) as log:
            log_lines = log.readlines()

        log_lines = [line.strip() for line in log_lines]

        mode_str  = log_lines[1][6:]
        if mode_str == 'TIF':
            mode = 1
        elif mode_str == 'SAVE AS PDF':
            mode = 2
        elif mode_str == 'PRINT TO PDF':
            mode = 3

        directory = log_lines[2][11:]

        converted = log_lines[5:]

        pds_files = self.read_pds_files(directory)
        pds_files = [pds for pds in pds_files if pds not in converted]

        return mode, directory, pds_files


    @staticmethod
    def focus_on_window(window):
        """Deprecated as Windows is not communicating with pwa as expected"""
        if window.isActive == False:
            pwa.application.Application().connect(best_match=window.title).top_window().set_focus()
            print('\n')

    def prepare_pdsview(self):
        title = "Schlumberger PDSView"
        try:
            pdsview = pag.getWindowsWithTitle(title)[0]
        except:
            sys.exit(input("\nPlease launch 'Schlumberger PDSView' and retry."))

        # self.focus_on_window(pdsview)
        #pdsview.activate()
        pdsview.size = (500, 500)
        pdsview.topleft = (0, 0)
        pdsview.size = (1920, 1080)
        pdsview.maximize()
        time.sleep(1)
        pag.click(pdsview.left + 120, pdsview.top + 120)

        return pdsview

    @staticmethod
    def convert(filename, directory, mode):

        def load_pds(filename):
            assert filename.upper().endswith(('.PDS')), "%s is not a .PDS file. Aborting." % filename
            # Load .PDS
            time.sleep(0.5)
            pag.hotkey('ctrl', 'o')
            time.sleep(0.25)
            pag.getActiveWindow().size = (1920, 1080)
            time.sleep(1.5)
            pag.write(os.path.abspath(filename))
            pag.press('enter')
            while pag.getWindowsWithTitle('File Load') == []:
                time.sleep(0.5)
            time.sleep(0.5)
            pag.press('enter')

        def save_as_tif(filename, directory):
            pag.hotkey('alt', 'f')
            pag.press('down', presses=2)
            pag.press('enter')
            pag.press('tab')
            pag.press('down', presses=4)
            pag.press('enter', presses=2)
            while pag.getWindowsWithTitle('Confirm Save As') != []:
                time.sleep(1)
            while pag.getWindowsWithTitle('Please wait...') != []:
                time.sleep(1)
            relpath = os.path.relpath(filename, start=directory)
            print(f"'{relpath[:-4]}.TIF' created.")

        def save_as_pdf(filename, directory):
            pag.hotkey('alt', 'f')
            pag.press('down', presses=2)
            pag.press('enter')
            pag.sleep(1)
            pag.press('tab')
            pag.press('down', presses=5)
            pag.press('enter', presses=2)
            while pag.getWindowsWithTitle('Confirm Save As') != []:
                time.sleep(1)
            while pag.getWindowsWithTitle('Please wait...') != []:
                time.sleep(1)
            relpath = os.path.relpath(filename, start=directory)
            print(f"'{relpath[:-4]}.PDF' created.")


        def print_to_pdf(filename, directory):
            pag.hotkey('ctrl', 'p')
            pag.press('enter')
            while pag.getWindowsWithTitle('Save Print Output As') == []:
                time.sleep(1)
            pag.write(os.path.abspath(filename[:-4]))
            pag.press('enter')
            while pag.getWindowsWithTitle('PDSView Print') == []:
                time.sleep(1)
            pag.press('enter')
            relpath = os.path.relpath(filename, start=directory)
            print(f"'{relpath[:-4]}.PDF' created.")

        load_pds(filename)
        if mode == 1:
            save_as_tif(filename, directory)
        elif mode == 2:
            save_as_pdf(filename, directory)
        elif mode == 3:
            print_to_pdf(filename, directory)

    def bulk_convert(self):
        for file in self.pds_files:
            self.convert(file, self.directory, self.mode)
            self.count += 1
            self.update_log(file)
            self.save_log()

    def start_log(self):

        if self.mode == 1:
            mode_str = "TIF"
        elif self.mode == 2:
            mode_str = "SAVE AS PDF"
        elif self.mode == 3:
            mode_str = "PRINT TO PDF"

        log = f"""DATETIME: {str(self.starttime)}
MODE: {mode_str}
DIRECTORY: {self.directory}
        
CONVERTED:
"""
        return log

    def update_log(self, filename):
        self.log += f"{filename}\n"

    def save_log(self):
        name = f"PDSBot_{self.starttime.strftime('%Y%m%d_%H%M%S')}.txt"
        logpath = os.path.join('Logs', name)
        with open(logpath, 'w') as file:
            file.write(self.log)

    @staticmethod
    def timer(starttime, endtime):
        diff = (endtime - starttime).total_seconds()
        hrs = diff // 3600
        mins = diff % 3600 // 60
        secs = round(diff % 3600 % 60)
        return "%i hr %i min %i s" % (hrs, mins, secs)

if __name__ == "__main__":
    PDSBot()
