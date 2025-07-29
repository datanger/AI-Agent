import os
import sys
import subprocess
import time
import logging
import uno
from uno.com.sun.star.util import XModifyListener
from uno.com.sun.star.lang import XEventListener
import atexit

# --- Configuration ---
FILE_TO_MONITOR = "/home/niejie/work/Code/tools/AI-Agent/test.xlsx"
LOG_FILE = "excel_monitor.log"
LISTENER_PORT = "2002"

# --- Setup Logging ---
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(LOG_FILE),
        logging.StreamHandler()
    ]
)

# --- Global variable for the LibreOffice process ---
soffice_process = None

def cleanup_soffice_process():
    """Ensures the LibreOffice process is terminated on script exit."""
    global soffice_process
    if soffice_process:
        logging.info("Terminating LibreOffice process...")
        soffice_process.terminate()
        soffice_process.wait()
        logging.info("LibreOffice process terminated.")

atexit.register(cleanup_soffice_process)


class SheetModifyListener(XModifyListener, XEventListener):
    """
    A listener that gets notified of sheet modifications.
    It calls a callback function when a modification is detected.
    """
    def __init__(self, sheet_name, callback):
        self.sheet_name = sheet_name
        self.callback = callback

    def modified(self, event):
        """Callback method called by LibreOffice when the sheet is modified."""
        logging.info(f"Modification detected in sheet: '{self.sheet_name}'. Checking for changes...")
        self.callback(self.sheet_name)

    def disposing(self, event):
        """Callback method called when the listener is being disposed."""
        logging.info(f"Listener for sheet '{self.sheet_name}' is being disposed.")


class RealTimeMonitor:
    def __init__(self, file_path):
        self.file_path = os.path.abspath(file_path)
        self.desktop = None
        self.document = None
        self.listeners = []
        self.sheet_states = {}

    def start_libreoffice(self):
        """Starts a headless LibreOffice instance listening for connections."""
        global soffice_process
        if not os.path.exists(self.file_path):
            logging.error(f"File not found: {self.file_path}. Please create it first.")
            sys.exit(1)
            
        logging.info("Starting headless LibreOffice process...")
        command = [
            "soffice",
            "--headless",
            "--invisible",
            f"--accept=socket,host=localhost,port={LISTENER_PORT};urp;"
        ]
        soffice_process = subprocess.Popen(command)
        logging.info(f"LibreOffice process started with PID: {soffice_process.pid}. Waiting for it to initialize...")
        time.sleep(10) # Give LibreOffice time to start up

    def connect(self):
        """Connects to the running LibreOffice instance."""
        logging.info("Connecting to LibreOffice...")
        try:
            local_context = uno.getComponentContext()
            resolver = local_context.ServiceManager.createInstanceWithContext(
                "com.sun.star.bridge.UnoUrlResolver", local_context
            )
            context = resolver.resolve(
                f"uno:socket,host=localhost,port={LISTENER_PORT};urp;StarOffice.ComponentContext"
            )
            self.desktop = context.ServiceManager.createInstanceWithContext(
                "com.sun.star.frame.Desktop", context
            )
            logging.info("Successfully connected to LibreOffice.")
        except Exception as e:
            logging.error(f"Failed to connect to LibreOffice: {e}")
            logging.error("Is LibreOffice running? Is another instance using the same port?")
            sys.exit(1)

    def load_document(self):
        """Loads the spreadsheet document."""
        url = uno.systemPathToFileUrl(self.file_path)
        logging.info(f"Loading document: {url}")
        
        # Properties to open the document as read-write
        props = (
            uno.createUnoStruct('com.sun.star.beans.PropertyValue', {'Name': 'ReadOnly', 'Value': False}),
        )
        
        self.document = self.desktop.loadComponentFromURL(url, "_blank", 0, props)
        if not self.document:
            logging.error(f"Failed to load document: {self.file_path}")
            sys.exit(1)
        logging.info("Document loaded successfully.")

    def get_sheet_state(self, sheet):
        """Captures the current state of all non-empty cells in a sheet."""
        state = {}
        cursor = sheet.createCursor()
        cursor.gotoEndOfUsedArea(False)
        last_col = cursor.getRangeAddress().EndColumn
        last_row = cursor.getRangeAddress().EndRow

        for r in range(last_row + 1):
            for c in range(last_col + 1):
                cell = sheet.getCellByPosition(c, r)
                if cell.String: # Using .String gets the displayed text
                    state[cell.AbsoluteName] = cell.String
        return state

    def compare_and_log(self, sheet_name):
        """Compares the current sheet state with the last known state and logs changes."""
        sheet = self.document.Sheets.getByName(sheet_name)
        new_state = self.get_sheet_state(sheet)
        old_state = self.sheet_states.get(sheet_name, {})
        
        base_name = os.path.basename(self.file_path)
        all_cells = set(old_state.keys()) | set(new_state.keys())

        for cell_addr in sorted(list(all_cells)):
            old_value = old_state.get(cell_addr)
            new_value = new_state.get(cell_addr)

            if old_value != new_value:
                if old_value is None:
                    logging.info(f"[{base_name}][{sheet_name}] Event: Cell created {cell_addr} = '{new_value}'")
                elif new_value is None:
                    logging.info(f"[{base_name}][{sheet_name}] Event: Cell deleted {cell_addr} (was '{old_value}')")
                else:
                    logging.info(f"[{base_name}][{sheet_name}] Event: Cell updated {cell_addr} from '{old_value}' to '{new_value}'")
        
        # Update the state for the next comparison
        self.sheet_states[sheet_name] = new_state

    def run(self):
        """Main method to start monitoring."""
        self.start_libreoffice()
        self.connect()
        self.load_document()

        sheets = self.document.Sheets
        for sheet_name in sheets.getElementNames():
            logging.info(f"Setting up monitoring for sheet: '{sheet_name}'")
            # Get initial state
            sheet = sheets.getByName(sheet_name)
            self.sheet_states[sheet_name] = self.get_sheet_state(sheet)
            logging.info(f"Initial state captured for '{sheet_name}' with {len(self.sheet_states[sheet_name])} cells.")

            # Create and attach listener
            listener = SheetModifyListener(sheet_name, self.compare_and_log)
            self.listeners.append(listener)
            sheet.addModifyListener(listener)

        logging.info("\n*** Real-time monitoring active. Press Ctrl+C to stop. ***\n")
        logging.info("*** Open your file in LibreOffice Calc to see live changes. ***\n")
        
        try:
            while True:
                time.sleep(1)
        except KeyboardInterrupt:
            logging.info("Ctrl+C received. Shutting down.")
        finally:
            self.shutdown()

    def shutdown(self):
        """Cleans up resources."""
        if self.document:
            logging.info("Removing listeners...")
            for listener in self.listeners:
                # The listener needs to be attached to the object it's listening to for removal
                sheet = self.document.Sheets.getByName(listener.sheet_name)
                sheet.removeModifyListener(listener)
            
            logging.info("Closing document...")
            self.document.close(True)
        
        # The atexit handler will terminate the soffice process

if __name__ == "__main__":
    print("=== LibreOffice Real-Time Excel Monitor ===")
    print(f"Monitoring file: {FILE_TO_MONITOR}")
    print("This script will start a background LibreOffice process.")
    print("Please open the file in your normal LibreOffice Calc window to edit.")
    print("Press Ctrl+C in this terminal to stop the monitor.")
    print("----------------------------------------------------")
    
    monitor = RealTimeMonitor(FILE_TO_MONITOR)
    monitor.run()
