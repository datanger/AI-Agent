import os
import sys
import logging
import win32com.client
import pythoncom
import time

# --- Configuration ---
# 重要：请使用反斜杠'\\'或正斜杠'/'作为路径分隔符
FILE_TO_MONITOR = "test.xlsx" # <--- 请务必修改为你的实际文件路径
LOG_FILE = "excel_monitor.log"

# --- Setup Logging ---
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(LOG_FILE),
        logging.StreamHandler()
    ]
)


class WpsEventHandler:
    """
    Handles events fired by the WPS Application.
    The method name OnSheetChange is specific and defined by the WPS COM API,
    which is compatible with Excel's API.
    """
    def __init__(self):
        self.monitor_callback = None

    def OnSheetChange(self, sheet, target):
        """
        This method is called by WPS itself whenever a cell changes.
        'sheet' is the sheet object where the change occurred.
        'target' is the range object that was changed.
        """
        try:
            sheet_name = sheet.Name
            logging.info(f"Modification detected in WPS sheet: '{sheet_name}'. Range: {target.Address}")
            if self.monitor_callback:
                self.monitor_callback(sheet_name, target)
        except Exception as e:
            logging.error(f"Error in OnSheetChange event handler: {e}")

class RealTimeMonitor:
    def __init__(self, file_path):
        self.file_path = os.path.abspath(file_path)
        self.wps_app = None
        self.workbook = None
        self.sheet_states = {}

    def connect(self):
        """Connects to an existing WPS instance or starts a new one."""
        logging.info("Connecting to WPS Spreadsheets (ET)...")
        try:
            # The ProgID for WPS Spreadsheets is "et.Application"
            self.wps_app = win32com.client.DispatchWithEvents("et.Application", WpsEventHandler)
        except Exception as e:
            logging.error(f"Failed to connect to WPS. Is it installed correctly? Error: {e}")
            sys.exit(1)

        self.wps_app.Visible = True # Make WPS visible to the user
        self.wps_app.DisplayAlerts = False

    def load_document(self):
        """Opens the specified workbook."""
        try:
            self.workbook = self.wps_app.Workbooks.Open(self.file_path)
            logging.info(f"Successfully opened workbook in WPS: {os.path.basename(self.file_path)}")
        except Exception as e:
            logging.error(f"Failed to open workbook: {e}")
            logging.error("Please ensure the file exists and is not corrupted.")
            self.shutdown()
            sys.exit(1)

    def get_sheet_state(self, sheet):
        """Captures the current state of all non-empty cells in a sheet."""
        state = {}
        last_row = sheet.Cells(sheet.Rows.Count, 1).End(3).Row # 3 corresponds to xlUp in WPS
        last_col = sheet.Cells(1, sheet.Columns.Count).End(1).Column # 1 corresponds to xlToLeft in WPS

        for r in range(1, last_row + 1):
            for c in range(1, last_col + 1):
                cell = sheet.Cells(r, c)
                if cell.Value is not None:
                    state[cell.Address] = str(cell.Value)
        return state

    def compare_and_log(self, sheet_name, target_range):
        """Compares the changed cells and logs the differences."""
        sheet = self.workbook.Sheets(sheet_name)
        new_state = self.get_sheet_state(sheet) # Get the full new state
        old_state = self.sheet_states.get(sheet_name, {})
        base_name = os.path.basename(self.file_path)

        for cell in target_range:
            cell_addr = cell.Address
            old_value = old_state.get(cell_addr)
            new_value = str(cell.Value) if cell.Value is not None else None

            if old_value != new_value:
                if old_value is None and new_value is not None:
                    logging.info(f"[{base_name}][{sheet_name}] Event: Cell created {cell_addr} = '{new_value}'")
                elif new_value is None and old_value is not None:
                    logging.info(f"[{base_name}][{sheet_name}] Event: Cell deleted {cell_addr} (was '{old_value}')")
                else:
                    logging.info(f"[{base_name}][{sheet_name}] Event: Cell updated {cell_addr} from '{old_value}' to '{new_value}'")
        
        self.sheet_states[sheet_name] = new_state

    def run(self):
        """Main method to start monitoring."""
        if not os.path.exists(self.file_path):
            logging.error(f"File not found: {self.file_path}")
            sys.exit(1)

        self.connect()
        self.load_document()

        self.wps_app.Events.monitor_callback = self.compare_and_log

        for sheet in self.workbook.Sheets:
            logging.info(f"Capturing initial state for sheet: '{sheet.Name}'")
            self.sheet_states[sheet.Name] = self.get_sheet_state(sheet)
            logging.info(f"Initial state captured for '{sheet.Name}' with {len(self.sheet_states[sheet.Name])} cells.")

        logging.info("\n*** WPS Real-time monitoring active. Press Ctrl+C in this terminal to stop. ***")
        logging.info("*** Edit the file in the WPS window to see live changes. ***\n")

        pythoncom.PumpMessages()

    def shutdown(self):
        """Cleans up resources."""
        if self.workbook:
            self.workbook.Close(SaveChanges=False)
        if self.wps_app:
            self.wps_app.Quit()
        logging.info("WPS monitor shut down.")

if __name__ == "__main__":
    print("=== WPS Real-Time Monitor for Windows ===")
    print(f"Monitoring file: {FILE_TO_MONITOR}")
    print("This script will connect to or start a WPS Spreadsheets application.")
    print("Press Ctrl+C in this terminal to stop the monitor.")
    print("----------------------------------------------------")
    
    monitor = RealTimeMonitor(FILE_TO_MONITOR)
    try:
        monitor.run()
    except KeyboardInterrupt:
        logging.info("Ctrl+C received. Shutting down.")
    finally:
        monitor.shutdown()
