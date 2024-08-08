import sys
import threading
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import pystray
from pystray import MenuItem as item
from PIL import Image, ImageDraw
import tkinter as tk
import tkinter.simpledialog
from tkinter import scrolledtext
import datetime
from office365.runtime.auth.user_credential import UserCredential
from office365.sharepoint.client_context import ClientContext
import os
import os.path
import time
import setproctitle

# Set the process title
setproctitle.setproctitle("Monitoring Script")

# SharePoint credentials
sharepoint_url = 'https://compeltechnology.sharepoint.com/sites/shared'
user_credentials = UserCredential('sharepoint@compel.ws', 'CTI$h@r3!')

print("Hi")
class FileCreatedHandler(FileSystemEventHandler):
    def __init__(self, filenames):
        super().__init__()
        self.filenames = filenames
        self.log = []

    def on_created(self, event):
        for filename in self.filenames:
            print(event.src_path, filename, event.event_type )
            if event.src_path.endswith(filename):
                log_entry = f"{filename} was {event.event_type} at {event.src_path}"
                self.log.append(log_entry)
                upload_to_onedrive(event.src_path, self, filename)
                break


def create_image():
    script_dir = os.path.dirname(os.path.abspath(__file__))
    image_path = os.path.join(script_dir, 'cti_tray_64.png')
   
    try:
        image = Image.open(image_path)
        image = image.convert('RGB')  # Ensure image is in RGB mode
        image = image.resize((64, 64))  # Resize image to fit tray icon size
    except IOError as e:
        print(f"Unable to load image file: {e}")
        # In case of failure, fallback to creating a new image
        image = Image.new('RGB', (64, 64), 'black')
   

    return image

def check_logs(handler):
    # Create a new window for displaying logs
    log_window = tk.Toplevel()
    log_window.title("Log Entries")
    log_window.geometry("1000x800") #Set width = 1000 height = 6--
    # Create a scrolled text widget to display logs
    log_text = scrolledtext.ScrolledText(log_window, width=80, height=20)
    log_text.pack(fill="both", expand=True)

    # Append each log entry to the scrolled text widget
    for entry in handler.log:
        log_text.insert(tk.END, entry + "\n")
    
    # Disable editing in the scrolled text widget
    log_text.configure(state='disabled')

    # Function to close the log window
    def close_window():
        log_window.destroy()

    # Create a close button to close the log window
    close_button = tk.Button(log_window, text="Close", command=close_window)
    close_button.pack(pady=10)

    # Bring the log window to the front
    log_window.focus_force()
    log_window.grab_set()

    # Run the Tkinter main loop for the log window
    log_window.wait_window()


def upload_to_onedrive(file_path, handler, filename):
    ctx = ClientContext(sharepoint_url).with_credentials(user_credentials)
    list_title = 'Shared Documents/Order_Processing'
    target_folder = ctx.web.get_folder_by_server_relative_url(list_title)
    name = os.path.basename(file_path)
    # Set the number of retries to try opening the file
    retries = 3 
    delay=1

    os.chdir("G:\\python")
    print("pid", os.getegid())
    print(os.stat(filename))
    print(os.getcwd())
    print(os.listdir())
    print(os.path.isfile(filename))
    print(filename)
    print(list(filename))
    os.remove(filename)
    for attempt in range(retries):
        try:
            print("try #", attempt)
            with open(os.path.join(os.getcwd(), filename), 'rb') as content_file:
                file_content = content_file.read()
            break
        except PermissionError as error:
            print(f"Attemt {attempt+1}: Permission denied. Retrying...")
            time.sleep(delay)
        except BaseException as error:
            print('An exception occurred:', error)
    else:
        raise Exception("Failed to open file after multiple attempts.")
    
    print("success!")
    target_file = target_folder.upload_file(name, file_content).execute_query()
    log_entry = f"{datetime.datetime.now()}: {name} was uploaded to OneDrive"
    handler.log.append(log_entry)  # Log the upload
    #print(f"File has been uploaded to OneDrive: {target_file.serverRelativeUrl}")

    # Verify that the file exists before deleting
    if os.path.exists(file_path):
        os.remove(file_path)
        log_entry = f"{datetime.datetime.now()}: {name} was deleted after uploading"
        handler.log.append(log_entry)
        
    else:
        log_entry = f"{datetime.datetime.now()}: {name} does not exist after uploading"
        handler.log.append(log_entry)
        


def main_invoices():
    monitored_files = ["PQUOTE.txt", "PINVOICE.txt", "PINSTALL.txt"]
    handler = FileCreatedHandler(monitored_files)
    observer = Observer()
    observer.schedule(handler, path='G:\\Python', recursive=False)  # Monitor the parent directory
    observer.start()

    def on_exit(icon, item):
        observer.stop()
        observer.join()
        icon.stop()

    def on_check_logs(icon, item):
        check_logs(handler)

    icon_image = create_image()
    menu = (item('Check Logs', on_check_logs), item('Exit', on_exit))
    icon = pystray.Icon("file_monitor", icon_image, "Compel Utility", menu)
    icon.run()
    

if __name__ == "__main__":
    main_invoices()