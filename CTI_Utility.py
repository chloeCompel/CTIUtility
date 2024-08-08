import sys
import os
import time
import datetime
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
from office365.runtime.auth.user_credential import UserCredential
from office365.sharepoint.client_context import ClientContext
from PIL import Image
import pystray
from pystray import MenuItem as item
import getpass

# SharePoint credentials
sharepoint_url = 'https://compeltechnology.sharepoint.com/sites/shared'
user_credentials = UserCredential('sharepoint@compel.ws', 'CTI$h@r3!')

class FileCreatedHandler(FileSystemEventHandler):
    def __init__(self, filenames):
        super().__init__()
        self.filenames = filenames
        self.log = []

    def on_any_event(self, event):
        if event.event_type in ['created', 'modified', 'moved']:
            for filename in self.filenames:
                if event.src_path.endswith(filename):
                    log_entry = f"{filename} was {event.event_type} at {event.src_path}"
                    self.log.append(log_entry)
                    upload_to_onedrive(event.src_path, self)

def create_image():
    # Create an icon image for the system tray
    image = Image.new('RGB', (32, 32), 'blue')
    return image

def upload_to_onedrive(file_path, handler):
    ctx = ClientContext(sharepoint_url).with_credentials(user_credentials)
    list_title = 'Shared Documents/Order_Processing'
    target_folder = ctx.web.get_folder_by_server_relative_url(list_title)
    name = os.path.basename(file_path)
    # Create a text file with the current log
    print("first line")
    print(str(os.path.isfile(r"C:\Support\another_file.txt")))
    open(r"C:\Support\another_file.txt", 'rb')

    print("second line")
    with open(file_path, 'rb') as content_file:
        file_content = content_file.read()

    print("third line")
    target_file = target_folder.upload_file(name, file_content).execute_query()
    log_entry = f"{datetime.datetime.now()}: {name} was uploaded to OneDrive"
    handler.log.append(log_entry)

    if os.path.exists(file_path):
        os.remove(file_path)
        log_entry = f"{datetime.datetime.now()}: {name} was deleted after uploading"
        handler.log.append(log_entry)
    else:
        log_entry = f"{datetime.datetime.now()}: {name} does not exist after uploading"
        handler.log.append(log_entry)

    # Create a text file with the current log
    with open('monitor_log.txt', 'w') as log_file:
        for entry in handler.log:
            log_file.write(entry + "\n")

def main_invoices():
    monitored_files = ["PQUOTE.TXT", "PINVOICE.TXT", "PINSTALL.TXT"]
    handler = FileCreatedHandler(monitored_files)
    observer = Observer()
    observer.schedule(handler, path='g:\\python', recursive=False)
    observer.start()
    

    def on_exit(icon, item):
        observer.stop()
        observer.join()
        icon.stop()

    icon_image = create_image()
    menu = (item('Exit', on_exit),)
    icon = pystray.Icon("file_monitor", icon_image, "Compel Utility", menu)
    icon.run()

if __name__ == "__main__":
    main_invoices()
