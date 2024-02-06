import os
import re
import requests
import configparser
import time
from datetime import datetime
import logging
from logging.handlers import RotatingFileHandler, TimedRotatingFileHandler
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

class LogFileHandler(FileSystemEventHandler):
    def __init__(self, settings):
        self.settings = settings

    def on_modified(self, event):
        if not event.is_directory:
            log_dir = os.path.expandvars(self.settings['logfolder'])
            file_pattern = re.compile(r'MSTeams_(\d{4}-\d{2}-\d{2})_(\d{2}-\d{2}-\d{2}\.\d+)\.log')
            if file_pattern.match(os.path.basename(event.src_path)):
                process_log_file(event.src_path, self.settings)

def load_settings(config_file):
    """
    Loads configuration settings from a given file.
    """
    config = configparser.ConfigParser()
    config.read(config_file)
    return {
        'HAurl': config['WebhookSettings']['HAurl'],
        'token': config['WebhookSettings']['token'],
        'logfolder': config['WebhookSettings']['logfolder'].replace('%%', '%'),
        'language': {k: v for k, v in config['LanguageSettings'].items()},
        'icons': {k: v for k, v in config['IconSettings'].items()},
        'entities': {k: v for k, v in config['EntitySettings'].items()},
        'debug': {k: v for k, v in config['DebugSettings'].items()}        
    }

def configure_logging(settings):
    debug_settings = settings.get('debug', {})
    if debug_settings.get('enabled', 'no').lower() == 'yes':
        log_file_path = debug_settings.get('log_file_path', 'debug.log')
        max_size_mb = int(debug_settings.get('max_size_mb', 5))
        backup_count = int(debug_settings.get('backup_count', 3))
        rotate_interval_hours = int(debug_settings.get('rotate_interval_hours', 3))

        # Set up basic configuration for the logging system
        logging.basicConfig(level=logging.DEBUG,
                            format='%(asctime)s - %(levelname)s - %(message)s',
                            datefmt='%Y-%m-%d %H:%M:%S')

        # Use a rotating file handler or timed rotating file handler based on config
        if max_size_mb > 0:
            handler = RotatingFileHandler(log_file_path, maxBytes=max_size_mb*1024*1024, backupCount=backup_count)
        else:
            handler = TimedRotatingFileHandler(log_file_path, when="h", interval=rotate_interval_hours, backupCount=backup_count)

        logging.getLogger('').addHandler(handler)
    else:
        # Disable logging if not enabled
        logging.disable(logging.CRITICAL)


def start_log_read(settings):
        log_dir = os.path.expandvars(settings['logfolder'])
        if not os.path.exists(log_dir):
            print("Log directory does not exist, check configuration.")
            exit 

        file_pattern = re.compile(r'MSTeams_(\d{4}-\d{2}-\d{2})_(\d{2}-\d{2}-\d{2}\.\d+)\.log')
        latest_file = None
        latest_time = None

        for file_name in os.listdir(log_dir):
            match = file_pattern.match(file_name)
            if match:
                file_datetime = datetime.strptime(f"{match.group(1)} {match.group(2)}", '%Y-%m-%d %H-%M-%S.%f')
                if latest_time is None or file_datetime > latest_time:
                    latest_time = file_datetime
                    latest_file = file_name

        if latest_file:
            process_log_file(os.path.join(log_dir, latest_file), settings)


def send_to_home_assistant(base_url, token, entity_id, state, attributes):
    """
    Sends a request to Home Assistant to update an entity's state and attributes.
    Includes detailed logging of the request and response.
    """
    url = f"{base_url}/api/states/{entity_id}"
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    payload = {"state": state, "attributes": attributes}
    try:
        response = requests.post(url, json=payload, headers=headers)
        response.raise_for_status()  # This will throw an error for 4XX/5XX responses
        logging.info(f"Successfully updated {entity_id} in Home Assistant to '{state}' with attributes {attributes}.")
    except requests.RequestException as e:
        logging.error(f"Error updating {entity_id} in Home Assistant: {e}. Payload: {payload}")

def main_loop(settings):
    """
    Original: Main loop that continuously checks for the latest Teams log file and processes it.
    This update replaces the original main loop with an event-driven approach using watchdog.
    """
    path = os.path.expandvars(settings['logfolder'])
    if not os.path.exists(path):
        print("Log directory does not exist, check configuration.")
        return

    event_handler = LogFileHandler(settings)
    observer = Observer()
    observer.schedule(event_handler, path, recursive=False)
    observer.start()
    try:
        while True:
            time.sleep(10)  # Sleep time reduced to maintain responsiveness without busy waiting
    except KeyboardInterrupt:
        observer.stop()
    observer.join()



def infer_call_status_from_line(line):
    """
    Infers the call status from a single line of log text.
    """
    search_pattern_start = re.compile(r'WebViewWindowWin:.*tags=Call.*Window previously was visible = false')
    search_pattern_end = re.compile(r'BluetoothRadioManager: Device watcher is Started.')
    
    if search_pattern_start.search(line):
        return "In a call"
    elif search_pattern_end.search(line):
        return "Not in a call"
    
    return None



def process_log_file(file_path, settings):
    """
    Processes the latest log file to extract availability, call status, and notification count.
    """
    search_pattern = re.compile(r'availability: (\w+), unread notification count: (\d+)')
    call_status = "Not in a call"  # Default call status
    
    try:
        with open(file_path, 'r', encoding='utf-8') as file:
            for line in file:
                match = search_pattern.search(line)
                if match:
                    last_availability, last_notification_count = match.groups()
                
                inferred_status = infer_call_status_from_line(line)
                if inferred_status:
                    call_status = inferred_status  # Update call status based on the log line

        # Update Home Assistant with the last found availability, notification count, and call status
        if last_availability and last_notification_count:
            logging.info(f"Detected status: {last_availability}, Call Status: {call_status}, Notification Count: {last_notification_count}")
            update_home_assistant(last_availability, last_notification_count, call_status, settings)
    except Exception as e:
        logging.error(f"Error reading file {file_path}: {e}")



# Adjust `update_home_assistant` to handle call status
def update_home_assistant(availability, notification_count, call_status, settings):
    """
    Updates Home Assistant entities based on Teams availability, call status, and notification count.
    """
    logging.info(f"Updating Home Assistant: Availability={availability}, Call Status={call_status}, Notification Count={notification_count}")
    
    entity_id_status = settings['entities']['entitystatus']
    entity_id_activity = settings['entities']['entityactivity']  # Assuming you want to update this based on call status
    
    state = settings['language'].get(availability.lower(), "unknown")  # Ensure "available" is mapped correctly
    icon_key = 'inacall' if call_status == "In a call" else 'notinacall'
    icon = settings['icons'].get(icon_key, "mdi:account")  # Ensure default icon is set

    # Define attributes for the status entity
    attributes_status = {
        "friendly_name": settings['entities']['entitystatusname'],
        "icon": settings['icons'].get(state.lower(), "mdi:account")  # Use a specific icon for each status
    }

    # Define attributes for the activity entity based on call status
    activity_state = "in a call" if call_status == "In a call" else "not in a call"
    attributes_activity = {
        "friendly_name": settings['entities']['entityactivityname'],
        "icon": icon
    }

    print(f"Updating Availability: {availability}, Notification Count: {notification_count}, Call Status: {activity_state}")
    send_to_home_assistant(settings['HAurl'], settings['token'], entity_id_status, state, attributes_status)
    send_to_home_assistant(settings['HAurl'], settings['token'], entity_id_activity, activity_state, attributes_activity)

if __name__ == "__main__":
    settings = load_settings('MSTeamsSettings.config')
    start_log_read(settings)
    configure_logging(settings)
    main_loop(settings)