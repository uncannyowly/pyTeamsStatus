import os
import re
import requests
import configparser
import time
from datetime import datetime
import logging
from logging.handlers import RotatingFileHandler, TimedRotatingFileHandler


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
    Main loop that continuously checks for the latest Teams log file and processes it.
    """
    while True:
        log_dir = os.path.expandvars(settings['logfolder'])
        if not os.path.exists(log_dir):
            print("Log directory does not exist, check configuration.")
            break

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

        time.sleep(60)  # Wait for 60 seconds before checking for new log files again



def infer_call_status_from_line(line):
    """
    Infers the call status from a single line of log text.
    
    Parameters:
    - line: A string representing a single line of log text.
    
    Returns:
    - call_status: A string indicating the call status inferred from the line.
    """
    search_pattern_start = re.compile(r'WebViewWindowWin:.*tags=Call.*Window previously was visible = false')
    search_pattern_end = re.compile(r'BluetoothRadioManager: Device watcher is Started.')
    
    if search_pattern_start.search(line):
        return "In a call"
    elif search_pattern_end.search(line):
        return "Not in a call"
    
    # Return None if no call status could be inferred from the line
    return None



def process_log_file(file_path, settings):
    """
    Processes the latest log file to extract availability and call status.
    """
    search_pattern = re.compile(r'availability: (\w+), unread notification count: (\d+)')
    last_availability = None
    last_notification_count = None
    call_status = None
    
    try:
        with open(file_path, 'r', encoding='utf-8') as file:
            for line in file:
                # Check for availability and notification count
                match = search_pattern.search(line)
                if match:
                    last_availability, last_notification_count = match.groups()
                # Infer call status from log line
                inferred_status = infer_call_status_from_log(line)
                if inferred_status:
                    call_status = inferred_status
                    
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
    # Logic to update Home Assistant based on availability and call status
    logging.info(f"Updating Home Assistant: Availability={availability}, Call Status={call_status}, Notification Count={notification_count}")
    entity_id_status = settings['entities']['entitystatus']
    entity_id_activity = settings['entities']['entityactivity']  # Assuming you want to update this based on call status
    
    state = settings['language'].get(availability.lower(), "unknown")  # Ensure "available" is mapped correctly
    icon_key = 'inacall' if "call" in state.lower() else 'notinacall'
    icon = settings['icons'].get(icon_key, "mdi:account")  # Ensure default icon is set

    # Define attributes for the status entity
    attributes_status = {
        "friendly_name": settings['entities']['entitystatusname'],
        "icon": settings['icons'].get(state.lower(), "mdi:account")  # Use a specific icon for each status
    }

    # Optionally define attributes for the activity entity
    attributes_activity = {
        "friendly_name": settings['entities']['entityactivityname'],
        "icon": icon  # This could be based on whether in a call or not
    }

    # Update Home Assistant for both status and activity
    print(f"Updating Availability: {availability}, Notification Count: {notification_count}")
    send_to_home_assistant(settings['HAurl'], settings['token'], entity_id_status, state, attributes_status)
    send_to_home_assistant(settings['HAurl'], settings['token'], entity_id_activity, "in a call" if "inacall" in icon_key else "not in a call", attributes_activity)

if __name__ == "__main__":
    settings = load_settings('MSTeamsSettings.config')
    configure_logging(settings)

    # Debug print to verify entities loading
    print(settings['entities'])

    main_loop(settings)
