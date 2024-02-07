"""
Microsoft Teams Status Monitor and Home Assistant Updater
Date: 2024-02-07

This script continuously monitors Microsoft Teams log files for changes in the user's
availability, notification count, and call status. When a change is detected, it updates
specific entities in Home Assistant to reflect the new status.
"""

import os
import re
import requests
import configparser
import time
from datetime import datetime
import logging
from logging.handlers import RotatingFileHandler, TimedRotatingFileHandler

# Global variables to store the last known status of the user
last_known_availability = "unknown"
last_known_notification_count = "unknown"
last_known_call_status = "unknown"

def load_settings(config_file):
    """
    Loads configuration settings from the specified configuration file.

    Args:
        config_file (str): Path to the configuration file.

    Returns:
        dict: A dictionary of configuration settings.
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
    """
    Configures the logging system based on settings from the configuration file.

    Args:
        settings (dict): A dictionary of configuration settings.
    """
    debug_settings = settings.get('debug', {})
    if debug_settings.get('enabled', 'yes').lower() == 'yes':
        log_file_path = debug_settings.get('log_file_path', 'debug.log')
        max_size_mb = int(debug_settings.get('max_size_mb', 5))
        backup_count = int(debug_settings.get('backup_count', 3))
        rotate_interval_hours = int(debug_settings.get('rotate_interval_hours', 24))

        logging.basicConfig(level=logging.DEBUG,
                            format='%(asctime)s - %(levelname)s - %(message)s',
                            datefmt='%Y-%m-%d %H:%M:%S')
        handler = RotatingFileHandler(log_file_path, maxBytes=max_size_mb*1024*1024, backupCount=backup_count)
        logging.getLogger('').addHandler(handler)
    else:
        logging.disable(logging.CRITICAL)

def send_to_home_assistant(base_url, token, entity_id, state, attributes):
    """
    Sends a request to Home Assistant to update the specified entity's state and attributes.

    Args:
        base_url (str): The base URL of the Home Assistant instance.
        token (str): The API token for Home Assistant.
        entity_id (str): The ID of the entity to update.
        state (str): The new state of the entity.
        attributes (dict): A dictionary of attributes to update for the entity.
    """
    url = f"{base_url}/api/states/{entity_id}"
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    payload = {"state": state, "attributes": attributes}
    try:
        response = requests.post(url, json=payload, headers=headers)
        response.raise_for_status()
        logging.info(f"Successfully updated {entity_id} in Home Assistant to '{state}' with attributes {attributes}.")
    except requests.RequestException as e:
        logging.error(f"Error updating {entity_id} in Home Assistant: {e}. Payload: {payload}")

def update_home_assistant(availability, notification_count, call_status, settings):
    """
    Updates Home Assistant entities based on the latest Microsoft Teams status.

    Args:
        availability (str): The current availability status.
        notification_count (str): The current notification count.
        call_status (str): The current call status.
        settings (dict): A dictionary of configuration settings.
    """
    logging.info(f"Updating Home Assistant: Availability={availability}, Call Status={call_status}, Notification Count={notification_count}")
    
    # Update status entity in Home Assistant
    entity_id_status = settings['entities']['entitystatus']
    state = settings['language'].get(availability.lower(), "unknown")
    attributes_status = {
        "friendly_name": settings['entities']['entitystatusname'],
        "icon": settings['icons'].get(state.lower(), "mdi:account")
    }
    send_to_home_assistant(settings['HAurl'], settings['token'], entity_id_status, state, attributes_status)

    # Update activity entity in Home Assistant
    entity_id_activity = settings['entities']['entityactivity']
    activity_state = "in a call" if call_status == "In a call" else "not in a call"
    icon_key = 'inacall' if call_status == "In a call" else 'notinacall'
    attributes_activity = {
        "friendly_name": settings['entities']['entityactivityname'],
        "icon": settings['icons'].get(icon_key, "mdi:phone-off")
    }
    send_to_home_assistant(settings['HAurl'], settings['token'], entity_id_activity, activity_state, attributes_activity)

def process_last_lines_of_log(file_path, settings):
    """
    Processes the last few lines of the most recent log file for status updates.

    Args:
        file_path (str): The path to the log file to be processed.
        settings (dict): A dictionary of configuration settings.
    """
    global last_known_availability, last_known_notification_count, last_known_call_status

    search_pattern = re.compile(r'availability: (\w+), unread notification count: (\d+)')
    search_pattern_start = re.compile(r'WebViewWindowWin:.*tags=Call.*Window previously was visible = false')
    search_pattern_end = re.compile(r'BluetoothRadioManager: Device watcher is Started.')

    try:
        with open(file_path, 'r', encoding='utf-8') as file:
            lines = file.readlines()[-10:]  # Read only the last 10 lines

        # Check each line for status updates
        for line in reversed(lines):
            if match := search_pattern.search(line):
                new_availability, new_notification_count = match.groups()
                if new_availability != last_known_availability or new_notification_count != last_known_notification_count:
                    last_known_availability = new_availability
                    last_known_notification_count = new_notification_count
                    logging.info("Status update found in the log.")
                    update_home_assistant(last_known_availability, last_known_notification_count, last_known_call_status, settings)
                    break  # Update found, no need to check further

            if search_pattern_start.search(line):
                last_known_call_status = "In a call"
                update_home_assistant(last_known_availability, last_known_notification_count, last_known_call_status, settings)
                break
            elif search_pattern_end.search(line):
                last_known_call_status = "Not in a call"
                update_home_assistant(last_known_availability, last_known_notification_count, last_known_call_status, settings)
                break

    except Exception as e:
        logging.error(f"Error reading file {file_path}: {e}")

def process_log_file(file_path, settings):
    """
    Processes the specified log file to extract the user's availability, call status, and notification count.
    If the expected patterns are not found in the log file, the status values are set to "unknown".

    Args:
        file_path (str): The path to the log file to be processed.
        settings (dict): A dictionary containing configuration settings, including Home Assistant connection details.
    """
    # Compile regex patterns to search for availability, notification count, and call status
    search_pattern = re.compile(r'availability: (\w+), unread notification count: (\d+)')
    search_pattern_start = re.compile(r'WebViewWindowWin:.*tags=Call.*Window previously was visible = false')
    search_pattern_end = re.compile(r'BluetoothRadioManager: Device watcher is Started.')
    
    # Initialize default statuses
    call_status = "Not in a call"  # Default call status
    found_text = False  # Flag to indicate if any expected pattern is found in the log

    try:
        with open(file_path, 'r', encoding='utf-8') as file:
            for line in file:
                # Check for availability and notification count
                match = search_pattern.search(line)
                if match:
                    # Update last known statuses if a match is found
                    last_availability, last_notification_count = match.groups()
                    found_text = True  # Indicate that expected text has been found
                
                # Infer call status from the current line using the start and end patterns
                if search_pattern_start.search(line):
                    call_status = "In a call"
                    found_text = True
                elif search_pattern_end.search(line):
                    call_status = "Not in a call"
                    found_text = True

        # If no expected text is found in the entire file, set status values to "unknown"
        if not found_text:
            last_availability = "unknown"
            last_notification_count = "unknown"
            call_status = "unknown"
            logging.info("No specific patterns were found in the log file. Setting status to 'unknown'.")

        # Update Home Assistant with the detected statuses
        logging.info(f"Detected status: {last_availability}, Call Status: {call_status}, Notification Count: {last_notification_count}")
        update_home_assistant(last_availability, last_notification_count, call_status, settings)
    except Exception as e:
        logging.error(f"Error reading file {file_path}: {e}")

def startup_log_read(settings):
    """
    Reads the most recent log file at startup to determine the user's current status.
    This function is intended to be called when the script starts to initialize the status.

    Args:
        settings (dict): A dictionary containing configuration settings, including the log folder path.
    """
    log_dir = os.path.expandvars(settings['logfolder'])
    if not os.path.exists(log_dir):
        print("Log directory does not exist, check configuration.")
        exit()  # Exit the script if the log directory does not exist

    file_pattern = re.compile(r'MSTeams_(\d{4}-\d{2}-\d{2})_(\d{2}-\d{2}-\d{2}\.\d+)\.log')
    latest_file = None
    latest_time = None

    # Iterate through all files in the log directory to find the most recent log file
    for file_name in os.listdir(log_dir):
        match = file_pattern.match(file_name)
        if match:
            # Parse the datetime from the filename to determine if it's the most recent
            file_datetime = datetime.strptime(f"{match.group(1)} {match.group(2)}", '%Y-%m-%d %H-%M-%S.%f')
            if latest_time is None or file_datetime > latest_time:
                # Update the latest time and file if this file is more recent
                latest_time = file_datetime
                latest_file = file_name

    # If a latest file is found, process it to initialize the user's status
    if latest_file:
        process_log_file(os.path.join(log_dir, latest_file), settings)

def main_loop(settings):
    """
    The main loop of the script. Continuously checks for updates in the latest Teams log file and processes it.
    
    Args:
        settings (dict): A dictionary of configuration settings.
    """
    while True:
        log_dir = os.path.expandvars(settings['logfolder'])
        if not os.path.exists(log_dir):
            logging.error("Log directory does not exist, check configuration.")
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
            process_last_lines_of_log(os.path.join(log_dir, latest_file), settings)

        time.sleep(3)  # Wait for 3 seconds before checking for new log files again.

if __name__ == "__main__":
    settings = load_settings('MSTeamsSettings.config')
    configure_logging(settings)
    startup_log_read(settings) 
    main_loop(settings)
