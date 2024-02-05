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
    """
    url = f"{base_url}/api/states/{entity_id}"
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    payload = {"state": state, "attributes": attributes}
    try:
        response = requests.post(url, json=payload, headers=headers)
        response.raise_for_status()
        print(f"Successfully updated {entity_id} in Home Assistant.")
    except requests.RequestException as e:
        print(f"Error updating {entity_id} in Home Assistant: {e}")

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

def process_log_file(file_path, settings):
    """
    Processes the latest log file to extract availability and notification count.
    """
    search_pattern = re.compile(r'availability: (\w+), unread notification count: (\d+)')
    try:
        with open(file_path, 'r', encoding='utf-8') as file:
            for line in file:
                match = search_pattern.search(line)
                if match:
                    update_home_assistant(match.group(1), match.group(2), settings)
    except Exception as e:
        print(f"Error reading file {file_path}: {e}")

def update_home_assistant(availability, notification_count, settings):
    """
    Updates Home Assistant entities based on Teams availability and notification count.
    """
    # Adjusting to lowercase for consistent key access
    if 'entitystatus' not in settings['entities']:
        print("Error: 'entitystatus' key not found in settings['entities']")
        return  # Early return if 'entitystatus' is missing

    entity_id = settings['entities']['entitystatus']
    state = settings['language'].get(availability.lower(), availability)  # Lowercase for matching
    icon_key = 'inacall' if state.lower() in ["inacall", "onthephone"] else 'notinacall'
    icon = settings['icons'].get(icon_key.lower(), "mdi:account")  # Lowercase for consistency
    attributes = {
        "friendly_name": settings['entities']['entitystatusname'],
        "icon": icon
    }

    print(f"Latest Availability: {availability}, Latest Notification Count: {notification_count}")
    send_to_home_assistant(settings['HAurl'], settings['token'], entity_id, state, attributes)

if __name__ == "__main__":
    settings = load_settings('MSTeamsSettings.config')
    configure_logging(settings)

    # Debug print to verify entities loading
    print(settings['entities'])

    main_loop(settings)
