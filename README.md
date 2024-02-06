
# NEW Microsoft Teams Status Monitor for Home Assistant

## Introduction (EBOOZ)

In the era of remote work, integrating your digital workspace with your home automation system offers a seamless way to enhance your work-from-home experience. This solution provides a method to automate activities in your home automation system based on your Microsoft Teams status without needing organizational consent for Microsoft Graph API.

This script monitors the Microsoft Teams client logfile for changes in status and activity, updating corresponding sensors in Home Assistant. It's designed for the NEW version of Microsoft Teams, which logs status information in a readable format.

> [!NOTE]
> This solution only works for the NEW version of Microsoft Teams. The new version was updated to finally contain a status in the logs.txt file which can be read by the script. There however are still some limitations around calls that are being worked through. The original script from EBOOZ was also migrated from Powershell to Python. 

## Requirements

- Home Assistant setup with API access.
- Long-lived access token from Home Assistant ([see HA documentation](https://developers.home-assistant.io/docs/auth_api/#long-lived-access-token)).
- Python 3 and the `requests` library installed on the system where the script will run.
- The Microsoft Teams client installed and operational.

Before running the script, create the necessary sensors in Home Assistant by adding the following to your `configuration.yaml`:

```yaml
input_text:
  teams_status:
    name: Microsoft Teams status
    icon: mdi:microsoft-teams
  teams_activity:
    name: Microsoft Teams activity
    icon: mdi:phone-off

sensor:
  - platform: template
    sensors:
      teams_status: 
        friendly_name: "Microsoft Teams status"
        value_template: "{{states('input_text.teams_status')}}"
        icon_template: "{{state_attr('input_text.teams_status','icon')}}"
        unique_id: sensor.teams_status
      teams_activity:
        friendly_name: "Microsoft Teams activity"
        value_template: "{{states('input_text.teams_activity')}}"
        unique_id: sensor.teams_activity
```

After modifying `configuration.yaml`, restart Home Assistant to register the new sensors.

## Configuration and Execution

1. **Clone or download the script files** to a preferred location on your system, e.g., `C:\Scripts` for Windows.

2. **Edit the `MSTeamsSettings.config` file**:
    - Replace placeholders with your actual Long-lived access token, Home Assistant URL, and specify the path to your Teams logs.
    - Adjust the `[DebugSettings]` section according to your preferences for logging.

3. **Run the Script**:
    - For continuous monitoring, execute `MSTeams.py` using a command prompt or terminal:
      ```bash
      python MSTeams.py
      ```
    - To run as a service on Windows, you can use tools like NSSM (Non-Sucking Service Manager) or Task Scheduler.

## Debugging and Logging

The Python script supports logging based on the configuration in `MSTeamsSettings.config`. Configure `[DebugSettings]` to enable or disable logging, set the log file path, maximum size, backup count, and rotation interval.

Ensure to monitor the log file for any errors or issues, especially after initial setup or changes to your Home Assistant or Teams configuration.
