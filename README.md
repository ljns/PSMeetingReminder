# PSMeetingReminder
Simple PowerShell script that reads your Outlook Calendar items and is more annoying that a regular toast notification.

## Configuration
Some variables are set in the script anyway. Adjust as you see fit.

| Variable | Default | Description |
|-|-|-|
| `minutesWithin` | `5` | How close do you want a calendar event to be before you are notified |
| `checkFrequency` | `240` | How frequently should the script check your calendar (in seconds) |
| `prompt` | `$false` | Generates a user prompt after first beep if set to $true, else script will only beep. Leave $false if not running in ISE |
| `beepRepeatAmount` | `10` | How many times to beep before continuing anyway |

## How to run

Download the script and run `.\MeetingReminder.ps1` or open in PowerShell ISE and run (better)

# Examples
Running in PowerShell ISE

![image](https://user-images.githubusercontent.com/35964690/134339919-9a366f37-1f13-47a3-8d4b-7c9a12d70c66.png)

Example prompt generated when `$prompt` is set to `$true` and run from ISE

![image](https://user-images.githubusercontent.com/35964690/134341131-81914717-1d75-4d16-a4a3-5e425109054c.png)

Example prompt generated when `$prompt` is set to `$true` and run from PowerShell. 

![image](https://user-images.githubusercontent.com/35964690/134341710-b7153b02-ce37-4de5-ac1a-5ef866a2e1c3.png)

Running in PowerShell

![image](https://user-images.githubusercontent.com/35964690/134341293-3850d7fe-d4c5-4953-8662-38d8e2842ef8.png)

