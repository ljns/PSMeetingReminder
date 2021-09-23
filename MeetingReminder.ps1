$minutesWithin = 5 # how close do you want a calendar event to be before you are notified
$checkFrequency = 240 # how frequently should the script check your calendar (in seconds)
$prompt = $false # Generates a user prompt after first beep if set to $true, else script will only beep. Leave $false if not running in ISE
$beepRepeatAmount = 10 # how many times to beep before continuing anyway

# get all available calendar items for the day
function Get-OutlookCalendarItems
{
    $date = Get-Date -Format "MM/dd/yyyy"

    # load the required .NET types
    Add-Type -AssemblyName 'Microsoft.Office.Interop.Outlook'
    
    # access Outlook object model
    $outlook = New-Object -ComObject outlook.application

    # connect to the appropriate location
    $namespace = $outlook.GetNameSpace('MAPI')
    $calendar = [Microsoft.Office.Interop.Outlook.OlDefaultFolders]::olFolderCalendar
    $folder = $namespace.getDefaultFolder($Calendar)
    # get todays calendar items
    $folder.items | Where-Object {$_.Start -match $date} | Select-Object -Property Start, Subject, Organizer
}

while ($true)
{
    # clear the console
    cls

    # get existing date
    $dateTime = Get-Date -Format "dd/MM/yyyy HH:mm:ss"

    # get calendar items and add in some whitespace so output isn't obscured by progress bar
    "Fetching Outlook calendar items`n`n`n`n`n`n`n"

    $calendarItems = Get-OutlookCalendarItems
    
    # set a var that we'll use later to avoid multiple alerts for meetings occurring in the same window
    $alreadyAlerted = $false

    # set a var to reset that status of the user prompt if it has been used
    $ackResult = $null

    # if there are items in the calendar
    if ($calendarItems -ne $null)
    {
        # iterate through the discovered items to see if any come up soon
        foreach ($item in $calendarItems)
        {
            $span = New-TimeSpan -Start $dateTime -End $item.Start
    
            # if one does come up within the specified time period
            if ($span.Minutes -le $minutesWithin -and $span -gt 0)
            {
                # check if this is a repeat alert
                if ($alreadyAlerted)
                {
                    "$($item.Subject) also falls within the specifed time (start time of $($item.Start)), but an alert has already triggered"
                }
                # if it isn't
                if (!$alreadyAlerted)
                {
                    # do something (beep)
                    "'$($item.Subject)' is starting within the next $minutesWithin minutes. Alert will repeat $beepRepeatAmount times"

                    $i = 0
                    while ($i -lt $beepRepeatAmount -and $ackResult -ne 0)
                    {
                        $i++    
                        [console]::Beep(1000, 150)    
                        start-sleep 2

                        if ($prompt)
                        {
                            # create user prompt
                            $ackOk = New-Object System.Management.Automation.Host.ChoiceDescription "&Snooze", "This will stop the alert now but it may come back"
                            $ackOptions = [System.Management.Automation.Host.ChoiceDescription[]]($ackOk, (New-Object System.Management.Automation.Host.ChoiceDescription "&Ignore"))
                            $ackResult = $host.ui.PromptForChoice("Meeting Reminder", "Your $($item.Subject) meeting starts at $($item.Start)", $ackOptions, 0) 
                        }
                    }
                    # set var to true so we don't alert again immediately after if another item falls withing the allotted time
                    $alreadyAlerted = $true
                }
                
            }
            # if one does not come up within the specified time period (or has already passed)
            else
            {
                if ($span -lt 0)
                {
                    "'$($item.Subject)' has already started: $($item.Start)"
                }
                else
                {
                    "'$($item.Subject)' is not starting in the next $minutesWithin minutes (start time of $($item.Start))"
                }
            }
        }
    } 
    # if there are no items in the calendar
    else
    {
        "No calendar items"
    }
    # wait before checking again
    $time = $checkFrequency
    $timerLength = $time / 100
    for ($time; $time -ne 0; $time--) 
    {
        $min = [int](([string]($time/60)).split('.')[0])
        $text = " " + $min + " minutes " + ($time % 60) + " seconds"
        Write-Progress -Activity "Watiting for.." -Status "$text" -PercentComplete ($time/$timerLength)
        Start-Sleep 1
    }
    Write-Progress -Activity "Watiting for.." -Completed
}
