[CmdletBinding()]
param(
    [Parameter()]
    [DateTime]$StartDate = (Get-Date).Date,
    
    [Parameter()]
    [DateTime]$EndDate = (Get-Date).Date.AddDays(7),
    
    [Parameter()]
    [string]$StartTime = "08:00",
    
    [Parameter()]
    [string]$EndTime = "17:00",

    [Parameter()]
    [int]$MinimumSlotMinutes = 30
)

try {
    Write-Verbose "Script started with parameters:"
    Write-Verbose "StartDate: $StartDate"
    Write-Verbose "EndDate: $EndDate"
    Write-Verbose "StartTime: $StartTime"
    Write-Verbose "EndTime: $EndTime"
    Write-Verbose "Minimum slot duration: $MinimumSlotMinutes minutes"
    Write-Verbose "-------------------"

    if (!(Get-Module -ListAvailable -Name Microsoft.Graph)) {
        Write-Host "Microsoft Graph module is not installed. Please install it using:" -ForegroundColor Yellow
        Write-Host "Install-Module -Name Microsoft.Graph" -ForegroundColor Yellow
        exit
    }
    
    # Connect to Microsoft Graph
    Write-Verbose "Connecting to Microsoft Graph..."
    Connect-MgGraph -Scopes "Calendars.Read", "User.Read"
    Write-Verbose "Connected to Microsoft Graph successfully"

    # Get current user
    Write-Verbose "Getting current user information..."
    $context = Get-MgContext
    if (-not $context) {
        throw "Failed to get Microsoft Graph context. Please ensure you are properly connected."
    }

    Write-Verbose "Context information:"
    Write-Verbose "Account: $($context.Account)"
    Write-Verbose "TenantId: $($context.TenantId)"
    Write-Verbose "Scopes: $($context.Scopes -join ', ')"

    if (-not $context.Account -or [string]::IsNullOrEmpty($context.Account)) {
        throw "User account information is not available. Please ensure you are properly authenticated."
    }

    $userId = $context.Account
    Write-Verbose "Current user ID: $userId"

    # Convert time strings to DateTime objects
    Write-Verbose "Converting time parameters..."
    $startDateTime = [DateTime]::ParseExact($StartTime, "HH:mm", $null)
    $endDateTime = [DateTime]::ParseExact($EndTime, "HH:mm", $null)
    Write-Verbose "Time conversion completed"

    # Create the time window for availability (all times in local time, no extra conversion needed)
    Write-Verbose "Creating time window for availability check..."
    $timeWindow = @{
        startDateTime = $StartDate.Add($startDateTime.TimeOfDay).ToString("o")
        endDateTime = $EndDate.Add($endDateTime.TimeOfDay).ToString("o")
    }
    Write-Verbose "Time window created:"
    Write-Verbose "Start: $($timeWindow.startDateTime)"
    Write-Verbose "End: $($timeWindow.endDateTime)"
  

    # Get all calendar events in the range using -All
    Write-Verbose "Fetching all calendar events in one call using -All..."
    $calendarEvents = Get-MgUserCalendarView -UserId $userId -StartDateTime $timeWindow.startDateTime -EndDateTime $timeWindow.endDateTime -All
    Write-Verbose "Found $($calendarEvents.Count) calendar events"

    # Convert events to a more easily searchable format (include Subject, convert from UTC to local time)
    $busyPeriods = $calendarEvents | ForEach-Object {
        @{
            Start = ([DateTime]::Parse($_.Start.DateTime)).ToLocalTime()
            End = ([DateTime]::Parse($_.End.DateTime)).ToLocalTime()
            Subject = $_.Subject
        }
    } | Sort-Object Start

    # Initialize variables for scanning
    $currentDate = $StartDate
    $availableSlots = @()

    while ($currentDate -le $EndDate) {
        Write-Verbose "\n--- Processing $($currentDate.DayOfWeek) $($currentDate.ToString('yyyy-MM-dd')) ---"
        if ($currentDate.DayOfWeek -in @('Saturday', 'Sunday')) {
            Write-Verbose "Skipping weekend day: $($currentDate.DayOfWeek)"
            $currentDate = $currentDate.AddDays(1)
            continue
        }

        $dayStart = $currentDate.Add($startDateTime.TimeOfDay)
        $dayEnd = $currentDate.Add($endDateTime.TimeOfDay)
        Write-Verbose "Workday starts at $($dayStart.ToString('HH:mm')), ends at $($dayEnd.ToString('HH:mm'))"

        # Get all busy periods for this day, sorted
        $dayBusy = $busyPeriods | Where-Object { $_.Start.Date -eq $currentDate.Date } | Sort-Object Start
        Write-Verbose "Found $($dayBusy.Count) busy periods for this day."
        foreach ($b in $dayBusy) {
            Write-Verbose ("Busy: {0} to {1} | {2}" -f $b.Start.ToString('HH:mm'), $b.End.ToString('HH:mm'), $b.Subject)
        }

        $freeStart = $dayStart
        foreach ($busy in $dayBusy) {
            if ($busy.End -le $freeStart) { 
                Write-Verbose ("Skipping busy period ending before freeStart: {0} to {1}" -f $busy.Start.ToString('HH:mm'), $busy.End.ToString('HH:mm'))
                continue 
            }
            if ($busy.Start -gt $freeStart) {
                $gap = ($busy.Start - $freeStart).TotalMinutes
                Write-Verbose ("Gap found: {0} to {1} ({2} min)" -f $freeStart.ToString('HH:mm'), $busy.Start.ToString('HH:mm'), [math]::Round($gap))
                if ($gap -ge $MinimumSlotMinutes) {
                    Write-Verbose "--> Reporting as available."
                    $availableSlots += @{
                        Start = $freeStart
                        End = $busy.Start
                        Duration = $gap
                    }
                } else {
                    Write-Verbose "--> Gap too short, not reported."
                }
            }
            if ($busy.End -gt $freeStart) {
                $freeStart = $busy.End
            }
        }
        if ($freeStart -lt $dayEnd) {
            $gap = ($dayEnd - $freeStart).TotalMinutes
            Write-Verbose ("End-of-day gap: {0} to {1} ({2} min)" -f $freeStart.ToString('HH:mm'), $dayEnd.ToString('HH:mm'), [math]::Round($gap))
            if ($gap -ge $MinimumSlotMinutes) {
                Write-Verbose "--> Reporting as available."
                $availableSlots += @{
                    Start = $freeStart
                    End = $dayEnd
                    Duration = $gap
                }
            } else {
                Write-Verbose "--> Gap too short, not reported."
            }
        }
        $currentDate = $currentDate.AddDays(1)
    }

    # Group slots by date for better organization
    $groupedSlots = $availableSlots | Group-Object { $_.Start.Date }

    # Prepare output for clipboard
    $output = "USER AVAILABILITY`n====================="
    foreach ($dateGroup in $groupedSlots) {
        $date = [DateTime]$dateGroup.Name
        $output += "`n$($date.DayOfWeek) $($date.ToString('yyyy-MM-dd'))`n-------------------------"
        if ($dateGroup.Group.Count -eq 0) {
            $output += "`n  No available slots"
        } else {
            foreach ($slot in $dateGroup.Group) {
                $output += "`n  From {0} to {1}" -f $slot.Start.ToString('hh:mm tt'), $slot.End.ToString('hh:mm tt')
            }
        }
    }
    $output += "`n====================="

    # Output to console
    Write-Host $output
    # Copy to clipboard
    $output | Set-Clipboard
    Write-Host "`n(Output has been copied to the clipboard.)"

}
catch {
    Write-Host "`nAn error occurred:" -ForegroundColor Red
    Write-Host "Error Message: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Error Type: $($_.Exception.GetType().FullName)" -ForegroundColor Red
    Write-Host "Line Number: $($_.InvocationInfo.ScriptLineNumber)" -ForegroundColor Red
    Write-Host "Script Stack Trace:" -ForegroundColor Red
    Write-Host $_.ScriptStackTrace -ForegroundColor Red
}
finally {
    
}

