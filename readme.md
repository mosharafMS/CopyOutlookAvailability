# Outlook Availability Copier

A PowerShell script that retrieves your Microsoft Outlook calendar availability and copies it to your clipboard in a formatted text format.

## Features

- Retrieves calendar availability from Microsoft Outlook using Microsoft Graph API
- Configurable date range and time window
- Excludes weekends automatically
- Filters out slots shorter than a specified minimum duration
- Formats output in a clean, readable text format
- Automatically copies the formatted availability to your clipboard

## Prerequisites

- Windows PowerShell 5.1 or PowerShell 7+
- Microsoft Graph PowerShell module
- Microsoft 365 account with calendar access

## Installation

1. Install the Microsoft Graph PowerShell module:
```powershell
Install-Module -Name Microsoft.Graph
```

2. Download the `copyAvailability.ps1` script to your local machine.

## Usage

Run the script with default parameters:
```powershell
.\copyAvailability.ps1
```

Or specify custom parameters:
```powershell
.\copyAvailability.ps1 -StartDate "2024-03-20" -EndDate "2024-03-27" -StartTime "09:00" -EndTime "18:00" -MinimumSlotMinutes 60
```

### Parameters

- `StartDate`: Start date for availability check (default: today)
- `EndDate`: End date for availability check (default: 7 days from today)
- `StartTime`: Start time of working hours (default: "08:00")
- `EndTime`: End time of working hours (default: "17:00")
- `MinimumSlotMinutes`: Minimum duration for available slots in minutes (default: 30)

## Output Format

The script generates output in the following format:
```
USER AVAILABILITY
=====================
Monday 2024-03-18
-------------------------
  From 08:00 AM to 09:30 AM
  From 11:00 AM to 12:00 PM
  From 02:00 PM to 05:00 PM
=====================
```

The output is automatically copied to your clipboard for easy pasting.

## Error Handling

The script includes comprehensive error handling for:
- Missing Microsoft Graph module
- Authentication failures
- Invalid date/time parameters
- API connection issues

## Contributing

Feel free to submit issues and enhancement requests!

## License

This project is open source and available under the MIT License.
