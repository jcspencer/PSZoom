# Copyright 2020 James Spencer, Girton Grammar School Bendigo
#
# Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"),
# to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense,
# and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
#
# The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.
#
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY,
# WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.


Import-Module PSZoom

# All commands require an API key and API secret.
# You can generate the JWT key/secret from https://marketplace.zoom.us/develop/create, then click on
# 'Create' under JWT.
# When generating the JWT, choose a large timeframe, such as 6 months.

# Used for all Zoom API calls
$ZoomCredentials = @{
  ApiKey     = 'ZOOM_API_KEY_GOES_HERE'
  ApiSecret  = 'ZOOM_API_SECRET_GOES_HERE'
}

# Used to skip password authentication on Zoom recording download links
$JWT = "ZOOM_JWT_GOES_HERE"

# The local file share to download Zoom recordings to
$BaseOutputPath = "\\gfs01.girton.vic.edu.au\ZoomArchive$\"
# The number of days to go back when downloading recordings (default: 1)
$NumberOfDays = 1
# The list of names of the Zoom groups who should have their recordings archived
$GroupNames = @(
  "Staff",
  "StaffOpen"
)

# The text log file to record download logs to
$LogFile = "\\appsupport\Scripts\Zoom\downloader_$(Get-Date -Format "yyyy_MM").log"

# RClone configuration for syncing up to Sharepoint, etc.
$RCloneLogFile = "\\appsupport\Scripts\Zoom\rclone_$(Get-Date -Format "yyyy_MM").log"
$RCloneExe = "\\appsupport\Scripts\Zoom\rclone\rclone.exe"
$RCloneConfig = "\\appsupport\Scripts\Zoom\rclone.conf"
$RCloneSyncPath = "sharepoint:Recordings"

# A function to log a message to the log file,
# as well as printing it to the console
function Log-MessageToFile {
    param([String]$Message)
  
    $Date = (Get-Date).ToString()
    $LogLine = "[$Date] $Message"
  
    Write-Host $LogLine
    Add-Content $LogFile $LogLine
}

# Strip any invalid characters from a Zoom meeting name
# before writing it to disk
Function Remove-InvalidFileNameChars {
  param(
    [Parameter(Mandatory = $true,
      Position = 0,
      ValueFromPipeline = $true,
      ValueFromPipelineByPropertyName = $true)]
    [String]$Name
  )
  
  $invalidChars = [IO.Path]::GetInvalidFileNameChars() -join ''
  $re = "[{0}]" -f [RegEx]::Escape($invalidChars)
  return ($Name -replace $re)
}

# Prevent Invoke-WebRequest from showing progress bar, therefore speeding it up a lot.
$ProgressPreference = 'SilentlyContinue'

# Get the date to start downloading recordings from
$StartDate = (Get-Date).AddDays(-$NumberOfDays)
$StartDateString = (Get-Date -Date $StartDate -Format "yyyy-MM-dd")

# Get the date to end downloading recordings
$EndDate = Get-Date
$EndDateString = (Get-Date -Date $EndDate -Format "yyyy-MM-dd")
Log-MessageToFile "Downloading from $StartDateString to $EndDateString"

Log-MessageToFile "****** Starting up ******"


# Get all users on the Zoom instance
Log-MessageToFile "[ ] Grabbing Zoom user list..."
$ZoomUsers = (Get-ZoomUsers -status active -allpages @ZoomCredentials) + (Get-ZoomUsers -status inactive -allpages @ZoomCredentials)

# Get the IDs of all the groups the script has listed
$GroupIDs = Get-ZoomGroups @ZoomCredentials | ?{ $GroupNames -contains $_.name } | %{ $_.id }

# Get a list of all the matching users
$RecordingUsers = $ZoomUsers | ?{
  # Does the User belong to at least one of the groups in the list?
  @($_.group_ids | ?{$GroupIDs -contains $_}).length -ge 1
}

Log-MessageToFile "[*] Grabbed Zoom user list."

foreach ($User in $RecordingUsers) {
  Log-MessageToFile "--- User: $($User.email) [$($User.id)] ---"

  # Retrieve a list of Meeting recordings for the given user in the specified range
  $RecordingList = Get-ZoomRecordings -UserId $User.email -From $StartDateString -To $EndDateString @ZoomCredentials
  $Meetings = $RecordingList | select meetings | %{ $_.psobject.properties.value }  

  Log-MessageToFile "User has $($RecordingList.total_records) recordings."
  $Count = 1

  # Iterate through each recording
  foreach ($Recording in $RecordingList.meetings) {
    Log-MessageToFile "[ ] Downloading recording #$Count..."

    # Use the first part of the email to group the recordings (depends on org)
    $Username = $User.email.Split("@")[0]

    # e.g. "2020 April 02"
    $RecordingDate = Get-Date -Date $([datetime]::Parse($Recording.start_time)) -f "yyyy MMMM dd"

    # e.g. "15:39"
    $RecordingTime = Get-Date -Date $([datetime]::Parse($Recording.start_time)) -f "HH.mm"

    # Path will be in the form:
    # \\fileserver\share\jamesspencer\2020 April 20\15.30 - IT Staff Meeting
    $OutputUserFolder = $BaseOutputPath + $Username + "\" + $RecordingDate + "\"
    $OutputMeetingFolder = Remove-InvalidFileNameChars($RecordingTime + " - " + $Recording.topic.trim())
    
    # Make the folder for the meeting if it does not exist
    if(-not (Test-Path "$OutputUserFolder$OutputMeetingFolder")) {
      New-Item -ItemType "directory" -name $OutputMeetingFolder -Path $OutputUserFolder | Out-Null
    }

    # Go through each recording 
    foreach ($File in $Recording.recording_files) {
      if ($file.file_type -eq "TIMELINE") {
        $OutputFilename = "timeline.json"
      } else {
        # e.g. "shared_screen_with_speaker_view.mp4"
        $OutputFilename = (Remove-InvalidFileNameChars($File.recording_type)) + "." + $file.file_type
      }
      
      # Build the final output path
      $OutputFullPath = $OutputUserFolder + $OutputMeetingFolder + "\" + $OutputFilename

      # Only download a file if it doesn't already exist
      if(-not (Test-Path -Path $OutputFullPath)) {
        # Write the filesize to the log if it's been provided by the API
        # (this is not provided for all files)
        $size = if($file.file_size) { $file.file_size } else { 0 }
        Log-MessageToFile "[ ] Retrieving $outputFilename [$([math]::round($size/1MB, 2)) MB]"

        # Download the file and place it at the desired output path
        # (Use the JWT to allow downloading without the meeting password)
        Invoke-WebRequest -Uri "$($File.download_url)?access_token=$JWT" -OutFile $OutputFullPath 

        Log-MessageToFile "[*] Retrieved $outputFilename"
      }
    }

    $Count += 1
  }

  Log-MessageToFile
}

# Sync the recordings up to SharePoint if they do not already exist
Log-MessageToFile "****** Finished Download, Syncing to Sharepoint Online ******"
& $RCloneExe copy $BaseOutputPath $RCloneSyncPath --log-file $RCloneLogFile --config $RCloneConfig
Log-MessageToFile "****** Finished Sync ******"