########################################################
########################################################
[reflection.assembly]::LoadWithPartialName("System.Drawing")

# Define the root path where the media files are located
$rootPath = "\\Synology\photo\To SOrt"
$OrgaFolder = "\\Synology\photo\Organization"
$prefix = ""
$incnotfound=0
$foldercount=0
$totalFiles=0
$formattedPercentage=""
$formattedPercentageNotFound=""

#$PSDefaultParameterValues = $PSDefaultParameterValues.clone()
#$PSDefaultParameterValues += @{'Get-DateTakenProperty:ErrorAction' = 'SilentlyContinue'}

#function to get datetaken property
Function Get-DateTakenProperty{
Param([string]$path)
# Create a Bitmap object from the JPG file
try {
    $stream = [System.IO.File]::Open($path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read)
    $pic = New-Object System.Drawing.Bitmap($stream)
    $stream.Close()  # Close the stream when done

# Get the property item for "Date taken"
$bitearr = $pic.GetPropertyItem(36867).Value

$pic = $null

# Convert the byte array to a string
$string = [System.Text.Encoding]::ASCII.GetString($bitearr)

# Parse the date and time
$DateTime = [datetime]::ParseExact($string.Substring(0,19), "yyyy:MM:dd HH:mm:ss", $Null)

$percentNotFound = ($incnotfound / $totalFiles) * 100
$formattedPercentageNotFound = "{0:p2}" -f ($percentNotFound / 100)
#Write-Host "`r$totalFiles files - $formattedPercentage completed. $formattedPercentageNotFound of dateTaken not found."
return $DateTime
}
catch{
    $incnotfound++
    $percentNotFound = ($incnotfound / $totalFiles) * 100
    $formattedPercentageNotFound = "{0:p2}" -f ($percentNotFound / 100)
    #Write-Host "`r$totalFiles files - $formattedPercentage completed. $formattedPercentageNotFound of dateTaken not found." -NoNewline 
    return $null
}

}

# Function to create a directory if it doesn't exist
Function Create-DirectoryIfNotExist {
Param([string]$path)
if (-not (Test-Path $path)) {
New-Item -ItemType Directory -Path $path | Out-Null
}
}

# Function to rename the file with an increment if it already exists
Function Rename-WithIncrement {
Param([string]$path, [string]$name, [string]$extension)
$counter = 1
$newName = $name + "_" + $counter
$newFullPath = Join-Path -Path $path -ChildPath ($newName + $extension)
while (Test-Path $newFullPath) {
$counter++
$newName = $name + "_" + $counter
$newFullPath = Join-Path -Path $path -ChildPath ($newName + $extension)
}
return $newFullPath
}



    $mediaFiles = Get-ChildItem -Path $rootPath -Recurse | Where-Object { $_.Extension -match "\.(xcf|bmp|JPG|jpg|jpeg|png|mp4|mkv|hevc|mov|avi|dng|heic|cr2)$" }
    $mediaFiles+=$mediaroot
    $totalFiles = $mediaFiles.Count
    $filesProcessed = 0
    

    # Start a timer
    $timer = [Diagnostics.Stopwatch]::StartNew()
    [bool]$NewLoop=$true
    foreach ($file in $mediaFiles) {
        # Determine if the file is a photo or video
        $isPhoto = $file.Extension -match "\.(xcf|bmp|JPG|jpg|jpeg|png|dng|heic|cr2)$"
        # Get the date value
        $dateValue = if ($isPhoto) {
        # Attempt to get the 'Date Taken' from the photo metadata
        (Get-ItemProperty -Path $file.FullName).LastWriteTime
        } else {
        # Use the 'CreationTime' for videos
        $file.CreationTime
        }
        # Format the date into 'yyyy' and 'yyyy-MM' structure
        $year = $dateValue.ToString("yyyy")
        $yearMonth = $dateValue.ToString("yyyy-MM")
        # Define the new folder path based on the date
        $newFolderPath = Join-Path -Path $OrgaFolder -ChildPath "$year\$yearMonth"
        # Create the new folder if it doesn't exist
        Create-DirectoryIfNotExist -path $newFolderPath
        # Define the new file path
        $newFilePath = Join-Path -Path $newFolderPath -ChildPath ($prefix + $file.Name)
        # Check if the file already exists in the destination folder
        if (Test-Path $newFilePath) {
        # Rename the file with an increment
        $newFilePath = Rename-WithIncrement -path $newFolderPath -name ($prefix + $file.BaseName) -extension $file.Extension
        }
        # Move the file to the new folder
        #robocopy $file.DirectoryName $newFolderPath ($prefix + $file.PSChildName) /MOV
        Move-Item -Path $file.FullName -Destination $newFilePath -Force

        # Increment the count of files processed
        $filesProcessed++

        # Calculate the percentage of completion
        $percentComplete = ($filesProcessed / $totalFiles) * 100
        $formattedPercentage = "{0:p2}" -f ($percentComplete / 100)

        $percentNotFound = ($incnotfound / $totalFiles) * 100
        $formattedPercentageNotFound = "{0:p2}" -f ($percentNotFound / 100)
        
        # Check if the timer has reached 1 minute
        $direct = $SubFolder.FullName
        if ($timer.Elapsed.TotalMilliseconds -ge 100) {
            if($NewLoop){
                Write-Host "`r$foldercount / $foldertotal - $totalFiles files - $formattedPercentage completed. $formattedPercentageNotFound of dateTaken not found - ($direct)"
                $NewLoop=$false
            }
            else {
                Write-Host "`r$foldercount / $foldertotal - $totalFiles files - $formattedPercentage completed. $formattedPercentageNotFound of dateTaken not found - ($direct)" -NoNewline
            }
            # Reset the timer
            $timer.Restart()
        }
    }

Write-Host "Media reorganization complete."