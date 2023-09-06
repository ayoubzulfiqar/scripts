param (
    [string]$FolderPath
)

# Check if the folder path is valid
if (-not (Test-Path -Path $FolderPath -PathType Container)) {
    Write-Host "Folder path '$FolderPath' is not valid."
    exit
}

# Define an array of valid media file extensions
$ValidMediaExtensions = @(".mp3", ".mp4", ".avi", ".mkv", ".wmv", ".flv")

# Recursively search for media files in the specified folder
$MediaFiles = Get-ChildItem -Path $FolderPath -Recurse |
    Where-Object { $ValidMediaExtensions -contains $_.Extension }

# Check if any media files were found
if ($MediaFiles.Count -eq 0) {
    Write-Host "No media files found in the specified folder."
    exit
}

# Get the path of the default media player executable (e.g., VLC)
$DefaultMediaPlayerPath = (Get-Command -Name "vlc" -ErrorAction SilentlyContinue).Source

if (-not $DefaultMediaPlayerPath) {
    Write-Host "Default media player (e.g., VLC) not found on your system."
    exit
}

# Play each media file using the default media player
foreach ($MediaFile in $MediaFiles) {
    Write-Host "Playing: $($MediaFile.FullName)"
    Start-Process -FilePath $DefaultMediaPlayerPath -ArgumentList $MediaFile.FullName -Wait
}

Write-Host "All media files played successfully."
