# PowerShell Script to automatically download the latest Nasa APOD image and set is as the windows background image

param(
  [switch]$Update
)

####### OPTIONS #########

# Images will be downloaded and saved to this folder
$downloadPath = "$([Environment]::GetFolderPath("MyPictures"))\APOD"
# Name of scheduled task
$taskName = "NASA APOD BG Image Updater"

##############################################################################################################

# Define custom lib to set bg image
$setwallpapersrc = @"
using System.Runtime.InteropServices;

public class Wallpaper
{
  public const int SetDesktopWallpaper = 20;
  public const int UpdateIniFile = 0x01;
  public const int SendWinIniChange = 0x02;
  [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
  private static extern int SystemParametersInfo(int uAction, int uParam, string lpvParam, int fuWinIni);
  public static void SetWallpaper(string path)
  {
    SystemParametersInfo(SetDesktopWallpaper, 0, path, UpdateIniFile | SendWinIniChange);
  }
}
"@
Add-Type -TypeDefinition $setwallpapersrc

##################

# Function to run image update
function Update() {
  $imagePath = Get-CurrentApodImage

  if ($imagePath -is [string] -and -not ([System.String]::IsNullOrEmpty($imagePath))) {
    Set-BGImage -imagePath $imagePath
  }
}

# Download current APOD image
function Get-CurrentApodImage() {
  # create download dir if it doesn't exist
  if (-not (Test-Path -LiteralPath $downloadPath)) {
    New-Item -ItemType Directory -Path $downloadPath
  }

  # Get current APOD page
  $url = [System.Uri]"https://apod.nasa.gov/apod/astropix.html"

  $segments = $url.Segments | Select-Object -SkipLast 1
  $base = "$($url.Scheme)://$($url.Host)$($segments -join `"`")"

  $result = Invoke-WebRequest -Uri $url -TimeoutSec 10 -ErrorAction Stop

  $srcpattern = '(?i)href="(.*?)"'
  $src = ([regex]$srcpattern ).Matches($result.Content)

  $image = $null
  # Get image url
  $src | ForEach-Object { 
      if ($_.Groups[1].Value.EndsWith(".jpg") -or $_.Groups[1].Value.EndsWith(".jpeg")) {
          if ([System.String]::IsNullOrEmpty($image)) {
            $image = [System.Uri]($base + $_.Groups[1].Value)
          } else { return }
      }
  }

  $fileName = (Get-Date -Format "yyyyMMdd") + "_" + $image.Segments[$image.Segments.Count-1]
  $fullDLPath = ($downloadPath + "\" + $fileName)

  # Check if file already exists
  if (Test-Path -LiteralPath $fullDLPath) {
    Write-Host -ForegroundColor Red "Current image already exists, aborting!"
    return $false;
  }

  ### Download image
  $ProgressPreference = 'SilentlyContinue'
  Invoke-WebRequest -Uri $image -TimeoutSec 10 -ErrorAction Stop -OutFile $fullDLPath

  return $fullDLPath
}

function Set-BGImage($imagePath) {
  # Set background image
  [Wallpaper]::SetWallpaper($imagePath)
}

# Create new scheduled task to update bg image on every login
function New-SchedTask {
  if ([System.String]::IsNullOrEmpty($taskName)) {
    Write-Host -ForegroundColor Red "Taskname cannot be empty!"
    return $false
  }

  $taskAction = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-ExecutionPolicy Bypass -WindowStyle Hidden -NonInteractive -File $($PSCommandPath) -Update"
  $taskTrigger = New-ScheduledTaskTrigger -AtLogOn
  $taskSettings = New-ScheduledTaskSettingsSet -RunOnlyIfNetworkAvailable -ExecutionTimeLimit (New-TimeSpan -Minutes 10) 
  $taskPrincipal = New-ScheduledTaskPrincipal -UserId ($env:USERDOMAIN + "\" + $env:USERNAME)
  $task = New-ScheduledTask -Action $taskAction -Trigger $taskTrigger -Settings $taskSettings -Principal $taskPrincipal

  # Register task
  Register-ScheduledTask -TaskName $taskName -TaskPath "\" -InputObject $task
}

function Remove-SchedTask {
  if (-not (Get-ScheduledTask -TaskName $taskName)) {
    return $false
  }

  Unregister-ScheduledTask -TaskName $taskName -TaskPath "\" -Confirm:$false
  return $true
}

#########################

if ($Update.IsPresent) {
  # Run update
  Update

  # Test if scheduled task is already registered
}  else {
  if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {  
    $arguments = "& '" +$myinvocation.mycommand.definition + "'"
    Start-Process powershell -Verb runAs -ArgumentList $arguments
    break
  }

  Write-Host "####################################################################"
  Write-Host "#                                                                  #"
  Write-Host "#            NASA Astronomy Picture of the Day Updater             #"
  Write-Host "#                                                                  #"
  Write-Host "#              by ex0tiq <https://github.com/ex0tiq>               #"
  Write-Host "#                                                                  #"
  Write-Host "####################################################################"
  Write-Host ""

  if (Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue) {
    # Ask if task should be deleted
    $confirmation = Read-Host "Scheduled task already exists. Delete it? [y/N]"
    if ([System.String]::IsNullOrEmpty($confirmation)) { $confirmation = 'n' }
  
    if ($confirmation -eq 'y') {
      if (Remove-SchedTask) {
        Write-Host -ForegroundColor Green "[SUCCESS] Task has been deleted!"
        Start-Sleep -Seconds 5
      } else {
        Write-Host -ForegroundColor Red "[ERROR] Error while deleting task: $($Error[0].InnerException.Message)"
        Read-Host
      }
    }
    return 
    # If not, create new task
  } else {
    $confirmation = Read-Host "Scheduled task does not exists. Create it? [y/N]"
    if ([System.String]::IsNullOrEmpty($confirmation)) { $confirmation = 'n' }
  
    if ($confirmation -eq 'y') {
      if (New-SchedTask) {
        Write-Host -ForegroundColor Green "[SUCCESS] Task has been created!"
        Start-Sleep -Seconds 3
        Write-Host ""
        $confirm = Read-Host "Run task now? [y/N]"
        if ([System.String]::IsNullOrEmpty($confirm)) { $confirm = 'n' }

        if ($confirm -eq 'y') {
          Update
        }
      } else {
        Write-Host -ForegroundColor Red "[ERROR] Error while creating task: $($Error[0].InnerException.Message)"
        Read-Host
      }
    }
    return 
  }
}
