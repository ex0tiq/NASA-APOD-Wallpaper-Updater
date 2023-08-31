# NASA APOD Background Updater
This PowerShell script automates the configuration of a Windows scheduled task that runs on every user login.
On every login this script fetches the latest NASA Astronomy Picture of the Day, downloads it and sets it as the Windows background image.

## Configuration
At the top of the script you can customize two variables.    
* `$downloadPath`: The path where the downloaded images will be stored. Defaults to "APOD" subfolder of the current users "Pictures" library.   
**This folder will be automatically created if it doesn't exist!**
* `$taskName`: Name of the task that will be created. Only change this when no task has been created yet. Otherwise the script won't find the old task.

## How to setup
1. Download and save the script to a path where it can stay. This is important, as the scheduled task needs to be able to execute this script on every login.  
A working path would be `C:\Scripts`.
2. Run the script with right click -> Run with PowerShell   
The script will automatically ask for elevated privileges if needed. Confirm any UAC prompts with yes.
    * If there is no update task present, the script will ask if it should create one. Confirm this by entering 'y'.   
    * If the update task is present, the script will ask you if you want to delete it. Confirm this by entering 'y'.

## Use without scheduled task
If you don't want to run this script as a scheduled task, you can append the parameter "-Update" to the script call.   
The script will then only run once and update the current background image.