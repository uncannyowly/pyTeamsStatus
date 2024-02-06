#Modify these: 
$installDir = "C:\git\pyTeamsStatus" #Where you cloned repo to
$pythonDir = "$env:Userprofile\scoop\apps\python\current\python.exe" #Assumes you used Scoop to install Python 3

<# 
    ################################################
    Shouldn't need to modify below here... hopefully 
    ################################################
#> 

Set-ExecutionPolicy -ExecutionPolicy Unrestricted -Scope CurrentUser 

#Unsure if this is still necessary? 
Unblock-File $installDir\MSTeamsSettings.config
Unblock-File $installDir\Get-TeamsStatus.py 

$logDir = "$installDir\logs"
if (!($logDir)){mkdir $logDir|Out-Null}

$installScript = "$installDir\Get-TeamsStatus.py"
if (!($installScript)){  
    throw "Err! No script found. Review path.";
    exit 1
}

Start-Process -FilePath .\nssm.exe -ArgumentList "install 'Microsoft Teams Status Monitor' $pythonDir $installScript " -NoNewWindow -Wait

#Get user who is running script's home dir, update config if they didn't do so before installing service. 
$UserPath = "$env:Userprofile\AppData\Local"

(Get-Content -raw .\MSTeamsSettings.config).replace('%%localappdata%%',$UserPath) | Out-File .\MSTeamsSettings.config

Start-Service -Name "Microsoft Teams Status Monitor"