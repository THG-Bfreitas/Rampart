#Sagitta Software preqes for rampart
$DownloadURL = "https://wp40.sagitta-online.com/rampart/Sagitta/installs/AMSSagittaActivexControls.msi"

$software = "AMSSagittaActivexControls"
$installed = (Get-ItemProperty HKLM:\SOFTWARE\Classes\Installer\Products\* | Where { $_.ProductName -match $software }) -ne $null

If(-Not $installed) {
	Write-Host "'$software' NOT is installed."

mkdir C:\hilb\ -ErrorAction SilentlyContinue
# Download Installer to 'C:\HILB\'
Invoke-RestMethod -Uri $DownloadURL -OutFile C:\hilb\AMSSagittaActivexControls.msi
#Sleep for download to complete
Start-sleep -seconds 10



# Silently Install the Application

Start-Process 'C:\hilb\AMSSagittaActivexControls.msi' -ArgumentList '/qn'

#Sleep for installation to complete
Start-sleep -seconds 10

} Else { Write-host $software already installed Skipping. } 




#Sagitta Software preqes for rampart
$DownloadURL = "https://wp40.sagitta-online.com/rampart/Sagitta/installs/Sagitta%20Launcher.msi"

$software = "Sagitta Launcher"
$installed = (Get-ItemProperty HKLM:\SOFTWARE\Classes\Installer\Products\* | Where { $_.ProductName -match $software }) -ne $null

If(-Not $installed) {
	Write-Host "'$software' NOT is installed."

mkdir C:\hilb\ -ErrorAction SilentlyContinue
# Download Installer to 'C:\HILB\'
Invoke-RestMethod -Uri $DownloadURL -OutFile C:\hilb\Sagitta%20Launcher.msi
#Sleep for download to complete
Start-sleep -seconds 10



# Silently Install the Application

Start-Process 'C:\hilb\Sagitta%20Launcher.msi' -ArgumentList '/qn'

#Sleep for installation to complete
Start-sleep -seconds 10

} Else { Write-host $software already installed Skipping. } 



#Sagitta Software preqes for rampart
$DownloadURL = "https://wp40.sagitta-online.com/rampart/Sagitta/installs/SagittaDll.msi"

$software = "SagittaDll"
$installed = (Get-ItemProperty HKLM:\SOFTWARE\Classes\Installer\Products\* | Where { $_.ProductName -match $software }) -ne $null

If(-Not $installed) {
	Write-Host "'$software' NOT is installed."

mkdir C:\hilb\ -ErrorAction SilentlyContinue
# Download Installer to 'C:\HILB\'
Invoke-RestMethod -Uri $DownloadURL -OutFile C:\hilb\SagittaDll.msi
#Sleep for download to complete
Start-sleep -seconds 10



# Silently Install the Application

Start-Process 'C:\hilb\SagittaDll.msi' -ArgumentList '/qn'

#Sleep for installation to complete
Start-sleep -seconds 10

} Else { Write-host $software already installed Skipping. } 



#Sagitta Software preqes for rampart
$DownloadURL = "https://wp40.sagitta-online.com/rampart/Sagitta/installs/AMSVBRuntime.msi"

$software = "AMSVBRuntime"
$installed = (Get-ItemProperty HKLM:\SOFTWARE\Classes\Installer\Products\* | Where { $_.ProductName -match $software }) -ne $null

If(-Not $installed) {
	Write-Host "'$software' NOT is installed."

mkdir C:\hilb\ -ErrorAction SilentlyContinue
# Download Installer to 'C:\HILB\'
Invoke-RestMethod -Uri $DownloadURL -OutFile C:\hilb\AMSVBRuntime.msi
#Sleep for download to complete
Start-sleep -seconds 10



# Silently Install the Application

Start-Process 'C:\hilb\AMSVBRuntime.msi' -ArgumentList '/qn'

#Sleep for installation to complete
Start-sleep -seconds 10

} Else { Write-host $software already installed Skipping. } 


#Sagitta Software preqes for rampart
$DownloadURL = "https://wp40.sagitta-online.com/rampart/Sagitta/installs/ImportIt.msi"

$software = "ImportIt"
$installed = (Get-ItemProperty HKLM:\SOFTWARE\Classes\Installer\Products\* | Where { $_.ProductName -match $software }) -ne $null

If(-Not $installed) {
	Write-Host "'$software' NOT is installed."

mkdir C:\hilb\ -ErrorAction SilentlyContinue
# Download Installer to 'C:\HILB\'
Invoke-RestMethod -Uri $DownloadURL -OutFile C:\hilb\ImportIt.msi
#Sleep for download to complete
Start-sleep -seconds 10



# Silently Install the Application

Start-Process 'C:\hilb\ImportIt.msi' -ArgumentList '/qn'

#Sleep for installation to complete
Start-sleep -seconds 10

} Else { Write-host $software already installed Skipping. } 


#add url for IE
Set-Location "HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings"
Set-Location ZoneMap\Domains
New-Item sagitta-online.com
Set-Location sagitta-online.com
New-Item www
Set-Location www
New-ItemProperty . -Name http -Value 2 -Type DWORD

#add url for IE
Set-Location "HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings"
Set-Location ZoneMap\Domains
New-Item sagitta-online.com
Set-Location sagitta-online.com
New-Item wp40
Set-Location wp40
New-ItemProperty . -Name https -Value 2 -Type DWORD
