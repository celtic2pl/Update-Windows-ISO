<#
    .SYNOPSIS
        Script To Update Windows Server ISO Image.
    .DESCRIPTION
        Update Windows Server ISO For Lastes Updates,
        Using xcopy, PSWindowsUpdate, kbupdate, Dism, 
        And Oscdimg from Assessment and Deployment Kit
    .NOTES
    Version    : 0.0.0 experimental
    Author     : CeLTic
    Created on : 2021-08-19
    License    : GPL3
    Tested on  :
    SW_DVD9_Windows_Svr_Std_and_DataCtr_2012_R2_64Bit_English_-4_MLF_X19-82891

#>

#Requires -RunAsAdministrator

#Enable TLS1.2
#Soruce: https://docs.microsoft.com/en-us/dotnet/framework/network-programming/tls
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12


$Temp = 'C:\Temp'
$Dism_Temp = 'C:\Temp\Dism_Temp'
$Src_ISO = 'C:\Temp\Src_ISO'
$Updates = 'C:\Temp\Updates'
$Final_ISO ='C:\Temp\ISO'
$DVD = 'D:\'

$Features = @(
'Telnet-Client', 
'NET-Framework-Core'
)

$BGColor = 'DarkGreen'

Write-Host -BackgroundColor Red "This is experiment/experminental version!!!"
Write-Host -BackgroundColor Red "You can use it, but it’s at your own risk.!!!"
Write-Host
Write-Host

#Checking for powershell 5
#https://download.microsoft.com/download/6/F/5/6F5FF66C-6775-42B0-86C4-47D41F2DA187/Win8.1AndW2K12R2-KB3191564-x64.msu
Write-Host -BackgroundColor $BGColor "Checkin version of powershell"
if ($PSVersionTable.PSVersion.Major -ge 5){
    Write-Host -BackgroundColor $BGColor "Version 5 Found. Continuing"
}
else {
    Write-Host -BackgroundColor $BGColor "Installing Powershell Version 5"
    New-Item -Path $Temp  -Type 'Directory' -Force
    Invoke-WebRequest https://go.microsoft.com/fwlink/?linkid=839516 -OutFile $Temp\Win8.1AndW2K12R2-KB3191564-x64.msu
    [System.Windows.MessageBox]::Show('Computer Will Be Restarted In 30 Seconds')
    Write-Host -BackgroundColor $BGColor "Computer Will Be Restarted In 30 Seconds"
    Start-Process -Wait wusa.exe "/quiet /forcerestart $Temp\Win8.1AndW2K12R2-KB3191564-x64.msu "
    Break 
}

#Check Operating System
if ( -not ((Get-WmiObject -class Win32_OperatingSystem).Caption -eq ("Microsoft Windows Server 2012 R2 Standard")) ){ 
    Write-Host -BackgroundColor $BGColor "Only Windows 2012 R2 Is Now Supported"
}

#Check Free Space
$Free =(Get-WMIObject Win32_Logicaldisk -filter "deviceid='C:'").FreeSpace
if (-not ($Free -ge 15000000000)) { 
    Write-Host -BackgroundColor $BGColor "Sorry, You Need min. 15GB Free Space"
    Break 
}

#Create Directories
New-Item -Path $Dism_Temp -Type 'Directory' -Force
New-Item -Path $Src_ISO -Type 'Directory' -Force
New-Item -Path $Updates -Type 'Directory' -Force
New-Item -Path $Final_ISO -Type 'Directory' -Force

#Install-WindowsFeature -Name Telnet-Client, NET-Framework-Core
#Write-Host -BackgroundColor $BGColor "Installing Selected Features"
#ForEach ($Feature in $Features){
#    if (-not ((Get-WindowsFeature -Name $Feature).Installed)){
#        Write-Host -BackgroundColor $BGColor "Installing: $Feature"
#        Install-WindowsFeature -Name $Feature -Restart
#    }
#}

Write-Host -BackgroundColor $BGColor "Installing Required Modules From PSGallery"
Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force
Install-Module -Name PSWindowsUpdate -Force
Install-Module -Name kbupdate -Force

Write-Host -BackgroundColor $BGColor "Installing Assessment and Deployment Kit"
wget https://go.microsoft.com/fwlink/?linkid=2120254 -Outfile C:\Temp\adksetup.exe
Start-Process -Wait C:\Temp\adksetup.exe "/quiet /features OptionId.DeploymentTools"
$env:Path += ";C:\Program Files (x86)\Windows Kits\10\Assessment and Deployment Kit\Deployment Tools\amd64\Oscdimg"


Write-Host -BackgroundColor $BGColor "Searching For Updates..."
Write-Host -BackgroundColor $BGColor "This Is Long Process, Be Patient"
Write-Host -BackgroundColor $BGColor "This Will Take Approx Max ~1-2 Hours"
$Scanned_Updates = Get-WindowsUpdate

Write-Host -BackgroundColor $BGColor "Downloading Updates..."
Get-KbUpdate -Name $Scanned_Updates.KB -Architecture x64 -OperatingSystem 'Windows Server 2012 R2' | Save-KbUpdate -Path $Updates
Remove-Item -Path $Updates\*.exe
$Downloaded_Updates = (Get-ChildItem C:\Temp\Updates).Name

Write-Host -BackgroundColor $BGColor "Copying ISO To Hard Drive"
xcopy /E /H $DVD $Src_ISO /Y

Write-Host -BackgroundColor $BGColor "Mounting WIM Image"
Dism /Get-ImageInfo /imagefile:$Src_ISO\sources\install.wim
Dism /Mount-Image /ImageFile:$Src_ISO\sources\install.wim  /MountDir:$Dism_Temp /Name:"Windows Server 2012 R2 SERVERSTANDARD"

Write-Host -BackgroundColor $BGColor "Customizing Image..."
Dism /Image:$Dism_Temp /Enable-Feature:"TelnetClient"
#Dism /Image:$Dism_Temp /Enable-Feature:"NetFx3ServerFeatures"
Dism /Image:$Dism_Temp /Enable-Feature:"WindowsServerBackup"
Dism /Image:$Dism_Temp /Enable-Feature:"InkAndHandwritingServices"
Dism /Image:$Dism_Temp /Enable-Feature:"ServerMediaFoundation"
Dism /Image:$Dism_Temp /Enable-Feature:"DesktopExperience" 

Write-Host -BackgroundColor $BGColor "Applaying Updates..."
ForEach ($File in $Downloaded_Updates){
    #Write-Host $Updates\$File
    Dism /Image:$Dism_Temp /Add-Package /PackagePath:$Updates\$File
}

#Adding Powershell5 To Image
$PS5Installler = "$Temp\Win8.1AndW2K12R2-KB3191564-x64.msu"
if ([System.IO.File]::Exists($PS5Installler)) {
    Dism /Image:$Dism_Temp /Add-Package /PackagePath:$PS5Installler
}

Write-Host -BackgroundColor $BGColor "Finalizing. Commiting Channges..."
Dism /Unmount-Image /MountDir:$Dism_Temp  /Commit

Write-Host -BackgroundColor $BGColor "Creating ISO File in:" $Final_ISO 
Start-Process -Wait oscdimg.exe "-m -o -u2 -udfver102 -bootdata:2#p0,e,b$Src_ISO\boot\etfsboot.com#pEF,e,b$Src_ISO\efi\microsoft\boot\efisys.bin $Src_ISO\ $Final_ISO\Win2012R2.iso"
