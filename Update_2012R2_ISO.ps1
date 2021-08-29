<#
    .SYNOPSIS
        Script To Update Windows Server ISO Image.
    .DESCRIPTION
        Update Windows Server ISO For Lastes Updates,
        Using xcopy, PSWindowsUpdate, kbupdate, Dism, 
        And Oscdimg from Assessment and Deployment Kit
    .NOTES
    Version    : 0.0.2 experimental
    Author     : CeLTic
    Created on : 2021-08-19
	Updated on : 2021-08-29
    License    : GPL3
    Tested on  :
    SW_DVD9_Windows_Svr_Std_and_DataCtr_2012_R2_64Bit_English_-4_MLF_X19-82891
	9600.17050.WINBLUE_REFRESH.140317-1640_X64FRE_SERVER_EVAL_EN-US-IR3_SSS_X64FREE_EN-US_DV9

#>

#Requires -RunAsAdministrator

#Enable TLS1.2
#Soruce: https://docs.microsoft.com/en-us/dotnet/framework/network-programming/tls
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$OperatingSystem = 'Windows Server 2012 R2'

$DVD = "D:\"
$Temp = "C:\Temp"
$Dism_Temp = "$Temp\Dism"
$Src_ISO = "$$Temp\Src_ISO"
$Updates = "$Temp\Updates"
$Final_ISO = "$Temp\ISO"
$Logs = "$Temp\Logs"


$BGColor_Info = 'DarkGreen'
$BGColor_Warn = 'Red'

$Supported_Systems = @( 'Microsoft Windows Server 2012 R2 Standard',
'Microsoft Windows Server 2012 R2 Standard Evaluation',
'Microsoft Windows Server 2012 R2 Datacenter Evaluation'
)

$Supported_InstallationType = @( 'Server')

<#
Index       : 1
Name        : Windows Server 2012 R2 SERVERSTANDARDCORE
Description : Windows Server 2012 R2 SERVERSTANDARDCORE
Size        : 6,898,373,863 bytes

Index       : 2
Name        : Windows Server 2012 R2 SERVERSTANDARD
Description : Windows Server 2012 R2 SERVERSTANDARD
Size        : 22,658,273,835 bytes

Index       : 3
Name        : Windows Server 2012 R2 SERVERDATACENTERCORE
Description : Windows Server 2012 R2 SERVERDATACENTERCORE
Size        : 6,871,511,192 bytes

Index       : 4
Name        : Windows Server 2012 R2 SERVERDATACENTER
Description : Windows Server 2012 R2 SERVERDATACENTER
Size        : 12,065,366,117 bytes
#>


Function Format-FileSize() {
    Param ([int]$size)
    If     ($size -gt 1TB) {[string]::Format("{0:0.00} TB", $size / 1TB)}
    ElseIf ($size -gt 1GB) {[string]::Format("{0:0.00} GB", $size / 1GB)}
    ElseIf ($size -gt 1MB) {[string]::Format("{0:0.00} MB", $size / 1MB)}
    ElseIf ($size -gt 1KB) {[string]::Format("{0:0.00} kB", $size / 1KB)}
    ElseIf ($size -gt 0)   {[string]::Format("{0:0.00} B", $size)}
    Else                   {""}
}

Write-Host -BackgroundColor $BGColor_Info "Creating Directories..."
$q = New-Item -Path $Dism_Temp -Type 'Directory' -Force
$q = New-Item -Path $Src_ISO -Type 'Directory' -Force
$q = New-Item -Path $Updates -Type 'Directory' -Force
$q = New-Item -Path $Final_ISO -Type 'Directory' -Force
$q = New-Item -Path $Logs -Type 'Directory' -Force
Write-Host -BackgroundColor $BGColor_Info "Done.`n`n"


Write-Host -BackgroundColor $BGColor_Warn "This is experiment/experminental version!!!"
Write-Host -BackgroundColor $BGColor_Warn "You can use it, but it’s at your own risk.!!!`n`n"

#TODO
#TimeZone ?

Write-Host -BackgroundColor $BGColor_Info "Checking version of powershell..."
#https://download.microsoft.com/download/6/F/5/6F5FF66C-6775-42B0-86C4-47D41F2DA187/Win8.1AndW2K12R2-KB3191564-x64.msu
if ($PSVersionTable.PSVersion.Major -ge 5){
    Write-Host -BackgroundColor $BGColor_Info "Version 5 Found. Continuing"
}
else {
	Write-Host -BackgroundColor $BGColor_Warn "Not Found!!!"
    Write-Host -BackgroundColor $BGColor_Info "Installing Powershell Version 5..."
    Invoke-WebRequest https://go.microsoft.com/fwlink/?linkid=839516 -OutFile $Updates\Win8.1AndW2K12R2-KB3191564-x64.msu
    Write-Host -BackgroundColor $BGColor_Warn "Computer Will Be Restarted In 30 Seconds"
    Start-Process -Wait wusa.exe "/quiet /forcerestart $Updates\Win8.1AndW2K12R2-KB3191564-x64.msu "
    Break 
}

Write-Host -BackgroundColor $BGColor_Info "Checking Operating System Support..."
$System = (Get-WmiObject -class Win32_OperatingSystem).Caption
if ( -not ($System -in $Supported_Systems) ){ 
    Write-Host -BackgroundColor $BGColor_Warn "Sorry, $System Is Not Suppoted"
    Break
}
Write-Host -BackgroundColor $BGColor_Info "System: $System"
Write-Host -BackgroundColor $BGColor_Info "Ok. Supported`n`n"

Write-Host -BackgroundColor $BGColor_Info "Checking Type of Operating System..."
$InstallationType = (Get-ItemProperty -Path "HKLM:\Software\Microsoft\Windows NT\CurrentVersion" -Name "InstallationType").InstallationType
if ( -not ($InstallationType-in $Supported_InstallationType) ){ 
    Write-Host -BackgroundColor $BGColor_Warn "Sorry, Type: $InstallationType Is Not Suppoted"
    Break
}
Write-Host -BackgroundColor $BGColor_Info "Ok. Yout Version is: $InstallationType `n`n"

Write-Host -BackgroundColor $BGColor_Info "Checking Free Space..."
$Free =(Get-WMIObject Win32_Logicaldisk -filter "deviceid='C:'").FreeSpace
if (-not ($Free -ge 15000000000)) { 
    Write-Host -BackgroundColor $BGColor_Warn "Sorry, You Need min. 15GB Free Space"
    Break 
}
Write-Host -BackgroundColor $BGColor_Info "Space OK.`n`n"

Write-Host -BackgroundColor $BGColor_Info "Installing Required Modules From PSGallery"
Write-Host -BackgroundColor $BGColor_Info "Installing NuGet..."
$q = Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force
Write-Host -BackgroundColor $BGColor_Info "Installed [NuGet]`n`n"
Write-Host -BackgroundColor $BGColor_Info "Installing PSWindowsUpdate..."
$q = Install-Module -Name PSWindowsUpdate -Force
Write-Host -BackgroundColor $BGColor_Info "Installed [PSWindowsUpdate]`n`n"
Write-Host -BackgroundColor $BGColor_Info "Installing kbupdate..."
$q = Install-Module -Name kbupdate -Force
Write-Host -BackgroundColor $BGColor_Info "Installed [kbupdate]`n`n"

Write-Host -BackgroundColor $BGColor_Info "Installing Assessment and Deployment Kit"
wget https://go.microsoft.com/fwlink/?linkid=2120254 -Outfile C:\Temp\adksetup.exe
Start-Process $Temp\adksetup.exe "/quiet /features OptionId.DeploymentTools"
$env:Path += ";C:\Program Files (x86)\Windows Kits\10\Assessment and Deployment Kit\Deployment Tools\amd64\Oscdimg"
Write-Host -BackgroundColor $BGColor_Info "Installing in background [ADK]`n`n"

Write-Host -BackgroundColor $BGColor_Info "Copying ISO To Hard Drive..."
#TODO
#Copy-ItemWithProgress
xcopy /E /H $DVD $Src_ISO /Y
Write-Host -BackgroundColor $BGColor_Info "Copying Finished.`n`n"

Write-Host -BackgroundColor $BGColor_Info "Searching For Updates..."
Write-Host -BackgroundColor $BGColor_Info "This Is Long Process, Be Patient"
Write-Host -BackgroundColor $BGColor_Info "This Will Take Approx Max ~0.5-2 Hours"
$Scanned_Updates = Get-WindowsUpdate
$Scanned_Updates | Select-Object KB,Size,Title,LastDeploymentChangeTime |Format-List > $Logs\Scanned_Updates.txt
$Updates_Count = [int]$Scanned_Updates.Count
Write-Host -BackgroundColor $BGColor_Info "Found: $Updates_Count"
Write-Host -BackgroundColor $BGColor_Info "Searching Finished.`n`n"


Write-Host -BackgroundColor $BGColor_Info "Downloading Updates..."
Get-KbUpdate -Name $Scanned_Updates.KB -Architecture x64 -OperatingSystem $OperatingSystem | Save-KbUpdate -Path $Updates
Write-Host -BackgroundColor $BGColor_Info "Downloading Finished.`n`n"
Get-ChildItem $Updates | Select-Object Name,Length,LastWriteTime | Format-List > $Logs\Downloaded_Updates.txt
Remove-Item -Path $Updates\*.exe
$Downloaded_Updates = (Get-ChildItem $Updates).Name
$Downloaded_Count = [int]$Downloaded_Updates.Count
#TODO
# Sort Updates by time
#| Sort-Object LastWriteTime


Write-Host -BackgroundColor $BGColor_Info "Mounting WIM Image..."
#Dism /Get-ImageInfo /imagefile:$Src_ISO\sources\install.wim
#TODO
#Searching by Index
Get-WindowsImage -ImagePath $Src_ISO\sources\install.wim
#Dism /Mount-Image /ImageFile:$Src_ISO\sources\install.wim  /MountDir:$Dism_Temp /Name:"Windows Server 2012 R2 SERVERSTANDARD"
Mount-WindowsImage -ImagePath $Src_ISO\sources\install.wim -Path $Dism_Temp -Index 2 -LogPath $Logs\Mount_Wim.txt
Write-Host -BackgroundColor $BGColor_Info "Mounted.`n`n"

Write-Host -BackgroundColor $BGColor_Info "Applying Updates..."
$i = 1
ForEach ($File in $Downloaded_Updates){
    Write-Host -BackgroundColor $BGColor_Info "Applying: $i / $Downloaded_Count"
	Write-Host -BackgroundColor $BGColor_Info "File: $File"
	#Get-ChildItem | Select-Object Name, @{Name="Size";Expression={Format-FileSize($_.Length)}}
	#Write-Host((Get-Item C:\Temp\Updates\$file).length)
	Get-Item $Updates\$File | Select-Object Name, @{Name="Size";Expression={Format-FileSize($_.Length)}}
    #Dism /Image:$Dism_Temp /Add-Package /PackagePath:$Updates\$File
	Add-WindowsPackage -Path $Dism_Temp -PackagePath $Updates\$File -LogPath $Logs\Pkg_$File.txt #-PreventPending
	$i++
}
Write-Host -BackgroundColor $BGColor_Info "Applaying Finished.`n`n"

Write-Host -BackgroundColor $BGColor_Info "Finalizing. Commiting/Saving Channges..."
#Dism /Unmount-Image /MountDir:$Dism_Temp  /Commit
Dismount-WindowsImage -Path $Dism_Temp -Save -LogPath $Logs\Dismount_Wim.txt
Write-Host -BackgroundColor $BGColor_Info "Finalizing Done.`n`n"

Write-Host -BackgroundColor $BGColor_Info "Creating ISO File"
Start-Process -Wait oscdimg.exe "-m -o -u2 -udfver102 -bootdata:2#p0,e,b$Src_ISO\boot\etfsboot.com#pEF,e,b$Src_ISO\efi\microsoft\boot\efisys.bin $Src_ISO\ $Final_ISO\Win2012R2.iso"
Write-Host -BackgroundColor $BGColor_Info "Creating ISO Finished in: $Final_ISO `n`n"
