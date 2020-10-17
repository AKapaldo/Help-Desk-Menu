<#
.Synopsis
Help Desk Ninja is a collection of scripts to make common help desk tasks quicker.
.Description
Help Desk Ninja is a collection of scripts to make common help desk tasks quicker.
.Parameter ComputerName
System number of the remote computer to connect to.
.Notes
To install this module, place the file in your c:\Users\<Username>\Documents\MicrosoftPowerShell\Modules\helpdesk\ folder. You will need to create these folders. They do not exist by default.
Author: A. Kapaldo
#>
Function Pause ($Message="Press any key to continue..."){ 
    "" 
    Write-Host $Message 
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown") | Out-Null
    Load-Helpdesk -ComputerName $ComputerName}
function Load-Helpdesk {
[cmdletbinding()]
param(	
[Parameter()]
[string]$ComputerName=$(read-host -prompt "Enter the System Number")
	)
Process {	
Clear-Host
	$foregroundColor = "Green"
	$usercolor = "Yellow"
	$Title = "Help Desk Ninja Center"
	$Version = "- v1.0"
	$Host.UI.RawUI.WindowTitle = "Help Desk Ninja Center - v1.0"
IF ($COmputerName -eq "") {
Write-Warning "Invalid System Number"}
ELSE {
$TestConnection = Test-Connection -ComputerName $ComputerName -Count 1 -BufferSize 16 -ErrorAction SilentlyContinue
IF (!($TestConnection)) {
Write-Warning "Unable to Connect to $ComputerName"}
Else{
Clear-Host
    ##########################################
    #         Help Desk Ninja Menu           #
    ##########################################
    Write-Host -NoNewLine `n"#####################################################" -ForeGroundColor $foregroundcolor
    Write-Host -NoNewLine `n"# " -ForeGroundColor $foregroundcolor
    Write-Host -NoNewLine "========= "
    Write-Host -NoNewLine "$Title $Version" -ForeGroundColor $foregroundcolor
    Write-Host -NoNewLine " ========="
    Write-Host -NoNewLine " #" -ForeGroundColor $foregroundcolor
    Write-Host -NoNewLine `n"# " -ForeGroundColor $foregroundcolor
    Write-Host -NoNewLine "======== Successfully connected to $HostName ======="
    Write-Host -NoNewLine " #" -ForeGroundColor $foregroundcolor
    Write-Host -NoNewLine `n"#####################################################" -ForeGroundColor $foregroundcolor
    Write-Host `n"Type 'q' or hit enter to drop to shell"`n
    
    Write-Host -NoNewLine "<" -foregroundcolor $foregroundColor
    Write-Host -NoNewLine "Help Desk Ninja Functions"
    Write-Host -NoNewLine ">" -foregroundcolor $foregroundColor
    Write-Host -NoNewLine "["
    Write-Host -NoNewLine "A" -foregroundcolor $foregroundColor
    Write-Host -NoNewLine "]"

    Write-Host -NoNewLine `t`n "A1 - " -foregroundcolor $foregroundcolor
    Write-host -NoNewLine "Change Remote Machine"`n`n
	
    Write-Host -NoNewLine "<" -foregroundcolor $foregroundColor
    Write-Host -NoNewLine "Certificates"
    Write-Host -NoNewLine ">" -foregroundcolor $foregroundColor
    Write-Host -NoNewLine "["
    Write-Host -NoNewLine "C" -foregroundcolor $foregroundColor
    Write-Host -NoNewLine "]"


    Write-Host -NoNewLine `t`n "C1 - " -foregroundcolor $foregroundcolor
    Write-host -NoNewLine "Add Certificates Console to User's Desktop"`n`n

    Write-Host -NoNewLine "<" -foregroundcolor $foregroundColor
    Write-Host -NoNewLine "DameWare"
    Write-Host -NoNewLine ">" -foregroundcolor $foregroundColor
    Write-Host -NoNewLine "["
    Write-Host -NoNewLine "D" -foregroundcolor $foregroundColor
    Write-Host -NoNewLine "]"
    
    Write-Host -NoNewLine `t`n "D1 - " -foregroundcolor $foregroundcolor
    Write-host -NoNewLine "Reset DameWare"`n`n

    Write-Host -NoNewLine "<" -foregroundcolor $foregroundColor
    Write-Host -NoNewLine "Host Information/Status"
    Write-Host -NoNewLine ">" -foregroundcolor $foregroundColor
    Write-Host -NoNewLine "["
    Write-Host -NoNewLine "H" -foregroundcolor $foregroundColor
    Write-Host -NoNewLine "]"
    
    Write-Host -NoNewLine `t`n "H1 - " -foregroundcolor $foregroundcolor
    Write-host -NoNewLine "System Info/Uptime"
    Write-Host -NoNewLine `t`n "H2 - " -foregroundcolor $foregroundcolor
    Write-host -NoNewLine "Group Policy Update"
    Write-Host -NoNewLine `t`n "H3 - " -foregroundcolor $foregroundcolor
    Write-host -NoNewLine "AD User Info/Password Expiration"
    
    Write-Host -NoNewLine "<" -foregroundcolor $foregroundColor
    Write-Host -NoNewLine "Install"
    Write-Host -NoNewLine ">" -foregroundcolor $foregroundColor
    Write-Host -NoNewLine "["
    Write-Host -NoNewLine "I" -foregroundcolor $foregroundColor
    Write-Host -NoNewLine "]"
    
    Write-Host -NoNewLine `t`n "I1 - " -foregroundcolor $foregroundcolor
    Write-host -NoNewLine "Install Content Manager"
    Write-Host -NoNewLine `t`n "I2 - " -foregroundcolor $foregroundcolor
    Write-host -NoNewLine "Place Process Manager on System"
    Write-Host -NoNewLine `t`n "I3 - " -foregroundcolor $foregroundcolor
    Write-host -NoNewLine "Place User Tools Batch File on Desktop"
    Write-Host -NoNewLine `t`n "I4 - " -foregroundcolor $foregroundcolor
    Write-host -NoNewLine "Place JAMIS Batch File on Desktop"`n`n

    Write-Host -NoNewLine "<" -foregroundcolor $foregroundColor
    Write-Host -NoNewLine "Registry Edits"
    Write-Host -NoNewLine ">" -foregroundcolor $foregroundColor
    Write-Host -NoNewLine "["
    Write-Host -NoNewLine "R" -foregroundcolor $foregroundColor
    Write-Host -NoNewLine "]"

    Write-Host -NoNewLine `t`n "R1 - " -foregroundcolor $foregroundcolor
    Write-host -NoNewLine "Add Username Hint Box Back"
    Write-Host -NoNewLine `t`n "R2 - " -foregroundcolor $foregroundcolor
    Write-Host -NoNewLine "Force Outlook Skype Add In to Load on Startup"
    Write-Host -NoNewLine `t`n "R3 - " -foregroundcolor $foregroundcolor
    Write-host -NoNewLine "Force Outlook Entrust Add In to Load on Startup"`n`n

    Write-Host -NoNewLine "<" -foregroundcolor $foregroundColor
    Write-Host -NoNewLine "Storage"
    Write-Host -NoNewLine ">" -foregroundcolor $foregroundColor
    Write-Host -NoNewLine "["
    Write-Host -NoNewLine "S" -foregroundcolor $foregroundColor
    Write-Host -NoNewLine "]"

    Write-Host -NoNewLine `t`n "S1 - " -foregroundcolor $foregroundcolor
    Write-host -NoNewLine "Map a Z:\ Drive"
    Write-Host -NoNewLine `t`n "S2 - " -foregroundcolor $foregroundcolor
    Write-host -NoNewLine "Map any Drive"
    Write-Host -NoNewLine `t`n "S3 - " -foregroundcolor $foregroundcolor
    Write-host -NoNewLine "Z: Drive Map Batch File"`n`n

    $sel = Read-Host "Which option?"

    Switch ($sel) {
		"A1" {
		Clear-Host
		$ComputerName = Read-Host "Enter New Remote Machine"
		Pause
		}
       
	    "C1" {
        Copy-CertConsole -Computerame $ComputerName
		Pause
		}
		
		"D1" {
		Reset-DameWare -ComputerName $ComputerName
		Pause
		}

        "H1" {
		Function Get-SystemInfo {
		$wmi = Get-WmiObject -class Win32_OperatingSystem -computer $ComputerName
		$LBTime = $wmi.ConvertToDateTime($wmi.Lastbootuptime)
		[TimeSpan]$uptime = New-TimeSpan $LBTime $(get-date)
        Write-Host "$ComputerName has been up for:"
		Write-Output "$($uptime.days) Days $($uptime.hours) Hours $($uptime.minutes) Minutes $($uptime.seconds) Seconds"
		SystemInfo /s $HostName | findstr /i /c:"Host Name" /c:"OS Name" /c:"OS Version" /c:"Original Install Date" /c:"System Boot Time" /c:"System Up Time" /c:"System Manufacturer" /c:"System Model" /c:"System Type" /c:"Total Physical Memory"
			
    Pause      
}
	}
        
        "H2" {
        $result1 = Invoke-Command -ComputerName $ComputerName -ScriptBlock {Start-Process cmd.exe -ArgumentList "/c klist -lh 0 -li 0x3e7 purge"}
        Write-Host "$HostName - Group Policy Update"
		if ($result1 -ne $null)
		{
			if ($result1.ReturnValue -eq 0){
			    Write-Host "Kerberos Ticket Purge Ran Successfully!"
                Write-Host "Starting Group Policy Update."
                $result2 = Invoke-Command -ComputerName $ComputerName -ScriptBlock {Start-Process cmd.exe -ArgumentList "/c echo y | gpupdate /force /wait:0"
                    if ($result2.ReturnValue -eq 0) {
                    Write-Host "Group Policy Updated Successfully!"
                    }
                    ELSE{Write-Warning "Group Policy Failed to Update."}
                }

			}
            ELSE{Write-Warning "Group Policy Failed to Update."}
		}
        ELSE{Write-Warning "Group Policy Failed to Update."}
		Pause
        }
		
		"H3" {
		Clear-Host
		$User = Invoke-Command -ComputerName $ComputerName -ScriptBlock {(Get-WMIObject -class Win32_ComputerSystem | select -expand username).substring(3)}
		$Current = Read-Host "Run with $User ? Y or N"
		IF ($Current -eq "Y" -or $Current -eq "Yes"){
		Get-PassExpiry -User $User
		}
		Else {Get-PassExpiry}
		Pause
		}
		

        "I1" {
		Install-Software -ComputerName $ComputerName -SW 1
		Pause
          }
		"I2" {
		Install-Software -ComputerName $ComputerName -SW 2
		Pause
          }
		"I3" {
		$User = Invoke-Command -ComputerName $ComputerName -ScriptBlock {(Get-WMIObject -class Win32_ComputerSystem | select -expand username).substring(3)}
		Push-UserTools -ComputerName $ComputerName -User $User
		Pause
          }
		"I4" {
		Install-Software -ComputerName $ComputerName -SW 3
		Pause
          }
        
		
		"R1" {
		Clear-Host
		$CurrentValue = Invoke-Command -ComputerName $ComputerName -ScriptBlock {Get-ItemProperty 'Registry::Hkey_local_machine\SOFTWARE\Policies\Microsoft\Windows\SmartCardCredentialProvider' | Select -expand X509HintsNeeded}
		IF ($CurrentValue -eq 1){Write-Host "X509HintsNeeded is already set to $CurrentValue"}
		ELSE {
		$Change = Read-Host "X509HintsNeeded is set to $CurrentValue - Change to 1/enable? Y or N"
			IF ($Change -eq "Y" -OR $Change -eq "Yes"){
			Invoke-Command -ComputerName $ComputerName -ScriptBlock {Set-ItemProperty 'Registry::Hkey_local_machine\SOFTWARE\Policies\Microsoft\Windows\SmartCardCredentialProvider' -Name X509HintsNeeded -value 1 -PassThru}
			Pause}
			ELSE {Write-Warning "X509HintsNeeded not changed."}
		
		}}
		"R2" {
		Clear-Host
		$CurrentValue = Invoke-Command -ComputerName $ComputerName -ScriptBlock {Get-ItemProperty 'Registry::Hkey_local_machine\WOW6432Node\Microsoft\Office\Outlook\Addins\UCAddin.LyncAddin.1' | Select -Expand loadbehavior}
		IF ($CurrentValue -eq 3){Write-Host "LoadBehavior is already set to $CurrentValue"}
		ELSE {
		$Change = Read-Host "LoadBehavior is set to $CurrentValue - Change to 3/On Startup? Y or N"
			IF ($Change -eq "Y" -OR $Change -eq "Yes"){
			Invoke-Command -ComputerName $ComputerName -ScriptBlock {Set-ItemProperty 'Registry::Hkey_local_machine\WOW6432Node\Microsoft\Office\Outlook\Addins\UCAddin.LyncAddin.1' -Name loadbehavior -value 3 -PassThru}
			Pause}
			ELSE {Write-Warning "LoadBehavior not changed."}
		}}
		"R3" {
		Clear-Host
		$CurrentValue1 = Invoke-Command -ComputerName $ComputerName -ScriptBlock {Get-ItemProperty 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\Outlook\Addins\Entrust.SMIMEFormat' | Select -Expand loadbehavior}
		$CurrentValue2 = Invoke-Command -ComputerName $ComputerName -ScriptBlock {Get-ItemProperty 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Office\Outlook\Addins\Entrust.SMIMEFormat' | Select -Expand loadbehavior}
		IF ($CurrentValue1 -eq 3 -AND $CurrentValue2 -eq 3) {Write-Host "LoadBehavior is already set to $CurrentValue"}
		ELSE {
		$Change = Read-Host "LoadBehavior set to $CurrentValue1 and $CurrentValue2 - Change to 3/On Startup? Y or N."}
			IF ($Change -eq "Y" -OR $Change -eq "Yes"){
				IF ($CurrentValue1 -eq 3 -AND $CurrentValue2 -ne 3)
				{Invoke-Command -ComputerName $ComputerName -ScriptBlock {Set-ItemProperty 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Office\Outlook\Addins\Entrust.SMIMEFormat' -Name loadbehavior -Value 3 -PassThru}
				Pause}
				ELSEIF ($CurrentValue1 -ne 3 -AND $CurrentValue2 -eq 3) {
				Invoke-Command -ComputerName $ComputerName -ScriptBlock {Set-ItemProperty 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\Outlook\Addins\Entrust.SMIMEFormat' -Name loadbehavior -Value 3 -PassThru}
				Pause
				}
				ELSEIF ($CurrentValue1 -ne 3 -AND $CurrentValue2 -ne 3) {Invoke-Command -ComputerName $ComputerName -ScriptBlock {
				Set-ItemProperty 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Office\Outlook\Addins\Entrust.SMIMEFormat' -Name loadbehavior -Value 3 -PassThru
				Set-ItemProperty 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\Outlook\Addins\Entrust.SMIMEFormat' -Name loadbehavior -Value 3 -PassThru
				Pause}}
				ELSE {Write-Warning "Error: Script Failed"}
			}
			ELSE {Write-Warning "LoadBehavior not changed."}
		}
        "S1" {
            Clear-Host
            $User = Invoke-Command -ComputerName $ComputerName -ScriptBlock {(Get-WMIObject -class Win32_ComputerSystem | select -expand username).SubString(3)}
			$Path = Get-ADUser -Identity $User -Properties HomeDirectory | select -expand HomeDirectory
            $Cont = Read-Host "Map $Path to Drive Z:\? Y or N"
				IF ($Cont -eq "Y" -OR $Cont -eq "Yes"){
				Invoke-Command -ComputerName $ComputerName -ScriptBlock {New-PSDrive -Name "Z" -Root $Path -Persist; Get-PSDrive Z}}
				ELSE {Write-Host "Drive Z:\ Not Mapped."}
				Pause
			}
		"S2" {
		Clear-Host
		Invoke-Command -ComputerName $ComputerName -ScriptBlock {Get-PSDrive}
		$Drive = Read-Host "What Drive Letter? Letter only."
		$Path = Read-Host "What Path? Start with \\"
		Invoke-Command -ComputerName $ComputerName -ScriptBlock {New-PSDrive -Name "$Drive" -Root $Path -Persist; Get-PSDrive}
		Pause
		}
       
	   "S3" {
		Clear-Host
		$User = Invoke-Command -ComputerName $ComputerName -ScriptBlock {(Get-WMIObject -class Win32_ComputerSystem | select -expand username).substring(3)}
		$Path = Get-ADUser -Identity $User -Properties HomeDirectory | select -expand HomeDirectory
		Invoke-Command -ComputerName $ComputerName -ScriptBlock {Set-Content "C:\Users\$User\Desktop\Map Z Drive.bat" @"
		cls
		title Map Z:\ Drive
		echo Mapping Z:\ Drive
		echo.
		Net Use Z: $Path /Persistent:yes
		title Z:\ Drive Mapped
		echo Z:\ Drive Mapped
		echo.
"@		}
		Pause
	}
        {($_ -like "*q*") -or ($_ -eq "")} {
            
            Write-Host `n"No input or 'q' seen... dropping to shell" -foregroundColor $foregroundColor
            Write-Host "Type Load-Helpdesk to load the menu again"`n
            
            
        }          
            
    }
    ##########################################
    #        End Ninja Center Menu           #
    ##########################################


    }
	
		
	}}}
