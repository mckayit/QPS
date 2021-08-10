#Sophos Installation Check

$software = "Sophos Anti-Virus";
$installed64 = (Get-ItemProperty HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where { $_.DisplayName -eq $software }) -ne $null
$installed32 = (Get-ItemProperty HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\* | Where { $_.DisplayName -eq $software }) -ne $null

If(-Not $installed64 -and -Not $installed32) {
	Write-Host "'$software' is NOT installed.";


} else {
	    Write-Host "'$software' is installed, preparing to remove."
    
   
        #Disable Tamper Protection (may require reboot)
    
        REG ADD "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\services\SAVService" /t REG_DWORD /v Start /d 0x00000004 /f
        REG ADD "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\Sophos MCS Agent" /t REG_DWORD /v Start /d 0x00000004 /f
        REG ADD "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\Sophos Endpoint Defense\TamperProtection\Config" /t REG_DWORD /v SAVEnabled /d 0 /f
        REG ADD "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\Sophos Endpoint Defense\TamperProtection\Config" /t REG_DWORD /v SEDEnabled /d 0 /f
        REG ADD "HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Sophos\SAVService\TamperProtection" /t REG_DWORD /v Enabled /d 0 /f
        
        #Stop All Sophos Services
        
        net stop "Sophos AutoUpdate Service"
        net stop "Sophos Agent"
        net stop "SAVService"
        net stop "SAVAdminService"
        net stop "Sophos Message Router"
        net stop "Sophos Web Control Service"
        net stop "swi_service"
        net stop "swi_update"
        net stop "SntpService"
        net stop "Sophos System Protection Service"
        net stop "Sophos Web Control Service"
        net stop "Sophos Endpoint Defense Service"
        
        
        
        #Stop Sophos Services check
        
        wmic service where "caption like '%Sophos%'" call stopservice
        
        
        
        #Kill all Sophos Services
        
        taskkill /f /im ALMon.exe
        taskkill /f /im ALsvc.exe
        taskkill /f /im swi_fc.exe
        taskkill /f /im swi_filter.exe
        taskkill /f /im spa.exe
        
        #Uninstall Sophos Network Threat Protection
        $SNTPVer = Get-ChildItem -Path HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall, HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall  |
            Get-ItemProperty |
                Where-Object {$_.DisplayName -match "Sophos Network Threat Protection" } |
                    Select-Object -Property DisplayName, UninstallString
        
        ForEach ($ver in $SNTPVer) {
        
            If ($ver.UninstallString) {
        
                $uninst = $ver.UninstallString
                Start-Process cmd "/c $uninst /qn REBOOT=SUPPRESS /PASSIVE" -NoNewWindow -verbose
            }
        
        }
        
        
        
        #Uninstall Sophos System Protection
        $SSPVer = Get-ChildItem -Path HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall, HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall  |
            Get-ItemProperty |
                Where-Object {$_.DisplayName -match "Sophos System Protection" } |
                    Select-Object -Property DisplayName, UninstallString
        
        ForEach ($ver in $SSPVer) {
        
            If ($ver.UninstallString) {
        
                $uninst = $ver.UninstallString
                Start-Process cmd "/c $uninst /qn REBOOT=SUPPRESS /PASSIVE" -NoNewWindow -verbose
            }
        
        }
        
        
        
        #Uninstall Sophos Client Firewall
        $SCFVer = Get-ChildItem -Path HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall, HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall  |
            Get-ItemProperty |
                Where-Object {$_.DisplayName -match "Sophos Client Firewall" } |
                    Select-Object -Property DisplayName, UninstallString
        
        ForEach ($ver in $SCFVer) {
        
            If ($ver.UninstallString) {
        
                $uninst = $ver.UninstallString
                Start-Process cmd "/c $uninst /qn REBOOT=SUPPRESS /PASSIVE" -NoNewWindow -verbose
            }
        
        }
        
        
        
        #Uninstall Sophos Anti-Virus
        $SAVVer = Get-ChildItem -Path HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall, HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall  |
            Get-ItemProperty |
                Where-Object {$_.DisplayName -match "Sophos Anti-Virus" } |
                    Select-Object -Property DisplayName, UninstallString
        
        ForEach ($ver in $SAVVer) {
        
            If ($ver.UninstallString) {
        
                $uninst = $ver.UninstallString
                Start-Process cmd "/c $uninst /qn REBOOT=SUPPRESS /PASSIVE" -NoNewWindow -verbose
            }
        
        }
        
        
        
        #Uninstall Sophos Remote Management System
        $SRMSVer = Get-ChildItem -Path HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall, HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall  |
            Get-ItemProperty |
                Where-Object {$_.DisplayName -match "Sophos Remote Management System" } |
                    Select-Object -Property DisplayName, UninstallString
        
        ForEach ($ver in $SRMSVer) {
        
            If ($ver.UninstallString) {
        
                $uninst = $ver.UninstallString
                Start-Process cmd "/c $uninst /qn REBOOT=SUPPRESS /PASSIVE" -NoNewWindow -verbose
            }
        
        }
        
        
        
        #Uninstall Sophos AutoUpdate
        $SAUVer = Get-ChildItem -Path HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall, HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall  |
            Get-ItemProperty |
                Where-Object {$_.DisplayName -match "Sophos AutoUpdate" } |
                    Select-Object -Property DisplayName, UninstallString
        
        ForEach ($ver in $SAUVer) {
        
            If ($ver.UninstallString) {
        
                $uninst = $ver.UninstallString
                Start-Process cmd "/c $uninst /qn REBOOT=SUPPRESS /PASSIVE" -NoNewWindow -verbose
            }
        
        }
        
        
        
        #Uninstall Sophos Endpoint Defense
        $SEDVer = Get-ChildItem -Path HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall, HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall  |
            Get-ItemProperty |
                Where-Object {$_.DisplayName -match "Sophos Endpoint Defense" } |
                    Select-Object -Property DisplayName, UninstallString
        
        ForEach ($ver in $SEDVer) {
        
            If ($ver.UninstallString) {
        
                $uninst = $ver.UninstallString
                cmd /c "$uninst"
            }
        }
            
        
        
        #Stop Sophos Services - Second Sweep
        
        wmic service where "caption like '%Sophos%'" call stopservice
        
        
        
        #Sophos Services Removal
        
        sc.exe delete "SAVService"
        sc.exe delete "SAVAdminService"
        sc.exe delete "Sophos Web Control Service"
        sc.exe delete "Sophos AutoUpdate Service"
        sc.exe delete "Sophos Agent"
        sc.exe delete "SAVService"
        sc.exe delete "SAVAdminService"
        sc.exe delete "Sophos Message Router"
        sc.exe delete "swi_service"
        sc.exe delete "swi_update"
        sc.exe delete "SntpService"
        sc.exe delete "Sophos System Protection Service"
        sc.exe delete "Sophos Endpoint Defense Service"
        sc.exe delete "Sophos Device Control Service"
        
        
        
        # This function performs a second sweep to remove leftover components.
        Function Remove-Sophos
            {
                Param ($Application)
                # Get Sophos services and stop them.
                $Service = Get-Service | Where-Object {$_.Name -like "*Sophos*" -and $_.Status -like "Running"} -verbose
                try
                    {
                        Stop-Service $Service
                    }
                catch
                    {
                        Write-Error "Unable to stop the service $($Service.Name)"
                    }
        
        
                        Write-Output "Attempting to uninstall $($Application.Name)"
                                        try 
                                            {
                                                Uninstall-Package $Application.Name
                                                $Counter = 1
                                                Write-Verbose "Confirming that $($Application.Name) is uninstalled..."
                                                $Installed = Get-Package | Where-Object {$_.Name -like $Application.Name}
                                                While ($Installed -and $Counter -lt 4) 
                                                    {
                                                        Write-Warning "$($Application.Name) was not uninstalled, trying again... ($Counter)"
                                                        Uninstall-Package $Application.Name
                                                        $Counter++
                                                    }
        
                                                If ($Installed)
                                                    {
                                                        Write-Error "ERROR: Unable to uninstall $($Application.Name) after $Counter times"
                                                    }
        
                                                Else
                                                    {
                                                        Write-Output "Successfully removed $($Application.Name)"
                                                        $Counter = 0
                                                    }
                                            }
                                                catch 
                                            {
                                                Write-Error "Error: Failed to remove $($Application.Name)"
                                            }
            } # End of the function
        
        #Disable Tamper Protection (msiexec will re-inject registry keys)
    
        REG ADD "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\services\SAVService" /t REG_DWORD /v Start /d 0x00000004 /f
        REG ADD "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\Sophos MCS Agent" /t REG_DWORD /v Start /d 0x00000004 /f
        REG ADD "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\Sophos Endpoint Defense\TamperProtection\Config" /t REG_DWORD /v SAVEnabled /d 0 /f
        REG ADD "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\Sophos Endpoint Defense\TamperProtection\Config" /t REG_DWORD /v SEDEnabled /d 0 /f
        REG ADD "HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Sophos\SAVService\TamperProtection" /t REG_DWORD /v Enabled /d 0 /f
        
        # Gather the installed Sophos components.
        Write-Output "Gathering installed applications..."
        $AppArray = Get-Package | Where-Object {$_.Name -like "*Sophos*"}
        
        # Go through each of the apps in the order specified by Sophos and uninstall them.
        ForEach ($App in $AppArray)
            {
                switch ($App.Name)
                    {
                        "Sophos Remote Management System" 
                            {
                                Remove-Sophos $App
                            }	
                        "Sophos Network Threat Protection" 
                            {
                                Remove-Sophos $App
                            }
                        "Sophos Client Firewall" 
                            {
                                Remove-Sophos $App
                            }
                        "Sophos AutoUpdate" 
                            {
                                Remove-Sophos $App
                            }
                        "Sophos Diagnostic Utility" 
                            {
                                Remove-Sophos $App
                            }
                        "Sophos Exploit Prevention" 
                            {
                                Remove-Sophos $App
                            }
                        "Sophos Clean" 
                            {
                                Remove-Sophos $App
                            }
                        "Sophos Patch Agent" 
                            {
                                Remove-Sophos $App
                            }
                        "Sophos Compliance Agent" 
                            {
                                Remove-Sophos $App
                            }					
                        "Sophos Endpoint Defense" 
                            {
                                Remove-Sophos $App
                            }
                        "Sophos System Protection" 
                            {
                                Remove-Sophos $App
                            }
                        "Sophos Management Communication System" 
                            {
                                Remove-Sophos $App
                            }					
                        "Sophos SafeGuard components" 
                            {
                                Remove-Sophos $App
                            }
                        "Sophos Health" 
                            {
                                Remove-Sophos $App
                            }
                        "Sophos Heartbeat" 
                            {
                                Remove-Sophos $App
                            }
                        "Sophos Anti-Virus" 
                            {
                                Remove-Sophos $App
                            }					
                    }
            }
        
        #Disable Tamper Protection (msiexec will re-inject registry keys)
    
        REG ADD "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\services\SAVService" /t REG_DWORD /v Start /d 0x00000004 /f
        REG ADD "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\Sophos MCS Agent" /t REG_DWORD /v Start /d 0x00000004 /f
        REG ADD "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\Sophos Endpoint Defense\TamperProtection\Config" /t REG_DWORD /v SAVEnabled /d 0 /f
        REG ADD "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\Sophos Endpoint Defense\TamperProtection\Config" /t REG_DWORD /v SEDEnabled /d 0 /f
        REG ADD "HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Sophos\SAVService\TamperProtection" /t REG_DWORD /v Enabled /d 0 /f
        
        #Uninstall Sophos Remote Management System (Third sweep)
        $SRMSVer = Get-ChildItem -Path HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall, HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall  |
            Get-ItemProperty |
                Where-Object {$_.DisplayName -match "Sophos Remote Management System" } |
                    Select-Object -Property DisplayName, UninstallString
        
        ForEach ($ver in $SRMSVer) {
        
            If ($ver.UninstallString) {
        
                $uninst = $ver.UninstallString
                Start-Process cmd "/c $uninst /qn REBOOT=SUPPRESS /PASSIVE" -NoNewWindow -verbose
                }
        
        }
            #Uninstall Sophos AutoUpdate (Third sweep)
        $SRMSVer = Get-ChildItem -Path HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall, HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall  |
            Get-ItemProperty |
                Where-Object {$_.DisplayName -match "Sophos AutoUpdate" } |
                    Select-Object -Property DisplayName, UninstallString
        
        ForEach ($ver in $SRMSVer) {
        
            If ($ver.UninstallString) {
        
                $uninst = $ver.UninstallString
                Start-Process cmd "/c $uninst /qn REBOOT=SUPPRESS /PASSIVE" -NoNewWindow -verbose
                }
        
        }
    
        
        
        #Cleanup Registry Keys
    
        cmd /c "REG DELETE HKLM\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{23E4E25E-E963-4C62-A18A-49C73AA3F963} /f"
        cmd /c "REG DELETE HKLM\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{84748F71-7BF1-4F73-9340-D0785F4B0197} /f"
        cmd /c "REG DELETE HKLM\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{65323B2D-83D4-470D-A209-D769DB30BBDB} /f"
        cmd /c "REG DELETE HKLM\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{FED1005D-CBC8-45D5-A288-FFC7BB304121} /f"
        cmd /c "REG DELETE HKLM\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{7CD26A0C-9B59-4E84-B5EE-B386B2F7AA16} /f"
        cmd /c "REG DELETE HKLM\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{D929B3B5-56C6-46CC-B3A3-A1A784CBB8E4} /f"
        cmd /c "REG DELETE HKLM\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{FED1005D-CBC8-45D5-A288-FFC7BB304121} /f"
        cmd /c "REG DELETE HKLM\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{AFBCA1B9-496C-4AE6-98AE-3EA1CFF65C54} /f"
        cmd /c "REG DELETE HKLM\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{1AC3C833-D493-460C-816F-D26F30F79DC3} /f"
        cmd /c " for /f `"tokens=*`" %a in ('reg query `"HKLM\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall`" /f `"Sophos Anti-Virus`" /d /s /e ^| find /I `"HKEY`" ') do @(reg delete `"%~a`" /f)" 
        cmd /c " for /f `"tokens=*`" %a in ('reg query `"HKLM\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall`" /f `"Sophos AutoUpdate`" /d /s /e ^| find /I `"HKEY`" ') do @(reg delete `"%~a`" /f)"
        cmd /c " for /f `"tokens=*`" %a in ('reg query `"HKLM\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall`" /f `"Sophos Remote Management System`" /d /s /e ^| find /I `"HKEY`" ') do @(reg delete `"%~a`" /f)"
        cmd /c " for /f `"tokens=*`" %a in ('reg query `"HKLM\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall`" /f `"Sophos System Protection`" /d /s /e ^| find /I `"HKEY`" ') do @(reg delete `"%~a`" /f)"
        cmd /c " for /f `"tokens=*`" %a in ('reg query `"HKLM\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall`" /f `"Sophos Endpoint Defense`" /d /s /e ^| find /I `"HKEY`" ') do @(reg delete `"%~a`" /f)"
        cmd /c " for /f `"tokens=*`" %a in ('reg query `"HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall`" /f `"Sophos Anti-Virus`" /d /s /e ^| find /I `"HKEY`" ') do @(reg delete `"%~a`" /f)" 
        cmd /c " for /f `"tokens=*`" %a in ('reg query `"HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall`" /f `"Sophos AutoUpdate`" /d /s /e ^| find /I `"HKEY`" ') do @(reg delete `"%~a`" /f)"
        cmd /c " for /f `"tokens=*`" %a in ('reg query `"HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall`" /f `"Sophos Remote Management System`" /d /s /e ^| find /I `"HKEY`" ') do @(reg delete `"%~a`" /f)"
        cmd /c " for /f `"tokens=*`" %a in ('reg query `"HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall`" /f `"Sophos System Protection`" /d /s /e ^| find /I `"HKEY`" ') do @(reg delete `"%~a`" /f)"
        cmd /c " for /f `"tokens=*`" %a in ('reg query `"HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall`" /f `"Sophos Endpoint Defense`" /d /s /e ^| find /I `"HKEY`" ') do @(reg delete `"%~a`" /f)"
        REG Delete "HKLM\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Run" /v "Sophos AutoUpdate Monitor" /f
        
        #Directory Cleanup
        
        Remove-Item -LiteralPath "C:\Program Files\Sophos*" -Force -Recurse
        Remove-Item -LiteralPath "C:\Program Files\Sophos" -Force -Recurse
        Remove-Item -LiteralPath "C:\Program Files (x86)\Sophos" -Force -Recurse
        Remove-Item -LiteralPath "C:\ProgramData\Sophos" -Force -Recurse
        
        #Check for Sophos components
        
        Get-Package | Where-Object {$_.Name -like "*Sophos*"}
        
        #Check Firewall Profile Status
        get-netfirewallprofile -policystore activestore
        }  _Removal of 