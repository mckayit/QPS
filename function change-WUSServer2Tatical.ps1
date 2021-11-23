function change-WUSServer2Tatical
{
    <#
 #... <Short Description>
.SYNOPSIS
    Short description
.DESCRIPTION
    Long description
.PARAMETER one
    Specifies Pram details.
.PARAMETER two
    Specifies Pram details
.PARAMETER InputObject
    Specifies the object to be processed.  You can also pipe the objects to this command.
.EXAMPLE
    C:\PS>
    Example of how to use this cmdlet
.INPUTS
    Inputs to this cmdlet (if any)
.OUTPUTS
    Output from this cmdlet (if any)
.NOTES
    
    Lawrence McKay
    Lawrence@mckayit.com
    McKayIT Solutions Pty 
     
    Date:    15 Feb 2021
      
     ******* Update Version number below when a change is done.*******
     
    History
    Version         Date                Name           Detail
    ---------------------------------------------------------------------------------------
    0.0.1           9 Feb 2021         Lawrence       Initial Coding

#>

    
    
    begin 
    {
             
    }
    
    process 
    {
            
        try 
        {
            New-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate" -Name "WUServer" -Value "http://qps-scm-pr-01.prds.qldpol:8530" -Force -PropertyType String

            New-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate" -Name "WUStatusServer" -Value "http://qps-scm-pr-01.prds.qldpol:8530" -Force -PropertyType String
        }
        catch 
        {
            Write-Host 'ERROR : $(_.Exception.Message)' -ForegroundColor Magenta
        }
    }
    
    end
    {
        #  this bit restarts the services so it re-reads in the new REG settings.
        # 
        #  note   Reboot will put back the GPO WSUS server.      
        net stop wuauserv
        net stop cryptSvc
        net stop bits
        net stop msiserver

        del /f /q “%ALLUSERSPROFILE%\Application Data\Microsoft\Network\Downloader\qmgr*.dat”
        #del /f /s /q %SystemRoot%\SoftwareDistribution\*.*
        #del /f /s /q %SystemRoot%\system32\catroot2\*.*
        #del /f /q %SystemRoot%\WindowsUpdate.log

        net start wuauserv
        net start cryptSvc
        net start bits
        net start msiserver
           
        #wuinstall -AcceptAll -Install 
    }
}
change-WUSServer2Tatical
