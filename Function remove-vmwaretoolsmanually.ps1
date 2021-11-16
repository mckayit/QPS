Function remove-vmwaretoolsmanually
{
    <#
 #... removes VMware tools and informational error reporting
.SYNOPSIS
    removes VMware tools and informational error reporting
.DESCRIPTION
    Manually removes VMware tools and informational error reporting
    in the application Event log on 2019 servers you can get lots of informational events relating to windows reporting (Errors)
    The reason this happens is due to an application has crashed and logged data to C:\ProgramData\Microsoft\Windows\WER\ReportQueue
    the eventlog seems to poole this loaction regularry and contuniously report the errors and fill the eventlog.
    Removing all the sub folders will stop these errors from contuniously reporting.
.PARAMETER one
    None

    .EXAMPLE
    C:\PS>
    
.OUTPUTS
    Output from this cmdlet (if any)
.NOTE
    
    Lawrence McKay
    Lawrence@mckayit.com
    McKayIT Solutions Pty 
     
    Date:    11 Nov 2021
      
     ******* Update Version number below when a change is done.*******
     
    History
    Version         Date                Name           Detail
    ---------------------------------------------------------------------------------------
    0.0.1           10 Nov 2021         Lawrence       Initial Coding

#>

    function remove-QPSFolders
    {
        <#
     #... Removes DIR and REG keys recursively 
    .SYNOPSIS
    Removes DIR and REG keys recursively 
    
    .DESCRIPTION
    this function will remove all dir's or Regkey's recursivly.  
    There is ** NO ** Confirmation and will just do it.
    So be very carefull with what you add as a Var "Foldername"

  
    .EXAMPLE
    remove-QPSFolders -foldername  '<Folder path>'
    remove-QPSFolders -foldername  'C:\Program Files\VMware'

    remove-QPSFolders -foldername  '<Reg key>'
    remove-QPSFolders -foldername  "HKLM:\SOFTWARE\VMware, Inc."
    
    .NOTES
    Note 
        this script needs to be run twice (2x) to cleanup correctly.

        

    #>


        [CmdletBinding()]
        param( 
            [Parameter(Mandatory = $true)]
            $FolderName 
        
        )
        Begin
        {
     

        }

        Process
        {    
            if (Test-Path $FolderName)
            {
                Try
                {
                
                    Remove-Item $FolderName -Force -Recurse
                    Write-Host "$foldername has been removed." -ForegroundColor Green
                }
                catch
                {
                    Write-Host "$FolderName Failed to be removed" -ForegroundColor  red 
                }
            }
            else
            {
                Write-Host "$FolderName Doesn't Exists"  -ForegroundColor  Cyan 

            }
        }
        end
        {

        }

    }

    <# 
    Removes all the Windows Error Reporting folder contents.
    This removes the error from appearing in the Eventlog.
    this is required as the evenlog reports this as informational.
#>
    function remove-informationaleventlogforerrores
    {
        <#
 #... removes informationalError from eventlog 
.SYNOPSIS
    removes informationalError from eventlog
.DESCRIPTION
    in the application Event log on 2019 servers you can get lots of informational events relating to windows reporting (Errors)
    The reason this happens is due to an application has crashed and logged data to C:\ProgramData\Microsoft\Windows\WER\ReportQueue
    the eventlog seems to poole this loaction regularry and contuniously report the errors and fill the eventlog.
    Removing all the sub folders will stop these errors from contuniously reporting.
.PARAMETER one
    None

    .EXAMPLE
    C:\PS>remove-informationaleventlogforerrores
    
    
.OUTPUTS
    Output from this cmdlet (if any)
.NOTES
    
    Lawrence McKay
    Lawrence@mckayit.com
    McKayIT Solutions Pty 
     
    Date:    11 Nov 2021
      
     ******* Update Version number below when a change is done.*******
     
    History
    Version         Date                Name           Detail
    ---------------------------------------------------------------------------------------
    0.0.1           10 Nov 2021         Lawrence       Initial Coding

#>
     
        begin 
        {
            $WERDirs = (Get-ChildItem -Directory C:\ProgramData\Microsoft\Windows\WER\ReportQueue).FullName      
        }
    
        process 
        {
            foreach ($WER in $WERDIRs)
            { 
                try 
                {
                    remove-QPSFolders -foldername $wer 
                }
    
                catch 
                {

                }
            }
        }

    }
    
    end
    {
            
    }
}

# Removes the Informational reporting from eventlog
remove-informationaleventlogforerrores

# removes the VMware tools DIR's
remove-QPSFolders -foldername  'C:\Program Files\VMware'
remove-QPSFolders -foldername  'C:\ProgramData\VMware' 

# removes VMware Tools reg keys. 
remove-QPSFolders -foldername  "HKLM:\SOFTWARE\VMware, Inc."
remove-QPSFolders -foldername  "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{1D060220-2A64-4153-A6F5-C43B95C3BFC7}"
remove-QPSFolders -foldername  "HKCU:\Software\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Compatibility Assistant\Store" 
remove-QPSFolders -foldername  "HKLM:\SOFTWARE\Classes\Installer\Products\022060D146A235146A5F4CB3593CFB7C"
remove-QPSFolders -foldername  "HKLM:\SOFTWARE\Classes\TypeLib\{6B8C0665-86D9-4DC9-8D58-FABE31A495E3}"


#reboots the system within 2 Sec
shutdown /r /t 2
}