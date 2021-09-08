function set-ServerShares
{
    <#
 #... Sets the Server shares from a CSV filer.
.SYNOPSIS
    Sets the Server shares from a CSV filer.
.DESCRIPTION
    reades in a CSV file and then created the Directory structure then shars the DIR
     as per name and applies Default Permissions  Administrators:Full
.PARAMETER PATH2CSV
    path to CSV file
.EXAMPLE
    C:\PS>set-ServerShares -PATH2CSV c:\temp\shares.csv
    Example of how to use this cmdlet
.INPUTS
    Inputs to this cmdlet (if any)
.OUTPUTS
    Output from this cmdlet (if any)
.NOTES

the CSV file should have the following 
EG
Name	        path	              comment
APPLICATIONS	D:\APPLICATIONS	      Share Description (Optopnal)



    Lawrence McKay
    Lawrence@mckayit.com
    McKayIT Solutions Pty 
     
    Date:    8 Sep 2021
      
     ******* Update Version number below when a change is done.*******
     
    History
    Version         Date                Name           Detail
    ---------------------------------------------------------------------------------------
    0.0.1           8 Sep 2021         Lawrence       Initial Coding

#>
     
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true,
            HelpMessage = 'Enter The Output Folder')]
        [string]$path2CSVFile
    )
    
    
    begin 
    {
             
    }
    
    process 
    {
            
        try 
        {
            $share = import-csv $path2CSVFile
                
        }
        catch 
        {
            Write-Host 'ERROR : $file not Found or is invalid' -ForegroundColor Magenta
        }

        
        foreach ($sh in $share)
        {
            if (get-item -path $sh.path -ErrorAction SilentlyContinue)
            { 
                Write-Host 'Path already Exist: ' $sh.path -ForegroundColor Magenta 
            }

            if (!(get-item -path $sh.path -ErrorAction SilentlyContinue ))
            {
                new-item -path $sh.path -ItemType Directory -ErrorAction silentlycontinue 
            }
                     
            if (get-smbshare $sh.Name -ErrorAction SilentlyContinue)
            { 
                Write-host 'Error creating Share:' $sh.name -ForegroundColor Magenta
            }
            if (!(get-smbshare $sh.Name -ErrorAction SilentlyContinue))
            { 
                new-SmbShare -name $sh.name -path $sh.path -ChangeAccess "Users"  -Description $sh.comment -FullAccess "Administrators"
            }
           
        }

    }
    
    end
    {
        Get-SmbShare | select name, path, Description | ft -AutoSize -Wrap 

    }
}
