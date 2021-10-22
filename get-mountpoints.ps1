function get-disksizeWithMountpoints
{
    <#
 #... get All Moiuntpoints on a system
.SYNOPSIS
    get all mount points of disks on a system
.DESCRIPTION
    displays all disks sizings that are Fixed disk type. (3)  
    this will also show the disks that have mountpoints and where thewy are mounted.. 
    
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
     
    Date:    9 aug 2021
      
     ******* Update Version number below when a change is done.*******
     
    History
    Version         Date                Name           Detail
    ---------------------------------------------------------------------------------------
    0.0.1           9 Aug 2021         Lawrence       Initial Coding

#>
     
  

    begin 
    {
        #Setting the info on how the sizing are done.
                  
    }
    
    process 
    {

        try
        {
            $volumes = Get-WmiObject win32_volume -Filter "DriveType='3'" | Sort-Object DriveLetter # -Descending


            foreach ($volume in $volumes)
            {
                $TotalGB = [math]::round(($volume.Capacity / 1gb), 2) 
            
                $FreeGB = [math]::round(($volume.FreeSpace / 1Gb), 2) 
        
                $FreePerc = [math]::round(((($volume.FreeSpace / 1GB) / ($volume.Capacity / 1GB)) * 100), 0) 
    
                [PSCustomObject] @{
                    Name            = $volume.name
                    Label           = $volume.label
                    DriveLetter     = $volume.driveletter
                    FileSystem      = $volume.filesystem
                    "Capacity(GB)"  = $TotalGB
                    "FreeSpace(GB)" = $FreeGB
                    "Free(%)"       = $FreePerc
                }
            }
        }
    
  

        catch 
        {
            Write-Host 'ERROR : $(_.Exception.Message)' -ForegroundColor Magenta
        }

    }

    end 
    {
            
    }

}
