<# 
    ********************************************************************************  
    *                                                                              *  
    *        This script Loads all my common f u n c t i o n s                     *
    *                                                                              *  
    ********************************************************************************    
    Note.
    All the standard Functions I use loaded.
      
    *******************
    Copyright Notice.
    *******************
    This Program is free software: you can redistribute it and/or modify
    it under the terms of the GNU General Public License as published by
    the Free Software Foundation, either version 3 of the License, or
    (at your option) any later version.

    This program is distributed in the hope that it will be useful,
    but WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
    GNU General Public License for more details.

    You should have received a copy of the GNU General Public License
    along with this program.  If not, see <http://www.gnu.org/licenses/>.

    test-exchangeonlineconnected

    Author:
    Lawrence McKay
    Lawrence@mckayit.com
    McKayIT Solutions Pty Ltd
    
    Date:   26  March 2018
   

    ******* Update Version number below when a change is done.*******

    History
    Version         Date                Name           Detail
    ---------------------------------------------------------------------------------------
    0.0.1           15 Sept 20121       Lawrence       Initial Coding



    #>

$version = "Version 0.0.31"
Write-Host $version -ForegroundColor Green


#$global:Functisloaded = "YES"

#Write-host "`nLoading $PSCommandPath"  -ForegroundColor green -BackgroundColor Red
$FormatEnumerationLimit = -1

Function get-help1 
{
    <#
    #... Displays the Functions in QHSupport.
    Syntax   Get-help1

    This will generate the report for all of the batchname starting with Batch14*
#>

    param
    (
	
    )
   	
    BEGIN
    {
        Write-Host $version -ForegroundColor Green
    }

    PROCESS
    {
        #... Display Functions with the comments
        $DDSPLAY = ""
        #Reads in the Current Powershell script file
        $DDSPLAY = get-content $PSCommandPath 

        foreach ($line in $DDSPLAY)
        {
            if ($line.Trim().StartsWith('Function', "CurrentCultureIgnoreCase") -or $line.Trim().Startswith('#...', "CurrentCultureIgnoreCase"))
            {
                $1 = $line

                if ($1.Trim().StartsWith('#...', "CurrentCultureIgnoreCase"))
                {
                
                    #Removes the first 4 char's Eg "#... "
                    $linedes = $line.trim().substring(4)
                    Write-Host $linedes -f Gray -NoNewline
                }
                Elseif (!($1.Trim().Startswith('#...')))
                {
                    Write-host ''
                }
                if ($1.Trim().StartsWith('Function', "CurrentCultureIgnoreCase"))
                {
                    $linelong = $line + "                                               "
                
                    #makes the Line length to be 50 so the comments all line up. Fills it up with a space.
                    $line = $linelong.substring(0, 50)
                    Write-host "  $line" -f green -NoNewline
                }
            }
        }
        
    }
    END
    {
        Write-Output  ""
    }
}
Function sync-qhmodule 
{
    #... copies QHO365MigrationOps.psm1 module to 
    [string]$sourcefiles = 'C:\Users\904223\OneDrive - Queensland Police Service\Github\QPS\QPS_support.psm1'
    [string]$destinationDir = 'c:\Windows\System32\WindowsPowerShell\v1.0\Modules\QPSSupport\'
    copy-item -force -Recurse $sourcefiles -Destination $destinationDir


    [string]$sourcefiles = '\\exc-mgtbk7p001\c$\Windows\System32\WindowsPowerShell\v1.0\Modules\QHSupport\*'
    [string]$destinationDir = 'c:\Windows\System32\WindowsPowerShell\v1.0\Modules\QHSupport\'
    copy-item -Recurse $sourcefiles -Destination $destinationDir -force 
    Remove-Module QHSupport -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 10
    import-module QHSupport -Verbose
}