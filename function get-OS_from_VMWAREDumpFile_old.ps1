function get-OS_from_VMWAREDumpFile

{
    <#
 #... Gets the OS from the $vmwareimport
.SYNOPSIS
    Gets the OS from the $vmwareimport
.DESCRIPTION
    Gets the OS from Ad using the hostname from the $vmwareimport

    
.PARAMETER $vmwareimport
    $vmwareimport import file


.PARAMETER prdsCreds
    gets creds for PRDS domain
.PARAMETER DESQLDCreds
   gets creds for PRDS domain
.PARAMETER ACDSCreds
   gets creds for PRDS domain
.PARAMETER DVDSCreds
   gets creds for DESQLD domain


.PARAMETER InputObject
    Specifies the object to be processed.  You can also pipe the objects to this command.
.EXAMPLE
    C:\PS>Get-serverOS -vmwareimport  -prdscreds -desqldcreds
    C:\PS>Get-serverOS -vmwareimport  -prdscreds -desqldcreds |export-csv d:\serversoftware.csv 
    
.INPUTS
    Inputs to this cmdlet (if any)
.OUTPUTS
    Output from this cmdlet (if any)
.NOTES
    
    Lawrence McKay
    Lawrence@mckayit.com
    McKayIT Solutions Pty 
     
    Date:    19 May 2021
      
     ******* Update Version number below when a change is done.*******
     
    History
    Version         Date                Name           Detail
    ---------------------------------------------------------------------------------------
    0.0.1           15 June 2021         Lawrence       Initial Coding

#>

    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $vmwareimportfile,
        [Parameter(Mandatory = $false)]
        $PRDSCREDS = (get-credential -Message "Enter your PRDS Admin Creds"  ),

        [Parameter(Mandatory = $false)]
        $desqldcreds = (get-credential -Message "Enter your DESQLD Admin Creds" ),

        [Parameter(Mandatory = $false)]
        $acdscreds = (get-credential -Message "Enter your ACDS Admin Creds" ),

        [Parameter(Mandatory = $false)]
        $dvdscreds = (get-credential -Message "Enter your DVDS Admin Creds" )
        
        
    )
    
    begin
    {
        $vmwareimporta = import-csv  $vmwareimportfile

        $i = 1
        foreach ($server in $vmwareimporta)
        {
            #progress bar.
            $paramWriteProgress = @{
                Activity         = "getting Server info"
                Status           = "Processing [$i] of [$($vmwareimporta.Count)] users"
                PercentComplete  = (($i / $vmwareimporta.Count) * 100)
                CurrentOperation = "processing the following server : [ $(($server).'DNS NAME')]"
            }
            Write-Progress @paramWriteProgress       
            $i++

        
            $os = "Not Detected"
            #ACDS     
            if (($server).'DNS NAME' -match "acds.accpol")
            {
         
                $name = (($server).'DNS NAME').trimend(".acds.accpol")

                try
                {
                    $osinfo = get-adcomputer -filter 'ipv4address -eq $server.'Primary IP Address' -Server acds.accpol -Credential $acdscreds -Properties OperatingSystem , OperatingSystemServicePack, OperatingSystemVersion, CanonicalName | Select-Object OperatingSystemServicePack, OperatingSystem , OperatingSystemVersion , CanonicalName
                    #$osinfo = get-adcomputer $name  -Server acds.accpol -Credential $acdscreds -Properties OperatingSystem , OperatingSystemServicePack, OperatingSystemVersion, CanonicalName | Select-Object OperatingSystemServicePack, OperatingSystem , OperatingSystemVersion , CanonicalName
                    $os = $OSINFo.OperatingSystem
                }
                Catch
                { $os = "Not Detected" }
            }
            #DVDS
            if (($server).'DNS NAME' -match "dvds.devpol")
            {
         
                $name = (($server).'DNS NAME').trimend(".dvds.devpol")

                try
                {
                    $osinfo = get-adcomputer $name  -Server dvds.devpol -Credential $dvdscreds -Properties OperatingSystem , OperatingSystemServicePack, OperatingSystemVersion, CanonicalName | Select-Object OperatingSystemServicePack, OperatingSystem , OperatingSystemVersion , CanonicalName
                    $os = $OSINFo.OperatingSystem
                }
                Catch
                { $os = "Not Detected" }
            }


            #DESQLD
            if (($server).'DNS NAME' -match "Desqld.internal")
            {
         
                $name = (($server).'DNS NAME').trimend(".Desqld.internal")

                try
                {
                    $osinfo = get-adcomputer  -filter 'ipv4address -eq $server.'Primary IP Address' -Server desqld.internal -Credential $DESQLDcreds -Properties OperatingSystem , OperatingSystemServicePack, OperatingSystemVersion, CanonicalName | Select-Object OperatingSystemServicePack, OperatingSystem , OperatingSystemVersion , CanonicalName
                    $os = $OSINFo.OperatingSystem
                }
                Catch
                { $os = "Not Detected" }
            }
       
            #PRDS
            if (($server).'DNS NAME' -match ".prds.qldpol")
            {
         
                $name = (($server).'DNS NAME').trimend(".prds.qldpol")

                try
                {
                    $OSINFo = get-adcomputer $name -server prds.qldpol  -credential $PRDSCREDS -Properties OperatingSystem , OperatingSystemServicePack, OperatingSystemVersion, CanonicalName | Select-Object OperatingSystemServicePack, OperatingSystem , OperatingSystemVersion , CanonicalName
                    $os = $OSINFo.OperatingSystem
                }
                catch
                {
                    $os = "Not Detected"
                }
            }
    
            if (($server).'DNS NAME' -match ".psba.qld.gov.au")
            {
                $os = "Not Detected"
            }

            [PSCustomObject] @{"VM UUID"     = $server."VM UUID"
                "DNS Name"                   = $server."DNS Name"
                "Template"                   = $server.Template
                "OperatingSystem"            = $OS
                "OperatingSystemVersion"     = $osinfo.OperatingSystemVersion
                "OperatingSystemServicePack" = $OSinfo.OperatingSystemServicePack
                "OperatingSystemHotfix"      = $OSinfo.OperatingSystemHotfix
                "CanonicalName"              = $OSinfo.CanonicalName
            }



        }
        end {}
    }
}
