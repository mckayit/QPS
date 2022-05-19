Function allmailboxstats
{

    <#
  #... get all the mailbox stats on the domain you are on.>
 .SYNOPSIS
     get all the mailbox stats on the domain you are on.
 .DESCRIPTION
     get all the mailbox stats on the domain you are on.
     It uses Get-mailbox to get all the mailboxes.
 .PARAMETER one
     Specifies Pram details.
 
 .EXAMPLE
     C:\PS>Get-allmailboxstats


 .INPUTS
     Inputs to this cmdlet (if any)
 .OUTPUTS
     Output from this cmdlet (if any)
 .NOTES
   
     Lawrence McKay
     Lawrence@mckayit.com
     McKayIT Solutions Pty 
    
     Date:    7 Mar 2022
      
      ******* Update Version number below when a change is done.*******
     
     History
     Version         Date                Name           Detail
     ---------------------------------------------------------------------------------------
     0.0.1           7 Mar 2022         Lawrence       Initial Coding

 #>


    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false,
            HelpMessage = 'Enter The Mailbox')]
        $mailboxes = $(Get-Mailbox -resultsize Unlimited )
    )

    $i = 1

    #Loop through mailbox list and collect the mailbox statistics
    foreach ($mb in $mailboxes)
    {
	
        $paramWriteProgress = @{
            Activity        = 'Collecting mailbox details'
            Status          = "Processing [$i] of [$($Mailboxes.Count)] Mailboxes"
            PercentComplete = (($i / $Mailboxes.Count) * 100)
                                
        }
                            
        Write-Progress @paramWriteProgress
                        
        $i++


        $stats = $mb | Get-MailboxStatistics | Select-Object DisplayName, LastLogonTime, TotalItemSize, itemcount
        $user = $mb | Get-Recipient | Select-Object UserPrincipalName, DistinguishedName, RetentionPolicy, PrimarySmtpAddress, EmailAddresses, HiddenFromAddressListsEnabled, IsMailboxEnabled, RecipientType, RecipientTypedetails
        $mbupn = $mb.UserPrincipalName
        $ad = Get-ADUser -Filter "UserPrincipalName -eq '$($mbupn)'" -Properties * 
        # Get-ADUser -Filter * -Properties proxyaddresses | Select-Object Name, @{L  "ProxyAddresses"; E = { ($_.ProxyAddresses -match '^smtp:') -join ";"}}           

        #Create a custom PS object to aggregate the data we're interested in
        [PSCustomObject] @{	
            "DisplayName"                   = $stats.DisplayName
            "SamAccountName"                = $mb.SamAccountName
            "UserPrincipalName"             = $mb.UserPrincipalName
            "EmailAddress"                  = $user.PrimarySmtpAddress
            "RetentionPolicy"               = $user.RetentionPolicy
            "Size(MB)"                      = $stats.totalitemsize.Value.ToKB()
            "LastLogon"                     = $stats.LastLogonTime
            "MBEnabled"                     = $mb.IsMailboxEnabled
            "HiddenFromAddressListsEnabled" = $mb.HiddenFromAddressListsEnabled
            "RecipientType"                 = $mb.RecipientType
            "RecipientTypeDetails"          = $mb.RecipientTypeDetails
            "AD USER Enabled"               = $ad.Enabled
            "AD Last LogOn Date"            = $ad.lastlogondate
            "AD When Changed"               = $ad.whenchanged
            "All SMTPs"                     = $ad |  Select-Object  @{L = ""; E = { ($_.ProxyAddresses -match 'smtp:') -join ";" } } 
            "DistinguishedName"             = $user.DistinguishedName

        }
    }
}
