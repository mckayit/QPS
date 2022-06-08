

$2010 = Get-mailbox -resultsize unlimited | where { $_.database -match "MB-DB" }

foreach ($mb in $2010)
{

    $perm = Get-MailboxPermission $mb
    $trustee = Get-RecipientPermission $mb | where { ($_.trustee -ne "NT AUTHORITY\SELF") } | select Identity, Trustee, AccessControlType, AccessRights, IsInherited 

    $results = [ordered]@{
        Mailboxname               = $mb.displayname
        MailboxEmailaddress       = $mb.PrimarySmtpAddress
        ExchangeServerName        = $mb.servername
        Office                    = $mb.office
        RecipientType             = $mb.RecipientType
        RecipientTypedetails      = $mb.RecipientTypedetails
        Permissions_Identity      = $Perm.identity
        Permissions_User          = $perm.user
        Permissions_Deny          = $perm.deny
        Permissions_AccessRights  = $perm.AccessRights
        Permissions_IsInherited   = $perm.IsInherited
        Trustee_Identity          = $Trustee.Identity
        Trustee_Trustee           = $Trustee.trustee
        Trustee_AccessControlType = $Trustee.AccessControlType
        Trustee_AccessRights      = $Trustee.AccessRights
        Trustee_IsInherited       = $Trustee.IsInherited
                
                
    }
    New-Object PSObject -Property $results

}
