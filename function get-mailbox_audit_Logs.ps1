#$mailbox = get-mailbox -resultsize unlimited
function get-mailbox_audit_Logs
{
    <#
 #... Gets All Mailbox Auditing.
.SYNOPSIS
    Gets All Mailbox Auditing.
.DESCRIPTION
    This Function will get the Gets All Mailbox Auditing from all onpremise Mailboxes if not added via CLI
    the imput will be the output from get-mailbox <mailboxname> 

.PARAMETER one
    Mailbox   This is optional.   
    imput shoud be retrieved from Get-mailbox <mailboxname>

.EXAMPLE
    C:\PS>get-mailbox_audit_Logs
    C:\PS>get-mailbox_audit_Logs -mailbox $mailboxOutput
    Example of how to use this cmdlet

.INPUTS
    Inputs to this cmdlet (if any)

.OUTPUTS
    Output from this cmdlet (if any)

.NOTES
    
    Lawrence McKay
    Lawrence@mckayit.com
    McKayIT Solutions Pty 
     
    Date:    3 March 2020
      
     ******* Update Version number below when a change is done.*******
     
    History
    Version         Date                Name           Detail
    ---------------------------------------------------------------------------------------
    0.0.1           8 Mar 2020         Lawrence       Initial Coding
    0.0.2          19 May 2022         Lawrence       Added Progress bar and an extra few fields.

#>





    [CmdletBinding()]
    param(
        $mailbox
     
    )

    begin
    {
        #checking Mailbox has not been added as a VAR   If not then get all mailbox details.

        if ($mailbox -eq $null)
        {
            write-host "Getting all mailboxes Onpremise to be Checked" -ForegroundColor cyan
            $mailbox = get-mailbox -resultsize unlimited | where { $_.servername -match "qps-xch-" }
        }
    }

    process 
    {
            
        try 
        {

            $i = 1
            foreach ($mb in $mailbox)
            {
       
                #progress bar Start
                $paramWriteProgress = @{
                    Activity        = "Processing  Mailbox "
                    Status          = "Processing  [$i] of [$($mailbox.Count)] Mailbox $($mb.name)"
                    PercentComplete = (($i / $mailbox.Count) * 100)

                }

                Write-Progress @paramWriteProgress
                $i++
                #progress bar End

                $search1 = Search-MailboxAuditLog -Identity $mb.name  -ResultSize 10 -ShowDetails
                if ($search1 -ne $null )
                {
                    foreach ($sea in $Search1)
                    {
                        
                        $sea | Add-Member -type noteproperty -name "MBname" -value $mb.name
                        $sea | Add-Member -type noteproperty -name "HiddenfromAddressbook" -value $mb.HiddenFromAddressListsEnabled
                        $sea  | select  MBname, 
                        HiddenfromAddressbook,
                        MailboxResolvedOwnerName,
                        Operation,
                        OperationResult,
                        LogonType,
                        ExternalAccess,
                        FolderPathName,
                        ClientInfoString,
                        ClientProcessName,
                        ClientVersion,
                        InternalLogonType,
                        MailboxOwnerUPN,
                        MailboxOwnerSid,
                        LogonUserDisplayName,
                        IsValid
                        
                    }
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

get-mailbox_audit_Logs | export-csv C:\temp\Audit_MB_accessreport.csv -NoTypeInformation
Send-MailMessage -To sharma.bharat@police.qld.gov.au -From code@prds.qldpol -Attachments C:\temp\Audit_MB_accessreport.csv -Body "QLD 2010 Mailbox  Audit report" -Subject "2010 Audit report" -SmtpServer localhost
