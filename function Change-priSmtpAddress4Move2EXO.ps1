function Change-priSmtpAddress4Move2EXO
{
    <#
 #... Sets the Routing Address to account before Move.
.SYNOPSIS
    Sets the Routing Address to account before Move
.DESCRIPTION
    This Function will get the import from commandline and then Add the Router SMTPAddress to "$sam@psbaqld.mail.onmicrosoft.com"
    Thius is required for moving to EXO
.PARAMETER one
    Samaccountname
.
.PARAMETER InputObject
    Specifies the object to be processed.  You can also pipe the objects to this command.
.EXAMPLE
    C:\PS>add-RoutingAddress4Move2EXO -samAccount $bulkSam
    C:\PS>add-RoutingAddress4Move2EXO -samAccount 904223
    Example of how to use this cmdlet
.INPUTS
    Inputs to this cmdlet (if any)
.OUTPUTS
    Output from this cmdlet (if any)
.NOTES
    
    Lawrence McKay
    Lawrence@mckayit.com
    McKayIT Solutions Pty 
     
    Date:    3 March 2021
      
     ******* Update Version number below when a change is done.*******
     
    History
    Version         Date                Name           Detail
    ---------------------------------------------------------------------------------------
    0.0.1           8 Mar 2021         Lawrence       Initial Coding

#>
     
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $SamAccount
     
    )
    
    
    begin 
    {
        $i = 1  
    }
    
    process 
    {
            
        try 
        {
            foreach ($sam in $samaccount)
            {

                #progress bar
                $paramWriteProgress = @{
                    Activity        = "Processing  User  "
                    Status          = "Processing  [$i] of [$($samaccount.Count)] Mailbox $($sam)"
                    PercentComplete = (($i / $samaccount.Count) * 100)

                }

                Write-Progress @paramWriteProgress
                $i++

                $note = ""

                $before = Get-mailbox  $sam 

                Try
                {
                    $newPRISMTPAddress = "$sam@police.qld.gov.au"
                    Set-Mailbox $sam -EmailAddressPolicyEnabled $false -WarningAction silentlycontinue
                    Set-Mailbox $sam -PrimarySmtpAddress $newPRISMTPAddress -WarningAction silentlycontinue
                    
                    
                }
                catch
                {
                    $Errormsg = $error[0].exception.message
                    Write-Host $Errormsg -ForegroundColor magenta
                    $note = $errormsg
                }
                
            

                [PSCustomObject] @{ 
                    Displayname        = $before.displayname
                    # Lastname                  = $before.Surname
                    Samaccountname     = $before.samaccountname
                    PrimarySMTPAddress = $before.PrimarySMTPAddress
                    NewPRI_SMTP        = $newPRISMTPAddress
                    Notes              = $note

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
