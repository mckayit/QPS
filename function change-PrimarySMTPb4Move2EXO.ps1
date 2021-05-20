function change-PrimarySMTPb4Move2EXO
{
    <#
 #... Sets the PrimarySMTPAddress to Match UPN
.SYNOPSIS
    Sets the PrimarySMTPAddress to Match UPN
.DESCRIPTION
    This Function will get the import from commandline and then Sets the PrimarySMTPAddress to Match UPN
    Thius is required for moving to EXO
.PARAMETER one
    Samaccountname
.
.PARAMETER InputObject
    Specifies the object to be processed.  You can also pipe the objects to this command.
.EXAMPLE
    C:\PS>change-PrimarySMTPb4Move2EXO -samAccount $bulkSam
    C:\PS>change-PrimarySMTPb4Move2EXO -samAccount 904223
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
    0.0.1           3 Mar 2021         Lawrence       Initial Coding

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
                    Status          = "Processing  [$i] of [$($Upn.Count)] Mailbox $($sam)"
                    PercentComplete = (($i / $#samaccount.Count) * 100)

                }

                Write-Progress @paramWriteProgress
                $i++

                $note = ""

                $before = Get-mailbox -filter { Samaccountname -eq $sam }

                Try
                {
                    $newPRISMTP = "$sam.userprincipalname"
                    set-mailbox $sam -primarySMTPAddress $newPRISMTP
                }
                catch
                {
                    $Errormsg = $error[0].exception.message
                    Write-Host $Errormsg -ForegroundColor magenta
                    $note = $errormsg
                }
                
            

                [PSCustomObject] @{ 
                    Displayname               = $before.displayname
                    Lastname                  = $before.Surname
                    Samaccountname            = $before.samaccountname
                    PrimarySMTPAddress        = $before.PrimarySMTPAddress
                    ChangedPrimarySMTPAddress = $newPRISMTP
                    Notes                     = $note

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
