function New-MoveRequestTo2016
{
    <#
	.SYNOPSIS
		A brief description of the New-MoveRequestTo2016 function.
	
	.DESCRIPTION
		A detailed description of the New-MoveRequestTo2016 function.
	
	.PARAMETER PrimarySMtpAddress
		A description of the PrimarySMtpAddress parameter.
	
	
	.PARAMETER BatchName
		A description of the BatchName parameter.
	
	.EXAMPLE
		PS C:\> New-MoveRequestTo2016 -PrimarySMtpAddress 'value1' -BATCHNAME 'value'
        PS C:\> New-MoveRequestTo2016 -PrimarySMtpAddress 'value1' -BATCHNAME 'value' -targetdatabase "DAG5DB1"
	
	.NOTES
		Additional information about the function.
#>
    [CmdletBinding()]
    param
    (
        $PrimarySmtpAddress,
        $batchName,
        $Targetdatabase = @('dag5db1', 'dag5db2')
	
    )
	
    begin
    {
        $i = 1
		
    }
    process
    {
        foreach ($SMTPAddress in $PrimarySmtpAddress)
        {

            if ($SMTPAddress.count -gt 1)
            {
                $paramWriteProgress = @{
                    Activity         = 'Adding MoveRequests'
                    Status           = "Processing [$i] of [$($SMTPAddress.Count)] users"
                    PercentComplete  = (($i / $SMTPAddress.Count) * 100)
                    CurrentOperation = "Completed : [$SMTPAddress]"
                }
                Write-Progress @paramWriteProgress
            }

            $database = (get-random $targetdatabase)
            try
            {
            
                new-moverequest -Identity $SMTPAddress -targetdatabase $database -ArchiveTargetDatabase $database -BatchName $batchname -whatif
                Write-host $database

                $prop = [ordered]@{
                    User      = $SMTPAddress
                    UserNo    = $i
                    BatchName = $BatchName
                    Status    = 'Added'
                    Details   = 'None'
                }

            }
            catch
            {
                $prop = [ordered]@{
                    User      = $SMTPAddress
                    UserNo    = $i
                    BatchName = $BatchName
                    Status    = 'FAILED'
                    Details   = "ERROR : $($_.Exception.Message)"
                }
            }
            finally
            {
                $obj = New-Object -TypeName System.Management.Automation.PSObject -Property $prop
                Write-Output $obj
				
                $i++
            }
        }
    }
    end
    {
        Write-Progress -Activity 'Testing WhatIf' -Completed
    }
}