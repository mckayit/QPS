function remove-UsersfromOffice365group
{

    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true,
            HelpMessage = 'Enter UPN')]
        $UPNs
    )

    $i = 0

    # $usersfull = Get-ADUser -Filter { UserPrincipalName -eq '4013011@police.qld.gov.au' } -Properties memberof


    foreach ($upn in $UPNS)
    {
        #counter Progress bar.
        $paramWriteProgress = @{
            Activity         = "processing UPN"
            Status           = "Processing [$i] of [$($upns.Count)] users"
            PercentComplete  = (($i / $upns.Count) * 100)
            CurrentOperation = "Processing the following account : [ $($upn))]"
        }
        Write-Progress @paramWriteProgress       
        $i++


        #get accound Groups for user.
        $user = Get-ADUser -Filter { UserPrincipalName -eq $upn } -Properties memberof 
        $sam = $user.samaccountname

        #removeds user from groups that start with 'QPS-P-Office365'
        foreach ($group in $user.memberof )
        {
             
            if ($group -match 'QPS-P-Office365')
            {
                write-host $group
                Remove-ADGroupMember -Identity $group -Members $sam -Confirm:$false
            }
        }
    }

}
 