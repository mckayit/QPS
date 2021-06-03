param
(
    [ValidateSet('qps-xch-dr-01', 'qps-xch-dr-02', 'qps-xch-dr-03', 'qps-xch-pr-01', 'qps-xch-pr-02', 'qps-xch-pr-03', 'kedmbx02.desqld.internal')]
    $Server = 'qkedmbx02.desqld.internal'
                            
    #$credential = $(Import-Clixml $env:USERPROFILE\Cred\OnPremCred.Clixml)
    $credential = Get-Credential
)
Try
{
    if (Test-Connection -ComputerName $Server -Count 1 -Quiet)
    {
        if ($credential -eq $null)
        {
            $OnPremCred = Get-Credential -UserName 'QH\' -Message "Enter your credential to connect to $server" -ErrorAction Stop
        }
        Else
        {
            $OnPremCred = $credential        
        }
        Write-host "INFO : Trying to connect to Server : $server" -ForegroundColor Cyan
        $session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "http://$server/PowerShell/" -Authentication Kerberos -Credential $OnPremCred
        Import-PSSession $Session -AllowClobber
        Write-host "`nSUCCESS : Successfully connected to : $server" -ForegroundColor Green
    }
    Else
    {
        Throw "ERROR : Server $server is not reachable..."
    }
}
Catch
{
    Write-host "$($error[0].Exception.Message)" -ForegroundColor Magenta
}
