<#	
	.NOTES
	===========================================================================
	 Created with: 	SAPIEN Technologies, Inc., PowerShell Studio 2018 v5.5.152
	 Created on:   	18/05/2018 12:18 PM
	 Created by:   	Rana Banerjee
	 Organization: 	Queensland Health
	 Filename:     	QHO365MigrationOps.psm1 
	===========================================================================
	.DESCRIPTION
		This PowerShell Module is created for the purpose of assisting in Office 365 migration tasks for Queensland Health.
#>

function Connect-QHOnpremExchange
{
<#
	.SYNOPSIS
		A brief description of the Connect-QHOnpremExchange function.
	
	.DESCRIPTION
		A detailed description of the Connect-QHOnpremExchange function.
	
	.PARAMETER Server
		A description of the Server parameter.
	
	.EXAMPLE
		PS C:\> Connect-QHOnpremExchange
	
	.NOTES
		Additional information about the function.
#>
	
	[CmdletBinding()]
	param
	(
		$Server
	)
	
	try
	{
		$StopWatch = [System.Diagnostics.StopWatch]::StartNew()
		$paramNewPSSession = @{
			ConfigurationName = 'Microsoft.Exchange'
			ConnectionUri	  = "http://$Server/PowerShell/"
			Authentication    = 'Kerberos'
			
		}
		#Add-PSSnapin -Name Microsoft.Exchange.Management.PowerShell.E2010
		$Session = New-PSSession @paramNewPSSession
		$paramImportModule = @{
			ModuleInfo = (Import-PSSession $Session -AllowClobber -DisableNameChecking)
			Global	   = $true
			ErrorAction = 'Stop'
			WarningAction = 'SilentlyContinue'
		}
		
		Import-Module @paramImportModule
		$msg = "INFO : Connected to $Server. The Function took $([math]::round($($StopWatch.Elapsed.TotalSeconds), 2)) seconds to Connect to On Premise Exchange"
		Write-Host $msg -ForegroundColor Cyan
	}
	catch
	{
		Write-Host "ERROR : $($_.Exception.Message)" -ForegroundColor Magenta
	}
}

function Connect-QhO365
{
    <#
	.SYNOPSIS
		A brief description of the Connect-OnPremExchange function.
	
	.DESCRIPTION
		A description of the file.
	
	.PARAMETER Server
		A description of the Server parameter.
	
	.PARAMETER credential
		A description of the credential parameter.
	
	.NOTES
		===========================================================================
		Created with: 	SAPIEN Technologies, Inc., PowerShell Studio 2018 v5.5.150
		Created on:   	5/04/2018 6:39 PM
		Created by:   	Rana Banerjee
		Organization: 	Queensland Health
		Filename:
		===========================================================================
#>
	
	param
	(
		[System.Management.Automation.Credential()]
		[ValidateNotNull()]
		[System.Management.Automation.PSCredential]$Credential = [System.Management.Automation.PSCredential]::Empty,
		[Switch]$UseIEProxy
	)
	try
	{
		$FormatEnumerationLimit = -1
		#Write-Host "INFO : Trying to Connect to Office 365" -ForegroundColor Cyan
		
		if ($credential -eq $null)
		{
			$credential = Get-Credential -Message "Enter your Credentials" -ErrorAction Stop
		}
		$paramNewPSSession = @{
			ConfigurationName = 'Microsoft.Exchange'
			ConnectionUri	  = 'https://outlook.office365.com/powershell-liveid/'
			Credential	      = $credential
			Authentication    = 'Basic'
			AllowRedirection  = $true
		}
		if ($UseIEProxy)
		{
			$proxysettings = New-PSSessionOption -ProxyAccessType IEConfig
			$paramNewPSSession.Add('SessionOption', "$proxySettings")
		}
		
		#Write-Host "$($paramNewPSSession | out-string)" -ForegroundColor Cyan
		
		$ExoSession = New-PSSession @paramNewPSSession
		
		$paramImportModule = @{
			ModuleInfo = (Import-PSSession $ExoSession -AllowClobber -DisableNameChecking)
			Global	   = $true
			ErrorAction = 'Stop'
			WarningAction = 'SilentlyContinue'
			Prefix	   = 'EXO'
		}
		Import-Module @paramImportModule
		#Import-Module MsOnline -ErrorAction Stop -Global
		
		#Write-Host "SUCCESS : Successfully Connected to Office 365" -ForegroundColor Green
	}
	catch
	{
		Write-host "ERROR : $($_.Exception.Message)" -ForegroundColor Magenta
	}
}

function Connect-QhMSOLService
{
    <#
	.SYNOPSIS
		A brief description of the Connect-OnPremExchange function.
	
	.DESCRIPTION
		A description of the file.
	
	.PARAMETER Server
		A description of the Server parameter.
	
	.PARAMETER credential
		A description of the credential parameter.
	
	.NOTES
		===========================================================================
		Created with: 	SAPIEN Technologies, Inc., PowerShell Studio 2018 v5.5.150
		Created on:   	5/04/2018 6:39 PM
		Created by:   	Rana Banerjee
		Organization: 	Queensland Health
		Filename:
		===========================================================================
#>
	
	param
	(
		[System.Management.Automation.Credential()]
		[ValidateNotNull()]
		[System.Management.Automation.PSCredential]$Credential = [System.Management.Automation.PSCredential]::Empty
	)
	try
	{
		$FormatEnumerationLimit = -1
		#Write-Host "INFO : Trying to Connect to Office 365" -ForegroundColor Cyan
		
		if ($credential -eq $null)
		{
			$credential = Get-Credential -Message "Enter your Credentials" -ErrorAction Stop
		}
		Connect-MsolService -Credential $Credential -ErrorAction Stop
		#Write-Host "SUCCESS : Successfully Connected to Office 365" -ForegroundColor Green
		return $true
	}
	catch
	{
		Write-host "ERROR : $($_.Exception.Message)" -ForegroundColor Magenta
		return $false
	}
}

function Connect-QhSkypeOnline
{
<#
	.SYNOPSIS
		A brief description of the Connect-OnPremExchange function.
	
	.DESCRIPTION
		A description of the file.
	
	.PARAMETER credential
		A description of the credential parameter.
	
	.PARAMETER Server
		A description of the Server parameter.
	
	.NOTES
		===========================================================================
		Created with: 	SAPIEN Technologies, Inc., PowerShell Studio 2018 v5.5.150
		Created on:   	5/04/2018 6:39 PM
		Created by:   	Rana Banerjee
		Organization: 	Queensland Health
		Filename:
		===========================================================================
#>
	
	[CmdletBinding()]
	param
	(
		[System.Management.Automation.Credential()]
		[ValidateNotNull()]
		[System.Management.Automation.PSCredential]$Credential = [System.Management.Automation.PSCredential]::Empty
	)
	
	try
	{
		$FormatEnumerationLimit = -1
		Write-Host "INFO : Trying to Connect to Skype Online" -ForegroundColor Cyan
		
		Import-Module SkypeOnlineConnector -ErrorAction Stop
		
		$paramNewCsOnlineSession = @{
			Credential		    = $Credential
			OverrideAdminDomain = "healthqld.onmicrosoft.com"
			ErrorAction		    = 'Stop'
		}
		
		$CSsession = New-CsOnlineSession @paramNewCsOnlineSession
		
		$paramImportModule = @{
			ModuleInfo = (Import-PSSession $CSSession -AllowClobber -DisableNameChecking)
			Global	   = $true
			ErrorAction = 'Stop'
			WarningAction = 'SilentlyContinue'
			Prefix	   = '365'
		}
		Import-Module @paramImportModule
		
		#Import-PSSession $Session -Prefix 365 -DisableNameChecking -AllowClobber -ErrorAction Stop >> $null
		
		Write-Host "SUCCESS : Successfully Connected to Skype Online. Commandlet prefix is 365" -ForegroundColor Green
	}
	catch
	{
		Write-host "ERROR : $($_.Exception.Message)" -ForegroundColor Magenta
	}
}

function Get-QHLatestWhatifResult
{
<#
	.SYNOPSIS
		A brief description of the Get-QHLatestWhatifResult function.
	
	.DESCRIPTION
		A detailed description of the Get-QHLatestWhatifResult function.
	
	.PARAMETER ParentDirectory
		A description of the ParentDirectory parameter.
	
	.PARAMETER BatchName
		A description of the BatchName parameter.
	
	.EXAMPLE
		PS C:\> Get-QHLatestWhatifResult
	
	.NOTES
		Additional information about the function.
#>
	param (
		$ParentDirectory = 'D:\Office365\Migrations\Batch',
		$BatchName
	)
	try
	{
		$batch = Get-ChildItem $ParentDirectory |
		Where-Object { ($_.Name -eq $batchName) -and ($_.PSIsContainer -eq 'True') } -ErrorAction Stop |
		Select-Object -ExpandProperty FullName
		
		if ($batch -ne $null)
		{
			$latestwif = Get-ChildItem -Path "$batch\Reports" -Filter *.csv |
			Where-Object { $_.Name -match 'Wif' } | Sort-Object Creationtime -Descending |
			Select-Object -first 1 | Select-Object -ExpandProperty Fullname
			if ($latestwif -ne $null)
			{
				return $latestwif
			}
			else
			{
				throw "No What if results exists. make sure WhatIf checks are run on the batch $BatchName."
			}
			
		}
		else
		{
			throw "Folder $BatchName does not exist at $ParentDirectory"
		}
	}
	catch
	{
		Write-Host "ERROR : $($_.exception.message)" -ForegroundColor Magenta
	}
}

function get-QHADinfo
{
<#
	.SYNOPSIS
		A brief description of the get-QHADinfo function.
	
	.DESCRIPTION
		A detailed description of the get-QHADinfo function.
	
	.PARAMETER UserName
		A description of the UserName parameter.
	
	.EXAMPLE
		PS C:\> get-QHADinfo -UserName 'value1'
	
	.NOTES
		Additional information about the function.
#>
	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory = $true)]
		[String]$UserName
	)
	
	try
	{
		$ad = get-aduser -Filter { UserPrincipalName -eq $UserName } -ErrorAction Stop
		[PSCustomObject][ordered] @{
			User		   = $UserName
			SamAccountName = $ad.SamAccountName
			GivenName	   = $ad.GivenName
			Surname	       = $ad.SurName
		}
	}
	catch
	{
		[PSCustomObject][ordered] @{
			User		   = $UserName
			SamAccountName = 'Could not Retrive. Please Investigate'
			GivenName	   = 'Could not Retrive. Please Investigate'
			Surname	       = 'Could not Retrive. Please Investigate'
		}
	}
}

function get-QHMBStats
{
<#
	.SYNOPSIS
		A brief description of the get-QHMBStats function.
	
	.DESCRIPTION
		A detailed description of the get-QHMBStats function.
	
	.PARAMETER UserName
		A description of the UserName parameter.
	
	.EXAMPLE
		PS C:\> get-QHMBStats -UserName 'value1'
	
	.NOTES
		Additional information about the function.
#>
	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory = $true)]
		[String]$UserName
	)
	
	try
	{
		$MbxSt = Get-MailboxStatistics $UserName -ErrorAction Stop
		[PSCustomObject][ordered] @{
			User		  = $UserName
			TotalItemSize = $MbxSt.TotalItemSize.Value.ToMB()
			ItemCount	  = $MbxSt.ItemCount
		}
		
	}
	catch
	{
		[PSCustomObject][ordered] @{
			User		  = $UserName
			TotalItemSize = 'Could not Retrive. Please Investigate'
			ItemCount	  = 'Could not Retrive. Please Investigate'
		}
	}
}

function get-QHMBInfo
{
<#
	.SYNOPSIS
		A brief description of the get-QHMBInfo function.
	
	.DESCRIPTION
		A detailed description of the get-QHMBInfo function.
	
	.PARAMETER UserName
		A description of the UserName parameter.
	
	.EXAMPLE
		PS C:\> get-QHMBInfo -UserName 'value1'
	
	.NOTES
		Additional information about the function.
#>
	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory = $true)]
		[String]$UserName
	)
	
	try
	{
		$Mbx = Get-Mailbox $UserName -ErrorAction Stop
		[PSCustomObject][ordered] @{
			User	 = $UserName
			LegacyDN = $Mbx.LegacyExchangeDN
		}
	}
	catch
	{
		[PSCustomObject][ordered] @{
			User	 = $UserName
			LegacyDN = 'Could not Retrive. Please Investigate'
		}
	}
}

function Create-QHReferenceList
{
<#
	.SYNOPSIS
		A brief description of the Create-QHReferenceList function.
	
	.DESCRIPTION
		A detailed description of the Create-QHReferenceList function.
	
	.PARAMETER UserCSV
		A description of the UserCSV parameter.
	
	.PARAMETER EVCsv
		A description of the EVCsv parameter.
	
	.EXAMPLE
		PS C:\> Create-QHReferenceList
	
	.NOTES
		Additional information about the function.
#>
	[CmdletBinding()]
	param
	(
		$UserCSV,
		$EVCsv = "D:\Office365\Migrations\Copy-of-EVReports\merged.csv"
	)
	
	begin
	{
		$global:ErrorActionPreference = [System.Management.Automation.ActionPreference]::Stop
		$i = 1
		$err = @{
			ForegroundColor = 'White'
			BackgroundColor = 'Red'
		}
		$OK = @{
			ForegroundColor = 'White'
			BackgroundColor = 'DarkGreen'
		}
		try
		{
			$UserCsvContent = Import-Csv $UserCSV -ErrorAction Stop | Where-Object lookup -eq passed
			$EVCsvContent = Import-Csv $EVCsv -ErrorAction Stop
		}
		catch
		{
			Write-Host "ERROR : $($_.Exception.Message)" @err
			break
		}
	}
	process
	{
		foreach ($line in $UserCsvContent)
		{
			try
			{
				$Adinfo = get-QHADinfo -UserName $line.EmailAddress.Trim()
				$mbxinfo = get-QHMBInfo -UserName $line.EmailAddress.Trim()
				$mbxStat = get-QHMBStats -UserName $line.EmailAddress.Trim()
				
				$prop = [ordered]@{
					EmailAddress = $line.EmailAddress.Trim()
					UserPrimarySMTP = $line.EmailAddress.Trim()
					SamAccountName = $Adinfo.SamAccountName
					GivenName    = $Adinfo.GivenName
					SurName	     = $Adinfo.Surname
					MailboxType  = $line.RecipientTypeDetails
					WhatIfCheck  = $line.Status
					'MailboxSize(MB)' = $mbxStat.TotalItemSize
					MailboxItemCount = $mbxStat.ItemCount
					LegacyExchangeDN = $mbxinfo.LegacyDN
					Details	     = 'None'
				}
				
				if ($Adinfo.Samaccountname -eq $null)
				{
					$prop.Add('ArchiveSizeMB', 'User is Null')
					$prop.Add('ArchiveItemCount', 'User is Null')
				}
				else
				{
					foreach ($EvLine in $EVCsvContent)
					{
						if (($EvLine -ne $null) -and ($EvLine.BillingAccountNameValue.Trim().StartsWith('QH\')))
						{
							if ($EvLine.BillingAccountNameValue.TrimStart('QH\') -eq $Adinfo.SamAccountname)
							{
								$ArchiveSize = $EvLine.TotalArchiveSizeValue
								$ArchiveItemCount = $EvLine.TotalItemsValue
								break
							}
							else
							{
								continue
							}
						}
						else
						{
							continue
						}
					}
					if ($ArchiveSize -ne $null -or $ArchiveItemCount -ne $null)
					{
						$prop.Add('ArchiveSizeMB', $ArchiveSize)
						$prop.Add('ArchiveItemCount', $ArchiveItemCount)
					}
					else
					{
						$prop.Add('ArchiveSizeMB', 'User Not Found')
						$prop.Add('ArchiveItemCount', 'User Not Found')
					}
				}
			}
			catch
			{
				$prop = [ordered]@{
					EmailAddress = $line.EmailAddress.Trim()
					UserPrimarySMTP = $line.EmailAddress.Trim()
					SamAccountName = 'ERROR'
					GivenName    = 'ERROR'
					SurName	     = 'ERROR'
					MailboxType  = $line.RecipientTypeDetails
					WhatIfCheck  = $line.Status
					'MailboxSize(MB)' = 'ERROR'
					MailboxItemCount = 'ERROR'
					LegacyExchangeDN = 'ERROR'
					Details	     = "$($_.Exception.Message)"
				}
			}
			finally
			{
				if ($UserCsvContent.count -gt 1)
				{
					$paramWriteProgress = @{
						Activity = 'Generating Reference List'
						Status   = "Processing [$i] of [$($UserCsvContent.Count)] users"
						PercentComplete = (($i / $UserCsvContent.Count) * 100)
						CurrentOperation = "Completed : [$($line.EmailAddress)]"
					}
					
					Write-Progress @paramWriteProgress
				}
				
				$obj = New-Object -TypeName PSObject -Property $prop
				Write-Output $obj
				$i++
			}
		}
	}
	end
	{
		$global:ErrorActionPreference = [System.Management.Automation.ActionPreference]::Continue
	}
}

function Remove-QHNonAcceptedDomainsEmailAlias
{
<#
	.SYNOPSIS
		Removes or reports  Email addresses with non accepted office 365 domains.
	
	.DESCRIPTION
		This commandlet will remove or report following non accepted domains.
		exchange.health.qld.gov.au, groupwise.qld.gov.au
	
	.PARAMETER UserPrincipalName
		Please enter UPN or EmailAddress
	
	.PARAMETER RemoveDomains
		This will remove the email Alias with the following domains
		exchange.health.qld.gov.au, groupwise.qld.gov.au
	
	.PARAMETER BatchName
		A description of the BatchName parameter.
	
	.EXAMPLE
		Remove-QHNonAcceptedDomainsEmailAlias -UserPrincipalName user1@health.qld.gov.au, user1@health.qld.gov.au -RemoveDomains groupwise.qld.gov.au, exchange.health.qld.gov.au -BatchName Batchx
	
	.NOTES
		Additional information about the function.
#>
	
	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory = $true,
				   ValueFromPipeline = $true,
				   ValueFromPipelineByPropertyName = $true,
				   HelpMessage = 'Please enter UPN or EmailAddress')]
		[ValidateNotNullOrEmpty()]
		[Alias('EmailAddress', 'PrimarySmtpAddress')]
		[String[]]$UserPrincipalName,
		[ValidateSet('exchange.health.qld.gov.au', 'groupwise.qld.gov.au')]
		[String[]]$RemoveDomains,
		[String]$BatchName
	)
	
	begin
	{
		$global:ErrorActionPreference = [System.Management.Automation.ActionPreference]::Stop
		$i = 1
		$domains = Get-EXOAcceptedDomain | Select-Object -ExpandProperty DomainName
	}
	process
	{
		foreach ($UPN in $UserPrincipalName)
		{
			$Accepted = @()
			$notAccepted = @()
			$removal = @()
			
			try
			{
				$UserRecipient = Get-Recipient $UPN -ErrorAction Stop
				if ($UserRecipient.RecipientTypeDetails.ToString() -notmatch 'Remote')
				{
					foreach ($SmtpEmail in $UserRecipient.EmailAddresses.SmtpAddress)
					{
						[MailAddress]$email = $SmtpEmail
						if ($domains -icontains $email.host)
						{
							[Array]$Accepted += $SmtpEmail
						}
						else
						{
							[Array]$notAccepted += $SmtpEmail
							if ($RemoveDomains.count -gt 0)
							{
								$NADRemoval = 'Remove'
								if ($RemoveDomains -icontains $email.host)
								{
									[Array]$removal += $SmtpEmail
								}
								else
								{
									continue
								}
							}
							else
							{
								$NADRemoval = 'ReportOnly'
							}
						}
					}
					
					if ($notAccepted.count -eq 0)
					{
						$notAccepted += 'None'
						$Action = 'NotRequired'
						$NADRemoval = 'NotApplicable'
					}
					else
					{
						if ($removal.Count -gt 0)
						{
							try
							{
								Set-Mailbox $UPN -EmailAddresses @{ remove = $removal } -ErrorAction Stop -WarningAction SilentlyContinue
								$Action = "Removed : $(($removal -join ',' | Out-String).Trim())"
							}
							catch
							{
								$Action = "FailedToRemove: $(($removal -join ',' | Out-String).Trim())"
							}
						}
						else
						{
							$Action = 'NothingToRemove'
						}
					}
					
					$Emlprop = [ordered]@{
						UPN		     = $UPN
						EmailAddress = $UserRecipient.PrimarySmtpAddress.ToString()
						BatchName    = $BatchName
						MailBoxType  = $UserRecipient.RecipientTypeDetails.ToString()
						AcceptedEmail = "$(($Accepted -join ',' | Out-String).Trim())"
						NonAcceptedEmail = "$(($notAccepted -join ',' | Out-String).Trim())"
						NonAcceptedDomainAction = $NADRemoval
						Action	     = $Action
						Details	     = 'Processed'
					}
				}
				else
				{
					$Emlprop = [ordered]@{
						UPN		     = $UPN
						EmailAddress = $UserRecipient.PrimarySmtpAddress.ToString()
						BatchName    = $BatchName
						MailBoxType  = $UserRecipient.RecipientTypeDetails.ToString()
						AcceptedEmail = 'None'
						NonAcceptedEmail = 'None'
						NonAcceptedDomainAction = 'None'
						Action	     = 'None'
						Details	     = 'ERROR : Already a cloud user'
					}
				}
			}
			catch
			{
				$Emlprop = [ordered]@{
					UPN					    = $UPN
					EmalAddress			    = 'ERROR'
					BatchName			    = $BatchName
					MailBoxType			    = 'ERROR'
					AcceptedEmail		    = 'ERROR'
					NonAcceptedEmail	    = 'ERROR'
					NonAcceptedDomainAction = 'ERROR'
					Action				    = 'ERROR'
					Details				    = "ERROR : $($_.Exception.Message)"
				}
			}
			finally
			{
				$obj = New-Object -TypeName System.Management.Automation.PSObject -Property $Emlprop
				Write-Output $obj
				if ($UserPrincipalName.count -gt 1)
				{
					$paramWriteProgress = @{
						Activity = 'Validating EmailAddresses Against Accepted Domains'
						Status   = "Processing [$i] of [$($UserPrincipalName.Count)] users"
						PercentComplete = (($i / $UserPrincipalName.Count) * 100)
						CurrentOperation = "Completed : [$UPN]"
					}
					
					Write-Progress @paramWriteProgress
				}
				$i++
				$Accepted = $null
				$notAccepted = $null
				$removal = $null
			}
		}
	}
	end
	{
		Write-Progress -Activity 'Validating EmailAddresses Against Accepted Domains' -Completed
		$global:ErrorActionPreference = [System.Management.Automation.ActionPreference]::Continue
	}
}

function Validate-QHMigrationUsers
{
<#
	.SYNOPSIS
		A brief description of the Validate-QHMigrationUsers function.
	
	.DESCRIPTION
		A detailed description of the Validate-QHMigrationUsers function.
	
	.PARAMETER UserPrincipalName
		A description of the UserPrincipalName parameter.
	
	.PARAMETER BatchName
		A description of the BatchName parameter.
	
	.EXAMPLE
		PS C:\> Validate-QHMigrationUsers -UserPrincipalName 'value1'
	
	.NOTES
		Additional information about the function.
#>
	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory = $true,
				   ValueFromPipeline = $true,
				   ValueFromPipelineByPropertyName = $true,
				   HelpMessage = 'Please enter UPN or EmailAddress')]
		[ValidateNotNullOrEmpty()]
		[Alias('EmailAddress', 'PrimarySmtpAddress')]
		[String[]]$UserPrincipalName,
		[String]$BatchName = 'None'
	)
	
	begin
	{
		$global:ErrorActionPreference = [System.Management.Automation.ActionPreference]::Stop
		$i = 1
		$HHSContext = @{
			"200ADL.CO.STH.HEALTH"					    = "DoH"
			"51WEMBLEY.LOGAN.STH.HEALTH"			    = "Metro South"
			"61MARY.CO.STH.HEALTH"					    = "DoH"
			"ABIOS.BS.STH.HEALTH"					    = "Metro South"
			"ADH.TBL.FNQ.HEALTH"					    = "Cairns and Hinterland"
			"ALLIED.TPCH.Chermside.BNN.HEALTH"		    = "Metro North"
			"ARCHIVED.SVC.FNQ.HEALTH"				    = "Cairns and Hinterland"
			"Ashgrove.NTH-BNE.BNN.HEALTH"			    = "Metro North"
			"Ashworth-House.ON-BNE.BNN.HEALTH"		    = "Metro North"
			"Aspley.ON-BNE.BNN.HEALTH"				    = "Metro North"
			"Ayr.BWN.NTH.HEALTH"					    = "Townsville"
			"BAB.INN.FNQ.HEALTH"					    = "Cairns and Hinterland"
			"Bald-Hills.ON-BNE.BNN.HEALTH"			    = "Metro North"
			"Baralaba-H.BANANA.CTL.HEALTH"			    = "Central Queensland"
			"Barcaldine-H.CWEST.CTL.HEALTH"			    = "Central West"
			"BAY.FRASER.WBY.HEALTH"					    = "Wide Bay"
			"BBG.BUNDY.WBY.HEALTH"					    = "Wide Bay"
			"BCH.TORRES.FNQ.HEALTH"					    = "Torres and Cape"
			"BDH.TORRES.FNQ.HEALTH"					    = "Torres and Cape"
			"Beaudesert-H.LOGAN.STH.HEALTH"			    = "Metro South"
			"Beenleigh-CH.LOGAN.STH.HEALTH"			    = "Metro South"
			"BHHS.TWB.SWQ.HEALTH"					    = "Darling Downs"
			"Biala.IN-BNE.BNN.HEALTH"				    = "Metro North"
			"Biloela-H.BANANA.CTL.HEALTH"			    = "Central Queensland"
			"Blackall-H.CWEST.CTL.HEALTH"			    = "Central West"
			"Blackwater-H.CHIGH.CTL.HEALTH"			    = "Central Queensland"
			"Boonah.WM.SWQ.HEALTH"					    = "West Moreton"
			"Bowen.BWN.NTH.HEALTH"					    = "Mackay"
			"Brighton.ON-BNE.BNN.HEALTH"			    = "Metro North"
			"Browns-ACC.LOGAN.STH.HEALTH"			    = "Metro South"
			"BSC.TPCH.Chermside.BNN.HEALTH"			    = "Metro North"
			"BSQ.QEII.QEIIHD.STH.HEALTH"			    = "Metro South"
			"BSS.QHSS.BS.STH.HEALTH"				    = "Metro South"
			"CANNONH.BS.STH.HEALTH"					    = "Metro South"
			"CARD_THOR.TPCH.Chermside.BNN.HEALTH"	    = "Metro North"
			"CBH.CNS.FNQ.HEALTH"					    = "Cairns and Hinterland"
			"CFH.CO.STH.HEALTH"						    = "DoH"
			"CH.QEII.QEIIHD.STH.HEALTH"				    = "Metro South"
			"Char-H.CHAR.SWQ.HEALTH"				    = "South West"
			"CHE.SBUR.WBY.HEALTH"					    = "Wide Bay"
			"Chin-H.NDowns.SWQ.HEALTH"				    = "Darling Downs"
			"Chtwrs.CT.NTH.HEALTH"					    = "Townsville"
			"Clermont-H.MKY.NTH.HEALTH"				    = "Mackay"
			"CLIENTS.BeachRd.SCG.BNN.HEALTH"		    = "Sunshine Coast"
			"CLIENTS.Birtinya.SCG.BNN.HEALTH"		    = "Sunshine Coast"
			"CLIENTS.BottleBrush.SCG.BNN.HEALTH"	    = "Sunshine Coast"
			"CLIENTS.BrisbaneRd.SCG.BNN.HEALTH"		    = "Sunshine Coast"
			"CLIENTS.Caloundra.SCG.BNN.HEALTH"		    = "Sunshine Coast"
			"CLIENTS.Gympie.SCG.BNN.HEALTH"			    = "Sunshine Coast"
			"CLIENTS.HortonPde.SCG.BNN.HEALTH"		    = "Sunshine Coast"
			"Clients.Kilcoy-H.Red-Cab.BNN.HEALTH"	    = "Metro North"
			"CLIENTS.Maleny.SCG.BNN.HEALTH"			    = "Sunshine Coast"
			"CLIENTS.Musgrave.SCG.BNN.HEALTH"		    = "Sunshine Coast"
			"CLIENTS.Nambour.SCG.BNN.HEALTH"		    = "Sunshine Coast"
			"Clients.Red-Cab.BNN.HEALTH"			    = "Metro North"
			"CLIENTS.SCUH.SCG.BNN.HEALTH"			    = "Sunshine Coast"
			"CLIENTS.SixthAve.SCG.BNN.HEALTH"		    = "Sunshine Coast"
			"CLON-H.ISA.NTH.HEALTH"					    = "North West"
			"Collinsville.BWN.NTH.HEALTH"			    = "Mackay"
			"COMMUNITY.TPCH.Chermside.BNN.HEALTH"	    = "Metro North"
			"ComPlz.WM.SWQ.HEALTH"					    = "West Moreton"
			"COOKTOWN.CNS-REG.FNQ.HEALTH"			    = "Torres and Cape"
			"Coorparoo.QEIIHD.STH.HEALTH"			    = "Metro South"
			"Corinda.QEIIHD.STH.HEALTH"				    = "Metro South"
			"CORP.PAH.BS.STH.HEALTH"				    = "Metro South"
			"CORP.TPCH.Chermside.BNN.HEALTH"		    = "Metro North"
			"CSS.PAH.BS.STH.HEALTH"					    = "Metro South"
			"Dalby-H.NDowns.SWQ.HEALTH"				    = "Darling Downs"
			"Disabled users.MKY.NTH.HEALTH"			    = "Mackay"
			"DisabledUsers.ISA-H.ISA.NTH.HEALTH"	    = "North West"
			"DisabledUsers.MORN-H.ISA.NTH.HEALTH"	    = "North West"
			"DOOM-H.ISA.NTH.HEALTH"					    = "North West"
			"DTS.IFS.IS.HEALTH"						    = "DoH"
			"Dysart-H.MKY.NTH.HEALTH"				    = "Mackay"
			"EDMON.CNS-REG.FNQ.HEALTH"				    = "Cairns and Hinterland"
			"Emerald-H.CHIGH.CTL.HEALTH"			    = "Central Queensland"
			"Esk.WM.SWQ.HEALTH"						    = "Darling Downs"
			"Eventide.CT.NTH.HEALTH"				    = "Townsville"
			"Eventide-NH.ROCK.CTL.HEALTH"			    = "Central Queensland"
			"EXPIRED.GCH.GC.STH.HEALTH"				    = "Gold Coast"
			"EXPIRED.Helensvale.GC.STH.HEALTH"		    = "Gold Coast"
			"EXPIRED.SHP.GC.STH.HEALTH"				    = "Gold Coast"
			"EXTERNAL.CBH.CNS.FNQ.HEALTH"			    = "Cairns and Hinterland"
			"External.MKY.NTH.HEALTH"				    = "Mackay"
			"EXTERNAL.PAH.BS.STH.HEALTH"			    = "Metro South"
			"EXTERNAL.TGH.TSV.NTH.HEALTH"			    = "Townsville"
			"FS.QHSS.BS.STH.HEALTH"					    = "Metro South"
			"GARDC.BS.STH.HEALTH"					    = "Metro South"
			"GCH.GC.STH.HEALTH"						    = "Gold Coast"
			"GCTECH.GC.STH.HEALTH"					    = "Gold Coast"
			"GDH.NBUR.WBY.HEALTH"					    = "Wide Bay"
			"GHS.TWB.SWQ.HEALTH"					    = "Darling Downs"
			"Gladstone-H.ROCK.CTL.HEALTH"			    = "Central Queensland"
			"GMO.CO.STH.HEALTH"						    = "DoH"
			"Goodna.WM.SWQ.HEALTH"					    = "West Moreton"
			"Goondi-H.SDowns.SWQ.HEALTH"			    = "Darling Downs"
			"GORDON.CNS-REG.FNQ.HEALTH"				    = "Cairns and Hinterland"
			"Greenslopes.QEIIHD.STH.HEALTH"			    = "Metro South"
			"GWISE.SVC.DTS.IFS.IS.HEALTH"			    = "DoH"
			"HDH.TBL.FNQ.HEALTH"					    = "Cairns and Hinterland"
			"Helensvale.GC.STH.HEALTH"				    = "Gold Coast"
			"Homehill.BWN.NTH.HEALTH"				    = "Townsville"
			"Hughenden.CT.NTH.HEALTH"				    = "Townsville"
			"HVB.FRASER.WBY.HEALTH"					    = "Wide Bay"
			"ID.CO.STH.HEALTH"						    = "DoH"
			"IDD.PAH.BS.STH.HEALTH"					    = "Metro South"
			"IDH.INN.FNQ.HEALTH"					    = "Cairns and Hinterland"
			"IGH.WM.SWQ.HEALTH"						    = "West Moreton"
			"IMSU.TPCH.Chermside.BNN.HEALTH"		    = "Metro North"
			"Inactive.Blackwater-H.CHIGH.CTL.HEALTH"    = "Central Queensland"
			"Inactive.Emerald-H.CHIGH.CTL.HEALTH"	    = "Central Queensland"
			"Inactive.Gladstone-H.ROCK.CTL.HEALTH"	    = "Central Queensland"
			"Inactive.Rockhampton-CH.ROCK.CTL.HEALTH"   = "Central Queensland"
			"Inactive.Rockhampton-H.ROCK.CTL.HEALTH"    = "Central Queensland"
			"Inactive.Rockhampton-PH.ROCK.CTL.HEALTH"   = "Central Queensland"
			"Inactive.Springsure-H.CHIGH.CTL.HEALTH"    = "Central Queensland"
			"Inactive.Yeppoon-H.ROCK.CTL.HEALTH"	    = "Central Queensland"
			"Inactive-Users.TGH.TSV.NTH.HEALTH"		    = "Townsville"
			"InalaCYMH.QEIIHD.STH.HEALTH"			    = "Metro South"
			"Ingham.TSV.NTH.HEALTH"					    = "Townsville"
			"INGLE-H.SDowns.SWQ.HEALTH"				    = "Darling Downs"
			"IS.PAH.BS.STH.HEALTH"					    = "Metro South"
			"ISA-H.ISA.NTH.HEALTH"					    = "North West"
			"Jando-H.NDowns.SWQ.HEALTH"				    = "Darling Downs"
			"June2014.EXpUsers.User2.QHB.CO.STH.HEALTH" = "DoH"
			"Kingston-OH.LOGAN.STH.HEALTH"			    = "Metro South"
			"Kirwan.TSV.NTH.HEALTH"					    = "Townsville"
			"KRY.SBUR.WBY.HEALTH"					    = "Wide Bay"
			"Laidley.WM.SWQ.HEALTH"					    = "West Moreton"
			"LOGAN-H.LOGAN.STH.HEALTH"				    = "Metro South"
			"Longreach-CH.CWEST.CTL.HEALTH"			    = "Central West"
			"Longreach-DO.CWEST.CTL.HEALTH"			    = "Central West"
			"Longreach-H.CWEST.CTL.HEALTH"			    = "Central West"
			"Mackay-CH.MKY.NTH.HEALTH"				    = "Mackay"
			"Mackay-H.MKY.NTH.HEALTH"				    = "Mackay"
			"MAIN.InalaCH.QEIIHD.STH.HEALTH"		    = "Metro South"
			"Main.QEII.QEIIHD.STH.HEALTH"			    = "Metro South"
			"MBH.FRASER.WBY.HEALTH"					    = "Wide Bay"
			"MBNC.MBC.BAYSD.STH.HEALTH"				    = "Metro South"
			"MDH.TBL.FNQ.HEALTH"					    = "Cairns and Hinterland"
			"Meadow.LOGAN.STH.HEALTH"				    = "Metro South"
			"MED.PAH.BS.STH.HEALTH"					    = "Metro South"
			"MEDICAL.TPCH.Chermside.BNN.HEALTH"		    = "Metro North"
			"MENTAL.CNS.FNQ.HEALTH"					    = "Cairns and Hinterland"
			"MH.PAH.BS.STH.HEALTH"					    = "Metro South"
			"Miles-H.NDowns.SWQ.HEALTH"				    = "Darling Downs"
			"Mill-H.SDowns.SWQ.HEALTH"				    = "Darling Downs"
			"MLNH.TWB.SWQ.HEALTH"					    = "Darling Downs"
			"Moranbah-H.MKY.NTH.HEALTH"				    = "Mackay"
			"MORN-H.ISA.NTH.HEALTH"					    = "North West"
			"Mosman.CT.NTH.HEALTH"					    = "Townsville"
			"MOSSMAN.CNS-REG.FNQ.HEALTH"			    = "Cairns and Hinterland"
			"Moura-H.BANANA.CTL.HEALTH"				    = "Central Queensland"
			"MtMorgan-H.ROCK.CTL.HEALTH"			    = "Central Queensland"
			"Nathan.TSV.NTH.HEALTH"					    = "Townsville"
			"NerangDent.GCDental.GC.STH.HEALTH"		    = "Gold Coast"
			"NORM-H.ISA.NTH.HEALTH"					    = "North West"
			"NorthWard.TSV.NTH.HEALTH"				    = "Townsville"
			"NorthWest.NTH-BNE.BNN.HEALTH"			    = "Metro North"
			"Nth-Rton-NH.ROCK.CTL.HEALTH"			    = "Central Queensland"
			"NURSING.TPCH.Chermside.BNN.HEALTH"		    = "Metro North"
			"ODM.PAH.BS.STH.HEALTH"					    = "Metro South"
			"OHS.TWB.SWQ.HEALTH"					    = "Darling Downs"
			"Oral.QEII.QEIIHD.STH.HEALTH"			    = "Metro South"
			"PalmBeachCH.GC.STH.HEALTH"				    = "Gold Coast"
			"PalmIs.TSV.NTH.HEALTH"					    = "Townsville"
			"PATH.PAH.BS.STH.HEALTH"				    = "Metro South"
			"PCH.TPCH.Chermside.BNN.HEALTH"			    = "Metro North"
			"PHS.QHSS.BS.STH.HEALTH"				    = "Metro South"
			"PHU.PAH.BS.STH.HEALTH"					    = "Metro South"
			"PineRivers-CH.ON-BNE.BNN.HEALTH"		    = "Metro North"
			"Prime.GC.STH.HEALTH"					    = "Gold Coast"
			"Proserpine-H.MKY.NTH.HEALTH"			    = "Mackay"
			"PRT.CO.STH.HEALTH"						    = "DoH"
			"RAD.CO.STH.HEALTH"						    = "DoH"
			"RAD-ONC.BS.STH.HEALTH"					    = "Metro South"
			"RCFNQ.CBH.CNS.FNQ.HEALTH"				    = "Cairns and Hinterland"
			"Red-H.RHC.BAYSD.STH.HEALTH"			    = "Metro South"
			"REHAB.PAH.BS.STH.HEALTH"				    = "Metro South"
			"RHP.GC.STH.HEALTH"						    = "Gold Coast"
			"RICHLNDS.BS.STH.HEALTH"				    = "Metro South"
			"Richmond.CT.NTH.HEALTH"				    = "Townsville"
			"Robina.GC.STH.HEALTH"					    = "Gold Coast"
			"Rockhampton-CH.ROCK.CTL.HEALTH"		    = "Central Queensland"
			"Rockhampton-H.ROCK.CTL.HEALTH"			    = "Central Queensland"
			"Rockhampton-PH.ROCK.CTL.HEALTH"		    = "Central Queensland"
			"Roma-CH.Roma.SWQ.HEALTH"				    = "South West"
			"Roma-DWS.Roma.SWQ.HEALTH"				    = "South West"
			"Roma-H.Roma.SWQ.HEALTH"				    = "South West"
			"Roma-ORH.Roma.SWQ.HEALTH"				    = "Darling Downs"
			"Sarina-H.MKY.NTH.HEALTH"				    = "Mackay"
			"SMITH.CNS-REG.FNQ.HEALTH"				    = "Cairns and Hinterland"
			"Springsure-H.CHIGH.CTL.HEALTH"			    = "Central Queensland"
			"SSI.PAH.BS.STH.HEALTH"					    = "Metro South"
			"Stan-H.SDowns.SWQ.HEALTH"				    = "Darling Downs"
			"StGeorge-H.Roma.SWQ.HEALTH"			    = "South West"
			"SthBnePH.BS.STH.HEALTH"				    = "Metro South"
			"SURG.PAH.BS.STH.HEALTH"				    = "Metro South"
			"System.Object[]"						    = "System.Object[]"
			"Tara-H.NDowns.SWQ.HEALTH"				    = "Darling Downs"
			"Taroom-H.NDowns.SWQ.HEALTH"			    = "Darling Downs"
			"TESTOU.DTS.IFS.IS.HEALTH"				    = ""
			"Texas-H.SDowns.SWQ.HEALTH"				    = "Darling Downs"
			"TGH.TSV.NTH.HEALTH"					    = "Townsville"
			"Theodore-H.BANANA.CTL.HEALTH"			    = "Central Queensland"
			"THS.TWB.SWQ.HEALTH"					    = "Darling Downs"
			"TI-HOSP.TORRES.FNQ.HEALTH"				    = "Torres and Cape"
			"TI-PRIM.TORRES.FNQ.HEALTH"				    = "Torres and Cape"
			"TOP.BS.STH.HEALTH"						    = "Metro South"
			"TPCH.Chermside.BNN.HEALTH"				    = "Metro North"
			"TPHU.CNS.FNQ.HEALTH"					    = "Torres and Cape"
			"TULLY.INN.FNQ.HEALTH"					    = "Cairns and Hinterland"
			"USER.179GREY.CO.STH.HEALTH"			    = "DoH"
			"USER.199GREY.CO.STH.HEALTH"			    = "DoH"
			"USER.BDH.IN-BNE.BNS.HEALTH"			    = "Metro North"
			"USER.Carrara.GC.STH.HEALTH"			    = "Gold Coast"
			"USER.ChapelSt.ON-BNE.BNS.HEALTH"		    = "Metro South"
			"USER.Enoggera.ON-BNE.BNS.HEALTH"		    = "Metro South"
			"USER.Finney-Rd.ON-BNE.BNS.HEALTH"		    = "Metro North"
			"USER.GARU.ON-BNE.BNS.HEALTH"			    = "Metro South"
			"USER.Halwyn.ON-BNE.BNS.HEALTH"			    = "Metro South"
			"USER.Helensvale.GC.STH.HEALTH"			    = "Gold Coast"
			"USER.Herston.IN-BNE.BNS.HEALTH"		    = "Metro North"
			"USER.ID.CO.STH.HEALTH"					    = "DoH"
			"USER.LCCH.CHQ.STH.HEALTH"				    = "Children's Health Queensland"
			"USER.MarinePde.GC.STH.HEALTH"			    = "Gold Coast"
			"USER.NerangSSP.GC.STH.HEALTH"			    = "Gold Coast"
			"USER.Nundah.ON-BNE.BNS.HEALTH"			    = "Metro South"
			"USER.Nundah-CH.ON-BNE.BNS.HEALTH"		    = "Metro South"
			"USER.PalmBeachCH.GC.STH.HEALTH"		    = "Gold Coast"
			"USER.QEII.QEIIHD.STH.HEALTH"			    = "Metro South"
			"USER.Robina.GC.STH.HEALTH"				    = "Gold Coast"
			"USER.SHP.GC.STH.HEALTH"				    = "Gold Coast"
			"USER.Stafford.ON-BNE.BNS.HEALTH"		    = "Metro North"
			"User1.QHB.CO.STH.HEALTH"				    = "DoH"
			"User2.QHB.CO.STH.HEALTH"				    = "DoH"
			"USERS.LOGAN-H.LOGAN.STH.HEALTH"		    = "Metro South"
			"VIL.FRASER.WBY.HEALTH"					    = "Wide Bay"
			"Vincent.TSV.NTH.HEALTH"				    = "Townsville"
			"Warehouse.TSV.NTH.HEALTH"				    = "Townsville"
			"WEB-EXT.CO.STH.HEALTH"					    = "DoH"
			"Whitsunday-CH.MKY.NTH.HEALTH"			    = "Mackay"
			"WHS.SDowns.SWQ.HEALTH"					    = "Darling Downs"
			"Winton-H.CWEST.CTL.HEALTH"				    = "Central West"
			"Woodridge-CH.LOGAN.STH.HEALTH"			    = "Metro South"
			"Woorabinda-H.ROCK.CTL.HEALTH"			    = "Central Queensland"
			"WPA-HOSP.CAPE.FNQ.HEALTH"				    = "Torres and Cape"
			"WPH.WM.SWQ.HEALTH"						    = "West Moreton"
			"WUJAL.CNS-REG.FNQ.HEALTH"				    = "Torres and Cape"
			"WYN-H.WHC.BAYSD.STH.HEALTH"			    = "Metro South"
			"YBH-HOSP.CNS-REG.FNQ.HEALTH"			    = "Cairns and Hinterland"
			"Yeppoon-H.ROCK.CTL.HEALTH"				    = "Central Queensland"
			"YerongaCYMH.QEIIHD.STH.HEALTH"			    = "Metro South"
		}
		
		try
		{
			$null = Get-ExoAcceptedDomain -ErrorAction Stop
		}
		catch
		{
			Write-Warning "Please make sure that you are connected to Exchange Online with session prefix 'EXO'."
			break
		}
		
		$null = Invoke-Command -Session (Get-PSSession | Where-Object { $_.ComputerName -eq 'outlook.office365.com' }) -ScriptBlock { Get-MigrationUser -Resultsize 'Unlimited' | Select-object Identity, BatchId } |
		Select-Object @{ n = 'Identity'; e = { $_.Identity.ToString() } }, @{ n = 'BatchID'; e = { $_.BatchId.ToString() } } -OutVariable ExistingMoveRequests
		
	}
	process
	{
		foreach ($UPN in $UserPrincipalName)
		{
			if ($UPN.Contains('@'))
			{
				if ($upn.Trim().Split("@")[1].ToLower().StartsWith('h'))
				{
					$User = $UPN.Trim().Split("@")[0] + '@health.qld.gov.au'
				}
				elseif ($upn.Trim().Split("@")[1].ToLower().StartsWith('m'))
				{
					$User = $UPN.Trim().Split("@")[0] + '@mhrt.qld.gov.au'
				}
				elseif ($upn.Trim().Split("@")[1].ToLower().StartsWith('t'))
				{
					$User = $UPN.Trim().Split("@")[0] + '@tpchfoundation.org.au'
				}
				else { }
				
				#$User = $UPN.Trim().Split("@")[0] + '@health.qld.gov.au'
				
				if ($UPN -ne $user)
				{
					$rectified = 'True'
				}
				else
				{
					$rectified = 'False'
				}
			}
			else
			{
				$User = $UPN.Trim()
				$rectified = 'N/A'
			}
			try
			{
				$mbx = Get-Recipient $User -ErrorAction Stop
				$HHS = $HHSContext["$($mbx.CustomAttribute10)"]
				if ($HHS -eq $null)
				{
					$HHS = 'Could not Lookup HHS'
				}
				if ($mbx.RecipientTypeDetails.ToString() -eq 'UserMailbox')
				{
					$ad = get-aduser $mbx.SamAccountName -ErrorAction 'Stop'
					
					if ($ad.UserPrincipalName -ne $null)
					{
						if ($ad.UserPrincipalName -ne $mbx.PrimarySMTPAddress.ToString())
						{
							$UPNCheck = 'MisMatch'
							$lookup = 'FAILED'
							$reason = "FAILED : UPN:$($ad.UserPrincipalName) did not match SMTP:$($mbx.PrimarySMTPAddress.ToString())"
						}
						else
						{
							$UPNCheck = 'OK'
							$lookup = 'PASSED'
							$reason = 'None'
						}
					}
					else
					{
						$UPNCheck = 'MisMatch'
						$lookup = 'FAILED'
						$reason = "FAILED : UPN not found"
					}
					
					$moveReq = $ExistingMoveRequests.Where({ $_.Identity -eq $Mbx.PrimarySmtpAddress.ToString() })
					
					if ($moveReq -ne $null)
					{
						if ($lookup -eq 'PASSED')
						{
							$lookup = 'FAILED'
							$reason = "Existing Move Request [$($moveReq.Identity) in $($moveReq.BatchID)]"
						}
						else
						{
							$reason = $reason + " & Existing Move Request [$($moveReq.Identity) in $($moveReq.BatchID)]"
						}
					}
					
					
					$prop = [ordered]@{
						FromOrignalCSV = $UPN
						Rectified	   = $rectified
						Batch		   = $BatchName
						EmailAddress   = $mbx.PrimarySMTPAddress.ToString()
						RecipientTypeDetails = $mbx.RecipientTypeDetails.ToString()
						SamAccountName = $mbx.SamAccountName
						HHS		       = $HHS
						UPNCheck	   = $UPNCheck
						Lookup		   = $lookup
						Details	       = $reason
					}
				}
				elseif ($mbx.RecipientTypeDetails.ToString() -eq 'SharedMailbox')
				{
					$ad = get-aduser $mbx.SamAccountName -ErrorAction 'Stop'
					
					$dom = $mbx.PrimarySmtpAddress.ToString().Split('@')[1]
					
					if ($ad.UserPrincipalName -ne $null)
					{
						if (!($ad.UserPrincipalName.EndsWith("$dom")))
						{
							$UPNCheck = 'MisMatch'
							$lookup = 'FAILED'
							$reason = "FAILED : UPN:$($ad.UserPrincipalName) did not have domain ending in ($dom)"
						}
						else
						{
							$UPNCheck = 'OK'
							$lookup = 'PASSED'
							$reason = 'None'
						}
					}
					else
					{
						$UPNCheck = 'MisMatch'
						$lookup = 'FAILED'
						$reason = "FAILED : UPN not found"
					}
					
					
					$moveReq = $ExistingMoveRequests.Where({ $_.Identity -eq $Mbx.PrimarySmtpAddress.ToString() })
					
					if ($moveReq -ne $null)
					{
						if ($lookup -eq 'PASSED')
						{
							$lookup = 'FAILED'
							$reason = "Existing Move Request [$($moveReq.Identity) in $($moveReq.BatchID)]"
						}
						else
						{
							$reason = $reason + " & Existing Move Request [$($moveReq.Identity) in $($moveReq.BatchID)]"
						}
					}
					
					$prop = [ordered]@{
						FromOrignalCSV = $UPN
						Rectified	   = $rectified
						Batch		   = $BatchName
						EmailAddress   = $mbx.PrimarySMTPAddress.ToString()
						RecipientTypeDetails = $mbx.RecipientTypeDetails.ToString()
						SamAccountName = $mbx.SamAccountName
						HHS		       = $HHS
						UPNCheck	   = $UPNCheck
						Lookup		   = $lookup
						Details	       = $reason
					}
				}
				elseif ($mbx.RecipientTypeDetails.ToString() -eq 'RoomMailbox')
				{
					$ad = get-aduser $mbx.SamAccountName -ErrorAction 'Stop'
					$dom = $mbx.PrimarySmtpAddress.ToString().Split('@')[1]
					
					if ($ad.UserPrincipalName -ne $null)
					{
						if (!($ad.UserPrincipalName.EndsWith("$dom")))
						{
							$UPNCheck = 'MisMatch'
							$lookup = 'FAILED'
							$reason = "FAILED : UPN:$($ad.UserPrincipalName) did not have domain ending in ($dom)"
						}
						else
						{
							$UPNCheck = 'OK'
							$lookup = 'PASSED'
							$reason = 'None'
						}
					}
					else
					{
						$UPNCheck = 'MisMatch'
						$lookup = 'FAILED'
						$reason = "FAILED : UPN not found"
					}
					
					
					$moveReq = $ExistingMoveRequests.Where({ $_.Identity -eq $Mbx.PrimarySmtpAddress.ToString() })
					
					if ($moveReq -ne $null)
					{
						if ($lookup -eq 'PASSED')
						{
							$lookup = 'FAILED'
							$reason = "Existing Move Request [$($moveReq.Identity) in $($moveReq.BatchID)]"
						}
						else
						{
							$reason = $reason + " & Existing Move Request [$($moveReq.Identity) in $($moveReq.BatchID)]"
						}
					}
					
					$prop = [ordered]@{
						FromOrignalCSV = $UPN
						Rectified	   = $rectified
						Batch		   = $BatchName
						EmailAddress   = $mbx.PrimarySMTPAddress.ToString()
						RecipientTypeDetails = $mbx.RecipientTypeDetails.ToString()
						SamAccountName = $mbx.SamAccountName
						HHS		       = $HHS
						UPNCheck	   = $UPNCheck
						Lookup		   = $lookup
						Details	       = $reason
					}
				}
				elseif ($mbx.RecipientTypeDetails.ToString().startsWith('Remote'))
				{
					$prop = [ordered]@{
						FromOrignalCSV = $UPN
						Rectified	   = $rectified
						Batch		   = $BatchName
						EmailAddress   = $mbx.PrimarySMTPAddress.ToString()
						RecipientTypeDetails = $mbx.RecipientTypeDetails.ToString()
						SamAccountName = $mbx.SamAccountName
						HHS		       = $HHS
						UPNCheck	   = 'Skipped'
						Lookup		   = 'SKIPPED'
						Details	       = 'Already Migrated'
					}
				}
				else
				{
					$prop = [ordered]@{
						FromOrignalCSV = $UPN
						Rectified	   = $rectified
						Batch		   = $BatchName
						EmailAddress   = $mbx.PrimarySMTPAddress.ToString()
						RecipientTypeDetails = $mbx.RecipientTypeDetails.ToString()
						SamAccountName = $mbx.SamAccountName
						HHS		       = $HHS
						UPNCheck	   = 'Skipped'
						Lookup		   = 'SKIPPED'
						Details	       = 'Unknown Mailbox'
					}
				}
			}
			catch
			{
				$prop = [ordered]@{
					FromOrignalCSV	     = $UPN
					Rectified		     = $rectified
					Batch			     = $BatchName
					EmailAddress		 = 'ERROR'
					RecipientTypeDetails = 'ERROR'
					SamAccountName	     = 'ERROR'
					HHS				     = 'ERROR'
					UPNCheck			 = 'ERROR'
					Lookup			     = 'FAILED'
					Details			     = "ERROR : $($_.Exception.Message)"
				}
			}
			finally
			{
				if ($UserPrincipalName.count -gt 1)
				{
					$paramWriteProgress = @{
						Activity = 'Validating EmailAddresses'
						Status   = "Processing [$i] of [$($UserPrincipalName.Count)] users"
						PercentComplete = (($i / $UserPrincipalName.Count) * 100)
						CurrentOperation = "Completed : [$UPN]"
					}
					
					Write-Progress @paramWriteProgress
				}
				$i++
				$obj = New-Object -TypeName System.Management.Automation.PSObject -Property $prop
				Write-Output $obj
			}
		}
	}
	end
	{
		Write-Progress -Activity 'Validating EmailAddresses' -Completed
		$global:ErrorActionPreference = [System.Management.Automation.ActionPreference]::Continue
	}
}

function Set-QHdir
{
<#
	.SYNOPSIS
		A brief description of the Set-QHdir function.
	
	.DESCRIPTION
		A detailed description of the Set-QHdir function.
	
	.PARAMETER Path
		A description of the Path parameter.
	
	.EXAMPLE
		PS C:\> Set-QHdir -Path 'value1'
	
	.NOTES
		Additional information about the function.
#>
	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory = $true)]
		[String]$Path
	)
	
	try
	{
		$msg = "INFO : $($MyInvocation.InvocationName) `t`t Checking if the Log Directory exists [$Path]"
		Write-Host "$msg" -ForegroundColor Yellow
		#$msg >> $Runlogfile
		if (! (Test-Path $Path -ErrorAction Stop))
		{
			$msg = "INFO : $($MyInvocation.InvocationName) `t`t path [$path] did not exist. Trying to create it now"
			Write-Host "$msg" -ForegroundColor Yellow
			mkdir $Path -Force -ErrorAction Stop | Out-Null
			$msg = "SUCCESS : $($MyInvocation.InvocationName) `t`t Created path [$path]"
			Write-Host "$msg" -ForegroundColor Green
			return $path
		}
		else
		{
			$msg = "INFO : $($MyInvocation.InvocationName) `t`t path [$path] Exists... No Action taken"
			Write-Host "$msg" -ForegroundColor Yellow
			return $Path
		}
	}
	catch
	{
		$Errormsg = "ERROR : $($MyInvocation.InvocationName) `t`t $($_.exception.message)"
		Write-Host $Errormsg -ForegroundColor magenta
	}
}

function Split-Array
{
<#
	.SYNOPSIS
		A brief description of the Split-Array function.
	
	.DESCRIPTION
		A detailed description of the Split-Array function.
	
	.PARAMETER inArray
		A description of the inArray parameter.
	
	.PARAMETER parts
		A description of the parts parameter.
	
	.PARAMETER size
		A description of the size parameter.
	
	.EXAMPLE
		PS C:\> Split-Array
	
	.NOTES
		Additional information about the function.
#>
	param (
		$inArray,
		[int]$parts,
		[int]$size
	)
	if ($parts)
	{
		$PartSize = [Math]::Ceiling($inArray.count / $parts)
	}
	if ($size)
	{
		$PartSize = $size
		$parts = [Math]::Ceiling($inArray.count / $size)
	}
	
	$outArray = New-Object 'System.Collections.Generic.List[psobject]'
	
	for ($i = 1; $i -le $parts; $i++)
	{
		$start = (($i - 1) * $PartSize)
		$end = (($i) * $PartSize) - 1
		if ($end -ge $inArray.count) { $end = $inArray.count - 1 }
		$outArray.Add(@($inArray[$start .. $end]))
	}
	return, $outArray
}

function Get-duration
{
<#
	.SYNOPSIS
		A brief description of the Get-duration function.
	
	.DESCRIPTION
		A detailed description of the Get-duration function.
	
	.PARAMETER StartTime
		A description of the StartTime parameter.
	
	.PARAMETER EndTime
		A description of the EndTime parameter.
	
	.EXAMPLE
		PS C:\> Get-duration
	
	.NOTES
		Additional information about the function.
#>
	param (
		$StartTime,
		$EndTime
	)
	try
	{
		$ErrorActionPreference = 'Stop'
		
		if ($EndTime -gt $StartTime)
		{
			$ts = $EndTime - $StartTime
			$tsobj = '{0:00}:{1:00}:{2:00}' -f $($ts.Hours + $ts.days * 24), $ts.Minutes, $ts.Seconds
			
			$prop = [ordered] @{
				StartTime = $StartTime.toString('dd-MM-yyyy HH:mm:ss')
				EndTime   = $EndTime.toString('dd-MM-yyyy HH:mm:ss')
				Duration  = $tsobj + ' (hh:mm:ss)'
			}
			$obj = New-Object -TypeName System.Management.Automation.PSObject -Property $prop
			Write-Output $obj
		}
		else
		{
			throw "StartTime:[$($StartTime.toString('dd-MM-yyyy HH:mm:ss'))] Cannot be After EndTime:[$($EndTime.toString('dd-MM-yyyy HH:mm:ss'))]  "
		}
		$ErrorActionPreference = 'Continue'
	}
	catch
	{
		Write-Host "ERROR : $($_.Exception.Message)" -ForegroundColor Magenta
		return $null
	}
}

function Get-QHHHSInfo
{
<#
	.SYNOPSIS
		A brief description of the Get-QHHHSInfo function.
	
	.DESCRIPTION
		A detailed description of the Get-QHHHSInfo function.
	
	.PARAMETER PrimarySmtpAddress
		A description of the PrimarySmtpAddress parameter.
	
	.EXAMPLE
		PS C:\> Get-QHHHSInfo -PrimarySmtpAddress 'value1'
	
	.NOTES
		Additional information about the function.
#>
	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory = $true,
				   ValueFromPipeline = $true,
				   ValueFromPipelineByPropertyName = $true)]
		[Alias('Mailbox', 'SamAccountName', 'UserPrincipalName', 'SMTPAddress')]
		[String[]]$PrimarySmtpAddress
	)
	begin
	{
		$HHSContext = @{
			"200ADL.CO.STH.HEALTH"					    = "DoH"
			"51WEMBLEY.LOGAN.STH.HEALTH"			    = "Metro South"
			"61MARY.CO.STH.HEALTH"					    = "DoH"
			"ABIOS.BS.STH.HEALTH"					    = "Metro South"
			"ADH.TBL.FNQ.HEALTH"					    = "Cairns and Hinterland"
			"ALLIED.TPCH.Chermside.BNN.HEALTH"		    = "Metro North"
			"ARCHIVED.SVC.FNQ.HEALTH"				    = "Cairns and Hinterland"
			"Ashgrove.NTH-BNE.BNN.HEALTH"			    = "Metro North"
			"Ashworth-House.ON-BNE.BNN.HEALTH"		    = "Metro North"
			"Aspley.ON-BNE.BNN.HEALTH"				    = "Metro North"
			"Ayr.BWN.NTH.HEALTH"					    = "Townsville"
			"BAB.INN.FNQ.HEALTH"					    = "Cairns and Hinterland"
			"Bald-Hills.ON-BNE.BNN.HEALTH"			    = "Metro North"
			"Baralaba-H.BANANA.CTL.HEALTH"			    = "Central Queensland"
			"Barcaldine-H.CWEST.CTL.HEALTH"			    = "Central West"
			"BAY.FRASER.WBY.HEALTH"					    = "Wide Bay"
			"BBG.BUNDY.WBY.HEALTH"					    = "Wide Bay"
			"BCH.TORRES.FNQ.HEALTH"					    = "Torres and Cape"
			"BDH.TORRES.FNQ.HEALTH"					    = "Torres and Cape"
			"Beaudesert-H.LOGAN.STH.HEALTH"			    = "Metro South"
			"Beenleigh-CH.LOGAN.STH.HEALTH"			    = "Metro South"
			"BHHS.TWB.SWQ.HEALTH"					    = "Darling Downs"
			"Biala.IN-BNE.BNN.HEALTH"				    = "Metro North"
			"Biloela-H.BANANA.CTL.HEALTH"			    = "Central Queensland"
			"Blackall-H.CWEST.CTL.HEALTH"			    = "Central West"
			"Blackwater-H.CHIGH.CTL.HEALTH"			    = "Central Queensland"
			"Boonah.WM.SWQ.HEALTH"					    = "West Moreton"
			"Bowen.BWN.NTH.HEALTH"					    = "Mackay"
			"Brighton.ON-BNE.BNN.HEALTH"			    = "Metro North"
			"Browns-ACC.LOGAN.STH.HEALTH"			    = "Metro South"
			"BSC.TPCH.Chermside.BNN.HEALTH"			    = "Metro North"
			"BSQ.QEII.QEIIHD.STH.HEALTH"			    = "Metro South"
			"BSS.QHSS.BS.STH.HEALTH"				    = "Metro South"
			"CANNONH.BS.STH.HEALTH"					    = "Metro South"
			"CARD_THOR.TPCH.Chermside.BNN.HEALTH"	    = "Metro North"
			"CBH.CNS.FNQ.HEALTH"					    = "Cairns and Hinterland"
			"CFH.CO.STH.HEALTH"						    = "DoH"
			"CH.QEII.QEIIHD.STH.HEALTH"				    = "Metro South"
			"Char-H.CHAR.SWQ.HEALTH"				    = "South West"
			"CHE.SBUR.WBY.HEALTH"					    = "Wide Bay"
			"Chin-H.NDowns.SWQ.HEALTH"				    = "Darling Downs"
			"Chtwrs.CT.NTH.HEALTH"					    = "Townsville"
			"Clermont-H.MKY.NTH.HEALTH"				    = "Mackay"
			"CLIENTS.BeachRd.SCG.BNN.HEALTH"		    = "Sunshine Coast"
			"CLIENTS.Birtinya.SCG.BNN.HEALTH"		    = "Sunshine Coast"
			"CLIENTS.BottleBrush.SCG.BNN.HEALTH"	    = "Sunshine Coast"
			"CLIENTS.BrisbaneRd.SCG.BNN.HEALTH"		    = "Sunshine Coast"
			"CLIENTS.Caloundra.SCG.BNN.HEALTH"		    = "Sunshine Coast"
			"CLIENTS.Gympie.SCG.BNN.HEALTH"			    = "Sunshine Coast"
			"CLIENTS.HortonPde.SCG.BNN.HEALTH"		    = "Sunshine Coast"
			"Clients.Kilcoy-H.Red-Cab.BNN.HEALTH"	    = "Metro North"
			"CLIENTS.Maleny.SCG.BNN.HEALTH"			    = "Sunshine Coast"
			"CLIENTS.Musgrave.SCG.BNN.HEALTH"		    = "Sunshine Coast"
			"CLIENTS.Nambour.SCG.BNN.HEALTH"		    = "Sunshine Coast"
			"Clients.Red-Cab.BNN.HEALTH"			    = "Metro North"
			"CLIENTS.SCUH.SCG.BNN.HEALTH"			    = "Sunshine Coast"
			"CLIENTS.SixthAve.SCG.BNN.HEALTH"		    = "Sunshine Coast"
			"CLON-H.ISA.NTH.HEALTH"					    = "North West"
			"Collinsville.BWN.NTH.HEALTH"			    = "Mackay"
			"COMMUNITY.TPCH.Chermside.BNN.HEALTH"	    = "Metro North"
			"ComPlz.WM.SWQ.HEALTH"					    = "West Moreton"
			"COOKTOWN.CNS-REG.FNQ.HEALTH"			    = "Torres and Cape"
			"Coorparoo.QEIIHD.STH.HEALTH"			    = "Metro South"
			"Corinda.QEIIHD.STH.HEALTH"				    = "Metro South"
			"CORP.PAH.BS.STH.HEALTH"				    = "Metro South"
			"CORP.TPCH.Chermside.BNN.HEALTH"		    = "Metro North"
			"CSS.PAH.BS.STH.HEALTH"					    = "Metro South"
			"Dalby-H.NDowns.SWQ.HEALTH"				    = "Darling Downs"
			"Disabled users.MKY.NTH.HEALTH"			    = "Mackay"
			"DisabledUsers.ISA-H.ISA.NTH.HEALTH"	    = "North West"
			"DisabledUsers.MORN-H.ISA.NTH.HEALTH"	    = "North West"
			"DOOM-H.ISA.NTH.HEALTH"					    = "North West"
			"DTS.IFS.IS.HEALTH"						    = "DoH"
			"Dysart-H.MKY.NTH.HEALTH"				    = "Mackay"
			"EDMON.CNS-REG.FNQ.HEALTH"				    = "Cairns and Hinterland"
			"Emerald-H.CHIGH.CTL.HEALTH"			    = "Central Queensland"
			"Esk.WM.SWQ.HEALTH"						    = "Darling Downs"
			"Eventide.CT.NTH.HEALTH"				    = "Townsville"
			"Eventide-NH.ROCK.CTL.HEALTH"			    = "Central Queensland"
			"EXPIRED.GCH.GC.STH.HEALTH"				    = "Gold Coast"
			"EXPIRED.Helensvale.GC.STH.HEALTH"		    = "Gold Coast"
			"EXPIRED.SHP.GC.STH.HEALTH"				    = "Gold Coast"
			"EXTERNAL.CBH.CNS.FNQ.HEALTH"			    = "Cairns and Hinterland"
			"External.MKY.NTH.HEALTH"				    = "Mackay"
			"EXTERNAL.PAH.BS.STH.HEALTH"			    = "Metro South"
			"EXTERNAL.TGH.TSV.NTH.HEALTH"			    = "Townsville"
			"FS.QHSS.BS.STH.HEALTH"					    = "Metro South"
			"GARDC.BS.STH.HEALTH"					    = "Metro South"
			"GCH.GC.STH.HEALTH"						    = "Gold Coast"
			"GCTECH.GC.STH.HEALTH"					    = "Gold Coast"
			"GDH.NBUR.WBY.HEALTH"					    = "Wide Bay"
			"GHS.TWB.SWQ.HEALTH"					    = "Darling Downs"
			"Gladstone-H.ROCK.CTL.HEALTH"			    = "Central Queensland"
			"GMO.CO.STH.HEALTH"						    = "DoH"
			"Goodna.WM.SWQ.HEALTH"					    = "West Moreton"
			"Goondi-H.SDowns.SWQ.HEALTH"			    = "Darling Downs"
			"GORDON.CNS-REG.FNQ.HEALTH"				    = "Cairns and Hinterland"
			"Greenslopes.QEIIHD.STH.HEALTH"			    = "Metro South"
			"GWISE.SVC.DTS.IFS.IS.HEALTH"			    = "DoH"
			"HDH.TBL.FNQ.HEALTH"					    = "Cairns and Hinterland"
			"Helensvale.GC.STH.HEALTH"				    = "Gold Coast"
			"Homehill.BWN.NTH.HEALTH"				    = "Townsville"
			"Hughenden.CT.NTH.HEALTH"				    = "Townsville"
			"HVB.FRASER.WBY.HEALTH"					    = "Wide Bay"
			"ID.CO.STH.HEALTH"						    = "DoH"
			"IDD.PAH.BS.STH.HEALTH"					    = "Metro South"
			"IDH.INN.FNQ.HEALTH"					    = "Cairns and Hinterland"
			"IGH.WM.SWQ.HEALTH"						    = "West Moreton"
			"IMSU.TPCH.Chermside.BNN.HEALTH"		    = "Metro North"
			"Inactive.Blackwater-H.CHIGH.CTL.HEALTH"    = "Central Queensland"
			"Inactive.Emerald-H.CHIGH.CTL.HEALTH"	    = "Central Queensland"
			"Inactive.Gladstone-H.ROCK.CTL.HEALTH"	    = "Central Queensland"
			"Inactive.Rockhampton-CH.ROCK.CTL.HEALTH"   = "Central Queensland"
			"Inactive.Rockhampton-H.ROCK.CTL.HEALTH"    = "Central Queensland"
			"Inactive.Rockhampton-PH.ROCK.CTL.HEALTH"   = "Central Queensland"
			"Inactive.Springsure-H.CHIGH.CTL.HEALTH"    = "Central Queensland"
			"Inactive.Yeppoon-H.ROCK.CTL.HEALTH"	    = "Central Queensland"
			"Inactive-Users.TGH.TSV.NTH.HEALTH"		    = "Townsville"
			"InalaCYMH.QEIIHD.STH.HEALTH"			    = "Metro South"
			"Ingham.TSV.NTH.HEALTH"					    = "Townsville"
			"INGLE-H.SDowns.SWQ.HEALTH"				    = "Darling Downs"
			"IS.PAH.BS.STH.HEALTH"					    = "Metro South"
			"ISA-H.ISA.NTH.HEALTH"					    = "North West"
			"Jando-H.NDowns.SWQ.HEALTH"				    = "Darling Downs"
			"June2014.EXpUsers.User2.QHB.CO.STH.HEALTH" = "DoH"
			"Kingston-OH.LOGAN.STH.HEALTH"			    = "Metro South"
			"Kirwan.TSV.NTH.HEALTH"					    = "Townsville"
			"KRY.SBUR.WBY.HEALTH"					    = "Wide Bay"
			"Laidley.WM.SWQ.HEALTH"					    = "West Moreton"
			"LOGAN-H.LOGAN.STH.HEALTH"				    = "Metro South"
			"Longreach-CH.CWEST.CTL.HEALTH"			    = "Central West"
			"Longreach-DO.CWEST.CTL.HEALTH"			    = "Central West"
			"Longreach-H.CWEST.CTL.HEALTH"			    = "Central West"
			"Mackay-CH.MKY.NTH.HEALTH"				    = "Mackay"
			"Mackay-H.MKY.NTH.HEALTH"				    = "Mackay"
			"MAIN.InalaCH.QEIIHD.STH.HEALTH"		    = "Metro South"
			"Main.QEII.QEIIHD.STH.HEALTH"			    = "Metro South"
			"MBH.FRASER.WBY.HEALTH"					    = "Wide Bay"
			"MBNC.MBC.BAYSD.STH.HEALTH"				    = "Metro South"
			"MDH.TBL.FNQ.HEALTH"					    = "Cairns and Hinterland"
			"Meadow.LOGAN.STH.HEALTH"				    = "Metro South"
			"MED.PAH.BS.STH.HEALTH"					    = "Metro South"
			"MEDICAL.TPCH.Chermside.BNN.HEALTH"		    = "Metro North"
			"MENTAL.CNS.FNQ.HEALTH"					    = "Cairns and Hinterland"
			"MH.PAH.BS.STH.HEALTH"					    = "Metro South"
			"Miles-H.NDowns.SWQ.HEALTH"				    = "Darling Downs"
			"Mill-H.SDowns.SWQ.HEALTH"				    = "Darling Downs"
			"MLNH.TWB.SWQ.HEALTH"					    = "Darling Downs"
			"Moranbah-H.MKY.NTH.HEALTH"				    = "Mackay"
			"MORN-H.ISA.NTH.HEALTH"					    = "North West"
			"Mosman.CT.NTH.HEALTH"					    = "Townsville"
			"MOSSMAN.CNS-REG.FNQ.HEALTH"			    = "Cairns and Hinterland"
			"Moura-H.BANANA.CTL.HEALTH"				    = "Central Queensland"
			"MtMorgan-H.ROCK.CTL.HEALTH"			    = "Central Queensland"
			"Nathan.TSV.NTH.HEALTH"					    = "Townsville"
			"NerangDent.GCDental.GC.STH.HEALTH"		    = "Gold Coast"
			"NORM-H.ISA.NTH.HEALTH"					    = "North West"
			"NorthWard.TSV.NTH.HEALTH"				    = "Townsville"
			"NorthWest.NTH-BNE.BNN.HEALTH"			    = "Metro North"
			"Nth-Rton-NH.ROCK.CTL.HEALTH"			    = "Central Queensland"
			"NURSING.TPCH.Chermside.BNN.HEALTH"		    = "Metro North"
			"ODM.PAH.BS.STH.HEALTH"					    = "Metro South"
			"OHS.TWB.SWQ.HEALTH"					    = "Darling Downs"
			"Oral.QEII.QEIIHD.STH.HEALTH"			    = "Metro South"
			"PalmBeachCH.GC.STH.HEALTH"				    = "Gold Coast"
			"PalmIs.TSV.NTH.HEALTH"					    = "Townsville"
			"PATH.PAH.BS.STH.HEALTH"				    = "Metro South"
			"PCH.TPCH.Chermside.BNN.HEALTH"			    = "Metro North"
			"PHS.QHSS.BS.STH.HEALTH"				    = "Metro South"
			"PHU.PAH.BS.STH.HEALTH"					    = "Metro South"
			"PineRivers-CH.ON-BNE.BNN.HEALTH"		    = "Metro North"
			"Prime.GC.STH.HEALTH"					    = "Gold Coast"
			"Proserpine-H.MKY.NTH.HEALTH"			    = "Mackay"
			"PRT.CO.STH.HEALTH"						    = "DoH"
			"RAD.CO.STH.HEALTH"						    = "DoH"
			"RAD-ONC.BS.STH.HEALTH"					    = "Metro South"
			"RCFNQ.CBH.CNS.FNQ.HEALTH"				    = "Cairns and Hinterland"
			"Red-H.RHC.BAYSD.STH.HEALTH"			    = "Metro South"
			"REHAB.PAH.BS.STH.HEALTH"				    = "Metro South"
			"RHP.GC.STH.HEALTH"						    = "Gold Coast"
			"RICHLNDS.BS.STH.HEALTH"				    = "Metro South"
			"Richmond.CT.NTH.HEALTH"				    = "Townsville"
			"Robina.GC.STH.HEALTH"					    = "Gold Coast"
			"Rockhampton-CH.ROCK.CTL.HEALTH"		    = "Central Queensland"
			"Rockhampton-H.ROCK.CTL.HEALTH"			    = "Central Queensland"
			"Rockhampton-PH.ROCK.CTL.HEALTH"		    = "Central Queensland"
			"Roma-CH.Roma.SWQ.HEALTH"				    = "South West"
			"Roma-DWS.Roma.SWQ.HEALTH"				    = "South West"
			"Roma-H.Roma.SWQ.HEALTH"				    = "South West"
			"Roma-ORH.Roma.SWQ.HEALTH"				    = "Darling Downs"
			"Sarina-H.MKY.NTH.HEALTH"				    = "Mackay"
			"SMITH.CNS-REG.FNQ.HEALTH"				    = "Cairns and Hinterland"
			"Springsure-H.CHIGH.CTL.HEALTH"			    = "Central Queensland"
			"SSI.PAH.BS.STH.HEALTH"					    = "Metro South"
			"Stan-H.SDowns.SWQ.HEALTH"				    = "Darling Downs"
			"StGeorge-H.Roma.SWQ.HEALTH"			    = "South West"
			"SthBnePH.BS.STH.HEALTH"				    = "Metro South"
			"SURG.PAH.BS.STH.HEALTH"				    = "Metro South"
			"System.Object[]"						    = "System.Object[]"
			"Tara-H.NDowns.SWQ.HEALTH"				    = "Darling Downs"
			"Taroom-H.NDowns.SWQ.HEALTH"			    = "Darling Downs"
			"TESTOU.DTS.IFS.IS.HEALTH"				    = ""
			"Texas-H.SDowns.SWQ.HEALTH"				    = "Darling Downs"
			"TGH.TSV.NTH.HEALTH"					    = "Townsville"
			"Theodore-H.BANANA.CTL.HEALTH"			    = "Central Queensland"
			"THS.TWB.SWQ.HEALTH"					    = "Darling Downs"
			"TI-HOSP.TORRES.FNQ.HEALTH"				    = "Torres and Cape"
			"TI-PRIM.TORRES.FNQ.HEALTH"				    = "Torres and Cape"
			"TOP.BS.STH.HEALTH"						    = "Metro South"
			"TPCH.Chermside.BNN.HEALTH"				    = "Metro North"
			"TPHU.CNS.FNQ.HEALTH"					    = "Torres and Cape"
			"TULLY.INN.FNQ.HEALTH"					    = "Cairns and Hinterland"
			"USER.179GREY.CO.STH.HEALTH"			    = "DoH"
			"USER.199GREY.CO.STH.HEALTH"			    = "DoH"
			"USER.BDH.IN-BNE.BNS.HEALTH"			    = "Metro North"
			"USER.Carrara.GC.STH.HEALTH"			    = "Gold Coast"
			"USER.ChapelSt.ON-BNE.BNS.HEALTH"		    = "Metro South"
			"USER.Enoggera.ON-BNE.BNS.HEALTH"		    = "Metro South"
			"USER.Finney-Rd.ON-BNE.BNS.HEALTH"		    = "Metro North"
			"USER.GARU.ON-BNE.BNS.HEALTH"			    = "Metro South"
			"USER.Halwyn.ON-BNE.BNS.HEALTH"			    = "Metro South"
			"USER.Helensvale.GC.STH.HEALTH"			    = "Gold Coast"
			"USER.Herston.IN-BNE.BNS.HEALTH"		    = "Metro North"
			"USER.ID.CO.STH.HEALTH"					    = "DoH"
			"USER.LCCH.CHQ.STH.HEALTH"				    = "Children's Health Queensland"
			"USER.MarinePde.GC.STH.HEALTH"			    = "Gold Coast"
			"USER.NerangSSP.GC.STH.HEALTH"			    = "Gold Coast"
			"USER.Nundah.ON-BNE.BNS.HEALTH"			    = "Metro South"
			"USER.Nundah-CH.ON-BNE.BNS.HEALTH"		    = "Metro South"
			"USER.PalmBeachCH.GC.STH.HEALTH"		    = "Gold Coast"
			"USER.QEII.QEIIHD.STH.HEALTH"			    = "Metro South"
			"USER.Robina.GC.STH.HEALTH"				    = "Gold Coast"
			"USER.SHP.GC.STH.HEALTH"				    = "Gold Coast"
			"USER.Stafford.ON-BNE.BNS.HEALTH"		    = "Metro North"
			"User1.QHB.CO.STH.HEALTH"				    = "DoH"
			"User2.QHB.CO.STH.HEALTH"				    = "DoH"
			"USERS.LOGAN-H.LOGAN.STH.HEALTH"		    = "Metro South"
			"VIL.FRASER.WBY.HEALTH"					    = "Wide Bay"
			"Vincent.TSV.NTH.HEALTH"				    = "Townsville"
			"Warehouse.TSV.NTH.HEALTH"				    = "Townsville"
			"WEB-EXT.CO.STH.HEALTH"					    = "DoH"
			"Whitsunday-CH.MKY.NTH.HEALTH"			    = "Mackay"
			"WHS.SDowns.SWQ.HEALTH"					    = "Darling Downs"
			"Winton-H.CWEST.CTL.HEALTH"				    = "Central West"
			"Woodridge-CH.LOGAN.STH.HEALTH"			    = "Metro South"
			"Woorabinda-H.ROCK.CTL.HEALTH"			    = "Central Queensland"
			"WPA-HOSP.CAPE.FNQ.HEALTH"				    = "Torres and Cape"
			"WPH.WM.SWQ.HEALTH"						    = "West Moreton"
			"WUJAL.CNS-REG.FNQ.HEALTH"				    = "Torres and Cape"
			"WYN-H.WHC.BAYSD.STH.HEALTH"			    = "Metro South"
			"YBH-HOSP.CNS-REG.FNQ.HEALTH"			    = "Cairns and Hinterland"
			"Yeppoon-H.ROCK.CTL.HEALTH"				    = "Central Queensland"
			"YerongaCYMH.QEIIHD.STH.HEALTH"			    = "Metro South"
		}
	}
	process
	{
		foreach ($user in $PrimarySmtpAddress)
		{
			try
			{
				$mbx = Get-Recipient $user -ErrorAction Stop
				$readableHHS = $HHSContext["$($mbx.CustomAttribute10)"]
				if ($readableHHS -eq '')
				{
					$readableHHS = 'Could not Lookup HHS'
				}
				
				$prop = [ordered]@{
					User = $user
					HHS  = $mbx.CustomAttribute10.ToString()
					ReadableHHS = $readableHHS
					Details = 'None'
				}
			}
			catch
			{
				$prop = [ordered]@{
					User	    = $user
					HHS		    = 'None'
					ReadableHHS = 'None'
					Details	    = "ERROR : $($_.Exception.Message)"
				}
			}
			finally
			{
				$obj = New-Object -TypeName System.Management.Automation.PSObject -Property $prop
				Write-Output $obj
			}
		}
	}
	end
	{
	}
}

function Export-QHMailboxInfo
{
<#
	.SYNOPSIS
		A brief description of the Export-QHMailboxInfo function.
	
	.DESCRIPTION
		A detailed description of the Export-QHMailboxInfo function.
	
	.PARAMETER UserName
		A description of the UserName parameter.
	
	.PARAMETER SnapshotFolder
		A description of the SnapshotFolder parameter.
	
	.PARAMETER batchName
		A description of the batchName parameter.
	
	.EXAMPLE
		PS C:\> Export-QHMailboxInfo
	
	.NOTES
		Additional information about the function.
#>
	[CmdletBinding()]
	param
	(
		[String[]]$UserName,
		[String]$SnapshotFolder = 'SnapshotTextFiles',
		[String]$batchName
	)
	
	begin
	{
		$global:ErrorActionPreference = [System.Management.Automation.ActionPreference]::Stop
		$date = Get-date
		if (Navigate-QHMigrationFolder $batchName)
		{
		}
		$batchdir = Get-QHbatchDirectory -batchName $batchName
		$logfolder = Set-QHdir -path "$batchdir\SnapShotTextFiles"
		$i = 1
	}
	process
	{
		foreach ($user in $UserName)
		{
			$AuditFile = $null
			
			$fileusr = $user
			
			[System.IO.Path]::GetInvalidFileNameChars() | ForEach-Object { $fileusr = $fileusr.replace($_, '_') }
			
			$AuditFile = "$logfolder\$($fileusr)_$($date.ToString('yyyyMMdd-HHmmss')).txt"
			"$user Audit : $date" >> $AuditFile
			
			"------ Mailbox Settings --------------------------------------------------------" >> $AuditFile
			$MbxAudit = Get-QHmailboxInfoFullAudit -UserName $user
			$MbxAudit >> $AuditFile
			
			"------ Mailbox Statistics ------------------------------------------------------" >> $AuditFile
			$MbxStatAudit = Get-QHmailboxStatsFullAudit -UserName $user
			$MbxAudit >> $AuditFile
			
			"------ Mailbox Full Access Permissions -----------------------------------------" >> $AuditFile
			$MbxFullAccess = Get-QHMailboxFullAccessPermissionFullAudit -UserName $user
			$MbxFullAccess >> $AuditFile
			
			"------ Mailbox Send-as Permissions ---------------------------------------------" >> $AuditFile
			$MbxSendAs = Get-SendAsPermissionFullAudit -UserName $user
			$MbxSendAs >> $AuditFile
			$AuditFile = $null
			
			if ($UserName.count -gt 1)
			{
				$paramWriteProgress = @{
					Activity = 'Exporting Mailbox Informaiton'
					Status   = "Processing [$i] of [$($UserName.Count)] users"
					PercentComplete = (($i / $UserName.Count) * 100)
					CurrentOperation = "Completed : [$user]"
				}
				Write-Progress @paramWriteProgress
			}
			$i++
		}
		
		Write-Progress -Activity 'Exporting Mailbox Informaiton' -Completed
	}
	end
	{
		$global:ErrorActionPreference = [System.Management.Automation.ActionPreference]::Continue
	}
}

function Get-QHmailboxInfoFullAudit
{
<#
	.SYNOPSIS
		A brief description of the Get-QHmailboxInfoFullAudit function.
	
	.DESCRIPTION
		A detailed description of the Get-QHmailboxInfoFullAudit function.
	
	.PARAMETER UserName
		A description of the UserName parameter.
	
	.EXAMPLE
		PS C:\> Get-QHmailboxInfoFullAudit
	
	.NOTES
		Additional information about the function.
#>
	[CmdletBinding()]
	param
	(
		$UserName
	)
	
	try
	{
		$mbx = Get-Mailbox $UserName -ErrorAction Stop | Select-Object *
		return $mbx
	}
	catch
	{
		$errormsg = "ERROR : $($_.Exception.Message)"
		return $errormsg
	}
}

function Get-QHmailboxStatsFullAudit
{
<#
	.SYNOPSIS
		A brief description of the Get-QHmailboxStatsFullAudit function.
	
	.DESCRIPTION
		A detailed description of the Get-QHmailboxStatsFullAudit function.
	
	.PARAMETER UserName
		A description of the UserName parameter.
	
	.EXAMPLE
		PS C:\> Get-QHmailboxStatsFullAudit
	
	.NOTES
		Additional information about the function.
#>
	[CmdletBinding()]
	param
	(
		$UserName
	)
	
	try
	{
		$mbxStats = Get-MailboxStatistics $UserName -ErrorAction Stop | Select-Object *
		return $mbxStats
	}
	catch
	{
		$errormsg = "ERROR : $($_.Exception.Message)"
		return $errormsg
	}
}

function Get-QHMailboxFullAccessPermissionFullAudit
{
<#
	.SYNOPSIS
		A brief description of the Get-QHMailboxFullAccessPermissionFullAudit function.
	
	.DESCRIPTION
		A detailed description of the Get-QHMailboxFullAccessPermissionFullAudit function.
	
	.PARAMETER UserName
		A description of the UserName parameter.
	
	.EXAMPLE
		PS C:\> Get-QHMailboxFullAccessPermissionFullAudit
	
	.NOTES
		Additional information about the function.
#>
	[CmdletBinding()]
	param
	(
		$UserName
	)
	
	try
	{
		$FullAccess = Get-MailboxPermission $UserName -ErrorAction Stop |
		Where-Object { ($_.User -notlike "*SELF*") -and ($_.IsInherited -eq $false) } | Select-Object *
		return $FullAccess
	}
	catch
	{
		$errormsg = "ERROR : $($_.Exception.Message)"
		return $errormsg
	}
}

function Get-SendAsPermissionFullAudit
{
<#
	.SYNOPSIS
		A brief description of the Get-SendAsPermissionFullAudit function.
	
	.DESCRIPTION
		A detailed description of the Get-SendAsPermissionFullAudit function.
	
	.PARAMETER UserName
		A description of the UserName parameter.
	
	.EXAMPLE
		PS C:\> Get-SendAsPermissionFullAudit
	
	.NOTES
		Additional information about the function.
#>
	[CmdletBinding()]
	param
	(
		$UserName
	)
	
	try
	{
		$SendAs = Get-Mailbox $UserName -ErrorAction Stop | Get-ADPermission |
		Where-Object { ($_.User -notlike "*SELF*") -and ($_.extendedrights -like "*send-as*") -and ($_.IsInherited -eq $false) } |
		Select-Object *
		return $SendAs
	}
	catch
	{
		$errormsg = "ERROR : $($_.Exception.Message)"
		return $errormsg
	}
}

function Get-QHbatchDirectory
{
<#
	.SYNOPSIS
		A brief description of the Get-QHbatchDirectory function.
	
	.DESCRIPTION
		A detailed description of the Get-QHbatchDirectory function.
	
	.PARAMETER ParentDirectory
		A description of the ParentDirectory parameter.
	
	.PARAMETER BatchName
		A description of the BatchName parameter.
	
	.EXAMPLE
		PS C:\> Get-QHbatchDirectory
	
	.NOTES
		Additional information about the function.
#>
	[CmdletBinding()]
	param
	(
		$ParentDirectory = 'D:\Office365\Migrations\Batch',
		$BatchName
	)
	
	try
	{
		$batch = Get-ChildItem $ParentDirectory -ErrorAction Stop |
		Where-Object { ($_.Name -eq $batchName) -and ($_.PSIsContainer -eq 'True') } |
		Select-Object -ExpandProperty FullName
		
		if ($batch -ne $null)
		{
			return $batch
		}
		else
		{
			throw "Folder $BatchName does not exist at $ParentDirectory"
		}
	}
	catch
	{
		Write-Host "ERROR : $($_.exception.message)" -ForegroundColor Magenta
		return $null
	}
}

function Remove-QHNonAcceptedEmails
{
<#
	.SYNOPSIS
		A brief description of the Remove-QHNonAcceptedEmails function.
	
	.DESCRIPTION
		A detailed description of the Remove-QHNonAcceptedEmails function.
	
	.PARAMETER UserPrincipalName
		A description of the UserPrincipalName parameter.
	
	.PARAMETER domains
		A description of the domains parameter.
	
	.EXAMPLE
		PS C:\> Remove-QHNonAcceptedEmails -UserPrincipalName 'value1' -domains $domains
	
	.NOTES
		Additional information about the function.
#>
	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory = $true,
				   ValueFromPipeline = $true,
				   ValueFromPipelineByPropertyName = $true,
				   HelpMessage = 'Please enter UPN or EmailAddress')]
		[ValidateNotNullOrEmpty()]
		[Alias('EmailAddress', 'PrimarySmtpAddress')]
		[String[]]$UserPrincipalName,
		[Parameter(Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		[Array]$domains = @('groupwise.qld.gov.au', 'exchange.health.qld.gov.au')
	)
	
	begin
	{
		
	}
	process
	{
		foreach ($UPN in $UserPrincipalName)
		{
			try
			{
				$UserRecipient = Get-Recipient $UPN -ErrorAction Stop
				if ($UserRecipient.RecipientTypeDetails.ToString() -notmatch 'Remote')
				{
					foreach ($SmtpEmail in $UserRecipient.EmailAddresses.SmtpAddress)
					{
						[MailAddress]$email = $SmtpEmail
						if ($domains -icontains $email.host)
						{
							Set-Mailbox $Upn -EmailAddresses @{ remove = "$SmtpEmail" } -ErrorAction Stop -WarningAction SilentlyContinue
							$status = 'Removed'
							Start-Sleep -Milliseconds 200
						}
						else
						{
							$status = 'NoAction'
						}
						
						$Emlprop = [ordered]@{
							UPN	     = $UPN
							UserName = $UserRecipient.PrimarySmtpAddress.ToString()
							SamAccountName = $UserRecipient.SamAccountName
							MailBoxType = $UserRecipient.RecipientTypeDetails.ToString()
							EmailAddress = $Email.Address
							Status   = $Status
							Details  = 'Processed'
						}
						
						$obj = New-Object -TypeName System.Management.Automation.PSObject -Property $Emlprop
						Write-Output $obj
					}
				}
				else
				{
					$Emlprop = [ordered]@{
						UPN	     = $UPN
						UserName = $UserRecipient.PrimarySmtpAddress.ToString()
						SamAccountName = $UserRecipient.SamAccountName
						MailBoxType = $UserRecipient.RecipientTypeDetails.ToString()
						EmailAddress = 'None'
						Status   = 'ERROR'
						Details  = 'Already a cloud user'
					}
					$obj = New-Object -TypeName System.Management.Automation.PSObject -Property $Emlprop
					Write-Output $obj
				}
			}
			catch
			{
				$Emlprop = [ordered]@{
					UPN		       = $UPN
					UserName	   = 'None'
					SamAccountName = 'None'
					MailBoxType    = 'None'
					EmailAddress   = 'None'
					Status		   = 'ERROR'
					Details	       = $($_.Exception.Message)
				}
				
				$obj = New-Object -TypeName System.Management.Automation.PSObject -Property $Emlprop
				Write-Output $obj
			}
		}
	}
	end
	{
	}
}

function New-QHMoveRequestToO365
{
<#
	.SYNOPSIS
		A brief description of the New-QHMoveRequestToO365 function.
	
	.DESCRIPTION
		A detailed description of the New-QHMoveRequestToO365 function.
	
	.PARAMETER UserPrincipalName
		A description of the UserPrincipalName parameter.
	
	.PARAMETER OnPremCred
		A description of the OnPremCred parameter.
	
	.PARAMETER MrsEndPoint
		A description of the MrsEndPoint parameter.
	
	.PARAMETER TargetDeliveryDomain
		A description of the TargetDeliveryDomain parameter.
	
	.PARAMETER BatchName
		A description of the BatchName parameter.
	
	.EXAMPLE
		PS C:\> New-QHMoveRequestToO365 -UserPrincipalName 'value1'
	
	.NOTES
		Additional information about the function.
#>
	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory = $true,
				   ValueFromPipeline = $true,
				   ValueFromPipelineByPropertyName = $true,
				   HelpMessage = 'Please enter UPN or EmailAddress')]
		[ValidateNotNullOrEmpty()]
		[Alias('EmailAddress', 'PrimarySmtpAddress')]
		[String[]]$UserPrincipalName,
		[Parameter(Mandatory = $false)]
		[System.Management.Automation.Credential()]
		[ValidateNotNull()]
		[System.Management.Automation.PSCredential]$OnPremCred = [System.Management.Automation.PSCredential]::Empty,
		[String]$MrsEndPoint = 'mrs.health.qld.gov.au',
		[String]$TargetDeliveryDomain = 'healthqld.mail.onmicrosoft.com',
		[String]$BatchName = 'NotDefined'
	)
	
	begin
	{
		$i = 1
		
	}
	process
	{
		foreach ($UPN in $UserPrincipalName)
		{
			try
			{
				$recipient = Get-Recipient $UPN -ErrorAction Stop
				
				$whatifparameters = @{
					whatif					   = $true
					Identity				   = $UPN
					BatchName				   = $BatchName
					Remote					   = $true
					RemoteHostName			   = $MrsEndPoint
					RemoteCredential		   = $OnpremCred
					TargetDeliveryDomain	   = $TargetDeliveryDomain
					SuspendWhenReadyToComplete = $true
					BadItemLimit			   = 100
					LargeItemLimit			   = 100
					AcceptLargeDataLoss	       = $true
					ErrorAction			       = 'Stop'
					WarningAction			   = 'SilentlyContinue'
				}
				
				New-ExoMoveRequest @whatifparameters
				
				$prop = [ordered]@{
					User				 = $UPN
					RecipientTypeDetails = $recipient.RecipientTypeDetails
					UserNo			     = $i
					BatchName		     = $BatchName
					Status			     = 'SUCCESS'
					Details			     = 'None'
				}
			}
			catch
			{
				$prop = [ordered]@{
					User				 = $UPN
					RecipientTypeDetails = $recipient.RecipientTypeDetails
					UserNo			     = $i
					BatchName		     = $BatchName
					Status			     = 'FAILED'
					Details			     = "ERROR : $($_.Exception.Message)"
				}
			}
			finally
			{
				$obj = New-Object -TypeName System.Management.Automation.PSObject -Property $prop
				Write-Output $obj
				if ($UserPrincipalName.count -gt 1)
				{
					$paramWriteProgress = @{
						Activity = 'Testing WhatIf'
						Status   = "Processing [$i] of [$($UserPrincipalName.Count)] users"
						PercentComplete = (($i / $UserPrincipalName.Count) * 100)
						CurrentOperation = "Completed : [$UPN]"
					}
					Write-Progress @paramWriteProgress
				}
				$i++
			}
		}
	}
	end
	{
		Write-Progress -Activity 'Testing WhatIf' -Completed
	}
}

function Enable-QHLitigationHold
{
<#
	.SYNOPSIS
		A brief description of the Enable-QHLitigationHold function.
	
	.DESCRIPTION
		A detailed description of the Enable-QHLitigationHold function.
	
	.PARAMETER UserPrincipalName
		A description of the UserPrincipalName parameter.
	
	.PARAMETER BatchName
		A description of the BatchName parameter.
	
	.EXAMPLE
		PS C:\> Enable-QHLitigationHold -UserPrincipalName 'value1'
	
	.NOTES
		Additional information about the function.
#>
	param (
		[Parameter(Mandatory = $true,
				   ValueFromPipeline = $true,
				   ValueFromPipelineByPropertyName = $true,
				   HelpMessage = 'Please enter UPN or EmailAddress')]
		[ValidateNotNullOrEmpty()]
		[Alias('EmailAddress', 'PrimarySmtpAddress')]
		[String[]]$UserPrincipalName,
		[String]$BatchName
	)
	
	begin
	{
		$i = 1
		if ($BatchName -eq $null)
		{
			$BatchName = 'NotDefined'
		}
	}
	process
	{
		foreach ($UPN in $UserPrincipalName)
		{
			try
			{
				Set-ExoMailbox -identity $UPN -LitigationHoldEnabled $true -ErrorAction 'Stop'
				$prop = [ordered] @{
					User		   = $UPN
					BatchName	   = $BatchName
					LitigationHold = 'Success'
					Details	       = 'None'
				}
			}
			catch
			{
				$prop = [ordered] @{
					User		   = $UPN
					BatchName	   = $BatchName
					LitigationHold = 'Failed'
					Details	       = "ERROR : $($_.Exception.Message)"
				}
			}
			finally
			{
				$obj = New-Object -TypeName System.Management.Automation.PSObject -Property $prop
				Write-Output $obj
				
				if ($UserPrincipalName.count -gt 1)
				{
					$paramWriteProgress = @{
						Activity = 'Enabling Litigation Hold'
						Status   = "Processing [$i] of [$($UserPrincipalName.Count)] users"
						PercentComplete = (($i / $UserPrincipalName.Count) * 100)
						CurrentOperation = "Completed : [$UPN]"
					}
					Write-Progress @paramWriteProgress
					$i++
				}
			}
		}
	}
	end
	{
		Write-Progress -Activity 'Enabling Litigation Hold' -Completed
	}
}

function Enable-QHAuditingAndAddressPolicy
{
<#
	.SYNOPSIS
		A brief description of the Enable-QHAuditingAndAddressPolicy function.
	
	.DESCRIPTION
		A detailed description of the Enable-QHAuditingAndAddressPolicy function.
	
	.PARAMETER UserPrincipalName
		A description of the UserPrincipalName parameter.
	
	.PARAMETER BatchName
		A description of the BatchName parameter.
	
	.EXAMPLE
		PS C:\> Enable-QHAuditingAndAddressPolicy -UserPrincipalName 'value1'
	
	.NOTES
		Additional information about the function.
#>
	param (
		[Parameter(Mandatory = $true,
				   ValueFromPipeline = $true,
				   ValueFromPipelineByPropertyName = $true,
				   HelpMessage = 'Please enter UPN or EmailAddress')]
		[ValidateNotNullOrEmpty()]
		[Alias('EmailAddress', 'PrimarySmtpAddress')]
		[String[]]$UserPrincipalName,
		[String]$BatchName
	)
	
	begin
	{
		$i = 1
		if ($BatchName -eq $null)
		{
			$BatchName = 'NotDefined'
		}
	}
	process
	{
		foreach ($UPN in $UserPrincipalName)
		{
			try
			{
				$paramSetExoMailbox = @{
					identity				  = $UPN
					AuditEnabled			  = $true
					AuditLogAgeLimit		  = 2555
					AddressBookPolicy		  = "Main QH ABP"
					SingleItemRecoveryEnabled = $true
					RetainDeletedItemsfor	  = 30
					ErrorAction			      = 'Stop'
				}
				
				Set-ExoMailbox @paramSetExoMailbox
				
				$prop = [ordered] @{
					User					 = $UPN
					BatchName			     = $BatchName
					AuditingAndAddressPolicy = 'Success'
					Details				     = 'None'
				}
			}
			catch
			{
				$prop = [ordered] @{
					User					 = $UPN
					BatchName			     = $BatchName
					AuditingAndAddressPolicy = 'Failed'
					Details				     = "ERROR : $($_.Exception.Message)"
				}
			}
			finally
			{
				$obj = New-Object -TypeName System.Management.Automation.PSObject -Property $prop
				Write-Output $obj
				
				if ($UserPrincipalName.count -gt 1)
				{
					$paramWriteProgress = @{
						Activity = 'Setting Auditing and Address Policy'
						Status   = "Processing [$i] of [$($UserPrincipalName.Count)] users"
						PercentComplete = (($i / $UserPrincipalName.Count) * 100)
						CurrentOperation = "Completed : [$UPN]"
					}
					Write-Progress @paramWriteProgress
					$i++
				}
			}
		}
	}
	end
	{
		Write-Progress -Activity 'Setting Auditing and Address Policy' -Completed
	}
}

function Remove-QHEVGroupMemberShip
{
<#
	.SYNOPSIS
		A brief description of the Remove-QHEVGroupMemberShip function.
	
	.DESCRIPTION
		A detailed description of the Remove-QHEVGroupMemberShip function.
	
	.PARAMETER UserPrincipalName
		A description of the UserPrincipalName parameter.
	
	.PARAMETER BatchName
		A description of the BatchName parameter.
	
	.PARAMETER AddToStagingGrp
		A description of the AddToStagingGrp parameter.
	
	.EXAMPLE
		PS C:\> Remove-QHEVGroupMemberShip -UserPrincipalName 'value1'
	
	.NOTES
		Additional information about the function.
#>
	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory = $true,
				   ValueFromPipeline = $true,
				   ValueFromPipelineByPropertyName = $true)]
		[Alias('EmailAddress', 'PrimarySmtpAddress')]
		[String[]]$UserPrincipalName,
		[String]$BatchName = 'None',
		[Switch]$AddToStagingGrp
	)
	
	begin
	{
		$global:ErrorActionPreference = [System.Management.Automation.ActionPreference]::Stop
		$i = 1
		
	}
	process
	{
		foreach ($UPN in $UserPrincipalName)
		{
			try
			{
				$recipient = Get-Recipient $UPN -ErrorAction Stop
				[Array]$EVgroups = Get-AdUser $recipient.SamAccountName -Properties Memberof |
				Select-Object -ExpandProperty Memberof | Where-Object { $_ -like "CN=ACT-Outlook-Archiving*" }
				
				if ($EVgroups -ne $null)
				{
					foreach ($Group in $EVgroups)
					{
						$paramRemoveAdGroupMember = @{
							identity = $Group
							members  = $recipient.SamAccountName
							confirm  = $false
							ErrorAction = 'Stop'
						}
						
						Remove-AdGroupMember @paramRemoveAdGroupMember
					}
					$removedState = 'Removed'
					$GrpStr = $($EVgroups -join ',' | Out-String)
					
				}
				else
				{
					$removedState = 'NoAction'
					$GrpStr = 'No EV Group Membership'
				}
				
				if ($AddToStagingGrp)
				{
					$paramAddADGroupMember = @{
						Identity = 'Act-Outlook-Archiving-O365Staging'
						Members  = $recipient.SamAccountName
						ErrorAction = 'Stop'
					}
					
					Add-ADGroupMember @paramAddADGroupMember
					
					$staging = 'AddedTo:O365StagingGrp'
					
				}
				else
				{
					$staging = 'Skipped:NotSelected'
				}
				
				$prop = [ordered]@{
					EmailAddress   = $UPN
					BatchName	   = $BatchName
					SamAccountName = $recipient.SamAccountName
					Status		   = $removedState
					StagingGrp	   = $staging
					Groups		   = $GrpStr.Trim()
					Details	       = 'None'
				}
			}
			catch
			{
				$prop = [ordered]@{
					EmailAddress   = $UPN
					BatchName	   = $BatchName
					SamAccountName = 'Error'
					Status		   = 'Error'
					StagingGrp	   = 'Error'
					Groups		   = 'Error'
					Details	       = "ERROR : $($_.Exception.Message)"
				}
			}
			finally
			{
				$obj = New-Object -TypeName System.Management.Automation.PSObject -Property $prop
				Write-Output $obj
				
				if ($UserPrincipalName.count -gt 1)
				{
					$paramWriteProgress = @{
						Activity = 'Removing Users from EV Group'
						Status   = "Processing [$i] of [$($UserPrincipalName.Count)] users"
						PercentComplete = (($i / $UserPrincipalName.Count) * 100)
						CurrentOperation = "Completed : [$UPN]"
					}
					Write-Progress @paramWriteProgress
				}
				$i++
			}
		}
	}
	end
	{
		Write-Progress -Activity 'Removing Users from EV Group' -Completed
		$global:ErrorActionPreference = [System.Management.Automation.ActionPreference]::Continue
	}
}

function Set-QHMailboxQuota50GB
{
<#
	.SYNOPSIS
		A brief description of the Set-QHMailboxQuota50GB function.
	
	.DESCRIPTION
		A detailed description of the Set-QHMailboxQuota50GB function.
	
	.PARAMETER UserPrincipalName
		A description of the UserPrincipalName parameter.
	
	.PARAMETER DomainController
		A description of the DomainController parameter.
	
	.PARAMETER BatchName
		A description of the BatchName parameter.
	
	.EXAMPLE
		PS C:\> Set-QHMailboxQuota50GB -UserPrincipalName 'value1' -DomainController EAD-WDCBTPP01
	
	.NOTES
		Additional information about the function.
#>
	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory = $true,
				   ValueFromPipeline = $true,
				   ValueFromPipelineByPropertyName = $true)]
		[Alias('EmailAddress', 'PrimarySmtpAddress')]
		[String[]]$UserPrincipalName,
		[Parameter(Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		[ValidateSet('EAD-WDCBTPP01', 'EAD-WDCBK7P01', 'EAD-WDCBK7P03', 'EAD-WDCBK7P02', 'EAD-WDCBTPP02', 'EAD-WDCBTPP03', 'EAD-WDCBTPP04', 'EAD-WDCBK7P04', 'EAD-WDCBK7P05', 'EAD-WDCBTPP05', 'EAD-WDCBK7P06', 'EAD-WDCBTPP06', 'EAD-WDCBK7P07', 'EAD-WDCBTPP07', 'EAD-WDCBK7P08', 'EAD-WDCBTPP08')]
		[String]$DomainController,
		[String]$BatchName
	)
	
	begin
	{
		$global:ErrorActionPreference = [System.Management.Automation.ActionPreference]::Stop
		$i = 1
		
		if (!(Test-Connection -ComputerName $DomainController -Quiet))
		{
			Write-Warning "[$DomainController] is Unreachable. Plase Execute the command with a different Domain Controller"
			break
		}
		if ($BatchName = $null)
		{
			$BatchName = 'None'
		}
	}
	process
	{
		foreach ($UPN in $UserPrincipalName)
		{
			try
			{
				$Old = Get-Mailbox $UPN -ErrorAction Stop -WarningAction SilentlyContinue -DomainController $DomainController
				
				if (($Old.UseDatabaseQuotaDefaults -eq $false) -and ($Old.ProhibitSendQuota -gt 40GB))
				{
					$prop = [ordered]@{
						EmailAddress		 = $UPN
						OLD_QuotaInheritance = $Old.UseDatabaseQuotaDefaults
						NEW_QuotaInheritance = 'SKIPPED'
						OLD_ProhibitSendRecieveQuota = $Old.ProhibitSendReceiveQuota
						NEW_ProhibitSendRecieveQuota = 'SKIPPED'
						OLD_ProhibitSendQuota = $Old.ProhibitSendQuota
						New_ProhibitSendQuota = 'SKIPPED'
						OLD_IssueWarningQuota = $Old.IssueWarningQuota
						NEW_IssueWarningQuota = 'SKIPPED'
						Status			     = 'SKIPPED'
						Details			     = "SKIPPED : The Mailbox Already has Quotas set more than 40 GB. No Action required"
					}
				}
				elseif (($Old.UseDatabaseQuotaDefaults -eq $false) -and ($Old.ProhibitSendQuota -lt 40GB))
				{
					$paramSetMailbox = @{
						Identity				 = $UPN
						ErrorAction			     = 'Stop'
						WarningAction		     = 'SilentlyContinue'
						ProhibitSendQuota	     = 50GB
						ProhibitSendReceiveQuota = 50GB
						IssueWarningQuota	     = 49GB
						DomainController		 = $DomainController
					}
					
					Set-Mailbox @paramSetMailbox
					$New = Get-Mailbox $UPN -ErrorAction Stop -WarningAction SilentlyContinue -DomainController $DomainController
					$prop = [ordered]@{
						EmailAddress		 = $UPN
						BatchName		     = $BatchName
						OLD_QuotaInheritance = $Old.UseDatabaseQuotaDefaults
						NEW_QuotaInheritance = $New.UseDatabaseQuotaDefaults
						OLD_ProhibitSendRecieveQuota = $old.ProhibitSendReceiveQuota
						NEW_ProhibitSendRecieveQuota = $New.ProhibitSendReceiveQuota
						OLD_ProhibitSendQuota = $Old.ProhibitSendQuota
						New_ProhibitSendQuota = $New.ProhibitSendQuota
						OLD_IssueWarningQuota = $Old.IssueWarningQuota
						NEW_IssueWarningQuota = $New.IssueWarningQuota
						Status			     = 'PASSED'
						Details			     = 'None'
					}
				}
				elseif ($Old.UseDatabaseQuotaDefaults -eq $true)
				{
					$oldDB = Get-MailboxDatabase $Old.Database -ErrorAction Stop -DomainController $DomainController
					$paramSetMailbox = @{
						Identity				 = $UPN
						ErrorAction			     = 'Stop'
						WarningAction		     = 'SilentlyContinue'
						UseDatabaseQuotaDefaults = $false
						ProhibitSendQuota	     = 50GB
						ProhibitSendReceiveQuota = 50GB
						IssueWarningQuota	     = 49GB
						DomainController		 = $DomainController
					}
					
					Set-Mailbox @paramSetMailbox
					$New = Get-Mailbox $UPN -ErrorAction Stop -WarningAction SilentlyContinue -DomainController $DomainController
					$prop = [ordered]@{
						EmailAddress		 = $UPN
						BatchName		     = $BatchName
						OLD_QuotaInheritance = $Old.UseDatabaseQuotaDefaults
						NEW_QuotaInheritance = $New.UseDatabaseQuotaDefaults
						OLD_ProhibitSendRecieveQuota = $oldDB.ProhibitSendReceiveQuota
						NEW_ProhibitSendRecieveQuota = $New.ProhibitSendReceiveQuota
						OLD_ProhibitSendQuota = $OldDB.ProhibitSendQuota
						New_ProhibitSendQuota = $New.ProhibitSendQuota
						OLD_IssueWarningQuota = $OldDB.IssueWarningQuota
						NEW_IssueWarningQuota = $New.IssueWarningQuota
						Status			     = 'PASSED'
						Details			     = 'None'
					}
				}
			}
			catch
			{
				$prop = [ordered]@{
					EmailAddress				 = $UPN
					BatchName				     = $BatchName
					OLD_QuotaInheritance		 = 'ERROR'
					NEW_QuotaInheritance		 = 'ERROR'
					OLD_ProhibitSendRecieveQuota = 'ERROR'
					NEW_ProhibitSendRecieveQuota = 'ERROR'
					OLD_ProhibitSendQuota	     = 'ERROR'
					New_ProhibitSendQuota	     = 'ERROR'
					OLD_IssueWarningQuota	     = 'ERROR'
					NEW_IssueWarningQuota	     = 'ERROR'
					Status					     = 'FAILED'
					Details					     = "ERROR : $($_.Exception.Message)"
				}
			}
			finally
			{
				$obj = New-Object -TypeName System.Management.Automation.PSObject -Property $prop
				Write-Output $obj
				if ($UserPrincipalName.count -gt 1)
				{
					$paramWriteProgress = @{
						Activity = 'Setting Quotas for Users'
						Status   = "Processing [$i] of [$($UserPrincipalName.Count)] users"
						PercentComplete = (($i / $UserPrincipalName.Count) * 100)
						CurrentOperation = "Completed : [$UPN]"
					}
					Write-Progress @paramWriteProgress
				}
				$i++
			}
		}
	}
	end
	{
		Write-Progress -Activity 'Setting Quotas for Users' -Completed
		$global:ErrorActionPreference = [System.Management.Automation.ActionPreference]::Continue
	}
}

function Remove-QHE3ADGrpmembership
{
<#
	.SYNOPSIS
		A brief description of the Remove-QHE3ADGrpmembership function.
	
	.DESCRIPTION
		A detailed description of the Remove-QHE3ADGrpmembership function.
	
	.PARAMETER SamAccountName
		A description of the SamAccountName parameter.
	
	.PARAMETER AdGroup
		A description of the AdGroup parameter.
	
	.EXAMPLE
		PS C:\> Remove-QHE3ADGrpmembership
	
	.NOTES
		Additional information about the function.
#>
	[CmdletBinding()]
	param
	(
		$SamAccountName,
		[ValidateSet('ADM-O365-LIC-E3-OfficeProPlusOnly', 'ACT-SaaS-O365-GenericAccountOfficeProPlusOnly')]
		$AdGroup
	)
	
	try
	{
		Write-Verbose "Trying to remove $SamAccountName from $AdGroup"
		$paramRemoveAdGroupmember = @{
			Identity	  = $AdGroup
			members	      = $SamAccountName
			ErrorAction   = 'Stop'
			WarningAction = 'SilentlyContinue'
			Confirm	      = $False
		}
		
		Remove-AdGroupmember @paramRemoveAdGroupmember
		Write-Verbose "Successfully Removed"
		return "Removed:$AdGroup"
	}
	catch
	{
		Write-Verbose "Error : $($_.Exception.Message)"
		return "Failed:$AdGroup"
	}
}

function Add-QHGenericGroupMember
{
<#
	.SYNOPSIS
		A brief description of the Add-QHGenericGroupMember function.
	
	.DESCRIPTION
		A detailed description of the Add-QHGenericGroupMember function.
	
	.PARAMETER SamAccountName
		A description of the SamAccountName parameter.
	
	.PARAMETER AdGroup
		A description of the AdGroup parameter.
	
	.EXAMPLE
		PS C:\> Add-QHGenericGroupMember
	
	.NOTES
		Additional information about the function.
#>
	[CmdletBinding()]
	param
	(
		$SamAccountName,
		[ValidateSet('ACT-SaaS-O365-GenericAccountOfficeProPlusAndMailbox', 'ACT-SaaS-O365-GenericAccountMailboxOnly')]
		$AdGroup
	)
	
	try
	{
		Write-Verbose "Trying to remove $SamAccountName from $AdGroup"
		$paramRemoveAdGroupmember = @{
			Identity	  = $AdGroup
			members	      = $SamAccountName
			ErrorAction   = 'Stop'
			WarningAction = 'SilentlyContinue'
			Confirm	      = $False
		}
		
		Add-AdGroupMember @paramRemoveAdGroupmember
		Write-Verbose "Successfully Added"
		return "Added:$AdGroup"
	}
	catch
	{
		Write-Verbose "Error : $($_.Exception.Message)"
		return "FailedAdd:$AdGroup"
	}
}

function Find-QHAduser
{
<#
	.SYNOPSIS
		A brief description of the Find-QHAduser function.
	
	.DESCRIPTION
		A detailed description of the Find-QHAduser function.
	
	.PARAMETER DisplayName
		A description of the DisplayName parameter.
	
	.EXAMPLE
		PS C:\> Find-QHAduser
	
	.NOTES
		Additional information about the function.
#>
	[CmdletBinding()]
	param
	(
		$DisplayName
	)
	
	begin
	{
		$HHSContext = @{
			"200ADL.CO.STH.HEALTH"					    = "DoH"
			"51WEMBLEY.LOGAN.STH.HEALTH"			    = "Metro South"
			"61MARY.CO.STH.HEALTH"					    = "DoH"
			"ABIOS.BS.STH.HEALTH"					    = "Metro South"
			"ADH.TBL.FNQ.HEALTH"					    = "Cairns and Hinterland"
			"ALLIED.TPCH.Chermside.BNN.HEALTH"		    = "Metro North"
			"ARCHIVED.SVC.FNQ.HEALTH"				    = "Cairns and Hinterland"
			"Ashgrove.NTH-BNE.BNN.HEALTH"			    = "Metro North"
			"Ashworth-House.ON-BNE.BNN.HEALTH"		    = "Metro North"
			"Aspley.ON-BNE.BNN.HEALTH"				    = "Metro North"
			"Ayr.BWN.NTH.HEALTH"					    = "Townsville"
			"BAB.INN.FNQ.HEALTH"					    = "Cairns and Hinterland"
			"Bald-Hills.ON-BNE.BNN.HEALTH"			    = "Metro North"
			"Baralaba-H.BANANA.CTL.HEALTH"			    = "Central Queensland"
			"Barcaldine-H.CWEST.CTL.HEALTH"			    = "Central West"
			"BAY.FRASER.WBY.HEALTH"					    = "Wide Bay"
			"BBG.BUNDY.WBY.HEALTH"					    = "Wide Bay"
			"BCH.TORRES.FNQ.HEALTH"					    = "Torres and Cape"
			"BDH.TORRES.FNQ.HEALTH"					    = "Torres and Cape"
			"Beaudesert-H.LOGAN.STH.HEALTH"			    = "Metro South"
			"Beenleigh-CH.LOGAN.STH.HEALTH"			    = "Metro South"
			"BHHS.TWB.SWQ.HEALTH"					    = "Darling Downs"
			"Biala.IN-BNE.BNN.HEALTH"				    = "Metro North"
			"Biloela-H.BANANA.CTL.HEALTH"			    = "Central Queensland"
			"Blackall-H.CWEST.CTL.HEALTH"			    = "Central West"
			"Blackwater-H.CHIGH.CTL.HEALTH"			    = "Central Queensland"
			"Boonah.WM.SWQ.HEALTH"					    = "West Moreton"
			"Bowen.BWN.NTH.HEALTH"					    = "Mackay"
			"Brighton.ON-BNE.BNN.HEALTH"			    = "Metro North"
			"Browns-ACC.LOGAN.STH.HEALTH"			    = "Metro South"
			"BSC.TPCH.Chermside.BNN.HEALTH"			    = "Metro North"
			"BSQ.QEII.QEIIHD.STH.HEALTH"			    = "Metro South"
			"BSS.QHSS.BS.STH.HEALTH"				    = "Metro South"
			"CANNONH.BS.STH.HEALTH"					    = "Metro South"
			"CARD_THOR.TPCH.Chermside.BNN.HEALTH"	    = "Metro North"
			"CBH.CNS.FNQ.HEALTH"					    = "Cairns and Hinterland"
			"CFH.CO.STH.HEALTH"						    = "DoH"
			"CH.QEII.QEIIHD.STH.HEALTH"				    = "Metro South"
			"Char-H.CHAR.SWQ.HEALTH"				    = "South West"
			"CHE.SBUR.WBY.HEALTH"					    = "Wide Bay"
			"Chin-H.NDowns.SWQ.HEALTH"				    = "Darling Downs"
			"Chtwrs.CT.NTH.HEALTH"					    = "Townsville"
			"Clermont-H.MKY.NTH.HEALTH"				    = "Mackay"
			"CLIENTS.BeachRd.SCG.BNN.HEALTH"		    = "Sunshine Coast"
			"CLIENTS.Birtinya.SCG.BNN.HEALTH"		    = "Sunshine Coast"
			"CLIENTS.BottleBrush.SCG.BNN.HEALTH"	    = "Sunshine Coast"
			"CLIENTS.BrisbaneRd.SCG.BNN.HEALTH"		    = "Sunshine Coast"
			"CLIENTS.Caloundra.SCG.BNN.HEALTH"		    = "Sunshine Coast"
			"CLIENTS.Gympie.SCG.BNN.HEALTH"			    = "Sunshine Coast"
			"CLIENTS.HortonPde.SCG.BNN.HEALTH"		    = "Sunshine Coast"
			"Clients.Kilcoy-H.Red-Cab.BNN.HEALTH"	    = "Metro North"
			"CLIENTS.Maleny.SCG.BNN.HEALTH"			    = "Sunshine Coast"
			"CLIENTS.Musgrave.SCG.BNN.HEALTH"		    = "Sunshine Coast"
			"CLIENTS.Nambour.SCG.BNN.HEALTH"		    = "Sunshine Coast"
			"Clients.Red-Cab.BNN.HEALTH"			    = "Metro North"
			"CLIENTS.SCUH.SCG.BNN.HEALTH"			    = "Sunshine Coast"
			"CLIENTS.SixthAve.SCG.BNN.HEALTH"		    = "Sunshine Coast"
			"CLON-H.ISA.NTH.HEALTH"					    = "North West"
			"Collinsville.BWN.NTH.HEALTH"			    = "Mackay"
			"COMMUNITY.TPCH.Chermside.BNN.HEALTH"	    = "Metro North"
			"ComPlz.WM.SWQ.HEALTH"					    = "West Moreton"
			"COOKTOWN.CNS-REG.FNQ.HEALTH"			    = "Torres and Cape"
			"Coorparoo.QEIIHD.STH.HEALTH"			    = "Metro South"
			"Corinda.QEIIHD.STH.HEALTH"				    = "Metro South"
			"CORP.PAH.BS.STH.HEALTH"				    = "Metro South"
			"CORP.TPCH.Chermside.BNN.HEALTH"		    = "Metro North"
			"CSS.PAH.BS.STH.HEALTH"					    = "Metro South"
			"Dalby-H.NDowns.SWQ.HEALTH"				    = "Darling Downs"
			"Disabled users.MKY.NTH.HEALTH"			    = "Mackay"
			"DisabledUsers.ISA-H.ISA.NTH.HEALTH"	    = "North West"
			"DisabledUsers.MORN-H.ISA.NTH.HEALTH"	    = "North West"
			"DOOM-H.ISA.NTH.HEALTH"					    = "North West"
			"DTS.IFS.IS.HEALTH"						    = "DoH"
			"Dysart-H.MKY.NTH.HEALTH"				    = "Mackay"
			"EDMON.CNS-REG.FNQ.HEALTH"				    = "Cairns and Hinterland"
			"Emerald-H.CHIGH.CTL.HEALTH"			    = "Central Queensland"
			"Esk.WM.SWQ.HEALTH"						    = "Darling Downs"
			"Eventide.CT.NTH.HEALTH"				    = "Townsville"
			"Eventide-NH.ROCK.CTL.HEALTH"			    = "Central Queensland"
			"EXPIRED.GCH.GC.STH.HEALTH"				    = "Gold Coast"
			"EXPIRED.Helensvale.GC.STH.HEALTH"		    = "Gold Coast"
			"EXPIRED.SHP.GC.STH.HEALTH"				    = "Gold Coast"
			"EXTERNAL.CBH.CNS.FNQ.HEALTH"			    = "Cairns and Hinterland"
			"External.MKY.NTH.HEALTH"				    = "Mackay"
			"EXTERNAL.PAH.BS.STH.HEALTH"			    = "Metro South"
			"EXTERNAL.TGH.TSV.NTH.HEALTH"			    = "Townsville"
			"FS.QHSS.BS.STH.HEALTH"					    = "Metro South"
			"GARDC.BS.STH.HEALTH"					    = "Metro South"
			"GCH.GC.STH.HEALTH"						    = "Gold Coast"
			"GCTECH.GC.STH.HEALTH"					    = "Gold Coast"
			"GDH.NBUR.WBY.HEALTH"					    = "Wide Bay"
			"GHS.TWB.SWQ.HEALTH"					    = "Darling Downs"
			"Gladstone-H.ROCK.CTL.HEALTH"			    = "Central Queensland"
			"GMO.CO.STH.HEALTH"						    = "DoH"
			"Goodna.WM.SWQ.HEALTH"					    = "West Moreton"
			"Goondi-H.SDowns.SWQ.HEALTH"			    = "Darling Downs"
			"GORDON.CNS-REG.FNQ.HEALTH"				    = "Cairns and Hinterland"
			"Greenslopes.QEIIHD.STH.HEALTH"			    = "Metro South"
			"GWISE.SVC.DTS.IFS.IS.HEALTH"			    = "DoH"
			"HDH.TBL.FNQ.HEALTH"					    = "Cairns and Hinterland"
			"Helensvale.GC.STH.HEALTH"				    = "Gold Coast"
			"Homehill.BWN.NTH.HEALTH"				    = "Townsville"
			"Hughenden.CT.NTH.HEALTH"				    = "Townsville"
			"HVB.FRASER.WBY.HEALTH"					    = "Wide Bay"
			"ID.CO.STH.HEALTH"						    = "DoH"
			"IDD.PAH.BS.STH.HEALTH"					    = "Metro South"
			"IDH.INN.FNQ.HEALTH"					    = "Cairns and Hinterland"
			"IGH.WM.SWQ.HEALTH"						    = "West Moreton"
			"IMSU.TPCH.Chermside.BNN.HEALTH"		    = "Metro North"
			"Inactive.Blackwater-H.CHIGH.CTL.HEALTH"    = "Central Queensland"
			"Inactive.Emerald-H.CHIGH.CTL.HEALTH"	    = "Central Queensland"
			"Inactive.Gladstone-H.ROCK.CTL.HEALTH"	    = "Central Queensland"
			"Inactive.Rockhampton-CH.ROCK.CTL.HEALTH"   = "Central Queensland"
			"Inactive.Rockhampton-H.ROCK.CTL.HEALTH"    = "Central Queensland"
			"Inactive.Rockhampton-PH.ROCK.CTL.HEALTH"   = "Central Queensland"
			"Inactive.Springsure-H.CHIGH.CTL.HEALTH"    = "Central Queensland"
			"Inactive.Yeppoon-H.ROCK.CTL.HEALTH"	    = "Central Queensland"
			"Inactive-Users.TGH.TSV.NTH.HEALTH"		    = "Townsville"
			"InalaCYMH.QEIIHD.STH.HEALTH"			    = "Metro South"
			"Ingham.TSV.NTH.HEALTH"					    = "Townsville"
			"INGLE-H.SDowns.SWQ.HEALTH"				    = "Darling Downs"
			"IS.PAH.BS.STH.HEALTH"					    = "Metro South"
			"ISA-H.ISA.NTH.HEALTH"					    = "North West"
			"Jando-H.NDowns.SWQ.HEALTH"				    = "Darling Downs"
			"June2014.EXpUsers.User2.QHB.CO.STH.HEALTH" = "DoH"
			"Kingston-OH.LOGAN.STH.HEALTH"			    = "Metro South"
			"Kirwan.TSV.NTH.HEALTH"					    = "Townsville"
			"KRY.SBUR.WBY.HEALTH"					    = "Wide Bay"
			"Laidley.WM.SWQ.HEALTH"					    = "West Moreton"
			"LOGAN-H.LOGAN.STH.HEALTH"				    = "Metro South"
			"Longreach-CH.CWEST.CTL.HEALTH"			    = "Central West"
			"Longreach-DO.CWEST.CTL.HEALTH"			    = "Central West"
			"Longreach-H.CWEST.CTL.HEALTH"			    = "Central West"
			"Mackay-CH.MKY.NTH.HEALTH"				    = "Mackay"
			"Mackay-H.MKY.NTH.HEALTH"				    = "Mackay"
			"MAIN.InalaCH.QEIIHD.STH.HEALTH"		    = "Metro South"
			"Main.QEII.QEIIHD.STH.HEALTH"			    = "Metro South"
			"MBH.FRASER.WBY.HEALTH"					    = "Wide Bay"
			"MBNC.MBC.BAYSD.STH.HEALTH"				    = "Metro South"
			"MDH.TBL.FNQ.HEALTH"					    = "Cairns and Hinterland"
			"Meadow.LOGAN.STH.HEALTH"				    = "Metro South"
			"MED.PAH.BS.STH.HEALTH"					    = "Metro South"
			"MEDICAL.TPCH.Chermside.BNN.HEALTH"		    = "Metro North"
			"MENTAL.CNS.FNQ.HEALTH"					    = "Cairns and Hinterland"
			"MH.PAH.BS.STH.HEALTH"					    = "Metro South"
			"Miles-H.NDowns.SWQ.HEALTH"				    = "Darling Downs"
			"Mill-H.SDowns.SWQ.HEALTH"				    = "Darling Downs"
			"MLNH.TWB.SWQ.HEALTH"					    = "Darling Downs"
			"Moranbah-H.MKY.NTH.HEALTH"				    = "Mackay"
			"MORN-H.ISA.NTH.HEALTH"					    = "North West"
			"Mosman.CT.NTH.HEALTH"					    = "Townsville"
			"MOSSMAN.CNS-REG.FNQ.HEALTH"			    = "Cairns and Hinterland"
			"Moura-H.BANANA.CTL.HEALTH"				    = "Central Queensland"
			"MtMorgan-H.ROCK.CTL.HEALTH"			    = "Central Queensland"
			"Nathan.TSV.NTH.HEALTH"					    = "Townsville"
			"NerangDent.GCDental.GC.STH.HEALTH"		    = "Gold Coast"
			"NORM-H.ISA.NTH.HEALTH"					    = "North West"
			"NorthWard.TSV.NTH.HEALTH"				    = "Townsville"
			"NorthWest.NTH-BNE.BNN.HEALTH"			    = "Metro North"
			"Nth-Rton-NH.ROCK.CTL.HEALTH"			    = "Central Queensland"
			"NURSING.TPCH.Chermside.BNN.HEALTH"		    = "Metro North"
			"ODM.PAH.BS.STH.HEALTH"					    = "Metro South"
			"OHS.TWB.SWQ.HEALTH"					    = "Darling Downs"
			"Oral.QEII.QEIIHD.STH.HEALTH"			    = "Metro South"
			"PalmBeachCH.GC.STH.HEALTH"				    = "Gold Coast"
			"PalmIs.TSV.NTH.HEALTH"					    = "Townsville"
			"PATH.PAH.BS.STH.HEALTH"				    = "Metro South"
			"PCH.TPCH.Chermside.BNN.HEALTH"			    = "Metro North"
			"PHS.QHSS.BS.STH.HEALTH"				    = "Metro South"
			"PHU.PAH.BS.STH.HEALTH"					    = "Metro South"
			"PineRivers-CH.ON-BNE.BNN.HEALTH"		    = "Metro North"
			"Prime.GC.STH.HEALTH"					    = "Gold Coast"
			"Proserpine-H.MKY.NTH.HEALTH"			    = "Mackay"
			"PRT.CO.STH.HEALTH"						    = "DoH"
			"RAD.CO.STH.HEALTH"						    = "DoH"
			"RAD-ONC.BS.STH.HEALTH"					    = "Metro South"
			"RCFNQ.CBH.CNS.FNQ.HEALTH"				    = "Cairns and Hinterland"
			"Red-H.RHC.BAYSD.STH.HEALTH"			    = "Metro South"
			"REHAB.PAH.BS.STH.HEALTH"				    = "Metro South"
			"RHP.GC.STH.HEALTH"						    = "Gold Coast"
			"RICHLNDS.BS.STH.HEALTH"				    = "Metro South"
			"Richmond.CT.NTH.HEALTH"				    = "Townsville"
			"Robina.GC.STH.HEALTH"					    = "Gold Coast"
			"Rockhampton-CH.ROCK.CTL.HEALTH"		    = "Central Queensland"
			"Rockhampton-H.ROCK.CTL.HEALTH"			    = "Central Queensland"
			"Rockhampton-PH.ROCK.CTL.HEALTH"		    = "Central Queensland"
			"Roma-CH.Roma.SWQ.HEALTH"				    = "South West"
			"Roma-DWS.Roma.SWQ.HEALTH"				    = "South West"
			"Roma-H.Roma.SWQ.HEALTH"				    = "South West"
			"Roma-ORH.Roma.SWQ.HEALTH"				    = "Darling Downs"
			"Sarina-H.MKY.NTH.HEALTH"				    = "Mackay"
			"SMITH.CNS-REG.FNQ.HEALTH"				    = "Cairns and Hinterland"
			"Springsure-H.CHIGH.CTL.HEALTH"			    = "Central Queensland"
			"SSI.PAH.BS.STH.HEALTH"					    = "Metro South"
			"Stan-H.SDowns.SWQ.HEALTH"				    = "Darling Downs"
			"StGeorge-H.Roma.SWQ.HEALTH"			    = "South West"
			"SthBnePH.BS.STH.HEALTH"				    = "Metro South"
			"SURG.PAH.BS.STH.HEALTH"				    = "Metro South"
			"System.Object[]"						    = "System.Object[]"
			"Tara-H.NDowns.SWQ.HEALTH"				    = "Darling Downs"
			"Taroom-H.NDowns.SWQ.HEALTH"			    = "Darling Downs"
			"TESTOU.DTS.IFS.IS.HEALTH"				    = ""
			"Texas-H.SDowns.SWQ.HEALTH"				    = "Darling Downs"
			"TGH.TSV.NTH.HEALTH"					    = "Townsville"
			"Theodore-H.BANANA.CTL.HEALTH"			    = "Central Queensland"
			"THS.TWB.SWQ.HEALTH"					    = "Darling Downs"
			"TI-HOSP.TORRES.FNQ.HEALTH"				    = "Torres and Cape"
			"TI-PRIM.TORRES.FNQ.HEALTH"				    = "Torres and Cape"
			"TOP.BS.STH.HEALTH"						    = "Metro South"
			"TPCH.Chermside.BNN.HEALTH"				    = "Metro North"
			"TPHU.CNS.FNQ.HEALTH"					    = "Torres and Cape"
			"TULLY.INN.FNQ.HEALTH"					    = "Cairns and Hinterland"
			"USER.179GREY.CO.STH.HEALTH"			    = "DoH"
			"USER.199GREY.CO.STH.HEALTH"			    = "DoH"
			"USER.BDH.IN-BNE.BNS.HEALTH"			    = "Metro North"
			"USER.Carrara.GC.STH.HEALTH"			    = "Gold Coast"
			"USER.ChapelSt.ON-BNE.BNS.HEALTH"		    = "Metro South"
			"USER.Enoggera.ON-BNE.BNS.HEALTH"		    = "Metro South"
			"USER.Finney-Rd.ON-BNE.BNS.HEALTH"		    = "Metro North"
			"USER.GARU.ON-BNE.BNS.HEALTH"			    = "Metro South"
			"USER.Halwyn.ON-BNE.BNS.HEALTH"			    = "Metro South"
			"USER.Helensvale.GC.STH.HEALTH"			    = "Gold Coast"
			"USER.Herston.IN-BNE.BNS.HEALTH"		    = "Metro North"
			"USER.ID.CO.STH.HEALTH"					    = "DoH"
			"USER.LCCH.CHQ.STH.HEALTH"				    = "Children's Health Queensland"
			"USER.MarinePde.GC.STH.HEALTH"			    = "Gold Coast"
			"USER.NerangSSP.GC.STH.HEALTH"			    = "Gold Coast"
			"USER.Nundah.ON-BNE.BNS.HEALTH"			    = "Metro South"
			"USER.Nundah-CH.ON-BNE.BNS.HEALTH"		    = "Metro South"
			"USER.PalmBeachCH.GC.STH.HEALTH"		    = "Gold Coast"
			"USER.QEII.QEIIHD.STH.HEALTH"			    = "Metro South"
			"USER.Robina.GC.STH.HEALTH"				    = "Gold Coast"
			"USER.SHP.GC.STH.HEALTH"				    = "Gold Coast"
			"USER.Stafford.ON-BNE.BNS.HEALTH"		    = "Metro North"
			"User1.QHB.CO.STH.HEALTH"				    = "DoH"
			"User2.QHB.CO.STH.HEALTH"				    = "DoH"
			"USERS.LOGAN-H.LOGAN.STH.HEALTH"		    = "Metro South"
			"VIL.FRASER.WBY.HEALTH"					    = "Wide Bay"
			"Vincent.TSV.NTH.HEALTH"				    = "Townsville"
			"Warehouse.TSV.NTH.HEALTH"				    = "Townsville"
			"WEB-EXT.CO.STH.HEALTH"					    = "DoH"
			"Whitsunday-CH.MKY.NTH.HEALTH"			    = "Mackay"
			"WHS.SDowns.SWQ.HEALTH"					    = "Darling Downs"
			"Winton-H.CWEST.CTL.HEALTH"				    = "Central West"
			"Woodridge-CH.LOGAN.STH.HEALTH"			    = "Metro South"
			"Woorabinda-H.ROCK.CTL.HEALTH"			    = "Central Queensland"
			"WPA-HOSP.CAPE.FNQ.HEALTH"				    = "Torres and Cape"
			"WPH.WM.SWQ.HEALTH"						    = "West Moreton"
			"WUJAL.CNS-REG.FNQ.HEALTH"				    = "Torres and Cape"
			"WYN-H.WHC.BAYSD.STH.HEALTH"			    = "Metro South"
			"YBH-HOSP.CNS-REG.FNQ.HEALTH"			    = "Cairns and Hinterland"
			"Yeppoon-H.ROCK.CTL.HEALTH"				    = "Central Queensland"
			"YerongaCYMH.QEIIHD.STH.HEALTH"			    = "Metro South"
		}
	}
	process
	{
		$AD = Get-ADUser -Filter { Displayname -like $user } -Properties *
		
		if ($AD -ne $Null)
		{
			if ($AD.Count -gt 1)
			{
				$i = 1
				foreach ($item in $ad)
				{
					$HHS = $HHSContext["$($item.AUQHLANCX)"]
					
					if ($HHS -eq $null)
					{
						$HHS = $($item.AUQHLANCX)
					}
					$prop = [ordered]@{
						DisplayName = $DisplayName
						HHS		    = $HHS
						SamAccount  = $item.SamAccountName
						EmailAddress = $item.Mail
						UPN		    = $item.UserPrincipalName
						Enabled	    = $item.Enabled
						LastLogon   = $item.LastLogonDate
						type	    = $item.ObjectClass
						MatchCount  = $i
						MultipleMatch = 'True'
						Details	    = 'None'
					}
					$obj = New-Object -TypeName System.Management.Automation.PSObject -Property $Prop
					Write-Output $obj
					$i++
				}
			}
			else
			{
				$HHS = $HHSContext["$($AD.AUQHLANCX)"]
				
				if ($HHS -eq $null)
				{
					$HHS = $($AD.AUQHLANCX)
				}
				$prop = [ordered]@{
					DisplayName = $DisplayName
					HHS		    = $HHS
					SamAccount  = $AD.SamAccountName
					EmailAddress = $AD.Mail
					UPN		    = $AD.UserPrincipalName
					Enabled	    = $AD.Enabled
					LastLogon   = $AD.LastLogonDate
					type	    = $AD.ObjectClass
					MatchCount  = 1
					MultipleMatch = 'False'
					Details	    = 'None'
				}
				$obj = New-Object -TypeName System.Management.Automation.PSObject -Property $Prop
				Write-Output $obj
			}
		}
		else
		{
			$prop = [ordered]@{
				DisplayName   = $DisplayName
				HHS		      = 'None'
				SamAccount    = 'None'
				EmailAddress  = 'None'
				UPN		      = 'None'
				Enabled	      = 'None'
				LastLogon	  = 'None'
				type		  = 'None'
				MatchCount    = 'None'
				MultipleMatch = 'None'
				Details	      = "Couuld not Find the user with Display Name [$DisplayName]"
			}
			$obj = New-Object -TypeName System.Management.Automation.PSObject -Property $Prop
			Write-Output $obj
		}
	}
	end
	{
		
	}
}

function Disable-QHIMapandPop
{
<#
	.SYNOPSIS
		A brief description of the Disable-QHIMapandPop function.
	
	.DESCRIPTION
		A detailed description of the Disable-QHIMapandPop function.
	
	.PARAMETER UserPrincipalName
		A description of the UserPrincipalName parameter.
	
	.PARAMETER BatchName
		A description of the BatchName parameter.
	
	.EXAMPLE
		PS C:\> Disable-QHIMapandPop
	
	.NOTES
		Additional information about the function.
#>
	[CmdletBinding()]
	param
	(
		[String[]]$UserPrincipalName,
		[String]$BatchName
	)
	begin
	{
		if (!($BatchName))
		{
			$BatchName = 'None'
		}
		$i = 1
	}
	process
	{
		foreach ($UPN in $UserPrincipalName)
		{
			try
			{
				Set-ExoCasMailbox -identity $UPN -PopEnabled $false -ImapEnabled $false -ErrorAction Stop
				$prop = [ordered]@{
					EmailAddress	  = $UPN
					BatchName		  = $BatchName
					DisableIMAPandPOP = 'Success'
					Details		      = 'None'
				}
			}
			catch
			{
				$prop = [ordered]@{
					EmailAddress	  = $UPN
					BatchName		  = $BatchName
					DisableIMAPandPOP = 'Failed'
					Details		      = "ERROR : $($_.Exception.Message)"
				}
			}
			finally
			{
				$obj = New-Object -TypeName System.Management.Automation.PSObject -Property $prop
				Write-Output $obj
				if ($UserPrincipalName.count -gt 1)
				{
					$paramWriteProgress = @{
						Activity = 'Disabling IMAP and POP for Users'
						Status   = "Processing [$i] of [$($UserPrincipalName.Count)] users"
						PercentComplete = (($i / $UserPrincipalName.Count) * 100)
						CurrentOperation = "Completed : [$UPN]"
					}
					Write-Progress @paramWriteProgress
				}
				$i++
			}
		}
	}
	end
	{
		Write-Progress -Activity 'Disabling IMAP and POP for Users' -Completed
	}
}

function Navigate-QHMigrationFolder
{
<#
	.SYNOPSIS
		A brief description of the Navigate-QHMigrationFolder function.
	
	.DESCRIPTION
		A detailed description of the Navigate-QHMigrationFolder function.
	
	.PARAMETER BatchName
		A description of the BatchName parameter.
	
	.PARAMETER ParentDir
		A description of the ParentDir parameter.
	
	.EXAMPLE
		PS C:\> Navigate-QHMigrationFolder
	
	.NOTES
		Additional information about the function.
#>
	param (
		$BatchName,
		$ParentDir = "D:\Office365\Migrations\Batch"
	)
	try
	{
		$location = "$ParentDir\$BatchName"
		Set-Location $location -ErrorAction Stop
		return $true
	}
	catch
	{
		Write-host "ERROR : Please make sure the directory [$location] Exists" -ForegroundColor Magenta
		return $false
	}
}

function Add-QHRoutngAddress # need to only check the domain for on microsoft
{
<#
	.SYNOPSIS
		A brief description of the Add-QHRoutngAddress function.
	
	.DESCRIPTION
		A detailed description of the Add-QHRoutngAddress function.
	
	.PARAMETER UserPrincipalName
		A description of the UserPrincipalName parameter.
	
	.PARAMETER BatchName
		A description of the BatchName parameter.
	
	.EXAMPLE
		PS C:\> Add-QHRoutngAddress
	
	.NOTES
		Additional information about the function.
#>
	[CmdletBinding()]
	param
	(
		[String[]]$UserPrincipalName,
		[String]$BatchName = 'None'
	)
	
	begin
	{
		$i = 1
	}
	process
	{
		foreach ($UPN in $UserPrincipalName)
		{
			try
			{
				$recipient = Get-Recipient $UPN -ErrorAction Stop
				if ($recipient.RecipientTypeDetails -notmatch 'remote')
				{
					$RoutingAddress = $recipient.PrimarySmtpAddress.ToString().Split('@')[0] + '@healthqld.mail.onmicrosoft.com'
					
					[Array]$smtpaddresses = $recipient.EmailAddresses.Smtpaddress
					
					if ($smtpaddresses -inotcontains $RoutingAddress)
					{
						#add routingAddress
						Set-Mailbox $recipient.PrimarySmtpAddress.ToString() -EmailAddresses @{ Add = $RoutingAddress } -ErrorAction Stop -WarningAction SilentlyContinue
						
						$prop = [Ordered]@{
							EmailAddress		 = $UPN
							BatchName		     = $BatchName
							RecipientTypeDetails = $recipient.RecipientTypeDetails
							Added			     = $RoutingAddress
							Details			     = "Added: $RoutingAddress"
						}
					}
					else
					{
						$prop = [Ordered]@{
							EmailAddress		 = $UPN
							BatchName		     = $BatchName
							RecipientTypeDetails = $recipient.RecipientTypeDetails
							Added			     = 'Skipped'
							Details			     = "Routing Address Already Exists. No Action needed"
						}
					}
				}
				else
				{
					$prop = [Ordered]@{
						EmailAddress		 = $UPN
						BatchName		     = $BatchName
						RecipientTypeDetails = $recipient.RecipientTypeDetails
						Added			     = 'Skipped'
						Details			     = "Already a Cloud User"
					}
				}
			}
			catch
			{
				$prop = [Ordered]@{
					EmailAddress		 = $UPN
					BatchName		     = $BatchName
					RecipientTypeDetails = 'ERROR'
					Added			     = 'ERROR'
					Details			     = "ERROR : $($_.Exception.Message)"
				}
			}
			finally
			{
				$obj = New-Object -TypeName System.Management.Automation.PSObject -Property $prop
				Write-Output $obj
				if ($UserPrincipalName.count -gt 1)
				{
					$paramWriteProgress = @{
						Activity = 'Remediating Routing Address'
						Status   = "Processing [$i] of [$($UserPrincipalName.Count)] users"
						PercentComplete = (($i / $UserPrincipalName.Count) * 100)
						CurrentOperation = "Completed : [$UPN]"
					}
					Write-Progress @paramWriteProgress
				}
				$i++
			}
		}
	}
	end
	{
		Write-Progress -Activity 'Remediating Routing Address' -Completed
	}
}

function Add-zQHRoutngAddress
{
<#
	.SYNOPSIS
		A brief description of the Add-zQHRoutngAddress function.
	
	.DESCRIPTION
		A detailed description of the Add-zQHRoutngAddress function.
	
	.PARAMETER UserPrincipalName
		A description of the UserPrincipalName parameter.
	
	.PARAMETER BatchName
		A description of the BatchName parameter.
	
	.EXAMPLE
		PS C:\> Add-zQHRoutngAddress
	
	.NOTES
		Additional information about the function.
#>
	[CmdletBinding()]
	param
	(
		[String[]]$UserPrincipalName,
		[String]$BatchName = 'None'
	)
	
	begin
	{
		$global:ErrorActionPreference = [System.Management.Automation.ActionPreference]::Stop
		$i = 1
		$RoutingDomain = 'healthqld.mail.onmicrosoft.com'
		Write-Verbose "Start Process"
	}
	process
	{
		foreach ($UPN in $UserPrincipalName)
		{
			Write-Verbose "Trying to get recipient $UPN"
			try
			{
				$recipient = Get-Recipient $UPN -ErrorAction Stop
				Write-Verbose "Checking if the $UPN is a Cloud User"
				if ($recipient.RecipientTypeDetails -notmatch 'remote')
				{
					$EmailaddressColl = @()
					
					$RoutingAddress = $recipient.PrimarySmtpAddress.ToString().Split('@')[0] + '@' + $RoutingDomain
					
					foreach ($email in $recipient.EmailAddresses.SmtpAddress)
					{
						Write-Verbose "Adding $($email.ToString()) to check list"
						[mailaddress]$email = $email
						$EmailaddressColl += $email
						
					}
					
					if ($EmailaddressColl.host -inotcontains $RoutingDomain)
					{
						#add routingAddress
						Write-Verbose "Trying to Add $RoutingAddress to EmailAddresses"
						Set-Mailbox $recipient.PrimarySmtpAddress.ToString() -EmailAddresses @{ Add = $RoutingAddress } -ErrorAction Stop -WarningAction SilentlyContinue
						
						$prop = [Ordered]@{
							EmailAddress		 = $UPN
							BatchName		     = $BatchName
							RecipientTypeDetails = $recipient.RecipientTypeDetails
							Action			     = 'Added'
							Details			     = "Added: $RoutingAddress to $UPN"
						}
					}
					else
					{
						Write-Verbose "No action needed"
						$prop = [Ordered]@{
							EmailAddress		 = $UPN
							BatchName		     = $BatchName
							RecipientTypeDetails = $recipient.RecipientTypeDetails
							Action			     = 'AlreadyExists'
							Details			     = "Routing Address Already Exists. No Action needed"
						}
					}
				}
				else
				{
					Write-Verbose "Already a cloud user, no action needed"
					$prop = [Ordered]@{
						EmailAddress		 = $UPN
						BatchName		     = $BatchName
						RecipientTypeDetails = $recipient.RecipientTypeDetails
						Action			     = 'Skipped'
						Details			     = "Already a Cloud User"
					}
				}
			}
			catch
			{
				Write-Verbose "Error Occured, Reporting the error"
				$prop = [Ordered]@{
					EmailAddress		 = $UPN
					BatchName		     = $BatchName
					RecipientTypeDetails = 'ERROR'
					Action			     = 'ERROR'
					Details			     = "ERROR : $($_.Exception.Message)"
				}
			}
			finally
			{
				Write-Verbose "Trying to generate object"
				$obj = New-Object -TypeName System.Management.Automation.PSObject -Property $prop
				Write-Output $obj
				if ($UserPrincipalName.count -gt 1)
				{
					$paramWriteProgress = @{
						Activity = 'Remediating Routing Address'
						Status   = "Processing [$i] of [$($UserPrincipalName.Count)] users"
						PercentComplete = (($i / $UserPrincipalName.Count) * 100)
						CurrentOperation = "Completed : [$UPN]"
					}
					Write-Progress @paramWriteProgress
				}
				Write-Verbose "Completed Process for $UPN"
				$i++
			}
		}
	}
	end
	{
		Write-Progress -Activity 'Remediating Routing Address' -Completed
		Write-Verbose "End Process"
		$global:ErrorActionPreference = [System.Management.Automation.ActionPreference]::Continue
	}
}

<#function Add-zQHRoutngAddress
{
	[CmdletBinding()]
	param
	(
		[String[]]$UserPrincipalName,
		[String]$BatchName = 'None'
	)
	
	BEGIN
	{
		$global:ErrorActionPreference = [System.Management.Automation.ActionPreference]::Stop
		$i = 1
		$RoutingDomain = 'healthqld.mail.onmicrosoft.com'
		Write-Verbose "Start Process"
	}
	PROCESS
	{
		foreach ($UPN in $UserPrincipalName)
		{
			Write-Verbose "Trying to get recipient $UPN"
			Try
			{
				$recipient = Get-Recipient $UPN -ErrorAction Stop
				Write-Verbose "Checking if the $UPN is a Cloud User"
				if ($recipient.RecipientTypeDetails -notmatch 'remote')
				{
					$EmailaddressColl = @()
					
					$RoutingAddress = $recipient.PrimarySmtpAddress.ToString().Split('@')[0] + '@' + $RoutingDomain
					
					$SmtpEmailAddresses = $($recipient.EmailAddresses | Where-Object { $_.ToLower().StartsWith("smtp", "CurrentCultureIgnoreCase") })
					
					foreach ($email in $SmtpEmailAddresses)
					{
						Write-Verbose "Adding $($email.ToString()) to check list"
						
						[mailaddress]$xemail = $email.ToLower() -replace "smtp:", ""
						$EmailaddressColl += $xemail
						
					}
					
					if ($EmailaddressColl.host -inotcontains $RoutingDomain)
					{
						#add routingAddress
						Write-Verbose "Trying to Add $RoutingAddress to EmailAddresses"
						Set-Mailbox $recipient.PrimarySmtpAddress.ToString() -EmailAddresses @{ Add = $RoutingAddress } -ErrorAction Stop -WarningAction SilentlyContinue
						
						$prop = [Ordered]@{
							EmailAddress = $UPN
							BatchName = $BatchName
							RecipientTypeDetails = $recipient.RecipientTypeDetails
							Action = "Added"
							Details = "Added $RoutingAddress to $UPN"
						}
					}
					Else
					{
						Write-Verbose "No action needed"
						$prop = [Ordered]@{
							EmailAddress = $UPN
							BatchName = $BatchName
							RecipientTypeDetails = $recipient.RecipientTypeDetails
							Action = 'AlreadyExists'
							Details = "Routing Address Already Exists. No Action needed"
						}
					}
				}
				Else
				{
					Write-Verbose "Already a cloud user, no action needed"
					$prop = [Ordered]@{
						EmailAddress = $UPN
						BatchName = $BatchName
						RecipientTypeDetails = $recipient.RecipientTypeDetails
						Action = 'Skipped'
						Details = "Already a Cloud User"
					}
				}
			}
			Catch
			{
				Write-Verbose "Error Occured, Reporting the error"
				$prop = [Ordered]@{
					EmailAddress = $UPN
					BatchName = $BatchName
					RecipientTypeDetails = 'ERROR'
					Action = 'ERROR'
					Details = "ERROR : $($_.Exception.Message)"
				}
			}
			Finally
			{
				Write-Verbose "Trying to generate object"
				$obj = New-Object -TypeName System.Management.Automation.PSObject -Property $prop
				Write-Output $obj
				if ($UserPrincipalName.count -gt 1)
				{
					$paramWriteProgress = @{
						Activity = 'Remediating Routing Address'
						Status = "Processing [$i] of [$($UserPrincipalName.Count)] users"
						PercentComplete = (($i / $UserPrincipalName.Count) * 100)
						CurrentOperation = "Completed : [$UPN]"
					}
					Write-Progress @paramWriteProgress
				}
				Write-Verbose "Completed Process for $UPN"
				$i++
			}
		}
	}
	END
	{
		Write-Progress -Activity 'Remediating Routing Address' -Completed
		Write-Verbose "End Process"
		$global:ErrorActionPreference = [System.Management.Automation.ActionPreference]::Continue
	}
}#>

function Set-xSkypeSettings # recursive function for skype settings due to Flaky Skype online sessions 
{
<#
	.SYNOPSIS
		A brief description of the Set-xSkypeSettings function.
	
	.DESCRIPTION
		A detailed description of the Set-xSkypeSettings function.
	
	.PARAMETER user
		A description of the user parameter.
	
	.PARAMETER counter
		A description of the counter parameter.
	
	.PARAMETER Retry
		A description of the Retry parameter.
	
	.PARAMETER BatchName
		A description of the BatchName parameter.
	
	.PARAMETER inputObject
		A description of the inputObject parameter.
	
	.PARAMETER Cred
		A description of the Cred parameter.
	
	.EXAMPLE
		PS C:\> Set-xSkypeSettings
	
	.NOTES
		Additional information about the function.
#>
	[CmdletBinding()]
	param
	(
		[string]$user,
		[int]$counter = 0,
		[int]$Retry = 3,
		[String]$BatchName,
		[object]$inputObject = $null,
		[System.Management.Automation.Credential()]
		[ValidateNotNull()]
		[System.Management.Automation.PSCredential]$Cred = [System.Management.Automation.PSCredential]::Empty
	)
	
	if ($counter -lt $Retry)
	{
		try
		{
			Grant-365CsConferencingPolicy -Identity $user -PolicyName "BposSAllModalityMinVideoBW" -ErrorAction Stop -WarningAction SilentlyContinue
			Grant-365CsExternalAccessPolicy -Identity $user -PolicyName "FederationOnly" -ErrorAction Stop -WarningAction SilentlyContinue
			
			$counter++
			
			$prop = [ordered]@{
				UserPrincipalName = $user
				BatchName		  = $BatchName
				Attempt		      = $counter
				status		      = 'Success'
				Details		      = 'Applied : [BposSAllModalityMinVideoBW] and [FederationOnly]'
			}
			
			$obj = New-Object -TypeName psobject -Property $prop
			Write-Output $obj
		}
		catch
		{
			$counter++
			$prop = [ordered]@{
				UserPrincipalName = $user
				BatchName		  = $BatchName
				Attempt		      = $counter
				status		      = 'Failed'
				Details		      = "ERROR : $($_.Exception.Message)"
			}
			
			$obj = New-Object -TypeName psobject -Property $prop
			
			Set-Location D:\Office365\Rb
			Get-PSSession | Where-Object {
				$_.ComputerName -match 'adminau1'
			} | Remove-PSSession -ErrorAction SilentlyContinue
			Start-Sleep -Milliseconds 500
			Connect-QhSkypeOnline -Credential $Cred
			
			#Write-Output $obj
			#Write-host "Tried $counter" -ForegroundColor Cyan
			
			$paramSetxSkypeSettings = @{
				user	    = $user
				counter	    = $counter
				Retry	    = $Retry
				BatchName   = $BatchName
				inputObject = $obj
				Cred	    = $Cred
			}
			
			Set-xSkypeSettings @paramSetxSkypeSettings
			#get-blah -user $user  $counter -count $count -input $obj
		}
	}
	else
	{
		#Write-Host "Failed Despite $Retry Attempts" -ForegroundColor Red
		Write-Output $inputObject
		
	}
}

function Set-QHSkypeSettings
{
<#
	.SYNOPSIS
		A brief description of the Set-QHSkypeSettings function.
	
	.DESCRIPTION
		A detailed description of the Set-QHSkypeSettings function.
	
	.PARAMETER BatchName
		A description of the BatchName parameter.
	
	.PARAMETER UserPrincipalName
		A description of the UserPrincipalName parameter.
	
	.PARAMETER Cred
		A description of the Cred parameter.
	
	.PARAMETER ShowProgress
		A description of the ShowProgress parameter.
	
	.EXAMPLE
		PS C:\> Set-QHSkypeSettings
	
	.NOTES
		Additional information about the function.
#>
	param (
		[String]$BatchName,
		[String[]]$UserPrincipalName,
		[System.Management.Automation.Credential()]
		[ValidateNotNull()]
		[System.Management.Automation.PSCredential]$Cred = [System.Management.Automation.PSCredential]::Empty,
		[Switch]$ShowProgress
	)
	begin
	{
		$i = 1
	}
	process
	{
		foreach ($UPN in $UserPrincipalName)
		{
			try
			{
				$recipient = Get-Recipient $UPN -ErrorAction Stop
				if ($recipient.RecipientTypeDetails -match 'UserMailbox')
				{
					$obj = Set-xSkypeSettings -user $UPN -Retry 2 -BatchName $BatchName -Cred $Cred
				}
				else
				{
					$prop = [ordered]@{
						UserPrincipalName = $UPN
						BatchName		  = $BatchName
						Attempt		      = 'None'
						status		      = 'Skipped'
						Details		      = "$UPN in not a user mailbox. The mailbox type is $($recipient.RecipientTypeDetails)"
					}
					$obj = New-Object -TypeName System.Management.Automation.PSObject -Property $prop
				}
			}
			catch
			{
				$prop = [ordered]@{
					UserPrincipalName = $UPN
					BatchName		  = $BatchName
					Attempt		      = 'ERROR'
					status		      = 'ERROR'
					Details		      = "ERROR : $($_.Exception.Message)"
				}
				$obj = New-Object -TypeName System.Management.Automation.PSObject -Property $prop
			}
			finally
			{
				Write-Output $obj
				if ($UserPrincipalName.count -gt 1)
				{
					if ($ShowProgress)
					{
						$paramWriteProgress = @{
							Activity = 'Applying Skype Settigs for External Access and Conferencing'
							Status   = "Processing [$i] of [$($UserPrincipalName.Count)] users"
							PercentComplete = (($i / $UserPrincipalName.Count) * 100)
							CurrentOperation = "Completed : [$UPN]"
						}
						Write-Progress @paramWriteProgress
					}
				}
				$i++
			}
		}
	}
	end
	{
		Write-Progress -Activity 'Applying Skype Settigs for External Access and Conferencing' -Completed
	}
}

function Validate-QHADSipAddress
{
<#
	.SYNOPSIS
		A brief description of the Validate-QHADSipAddress function.
	
	.DESCRIPTION
		A detailed description of the Validate-QHADSipAddress function.
	
	.PARAMETER UserPrincipalName
		A description of the UserPrincipalName parameter.
	
	.PARAMETER Remediate
		A description of the Remediate parameter.
	
	.PARAMETER DomainController
		A description of the DomainController parameter.
	
	.PARAMETER BatchName
		A description of the BatchName parameter.
	
	.EXAMPLE
		PS C:\> Validate-QHADSipAddress -DomainController EAD-WDCBTPP01
	
	.NOTES
		Additional information about the function.
#>
	[CmdletBinding()]
	param
	(
		[String[]]$UserPrincipalName,
		[Switch]$Remediate,
		[Parameter(Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		[ValidateSet('EAD-WDCBTPP01', 'EAD-WDCBK7P01', 'EAD-WDCBK7P03', 'EAD-WDCBK7P02', 'EAD-WDCBTPP02', 'EAD-WDCBTPP03', 'EAD-WDCBTPP04', 'EAD-WDCBK7P04', 'EAD-WDCBK7P05', 'EAD-WDCBTPP05', 'EAD-WDCBK7P06', 'EAD-WDCBTPP06', 'EAD-WDCBK7P07', 'EAD-WDCBTPP07', 'EAD-WDCBK7P08', 'EAD-WDCBTPP08')]
		[String]$DomainController,
		[String]$BatchName
	)
	begin
	{
		$global:ErrorActionPreference = [System.Management.Automation.ActionPreference]::Stop
		$i = 1
		
		if (!(Test-Connection -ComputerName $DomainController -Quiet))
		{
			Write-Warning "[$DomainController] is Unreachable. Plase Execute the command with a different Domain Controller"
			break
		}
	}
	process
	{
		foreach ($UPN in $UserPrincipalName)
		{
			try
			{
				$recipient = Get-Recipient $UPN -DomainController $DomainController -ErrorAction Stop
				if ($recipient.RecipientTypeDetails -match 'UserMailbox')
				{
					$adinfo = Get-ADUser -Server $DomainController -Filter { UserPrincipalName -eq $UPN } -Properties * -ErrorAction Stop
					
					$AdProxyAddresses = $adinfo | Select-Object -ExpandProperty ProxyAddresses
					
					if ($AdProxyAddresses -ne $null)
					{
						foreach ($address in $AdProxyAddresses)
						{
							if ($address.ToLower().Startswith('sip', "CurrentCultureIgnoreCase"))
							{
								$rawSip = $address
								$sip = $address.ToLower() -replace 'sip:', ''
								break
							}
							else
							{
								$rawsip = $null
								$sip = $null
							}
						}
						if ($sip -ne $null)
						{
							$SipHostingProvider = try { Get-Csuser $UPN -EA Stop | Select-Object -ExpandProperty HostingProvider }
							catch { 'FailedToRetrive' }
							
							if ($sip -eq $UPN)
							{
								$sipCheck = 'Passed'
								$action = 'NOTNEEDED'
								$detail = "INFO : The SIP [$rawSip] matched [$UPN]"
							}
							else
							{
								if ($Remediate)
								{
									try
									{
										Set-CsUser -Identity $UPN -SipAddress "SIP:$UPN" -ErrorAction Stop
										$action = 'REMEDIATED'
										$detail = "REMEDIATED : The SIP [$rawSip] did not matched [$UPN]. This is Remediated"
									}
									catch
									{
										$action = 'FAILED'
										$detail = "FAILED : The SIP [$rawSip] did not matched [$UPN].Failed to Remediate with error $($_.Exception.message)"
									}
								}
								else
								{
									$action = 'REPORTONLY'
									$detail = "WARNING : The SIP [$rawSip] did not matched [$UPN] No Action taken"
								}
								$sipCheck = 'MisMatchSip'
							}
						}
						else
						{
							if ($Remediate)
							{
								try
								{
									$paramEnableCsuser = @{
										Identity	  = $UPN
										RegistrarPool = 'lync-bk7pool.qh.health.qld.gov.au'
										SipAddress    = "SIP:$UPN"
										DomainController = $DomainController
										WarningAction = 'SilentlyContinue'
										ErrorAction   = 'Stop'
									}
									
									Enable-Csuser @paramEnableCsuser
									
									
									$paramSetCsUser = @{
										Identity						  = $UPN
										AudioVideoDisabled			      = $True
										RemoteCallControlTelephonyEnabled = $False
										EnterpriseVoiceEnabled		      = $False
										DomainController				  = $DomainController
										WarningAction					  = 'SilentlyContinue'
										ErrorAction					      = 'Stop'
									}
									
									Set-CsUser @paramSetCsUser
									
									
									
									$paramGrantCsConferencingPolicy = @{
										Identity		 = $UPN
										PolicyName	     = 'Queensland Health Default Meeting Policy 1'
										DomainController = $DomainController
										WarningAction    = 'SilentlyContinue'
										ErrorAction	     = 'Stop'
									}
									
									Grant-CsConferencingPolicy @paramGrantCsConferencingPolicy
									
									
									$action = 'PROVISIONEDONPREM'
									$detail = "REMEDIATED : Enabled user for Skype on premise."
									
								}
								catch
								{
									$action = 'FAILED:PROVISIONING'
									$detail = "ERROR : Failed to Enable Skype with error, $($_.Exception.Message)."
								}
							}
							else
							{
								#reportOnly
								$action = 'REQUIREPROVISIONING'
								$detail = "WARNING : User not Provisionedfor Skype."
							}
							
							$SipHostingProvider = 'None'
							$sipCheck = 'NonSkypeUser'
							
						}
					}
					else
					{
						throw "No ProxyAddresses found for this user Mailbox"
					}
					
					$prop = [ordered]@{
						UserPrincipalName    = $UPN
						RecipienttypeDetails = $recipient.RecipientTypeDetails
						HostingProvider	     = $SipHostingProvider
						SipStatus		     = $sipCheck
						Action			     = $action
						Details			     = $detail
					}
				}
				else
				{
					$prop = [ordered]@{
						UserPrincipalName    = $UPN
						RecipienttypeDetails = $recipient.RecipientTypeDetails
						HostingProvider	     = 'SKIPPED'
						SipStatus		     = 'SKIPPED'
						Action			     = 'SKIPPED'
						Details			     = "SKIPPED : Not a User Mailbox"
					}
				}
			}
			catch
			{
				$prop = [ordered]@{
					UserPrincipalName    = $UPN
					RecipienttypeDetails = 'ERROR'
					HostingProvider	     = 'ERROR'
					SipStatus		     = 'ERROR'
					Action			     = 'ERROR'
					Details			     = "ERROR : $($_.Exception.message)"
				}
			}
			finally
			{
				$obj = New-Object -TypeName psobject -Property $prop
				Write-Output $obj
				
				if ($UserPrincipalName.count -gt 1)
				{
					$paramWriteProgress = @{
						Activity = 'Testing SIP Addresses with UPN'
						Status   = "Processing [$i] of [$($UserPrincipalName.Count)] users"
						PercentComplete = (($i / $UserPrincipalName.Count) * 100)
						CurrentOperation = "Completed : [$UPN]"
					}
					Write-Progress @paramWriteProgress
				}
				$i++
			}
		}
	}
	end
	{
		Write-Progress -Activity 'Testing SIP Addresses with UPN' -Completed
		$global:ErrorActionPreference = [System.Management.Automation.ActionPreference]::Continue
	}
}

function Move-QHSkypeUserToOkypeOnline
{
<#
	.SYNOPSIS
		A brief description of the Move-QHSkypeUserToOkypeOnline function.
	
	.DESCRIPTION
		A detailed description of the Move-QHSkypeUserToOkypeOnline function.
	
	.PARAMETER UserPrincipalName
		A description of the UserPrincipalName parameter.
	
	.PARAMETER BatchName
		A description of the BatchName parameter.
	
	.EXAMPLE
		PS C:\> Move-QHSkypeUserToOkypeOnline
	
	.NOTES
		Additional information about the function.
#>
	[CmdletBinding()]
	param
	(
		[String[]]$UserPrincipalName,
		[String]$BatchName
	)
	
	begin
	{
		$i = 1
		if ($BatchName -eq $null)
		{
			$BatchName = 'None'
		}
	}
	process
	{
		foreach ($UPN in $UserPrincipalName)
		{
			try
			{
				$recipient = Get-Recipient $UPN -ErrorAction Stop
				
				if ($recipient.RecipientTypeDetails -match 'User')
				{
					Move-CsUser -Identity $UPN -Target sipfed.online.lync.com -confirm:$false -ErrorAction Stop -WarningAction SilentlyContinue
					
					$prop = [Ordered]@{
						UserPrincipalName    = $UPN
						BatchName		     = $BatchName
						RecipientTypeDetails = $recipient.RecipientTypeDetails
						Status			     = 'Success'
						Details			     = 'None'
					}
				}
				else
				{
					$prop = [Ordered]@{
						UserPrincipalName    = $UPN
						BatchName		     = $BatchName
						RecipientTypeDetails = $recipient.RecipientTypeDetails
						Status			     = 'Skipped'
						Details			     = 'Skipped : Not a User Mailbox'
					}
				}
				
			}
			catch
			{
				$prop = [Ordered]@{
					UserPrincipalName    = $UPN
					BatchName		     = $BatchName
					RecipientTypeDetails = $recipient.RecipientTypeDetails
					Status			     = 'Failed'
					Details			     = "ERROR : $($_.Exception.Message)"
				}
			}
			finally
			{
				$obj = New-Object -TypeName System.Management.Automation.PSObject -Property $prop
				Write-Output $obj
				if ($UserPrincipalName.count -gt 1)
				{
					$paramWriteProgress = @{
						Activity = 'Moving Skype users to Office 365'
						Status   = "Processing [$i] of [$($UserPrincipalName.Count)] users"
						PercentComplete = (($i / $UserPrincipalName.Count) * 100)
						CurrentOperation = "Completed : [$UPN]"
					}
					Write-Progress @paramWriteProgress
				}
				$i++
			}
		}
	}
	end
	{
		Write-Progress -Activity 'Moving Skype users to Office 365' -Completed
	}
}

function Set-QHFunctionTemplate #Template
{
<#
	.SYNOPSIS
		A brief description of the Set-QHFunctionTemplate function.
	
	.DESCRIPTION
		A detailed description of the Set-QHFunctionTemplate function.
	
	.PARAMETER UserPrincipalName
		A description of the UserPrincipalName parameter.
	
	.PARAMETER ShowProgress
		A description of the ShowProgress parameter.
	
	.EXAMPLE
		PS C:\> Set-QHFunctionTemplate -UserPrincipalName 'value1'
	
	.NOTES
		Additional information about the function.
#>
	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory = $true)]
		[String[]]$UserPrincipalName,
		[switch]$ShowProgress
	)
	
	begin
	{
		$i = 1
		$Global:ErrorActionPreference = [System.Management.Automation.ActionPreference]::Stop
	}
	process
	{
		foreach ($UPN in $UserPrincipalName)
		{
			try
			{
				$recipient = Get-Recipient $UPN -ErrorAction Stop
			}
			catch
			{
				
			}
			finally
			{
				if ($UserPrincipalName.count -gt 1)
				{
					if ($ShowProgress)
					{
						$paramWriteProgress = @{
							Activity = 'Doing Some Processing'
							Status   = "Processing [$i] of [$($UserPrincipalName.Count)] users"
							PercentComplete = (($i / $UserPrincipalName.Count) * 100)
							CurrentOperation = "Completed : [$UPN]"
						}
						Write-Progress @paramWriteProgress
					}
				}
				$i++
			}
		}
	}
	end
	{
		Write-Progress -Activity 'Doing Some Processing' -Completed
		$Global:ErrorActionPreference = [System.Management.Automation.ActionPreference]::Continue
	}
}

function get-blah # sample recursive function
{
<#
	.SYNOPSIS
		A brief description of the get-blah function.
	
	.DESCRIPTION
		A detailed description of the get-blah function.
	
	.PARAMETER user
		A description of the user parameter.
	
	.PARAMETER counter
		A description of the counter parameter.
	
	.PARAMETER Retry
		A description of the Retry parameter.
	
	.PARAMETER inputObject
		A description of the inputObject parameter.
	
	.EXAMPLE
		PS C:\> get-blah
	
	.NOTES
		Additional information about the function.
#>
	[CmdletBinding()]
	param
	(
		[string]$user,
		[int]$counter = 0,
		[int]$Retry = 3,
		[object]$inputObject = $null
	)
	
	if ($counter -lt $Retry)
	{
		try
		{
			$null = Get-Recipient $user -ErrorAction Stop
			$counter++
			
			$prop = [ordered]@{
				User    = $user
				Attempt = $counter
				status  = 'Success'
				Details = 'None'
			}
			
			$obj = New-Object -TypeName psobject -Property $prop
			Write-Output $obj
		}
		catch
		{
			$counter++
			$prop = [ordered]@{
				User    = $user
				Attempt = $counter
				status  = 'Failure'
				Details = "ERROR : $($_.Exception.Message)"
			}
			
			$obj = New-Object -TypeName psobject -Property $prop
			#Write-Output $obj
			Write-host "Tried $counter" -ForegroundColor Cyan
			
			get-blah -user $user -counter $counter -Retry $Retry -inputObject $obj
			#get-blah -user $user  $counter -count $count -input $obj
		}
	}
	else
	{
		Write-Host "Failed Despite $Retry Attempts" -ForegroundColor Red
		Write-Output $inputObject
		
	}
}

function Set-QHGenericLicenseAttribute
{
<#
	.SYNOPSIS
		A brief description of the Set-QHGenericLicenseAttribute function.
	
	.DESCRIPTION
		A detailed description of the Set-QHGenericLicenseAttribute function.
	
	.PARAMETER SamAccountName
		A description of the SamAccountName parameter.
	
	.EXAMPLE
		PS C:\> Set-QHGenericLicenseAttribute -SamAccountName 'value1'
	
	.NOTES
		Additional information about the function.
#>
	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory = $true)]
		[String[]]$SamAccountName
	)
	begin
	{
		$i = 1
	}
	process
	{
		foreach ($SamAccount in $SamAccountName)
		{
			try
			{
				Write-Verbose "Getting AD info for user: $SamAccount"
				$AD = Get-AdUser $SamAccount -properties * -ErrorAction Stop
				Write-Verbose "Retrived AD info for user : $SamAccount"
				$xdate = Get-date
				$domainsuffix = "health.qld.gov.au"
				$UPN = "$($AD.SamAccountName)@$($DomainSuffix)"
				$NewExtAttrib8Value = "o365genericnomailbox,$UPN,$($xDate.ToString('yyyyMMddHHmm'))"
				Write-Verbose "Trying to set ExtenAttribute8 to : $NewExtAttrib8Value"
				Set-AdUser $SamAccount -replace @{ 'extensionAttribute8' = $NewExtAttrib8Value } -UserPrincipalName $UPN -ErrorAction Stop
				Write-Verbose "Successfully set ExtenAttribute8 to : $NewExtAttrib8Value"
				Write-Verbose "Trying to add $SamAccount to AD group : ACT-SaaS-O365-GenericAccountOfficeProPlusOnly"
				Add-ADGroupMember -Identity "ACT-SaaS-O365-GenericAccountOfficeProPlusOnly" -Members $Ad.SamAccountName
				Write-Verbose "Added $SamAccount to AD group : ACT-SaaS-O365-GenericAccountOfficeProPlusOnly"
				$prop = [Ordered] @{
					User				    = $SamAccount
					OLD_ExtensionAttribute8 = $AD.ExtensionAttribute8
					NEW_EXtensionAttribute8 = $NewExtAttrib8Value
					UPN					    = $UPN
					'ACT-SaaS-O365'		    = 'Added'
					Status				    = 'Success'
					Details				    = 'None'
				}
				
			}
			catch
			{
				$prop = [Ordered] @{
					User				    = $SamAccount
					OLD_ExtensionAttribute8 = $AD.ExtensionAttribute8
					NEW_EXtensionAttribute8 = $NewExtAttrib8Value
					UPN					    = $UPN
					'ACT-SaaS-O365'		    = 'Failed'
					Status				    = 'Failed'
					Details				    = "ERROR : $($_.Exception.Message)"
				}
			}
			finally
			{
				$obj = New-Object -TypeName System.Management.Automation.PSObject -Property $prop
				Write-Output $obj
			}
		}
	}
	end
	{
		
	}
}

function Worker-EnableLitigationHold
{
<#
	.SYNOPSIS
		A brief description of the Worker-EnableLitigationHold function.
	
	.DESCRIPTION
		A detailed description of the Worker-EnableLitigationHold function.
	
	.PARAMETER UserName
		A description of the UserName parameter.
	
	.EXAMPLE
		PS C:\> Worker-EnableLitigationHold
	
	.NOTES
		Additional information about the function.
#>
	param (
		$UserName
	)
	try
	{
		Set-ExoMailbox -Identity $UserName -litigationholdEnabled:$true -ErrorAction Stop -WarningAction SilentlyContinue
		Write-Output 'Success'
	}
	catch
	{
		Write-Output "Failed : $($_.Exception.Message)"
	}
}

function Worker-SetRegionalSettings
{
	param (
		$UserName
	)
	try
	{
		$paramSetEXOMailboxRegionalConfiguration = @{
			Identity	  = $UserName
			Language	  = 'en-AU'
			TimeZone	  = 'E. Australia Standard Time'
			ErrorAction   = 'Stop'
			WarningAction = 'SilentlyContinue'
		}
		
		Set-EXOMailboxRegionalConfiguration @paramSetEXOMailboxRegionalConfiguration
		Write-Output 'Success'
		
	}
	catch
	{
		Write-Output "Failed : $($_.Exception.Message)"
	}
}

function Worker-DisableImapPop
{
<#
	.SYNOPSIS
		A brief description of the Worker-DisableImapPop function.
	
	.DESCRIPTION
		A detailed description of the Worker-DisableImapPop function.
	
	.PARAMETER UserName
		A description of the UserName parameter.
	
	.EXAMPLE
		PS C:\> Worker-DisableImapPop
	
	.NOTES
		Additional information about the function.
#>
	param (
		$UserName
	)
	try
	{
		Set-EXOCASMailbox -identity $UserName -PopEnabled $false -ImapEnabled $false -ErrorAction 'Stop' -WarningAction 'SilentlyContinue'
		Write-Output 'Success'
	}
	catch
	{
		Write-Output "Failed : $($_.Exception.Message)"
	}
}

function Worker-EnableAuditingAddPol
{
<#
	.SYNOPSIS
		A brief description of the Worker-EnableAuditingAddPol function.
	
	.DESCRIPTION
		A detailed description of the Worker-EnableAuditingAddPol function.
	
	.PARAMETER UserName
		A description of the UserName parameter.
	
	.EXAMPLE
		PS C:\> Worker-EnableAuditingAddPol
	
	.NOTES
		Additional information about the function.
#>
	param (
		$UserName
	)
	try
	{
		$paramAudit = @{
			identity				  = $UPN
			AuditEnabled			  = $true
			AuditLogAgeLimit		  = 2555
			AddressBookPolicy		  = "Main QH ABP"
			SingleItemRecoveryEnabled = $true
			RetainDeletedItemsfor	  = 30
			ErrorAction			      = 'Stop'
			WarningAction			  = 'SilentlyContinue'
		}
		
		Set-ExoMailbox @paramAudit
		
		Write-Output 'Success'
	}
	catch
	{
		Write-Output "Failed : $($_.Exception.Message)"
	}
}

function Worker-SkypeSettings
{
<#
	.SYNOPSIS
		A brief description of the Worker-SkypeSettings function.
	
	.DESCRIPTION
		A detailed description of the Worker-SkypeSettings function.
	
	.PARAMETER UserName
		A description of the UserName parameter.
	
	.EXAMPLE
		PS C:\> Worker-SkypeSettings
	
	.NOTES
		Additional information about the function.
#>
	param (
		$UserName
	)
	try
	{
		Grant-365CsConferencingPolicy -Identity $UserName -PolicyName "BposSAllModalityMinVideoBW" -ErrorAction Stop -WarningAction SilentlyContinue
		Grant-365CsExternalAccessPolicy -Identity $UserName -PolicyName "FederationOnly" -ErrorAction Stop -WarningAction SilentlyContinue
		Write-Output 'Success'
	}
	catch
	{
		Write-Output "Failed : $($_.Exception.Message)"
	}
}

function Worker-MoveSkypeUserstoO365
{
<#
	.SYNOPSIS
		A brief description of the Worker-MoveSkypeUserstoO365 function.
	
	.DESCRIPTION
		A detailed description of the Worker-MoveSkypeUserstoO365 function.
	
	.PARAMETER UserName
		A description of the UserName parameter.
	
	.EXAMPLE
		PS C:\> Worker-MoveSkypeUserstoO365
	
	.NOTES
		Additional information about the function.
#>
	param (
		$UserName
	)
	try
	{
		Move-CsUser -Identity $UserName -Target sipfed.online.lync.com -confirm:$false -ErrorAction Stop -WarningAction SilentlyContinue
		Write-Output 'Success'
	}
	catch
	{
		Write-Output "Failed : $($_.Exception.Message)"
	}
	
}

function Start-QHPostMigrationTasks
{
<#
	.SYNOPSIS
		This will apply necessary settings for users post migration. Setting details in the description.
	
	.DESCRIPTION
		This will apply following setting for migrated mailboxes :
		- Litigation hold : Enabled
		- Regional settings : Language to 'en-AU' and timezone to 'E. Australia Standard Time'
		- Disables : IMAP and POP
		- Enables : Auditing, Sets AuditLogAge  to 2555, Sets Single Item recovery and retains deleted item to 30days
		- SkypeSettings : Sets ConferencingPolicy to 'BposSAllModalityMinVideoBW' and External Access Policy to 'FederationOnly'
	
	.PARAMETER UserPrincipalName
		This is the UserPrincipalName of the migrated user.
	
	.PARAMETER BatchName
		BatchName is only is only used as a reference and better output.
	
	.PARAMETER ShowProgress
		if this option is selected then it will show a progress bar for more than one user.
	
	.EXAMPLE
		PS C:\> Start-PostMigrationTasks -UserPrincipalName user1@health.qld.gov.au, User2@health.qld.gov.au -BatchName 'Batchx' -ShowProgress
	
	.NOTES
		Additional information about the function.
#>
	
	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory = $true)]
		[String[]]$UserPrincipalName,
		[Parameter(Mandatory = $true)]
		[String]$BatchName = 'None',
		[String]$AdminUserName = 'None',
		[Parameter(Mandatory = $false)]
		[Switch]$ShowProgress
	)
	
	begin
	{
		$i = 1
	}
	process
	{
		foreach ($UPN in $UserPrincipalName)
		{
			try
			{
				$recipient = Get-ExoRecipient $UPN -ErrorAction Stop
				
				if ($recipient.RecipientTypeDetails -match 'UserMailbox')
				{
					$moveSkypeUserstoO365 = 'Skipped:Module not available' #Worker-MoveSkypeUserstoO365 -UserName $UPN
					$LitigationHold = Worker-EnableLitigationHold -UserName $UPN
					$RegionalSettings = Worker-SetRegionalSettings -UserName $UPN
					$DisablePopImap = Worker-DisableImapPop -UserName $UPN
					$EnableAuditAdd = Worker-EnableAuditingAddPol -UserName $UPN
					$SkypeSettings = "DoneSeperately" #Worker-SkypeSettings -UserName $UPN
				}
				elseif ($recipient.RecipientTypeDetails -match 'SharedMailbox')
				{
					$moveSkypeUserstoO365 = "Skipped : Not a User mailbox"
					$LitigationHold = Worker-EnableLitigationHold -UserName $UPN
					$RegionalSettings = Worker-SetRegionalSettings -UserName $UPN
					$DisablePopImap = Worker-DisableImapPop -UserName $UPN
					$EnableAuditAdd = Worker-EnableAuditingAddPol -UserName $UPN
					$SkypeSettings = "Skipped : Not a User mailbox"
				}
				
				$prop = [Ordered] @{
					UserPrincipalName    = $UPN
					BatchName		     = $BatchName
					RecipientTypeDetails = $recipient.RecipientTypeDetails
					MoveSkypeUsersToO365 = $moveSkypeUserstoO365
					LitigationHold	     = $LitigationHold
					RegionalSettings	 = $RegionalSettings
					DisabledImapPOP	     = $DisablePopImap
					AuditAndAddress	     = $EnableAuditAdd
					SkypeSettings	     = $SkypeSettings
					SessionAdmin		 = $AdminUserName
					Details			     = 'None'
				}
			}
			catch
			{
				$prop = [Ordered] @{
					UserPrincipalName    = $UPN
					BatchName		     = $BatchName
					RecipientTypeDetails = 'Error'
					MoveSkypeUsersToO365 = 'Error'
					LitigationHold	     = 'Error'
					RegionalSettings	 = 'Error'
					DisabledImapPOP	     = 'Error'
					AuditAndAddress	     = 'Error'
					SkypeSettings	     = 'Error'
					SessionAdmin		 = $AdminUserName
					Details			     = "ERROR : $($_.Exception.Message)"
				}
			}
			finally
			{
				$obj = New-Object -TypeName System.Management.Automation.PSObject -Property $prop
				Write-Output $obj
				if ($UserPrincipalName.count -gt 1)
				{
					if ($ShowProgress)
					{
						$paramWriteProgress = @{
							Activity = 'Executing Post Migration Tasks'
							Status   = "Processing [$i] of [$($UserPrincipalName.Count)] users"
							PercentComplete = (($i / $UserPrincipalName.Count) * 100)
							CurrentOperation = "Completed : [$UPN]"
						}
						Write-Progress @paramWriteProgress
					}
				}
				$i++
			}
		}
	}
	end
	{
		Write-Progress -Activity 'Executing Post Migration Tasks' -Completed
	}
}

function Get-QHCalendarDelegates
{
<#
	.SYNOPSIS
		A brief description of the Get-QHCalendarDelegates function.
	
	.DESCRIPTION
		A detailed description of the Get-QHCalendarDelegates function.
	
	.PARAMETER UserPrincipalName
		A description of the UserPrincipalName parameter.
	
	.PARAMETER BatchName
		A description of the BatchName parameter.
	
	.PARAMETER ShowProgress
		A description of the ShowProgress parameter.
	
	.EXAMPLE
		PS C:\> Get-QHCalendarDelegates -UserPrincipalName 'value1' -BatchName 'value2'
	
	.NOTES
		Additional information about the function.
#>
	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory = $true)]
		[String[]]$UserPrincipalName,
		[Parameter(Mandatory = $true)]
		[String]$BatchName,
		[Switch]$ShowProgress
	)
	
	begin
	{
		$Global:ErrorActionPreference = [System.Management.Automation.ActionPreference]::Stop
		$i = 1
		$date = (get-date).ToString('dd-MM-yyyy')
	}
	process
	{
		foreach ($UPN in $UserPrincipalName)
		{
			try
			{
				$calDel = Get-CalendarProcessing -Identity $UPN -ErrorAction Stop
				
				if ($calDel.ResourceDelegates -ne $null)
				{
					$prop = [ordered] @{
						Date			  = $date
						UserPrincipalName = $UPN
						BatchName		  = $BatchName
						CalendarDelegates = $($calDel.ResourceDelegates.Name -join ',').ToString().Trim()
						TotalDelegates    = $calDel.ResourceDelegates.Name.Count
						Details		      = 'None'
					}
				}
				else
				{
					$prop = [ordered] @{
						Date			  = $date
						UserPrincipalName = $UPN
						BatchName		  = $BatchName
						CalendarDelegates = 'NoCalendarDelegates'
						TotalDelegates    = 0
						Details		      = 'None'
					}
				}
			}
			catch
			{
				$prop = [ordered] @{
					Date			  = $date
					UserPrincipalName = $UPN
					BatchName		  = $BatchName
					CalendarDelegates = 'ERROR'
					TotalDelegates    = 'ERROR'
					Details		      = "ERROR : $($_.Exception.Message)"
				}
			}
			finally
			{
				$obj = New-Object -TypeName System.Management.Automation.PSObject -Property $prop
				Write-Output $obj
				
				if ($UserPrincipalName.count -gt 1)
				{
					if ($ShowProgress)
					{
						$paramWriteProgress = @{
							Activity = 'Retriving Calendar Delegates'
							Status   = "Processing [$i] of [$($UserPrincipalName.Count)] users"
							PercentComplete = (($i / $UserPrincipalName.Count) * 100)
							CurrentOperation = "Completed : [$UPN]"
						}
						Write-Progress @paramWriteProgress
					}
				}
				$i++
			}
		}
	}
	end
	{
		Write-Progress -Activity 'Retriving Calendar Delegates' -Completed
		$Global:ErrorActionPreference = [System.Management.Automation.ActionPreference]::Continue
	}
}

function Get-xQHMailboxFolderPermission
{
<#
	.SYNOPSIS
		A brief description of the Get-xQHMailboxFolderPermission function.
	
	.DESCRIPTION
		A detailed description of the Get-xQHMailboxFolderPermission function.
	
	.PARAMETER UserPrincipalName
		A description of the UserPrincipalName parameter.
	
	.PARAMETER BatchName
		A description of the BatchName parameter.
	
	.PARAMETER ShowProgress
		A description of the ShowProgress parameter.
	
	.EXAMPLE
		PS C:\> Get-xQHMailboxFolderPermission -UserPrincipalName 'value1' -BatchName 'value2'
	
	.NOTES
		Additional information about the function.
#>
	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory = $true)]
		[String[]]$UserPrincipalName,
		[Parameter(Mandatory = $true)]
		[String]$BatchName,
		[Switch]$ShowProgress
	)
	
	begin
	{
		$Global:ErrorActionPreference = [System.Management.Automation.ActionPreference]::Stop
		$i = 1
		$date = (get-date).ToString('dd-MM-yyyy')
	}
	process
	{
		foreach ($UPN in $UserPrincipalName)
		{
			try
			{
				$result = @()
				
				$recipient = Get-Recipient $UPN -ErrorAction Stop
				if ($recipient.RecipientTypeDetails -notmatch 'Remote')
				{
					$topLevel = Get-MailboxFolderPermission -Identity "$($UPN)" -ErrorAction Stop | Where-Object {
						$_.User -notmatch 'Default|Anonymous'
					}
					$FolCal = Get-MailboxFolderPermission -Identity "$($UPN):\calendar" -ErrorAction Stop | Where-Object {
						$_.User -notmatch 'Default|Anonymous'
					}
					$FolInbox = Get-MailboxFolderPermission -Identity "$($UPN):\Inbox" -ErrorAction Stop | Where-Object {
						$_.User -notmatch 'Default|Anonymous'
					}
					$FolContacts = Get-MailboxFolderPermission -Identity "$($UPN):\Contacts" -ErrorAction Stop | Where-Object {
						$_.User -notmatch 'Default|Anonymous'
					}
					$FolSent = Get-MailboxFolderPermission -Identity "$($UPN):\Sent Items" -ErrorAction Stop | Where-Object {
						$_.User -notmatch 'Default|Anonymous'
					}
					
					#$result = @($TopLevel, $FolCal, $FolInbox, $FolContacts, $FolSent)
					
					
					$result += $TopLevel
					$result += $FolCal
					$result += $FolInbox
					$result += $FolContacts
					$result += $FolSent
					
					
					$prop = [ordered]@{
						Date				 = $date
						UserPrincipalName    = $UPN
						RecipientTypeDetails = $recipient.RecipientTypeDetails
						TopLevelAccess	     = if ($topLevel.User.ADRecipient.UserPrincipalName -ne $null)
						{
							$($topLevel.User.ADRecipient.UserPrincipalName -join ',').ToString().Trim()
						} Else {
							'None'
						}
						CalendarAccess	     = if ($FolCal.User.ADRecipient.UserPrincipalName -ne $null)
						{
							$($FolCal.User.ADRecipient.UserPrincipalName -join ',').ToString().Trim()
						} else {
							'None'
						}
						InboxAccess		     = if ($FolInbox.User.ADRecipient.UserPrincipalName -ne $null)
						{
							$($FolInbox.User.ADRecipient.UserPrincipalName -join ',').ToString().Trim()
						} else {
							'None'
						}
						ContactsAccess	     = if ($FolContacts.User.ADRecipient.UserPrincipalName -ne $null)
						{
							$($FolContacts.User.ADRecipient.UserPrincipalName -join ',').ToString().Trim()
						} else {
							'None'
						}
						SentItemsAccess	     = if ($FolSent.User.ADRecipient.UserPrincipalName -ne $null)
						{
							$($FolSent.User.ADRecipient.UserPrincipalName -join ',').ToString().Trim()
						} else {
							'None'
						}
						TotalUniqueUsers	 = $($result | Sort-Object User -Unique).Count
						Details			     = 'None'
					}
				}
				else
				{
					$prop = [ordered]@{
						Date				 = $date
						UserPrincipalName    = $UPN
						RecipientTypeDetails = $recipient.RecipientTypeDetails
						TopLevelAccess	     = 'Skipped'
						CalendarAccess	     = 'Skipped'
						InboxAccess		     = 'Skipped'
						ContactsAccess	     = 'Skipped'
						SentItemsAccess	     = 'Skipped'
						TotalUniqueUsers	 = 0
						Details			     = 'Skipped : The user is not an On Premise User'
					}
				}
			}
			catch
			{
				$prop = [ordered]@{
					Date				 = $date
					UserPrincipalName    = $UPN
					RecipientTypeDetails = 'ERROR'
					TopLevelAccess	     = 'ERROR'
					CalendarAccess	     = 'ERROR'
					InboxAccess		     = 'ERROR'
					ContactsAccess	     = 'ERROR'
					SentItemsAccess	     = 'ERROR'
					TotalUniqueUsers	 = 'ERROR'
					Details			     = "$($_.Exception.Message)"
				}
			}
			finally
			{
				$obj = New-Object -TypeName System.Management.Automation.PSObject -Property $prop
				Write-Output $obj
				
				if ($UserPrincipalName.count -gt 1)
				{
					if ($ShowProgress)
					{
						$paramWriteProgress = @{
							Activity = 'Retriving Mailbox Folder Permissions'
							Status   = "Processing [$i] of [$($UserPrincipalName.Count)] users"
							PercentComplete = (($i / $UserPrincipalName.Count) * 100)
							CurrentOperation = "Completed : [$UPN]"
						}
						Write-Progress @paramWriteProgress
					}
				}
				$i++
			}
		}
	}
	end
	{
		Write-Progress -Activity 'Retriving Mailbox Folder Permissions' -Completed
		$Global:ErrorActionPreference = [System.Management.Automation.ActionPreference]::Continue
	}
}

function Get-QHEwsEVState # OLD
{
<#
	.SYNOPSIS
		A brief description of the Get-QHEwsEVState function.
	
	.DESCRIPTION
		A detailed description of the Get-QHEwsEVState function.
	
	.PARAMETER UserPrincipalName
		A description of the UserPrincipalName parameter.
	
	.PARAMETER AuthUsername
		A description of the AuthUsername parameter.
	
	.PARAMETER AuthPassword
		A description of the AuthPassword parameter.
	
	.PARAMETER DownloadDirectory
		A description of the DownloadDirectory parameter.
	
	.PARAMETER ShowProgress
		A description of the ShowProgress parameter.
	
	.EXAMPLE
		PS C:\> Get-QHEwsEVState -UserPrincipalName 'value1' -AuthUsername 'value2' -AuthPassword 'value3'
	
	.NOTES
		Additional information about the function.
#>
	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory = $true)]
		[String[]]$UserPrincipalName,
		[Parameter(Mandatory = $true)]
		[String]$AuthUsername,
		[Parameter(Mandatory = $true)]
		[String]$AuthPassword,
		[Parameter(Mandatory = $false)]
		[String]$DownloadDirectory = $env:TEMP,
		[Switch]$ShowProgress
	)
	
	begin
	{
		$targetDir = Set-Qhdir -path "$DownloadDirectory\EV_EWSSettings"
		$EwsUrl = "https://outlook.office365.com/EWS/Exchange.asmx"
		$MessageClass = 'IPM.Note.EnterpriseVault.Settings'
		$Impersonate = $true
		
		#region Functions
		
		function LoadEWSManagedAPI
		{
			# Find and load the managed API
			
			if (![string]::IsNullOrEmpty($EWSManagedApiPath))
			{
				if ({ Test-Path $EWSManagedApiPath })
				{
					Add-Type -Path $EWSManagedApiPath
					return $true
				}
				Write-Verbose ([string]::Format("WARNING : Managed API not found at specified location: {0}", $EWSManagedApiPath))
			}
			
			$a = Get-ChildItem -Recurse "C:\Program Files (x86)\Microsoft\Exchange\Web Services" -ErrorAction SilentlyContinue | Where-Object { ($_.PSIsContainer -eq $false) -and ($_.Name -eq "Microsoft.Exchange.WebServices.dll") }
			if (!$a)
			{
				$a = Get-ChildItem -Recurse "C:\Program Files\Microsoft\Exchange\Web Services" -ErrorAction SilentlyContinue | Where-Object { ($_.PSIsContainer -eq $false) -and ($_.Name -eq "Microsoft.Exchange.WebServices.dll") }
			}
			
			if ($a)
			{
				# Load EWS Managed API
				Write-Verbose ([string]::Format("INFO : Using managed API {0} found at: {1}", $a.VersionInfo.FileVersion, $a.VersionInfo.FileName))
				Add-Type -Path $a.VersionInfo.FileName
				return $true
			}
			return $false
		}
		
		function TrustAllCerts
		{
			
			## Create a compilation environment
			$Provider = New-Object Microsoft.CSharp.CSharpCodeProvider
			$Compiler = $Provider.CreateCompiler()
			$Params = New-Object System.CodeDom.Compiler.CompilerParameters
			$Params.GenerateExecutable = $False
			$Params.GenerateInMemory = $True
			$Params.IncludeDebugInformation = $False
			$Params.ReferencedAssemblies.Add("System.DLL") | Out-Null
			
			$TASource = @'
        namespace Local.ToolkitExtensions.Net.CertificatePolicy {
        public class TrustAll : System.Net.ICertificatePolicy {
            public TrustAll()
            { 
            }
            public bool CheckValidationResult(System.Net.ServicePoint sp,
                                                System.Security.Cryptography.X509Certificates.X509Certificate cert, 
                                                System.Net.WebRequest req, int problem)
            {
                return true;
            }
        }
        }
'@
			$TAResults = $Provider.CompileAssemblyFromSource($Params, $TASource)
			$TAAssembly = $TAResults.CompiledAssembly
			
			## We now create an instance of the TrustAll and attach it to the ServicePointManager
			$TrustAll = $TAAssembly.CreateInstance("Local.ToolkitExtensions.Net.CertificatePolicy.TrustAll")
			[System.Net.ServicePointManager]::CertificatePolicy = $TrustAll
		}
		
		#endregion Functions
		
		if (!(LoadEWSManagedAPI))
		{
			Write-Host "Failed to locate EWS Managed API, cannot continue" -ForegroundColor Magenta
			break
		}
		$i = 1
	}
	process
	{
		foreach ($UPN in $UserPrincipalName)
		{
			try
			{
				Write-Verbose "Trying to Verify Mailbox on Office 365 for user $UPN"
				$null = Get-ExoMailbox $UPN -ErrorAction Stop
				
				Write-Verbose "Trying to Create a new EWS Service Object"
				$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013_SP1)
				
				if ($EwsUrl)
				{
					Write-Verbose "Using $EwsUrl to Establish EWS connection"
					$service.URL = New-Object Uri($EwsUrl)
				}
				else
				{
					Write-Verbose "Performing autodiscover for $UPN"
					$service.AutodiscoverUrl($UPN)
					Write-Verbose "EWS Url found: ", $service.Url
				}
				
				if ($AuthUsername -and $AuthPassword)
				{
					Write-Verbose "Applying given credentials for, $AuthUsername"
					if ($AuthDomain)
					{
						$service.Credentials = New-Object  Microsoft.Exchange.WebServices.Data.WebCredentials($AuthUsername, $AuthPassword, $AuthDomain)
					}
					else
					{
						$service.Credentials = New-Object  Microsoft.Exchange.WebServices.Data.WebCredentials($AuthUsername, $AuthPassword)
					}
				}
				else
				{
					Write-Verbose "Using default credentials"
					$service.UseDefaultCredentials = $true
				}
				if ($Impersonate)
				{
					Write-Verbose "Impersonating $UPN"
					$service.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $UPN)
					#$FolderId = $rootFolder
				}
				else
				{
					# If we're not impersonating, we will specify the mailbox in case we are accessing a mailbox that is not the authenticating account's
					$mbx = New-Object Microsoft.Exchange.WebServices.Data.Mailbox($UPN)
					$FolderId = New-Object Microsoft.Exchange.WebServices.Data.FolderId($rootFolder, $mbx)
				}
				
				Write-Verbose "Trying to Establish Connection with Inbox"
				$folderid = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox, $UPN)
				$Inbox = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service, $folderid)
				
				Write-Verbose "Trying to Create Search filter for message class $MessageClass in the hidden Emails"
				$searchFilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.ItemSchema]::ItemClass, $MessageClass)
				
				Write-Verbose "Creting a view for the inbox items"
				$View = New-Object Microsoft.Exchange.WebServices.Data.ItemView(1000)
				#Define Assoicated Traversal
				
				$View.Traversal = [Microsoft.Exchange.WebServices.Data.ItemTraversal]::Associated;
				$fiItems = $null
				
				do
				{
					Write-Verbose "Invoking Search in the inbox"
					$fiItems = $service.FindItems($Inbox.Id, $searchFilter, $View)
					#[Void]$service.LoadPropertiesForItems($fiItems,$psPropset)
					$View.Offset += $fiItems.Items.Count
				}
				while ($fiItems.MoreAvailable -eq $true)
				
				if ($fiItems -ne $null)
				{
					Write-Verbose "Search Returned results, Analyzing them now"
					if ($fiItems.TotalCount -gt 1)
					{
						$prop = [ordered]@{
							Mailbox   = $UPN
							ItemClass = $fiItems.Items[0].ItemClass
							HasAttachments = 'MultipleItems'
							AttachmentDetails = 'MultipleItems'
							EVState   = 'MultipleItems'
							path	  = 'MultipleItems'
							Details   = "Warning : Multiple Items returned for the Class $MessageClass, Please check via EWSEDITOR"
						}
					}
					else
					{
						$fiItems.Items[0].Load()
						Write-Verbose "Checking if EV Attachment exist"
						
						#$attachment = $fiItems.Items[0].Attachments[0]
						$attachment = $fiItems.Items[0] | Select-Object -ExpandProperty Attachments | Where-Object { $_.Name -eq 'EnterpriseVaultSettings.txt' }
						# if Attachment -ne $null then do the load else no attachment
						
						if ($attachment -ne $null)
						{
							$path = "$targetDir\$($UPN)_$($attachment.Name)"
							Write-Verbose "Saving EV settings file at $path"
							$attachment.Load($path)
							$AttchName = $attachment.Name
							$content = Get-Content -Encoding Unicode $path
							$mbxStateRaw = $content | Where-Object { $_.StartsWith('MailboxState') }
							
							if ($mbxStateRaw -eq 'MailboxState=1')
							{
								$EVSettingState = 'EV Enabled (Should be Disabled)'
							}
							elseif ($mbxStateRaw -eq 'MailboxState=2')
							{
								$EVSettingState = 'EV Disabled (Correct)'
							}
							else
							{
								$EVSettingState = 'Could not Identify'
							}
						}
						else
						{
							$EVSettingState = "NoEVSettingsAttachmentFound"
							$AttchName = 'None'
							$path = 'None'
						}
						
						$prop = [ordered]@{
							Mailbox   = $UPN
							ItemClass = $fiItems.Items[0].ItemClass
							HasAttachments = $fiItems.Items[0].HasAttachments
							AttachmentDetails = $AttchName
							EVState   = $EVSettingState
							Path	  = $path
							Details   = 'None'
						}
					}
				}
				else
				{
					$prop = [ordered]@{
						Mailbox		      = $UPN
						ItemClass		  = 'No Item found'
						HasAttachments    = 'No Item found'
						AttachmentDetails = 'No Item found'
						EVState		      = 'No Item found'
						Path			  = 'No Item found'
						Details		      = 'None'
					}
					
				}
			}
			catch
			{
				$prop = [ordered]@{
					Mailbox		      = $UPN
					ItemClass		  = 'ERROR'
					HasAttachments    = 'ERROR'
					AttachmentDetails = 'ERROR'
					EVState		      = 'ERROR'
					Path			  = 'ERROR'
					Details		      = "ERROR : $($_.Exception.Message)"
				}
			}
			finally
			{
				Write-Verbose "Generating Output for $UPN"
				$obj = New-Object -TypeName System.Management.Automation.PSObject -Property $prop
				Write-Output $obj
				if ($UserPrincipalName.count -gt 1)
				{
					if ($ShowProgress)
					{
						$paramWriteProgress = @{
							Activity = 'Retriving EV Status from Office 365'
							Status   = "Processing [$i] of [$($UserPrincipalName.Count)] users"
							PercentComplete = (($i / $UserPrincipalName.Count) * 100)
							CurrentOperation = "Completed : [$UPN]"
						}
						Write-Progress @paramWriteProgress
					}
				}
				$i++
				Write-Verbose "Completed Process for $UPN"
			}
		}
	}
	end
	{
		Write-Progress -Activity 'Retriving EV Status from Office 365' -Completed
	}
}

function Set-QHEwsEVState
{
<#
	.SYNOPSIS
		A brief description of the Set-QHEwsEVState function.
	
	.DESCRIPTION
		A detailed description of the Set-QHEwsEVState function.
	
	.PARAMETER UserPrincipalName
		A description of the UserPrincipalName parameter.
	
	.PARAMETER AuthUsername
		A description of the AuthUsername parameter.
	
	.PARAMETER AuthPassword
		A description of the AuthPassword parameter.
	
	.PARAMETER DownloadDirectory
		A description of the DownloadDirectory parameter.
	
	.PARAMETER ShowProgress
		A description of the ShowProgress parameter.
	
	.EXAMPLE
		PS C:\> Set-QHEwsEVState -UserPrincipalName 'value1' -AuthUsername 'value2' -AuthPassword 'value3'
	
	.NOTES
		Additional information about the function.
#>
	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory = $true)]
		[String[]]$UserPrincipalName,
		[Parameter(Mandatory = $true)]
		[String]$AuthUsername,
		[Parameter(Mandatory = $true)]
		[String]$AuthPassword,
		[Parameter(Mandatory = $false)]
		[String]$DownloadDirectory = $env:TEMP,
		[Switch]$ShowProgress
	)
	
	begin
	{
		$targetDir = Set-Qhdir -path "$DownloadDirectory\EV_EWSSettings"
		$EwsUrl = "https://outlook.office365.com/EWS/Exchange.asmx"
		$MessageClass = 'IPM.Note.EnterpriseVault.Settings'
		$Impersonate = $true
		
		#region Functions
		
		function LoadEWSManagedAPI
		{
			# Find and load the managed API
			
			if (![string]::IsNullOrEmpty($EWSManagedApiPath))
			{
				if ({ Test-Path $EWSManagedApiPath })
				{
					Add-Type -Path $EWSManagedApiPath
					return $true
				}
				Write-Verbose ([string]::Format("WARNING : Managed API not found at specified location: {0}", $EWSManagedApiPath))
			}
			
			$a = Get-ChildItem -Recurse "C:\Program Files (x86)\Microsoft\Exchange\Web Services" -ErrorAction SilentlyContinue | Where-Object { ($_.PSIsContainer -eq $false) -and ($_.Name -eq "Microsoft.Exchange.WebServices.dll") }
			if (!$a)
			{
				$a = Get-ChildItem -Recurse "C:\Program Files\Microsoft\Exchange\Web Services" -ErrorAction SilentlyContinue | Where-Object { ($_.PSIsContainer -eq $false) -and ($_.Name -eq "Microsoft.Exchange.WebServices.dll") }
			}
			
			if ($a)
			{
				# Load EWS Managed API
				Write-Verbose ([string]::Format("INFO : Using managed API {0} found at: {1}", $a.VersionInfo.FileVersion, $a.VersionInfo.FileName))
				Add-Type -Path $a.VersionInfo.FileName
				return $true
			}
			return $false
		}
		
		function TrustAllCerts
		{
    <#
    .SYNOPSIS
    Set certificate trust policy to trust self-signed certificates (for test servers).
    #>
			
			## Create a compilation environment
			$Provider = New-Object Microsoft.CSharp.CSharpCodeProvider
			$Compiler = $Provider.CreateCompiler()
			$Params = New-Object System.CodeDom.Compiler.CompilerParameters
			$Params.GenerateExecutable = $False
			$Params.GenerateInMemory = $True
			$Params.IncludeDebugInformation = $False
			$Params.ReferencedAssemblies.Add("System.DLL") | Out-Null
			
			$TASource = @'
        namespace Local.ToolkitExtensions.Net.CertificatePolicy {
        public class TrustAll : System.Net.ICertificatePolicy {
            public TrustAll()
            { 
            }
            public bool CheckValidationResult(System.Net.ServicePoint sp,
                                                System.Security.Cryptography.X509Certificates.X509Certificate cert, 
                                                System.Net.WebRequest req, int problem)
            {
                return true;
            }
        }
        }
'@
			$TAResults = $Provider.CompileAssemblyFromSource($Params, $TASource)
			$TAAssembly = $TAResults.CompiledAssembly
			
			## We now create an instance of the TrustAll and attach it to the ServicePointManager
			$TrustAll = $TAAssembly.CreateInstance("Local.ToolkitExtensions.Net.CertificatePolicy.TrustAll")
			[System.Net.ServicePointManager]::CertificatePolicy = $TrustAll
		}
		
		#endregion Functions
		
		if (!(LoadEWSManagedAPI))
		{
			Write-Host "Failed to locate EWS Managed API, cannot continue" -ForegroundColor Magenta
			break
		}
		$i = 1
	}
	process
	{
		foreach ($UPN in $UserPrincipalName)
		{
			try
			{
				Write-Verbose "Trying to Verify Mailbox on Office 365 for user $UPN"
				$null = Get-ExoMailbox $UPN -ErrorAction Stop
				
				Write-Verbose "Trying to Create a new EWS Service Object"
				$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013_SP1)
				
				if ($EwsUrl)
				{
					Write-Verbose "Using $EwsUrl to Establish EWS connection"
					$service.URL = New-Object Uri($EwsUrl)
				}
				else
				{
					Write-Verbose "Performing autodiscover for $UPN"
					$service.AutodiscoverUrl($UPN)
					Write-Verbose "EWS Url found: ", $service.Url
				}
				
				if ($AuthUsername -and $AuthPassword)
				{
					Write-Verbose "Applying given credentials for, $AuthUsername"
					if ($AuthDomain)
					{
						$service.Credentials = New-Object Microsoft.Exchange.WebServices.Data.WebCredentials($AuthUsername, $AuthPassword, $AuthDomain)
					}
					else
					{
						$service.Credentials = New-Object  Microsoft.Exchange.WebServices.Data.WebCredentials($AuthUsername, $AuthPassword)
					}
				}
				else
				{
					Write-Verbose "Using default credentials"
					$service.UseDefaultCredentials = $true
				}
				if ($Impersonate)
				{
					Write-Verbose "Impersonating $UPN"
					$service.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $UPN)
					#$FolderId = $rootFolder
				}
				else
				{
					# If we're not impersonating, we will specify the mailbox in case we are accessing a mailbox that is not the authenticating account's
					$mbx = New-Object Microsoft.Exchange.WebServices.Data.Mailbox($UPN)
					$FolderId = New-Object Microsoft.Exchange.WebServices.Data.FolderId($rootFolder, $mbx)
				}
				
				Write-Verbose "Trying to Establish Connection with Inbox"
				$folderid = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox, $UPN)
				$Inbox = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service, $folderid)
				
				Write-Verbose "Trying to Create Search filter for message class $MessageClass in the hidden Emails"
				$searchFilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.ItemSchema]::ItemClass, $MessageClass)
				
				Write-Verbose "Creting a view for the inbox items"
				$View = New-Object Microsoft.Exchange.WebServices.Data.ItemView(1000)
				#Define Assoicated Traversal
				
				$View.Traversal = [Microsoft.Exchange.WebServices.Data.ItemTraversal]::Associated;
				$fiItems = $null
				
				do
				{
					Write-Verbose "Invoking Search in the inbox"
					$fiItems = $service.FindItems($Inbox.Id, $searchFilter, $View)
					#[Void]$service.LoadPropertiesForItems($fiItems,$psPropset)
					$View.Offset += $fiItems.Items.Count
				}
				while ($fiItems.MoreAvailable -eq $true)
				
				if ($fiItems -ne $null)
				{
					Write-Verbose "Search Returned results, Analyzing them now"
					if ($fiItems.TotalCount -gt 1)
					{
						
						$prop = [ordered]@{
							Mailbox   = $UPN
							ItemClass = $fiItems.Items[0].ItemClass
							HasAttachments = 'MultipleItems'
							AttachmentDetails = 'MultipleItems'
							EVState   = 'MultipleItems'
							Status    = 'MultipleItems'
							OldSettings = 'MultipleItems'
							Details   = "Warning : Multiple Items returned for the Class $MessageClass, Please check via EWSEDITOR"
						}
					}
					else
					{
						$fiItems.Items[0].Load()
						Write-Verbose "Checking if EV Attachment exist"
						
						#$attachment = $fiItems.Items[0].Attachments[0]
						$attachment = $fiItems.Items[0] | Select-Object -ExpandProperty Attachments | Where-Object { $_.Name -eq 'EnterpriseVaultSettings.txt' }
						# if Attachment -ne $null then do the load else no attachment
						
						if ($attachment -ne $null)
						{
							$path = "$targetDir\$($UPN)_$($attachment.Name)"
							Write-Verbose "Saving EV settings file at $path"
							$attachment.Load($path)
							$AttchName = $attachment.Name
							$content = Get-Content -Encoding Unicode $path
							$mbxStateRaw = $content | Where-Object { $_.StartsWith('MailboxState') }
							
							if ($mbxStateRaw -eq 'MailboxState=1')
							{
								$EVSettingState = 'EV Enabled (Should be Disabled)'
								Write-Verbose "Removing any old EV Settings text file"
								
								
								$targetPath = "$env:TEMP\EnterpriseVaultSettings.txt"
								remove-item -Path $targetPath -ErrorAction SilentlyContinue -Force
								Write-Verbose "Generating new EV Settings file at $targetPath"
								
								$content -replace "MailboxState=1", "MailboxState=2" | Out-File $targetPath
								
								Write-Verbose "Removing Old EV Settings Attachment : $($attachment.Name)"
								if ($fiItems.Items[0].Attachments.Remove($attachment))
								{
									$fiItems.Items[0].Update([Microsoft.Exchange.WebServices.Data.ConflictResolutionMode]::AutoResolve)
									Write-Verbose "Removed Old Attachment"
									$null = $fiItems.Items[0].Attachments.AddFileAttachment($targetPath)
									$fiItems.Items[0].Update([Microsoft.Exchange.WebServices.Data.ConflictResolutionMode]::AutoResolve)
									Write-Verbose "Added New Attachment"
									remove-item -Path $targetPath -ErrorAction SilentlyContinue -Force
									$newSettings = 'Success'
								}
								else
								{
									$newSettings = 'Failed'
								}
							}
							elseif ($mbxStateRaw -eq 'MailboxState=2')
							{
								$EVSettingState = 'EV Disabled (Correct)'
								$newSettings = 'Skipped'
							}
							else
							{
								$EVSettingState = 'Could not Identify'
								$newSettings = 'Skipped'
							}
						}
						else
						{
							$EVSettingState = "NoEVSettingsAttachmentFound"
							$AttchName = 'None'
							$path = 'None'
							$newSettings = 'Skipped'
						}
						
						$prop = [ordered]@{
							Mailbox   = $UPN
							ItemClass = $fiItems.Items[0].ItemClass
							HasAttachments = $fiItems.Items[0].HasAttachments
							AttachmentDetails = $AttchName
							EVState   = $EVSettingState
							Status    = $newSettings
							OldSettings = $path
							Details   = 'None'
						}
					}
				}
				else
				{
					$prop = [ordered]@{
						Mailbox		      = $UPN
						ItemClass		  = 'No Item found'
						HasAttachments    = 'No Item found'
						AttachmentDetails = 'No Item found'
						EVState		      = 'No Item found'
						Status		      = 'No Item found'
						OldSettings	      = 'No Item found'
						Details		      = 'None'
					}
					
				}
			}
			catch
			{
				$prop = [ordered]@{
					Mailbox		      = $UPN
					ItemClass		  = 'ERROR'
					HasAttachments    = 'ERROR'
					AttachmentDetails = 'ERROR'
					EVState		      = 'ERROR'
					Status		      = 'ERROR'
					OldSettings	      = 'ERROR'
					Details		      = "ERROR : $($_.Exception.Message)"
				}
			}
			finally
			{
				Write-Verbose "Generating Output for $UPN"
				$obj = New-Object -TypeName System.Management.Automation.PSObject -Property $prop
				Write-Output $obj
				if ($UserPrincipalName.count -gt 1)
				{
					if ($ShowProgress)
					{
						$paramWriteProgress = @{
							Activity = 'Retriving EV Status from Office 365'
							Status   = "Processing [$i] of [$($UserPrincipalName.Count)] users"
							PercentComplete = (($i / $UserPrincipalName.Count) * 100)
							CurrentOperation = "Completed : [$UPN]"
						}
						Write-Progress @paramWriteProgress
					}
				}
				$i++
				Write-Verbose "Completed Process for $UPN"
			}
		}
	}
	end
	{
		Write-Progress -Activity 'Retriving EV Status from Office 365' -Completed
	}
}

function Start-QHPremigrationTasks
{
<#
	.SYNOPSIS
		A brief description of the Start-QHPremigrationTasks function.
	
	.DESCRIPTION
		A detailed description of the Start-QHPremigrationTasks function.
	
	.PARAMETER BatchName
		A description of the BatchName parameter.
	
	.PARAMETER Stage
		A description of the Stage parameter.
	
	.EXAMPLE
		PS C:\> Start-QHPremigrationTasks -BatchName 'value1' -Stage Stage1-Premigration
	
	.NOTES
		Additional information about the function.
#>
	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory = $true)]
		[String[]]$BatchName,
		[Parameter(Mandatory = $true)]
		[ValidateSet('Stage1-Premigration', 'Stage2-DisableArchive', 'Stage3-LicensingAndMFA', 'Stage4-DisableEV')]
		$Stage
	)
	
	begin
	{
		#Connect-QHOnpremExchange -Server exc-casbtpp005
		Write-Host "$Stage Process started for all the Batches" -ForegroundColor Gray
	}
	process
	{
		foreach ($batch in $BatchName)
		{
			Write-Host "START : Process for batch [$batch]" -ForegroundColor Cyan
			
			if (Navigate-QHMigrationFolder $batch)
			{
				try
				{
					$users = Import-Csv ".\$($batch)_Validation.csv" |
					Where-Object { $_.lookup -eq 'PASSED' } |
					Select-Object -ExpandProperty EmailAddress
					
					if ($users -ne $null)
					{
						if ($Stage -eq 'Stage1-Premigration')
						{
							Write-Host "INFO : Starting Process : NonAcceptedDomains" -ForegroundColor Yellow
							Remove-QHNonAcceptedDomainsEmailAlias -UserPrincipalName $users -RemoveDomains exchange.health.qld.gov.au, groupwise.qld.gov.au -BatchName $batch |
							Export-csv "$($batch)_NonAcceptedDomains.csv" -NoTypeInformation
							Write-Host "INFO : Completed Process : NonAcceptedDomains" -ForegroundColor Yellow
							#needs Exchange .net
							
							Write-Host "INFO : Starting Process : RoutingAddress" -ForegroundColor Yellow
							Add-zQHRoutngAddress -UserPrincipalName $users -BatchName $batch |
							Export-csv "$($batch)_RoutingAddress.csv" -NoTypeInformation
							Write-Host "INFO : Completed Process : RoutingAddress" -ForegroundColor Yellow
							
							Write-Host "INFO : Starting Process : SipAddress" -ForegroundColor Yellow
							Validate-QHADSipAddress -UserPrincipalName $users -BatchName $batch -DomainController EAD-WDCBK7P03 -Remediate |
							Export-Csv "$($batch)_SipAddressValidation.csv" -NoTypeInformation
							Write-Host "INFO : Completed Process : SipAddress" -ForegroundColor Yellow
							
							Write-Host "INFO : Starting Process : CalendarDelegates" -ForegroundColor Yellow
							Get-QHCalendarDelegates -UserPrincipalName $users -BatchName $batch -ShowProgress |
							Export-Csv  "$($batch)_CalendarDelegates.csv" -NoTypeInformation
							# needs  Exchange .net
							Write-Host "INFO : Completed Process : CalendarDelegates" -ForegroundColor Yellow
							
							Write-Host "INFO : Starting Process : MailboxFolderPermission" -ForegroundColor Yellow
							Get-xQHMailboxFolderPermission -UserPrincipalName $users -BatchName $batch -ShowProgress |
							Export-Csv  "$($batch)_MailboxFolderPermission.csv" -NoTypeInformation
							# needs  Exchange .net
							Write-Host "INFO : Completed Process : MailboxFolderPermission" -ForegroundColor Yellow
							
							Write-Host "INFO : Starting Process : Migration Csv for EndPoints" -ForegroundColor Yellow
							Generate-QhMigrationBatchCSV -EmailAddresses $users -BatchName $batch
							Write-Host "INFO : Completed Process : Migration Csv for EndPoints" -ForegroundColor Yellow
							
							Write-Host "INFO : Starting Process : ExportMailboxInfo" -ForegroundColor Yellow
							Export-QHMailboxInfo -UserName $Users -batchName $batch
							Write-Host "INFO : Completed Process : ExportMailboxInfo" -ForegroundColor Yellow
						}
						elseif ($Stage -eq 'Stage2-DisableArchive')
						{
							Write-Host "INFO : Starting Process : MailboxQuota" -ForegroundColor Yellow
							Set-QHMailboxQuota50GB -UserPrincipalName $users -DomainController EAD-WDCBK7P04 -BatchName $batch |
							Export-Csv "$($batch)_Quotas.csv" -NoTypeInformation
							Write-Host "INFO : Completed Process : MailboxQuota" -ForegroundColor Yellow
							
							Write-Host "INFO : Starting Process : EVGroupRemoval" -ForegroundColor Yellow
							Remove-QHEVGroupMemberShip -UserPrincipalName $Users -BatchName $batch -AddToStagingGrp |
							Export-Csv "$($batch)_RemoveEVGroups.csv" -NoTypeInformation
							Write-Host "INFO : Completed Process : EVGroupRemoval" -ForegroundColor Yellow
							
							Write-Host "INFO : Starting Process : Export Calendar Processing Settings" -ForegroundColor Yellow
							Get-QhRoomCalendarSettings -UserPrincipalName $Users -showProgress |
							Export-Csv "$($batch)_CalendarProcessingStats.csv" -NoTypeInformation
							Write-Host "INFO : Completed Process : Export Calendar Processing Settings" -ForegroundColor Yellow
							
							Write-Host "INFO : Starting Process : ReferenceList" -ForegroundColor Yellow
							Create-QHReferenceList -UserCSV ".\$($batch)_Validation.csv" |
							Export-csv "$($batch)_ReferenceList.csv" -NoTypeInformation
							Write-Host "INFO : Completed Process : ReferenceList" -ForegroundColor Yellow
						}
						elseif ($Stage -eq 'Stage3-LicensingAndMFA')
						{
							Write-Host "INFO : Starting Process : Assigning License" -ForegroundColor Yellow
							Set-iQHLicense -UserPrincipalName $users -ShowProgress |
							Export-csv "$($batch)_License.csv" -NoTypeInformation
							Write-Host "INFO : Completed Process : Assigning License" -ForegroundColor Yellow
							
							Write-Host "INFO : Starting Process : MFA" -ForegroundColor Yellow
							Enable-QHMFA -UserPrincipalName $Users -ShowProgress |
							Export-Csv "$($batch)_MFA.csv" -NoTypeInformation
							Write-Host "INFO : Completed Process : MFA" -ForegroundColor Yellow
						}
						elseif ($Stage -eq 'Stage4-DisableEV')
						{
							Write-Host "INFO : Starting Process : EV Disable on eve-admbtpp101" -ForegroundColor Yellow
							Disable-QhEV -batchName $batch
							Write-Host "INFO : Completed Process : EV Disable on eve-admbtpp101" -ForegroundColor Yellow
							
							Write-Host "INFO : Starting Process : Staging EVGroupRemoval" -ForegroundColor Yellow
							Remove-QHEVGroupMemberShip -UserPrincipalName $Users -BatchName $batch |
							Export-Csv "$($batch)_RemoveEVGroupsPostDisableEV.csv" -NoTypeInformation
							Write-Host "INFO : Completed Process : Staging EVGroupRemoval" -ForegroundColor Yellow
						}
						else
						{
							# Donothing	
						}
					}
					else
					{
						Write-Host "WARNING : No Users to Process for $batch. Please confirm the validation is run on the $batch" -ForegroundColor Yellow
					}
				}
				catch
				{
					Write-Host "ERROR : $($_.Exception.Message)" -ForegroundColor Magenta
				}
			}
			
			Write-Host "COMPLETED : Process for batch [$batch]" -ForegroundColor Cyan
		}
	}
	end
	{
		Write-Host "$Stage Process completed for all the Batches" -ForegroundColor Gray
	}
	#TODO: Place script here
}

function Get-QHEVStatus
{
<#
	.SYNOPSIS
		A brief description of the Get-QHEVStatus function.
	
	.DESCRIPTION
		A detailed description of the Get-QHEVStatus function.
	
	.PARAMETER UserPrincipalName
		A description of the UserPrincipalName parameter.
	
	.PARAMETER AuthUsername
		A description of the AuthUsername parameter.
	
	.PARAMETER AuthPassword
		A description of the AuthPassword parameter.
	
	.PARAMETER DownloadDirectory
		A description of the DownloadDirectory parameter.
	
	.PARAMETER ShowProgress
		A description of the ShowProgress parameter.
	
	.PARAMETER Server
		A description of the Server parameter.
	
	.EXAMPLE
		PS C:\> Get-QHEVStatus -UserPrincipalName 'value1' -AuthUsername 'value2' -AuthPassword 'value3' -Server exc-casbtpp004
	
	.NOTES
		Author : Rana Banerjee - rana.banerjee@health.qld.gov.au.
#>
	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory = $true)]
		[String[]]$UserPrincipalName,
		[Parameter(Mandatory = $true)]
		[String]$AuthUsername,
		[Parameter(Mandatory = $true)]
		[String]$AuthPassword,
		[Parameter(Mandatory = $false)]
		[String]$DownloadDirectory = $env:TEMP,
		[Parameter(Mandatory = $false)]
		[Switch]$ShowProgress,
		[Parameter(Mandatory = $true)]
		[ValidateSet('exc-casbtpp004', 'exc-casbtpp003', 'exc-casbtpp002', 'exc-casbtpp005', 'exc-chmndcp001', 'outlook.office365.com')]
		$Server
	)
	
	begin
	{
		$Global:ErrorActionPreference = [System.Management.Automation.ActionPreference]::Stop
		if ($Server -eq 'outlook.office365.com')
		{
			$ExVer = 'Exchange2013_SP1'
		}
		else
		{
			$ExVer = 'Exchange2010_SP2'
		}
		
		$targetDir = Set-Qhdir -path "$DownloadDirectory\EV_EWSSettings"
		$EwsUrl = "https://$Server/EWS/Exchange.asmx"
		$MessageClass = 'IPM.Note.EnterpriseVault.Settings'
		$Impersonate = $true
		
		#region Functions
		
		function LoadEWSManagedAPI
		{
			# Find and load the managed API
			
			if (![string]::IsNullOrEmpty($EWSManagedApiPath))
			{
				if ({ Test-Path $EWSManagedApiPath })
				{
					Add-Type -Path $EWSManagedApiPath
					return $true
				}
				Write-Verbose ([string]::Format("WARNING : Managed API not found at specified location: {0}", $EWSManagedApiPath))
			}
			
			$a = Get-ChildItem -Recurse "C:\Program Files (x86)\Microsoft\Exchange\Web Services" -ErrorAction SilentlyContinue | Where-Object { ($_.PSIsContainer -eq $false) -and ($_.Name -eq "Microsoft.Exchange.WebServices.dll") }
			if (!$a)
			{
				$a = Get-ChildItem -Recurse "C:\Program Files\Microsoft\Exchange\Web Services" -ErrorAction SilentlyContinue | Where-Object { ($_.PSIsContainer -eq $false) -and ($_.Name -eq "Microsoft.Exchange.WebServices.dll") }
			}
			
			if ($a)
			{
				# Load EWS Managed API
				Write-Verbose ([string]::Format("INFO : Using managed API {0} found at: {1}", $a.VersionInfo.FileVersion, $a.VersionInfo.FileName))
				Add-Type -Path $a.VersionInfo.FileName
				return $true
			}
			return $false
		}
		
		function TrustAllCerts
		{
    <#
    .SYNOPSIS
    Set certificate trust policy to trust self-signed certificates (for test servers).
    #>
			
			## Create a compilation environment
			$Provider = New-Object Microsoft.CSharp.CSharpCodeProvider
			$Compiler = $Provider.CreateCompiler()
			$Params = New-Object System.CodeDom.Compiler.CompilerParameters
			$Params.GenerateExecutable = $False
			$Params.GenerateInMemory = $True
			$Params.IncludeDebugInformation = $False
			$Params.ReferencedAssemblies.Add("System.DLL") | Out-Null
			
			$TASource = @'
        namespace Local.ToolkitExtensions.Net.CertificatePolicy {
        public class TrustAll : System.Net.ICertificatePolicy {
            public TrustAll()
            { 
            }
            public bool CheckValidationResult(System.Net.ServicePoint sp,
                                                System.Security.Cryptography.X509Certificates.X509Certificate cert, 
                                                System.Net.WebRequest req, int problem)
            {
                return true;
            }
        }
        }
'@
			$TAResults = $Provider.CompileAssemblyFromSource($Params, $TASource)
			$TAAssembly = $TAResults.CompiledAssembly
			
			## We now create an instance of the TrustAll and attach it to the ServicePointManager
			$TrustAll = $TAAssembly.CreateInstance("Local.ToolkitExtensions.Net.CertificatePolicy.TrustAll")
			[System.Net.ServicePointManager]::CertificatePolicy = $TrustAll
		}
		
		#endregion Functions
		
		if (!(LoadEWSManagedAPI))
		{
			Write-Host "Failed to locate EWS Managed API, cannot continue" -ForegroundColor Magenta
			break
		}
		$i = 1
		
		
	}
	process
	{
		foreach ($UPN in $UserPrincipalName)
		{
			try
			{
				if ($Server -eq 'outlook.office365.com')
				{
					Write-Verbose "Trying to Verify Mailbox on Office 365 for user $UPN"
					$null = Get-ExoMailbox $UPN -ErrorAction Stop
				}
				else
				{
					Write-Verbose "Trying to Verify Mailbox on Office 365 for user $UPN"
					$null = Get-Mailbox $UPN -ErrorAction Stop
				}
				
				Write-Verbose "Trying to Create a new EWS Service Object"
				$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::$ExVer)
				
				if ($EwsUrl)
				{
					Write-Verbose "Using $EwsUrl to Establish EWS connection"
					$service.URL = New-Object Uri($EwsUrl)
				}
				else
				{
					Write-Verbose "Performing autodiscover for $UPN"
					$service.AutodiscoverUrl($UPN)
					Write-Verbose "EWS Url found: ", $service.Url
				}
				
				if ($AuthUsername -and $AuthPassword)
				{
					Write-Verbose "Applying given credentials for, $AuthUsername"
					if ($AuthDomain)
					{
						$service.Credentials = New-Object  Microsoft.Exchange.WebServices.Data.WebCredentials($AuthUsername, $AuthPassword, $AuthDomain)
					}
					else
					{
						$service.Credentials = New-Object  Microsoft.Exchange.WebServices.Data.WebCredentials($AuthUsername, $AuthPassword)
					}
				}
				else
				{
					Write-Verbose "Using default credentials"
					$service.UseDefaultCredentials = $true
				}
				if ($Impersonate)
				{
					Write-Verbose "Impersonating $UPN"
					$service.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $UPN)
					#$FolderId = $rootFolder
				}
				else
				{
					# If we're not impersonating, we will specify the mailbox in case we are accessing a mailbox that is not the authenticating account's
					$mbx = New-Object Microsoft.Exchange.WebServices.Data.Mailbox($UPN)
					$FolderId = New-Object Microsoft.Exchange.WebServices.Data.FolderId($rootFolder, $mbx)
				}
				
				Write-Verbose "Trying to Establish Connection with Inbox"
				$folderid = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox, $UPN)
				$Inbox = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service, $folderid)
				
				Write-Verbose "Trying to Create Search filter for message class $MessageClass in the hidden Emails"
				$searchFilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.ItemSchema]::ItemClass, $MessageClass)
				
				Write-Verbose "Creting a view for the inbox items"
				$View = New-Object Microsoft.Exchange.WebServices.Data.ItemView(1000)
				#Define Assoicated Traversal
				
				$View.Traversal = [Microsoft.Exchange.WebServices.Data.ItemTraversal]::Associated;
				$fiItems = $null
				
				do
				{
					Write-Verbose "Invoking Search in the inbox"
					$fiItems = $service.FindItems($Inbox.Id, $searchFilter, $View)
					#[Void]$service.LoadPropertiesForItems($fiItems,$psPropset)
					$View.Offset += $fiItems.Items.Count
				}
				while ($fiItems.MoreAvailable -eq $true)
				
				if ($fiItems -ne $null)
				{
					Write-Verbose "Search Returned results, Analyzing them now"
					if ($fiItems.TotalCount -gt 1)
					{
						$prop = [ordered]@{
							Mailbox   = $UPN
							ItemClass = $fiItems.Items[0].ItemClass
							HasAttachments = 'MultipleItems'
							AttachmentDetails = 'MultipleItems'
							EVState   = 'MultipleItems'
							path	  = 'MultipleItems'
							Details   = "Warning : Multiple Items returned for the Class $MessageClass, Please check via EWSEDITOR"
						}
					}
					else
					{
						$fiItems.Items[0].Load()
						Write-Verbose "Checking if EV Attachment exist"
						
						#$attachment = $fiItems.Items[0].Attachments[0]
						$attachment = $fiItems.Items[0] | Select-Object -ExpandProperty Attachments | Where-Object { $_.Name -eq 'EnterpriseVaultSettings.txt' }
						# if Attachment -ne $null then do the load else no attachment
						
						if ($attachment -ne $null)
						{
							$path = "$targetDir\$($UPN)_$($attachment.Name)"
							Write-Verbose "Saving EV settings file at $path"
							$attachment.Load($path)
							$AttchName = $attachment.Name
							$content = Get-Content -Encoding Unicode $path
							$mbxStateRaw = $content | Where-Object { $_.StartsWith('MailboxState') }
							
							if ($mbxStateRaw -eq 'MailboxState=1')
							{
								$EVSettingState = 'EV Enabled (Should be Disabled)'
							}
							elseif ($mbxStateRaw -eq 'MailboxState=2')
							{
								$EVSettingState = 'EV Disabled (Correct)'
							}
							else
							{
								$EVSettingState = 'Could not Identify'
							}
						}
						else
						{
							$EVSettingState = "NoEVSettingsAttachmentFound"
							$AttchName = 'None'
							$path = 'None'
						}
						
						$prop = [ordered]@{
							Mailbox   = $UPN
							ItemClass = $fiItems.Items[0].ItemClass
							HasAttachments = $fiItems.Items[0].HasAttachments
							AttachmentDetails = $AttchName
							EVState   = $EVSettingState
							Path	  = $path
							Details   = 'None'
						}
					}
				}
				else
				{
					$prop = [ordered]@{
						Mailbox		      = $UPN
						ItemClass		  = 'No Item found'
						HasAttachments    = 'No Item found'
						AttachmentDetails = 'No Item found'
						EVState		      = 'No Item found'
						Path			  = 'No Item found'
						Details		      = 'None'
					}
					
				}
			}
			catch
			{
				$prop = [ordered]@{
					Mailbox		      = $UPN
					ItemClass		  = 'ERROR'
					HasAttachments    = 'ERROR'
					AttachmentDetails = 'ERROR'
					EVState		      = 'ERROR'
					Path			  = 'ERROR'
					Details		      = "ERROR : $($_.Exception.Message)"
				}
			}
			finally
			{
				Write-Verbose "Generating Output for $UPN"
				$obj = New-Object -TypeName System.Management.Automation.PSObject -Property $prop
				Write-Output $obj
				if ($UserPrincipalName.count -gt 1)
				{
					if ($ShowProgress)
					{
						$paramWriteProgress = @{
							Activity = 'Retriving EV Status from Office 365'
							Status   = "Processing [$i] of [$($UserPrincipalName.Count)] users"
							PercentComplete = (($i / $UserPrincipalName.Count) * 100)
							CurrentOperation = "Completed : [$UPN]"
						}
						Write-Progress @paramWriteProgress
					}
				}
				$i++
				Write-Verbose "Completed Process for $UPN"
			}
		}
	}
	end
	{
		Write-Progress -Activity 'Retriving EV Status from Office 365' -Completed
		$Global:ErrorActionPreference = [System.Management.Automation.ActionPreference]::Continue
	}
}

function get-QhMailboxFolderCountOnPrem
{
<#
	.SYNOPSIS
		A brief description of the get-QhMailboxFolderCountOnPrem function.
	
	.DESCRIPTION
		A detailed description of the get-QhMailboxFolderCountOnPrem function.
	
	.PARAMETER UserPrincipalName
		A description of the UserPrincipalName parameter.
	
	.PARAMETER ShowProgress
		A description of the ShowProgress parameter.
	
	.EXAMPLE
		PS C:\> get-QhMailboxFolderCountOnPrem
	
	.NOTES
		Additional information about the function.
#>
	[CmdletBinding()]
	param
	(
		[String[]]$UserPrincipalName,
		[Switch]$ShowProgress
	)
	
	begin
	{
		$global:ErrorActionPreference = [System.Management.Automation.ActionPreference]::Stop
		$i = 1
		$DomainController = Get-ADDomainController | Select-Object -ExpandProperty hostname
		
	}
	process
	{
		foreach ($UPN in $UserPrincipalName)
		{
			try
			{
				$mbxStats = Get-MailboxFolderStatistics $UPN -DomainController $DomainController -ErrorAction Stop
				$prop = [ordered]@{
					UserPrincipalName = $UPN
					Mailbox		      = 'OnPremise'
					TotalFolders	  = $mbxStats.Count
					Details		      = 'None'
				}
			}
			catch
			{
				$prop = [ordered]@{
					UserPrincipalName = $UPN
					Mailbox		      = 'OnPremise'
					TotalFolders	  = 'ERROR'
					Details		      = "$($_.Exception.Message)"
				}
			}
			finally
			{
				$obj = New-Object -TypeName psobject -Property $prop
				Write-Output $obj
				
				if ($ShowProgress)
				{
					if ($UserPrincipalName.count -gt 1)
					{
						$paramWriteProgress = @{
							Activity = 'Counting Folders for Mailboxes Onpremise'
							Status   = "Processing [$i] of [$($UserPrincipalName.Count)] users"
							PercentComplete = (($i / $UserPrincipalName.Count) * 100)
							CurrentOperation = "Completed : [$UPN]"
						}
						Write-Progress @paramWriteProgress
					}
				}
				$i++
			}
			
		}
	}
	end
	{
		Write-Progress -Activity 'Counting Folders for Mailboxes' -Completed
		$global:ErrorActionPreference = [System.Management.Automation.ActionPreference]::Continue
	}
}

function get-QhMailboxFolderCountO365
{
<#
	.SYNOPSIS
		A brief description of the get-QhMailboxFolderCountO365 function.
	
	.DESCRIPTION
		A detailed description of the get-QhMailboxFolderCountO365 function.
	
	.PARAMETER UserPrincipalName
		A description of the UserPrincipalName parameter.
	
	.PARAMETER ShowProgress
		A description of the ShowProgress parameter.
	
	.EXAMPLE
		PS C:\> get-QhMailboxFolderCountO365
	
	.NOTES
		Additional information about the function.
#>
	[CmdletBinding()]
	param
	(
		[String[]]$UserPrincipalName,
		[Switch]$ShowProgress
	)
	
	begin
	{
		$global:ErrorActionPreference = [System.Management.Automation.ActionPreference]::Stop
		$i = 1
		
		try
		{
			$null = Get-ExoAcceptedDomain -ErrorAction Stop
		}
		catch
		{
			Write-Host "ERROR : Please connect to Office 365 and re-run this commandlet" -ForegroundColor Magenta
			break
		}
	}
	process
	{
		foreach ($UPN in $UserPrincipalName)
		{
			try
			{
				$mbxStats = Get-ExoMailboxFolderStatistics $UPN
				$prop = [ordered]@{
					UserPrincipalName = $UPN
					Mailbox		      = 'O365'
					TotalFolders	  = $mbxStats.Count
					Details		      = 'None'
				}
			}
			catch
			{
				$prop = [ordered]@{
					UserPrincipalName = $UPN
					Mailbox		      = 'O365'
					TotalFolders	  = 'ERROR'
					Details		      = "$($_.Exception.Message)"
				}
			}
			finally
			{
				$obj = New-Object -TypeName psobject -Property $prop
				Write-Output $obj
				
				if ($ShowProgress)
				{
					if ($UserPrincipalName.count -gt 1)
					{
						$paramWriteProgress = @{
							Activity = 'Counting Folders for Mailboxes'
							Status   = "Processing [$i] of [$($UserPrincipalName.Count)] users"
							PercentComplete = (($i / $UserPrincipalName.Count) * 100)
							CurrentOperation = "Completed : [$UPN]"
						}
						Write-Progress @paramWriteProgress
					}
				}
				$i++
			}
			
		}
	}
	end
	{
		Write-Progress -Activity 'Counting Folders for Mailboxes' -Completed
		$global:ErrorActionPreference = [System.Management.Automation.ActionPreference]::Continue
	}
}

function Get-QHScheduledTasks
{
    <#
    .SYNOPSIS
        Get scheduled task information from a system
    
    .DESCRIPTION
        Get scheduled task information from a system

        Uses Schedule.Service COM object, falls back to SchTasks.exe as needed.
        When we fall back to SchTasks, we add empty properties to match the COM object output.

    .PARAMETER ComputerName
        One or more computers to run this against

    .PARAMETER Folder
        Scheduled tasks folder to query.  By default, "\"

    .PARAMETER Recurse
        If specified, recurse through folders below $folder.
        
        Note:  We also recurse if we use SchTasks.exe

    .PARAMETER Path
        If specified, path to export XML files
        
        Details:
            Naming scheme is computername-taskname.xml
            Please note that the base filename is used when importing a scheduled task.  Rename these as needed prior to importing!

    .PARAMETER Exclude
        If specified, exclude tasks matching this regex (we use -notmatch $exclude)

    .PARAMETER CompatibilityMode
        If specified, pull scheduled tasks only with the schtasks.exe command, which works against older systems.
    
        Notes:
            Export is not possible with this switch.
            Recurse is implied with this switch.
    
    .EXAMPLE
    
        #Get scheduled tasks from the root folder of server1 and c-is-ts-91
        Get-ScheduledTasks server1, c-is-ts-91

    .EXAMPLE

        #Get scheduled tasks from all folders on server1, not in a Microsoft folder
        Get-ScheduledTasks server1 -recurse -Exclude "\\Microsoft\\"

    .EXAMPLE
    
        #Get scheduled tasks from all folders on server1, not in a Microsoft folder, and export in XML format (can be used to import scheduled tasks)
        Get-ScheduledTasks server1 -recurse -Exclude "\\Microsoft\\" -path 'D:\Scheduled Tasks'

    .NOTES
    
        Properties returned    : When they will show up
            ComputerName       : All queries
            Name               : All queries
            Path               : COM object queries, added synthetically if we fail back from COM to SchTasks
            Enabled            : COM object queries
            Action             : All queries.  Schtasks.exe queries include both Action and Arguments in this property
            Arguments          : COM object queries
            UserId             : COM object queries
            LastRunTime        : All queries
            NextRunTime        : All queries
            Status             : All queries
            Author             : All queries
            RunLevel           : COM object queries
            Description        : COM object queries
            NumberOfMissedRuns : COM object queries

        Thanks to help from Brian Wilhite, Jaap Brasser, and Jan Egil's functions:
            http://gallery.technet.microsoft.com/scriptcenter/Get-SchedTasks-Determine-5e04513f
            http://gallery.technet.microsoft.com/scriptcenter/Get-Scheduled-tasks-from-3a377294
            http://blog.crayon.no/blogs/janegil/archive/2012/05/28/working_2D00_with_2D00_scheduled_2D00_tasks_2D00_from_2D00_windows_2D00_powershell.aspx

    .FUNCTIONALITY
        Computers

    #>
	[cmdletbinding(
				   DefaultParameterSetName = 'COM'
				   )]
	param (
		[parameter(
				   ValueFromPipeline = $true,
				   ValueFromPipelineByPropertyName = $true,
				   ValueFromRemainingArguments = $false,
				   Position = 0
				   )]
		[Alias("host", "server", "computer")]
		[string[]]$ComputerName = "localhost",
		[parameter()]
		[string]$folder = "\",
		[parameter(ParameterSetName = 'COM')]
		[switch]$recurse,
		[parameter(ParameterSetName = 'COM')]
		[validatescript({
				#Test path if provided, otherwise allow $null
				if ($_)
				{
					Test-Path -PathType Container -path $_
				}
				else
				{
					$true
				}
			})]
		[string]$Path = $null,
		[parameter()]
		[string]$Exclude = $null,
		[parameter(ParameterSetName = 'SchTasks')]
		[switch]$CompatibilityMode
	)
	begin
	{
		
		if (-not $CompatibilityMode)
		{
			$sch = New-Object -ComObject Schedule.Service
			
			#thanks to Jaap Brasser - http://gallery.technet.microsoft.com/scriptcenter/Get-Scheduled-tasks-from-3a377294
			function Get-AllTaskSubFolders
			{
				[cmdletbinding()]
				param (
					# Set to use $Schedule as default parameter so it automatically list all files
					# For current schedule object if it exists.
					$FolderRef = $sch.getfolder("\"),
					[switch]$recurse
				)
				
				#No recurse?  Return the folder reference
				if (-not $recurse)
				{
					$FolderRef
				}
				#Recurse?  Build up an array!
				else
				{
					try
					{
						#This will fail on older systems...
						$folders = $folderRef.getfolders(1)
						
						#Extract results into array
						$ArrFolders = @(
							if ($folders)
							{
								foreach ($fold in $folders)
								{
									$fold
									if ($fold.getfolders(1))
									{
										Get-AllTaskSubFolders -FolderRef $fold
									}
								}
							}
						)
					}
					catch
					{
						#If we failed and the expected error, return folder ref only!
						if ($_.tostring() -like '*Exception calling "GetFolders" with "1" argument(s): "The request is not supported.*')
						{
							$folders = $null
							Write-Warning "GetFolders failed, returning root folder only: $_"
							return $FolderRef
						}
						else
						{
							throw $_
						}
					}
					
					#Return only unique results
					$Results = @($ArrFolders) + @($FolderRef)
					$UniquePaths = $Results | Select-Object -ExpandProperty path -Unique
					$Results | Where-Object{
						$UniquePaths -contains $_.path
					}
				}
			} #Get-AllTaskSubFolders
		}
		
		function Get-SchTasks
		{
			[cmdletbinding()]
			param ([string]$computername,
				[string]$folder,
				[switch]$CompatibilityMode)
			
			#we format the properties to match those returned from com objects
			$result = @(schtasks.exe /query /v /s $computername /fo csv |
				convertfrom-csv |
				Where-Object{
					$_.taskname -ne "taskname" -and $_.taskname -match $($folder.replace("\", "\\"))
				} |
				Select-Object @{
					label		  = "ComputerName"; expression = {
						$computername
					}
				},
							  @{
					label	    = "Name"; expression = {
						$_.TaskName
					}
				},
							  @{
					label		     = "Action"; expression = {
						$_."Task To Run"
					}
				},
							  @{
					label			   = "LastRunTime"; expression = {
						$_."Last Run Time"
					}
				},
							  @{
					label			   = "NextRunTime"; expression = {
						$_."Next Run Time"
					}
				},
							  "Status",
							  "Author"
			)
			
			if ($CompatibilityMode)
			{
				#User requested compat mode, don't add props
				$result
			}
			else
			{
				#If this was a failback, we don't want to affect display of props for comps that don't fail... include empty props expected for com object
				#We also extract task name and path to parent for the Name and Path props, respectively
				foreach ($item in $result)
				{
					$name = @($item.Name -split "\\")[-1]
					$taskPath = $item.name
					$item | Select-Object ComputerName, @{
						label = "Name"; expression = {
							$name
						}
					}, @{
						label	  = "Path"; Expression = {
							$taskPath
						}
					}, Enabled, Action, Arguments, UserId, LastRunTime, NextRunTime, Status, Author, RunLevel, Description, NumberOfMissedRuns
				}
			}
		} #Get-SchTasks
	}
	process
	{
		#loop through computers
		foreach ($computer in $computername)
		{
			
			#bool in case com object fails, fall back to schtasks
			$failed = $false
			
			write-verbose "Running against $computer"
			try
			{
				
				#use com object unless in compatibility mode.  Set compatibility mode if this fails
				if (-not $compatibilityMode)
				{
					
					try
					{
						#Connect to the computer
						$sch.Connect($computer)
						
						if ($recurse)
						{
							$AllFolders = Get-AllTaskSubFolders -FolderRef $sch.GetFolder($folder) -recurse -ErrorAction stop
						}
						else
						{
							$AllFolders = Get-AllTaskSubFolders -FolderRef $sch.GetFolder($folder) -ErrorAction stop
						}
						Write-verbose "Looking through $($AllFolders.count) folders on $computer"
						
						foreach ($fold in $AllFolders)
						{
							
							#Get tasks in this folder
							$tasks = $fold.GetTasks(0)
							
							Write-Verbose "Pulling data from $($tasks.count) tasks on $computer in $($fold.name)"
							foreach ($task in $tasks)
							{
								
								#extract helpful items from XML
								$Author = ([regex]::split($task.xml, '<Author>|</Author>'))[1]
								$UserId = ([regex]::split($task.xml, '<UserId>|</UserId>'))[1]
								$Description = ([regex]::split($task.xml, '<Description>|</Description>'))[1]
								$Action = ([regex]::split($task.xml, '<Command>|</Command>'))[1]
								$Arguments = ([regex]::split($task.xml, '<Arguments>|</Arguments>'))[1]
								$RunLevel = ([regex]::split($task.xml, '<RunLevel>|</RunLevel>'))[1]
								$LogonType = ([regex]::split($task.xml, '<LogonType>|</LogonType>'))[1]
								
								#convert state to status
								switch ($task.State)
								{
									0 {
										$Status = "Unknown"
									}
									1 {
										$Status = "Disabled"
									}
									2 {
										$Status = "Queued"
									}
									3 {
										$Status = "Ready"
									}
									4 {
										$Status = "Running"
									}
								}
								
								#output the task details
								if (-not $exclude -or $task.Path -notmatch $Exclude)
								{
									$task | Select-Object @{
										label	  = "ComputerName"; expression = {
											$computer
										}
									},
														  Name,
														  Path,
														  Enabled,
														  @{
										label   = "Action"; expression = {
											$Action
										}
									},
														  @{
										label	   = "Arguments"; expression = {
											$Arguments
										}
									},
														  @{
										label   = "UserId"; expression = {
											$UserId
										}
									},
														  LastRunTime,
														  NextRunTime,
														  @{
										label   = "Status"; expression = {
											$Status
										}
									},
														  @{
										label   = "Author"; expression = {
											$Author
										}
									},
														  @{
										label	  = "RunLevel"; expression = {
											$RunLevel
										}
									},
														  @{
										label	     = "Description"; expression = {
											$Description
										}
									},
														  NumberOfMissedRuns
									
									#if specified, output the results in importable XML format
									if ($path)
									{
										$xml = $task.Xml
										$taskname = $task.Name
										$xml | Out-File $(Join-Path $path "$computer-$taskname.xml")
									}
								}
							}
						}
					}
					catch
					{
						Write-Warning "Could not pull scheduled tasks from $computer using COM object, falling back to schtasks.exe"
						try
						{
							Get-SchTasks -computername $computer -folder $folder -ErrorAction stop
						}
						catch
						{
							Write-Error "Could not pull scheduled tasks from $computer using schtasks.exe:`n$_"
							continue
						}
					}
				}
				
				#otherwise, use schtasks
				else
				{
					
					try
					{
						Get-SchTasks -computername $computer -folder $folder -CompatibilityMode -ErrorAction stop
					}
					catch
					{
						Write-Error "Could not pull scheduled tasks from $computer using schtasks.exe:`n$_"
						continue
					}
				}
				
			}
			catch
			{
				Write-Error "Error pulling Scheduled tasks from $computer`: $_"
				continue
			}
		}
	}
}

function Invoke-QHEVScheduledTask
{
<#
	.SYNOPSIS
		A brief description of the Invoke-QHEVScheduledTask function.
	
	.DESCRIPTION
		A detailed description of the Invoke-QHEVScheduledTask function.
	
	.PARAMETER Action
		A description of the Action parameter.
	
	.EXAMPLE
		PS C:\> Invoke-QHEVScheduledTask -Action Start
	
	.NOTES
		Additional information about the function.
#>
	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		[ValidateSet('Start', 'Stop')]
		[String]$Action
	)
	begin
	{
		$global:ErrorActionPreference = [System.Management.Automation.ActionPreference]::Stop
		$computer = 'eve-admbtpp101'
	}
	process
	{
		try
		{
			if ($Action -eq 'Start')
			{
				$res = schtasks.exe /Run /S $Computer /TN 'Disable for EV'
			}
			elseif ($Action -eq 'Stop')
			{
				$res = schtasks.exe /End /S $Computer /TN 'Disable for EV'
			}
			
			$prop = [Ordered] @{
				Computer = $Computer
				Task	 = 'DisableEV'
				Action   = $Action
				Details  = $res
			}
		}
		catch
		{
			$prop = [Ordered] @{
				Computer = $Computer
				Task	 = 'DisableEV'
				Action   = 'ERROR'
				Details  = "$($_.Exception.Message)"
			}
		}
		finally
		{
			$obj = New-Object -TypeName System.Management.Automation.PSObject -Property $prop
			Write-Output $obj
		}
	}
	end
	{
		$global:ErrorActionPreference = [System.Management.Automation.ActionPreference]::Continue
	}
	
	
}

function Disable-QhEV
{
<#
	.SYNOPSIS
		A brief description of the Disable-QhEV function.
	
	.DESCRIPTION
		A detailed description of the Disable-QhEV function.
	
	.PARAMETER batchName
		A description of the batchName parameter.
	
	.EXAMPLE
		PS C:\> Disable-QhEV
	
	.NOTES
		Additional information about the function.
#>
	[CmdletBinding()]
	param
	(
		$batchName
	)
	
	begin
	{
		$global:ErrorActionPreference = [System.Management.Automation.ActionPreference]::Stop
		
		$targetDir = '\\eve-admbtpp101\d$\admin\Scripts\DisableEV'
		Write-Verbose "The Target Dir is $targetDir"
		
		$TargetFilePath = "$targetDir\MigrationReferenceList.csv"
		Write-Verbose  "The Target File Path is $TargetFilePath"
		
		$computer = 'eve-admbtpp101'
		
		Write-Verbose "Getting the current State of Scheduled job"
		
		$xState = Get-QHScheduledTasks -ComputerName $Computer | Where-Object {
			$_.Name -eq 'Disable for EV'
		}
		
		Write-Verbose "The current status of the Scheduled job is $($xState.Status)"
		
		if ($xState.Status -ne 'Ready')
		{
			Write-Host "WARNING : The task ['Disable for EV'] is not ready on $Computer. Please confirm the the task is not running and retry." -ForegroundColor Magenta
			break
		}
	}
	process
	{
		try
		{
			Write-Verbose "Navigating to Folder for $batchName"
			if (Navigate-QHMigrationFolder $batchName)
			{
				$parent = Get-Location | Select-Object -ExpandProperty path
				$refListName = "$($batchName)_ReferenceList.csv"
				$reflistPath = Join-Path -Path $parent -ChildPath $refListName
				Write-Verbose "Testing if $reflistPath Exists"
				if (Test-Path $reflistPath)
				{
					Write-Verbose "testing if $TargetFilePath Exists"
					if (Test-Path $TargetFilePath)
					{
						Write-Verbose "Renaming $TargetFilePath"
						Rename-Item $TargetFilePath -NewName "MigrationReferenceList_ReNamedOn_$((Get-Date).ToString('dd-MM-yyyy-hh-mm-ss')).csv" -Force
					}
					Write-Verbose "Getting Reference List"
					$csv = Import-Csv $reflistPath
					Write-Verbose "Saving Reference List at $TargetFilePath"
					$csv | Export-Csv -LiteralPath $TargetFilePath -NoTypeInformation -ErrorAction Stop
					
					Write-verbose "Attempting to Start Disable EV Scheduled task on [eve-admbtpp101]"
					Invoke-QHEVScheduledTask -Action Start
					Start-Sleep -Seconds 2
					while
					(
						(Get-QHScheduledTasks -ComputerName $Computer | Where-Object {
								($_.Name -eq 'Disable for EV') -and ($_.Status -eq 'Running')
							} | Measure-Object | Select-Object -ExpandProperty count) -ne 0
					)
					{
						Clear-Host
						$status = Get-QHScheduledTasks -ComputerName $Computer | Where-Object {
							$_.Name -eq 'Disable for EV'
						}
						Write-Host "Task Status" -ForegroundColor Cyan
						Write-Host "$($status | Out-String)" -ForegroundColor Yellow
						Start-Sleep -seconds 2
					}
					Write-Host "Task Status Completed"
					$status = Get-QHScheduledTasks -ComputerName $Computer | Where-Object {
						$_.Name -eq 'Disable for EV'
					}
					
					Write-Host "$($status | Out-String)" -ForegroundColor Green
					
					Start-Sleep -Seconds 2
					
					Write-Verbose "Getting the Log File from \\eve-admbtpp101\d$\admin\Scripts\DisableEV"
					
					$Log = Get-ChildItem '\\eve-admbtpp101\d$\admin\Scripts\DisableEV' -Filter evdisable*.log |
					Sort-Object -Property LastWriteTime | Select-Object -Last 1
					
					Write-Verbose "Saving EV Logfile"
					
					Copy-Item $Log.VersionInfo.FileName -Destination "$($batchName)_$($Log.Name)" -ErrorAction Stop
					
					#check if the last write is not today
					#$Log.LastAccessTime.Date.ToString('dd-MM-yyyy')
					#$today = Get-Date -Format 'dd-MM-yyyy'
				}
				else
				{
					Write-Host "ERROR : Reference List [$refList] Not Found. Please Ensure the Reference List with Legacy DN Attribut is Present." -ForegroundColor Magenta
					break
				}
			}
			else
			{
				Write-Host "Exiting Function" -ForegroundColor Magenta
			}
		}
		catch
		{
			Write-Host "ERROR : $($_.Exception.Message)" -ForegroundColor Magenta
		}
	}
	end
	{
		$global:ErrorActionPreference = [System.Management.Automation.ActionPreference]::Continue
	}
}

function Move-zQHSkypeUserToOkypeOnline
{
<#
	.SYNOPSIS
		A brief description of the Move-zQHSkypeUserToOkypeOnline function.
	
	.DESCRIPTION
		A detailed description of the Move-zQHSkypeUserToOkypeOnline function.
	
	.PARAMETER UserPrincipalName
		A description of the UserPrincipalName parameter.
	
	.PARAMETER BatchName
		A description of the BatchName parameter.
	
	.PARAMETER Credential
		A description of the Credential parameter.
	
	.PARAMETER ShowProgress
		A description of the ShowProgress parameter.
	
	.EXAMPLE
		PS C:\> Move-zQHSkypeUserToOkypeOnline
	
	.NOTES
		Additional information about the function.
#>
	[CmdletBinding()]
	param
	(
		[String[]]$UserPrincipalName,
		[String]$BatchName = 'None',
		[System.Management.Automation.Credential()]
		[ValidateNotNull()]
		[System.Management.Automation.PSCredential]$Credential = [System.Management.Automation.PSCredential]::Empty,
		[Switch]$ShowProgress
	)
	
	begin
	{
		$Global:ErrorActionPreference = [System.Management.Automation.ActionPreference]::Stop
		$i = 1
	}
	process
	{
		foreach ($UPN in $UserPrincipalName)
		{
			try
			{
				$recipient = Get-Recipient $UPN -ErrorAction Stop
				
				if ($recipient.RecipientTypeDetails -match 'User')
				{
					$start = Get-Date
					Move-CsUser -Identity $UPN -Target sipfed.online.lync.com -confirm:$false -ErrorAction Stop -WarningAction SilentlyContinue -Credential $Credential
					$end = Get-date
					$duration = Get-duration -StartTime $start -EndTime $end
					$prop = [Ordered]@{
						UserPrincipalName = $UPN
						BatchName		  = $BatchName
						EnvAdmin		  = $env:USERNAME
						SFBAdmin		  = $Credential.UserName
						RecipientTypeDetails = $recipient.RecipientTypeDetails
						Status		      = 'Success'
						Duration		  = $duration.Duration
						Details		      = 'None'
					}
				}
				else
				{
					$prop = [Ordered]@{
						UserPrincipalName = $UPN
						BatchName		  = $BatchName
						EnvAdmin		  = $env:USERNAME
						SFBAdmin		  = $Credential.UserName
						RecipientTypeDetails = $recipient.RecipientTypeDetails
						Status		      = 'Skipped'
						Duration		  = 'Skipped'
						Details		      = 'Skipped : Not a User Mailbox'
					}
				}
			}
			catch
			{
				$end = Get-date
				$duration = Get-duration -StartTime $start -EndTime $end
				$prop = [Ordered]@{
					UserPrincipalName = $UPN
					BatchName		  = $BatchName
					EnvAdmin		  = $env:USERNAME
					SFBAdmin		  = $Credential.UserName
					RecipientTypeDetails = $recipient.RecipientTypeDetails
					Status		      = 'Failed'
					Duration		  = $duration.Duration
					Details		      = "ERROR : $($_.Exception.Message)"
				}
			}
			finally
			{
				$obj = New-Object -TypeName System.Management.Automation.PSObject -Property $prop
				Write-Output $obj
				if ($ShowProgress)
				{
					if ($UserPrincipalName.count -gt 1)
					{
						$paramWriteProgress = @{
							Activity = 'Moving Skype users to Office 365'
							Status   = "Processing [$i] of [$($UserPrincipalName.Count)] users"
							PercentComplete = (($i / $UserPrincipalName.Count) * 100)
							CurrentOperation = "Completed : [$UPN]"
						}
						Write-Progress @paramWriteProgress
					}
				}
				
				$i++
			}
		}
	}
	end
	{
		Write-Progress -Activity 'Moving Skype users to Office 365' -Completed
		$global:ErrorActionPreference = [System.Management.Automation.ActionPreference]::Continue
	}
}

function Start-QhMailboxFolderCountOnPrem-Parallel
{
<#
	.SYNOPSIS
		A brief description of the Start-QhMailboxFolderCountOnPrem-Parallel function.
	
	.DESCRIPTION
		A detailed description of the Start-QhMailboxFolderCountOnPrem-Parallel function.
	
	.PARAMETER BulkUsers
		A description of the BulkUsers parameter.
	
	.PARAMETER ParellelSessions
		A description of the ParellelSessions parameter.
	
	.PARAMETER OutputFile
		A description of the OutputFile parameter.
	
	.EXAMPLE
		PS C:\> Start-QhMailboxFolderCountOnPrem-Parallel -BulkUsers 'value1' -ParellelSessions $ParellelSessions -OutputFile 'value3'
	
	.NOTES
		Additional information about the function.
#>
	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory = $true)]
		[String[]]$BulkUsers,
		[Parameter(Mandatory = $true)]
		[ValidateRange(1, 10)]
		[int]$ParellelSessions,
		[Parameter(Mandatory = $true)]
		[String]$OutputFile
	)
	
	begin
	{
		
	}
	process
	{
		$PreScript = {
			Import-Module QHO365MigrationOps -WarningAction SilentlyContinue
			
		}
		
		$ScriptBlock = {
			param (
				$users
			)
			
			get-QhMailboxFolderCountOnPrem -UserPrincipalName $users
		}
		
		if ($ParellelSessions -ne $null -and $BulkUsers.Count -gt 3)
		{
			$dataSet = Split-Array $BulkUsers -parts $ParellelSessions
			$Sub = 1
			foreach ($set in $dataSet)
			{
				$users = $set
				Start-Job -Name "MBXFolderCountOnPrem_Sub$($Sub)" -InitializationScript $PreScript -ScriptBlock $scriptBlock -ArgumentList $users
				$sub++
			}
			#$completed = $null
			while (@(Get-Job -Name "MBXFolderCountOnPrem_Sub*" | Where-Object {
						$_.State -eq "Running"
					}).Count -ne 0)
			{
				Clear-Host
				Write-Host "Please Wait While Jobs Complete : Completed - $((Get-job | Receive-job -keep).count)" -ForegroundColor Yellow
				$jobStatus = Get-job | Out-String
				Write-Host $jobStatus -ForegroundColor Cyan
				Start-Sleep -Seconds 15
			}
			Start-Sleep -Seconds 3
			$data = Get-job | Receive-Job -Keep
		}
	}
	end
	{
		$data | Export-Csv $OutputFile
		Get-Job | Remove-Job -Force
	}
}

function Start-QhMailboxFolderCountO365-Parallel
{
<#
	.SYNOPSIS
		A brief description of the Start-QhMailboxFolderCountO365-Parallel function.
	
	.DESCRIPTION
		A detailed description of the Start-QhMailboxFolderCountO365-Parallel function.
	
	.PARAMETER BulkUsers
		A description of the BulkUsers parameter.
	
	.PARAMETER ParellelSessions
		A description of the ParellelSessions parameter.
	
	.PARAMETER OutputFile
		A description of the OutputFile parameter.
	
	.EXAMPLE
		PS C:\> Start-QhMailboxFolderCountO365-Parallel -BulkUsers 'value1' -ParellelSessions $ParellelSessions -OutputFile 'value3'
	
	.NOTES
		Additional information about the function.
#>
	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory = $true)]
		[String[]]$BulkUsers,
		[Parameter(Mandatory = $true)]
		[ValidateRange(1, 10)]
		[int]$ParellelSessions,
		[Parameter(Mandatory = $true)]
		[String]$OutputFile
	)
	
	begin
	{
		
	}
	process
	{
		$PreScript = {
			Import-Module QHO365MigrationOps -WarningAction SilentlyContinue
			
		}
		
		$ScriptBlock = {
			param (
				$users
			)
			
			get-QhMailboxFolderCountO365 -UserPrincipalName $users
		}
		
		if ($ParellelSessions -ne $null -and $BulkUsers.Count -gt 3)
		{
			$dataSet = Split-Array $BulkUsers -parts $ParellelSessions
			$Sub = 1
			foreach ($set in $dataSet)
			{
				$users = $set
				Start-Job -Name "MBXFolderCountO365_Sub$($Sub)" -InitializationScript $PreScript -ScriptBlock $scriptBlock -ArgumentList $users
				$sub++
			}
			#$completed = $null
			while (@(Get-Job -Name "MBXFolderCountO365_Sub*" | Where-Object {
						$_.State -eq "Running"
					}).Count -ne 0)
			{
				Clear-Host
				Write-Host "Please Wait While Jobs Complete : Completed - $((Get-job | Receive-job -keep).count)" -ForegroundColor Yellow
				$jobStatus = Get-job | Out-String
				Write-Host $jobStatus -ForegroundColor Cyan
				Start-Sleep -Seconds 15
			}
			Start-Sleep -Seconds 3
			$data = Get-job | Receive-Job -Keep
		}
	}
	end
	{
		$data | Export-Csv $OutputFile
		Get-Job | Remove-Job -Force
	}
}

function Generate-QhMigrationBatchCSV
{
<#
	.SYNOPSIS
		A brief description of the Generate-QhMigrationBatchCSV function.
	
	.DESCRIPTION
		A detailed description of the Generate-QhMigrationBatchCSV function.
	
	.PARAMETER EmailAddresses
		A description of the EmailAddresses parameter.
	
	.PARAMETER Sets
		A description of the Sets parameter.
	
	.PARAMETER BatchName
		A description of the BatchName parameter.
	
	.EXAMPLE
		PS C:\> Generate-QhMigrationBatchCSV -EmailAddresses 'value1' -BatchName 'value2'
	
	.NOTES
		Additional information about the function.
#>
	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory = $true)]
		[String[]]$EmailAddresses,
		[Parameter(Mandatory = $false)]
		[ValidateSet('1', '2', '3', '4')]
		[Int]$Sets,
		[Parameter(Mandatory = $true)]
		[String]$BatchName
	)
	
	begin
	{
		
	}
	process
	{
		if ($sets -eq 0)
		{
			Write-Verbose "Sets not mentioned, Splitting $($EmailAddresses.Count) objects as per your defined logic"
			switch ($($EmailAddresses.Count))
			{
				{ $_ -le 500 }{ $sets = 1 }
				{ $_ -gt 500 -and $_ -le 1000 }{ $sets = 2 }
				{ $_ -gt 1000 }{ $sets = 4 }
				default { }
			}
		}
		else
		{
			Write-Verbose "Sets are user defined."
		}
		Write-Verbose "Splitting $($EmailAddresses.count) objects into $sets Parts"
		
		$datasets = Split-Array -inArray $EmailAddresses -parts $sets
		
		$i = 1
		foreach ($part in $datasets)
		{
			Write-Verbose "Processing: Set $i with $($part.count) objects"
			try
			{
				Write-Verbose "Exporting: Set $i at $($batchName)_MRS$($i).csv"
				$part | Select-Object @{
					n  = 'EmailAddress'; e = {
						$_
					}
				} |
				Export-Csv "$($batchName)_MRS$($i).csv" -NoTypeInformation -ErrorAction Stop
				Write-Verbose "Processed set $i"
				$i++
			}
			catch
			{
				Write-Host "ERROR : $($_.Exception.Message)" -ForegroundColor Magenta
				break
			}
		}
	}
	end
	{
	}
}

function get-LastADSync
{
<#
	.SYNOPSIS
		A brief description of the get-LastADSync function.
	
	.DESCRIPTION
		A detailed description of the get-LastADSync function.
	
	.EXAMPLE
		PS C:\> get-LastADSync
	
	.NOTES
		Additional information about the function.
#>
	param (
	)
	begin
	{
		
	}
	process
	{
		try
		{
			$MsolInfo = Get-MsolCompanyInformation -ErrorAction Stop
			$lastSync = $msolInfo.LastDirSyncTime.ToLocalTime()
			$now = (get-date -ErrorAction Stop).ToLocalTime()
			$Duration = $now - $LastSync
			$durationMinsRaw = $duration.TotalMinutes
			$durationMins = [math]::Round($durationMinsRaw)
			
			$obj = [PsCustomObject][Ordered]@{
				CurrentTime	    = $now
				LastAADSyncTime = $lastSync
				TimeElapsed	    = "$durationMins Mins Ago"
				NextScheduledAADSync = $lastSync.AddMinutes(30)
				TimeToNextAADSync = "$([math]::Round(($lastSync.AddMinutes(30) - $now).TotalMinutes)) Mins Remaining"
				
			}
			Write-Output $obj
		}
		catch
		{
			#$ErrorMsg = "ERROR : $($MyInvocation.InvocationName) `t`t$($error[0].Exception.Message)"
			Write-Host "ERROR : $($_.Exception.Message)" -ForegroundColor Magenta
		}
	}
	end
	{
		
	}
}

function Get-QHCreds
{
<#
	.SYNOPSIS
		A brief description of the Get-QHCreds function.
	
	.DESCRIPTION
		A detailed description of the Get-QHCreds function.
	
	.PARAMETER UserName
		A description of the UserName parameter.
	
	.PARAMETER Password
		A description of the Password parameter.
	
	.EXAMPLE
		PS C:\> Get-QHCreds
	
	.NOTES
		Additional information about the function.
#>
	param (
		[String]$UserName,
		[String]$Password
	)
	try
	{
		$Global:ErrorActionPreference = 'Stop'
		$pass = "$Password" | ConvertTo-SecureString -asPlainText -Force
		$cred = New-Object System.Management.Automation.PSCredential($UserName, $Pass)
		$Global:ErrorActionPreference = 'Continue'
		return $cred
	}
	catch
	{
		Write-host "ERROR : $($_.Exception.Message)" -ForegroundColor Magenta
		return $null
	}
	
}

function Start-QHPostMigrationTasks-Parallel
{
<#
	.SYNOPSIS
		A brief description of the Start-QHPostMigrationTasks-Parallel function.
	
	.DESCRIPTION
		A detailed description of the Start-QHPostMigrationTasks-Parallel function.
	
	.PARAMETER batchName
		A description of the batchName parameter.
	
	.PARAMETER BulkUsers
		A description of the BulkUsers parameter.
	
	.PARAMETER ParellelSessions
		A description of the ParellelSessions parameter.
	
	.PARAMETER credCsv
		A description of the credCsv parameter.
	
	.EXAMPLE
		PS C:\> Start-QHPostMigrationTasks-Parallel
	
	.NOTES
		Additional information about the function.
#>
	[CmdletBinding()]
	param
	(
		$batchName,
		[String[]]$BulkUsers,
		[ValidateRange(1, 50)]
		[int]$ParellelSessions,
		[String]$credCsv
	)
	
	begin
	{
		try
		{
			$Creds = Import-Csv $credCsv -ErrorAction Stop | Select-Object -first $ParellelSessions
		}
		catch
		{
			Write-Host "Error : $($_.Exception.Message)" -ForegroundColor Magenta
			break
		}
	}
	process
	{
		$PreScript = {
			Import-Module QHO365MigrationOps -WarningAction SilentlyContinue # module which contains the functions.
			<#
			$Exservers = @(
				'EXC-CHMDC2P002'
				'EXC-CHMDC1P004'
				'exc-casbk7p003'
				'EXC-CHMNDCP001'
				'EXC-CHMDC2P001'
				'EXC-JRNDC1P001'
				'exc-casbk7p004'
				'exc-casbk7p005'
				'exc-casbk7p002'
				'EXC-CHMDC1P012'
				'EXC-CHMDC1P010'
				'EXC-JRNDC2P001'
				'EXC-CHMDC2P007'
				'exc-casbtpp006'
				'EXC-CHMDC2P006'
				'EXC-CHMDC2P009'
				'EXC-CHMDC2P004'
				'EXC-CHMDC2P010'
				'EXC-JRNDC2P002'
				'EXC-CHMDC2P003'
				'EXC-CHMDC1P002'
				'EXC-CHMDC2P011'
				'EXC-CHMDC2P008'
				'EXC-CHMDC2P005'
				'EXC-JRNDC1P002'
				'EXC-CHMDC2P012'
				'EXC-CHMDC1P008'
				'EXC-CHMDC1P003'
				'EXC-CHMDC1P005'
				'EXC-CHMDC1P009'
			)
			#>
			
			$ExServers = ${D:\Office365\Migrations\Batch\WorkingExchangeServers.txt}
			Connect-QHOnpremExchange -Server ($Exservers | Get-Random)
			
		}
		
		$ScriptBlock = {
			param (
				[String[]]$users,
				[String]$batchName,
				[pscredential]$cred
			)
			
			if (Navigate-QHMigrationFolder $batchName)
			{
				Connect-QhO365 -Credential $cred
				Start-QHPostMigrationTasks -UserPrincipalName $users -BatchName $batchName -AdminUserName $cred.UserName
			}
			
			Get-PSSession | Remove-PSSession
		}
		if ($ParellelSessions -ne $null -and $BulkUsers.Count -gt 3)
		{
			$dataSet = Split-Array $BulkUsers -parts $ParellelSessions
			$Sub = 0
			foreach ($set in $dataSet)
			{
				$user = $creds[$Sub].UserPrincipalName
				$password = $creds[$Sub].AppPassword
				
				$cred = Get-QHCreds -UserName $user -Password $password
				$users = $set
				Start-Job -Name "$($batchName)_PostMigrationTasks_Sub$($Sub)" -InitializationScript $PreScript -ScriptBlock $scriptBlock -ArgumentList $users, $batchName, $cred
				$sub++
			}
			#$completed = $null
			while (@(Get-Job -Name "$($batchName)*" | Where-Object {
						$_.State -eq "Running"
					}).Count -ne 0)
			{
				Clear-Host
				Write-Host "Please Wait While Jobs Complete : Completed - $((Get-job | Receive-job -keep).count)" -ForegroundColor Yellow
				$jobStatus = Get-job | Out-String
				Write-Host $jobStatus -ForegroundColor Cyan
				Start-Sleep -Seconds 5
			}
			Start-Sleep -Seconds 3
			
			$jobStatus = Get-job | Out-String
			
			Write-Host "All Jobs Completed - $((Get-job | Receive-job -keep).count)" -ForegroundColor Green
			$jobStatus = Get-job | Out-String
			
			Write-Host $jobStatus -ForegroundColor Green
			
			
			$data = Get-job | Receive-Job -Keep
			
			Navigate-QHMigrationFolder $batchName
			
			Write-Output $data | Export-Csv "$($batchName)_PostMigrationTasks.csv" -NoTypeInformation
			
		}
	}
	end
	{
		
	}
}

function Move-xQHUsersToSFB-Parallel
{
<#
	.SYNOPSIS
		A brief description of the Move-xQHUsersToSFB-Parallel function.
	
	.DESCRIPTION
		A detailed description of the Move-xQHUsersToSFB-Parallel function.
	
	.PARAMETER batchName
		A description of the batchName parameter.
	
	.PARAMETER BulkUsers
		A description of the BulkUsers parameter.
	
	.PARAMETER ParellelSessions
		A description of the ParellelSessions parameter.
	
	.PARAMETER credCsv
		A description of the credCsv parameter.
	
	.EXAMPLE
		PS C:\> Move-xQHUsersToSFB-Parallel
	
	.NOTES
		Additional information about the function.
#>
	[CmdletBinding()]
	param
	(
		[String]$batchName,
		[String[]]$BulkUsers,
		[ValidateRange(1, 20)]
		[int]$ParellelSessions,
		[String]$credCsv
	)
	
	begin
	{
		try
		{
			$Creds = Import-Csv $credCsv -ErrorAction Stop | Select-Object -first $ParellelSessions
		}
		catch
		{
			Write-Host "Error : $($_.Exception.Message)" -ForegroundColor Magenta
			break
		}
	}
	process
	{
		$PreScript = {
			$Exservers = ('exc-casbtpp004', 'exc-casbtpp003', 'exc-casbtpp002', 'exc-casbtpp005', 'exc-chmndcp001')
			Connect-QHOnpremExchange -Server ($Exservers | Get-Random)
			Import-Module QHO365MigrationOps -WarningAction SilentlyContinue # module which contains the functions.
			Import-Module SkypeForBusiness
		}
		
		$ScriptBlock = {
			param (
				[String[]]$users,
				[String]$batchName,
				[pscredential]$cred
			)
			
			if (Navigate-QHMigrationFolder $batchName)
			{
				#Connect-QhO365 -Credential $cred
				#Connect-QhSkypeOnline -Credential $cred
				Move-zQHSkypeUserToOkypeOnline -UserPrincipalName $users -BatchName $batchName -Credential $cred
			}
		}
		if ($ParellelSessions -ne $null -and $BulkUsers.Count -gt 3)
		{
			$dataSet = Split-Array $BulkUsers -parts $ParellelSessions
			$Sub = 0
			foreach ($set in $dataSet)
			{
				$user = $creds[$Sub].UserPrincipalName
				$password = $creds[$Sub].Password
				
				$cred = Get-QHCreds -UserName $user -Password $password
				$users = $set
				Start-Job -Name "$($batchName)_MoveSkypeUsers_Sub$($Sub)" -InitializationScript $PreScript -ScriptBlock $scriptBlock -ArgumentList $users, $batchName, $cred
				$sub++
			}
			#$completed = $null
			$stopwatch = [system.diagnostics.stopwatch]::StartNew()
			while (@(Get-Job -Name "$($batchName)*" | Where-Object {
						$_.State -eq "Running"
					}).Count -ne 0)
			{
				Clear-Host
				Write-Host "Please Wait While Jobs Complete : Completed - $((Get-job | Receive-job -keep).count) ElapsedTime: $($stopwatch.Elapsed.Hours):$($stopwatch.Elapsed.Minutes):$($stopwatch.Elapsed.Seconds)" -ForegroundColor Yellow
				$jobStatus = Get-job | Out-String
				Write-Host $jobStatus -ForegroundColor Cyan
				Start-Sleep -Seconds 5
			}
			Start-Sleep -Seconds 3
			Write-Host "All Jobs Completed : Completed - $((Get-job | Receive-job -keep).count) ElapsedTime: $($stopwatch.Elapsed.Hours):$($stopwatch.Elapsed.Minutes):$($stopwatch.Elapsed.Seconds)" -ForegroundColor Yellow
			$jobStatus = Get-job | Out-String
			Write-Host $jobStatus -ForegroundColor Green
			$stopwatch.Stop()
			$data = Get-job | Receive-Job -Keep
			
			Navigate-QHMigrationFolder $batchName
			
			Write-Output $data | Export-Csv "$($batchName)_MovetoSkypeOnline.csv" -NoTypeInformation
			
		}
		
	}
	end
	{
		
	}
}

function Set-QHSFBSettings-Parallel
{
<#
	.SYNOPSIS
		A brief description of the Set-QHSFBSettings-Parallel function.
	
	.DESCRIPTION
		A detailed description of the Set-QHSFBSettings-Parallel function.
	
	.PARAMETER batchName
		A description of the batchName parameter.
	
	.PARAMETER BulkUsers
		A description of the BulkUsers parameter.
	
	.PARAMETER ParellelSessions
		A description of the ParellelSessions parameter.
	
	.PARAMETER credCsv
		A description of the credCsv parameter.
	
	.EXAMPLE
		PS C:\> Set-QHSFBSettings-Parallel
	
	.NOTES
		Additional information about the function.
#>
	[CmdletBinding()]
	param
	(
		[String]$batchName,
		[String[]]$BulkUsers,
		[ValidateRange(1, 20)]
		[int]$ParellelSessions,
		[String]$credCsv
	)
	
	begin
	{
		try
		{
			$Creds = Import-Csv $credCsv -ErrorAction Stop | Select-Object -first $ParellelSessions
		}
		catch
		{
			Write-Host "Error : $($_.Exception.Message)" -ForegroundColor Magenta
			break
		}
	}
	process
	{
		$PreScript = {
			Import-Module QHO365MigrationOps -WarningAction SilentlyContinue # module which contains the functions.
			<#$Exservers = @(
				'EXC-CHMDC2P002'
				'EXC-CHMDC1P004'
				'exc-casbk7p003'
				'EXC-CHMNDCP001'
				'EXC-CHMDC2P001'
				'EXC-JRNDC1P001'
				'exc-casbk7p004'
				'exc-casbk7p005'
				'exc-casbk7p002'
				'EXC-CHMDC1P012'
				'EXC-CHMDC1P010'
				'EXC-JRNDC2P001'
				'EXC-CHMDC2P007'
				'exc-casbtpp006'
				'EXC-CHMDC2P006'
				'EXC-CHMDC2P009'
				'EXC-CHMDC2P004'
				'EXC-CHMDC2P010'
				'EXC-JRNDC2P002'
				'EXC-CHMDC2P003'
				'EXC-CHMDC1P002'
				'EXC-CHMDC2P011'
				'EXC-CHMDC2P008'
				'EXC-CHMDC2P005'
				'EXC-JRNDC1P002'
				'EXC-CHMDC2P012'
				'EXC-CHMDC1P008'
				'EXC-CHMDC1P003'
				'EXC-CHMDC1P005'
				'EXC-CHMDC1P009'
			)#>
			$ExServers = ${D:\Office365\Migrations\Batch\WorkingExchangeServers.txt}
			Connect-QHOnpremExchange -Server ($Exservers | Get-Random)
			#Import-Module SkypeForBusiness
		}
		
		$ScriptBlock = {
			param (
				[String[]]$users,
				[String]$batchName,
				[pscredential]$cred
			)
			
			if (Navigate-QHMigrationFolder $batchName)
			{
				#Connect-QhO365 -Credential $cred
				Connect-QhSkypeOnline -Credential $cred
				Set-QHSkypeSettings -BatchName $batchName -UserPrincipalName $users -Cred $cred
			}
			Get-PSSession | Remove-PSSession
		}
		if ($ParellelSessions -ne $null -and $BulkUsers.Count -gt $ParellelSessions)
		{
			$dataSet = Split-Array $BulkUsers -parts $ParellelSessions
			$Sub = 0
			foreach ($set in $dataSet)
			{
				$user = $creds[$Sub].UserPrincipalName
				$password = $creds[$Sub].Password
				
				$cred = Get-QHCreds -UserName $user -Password $password
				$users = $set
				Start-Job -Name "$($batchName)_SkypeSettings_Sub$($Sub)" -InitializationScript $PreScript -ScriptBlock $scriptBlock -ArgumentList $users, $batchName, $cred
				$sub++
			}
			#$completed = $null
			$stopwatch = [system.diagnostics.stopwatch]::StartNew()
			while (@(Get-Job -Name "$($batchName)*" | Where-Object { $_.State -eq "Running" }).Count -ne 0)
			{
				Clear-Host
				Write-Host "Please Wait While Jobs Complete : Completed - $((Get-job | Receive-job -keep).count) ElapsedTime: $($stopwatch.Elapsed.Hours):$($stopwatch.Elapsed.Minutes):$($stopwatch.Elapsed.Seconds)" -ForegroundColor Yellow
				$jobStatus = Get-job | Out-String
				Write-Host $jobStatus -ForegroundColor Cyan
				Start-Sleep -Seconds 5
			}
			Start-Sleep -Seconds 3
			Write-Host "All Jobs Completed : Completed - $((Get-job | Receive-job -keep).count) ElapsedTime: $($stopwatch.Elapsed.Hours):$($stopwatch.Elapsed.Minutes):$($stopwatch.Elapsed.Seconds)" -ForegroundColor Yellow
			$jobStatus = Get-job | Out-String
			Write-Host $jobStatus -ForegroundColor Green
			$stopwatch.Stop()
			$data = Get-job | Receive-Job -Keep
			
			Navigate-QHMigrationFolder $batchName
			
			Write-Output $data | Export-Csv "$($batchName)_SkypeSettings.csv" -NoTypeInformation
		}
		
	}
	end
	{
		Get-Job | Remove-Job
	}
}

function Set-Generic1 #inProgress
{
	param (
		$SamAccountName
	)
	begin
	{
		
	}
	process
	{
		try
		{
			
		}
		catch
		{
			
		}
	}
	end
	{
		
	}
}

function Remove-QhAdGroupMembership #InProgress
{
<#
	.SYNOPSIS
		A brief description of the Remove-QhAdGroupMembership function.
	
	.DESCRIPTION
		A detailed description of the Remove-QhAdGroupMembership function.
	
	.PARAMETER SamAccountName
		A description of the SamAccountName parameter.
	
	.PARAMETER GroupToRemoveFrom
		A description of the GroupToRemoveFrom parameter.
	
	.EXAMPLE
		PS C:\> Remove-QhAdGroupMembership -SamAccountName $SamAccountName -GroupToRemoveFrom ACT-SaaS-O365-GenericAccountOfficeProPlusOnly
	
	.NOTES
		Additional information about the function.
#>
	param
	(
		[Parameter(Mandatory = $true)]
		$SamAccountName,
		[Parameter(Mandatory = $true)]
		[ValidateSet('ACT-SaaS-O365-GenericAccountOfficeProPlusOnly', 'ADM-O365-LIC-E3-OfficeProPlusOnly')]
		$GroupToRemoveFrom
	)
	
	try
	{
		$paramRemoveAdGroupMember = @{
			identity    = $GroupToRemoveFrom
			members	    = $SamAccountName
			confirm	    = $false
			ErrorAction = 'Stop'
		}
		
		Remove-AdGroupMember @paramRemoveAdGroupMember
		
		$prop = [Ordered] @{
			ADGroup	       = $GroupToRemoveFrom
			SamAccountName = $SamAccountName
			Removed	       = 'Success'
			Details	       = 'None'
		}
		
	}
	catch
	{
		$prop = [Ordered] @{
			ADGroup	       = $GroupToRemoveFrom
			SamAccountName = $SamAccountName
			Removed	       = 'Failed'
			Details	       = "ERROR : $($_.Exception.Message)"
		}
	}
	finally
	{
		$obj = New-Object -TypeName System.Management.Automation.PSObject -Property $prop
		Write-Output $obj
	}
}

function Get-QHDuplicateMailboxStatus
{
<#
	.SYNOPSIS
		A brief description of the Get-QHDuplicateMailboxStatus function.
	
	.DESCRIPTION
		A detailed description of the Get-QHDuplicateMailboxStatus function.
	
	.PARAMETER UserPrincipalName
		A description of the UserPrincipalName parameter.
	
	.PARAMETER ShowProgress
		A description of the ShowProgress parameter.
	
	.EXAMPLE
		PS C:\> Get-QHDuplicateMailboxStatus
	
	.NOTES
		Additional information about the function.
#>
	param (
		[String[]]$UserPrincipalName,
		[Switch]$ShowProgress
	)
	begin
	{
		$i = 1
		$Global:ErrorActionPreference = [System.Management.Automation.ActionPreference]::Stop
		
		try
		{
			$null = Get-AcceptedDomain -ErrorAction Stop
			$null = Get-ExoAcceptedDomain -ErrorAction Stop
		}
		catch
		{
			Write-Host "ERROR : $($_.Exception.Message). Make sure you are connected to Onpremise Exchange and Exchange Online" -ForegroundColor Magenta
			break
		}
	}
	process
	{
		foreach ($UPN in $UserPrincipalName)
		{
			try
			{
				$onprem = Get-Mailbox $UPN -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
				$Online = Get-ExoMailbox $UPN -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
				
				if ($onprem -ne $null)
				{
					$xOnprem = $true
				}
				else
				{
					$xOnprem = $false
				}
				
				if ($online -ne $null)
				{
					$xOnline = $true
				}
				else
				{
					$xOnline = $false
				}
				
				if ($Online -ne $null -and $onprem -ne $null)
				{
					$remediation = 'Needed'
					$details = "Mailbox Created on $($Online.WhenMailboxCreated)"
				}
				else
				{
					$remediation = 'NotNeeded'
					$details = 'None'
				}
				
				$prop = [ordered] @{
					UserPrincipalName = $UPN
					OnPrem		      = $xOnprem
					Online		      = $xOnline
					Remediation	      = $remediation
					Details		      = "$details"
				}
				
			}
			catch
			{
				$prop = [ordered] @{
					UserPrincipalName = $UPN
					OnPrem		      = 'Error'
					Online		      = 'Error'
					Remediation	      = 'Error'
					Details		      = "ERROR : $($_.Exception.Message)"
				}
			}
			finally
			{
				$obj = New-Object -TypeName System.Management.Automation.PSObject -Property $prop
				Write-Output $obj
				if ($UserPrincipalName.count -gt 1)
				{
					if ($ShowProgress)
					{
						$paramWriteProgress = @{
							Activity = 'Finding Duplicate Mailboxes'
							Status   = "Processing [$i] of [$($UserPrincipalName.Count)] users"
							PercentComplete = (($i / $UserPrincipalName.Count) * 100)
							CurrentOperation = "Completed : [$UPN]"
						}
						Write-Progress @paramWriteProgress
					}
				}
				$i++
			}
		}
	}
	end
	{
		Write-Progress -Activity 'Finding Duplicate Mailboxes' -Completed
		$Global:ErrorActionPreference = [System.Management.Automation.ActionPreference]::Continue
	}
	
}

function Open-QhModuleFolder
{
<#
	.SYNOPSIS
		A brief description of the Open-QhModuleFolder function.
	
	.DESCRIPTION
		A detailed description of the Open-QhModuleFolder function.
	
	.EXAMPLE
		PS C:\> Open-QhModuleFolder
	
	.NOTES
		Additional information about the function.
#>
	$mod = Get-Module -ListAvailable QHO365MigrationOps
	$path = Split-Path $mod.Path -Parent
	Start-Process $path
}

function Set-zQHLicenseWithoutExchange
{
<#
	.SYNOPSIS
		A brief description of the Set-zQHLicenseWithoutExchange function.
	
	.DESCRIPTION
		A detailed description of the Set-zQHLicenseWithoutExchange function.
	
	.PARAMETER UserPrincipalName
		A description of the UserPrincipalName parameter.
	
	.PARAMETER ShowProgress
		A description of the ShowProgress parameter.
	
	.EXAMPLE
		PS C:\> Set-zQHLicenseWithoutExchange
	
	.NOTES
		Additional information about the function.
#>
	[CmdletBinding()]
	param
	(
		[String[]]$UserPrincipalName,
		[Switch]$ShowProgress
	)
	
	begin
	{
		$global:ErrorActionPreference = [System.Management.Automation.ActionPreference]::Stop
		$i = 1
	}
	process
	{
		foreach ($UPN in $UserPrincipalName)
		{
			$EnabledPack = $null
			$disabledPack = $null
			$EnabledPack = @()
			$disabledPack = @()
			
			try
			{
				Write-Verbose "Trying to get $UPN Info"
				$recipient = Get-Recipient $UPN -ErrorAction Stop
				Write-Verbose "Retrived $UPN info"
				
				Write-Verbose "Trying to retrive licensing details for $UPN"
				$licinfo = Test-xQHIFLicenseEnabled -UserPrincipalName $UPN -LicenseSku E1, E3, PowerBi, ExchangeOnlinePlan2, MFA
				Write-Verbose "Retrived Licensing details for $UPN"
				if ($recipient.RecipientTypeDetails -match 'UserMailbox')
				{
					Write-Verbose "The $UPN Type is $($recipient.RecipientTypeDetails). Trying to Set the Usage Location to AU"
					Set-MsolUser -UserPrincipalName $UPN -UsageLocation "AU" -ErrorAction SilentlyContinue
					Write-Verbose "Usage Location is now set to AU"
					
					Write-Verbose "Checking the License Assignment"
					if (($licinfo.LicenseAssignment -eq 'ADM-O365-LIC-E3-OfficeProPlusOnly') -or ($licinfo.LicenseAssignment -eq 'ACT-SaaS-O365-GenericAccountOfficeProPlusOnly'))
					{
						Write-Verbose "License Assignment is via a group"
						
						Write-Verbose "Checking if the License assignment is via the generic Group (ACT-SaaS)"
						if ($licinfo.LicenseAssignment -eq 'ACT-SaaS-O365-GenericAccountOfficeProPlusOnly')
						{
							Write-Verbose "License Assignment is via Generic Group. Report Only"
							$Adrem = 'Report : ACT-SaaS-O365-GenericAccountOfficeProPlusOnly'
						}
						else
						{
							Write-Verbose "License Assignment is via (ADM-O365)"
							Write-Verbose "Attempting to remove (ADM-O365)"
							$Adrem = Remove-QHE3ADGrpmembership -AdGroup 'ADM-O365-LIC-E3-OfficeProPlusOnly' -SamAccountName $recipient.SamAccountName
						}
					}
					else
					{
						Write-Verbose "License Assignment is Direct"
						$Adrem = $licinfo.LicenseAssignment
					}
					
					Write-Verbose "Checking if E1 is Enabled"
					if ($licinfo.E1 -eq 'Enabled')
					{
						$disabledPack += 'healthqld:STANDARDPACK'
						Write-Verbose "E1 is enabled, added for Removal"
					}
					
					Write-Verbose "Checking if ExchangeOnline Plan 2 is Enabled"
					if ($licinfo.ExchangeOnlinePlan2 -eq 'Enabled')
					{
						$disabledPack += 'healthqld:EXCHANGEENTERPRISE'
						Write-Verbose "ExchangeOnline Plan 2 is enabled, added for Removal"
					}
					
					Write-Verbose "Checking if MFA license is Assigned"
					if ($licinfo.MFA -eq 'Disabled')
					{
						#Apply MFA License and Enable
						$EnabledPack += 'healthqld:MFA_STANDALONE'
						Write-Verbose "No MFA license, Added MFA for assignment"
					}
					else
					{
						$mfa = 'AlreadyLicensedforMFA'
						Write-Verbose "MFA License Exists, No action taken"
					}
					
					Write-Verbose "Checking if PowerBi license is Assigned"
					if ($licinfo.PowerBi -eq 'Disabled')
					{
						$EnabledPack += 'healthqld:POWER_BI_STANDARD'
						Write-Verbose "No PowerBi license, Added PowerBi for assignment"
					}
					
					Write-Verbose "Checking if E3 license is Assigned"
					if ($licinfo.E3 -eq 'Enabled')
					{
						Write-Verbose "Checking if E3 Assignment is via AD Group"
						if (($licinfo.LicenseAssignment -eq 'ADM-O365-LIC-E3-OfficeProPlusOnly') -or ($licinfo.LicenseAssignment -eq 'ACT-SaaS-O365-GenericAccountOfficeProPlusOnly'))
						{
							
							$EnabledPack += 'healthqld:ENTERPRISEPACK'
							$paramNewMsolLicenseOptions = @{
								AccountSkuId  = 'healthqld:ENTERPRISEPACK'
								DisabledPlans = 'EXCHANGE_S_ENTERPRISE'
								ErrorAction   = 'Stop'
							}
							
							$options = New-MsolLicenseOptions @paramNewMsolLicenseOptions
						}
						else
						{
							#UpDateLicense
							
							$paramNewMsolLicenseOptions = @{
								AccountSkuId  = 'healthqld:ENTERPRISEPACK'
								DisabledPlans = 'EXCHANGE_S_ENTERPRISE'
								ErrorAction   = 'Stop'
							}
							
							$options = New-MsolLicenseOptions @paramNewMsolLicenseOptions
						}
					}
					else
					{
						#addLicense
						$EnabledPack += 'healthqld:ENTERPRISEPACK'
					}
					
					
					$LicenceParam = @{
						UserPrincipalName = $UPN
						ErrorAction	      = 'Stop'
					}
					
					
					if ($EnabledPack -ne $null)
					{
						$LicenceParam.Add('AddLicenses', $EnabledPack)
						$add = $(($EnabledPack -join ',').Trim())
					}
					else
					{
						$add = 'None'
					}
					
					
					if ($disabledPack -ne $null)
					{
						$LicenceParam.Add('RemoveLicenses', $disabledPack)
						$removed = $(($disabledPack -join ',').Trim())
					}
					else
					{
						$removed = 'None'
					}
					
					
					if ($options)
					{
						$LicenceParam.Add('LicenseOptions', $options)
						$mod = 'E3:EnabledAllServices'
					}
					else
					{
						$mod = 'None'
					}
					
					
					Set-MsolUserLicense @LicenceParam
					#Write-Host "Giving 5 seconds to update the Changes" -ForegroundColor Cyan
					
					#Start-Sleep -Seconds 5
					
					
					if ($EnabledPack -icontains 'healthqld:MFA_STANDALONE')
					{
						$mfa = 'PostLicensing' #Enable-QHMFA -EmailAddress $UPN
					}
					
					$newLicInfo = Test-xQHIFLicenseEnabled -UserPrincipalName $UPN -LicenseSku E1, E3, PowerBi, ExchangeOnlinePlan2, MFA
					
					$prop = [ordered] @{
						User				 = $UPN
						RecipientTypeDetails = $recipient.RecipientTypeDetails
						Added			     = $add
						RemovedADmemberShip  = $removed
						Modified			 = $mod
						ADGrpRemoval		 = $Adrem
						MFAState			 = $mfa
						OldLicense		     = $(($licinfo | Select-Object E1, E3, ExchangeOnlinePlan2, MFA, PowerBi | Out-string).Trim())
						NewLicense		     = $(($Newlicinfo | Select-Object E1, E3, ExchangeOnlinePlan2, MFA, PowerBi | Out-String).Trim())
						Status			     = 'Success'
						Details			     = 'None'
					}
					
					#remove AdgroupmemberShip Function (may be at the top)
					
				}
				elseif ($recipient.RecipientTypeDetails -match 'SharedMailbox')
				{
					Set-MsolUser -UserPrincipalName $UPN -UsageLocation "AU" -ErrorAction SilentlyContinue
					$LicenceParam = @{
						UserPrincipalName = $UPN
						ErrorAction	      = 'Stop'
					}
					
					if ($licinfo.ExchangeOnlinePlan2 -eq 'Disabled')
					{
						$EnabledPack += 'healthqld:EXCHANGEENTERPRISE'
					}
					
					if ($licinfo.E3 -eq 'enabled')
					{
						$disabledPack += 'healthqld:ENTERPRISEPACK'
					}
					
					if ($EnabledPack -ne $null)
					{
						$LicenceParam.Add('AddLicenses', $EnabledPack)
						$add = $(($EnabledPack -join ',').Trim())
					}
					if ($disabledPack -ne $null)
					{
						$LicenceParam.Add('RemoveLicenses', $disabledPack)
						$removed = $(($disabledPack -join ',').Trim())
					}
					
					Set-MsolUserLicense @LicenceParam
					
					$newLicInfo = Test-xQHIFLicenseEnabled -UserPrincipalName $UPN -LicenseSku E1, E3, PowerBi, ExchangeOnlinePlan2, MFA
					
					$prop = [ordered] @{
						User				 = $UPN
						RecipientTypeDetails = $recipient.RecipientTypeDetails
						Added			     = $add
						RemovedADmemberShip  = $removed
						Modified			 = 'None'
						ADGrpRemoval		 = 'None'
						MFAState			 = 'None'
						OldLicense		     = $(($licinfo | Select-Object E1, E3, ExchangeOnlinePlan2, MFA, PowerBi | Out-string).Trim())
						NewLicense		     = $(($Newlicinfo | Select-Object E1, E3, ExchangeOnlinePlan2, MFA, PowerBi | Out-String).Trim())
						Status			     = 'Success'
						Details			     = 'None'
					}
				}
				else
				{
					$prop = [ordered] @{
						User				 = $UPN
						RecipientTypeDetails = $recipient.RecipientTypeDetails
						Added			     = 'SKIPPED'
						RemovedADmemberShip  = 'SKIPPED'
						Modified			 = 'SKIPPED'
						ADGrpRemoval		 = 'SKIPPED'
						MFAState			 = 'SKIPPED'
						OldLicense		     = 'SKIPPED'
						NewLicense		     = 'SKIPPED'
						Status			     = 'SKIPPED'
						Details			     = "SKIPPED : The [$UPN] is not an OnPremise Shared or User Mailbox"
					}
				}
			}
			catch
			{
				$prop = [ordered] @{
					User				 = $UPN
					RecipientTypeDetails = $recipient.RecipientTypeDetails
					Added			     = 'ERROR'
					RemovedADmemberShip  = 'ERROR'
					Modified			 = 'ERROR'
					ADGrpRemoval		 = 'ERROR'
					MFAState			 = 'ERROR'
					OldLicense		     = 'ERROR'
					NewLicense		     = 'ERROR'
					Status			     = 'Failed'
					Details			     = "ERROR : $($_.Exception.Message)"
				}
			}
			finally
			{
				$obj = New-Object -TypeName System.Management.Automation.PSObject -Property $prop
				Write-Output $obj
				if ($ShowProgress)
				{
					if ($UserPrincipalName.count -gt 1)
					{
						$paramWriteProgress = @{
							Activity = 'Assigning Licenses With out Exchange Online'
							Status   = "Processing [$i] of [$($UserPrincipalName.Count)] users"
							PercentComplete = (($i / $UserPrincipalName.Count) * 100)
							CurrentOperation = "Completed : [$UPN]"
						}
						Write-Progress @paramWriteProgress
						$i++
					}
				}
			}
		}
	}
	end
	{
		Write-Progress -Activity 'Assigning Licenses With out Exchange Online' -Completed
		$global:ErrorActionPreference = [System.Management.Automation.ActionPreference]::Continue
	}
}

function Set-zQHLicenseBeta-Parallel
{
<#
	.SYNOPSIS
		A brief description of the Set-zQHLicenseBeta-Parallel function.
	
	.DESCRIPTION
		A detailed description of the Set-zQHLicenseBeta-Parallel function.
	
	.PARAMETER batchName
		A description of the batchName parameter.
	
	.PARAMETER BulkUsers
		A description of the BulkUsers parameter.
	
	.PARAMETER ParellelSessions
		A description of the ParellelSessions parameter.
	
	.PARAMETER Cred
		A description of the Cred parameter.
	
	.EXAMPLE
		PS C:\> Set-zQHLicenseBeta-Parallel
	
	.NOTES
		Additional information about the function.
#>
	[CmdletBinding()]
	param
	(
		[String]$batchName,
		[String[]]$BulkUsers,
		[ValidateRange(1, 100)]
		[int]$ParellelSessions,
		[System.Management.Automation.Credential()]
		[ValidateNotNull()]
		[System.Management.Automation.PSCredential]$Cred = [System.Management.Automation.PSCredential]::Empty
	)
	
	begin
	{
		
	}
	process
	{
		$PreScript = {
			Import-Module QHO365MigrationOps -WarningAction SilentlyContinue # module which contains the functions.
			$Exservers = ${D:\Office365\Migrations\Batch\ExchangeCasServers.txt}
			Connect-QHOnpremExchange -Server ($Exservers | Get-Random)
			#Import-Module SkypeForBusiness
		}
		
		$ScriptBlock = {
			param (
				[String[]]$users,
				[String]$batchName,
				[System.Management.Automation.Credential()]
				[ValidateNotNull()]
				[System.Management.Automation.PSCredential]$Cred = [System.Management.Automation.PSCredential]::Empty
			)
			
			if (Navigate-QHMigrationFolder $batchName)
			{
				#Connect-QhO365 -Credential $cred
				Connect-QhMSOLService -Credential $cred
				Set-zQHLicenseBeta -UserPrincipalName $users
			}
		}
		if ($ParellelSessions -ne $null -and $BulkUsers.Count -gt 5)
		{
			$dataSet = Split-Array $BulkUsers -parts $ParellelSessions
			$Sub = 0
			foreach ($set in $dataSet)
			{
				$users = $set
				Start-Job -Name "$($batchName)_Licensing_Sub$($Sub)" -InitializationScript $PreScript -ScriptBlock $scriptBlock -ArgumentList $users, $batchName, $cred
				$sub++
			}
			#$completed = $null
			$stopwatch = [system.diagnostics.stopwatch]::StartNew()
			while (@(Get-Job -Name "$($batchName)*" | Where-Object { $_.State -eq "Running" }).Count -ne 0)
			{
				Clear-Host
				Write-Host "Please Wait While Jobs Complete : Completed - $((Get-job | Receive-job -keep).count) ElapsedTime: $($stopwatch.Elapsed.Hours):$($stopwatch.Elapsed.Minutes):$($stopwatch.Elapsed.Seconds)" -ForegroundColor Yellow
				$jobStatus = Get-job | Out-String
				Write-Host $jobStatus -ForegroundColor Cyan
				Start-Sleep -Seconds 5
			}
			Start-Sleep -Seconds 3
			Write-Host "All Jobs Completed : Completed - $((Get-job | Receive-job -keep).count) ElapsedTime: $($stopwatch.Elapsed.Hours):$($stopwatch.Elapsed.Minutes):$($stopwatch.Elapsed.Seconds)" -ForegroundColor Yellow
			$jobStatus = Get-job | Out-String
			Write-Host $jobStatus -ForegroundColor Green
			$stopwatch.Stop()
			$data = Get-job | Receive-Job -Keep
			
			Navigate-QHMigrationFolder $batchName
			
			Write-Output $data | Export-Csv "$($batchName)_Licensing.csv" -NoTypeInformation
		}
		
	}
	end
	{
		Get-Job | Remove-Job
	}
}

function Set-zEnableMFABeta-Parallel
{
<#
	.SYNOPSIS
		A brief description of the Set-zEnableMFABeta-Parallel function.
	
	.DESCRIPTION
		A detailed description of the Set-zEnableMFABeta-Parallel function.
	
	.PARAMETER batchName
		A description of the batchName parameter.
	
	.PARAMETER BulkUsers
		A description of the BulkUsers parameter.
	
	.PARAMETER ParellelSessions
		A description of the ParellelSessions parameter.
	
	.PARAMETER Cred
		A description of the Cred parameter.
	
	.EXAMPLE
		PS C:\> Set-zEnableMFABeta-Parallel
	
	.NOTES
		Additional information about the function.
#>
	[CmdletBinding()]
	param
	(
		[String]$batchName,
		[String[]]$BulkUsers,
		[ValidateRange(1, 100)]
		[int]$ParellelSessions,
		[System.Management.Automation.Credential()]
		[ValidateNotNull()]
		[System.Management.Automation.PSCredential]$Cred = [System.Management.Automation.PSCredential]::Empty
	)
	
	begin
	{
		
	}
	process
	{
		$PreScript = {
			Import-Module QHO365MigrationOps -WarningAction SilentlyContinue # module which contains the functions.
			$Exservers = ${D:\Office365\Migrations\Batch\ExchangeCasServers.txt}
			Connect-QHOnpremExchange -Server ($Exservers | Get-Random)
			#Import-Module SkypeForBusiness
		}
		
		$ScriptBlock = {
			param (
				[String[]]$users,
				[String]$batchName,
				[System.Management.Automation.Credential()]
				[ValidateNotNull()]
				[System.Management.Automation.PSCredential]$Cred = [System.Management.Automation.PSCredential]::Empty
			)
			
			if (Navigate-QHMigrationFolder $batchName)
			{
				#Connect-QhO365 -Credential $cred
				Connect-QhMSOLService -Credential $cred
				Enable-QHMFA -UserPrincipalName $users
			}
		}
		if ($ParellelSessions -ne $null -and $BulkUsers.Count -gt 5)
		{
			$dataSet = Split-Array $BulkUsers -parts $ParellelSessions
			$Sub = 0
			foreach ($set in $dataSet)
			{
				$users = $set
				Start-Job -Name "$($batchName)_MFA_Sub$($Sub)" -InitializationScript $PreScript -ScriptBlock $scriptBlock -ArgumentList $users, $batchName, $cred
				$sub++
			}
			#$completed = $null
			$stopwatch = [system.diagnostics.stopwatch]::StartNew()
			while (@(Get-Job -Name "$($batchName)*" | Where-Object { $_.State -eq "Running" }).Count -ne 0)
			{
				Clear-Host
				Write-Host "Please Wait While Jobs Complete : Completed - $((Get-job | Receive-job -keep).count) ElapsedTime: $($stopwatch.Elapsed.Hours):$($stopwatch.Elapsed.Minutes):$($stopwatch.Elapsed.Seconds)" -ForegroundColor Yellow
				$jobStatus = Get-job | Format-Table -AutoSize | Out-String
				Write-Host $jobStatus -ForegroundColor Cyan
				Start-Sleep -Seconds 5
			}
			Start-Sleep -Seconds 3
			Write-Host "All Jobs Completed : Completed - $((Get-job | Receive-job -keep).count) ElapsedTime: $($stopwatch.Elapsed.Hours):$($stopwatch.Elapsed.Minutes):$($stopwatch.Elapsed.Seconds)" -ForegroundColor Yellow
			$jobStatus = Get-job | Format-Table -AutoSize | Out-String
			Write-Host $jobStatus -ForegroundColor Green
			$stopwatch.Stop()
			$data = Get-job | Receive-Job -Keep
			
			Navigate-QHMigrationFolder $batchName
			
			Write-Output $data | Export-Csv "$($batchName)_MFA.csv" -NoTypeInformation
		}
		
	}
	end
	{
		Get-Job | Remove-Job
	}
}

function Get-QHMFAStatus-Parallel
{
<#
	.SYNOPSIS
		A brief description of the Get-QHMFAStatus-Parallel function.
	
	.DESCRIPTION
		A detailed description of the Get-QHMFAStatus-Parallel function.
	
	.PARAMETER batchName
		A description of the batchName parameter.
	
	.PARAMETER BulkUsers
		A description of the BulkUsers parameter.
	
	.PARAMETER ParellelSessions
		A description of the ParellelSessions parameter.
	
	.PARAMETER Cred
		A description of the Cred parameter.
	
	.EXAMPLE
		PS C:\> Get-QHMFAStatus-Parallel
	
	.NOTES
		Additional information about the function.
#>
	[CmdletBinding()]
	param
	(
		[String]$batchName,
		[String[]]$BulkUsers,
		[ValidateRange(1, 100)]
		[int]$ParellelSessions,
		[System.Management.Automation.Credential()]
		[ValidateNotNull()]
		[System.Management.Automation.PSCredential]$Cred = [System.Management.Automation.PSCredential]::Empty
	)
	
	begin
	{
		
	}
	process
	{
		$PreScript = {
			Import-Module QHO365MigrationOps -WarningAction SilentlyContinue # module which contains the functions.
			#$Exservers = ${D:\Office365\Migrations\Batch\ExchangeCasServers.txt}
			#Connect-QHOnpremExchange -Server ($Exservers | Get-Random)
			#Import-Module SkypeForBusiness
		}
		
		$ScriptBlock = {
			param (
				[String[]]$users,
				[String]$batchName,
				[System.Management.Automation.Credential()]
				[ValidateNotNull()]
				[System.Management.Automation.PSCredential]$Cred = [System.Management.Automation.PSCredential]::Empty
			)
			
			if (Navigate-QHMigrationFolder $batchName)
			{
				#Connect-QhO365 -Credential $cred
				Connect-QhMSOLService -Credential $cred
				Get-QHMFAStatus -UserPrincipalName $users
			}
		}
		if ($ParellelSessions -ne $null -and $BulkUsers.Count -gt 5)
		{
			$dataSet = Split-Array $BulkUsers -parts $ParellelSessions
			$Sub = 0
			foreach ($set in $dataSet)
			{
				$users = $set
				Start-Job -Name "$($batchName)_MfaStatus_Sub$($Sub)" -InitializationScript $PreScript -ScriptBlock $scriptBlock -ArgumentList $users, $batchName, $cred
				$sub++
			}
			#$completed = $null
			$stopwatch = [system.diagnostics.stopwatch]::StartNew()
			while (@(Get-Job -Name "$($batchName)*" | Where-Object { $_.State -eq "Running" }).Count -ne 0)
			{
				Clear-Host
				Write-Host "Please Wait While Jobs Complete : Completed - $((Get-job | Receive-job -keep).count) ElapsedTime: $($stopwatch.Elapsed.Hours):$($stopwatch.Elapsed.Minutes):$($stopwatch.Elapsed.Seconds)" -ForegroundColor Yellow
				$jobStatus = Get-job | Out-String
				Write-Host $jobStatus -ForegroundColor Cyan
				Start-Sleep -Seconds 5
			}
			Start-Sleep -Seconds 3
			Write-Host "All Jobs Completed : Completed - $((Get-job | Receive-job -keep).count) ElapsedTime: $($stopwatch.Elapsed.Hours):$($stopwatch.Elapsed.Minutes):$($stopwatch.Elapsed.Seconds)" -ForegroundColor Yellow
			$jobStatus = Get-job | Out-String
			Write-Host $jobStatus -ForegroundColor Green
			$stopwatch.Stop()
			$data = Get-job | Receive-Job -Keep
			
			Navigate-QHMigrationFolder $batchName
			
			Write-Output $data | Export-Csv "$($batchName)_MFAStatus.csv" -NoTypeInformation
		}
		
	}
	end
	{
		Get-Job | Remove-Job
	}
}

function Get-QHLicenseStatus
{
<#
	.SYNOPSIS
		A brief description of the Get-QHLicenseStatus function.
	
	.DESCRIPTION
		A detailed description of the Get-QHLicenseStatus function.
	
	.PARAMETER UserPrincipalName
		A description of the UserPrincipalName parameter.
	
	.PARAMETER ShowProgress
		A description of the ShowProgress parameter.
	
	.EXAMPLE
		PS C:\> Get-QHLicenseStatus -UserPrincipalName 'value1'
	
	.NOTES
		Additional information about the function.
#>
	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory = $true,
				   ValueFromPipeline = $true,
				   ValueFromPipelineByPropertyName = $true,
				   Position = 0,
				   HelpMessage = 'Enter UserPrincipal Name ?')]
		[String[]]$UserPrincipalName,
		[switch]$ShowProgress
	)
	
	begin
	{
		$i = 1
		$Global:ErrorActionPreference = [System.Management.Automation.ActionPreference]::Stop
	}
	process
	{
		foreach ($UPN in $UserPrincipalName)
		{
			try
			{
				$msol = Get-MsolUser -UserPrincipalName $UPN -ErrorAction Stop
				
				$lic = $msol.Licenses | Select-Object @{
					n = 'AccountSku'; e = {
						switch ($_ | Select-Object AccountSkuid | Select-Object -ExpandProperty AccountSkuid)
						{
							'healthqld:ENTERPRISEPACK' { 'E3' }
							'healthqld:STANDARDPACK' { 'E1' }
							'healthqld:MFA_STANDALONE' { 'MFA' }
							'healthqld:POWER_BI_STANDARD' { 'PowerBi Free' }
							'healthqld:EXCHANGEENTERPRISE' { 'ExchangeOnlinePlan2' }
							'healthqld:EMS' { 'Enterprise Mobility Security' }
							'healthqld:FLOW_FREE' { 'Microsoft Flow Free' }
							'healthqld:POWERAPPS_INDIVIDUAL_USER' { 'PowerApps and Logic Flows' }
							'healthqld:MCOEV' { 'Phone System' }
							'healthqld:POWER_BI_PRO' { 'PowerBi Pro' }
							'healthqld:POWER_BI_ADDON' { 'Power BI for Office 365 Add-On' }
							'healthqld:POWER_BI_INDIVIDUAL_USER' { 'Power BI Individual User' }
							'healthqld:ENTERPRISEWITHSCAL' { 'Enterprise Plan E4' }
							'healthqld:PROJECTONLINE_PLAN_1' { 'Project Online' }
							'healthqld:PROJECTCLIENT' { 'Project Pro for Office 365' }
							'healthqld:VISIOCLIENT' { 'Visio Pro Online' }
							'healthqld:STREAM' { 'Microsoft Stream' }
							'healthqld:POWERAPPS_VIRAL' { 'Microsoft Power Apps & Flow' }
							'healthqld:PROJECTESSENTIALS' { 'Project Lite' }
							'healthqld:PROJECTPROFESSIONAL' { 'Project Professional' }
							'healthqld:SPZA_IW' { 'App Connect' }
							'healthqld:PBI_PREMIUM_P1_ADDON' { 'Power Bi Premium' }
							'healthqld:DYN365_ENTERPRISE_P1_IW' { 'Dynamics 365 P1 Trial for Information Workers' }
							default { "$_" }
						}
					}
				},
													  @{
					n							    = 'Assignment'; e = {
						$_.GroupsAssigningLicense.Guid | ForEach-Object {
							if ($_ -match $msol.ObjectId.Guid) { "Direct" }
							elseif ($_ -eq $null) { "Direct:NoGUID" }
							else { $(Get-MsolGroup -ObjectId $_ | Select-Object -ExpandProperty Displayname) }
						}
					}
				}
				
				$prop = [ordered]@{
					UserPrincipalName = $UPN
				}
				foreach ($item in $lic)
				{
					$prop.Add($item.AccountSku, $item.Assignment)
					
				}
				$prop.Add('Details', 'None')
			}
			catch
			{
				$prop = [ordered]@{
					UserPrincipalName = $UPN
					Details		      = "ERROR : $($_.Exception.Message)"
				}
			}
			finally
			{
				$obj = New-Object -TypeName System.Management.Automation.PSObject -Property $prop
				Write-Output $obj
				
				if ($UserPrincipalName.count -gt 1)
				{
					if ($ShowProgress)
					{
						$paramWriteProgress = @{
							Activity = 'Getting MFA License Status'
							Status   = "Processing [$i] of [$($UserPrincipalName.Count)] users"
							PercentComplete = (($i / $UserPrincipalName.Count) * 100)
							CurrentOperation = "Completed : [$UPN]"
						}
						Write-Progress @paramWriteProgress
					}
				}
				$i++
			}
		}
	}
	end
	{
		Write-Progress -Activity 'Getting MFA License Status' -Completed
		$Global:ErrorActionPreference = [System.Management.Automation.ActionPreference]::Continue
	}
}

<#
# on Server VMs Having SFB Module
enable-wsmancredssp -role server -force
Restart-Service winrm
Set-PSSessionConfiguration -Name Microsoft.PowerShell -showSecurityDescriptorUI #add users for PS remotingaccess

#On Clients
enable-wsmancredssp -role client -delegatecomputer * -force
Enable-WSManCredSSP -Role client -DelegateComputer *.health.qld.gov.au

Restart-Service winrm

# Enabling Remoting and Credssp
Enable-PSRemoting -Force -Verbose
Enable-WSManCredSSP -Force Server -Verbose

# Enable Client CredSSP
Enable-WSManCredSSP -Force Client -DelegateComputer *


#>

function Set-iQHLicense
{
<#
	.SYNOPSIS
		A brief description of the Set-iQHLicense function.
	
	.DESCRIPTION
		A detailed description of the Set-iQHLicense function.
	
	.PARAMETER UserPrincipalName
		A description of the UserPrincipalName parameter.
	
	.PARAMETER ShowProgress
		A description of the ShowProgress parameter.
	
	.PARAMETER disabledServices
		A description of the disabledServices parameter.
	
	.EXAMPLE
		PS C:\> Set-iQHLicense
	
	.NOTES
		Additional information about the function.
#>
	[CmdletBinding()]
	param
	(
		[String[]]$UserPrincipalName,
		[Switch]$ShowProgress,
		[ValidateSet('E3:ExchangeOnline', 'E3:SkypeForBusinessOnline', 'E3:SharepointOnline')]
		[String[]]$disabledServices
	)
	
	begin
	{
		$global:ErrorActionPreference = [System.Management.Automation.ActionPreference]::Stop
		$i = 1
		
		$xdisabledServices = switch ($disabledServices)
		{
			'E3:ExchangeOnline' {
				'EXCHANGE_S_ENTERPRISE'
			}
			'E3:SkypeForBusinessOnline' {
				'MCOSTANDARD'
			}
			'E3:SharepointOnline' {
				'SHAREPOINTENTERPRISE', 'SHAREPOINTWAC'
			}
			default
			{
				$_
			}
		}
		
	}
	process
	{
		foreach ($UPN in $UserPrincipalName)
		{
			$EnabledPack = @()
			$disabledPack = @()
			$licinfo = $null
			$Newlicinfo = $null
			
			try
			{
				Write-Verbose "Trying to get $UPN Info"
				$recipient = Get-Recipient $UPN -ErrorAction Stop
				Write-Verbose "Retrived $UPN info"
				
				if ($recipient.RecipientTypeDetails -match 'UserMailbox')
				{
					Write-Verbose "Trying to retrive licensing details for $UPN"
					$licinfo = Get-QHLicenseStatus -UserPrincipalName $UPN
					Write-Verbose "Retrived Licensing details for $UPN"
					
					Write-Verbose "The $UPN Type is $($recipient.RecipientTypeDetails). Trying to Set the Usage Location to AU"
					Set-MsolUser -UserPrincipalName $UPN -UsageLocation "AU" -ErrorAction SilentlyContinue
					Write-Verbose "Usage Location is now set to AU"
					
					#check for Plans to Disable
					
					Write-Verbose "Current License as Follows"
					
					Write-Verbose "$($licinfo | Out-String)"
					
					if ($licinfo.E1 -ne $null)
					{
						$disabledPack += 'healthqld:STANDARDPACK'
					}
					
					if ($licinfo.ExchangeOnlinePlan2 -ne $null)
					{
						$disabledPack += 'healthqld:EXCHANGEENTERPRISE'
					}
					
					#Check for Enabled Plans
					
					if ($licinfo.'PowerBi Free' -eq $null)
					{
						$EnabledPack += 'healthqld:POWER_BI_STANDARD'
					}
					
					#E3 Ops and Checks
					
					if ($licinfo.E3 -eq $null)
					{
						$EnabledPack += 'healthqld:ENTERPRISEPACK'
						$Adrem = 'NoProPlusGrpMember'
						
						if ($licinfo.MFA -eq $null)
						{
							$EnabledPack += 'healthqld:MFA_STANDALONE'
						}
						
						if ($disabledServices -ne $null)
						{
							$paramNewMsolLicenseOptions = @{
								AccountSkuId  = 'healthqld:ENTERPRISEPACK'
								ErrorAction   = 'Stop'
								DisabledPlans = $xdisabledServices
							}
						}
						$options = New-MsolLicenseOptions @paramNewMsolLicenseOptions
						
					}
					else
					{
						if ($licinfo.E3 -icontains 'ADM-O365-LIC-E3-OfficeProPlusOnly')
						{
							#removes Group Membership
							$paramRemoveQHE3ADGrpmembership = @{
								AdGroup	       = 'ADM-O365-LIC-E3-OfficeProPlusOnly'
								SamAccountName = $recipient.SamAccountName
							}
							
							$Adrem = Remove-QHE3ADGrpmembership @paramRemoveQHE3ADGrpmembership
						}
						else
						{
							$Adrem = 'NoProPlusGrpMember'
						}
						
						if ($licinfo.E3 -icontains 'ACT-SaaS-O365-GenericAccountOfficeProPlusOnly')
						{
							if ($licinfo.MFA -ne $null)
							{
								$disabledPack += 'healthqld:MFA_STANDALONE'
							}
							
							if ($licinfo.E3 -icontains 'Direct' -or $licinfo.E3 -icontains 'Direct:NoGUID')
							{
								#remove Direct Licensing
								$disabledPack += 'healthqld:ENTERPRISEPACK'
							}
							#Process Generic - Code Pending.
							
							$paramRemoveQHE3ADGrpmembershipGen = @{
								AdGroup	       = 'ACT-SaaS-O365-GenericAccountOfficeProPlusOnly'
								SamAccountName = $recipient.SamAccountName
							}
							
							$rgen = Remove-QHE3ADGrpmembership @paramRemoveQHE3ADGrpmembershipGen
							
							$agen = Add-QHGenericGroupMember -SamAccountName $recipient.SamAccountName -AdGroup ACT-SaaS-O365-GenericAccountOfficeProPlusAndMailbox
							
							$gen = "$rgen $agen"
						}
						else
						{
							$gen = $false
							
							if ($licinfo.MFA -eq $null)
							{
								$EnabledPack += 'healthqld:MFA_STANDALONE'
							}
							
							if ($licinfo.E3 -icontains 'Direct' -or $licinfo.E3 -icontains 'Direct:NoGUID')
							{
								if ($disabledServices -ne $null)
								{
									$paramNewMsolLicenseOptions = @{
										AccountSkuId  = 'healthqld:ENTERPRISEPACK'
										ErrorAction   = 'Stop'
										DisabledPlans = $xdisabledServices
									}
								}
								else
								{
									$paramNewMsolLicenseOptions = @{
										AccountSkuId = 'healthqld:ENTERPRISEPACK'
										ErrorAction  = 'Stop'
									}
								}
								
								$options = New-MsolLicenseOptions @paramNewMsolLicenseOptions
								
							}
							else
							{
								$EnabledPack += 'healthqld:ENTERPRISEPACK'
								
								if ($disabledServices -ne $null)
								{
									$paramNewMsolLicenseOptions = @{
										AccountSkuId  = 'healthqld:ENTERPRISEPACK'
										ErrorAction   = 'Stop'
										DisabledPlans = $xdisabledServices
									}
								}
								$options = New-MsolLicenseOptions @paramNewMsolLicenseOptions
							}
						}
					}
					
					$LicenceParam = @{
						UserPrincipalName = $UPN
						ErrorAction	      = 'Stop'
					}
					
					if ($EnabledPack -ne $null)
					{
						$LicenceParam.Add('AddLicenses', $EnabledPack)
						$add = $(($EnabledPack -join ',').Trim())
					}
					else
					{
						$add = 'None'
					}
					
					if ($disabledPack -ne $null)
					{
						$LicenceParam.Add('RemoveLicenses', $disabledPack)
						$removed = $(($disabledPack -join ',').Trim())
					}
					else
					{
						$removed = 'None'
					}
					
					if ($options)
					{
						$LicenceParam.Add('LicenseOptions', $options)
						$mod = "E3:EnabledAllServices"
					}
					else
					{
						$mod = 'None'
					}
					
					# Create Command Param
					
					Set-MsolUserLicense @LicenceParam
					
					$Newlicinfo = Get-QHLicenseStatus -UserPrincipalName $UPN
					
					Write-Verbose "New License as Follows"
					
					Write-Verbose "$($Newlicinfo | Out-String)"
					
					$prop = [ordered] @{
						UserPrincipalName    = $UPN
						RecipientTypeDetails = $recipient.RecipientTypeDetails
						EnabledPack		     = $add
						DisabledPack		 = $removed
						Modified			 = $mod
						IsGeneric		     = $gen
						ADGrpRemoval		 = $Adrem
						Status			     = 'Success'
						Details			     = 'None'
					}
					
					#remove AdgroupmemberShip Function (may be at the top)
					
				}
				elseif ($recipient.RecipientTypeDetails -match 'SharedMailbox')
				{
					$gen = 'NotApplicable'
					Set-MsolUser -UserPrincipalName $UPN -UsageLocation "AU" -ErrorAction SilentlyContinue
					
					$licinfo = Get-QHLicenseStatus -UserPrincipalName $UPN
					
					Write-Verbose "Current License as Follows"
					Write-Verbose "$($licinfo | Out-String)"
					
					$LicenceParam = @{
						UserPrincipalName = $UPN
						ErrorAction	      = 'Stop'
					}
					
					if ($licinfo.ExchangeOnlinePlan2 -eq $null)
					{
						$EnabledPack += 'healthqld:EXCHANGEENTERPRISE'
					}
					
					if ($licinfo.E3 -ne $Null)
					{
						if ($licinfo.E3 -icontains 'ADM-O365-LIC-E3-OfficeProPlusOnly')
						{
							$paramRemoveQHE3ADGrpmembership = @{
								AdGroup	       = 'ADM-O365-LIC-E3-OfficeProPlusOnly'
								SamAccountName = $recipient.SamAccountName
							}
							
							$Adrem = Remove-QHE3ADGrpmembership @paramRemoveQHE3ADGrpmembership
						}
						
						if ($licinfo.E3 -match 'Direct')
						{
							$disabledPack += 'healthqld:ENTERPRISEPACK'
							$Adrem = 'NoProPlusGrpMember'
						}
					}
					else
					{
						$Adrem = 'NoProPlusGrpMember'
					}
					
					if ($licinfo.E1 -ne $null)
					{
						$disabledPack += 'healthqld:STANDARDPACK'
					}
					
					if ($EnabledPack -ne $null)
					{
						$LicenceParam.Add('AddLicenses', $EnabledPack)
						$add = $(($EnabledPack -join ',').Trim())
					}
					else
					{
						$add = 'None'
					}
					
					if ($disabledPack -ne $null)
					{
						$LicenceParam.Add('RemoveLicenses', $disabledPack)
						$removed = $(($disabledPack -join ',').Trim())
					}
					else
					{
						$removed = 'None'
					}
					
					$mod = 'NotApplicable'
					
					Set-MsolUserLicense @LicenceParam
					
					$Newlicinfo = Get-QHLicenseStatus -UserPrincipalName $UPN
					
					Write-Verbose "New License as Follows"
					Write-Verbose "$($Newlicinfo | Out-String)"
					
					$prop = [ordered] @{
						UserPrincipalName    = $UPN
						RecipientTypeDetails = $recipient.RecipientTypeDetails
						EnabledPack		     = $add
						DisabledPack		 = $removed
						Modified			 = $mod
						IsGeneric		     = $gen
						ADGrpRemoval		 = $Adrem
						Status			     = 'Success'
						Details			     = 'None'
					}
				}
				else
				{
					$prop = [ordered] @{
						UserPrincipalName    = $UPN
						RecipientTypeDetails = $recipient.RecipientTypeDetails
						EnabledPack		     = 'Skipped'
						DisabledPack		 = 'Skipped'
						Modified			 = 'Skipped'
						IsGeneric		     = $gen
						ADGrpRemoval		 = 'Skipped'
						Status			     = 'SKIPPED'
						Details			     = "The recipient is not a User or a Shared Mailbox. it is $($recipient.RecipientTypeDetails)"
					}
				}
			}
			catch
			{
				$prop = [ordered] @{
					UserPrincipalName    = $UPN
					RecipientTypeDetails = $recipient.RecipientTypeDetails
					EnabledPack		     = $add
					DisabledPack		 = $removed
					Modified			 = $mod
					IsGeneric		     = $gen
					ADGrpRemoval		 = $Adrem
					Status			     = 'Failed'
					Details			     = "ERROR : $($_.Exception.Message)"
				}
			}
			finally
			{
				$obj = New-Object -TypeName System.Management.Automation.PSObject -Property $prop
				Write-Output $obj
				if ($ShowProgress)
				{
					if ($UserPrincipalName.count -gt 1)
					{
						$paramWriteProgress = @{
							Activity = 'Assigning Licenses'
							Status   = "Processing [$i] of [$($UserPrincipalName.Count)] users"
							PercentComplete = (($i / $UserPrincipalName.Count) * 100)
							CurrentOperation = "Completed : [$UPN]"
						}
						Write-Progress @paramWriteProgress
					}
					$i++
				}
			}
		}
	}
	end
	{
		Write-Progress -Activity 'Assigning Licenses' -Completed
		$global:ErrorActionPreference = [System.Management.Automation.ActionPreference]::Continue
	}
}

function Get-iQHLicenseStatus
{
<#
	.SYNOPSIS
		A brief description of the Get-iQHLicenseStatus function.
	
	.DESCRIPTION
		A detailed description of the Get-iQHLicenseStatus function.
	
	.PARAMETER UserPrincipalName
		A description of the UserPrincipalName parameter.
	
	.PARAMETER ShowProgress
		A description of the ShowProgress parameter.
	
	.EXAMPLE
		PS C:\> Get-iQHLicenseStatus -UserPrincipalName 'value1'
	
	.NOTES
		Additional information about the function.
#>
	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory = $true,
				   ValueFromPipeline = $true,
				   ValueFromPipelineByPropertyName = $true,
				   Position = 0,
				   HelpMessage = 'Enter UserPrincipal Name ?')]
		[String[]]$UserPrincipalName,
		[switch]$ShowProgress
	)
	
	begin
	{
		$i = 1
		$Global:ErrorActionPreference = [System.Management.Automation.ActionPreference]::Stop
		$Plans = Get-MsolAccountSku -ErrorAction Stop | Select-Object -ExpandProperty AccountSkuId
	}
	process
	{
		foreach ($UPN in $UserPrincipalName)
		{
			try
			{
				$msol = Get-MsolUser -UserPrincipalName $UPN -ErrorAction Stop
				
				$prop = [ordered]@{
					UserPrincipalName = $UPN
				}
				
				foreach ($Plan in $Plans)
				{
					$lic = $msol.Licenses.Where({ $Plan -eq ($_ | Select-Object AccountSkuId | Select-Object -ExpandProperty AccountSkuId) })
					if ($lic -ne $null)
					{
						$licobj = $lic | Select-Object @{
							n = 'AccountSku'; e = {
								switch ($lic | Select-Object AccountSkuid | Select-Object -ExpandProperty AccountSkuid)
								{
									'healthqld:ENTERPRISEPACK' { 'E3' }
									'healthqld:STANDARDPACK' { 'E1' }
									'healthqld:MFA_STANDALONE' { 'MFA' }
									'healthqld:POWER_BI_STANDARD' { 'PowerBi Free' }
									'healthqld:EXCHANGEENTERPRISE' { 'ExchangeOnlinePlan2' }
									'healthqld:EMS' { 'Enterprise Mobility Security' }
									'healthqld:FLOW_FREE' { 'Microsoft Flow Free' }
									'healthqld:POWERAPPS_INDIVIDUAL_USER' { 'PowerApps and Logic Flows' }
									'healthqld:MCOEV' { 'Phone System' }
									'healthqld:POWER_BI_PRO' { 'PowerBi Pro' }
									'healthqld:POWER_BI_ADDON' { 'Power BI for Office 365 Add-On' }
									'healthqld:POWER_BI_INDIVIDUAL_USER' { 'Power BI Individual User' }
									'healthqld:ENTERPRISEWITHSCAL' { 'Enterprise Plan E4' }
									'healthqld:PROJECTONLINE_PLAN_1' { 'Project Online' }
									'healthqld:PROJECTCLIENT' { 'Project Pro for Office 365' }
									'healthqld:VISIOCLIENT' { 'Visio Pro Online' }
									'healthqld:STREAM' { 'Microsoft Stream' }
									'healthqld:POWERAPPS_VIRAL' { 'Microsoft Power Apps & Flow' }
									'healthqld:PROJECTESSENTIALS' { 'Project Lite' }
									'healthqld:PROJECTPROFESSIONAL' { 'Project Professional' }
									'healthqld:SPZA_IW' { 'App Connect' }
									'healthqld:PBI_PREMIUM_P1_ADDON' { 'Power Bi Premium' }
									'healthqld:DYN365_ENTERPRISE_P1_IW' { 'Dynamics 365 P1 Trial for Information Workers' }
									default { "$_" }
								}
							}
						},
													   @{
							n								  = 'Assignment'; e = {
								$lic.GroupsAssigningLicense.Guid | ForEach-Object {
									if ($_ -match $msol.ObjectId.Guid) { "Direct" }
									elseif ($_ -eq $null) { "Direct:NoGUID" }
									else { $(Get-MsolGroup -ObjectId $_ | Select-Object -ExpandProperty Displayname) }
								}
							}
						}
					}
					else
					{
						$lic = switch ($Plan)
						{
							'healthqld:ENTERPRISEPACK' { 'E3' }
							'healthqld:STANDARDPACK' { 'E1' }
							'healthqld:MFA_STANDALONE' { 'MFA' }
							'healthqld:POWER_BI_STANDARD' { 'PowerBi Free' }
							'healthqld:EXCHANGEENTERPRISE' { 'ExchangeOnlinePlan2' }
							'healthqld:EMS' { 'Enterprise Mobility Security' }
							'healthqld:FLOW_FREE' { 'Microsoft Flow Free' }
							'healthqld:POWERAPPS_INDIVIDUAL_USER' { 'PowerApps and Logic Flows' }
							'healthqld:MCOEV' { 'Phone System' }
							'healthqld:POWER_BI_PRO' { 'PowerBi Pro' }
							'healthqld:POWER_BI_ADDON' { 'Power BI for Office 365 Add-On' }
							'healthqld:POWER_BI_INDIVIDUAL_USER' { 'Power BI Individual User' }
							'healthqld:ENTERPRISEWITHSCAL' { 'Enterprise Plan E4' }
							'healthqld:PROJECTONLINE_PLAN_1' { 'Project Online' }
							'healthqld:PROJECTCLIENT' { 'Project Pro for Office 365' }
							'healthqld:VISIOCLIENT' { 'Visio Pro Online' }
							'healthqld:STREAM' { 'Microsoft Stream' }
							'healthqld:POWERAPPS_VIRAL' { 'Microsoft Power Apps & Flow' }
							'healthqld:PROJECTESSENTIALS' { 'Project Lite' }
							'healthqld:PROJECTPROFESSIONAL' { 'Project Professional' }
							'healthqld:SPZA_IW' { 'App Connect' }
							'healthqld:PBI_PREMIUM_P1_ADDON' { 'Power Bi Premium' }
							'healthqld:DYN365_ENTERPRISE_P1_IW' { 'Dynamics 365 P1 Trial for Information Workers' }
							default { "$_" }
						}
						
						$licobj = [PScustomobject]@{
							AccountSku = $lic
							Assignment = $null
						}
					}
					
					$prop.Add($licobj.AccountSku, $licobj.Assignment)
					
				}
				$prop.Add('Details', 'None')
			}
			catch
			{
				$prop = [ordered]@{
					UserPrincipalName = $UPN
					Details		      = "ERROR : $($_.Exception.Message)"
				}
			}
			finally
			{
				$obj = New-Object -TypeName System.Management.Automation.PSObject -Property $prop
				Write-Output $obj
				
				if ($UserPrincipalName.count -gt 1)
				{
					if ($ShowProgress)
					{
						$paramWriteProgress = @{
							Activity = 'Getting MFA License Status'
							Status   = "Processing [$i] of [$($UserPrincipalName.Count)] users"
							PercentComplete = (($i / $UserPrincipalName.Count) * 100)
							CurrentOperation = "Completed : [$UPN]"
						}
						Write-Progress @paramWriteProgress
					}
				}
				$i++
			}
		}
	}
	end
	{
		Write-Progress -Activity 'Getting MFA License Status' -Completed
		$Global:ErrorActionPreference = [System.Management.Automation.ActionPreference]::Continue
	}
}

function Set-iiQHLicense
{
<#
	.SYNOPSIS
		A brief description of the Set-iiQHLicense function.
	
	.DESCRIPTION
		A detailed description of the Set-iiQHLicense function.
	
	.PARAMETER UserPrincipalName
		A description of the UserPrincipalName parameter.
	
	.PARAMETER ShowProgress
		A description of the ShowProgress parameter.
	
	.PARAMETER disabledServices
		A description of the disabledServices parameter.
	
	.EXAMPLE
		PS C:\> Set-iiQHLicense
	
	.NOTES
		Additional information about the function.
#>
	[CmdletBinding()]
	param
	(
		[String[]]$UserPrincipalName,
		[Switch]$ShowProgress,
		[ValidateSet('E3:ExchangeOnline', 'E3:SkypeForBusinessOnline', 'E3:SharepointOnline')]
		[String[]]$disabledServices
	)
	
	begin
	{
		$global:ErrorActionPreference = [System.Management.Automation.ActionPreference]::Stop
		$i = 1
		
		$DisabledPlans = switch ($disabledServices)
		{
			'E3:ExchangeOnline' {
				'EXCHANGE_S_ENTERPRISE'
			}
			'E3:SkypeForBusinessOnline' {
				'MCOSTANDARD'
			}
			'E3:SharepointOnline' {
				'SHAREPOINTENTERPRISE', 'SHAREPOINTWAC'
			}
			default
			{
				$_
			}
		}
		
	}
	process
	{
		foreach ($UPN in $UserPrincipalName)
		{
			$EnabledPack = @()
			$disabledPack = @()
			$licinfo = $null
			$Newlicinfo = $null
			
			try
			{
				Write-Verbose "Trying to get $UPN Info"
				$recipient = Get-Recipient $UPN -ErrorAction Stop
				Write-Verbose "Retrived $UPN info"
				
				if ($recipient.RecipientTypeDetails -match 'UserMailbox')
				{
					Write-Verbose "Trying to retrive licensing details for $UPN"
					$licinfo = Get-QHLicenseStatus -UserPrincipalName $UPN
					Write-Verbose "Retrived Licensing details for $UPN"
					
					Write-Verbose "The $UPN Type is $($recipient.RecipientTypeDetails). Trying to Set the Usage Location to AU"
					Set-MsolUser -UserPrincipalName $UPN -UsageLocation "AU" -ErrorAction SilentlyContinue
					Write-Verbose "Usage Location is now set to AU"
					
					#check for Plans to Disable
					
					Write-Verbose "Current License as Follows"
					
					Write-Verbose "$($licinfo | Out-String)"
					
					if ($licinfo.E1 -ne $null)
					{
						$disabledPack += 'healthqld:STANDARDPACK'
					}
					
					if ($licinfo.ExchangeOnlinePlan2 -ne $null)
					{
						$disabledPack += 'healthqld:EXCHANGEENTERPRISE'
					}
					
					#Check for Enabled Plans
					
					if ($licinfo.'PowerBi Free' -eq $null)
					{
						$EnabledPack += 'healthqld:POWER_BI_STANDARD'
					}
					
					#E3 Ops and Checks
					
					if ($licinfo.E3 -eq $null)
					{
						$EnabledPack += 'healthqld:ENTERPRISEPACK'
						$Adrem = 'NoProPlusGrpMember'
						
						if ($licinfo.MFA -eq $null)
						{
							$EnabledPack += 'healthqld:MFA_STANDALONE'
						}
						
						if ($disabledServices -ne $null)
						{
							$paramNewMsolLicenseOptions = @{
								AccountSkuId  = 'healthqld:ENTERPRISEPACK'
								ErrorAction   = 'Stop'
								DisabledPlans = $DisabledPlans
							}
						}
						$options = New-MsolLicenseOptions @paramNewMsolLicenseOptions
						
					}
					else
					{
						if ($licinfo.E3 -icontains 'ADM-O365-LIC-E3-OfficeProPlusOnly')
						{
							#removes Group Membership
							$paramRemoveQHE3ADGrpmembership = @{
								AdGroup	       = 'ADM-O365-LIC-E3-OfficeProPlusOnly'
								SamAccountName = $recipient.SamAccountName
							}
							
							$Adrem = Remove-QHE3ADGrpmembership @paramRemoveQHE3ADGrpmembership
						}
						else
						{
							$Adrem = 'NoProPlusGrpMember'
						}
						
						if ($licinfo.E3 -icontains 'ACT-SaaS-O365-GenericAccountOfficeProPlusOnly')
						{
							if ($licinfo.MFA -ne $null)
							{
								$disabledPack += 'healthqld:MFA_STANDALONE'
							}
							
							if ($licinfo.E3 -icontains 'Direct' -or $licinfo.E3 -icontains 'Direct:NoGUID')
							{
								#remove Direct Licensing
								$disabledPack += 'healthqld:ENTERPRISEPACK'
							}
							#Process Generic - Code Pending.
							
							$paramRemoveQHE3ADGrpmembershipGen = @{
								AdGroup	       = 'ACT-SaaS-O365-GenericAccountOfficeProPlusOnly'
								SamAccountName = $recipient.SamAccountName
							}
							
							$rgen = Remove-QHE3ADGrpmembership @paramRemoveQHE3ADGrpmembershipGen
							
							$agen = Add-QHGenericGroupMember -SamAccountName $recipient.SamAccountName -AdGroup ACT-SaaS-O365-GenericAccountOfficeProPlusAndMailbox
							
							$gen = "$rgen $agen"
						}
						else
						{
							$gen = $false
							
							if ($licinfo.MFA -eq $null)
							{
								$EnabledPack += 'healthqld:MFA_STANDALONE'
							}
							
							if ($licinfo.E3 -icontains 'Direct' -or $licinfo.E3 -icontains 'Direct:NoGUID')
							{
								if ($disabledServices -ne $null)
								{
									$paramNewMsolLicenseOptions = @{
										AccountSkuId  = 'healthqld:ENTERPRISEPACK'
										ErrorAction   = 'Stop'
										DisabledPlans = $DisabledPlans
									}
								}
								else
								{
									$paramNewMsolLicenseOptions = @{
										AccountSkuId  = 'healthqld:ENTERPRISEPACK'
										ErrorAction   = 'Stop'
										DisabledPlans = $null
									}
								}
								
								$options = New-MsolLicenseOptions @paramNewMsolLicenseOptions
								
							}
							else
							{
								$EnabledPack += 'healthqld:ENTERPRISEPACK'
								
								if ($disabledServices -ne $null)
								{
									$paramNewMsolLicenseOptions = @{
										AccountSkuId  = 'healthqld:ENTERPRISEPACK'
										ErrorAction   = 'Stop'
										DisabledPlans = $DisabledPlans
									}
								}
								$options = New-MsolLicenseOptions @paramNewMsolLicenseOptions
							}
						}
					}
					
					$LicenceParam = @{
						UserPrincipalName = $UPN
						ErrorAction	      = 'Stop'
					}
					
					if ($EnabledPack -ne $null)
					{
						$LicenceParam.Add('AddLicenses', $EnabledPack)
						$add = $(($EnabledPack -join ',').Trim())
					}
					else
					{
						$add = 'None'
					}
					
					if ($disabledPack -ne $null)
					{
						$LicenceParam.Add('RemoveLicenses', $disabledPack)
						$removed = $(($disabledPack -join ',').Trim())
					}
					else
					{
						$removed = 'None'
					}
					
					if ($options)
					{
						$LicenceParam.Add('LicenseOptions', $options)
						$mod = "E3:EnabledAllServices"
					}
					else
					{
						$mod = 'None'
					}
					
					# Create Command Param
					
					Set-MsolUserLicense @LicenceParam
					
					$Newlicinfo = Get-QHLicenseStatus -UserPrincipalName $UPN
					
					Write-Verbose "New License as Follows"
					
					Write-Verbose "$($Newlicinfo | Out-String)"
					
					$prop = [ordered] @{
						UserPrincipalName    = $UPN
						RecipientTypeDetails = $recipient.RecipientTypeDetails
						EnabledPack		     = $add
						DisabledPack		 = $removed
						Modified			 = $mod
						IsGeneric		     = $gen
						ADGrpRemoval		 = $Adrem
						Status			     = 'Success'
						Details			     = 'None'
					}
					
					#remove AdgroupmemberShip Function (may be at the top)
					
				}
				elseif ($recipient.RecipientTypeDetails -match 'SharedMailbox')
				{
					#$gen = 'NotApplicable'
					Set-MsolUser -UserPrincipalName $UPN -UsageLocation "AU" -ErrorAction SilentlyContinue
					
					$licinfo = Get-QHLicenseStatus -UserPrincipalName $UPN
					
					Write-Verbose "Current License as Follows"
					Write-Verbose "$($licinfo | Out-String)"
					
					$LicenceParam = @{
						UserPrincipalName = $UPN
						ErrorAction	      = 'Stop'
					}
					
					if ($licinfo.E3 -ne $Null)
					{
						if ($licinfo.E3 -icontains 'Direct' -or $licinfo.E3 -icontains 'Direct:NoGUID')
						{
							#remove Direct Licensing
							$disabledPack += 'healthqld:ENTERPRISEPACK'
						}
						
						if ($licinfo.E3 -icontains 'ADM-O365-LIC-E3-OfficeProPlusOnly')
						{
							$paramRemoveQHE3ADGrpmembership = @{
								AdGroup	       = 'ADM-O365-LIC-E3-OfficeProPlusOnly'
								SamAccountName = $recipient.SamAccountName
							}
							
							$Adrem = Remove-QHE3ADGrpmembership @paramRemoveQHE3ADGrpmembership
						}
						else
						{
							#$disabledPack += 'healthqld:ENTERPRISEPACK'
							$Adrem = 'NoProPlusGrpMember'
						}
						
						if ($licinfo.E3 -icontains 'ACT-SaaS-O365-GenericAccountOfficeProPlusOnly')
						{
							if ($licinfo.MFA -ne $null)
							{
								$disabledPack += 'healthqld:MFA_STANDALONE'
							}
							if ($licinfo.ExchangeOnlinePlan2 -ne $null)
							{
								$disabledPack += 'healthqld:EXCHANGEENTERPRISE'
							}
							
							#Process Generic - Code Pending.
							
							$paramRemoveQHE3ADGrpmembershipGen = @{
								AdGroup	       = 'ACT-SaaS-O365-GenericAccountOfficeProPlusOnly'
								SamAccountName = $recipient.SamAccountName
							}
							
							$rgen = Remove-QHE3ADGrpmembership @paramRemoveQHE3ADGrpmembershipGen
							
							$agen = Add-QHGenericGroupMember -SamAccountName $recipient.SamAccountName -AdGroup 'ACT-SaaS-O365-GenericAccountOfficeProPlusAndMailbox'
							
							$gen = "$rgen $agen"
						}
						else
						{
							if ($licinfo.ExchangeOnlinePlan2 -eq $null)
							{
								$EnabledPack += 'healthqld:EXCHANGEENTERPRISE'
							}
						}
					}
					else
					{
						$Adrem = 'NoProPlusGrpMember'
						
						if ($licinfo.ExchangeOnlinePlan2 -eq $null)
						{
							$EnabledPack += 'healthqld:EXCHANGEENTERPRISE'
						}
					}
					
					if ($licinfo.E1 -ne $null)
					{
						$disabledPack += 'healthqld:STANDARDPACK'
					}
					
					if ($EnabledPack -ne $null)
					{
						$LicenceParam.Add('AddLicenses', $EnabledPack)
						$add = $(($EnabledPack -join ',').Trim())
					}
					else
					{
						$add = 'None'
					}
					
					if ($disabledPack -ne $null)
					{
						$LicenceParam.Add('RemoveLicenses', $disabledPack)
						$removed = $(($disabledPack -join ',').Trim())
					}
					else
					{
						$removed = 'None'
					}
					
					$mod = 'NotApplicable'
					
					Set-MsolUserLicense @LicenceParam
					
					$Newlicinfo = Get-QHLicenseStatus -UserPrincipalName $UPN
					
					Write-Verbose "New License as Follows"
					Write-Verbose "$($Newlicinfo | Out-String)"
					
					$prop = [ordered] @{
						UserPrincipalName    = $UPN
						RecipientTypeDetails = $recipient.RecipientTypeDetails
						EnabledPack		     = $add
						DisabledPack		 = $removed
						Modified			 = $mod
						IsGeneric		     = $gen
						ADGrpRemoval		 = $Adrem
						Status			     = 'Success'
						Details			     = 'None'
					}
				}
				else
				{
					$prop = [ordered] @{
						UserPrincipalName    = $UPN
						RecipientTypeDetails = $recipient.RecipientTypeDetails
						EnabledPack		     = 'Skipped'
						DisabledPack		 = 'Skipped'
						Modified			 = 'Skipped'
						IsGeneric		     = $gen
						ADGrpRemoval		 = 'Skipped'
						Status			     = 'SKIPPED'
						Details			     = "The recipient is not a User or a Shared Mailbox. it is $($recipient.RecipientTypeDetails)"
					}
				}
			}
			catch
			{
				$prop = [ordered] @{
					UserPrincipalName    = $UPN
					RecipientTypeDetails = $recipient.RecipientTypeDetails
					EnabledPack		     = $add
					DisabledPack		 = $removed
					Modified			 = "$mod Disabled:$DisabledPlans"
					IsGeneric		     = $gen
					ADGrpRemoval		 = $Adrem
					Status			     = 'Failed'
					Details			     = "ERROR : $($_.Exception.Message)"
				}
			}
			finally
			{
				$obj = New-Object -TypeName System.Management.Automation.PSObject -Property $prop
				Write-Output $obj
				if ($ShowProgress)
				{
					if ($UserPrincipalName.count -gt 1)
					{
						$paramWriteProgress = @{
							Activity = 'Assigning Licenses'
							Status   = "Processing [$i] of [$($UserPrincipalName.Count)] users"
							PercentComplete = (($i / $UserPrincipalName.Count) * 100)
							CurrentOperation = "Completed : [$UPN]"
						}
						Write-Progress @paramWriteProgress
					}
					$i++
				}
			}
		}
	}
	end
	{
		Write-Progress -Activity 'Assigning Licenses' -Completed
		$global:ErrorActionPreference = [System.Management.Automation.ActionPreference]::Continue
	}
}

function Set-ii3QHLicense
{
<#
	.SYNOPSIS
		A brief description of the Set-ii3QHLicense function.
	
	.DESCRIPTION
		A detailed description of the Set-ii3QHLicense function.
	
	.PARAMETER UserPrincipalName
		A description of the UserPrincipalName parameter.
	
	.PARAMETER disabledServices
		A description of the disabledServices parameter.
	
	.PARAMETER ShowProgress
		A description of the ShowProgress parameter.
	
	.EXAMPLE
		PS C:\> Set-ii3QHLicense
	
	.NOTES
		Additional information about the function.
#>
	[CmdletBinding()]
	param
	(
		[String[]]$UserPrincipalName,
		[ValidateSet('E3:ExchangeOnline', 'E3:SkypeForBusinessOnline', 'E3:SharepointOnline')]
		[String[]]$disabledServices,
		[Switch]$ShowProgress
	)
	
	begin
	{
		$global:ErrorActionPreference = [System.Management.Automation.ActionPreference]::Stop
		$i = 1
		
		$DisabledPlans = switch ($disabledServices)
		{
			'E3:ExchangeOnline' {
				'EXCHANGE_S_ENTERPRISE'
			}
			'E3:SkypeForBusinessOnline' {
				'MCOSTANDARD'
			}
			'E3:SharepointOnline' {
				'SHAREPOINTENTERPRISE', 'SHAREPOINTWAC'
			}
			default
			{
				$_
			}
		}
		
		$genericGroups = @('ACT-SaaS-O365-Non-Employee', 'ACT-SaaS-O365-GenericAccountMailboxOnly', 'ACT-SaaS-O365-GenericAccountFullLicence', 'ACT-SaaS-O365-GenericAccountOfficeProPlusAndMailbox')
	}
	process
	{
		foreach ($UPN in $UserPrincipalName)
		{
			$EnabledPack = @()
			$disabledPack = @()
			$licinfo = $null
			$Newlicinfo = $null
			
			try
			{
				Write-Verbose "Trying to get $UPN Info"
				$recipient = Get-Recipient $UPN -ErrorAction Stop
				Write-Verbose "Retrived $UPN info"
				
				if ($recipient.RecipientTypeDetails -match 'UserMailbox')
				{
					Write-Verbose "Trying to retrive licensing details for $UPN"
					$licinfo = Get-QHLicenseStatus -UserPrincipalName $UPN
					Write-Verbose "Retrived Licensing details for $UPN"
					
					Write-Verbose "The $UPN Type is $($recipient.RecipientTypeDetails). Trying to Set the Usage Location to AU"
					Set-MsolUser -UserPrincipalName $UPN -UsageLocation "AU" -ErrorAction SilentlyContinue
					Write-Verbose "Usage Location is now set to AU"
					
					#check for Plans to Disable
					
					Write-Verbose "Current License as Follows"
					
					Write-Verbose "$($licinfo | Out-String)"
					
					if ($licinfo.E1 -ne $null)
					{
						#remove E1
						$disabledPack += 'healthqld:STANDARDPACK'
					}
					if ($licinfo.ExchangeOnlinePlan2 -ne $null)
					{
						#remove Exchange Online Plan2
						$disabledPack += 'healthqld:EXCHANGEENTERPRISE'
					}
					
					if ($licinfo.E3 -eq $null)
					{
						$gen = $false
						$Adrem = 'None'
						#Add E3
						$EnabledPack += 'healthqld:ENTERPRISEPACK'
						
						#Add MFA
						if ($licinfo.MFA -eq $null)
						{
							$EnabledPack += 'healthqld:MFA_STANDALONE'
						}
						
						#Add Power Bi Free
						if ($licinfo.'PowerBi Free' -eq $null)
						{
							$EnabledPack += 'healthqld:POWER_BI_STANDARD'
						}
						
						#check for disabled Services.
						if ($disabledServices -ne $null)
						{
							$paramNewMsolLicenseOptions = @{
								AccountSkuId  = 'healthqld:ENTERPRISEPACK'
								ErrorAction   = 'Stop'
								DisabledPlans = $DisabledPlans
							}
						}
						
						$options = New-MsolLicenseOptions @paramNewMsolLicenseOptions
					}
					else
					{
						if ($licinfo.E3 -icontains 'ACT-SaaS-O365-GenericAccountOfficeProPlusOnly')
						{
							#remove from this group and add to the new group
							$paramRemoveQHE3ADGrpmembershipGen = @{
								AdGroup	       = 'ACT-SaaS-O365-GenericAccountOfficeProPlusOnly'
								SamAccountName = $recipient.SamAccountName
							}
							$rgen = Remove-QHE3ADGrpmembership @paramRemoveQHE3ADGrpmembershipGen
							
							#add to Generic ProPlus and Mailbox Group
							$paramAddQHGenericGroupMember = @{
								SamAccountName = $recipient.SamAccountName
								AdGroup	       = 'ACT-SaaS-O365-GenericAccountOfficeProPlusAndMailbox'
							}
							
							$agen = Add-QHGenericGroupMember @paramAddQHGenericGroupMember
							
							$gen = "$rgen $agen"
							
							if ($licinfo.E3 -icontains 'Direct' -or $licinfo.E3 -contains 'Direct:NoGUID')
							{
								#removeDirectLicensing
								$disabledPack += 'healthqld:ENTERPRISEPACK'
							}
							if ($licinfo.E3 -icontains 'ADM-O365-LIC-E3-OfficeProPlusOnly')
							{
								#remove from this group
								$paramRemoveQHE3ADGrpmembership = @{
									AdGroup	       = 'ADM-O365-LIC-E3-OfficeProPlusOnly'
									SamAccountName = $recipient.SamAccountName
								}
								
								$Adrem = Remove-QHE3ADGrpmembership @paramRemoveQHE3ADGrpmembership
							}
							else
							{
								$Adrem = 'None'
							}
							
							#removeMFA License if exists
							if ($licinfo.MFA -ne $null)
							{
								$disabledPack += 'healthqld:MFA_STANDALONE'
							}
							
							#remove PowerBi free if exists
							if ($licinfo.'PowerBi Free' -ne $null)
							{
								$disabledPack += 'healthqld:POWER_BI_STANDARD'
							}
						}
						elseif ($licinfo.E3 -icontains 'ADM-O365-LIC-E3-OfficeProPlusOnly')
						{
							$gen = $false
							#remove from this Group
							
							$paramRemoveQHE3ADGrpmembership = @{
								AdGroup	       = 'ADM-O365-LIC-E3-OfficeProPlusOnly'
								SamAccountName = $recipient.SamAccountName
							}
							
							$Adrem = Remove-QHE3ADGrpmembership @paramRemoveQHE3ADGrpmembership
							
							if ($licinfo.E3 -icontains 'Direct' -or $licinfo.E3 -icontains 'Direct:NoGUID')
							{
								#set the license
								if ($disabledServices -ne $null)
								{
									$paramNewMsolLicenseOptions = @{
										AccountSkuId  = 'healthqld:ENTERPRISEPACK'
										ErrorAction   = 'Stop'
										DisabledPlans = $DisabledPlans
									}
								}
								else
								{
									$paramNewMsolLicenseOptions = @{
										AccountSkuId  = 'healthqld:ENTERPRISEPACK'
										ErrorAction   = 'Stop'
										DisabledPlans = $null
									}
								}
								
								$options = New-MsolLicenseOptions @paramNewMsolLicenseOptions
							}
							else
							{
								#Add the License
								$EnabledPack += 'healthqld:ENTERPRISEPACK'
								
								if ($disabledServices -ne $null)
								{
									$paramNewMsolLicenseOptions = @{
										AccountSkuId  = 'healthqld:ENTERPRISEPACK'
										ErrorAction   = 'Stop'
										DisabledPlans = $DisabledPlans
									}
								}
								
								$options = New-MsolLicenseOptions @paramNewMsolLicenseOptions
							}
							
							#Add MFA
							if ($licinfo.MFA -eq $null)
							{
								
								$EnabledPack += 'healthqld:MFA_STANDALONE'
							}
							
							#Add Power Bi Free
							if ($licinfo.'PowerBi Free' -eq $null)
							{
								$EnabledPack += 'healthqld:POWER_BI_STANDARD'
							}
						}
						elseif ($licinfo.E3.Where({ $_ -notlike 'Direct*' }) -in $genericGroups)
						{
							# No Action Needed. Already is in Desired Generic Group
							# clarify if we need to remove the other direct licensing like DirectE2 etc
							
							$gen = $licinfo.E3.Where({ $_ -notlike 'Direct*' }) + " :NoActionNeeded"
							$Adrem = 'None'
							$DisabledPlans = 'NotApplicable'
							
						}
						elseif ($licinfo.E3 -icontains 'Direct' -or $licinfo.E3 -icontains 'Direct:NoGUID')
						{
							$gen = $false
							$Adrem = 'None'
							#set E3 License Option
							if ($disabledServices -ne $null)
							{
								$paramNewMsolLicenseOptions = @{
									AccountSkuId  = 'healthqld:ENTERPRISEPACK'
									ErrorAction   = 'Stop'
									DisabledPlans = $DisabledPlans
								}
							}
							else
							{
								$paramNewMsolLicenseOptions = @{
									AccountSkuId  = 'healthqld:ENTERPRISEPACK'
									ErrorAction   = 'Stop'
									DisabledPlans = $null
								}
							}
							
							$options = New-MsolLicenseOptions @paramNewMsolLicenseOptions
							
							#check MFA
							if ($licinfo.MFA -eq $null)
							{
								$EnabledPack += 'healthqld:MFA_STANDALONE'
							}
							
							#Add Power Bi Free
							if ($licinfo.'PowerBi Free' -eq $null)
							{
								$EnabledPack += 'healthqld:POWER_BI_STANDARD'
							}
						}
						else
						{
							#report and do nothing.
						}
					}
					
					$LicenceParam = @{
						UserPrincipalName = $UPN
						ErrorAction	      = 'Stop'
					}
					
					if ($disabledPack -ne $null)
					{
						$LicenceParam.Add('RemoveLicenses', $disabledPack)
						$removed = $(($disabledPack -join ',').Trim())
					}
					else
					{
						$removed = 'None'
					}
					
					if ($EnabledPack -ne $null)
					{
						$LicenceParam.Add('AddLicenses', $EnabledPack)
						$add = $(($EnabledPack -join ',').Trim())
					}
					else
					{
						$add = 'None'
					}
					
					if ($options -ne $null)
					{
						$LicenceParam.Add('LicenseOptions', $options)
						$mod = "E3:EnabledAllServices"
					}
					else
					{
						$mod = 'None'
					}
					if ($EnabledPack -ne $null -or $disabledPack -ne $null -or $options -ne $null)
					{
						Set-MsolUserLicense @LicenceParam
					}
					
					$Newlicinfo = Get-QHLicenseStatus -UserPrincipalName $UPN
					
					Write-Verbose "New License as Follows"
					
					Write-Verbose "$($Newlicinfo | Out-String)"
					
					if ($DisabledPlans -ne $null)
					{
						$D = " DisabledServices:$DisabledPlans"
					}
					else
					{
						$D = " DisabledServices:None"
					}
					
					$prop = [ordered] @{
						UserPrincipalName    = $UPN
						RecipientTypeDetails = $recipient.RecipientTypeDetails
						EnabledPack		     = $add
						DisabledPack		 = $removed
						Modified			 = $mod + $D
						IsGeneric		     = $gen
						ADGrpRemoval		 = $Adrem
						Status			     = 'Success'
						Details			     = 'None'
					}
				}
				elseif ($recipient.RecipientTypeDetails -match 'SharedMailbox')
				{
					Write-Verbose "Trying to retrive licensing details for $UPN"
					$licinfo = Get-QHLicenseStatus -UserPrincipalName $UPN
					Write-Verbose "Retrived Licensing details for $UPN"
					
					Write-Verbose "The $UPN Type is $($recipient.RecipientTypeDetails). Trying to Set the Usage Location to AU"
					Set-MsolUser -UserPrincipalName $UPN -UsageLocation "AU" -ErrorAction SilentlyContinue
					Write-Verbose "Usage Location is now set to AU"
					
					#Check for Plans to Disable
					
					Write-Verbose "Current License as Follows"
					
					Write-Verbose "$($licinfo | Out-String)"
					
					#Check if the Login is Enabled
					$AD = get-AdUser $recipient.SamAccountname -ErrorAction Stop
					
					#Remove E1 if exists
					if ($licinfo.E1 -ne $null)
					{
						$disabledPack += 'healthqld:STANDARDPACK'
					}
					
					#Remove Power Bi if exists
					if ($licinfo.'PowerBi Free' -ne $null)
					{
						$disabledPack += 'healthqld:POWER_BI_STANDARD'
					}
					
					#Remove MFA If exists
					
					if ($licinfo.MFA -ne $null)
					{
						$disabledPack += 'healthqld:MFA_STANDALONE'
					}
					
					# E3 Ops if Enabled
					
					if ($licinfo.E3 -eq $null)
					{
						#check if login enabled
						
						if ($AD.Enabled)
						{
							#Disable Exchange Online Plan 2 if Enabled.
							if ($licinfo.ExchangeOnlinePlan2 -ne $null)
							{
								if ($licinfo.ExchangeOnlinePlan2 -icontains 'Direct' -or $licinfo.ExchangeOnlinePlan2 -icontains 'Direct:NoGUID')
								{
									$disabledPack += 'healthqld:EXCHANGEENTERPRISE'
								}
								
								if ($licinfo.ExchangeOnlinePlan2.Where({ $_ -notlike 'Direct*' }) -in $genericGroups)
								{
									# No Action Needed. Already is in Desired Generic Group
									# clarify if we need to remove the other direct licensing like DirectE2 etc
									
									$gen = $licinfo.ExchangeOnlinePlan2.Where({ $_ -notlike 'Direct*' }) + " :NoActionNeeded"
									$Adrem = 'None'
									$DisabledPlans = 'NotApplicable'
								}
								else
								{
									#Add to Generic Group
									$paramAddQHGenericGroupMember = @{
										SamAccountName = $recipient.SamAccountName
										AdGroup	       = 'ACT-SaaS-O365-GenericAccountMailboxOnly'
									}
									
									$agen = Add-QHGenericGroupMember @paramAddQHGenericGroupMember
									
									$gen = "$agen"
									$Adrem = 'None'
								}
							}
						}
						else
						{
							if ($licinfo.ExchangeOnlinePlan2 -eq $null)
							{
								$EnabledPack += 'healthqld:EXCHANGEENTERPRISE'
							}
							
							$Adrem = 'None'
							
							$gen = $false
						}
					}
					else
					{
						#Remove DirectLicensing if Exists
						
						if ($licinfo.E3 -icontains 'Direct' -or $licinfo.E3 -contains 'Direct:NoGUID')
						{
							$disabledPack += 'healthqld:ENTERPRISEPACK'
						}
						
						#remove from the other licensing group if exists
						
						if ($licinfo.E3 -icontains 'ADM-O365-LIC-E3-OfficeProPlusOnly')
						{
							$paramRemoveQHE3ADGrpmembership = @{
								AdGroup	       = 'ADM-O365-LIC-E3-OfficeProPlusOnly'
								SamAccountName = $recipient.SamAccountName
							}
							
							$Adrem = Remove-QHE3ADGrpmembership @paramRemoveQHE3ADGrpmembership
						}
						else
						{
							$Adrem = 'None'
						}
						
						# Check if it is logon enabled Generic
						
						if ($licinfo.E3 -icontains 'ACT-SaaS-O365-GenericAccountOfficeProPlusOnly')
						{
							$paramRemoveQHE3ADGrpmembershipGen = @{
								AdGroup	       = 'ACT-SaaS-O365-GenericAccountOfficeProPlusOnly'
								SamAccountName = $recipient.SamAccountName
							}
							$rgen = Remove-QHE3ADGrpmembership @paramRemoveQHE3ADGrpmembershipGen
							
							#add to Generic ProPlus and Mailbox Group
							$paramAddQHGenericGroupMember = @{
								SamAccountName = $recipient.SamAccountName
								AdGroup	       = 'ACT-SaaS-O365-GenericAccountOfficeProPlusAndMailbox'
							}
							
							$agen = Add-QHGenericGroupMember @paramAddQHGenericGroupMember
							
							$gen = "$rgen $agen"
						}
						
						if ($licinfo.E3.Where({ $_ -notlike 'Direct*' }) -in $genericGroups)
						{
							# No Action Needed. Already is in Desired Generic Group
							# clarify if we need to remove the other direct licensing like DirectE2 etc
							
							$gen = $licinfo.E3.Where({ $_ -notlike 'Direct*' }) + " :NoActionNeeded"
							$Adrem = 'None'
							$DisabledPlans = 'NotApplicable'
							
						}
						
					}
					
					$LicenceParam = @{
						UserPrincipalName = $UPN
						ErrorAction	      = 'Stop'
					}
					
					if ($disabledPack -ne $null)
					{
						$LicenceParam.Add('RemoveLicenses', $disabledPack)
						$removed = $(($disabledPack -join ',').Trim())
					}
					else
					{
						$removed = 'None'
					}
					
					if ($EnabledPack -ne $null)
					{
						$LicenceParam.Add('AddLicenses', $EnabledPack)
						$add = $(($EnabledPack -join ',').Trim())
					}
					else
					{
						$add = 'None'
					}
					
					if ($options -ne $null)
					{
						$LicenceParam.Add('LicenseOptions', $options)
						$mod = "E3:EnabledAllServices"
					}
					else
					{
						$mod = 'None'
					}
					if ($EnabledPack -ne $null -or $disabledPack -ne $null -or $options -ne $null)
					{
						Set-MsolUserLicense @LicenceParam
					}
					
					$Newlicinfo = Get-QHLicenseStatus -UserPrincipalName $UPN
					
					Write-Verbose "New License as Follows"
					
					Write-Verbose "$($Newlicinfo | Out-String)"
					
					$prop = [ordered] @{
						UserPrincipalName    = $UPN
						RecipientTypeDetails = $recipient.RecipientTypeDetails
						EnabledPack		     = $add
						DisabledPack		 = $removed
						Modified			 = $mod
						IsGeneric		     = $gen
						ADGrpRemoval		 = $Adrem
						Status			     = 'Success'
						Details			     = 'None'
					}
				}
				else
				{
					$gen = 'NotApplicable'
					$prop = [ordered] @{
						UserPrincipalName    = $UPN
						RecipientTypeDetails = $recipient.RecipientTypeDetails
						EnabledPack		     = 'Skipped'
						DisabledPack		 = 'Skipped'
						Modified			 = 'Skipped'
						IsGeneric		     = $gen
						ADGrpRemoval		 = 'Skipped'
						Status			     = 'SKIPPED'
						Details			     = "The recipient is not a User or a Shared Mailbox. it is $($recipient.RecipientTypeDetails)"
					}
				}
			}
			catch
			{
				$prop = [ordered] @{
					UserPrincipalName    = $UPN
					RecipientTypeDetails = $recipient.RecipientTypeDetails
					EnabledPack		     = $add
					DisabledPack		 = $removed
					Modified			 = $mod
					IsGeneric		     = $gen
					ADGrpRemoval		 = $Adrem
					Status			     = 'Failed'
					Details			     = "ERROR : $($_.Exception.Message)"
				}
			}
			finally
			{
				$obj = New-Object -TypeName System.Management.Automation.PSObject -Property $prop
				Write-Output $obj
				if ($ShowProgress)
				{
					if ($UserPrincipalName.count -gt 1)
					{
						$paramWriteProgress = @{
							Activity = 'Assigning Licenses'
							Status   = "Processing [$i] of [$($UserPrincipalName.Count)] users"
							PercentComplete = (($i / $UserPrincipalName.Count) * 100)
							CurrentOperation = "Completed : [$UPN]"
						}
						Write-Progress @paramWriteProgress
					}
					$i++
				}
			}
		}
	}
	end
	{
		Write-Progress -Activity 'Assigning Licenses' -Completed
		$global:ErrorActionPreference = [System.Management.Automation.ActionPreference]::Continue
	}
}

function Set-QHLicenseWithGeneric
{
<#
	.SYNOPSIS
		A brief description of the Set-QHLicenseWithGeneric function.
	
	.DESCRIPTION
		A detailed description of the Set-QHLicenseWithGeneric function.
	
	.PARAMETER UserPrincipalName
		A description of the UserPrincipalName parameter.
	
	.PARAMETER DisabledServices
		A description of the DisabledServices parameter.
	
	.PARAMETER BatchName
		A description of the BatchName parameter.
	
	.PARAMETER Progress
		A description of the Progress parameter.
	
	.EXAMPLE
		PS C:\> Set-QHLicenseWithGeneric -UserPrincipalName 'value1'
	
	.NOTES
		Additional information about the function.
#>
	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory = $true)]
		[String[]]$UserPrincipalName,
		[ValidateSet('E3:ExchangeOnline', 'E3:SkypeForBusinessOnline', 'E3:SharepointOnline')]
		[String[]]$DisabledServices,
		[String]$BatchName = 'None',
		[switch]$Progress
	)
	
	begin
	{
		$i = 1
		$Global:ErrorActionPreference = [System.Management.Automation.ActionPreference]::Stop
		
		$xDisabledServices = switch ($disabledServices)
		{
			'E3:ExchangeOnline' {
				'EXCHANGE_S_ENTERPRISE'
			}
			'E3:SkypeForBusinessOnline' {
				'MCOSTANDARD'
			}
			'E3:SharepointOnline' {
				'SHAREPOINTENTERPRISE', 'SHAREPOINTWAC'
			}
			default
			{
				$_
			}
		}
	}
	process
	{
		foreach ($UPN in $UserPrincipalName)
		{
			$paramNewMsolLicenseOptions = $null
			$enabledPack = @()
			$disabledPack = @()
			$xgenlic = $null
			$options = $null
			
			try
			{
				$xgenlic = Validate-QHGeneric -UserPrincipalName $UPN
				
				if ($xgenlic.IsGeneric)
				{
					#no Processing needed
					$prop = [Ordered] @{
						UserPrincipalName    = $UPN
						RecipientTypeDetails = $xgenlic.RecipientTypeDetails
						IsGeneric		     = $xgenlic.IsGeneric
						ADAccountEnabled	 = $xgenlic.ADAccountEnabled
						GenericDetails	     = $xgenlic.GenericDetails
						DisabledDirectLicenses = $xgenlic.DisabledDirectLicenses
						EnabledDirectLicenses = 'NotApplicable'
						DisabledServices	 = 'NotApplicable'
						Status			     = $xgenlic.Status
						Details			     = $xgenlic.Details
					}
				}
				elseif ($xgenlic.IsGeneric -eq $false)
				{
					Set-MsolUser -UserPrincipalName $UPN -UsageLocation "AU" -ErrorAction SilentlyContinue
					#$ad = get-AdUser $recipient.SamAccountName -properties * -ErrorAction Stop
					#$groups = $ad.MemberOf.ForEach({ $_.Split(',')[0].Split('=')[1] })
					$lic = Get-iQHLicenseStatus $UPN
					# Processing needed
					if ($xgenlic.RecipientTypeDetails -match 'User')
					{
						#Process for regular Shared Mailbox
						
						if ($lic.E1 -match 'Direct') { $disabledPack += 'healthqld:STANDARDPACK' }
						if ($lic.ExchangeOnlinePlan2 -match 'direct') { $disabledPack += 'healthqld:EXCHANGEENTERPRISE' }
						
						
						if ($lic.MFA -eq $null) { $enabledPack += 'healthqld:MFA_STANDALONE' }
						if ($lic.'PowerBi Free' -eq $null) { $enabledPack += 'healthqld:POWER_BI_STANDARD' }
						if ($lic.E3 -eq $null)
						{
							$enabledPack += 'healthqld:ENTERPRISEPACK'
							
							if ($DisabledServices -ne $null)
							{
								$paramNewMsolLicenseOptions = @{
									AccountSkuId  = 'healthqld:ENTERPRISEPACK'
									ErrorAction   = 'Stop'
									DisabledPlans = $xDisabledServices
								}
								
								$options = New-MsolLicenseOptions @paramNewMsolLicenseOptions
							}
						}
						else
						{
							if ($lic.E3 -match 'Direct')
							{
								# Update License
								if ($xDisabledServices -ne $null)
								{
									$paramNewMsolLicenseOptions = @{
										AccountSkuId  = 'healthqld:ENTERPRISEPACK'
										ErrorAction   = 'Stop'
										DisabledPlans = $xDisabledServices
									}
								}
								else
								{
									$paramNewMsolLicenseOptions = @{
										AccountSkuId = 'healthqld:ENTERPRISEPACK'
										ErrorAction  = 'Stop'
									}
								}
								
								$options = New-MsolLicenseOptions @paramNewMsolLicenseOptions
							}
							else
							{
								$enabledPack += 'healthqld:ENTERPRISEPACK'
								if ($DisabledServices -ne $null)
								{
									$paramNewMsolLicenseOptions = @{
										AccountSkuId  = 'healthqld:ENTERPRISEPACK'
										ErrorAction   = 'Stop'
										DisabledPlans = $xDisabledServices
									}
									
									$options = New-MsolLicenseOptions @paramNewMsolLicenseOptions
									
								}
								
							}
							if ($lic.E3 -match 'ADM-O365-LIC-E3-OfficeProPlusOnly')
							{
								$paramRemoveQHE3ADGrpmembership = @{
									AdGroup	       = 'ADM-O365-LIC-E3-OfficeProPlusOnly'
									SamAccountName = $xgenlic.SamAccountName #Need to get Sam Account name
								}
								
								$Adrem = Remove-QHE3ADGrpmembership @paramRemoveQHE3ADGrpmembership
							}
							else
							{
								$Adrem = 'NoProPlusGroupMemberShip'
							}
						}
						
						$LicenceParam = @{
							UserPrincipalName = $UPN
							ErrorAction	      = 'Stop'
						}
						
						if ($disabledPack -ne $null)
						{
							$LicenceParam.Add('RemoveLicenses', $disabledPack)
							$removed = $(($disabledPack -join ',').Trim())
						}
						else
						{
							$removed = 'None'
						}
						
						if ($EnabledPack -ne $null)
						{
							$LicenceParam.Add('AddLicenses', $EnabledPack)
							$add = $(($EnabledPack -join ',').Trim())
						}
						else
						{
							$add = 'None'
						}
						
						if ($options -ne $null)
						{
							$LicenceParam.Add('LicenseOptions', $options)
							#$mod = "E3:EnabledAllServices"
						}
						
						if ($EnabledPack -ne $null -or $disabledPack -ne $null -or $options -ne $null)
						{
							Set-MsolUserLicense @LicenceParam
						}
						
						
						if ($xDisabledServices -ne $null)
						{
							$D = $(($DisabledServices -join ',').Trim())
						}
						else
						{
							$D = "None"
						}
						
						$prop = [Ordered] @{
							UserPrincipalName    = $UPN
							RecipientTypeDetails = $xgenlic.RecipientTypeDetails
							IsGeneric		     = $xgenlic.IsGeneric
							ADAccountEnabled	 = $xgenlic.ADAccountEnabled
							GenericDetails	     = $xgenlic.GenericDetails
							DisabledDirectLicenses = $removed
							EnabledDirectLicenses = $add
							DisabledServices	 = $D
							Status			     = 'Success'
							Details			     = 'None ' + $Adrem
						}
					}
					elseif ($xgenlic.RecipientTypeDetails -match 'Shared')
					{
						#process for Regular Mailbox
						if ($lic.E1 -ne $null) { $disabledPack += 'healthqld:STANDARDPACK' }
						if ($lic.'PowerBi Free' -ne $null) { $disabledPack += 'healthqld:POWER_BI_STANDARD' }
						if ($lic.MFA -ne $null) { $disabledPack += 'healthqld:MFA_STANDALONE' }
						if ($lic.E3 -ne $null)
						{
							if ($lic.E3 -match 'Direct')
							{
								$disabledPack += 'healthqld:ENTERPRISEPACK'
							}
							if ($lic.E3 -match 'ADM-O365-LIC-E3-OfficeProPlusOnly')
							{
								$paramRemoveQHE3ADGrpmembership = @{
									AdGroup	       = 'ADM-O365-LIC-E3-OfficeProPlusOnly'
									SamAccountName = $xgenlic.SamAccountName
								}
								
								$Adrem = Remove-QHE3ADGrpmembership @paramRemoveQHE3ADGrpmembership
							}
							else
							{
								$Adrem = 'NoProPlusGroupMemberShip'
							}
						}
						if ($lic.ExchangeOnlinePlan2 -eq $null) { $enabledPack += 'healthqld:EXCHANGEENTERPRISE' }
						
						$LicenceParam = @{
							UserPrincipalName = $UPN
							ErrorAction	      = 'Stop'
						}
						
						if ($disabledPack -ne $null)
						{
							$LicenceParam.Add('RemoveLicenses', $disabledPack)
							$removed = $(($disabledPack -join ',').Trim())
						}
						else
						{
							$removed = 'None'
						}
						
						if ($EnabledPack -ne $null)
						{
							$LicenceParam.Add('AddLicenses', $EnabledPack)
							$add = $(($EnabledPack -join ',').Trim())
						}
						else
						{
							$add = 'None'
						}
						
						if ($options -ne $null)
						{
							$LicenceParam.Add('LicenseOptions', $options)
							#$mod = "E3:EnabledAllServices"
						}
						
						if ($EnabledPack -ne $null -or $disabledPack -ne $null -or $options -ne $null)
						{
							Set-MsolUserLicense @LicenceParam
						}
						$D = 'NotApplicable'
						
						$prop = [Ordered] @{
							UserPrincipalName    = $UPN
							RecipientTypeDetails = $xgenlic.RecipientTypeDetails
							IsGeneric		     = $xgenlic.IsGeneric
							ADAccountEnabled	 = $xgenlic.ADAccountEnabled
							GenericDetails	     = $xgenlic.GenericDetails
							DisabledDirectLicenses = $removed
							EnabledDirectLicenses = $add
							DisabledServices	 = $D
							Status			     = 'Success'
							Details			     = 'None ' + $Adrem
						}
					}
				}
				else
				{
					# Is null throw error
					$prop = [Ordered] @{
						UserPrincipalName    = $UPN
						RecipientTypeDetails = $xgenlic.RecipientTypeDetails
						IsGeneric		     = $xgenlic.IsGeneric
						ADAccountEnabled	 = $xgenlic.ADAccountEnabled
						GenericDetails	     = $xgenlic.GenericDetails
						DisabledDirectLicenses = $xgenlic.DisabledDirectLicenses
						EnabledDirectLicenses = 'NotApplicable'
						DisabledServices	 = 'NotApplicable'
						Status			     = $xgenlic.Status
						Details			     = $xgenlic.Details
					}
				}
			}
			catch
			{
				$prop = [Ordered] @{
					UserPrincipalName    = $UPN
					RecipientTypeDetails = $xgenlic.RecipientTypeDetails
					IsGeneric		     = $xgenlic.IsGeneric
					ADAccountEnabled	 = $xgenlic.ADAccountEnabled
					GenericDetails	     = $xgenlic.GenericDetails
					DisabledDirectLicenses = $removed
					EnabledDirectLicenses = $add
					DisabledServices	 = $D
					Status			     = 'Failed'
					Details			     = "ERROR : $($_.Exception.Message)"
				}
			}
			finally
			{
				$obj = New-Object -TypeName System.Management.Automation.PSObject -Property $prop
				Write-Output $obj
				if ($UserPrincipalName.count -gt 1)
				{
					if ($Progress)
					{
						$paramWriteProgress = @{
							Activity = 'Doing Some Processing'
							Status   = "Processing [$i] of [$($UserPrincipalName.Count)] users"
							PercentComplete = (($i / $UserPrincipalName.Count) * 100)
							CurrentOperation = "Completed : [$UPN]"
						}
						Write-Progress @paramWriteProgress
					}
				}
				$i++
			}
		}
	}
	end
	{
		Write-Progress -Activity 'Doing Some Processing' -Completed
		$Global:ErrorActionPreference = [System.Management.Automation.ActionPreference]::Continue
	}
}

function Add-QHUserToLocalAdminGroup
{
<#
	.SYNOPSIS
		A brief description of the Add-QHUserToLocalAdminGroup function.
	
	.DESCRIPTION
		A detailed description of the Add-QHUserToLocalAdminGroup function.
	
	.PARAMETER UserSamAccountName
		A description of the UserSamAccountName parameter.
	
	.PARAMETER Domain
		A description of the Domain parameter.
	
	.PARAMETER ComputerName
		A description of the ComputerName parameter.
	
	.PARAMETER ShowProgress
		A description of the ShowProgress parameter.
	
	.EXAMPLE
		PS C:\> Add-QHUserToLocalAdminGroup -UserSamAccountName 'value1' -Domain 'value2' -ComputerName 'value3'
	
	.NOTES
		Additional information about the function.
#>
	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory = $true)]
		[String[]]$UserSamAccountName,
		[Parameter(Mandatory = $true)]
		[String]$Domain = 'QH',
		[Parameter(Mandatory = $true)]
		[String]$ComputerName = $env:COMPUTERNAME,
		[switch]$ShowProgress
	)
	
	begin
	{
		$i = 1
		$Global:ErrorActionPreference = [System.Management.Automation.ActionPreference]::Stop
		if (!(Test-Connection -ComputerName $ComputerName -Count 1 -Quiet))
		{
			Write-Warning "$ComputerName is not reachable"
			break
		}
		
	}
	process
	{
		foreach ($User in $UserSamAccountName)
		{
			try
			{
				$LocalAdminGroup = [ADSI]"WinNT://$ComputerName/Administrators"
				$LocalAdminGroup.Add("WinNT://$Domain/$User")
				
				$prop = [Ordered] @{
					User		 = $User
					ComputerName = $ComputerName
					ops		     = 'Added to Administrators group'
					Status	     = 'Success'
					details	     = 'None'
				}
			}
			catch
			{
				$prop = [Ordered] @{
					User		 = $User
					ComputerName = $ComputerName
					ops		     = 'Added to Administrators group'
					Status	     = 'Success'
					details	     = "ERROR : $($_.Exception.Message)"
				}
			}
			finally
			{
				$obj = New-Object -TypeName System.Management.Automation.PSObject -Property $prop
				Write-Output $prop
				
				if ($UserPrincipalName.count -gt 1)
				{
					if ($ShowProgress)
					{
						$paramWriteProgress = @{
							Activity = 'Doing Some Processing'
							Status   = "Processing [$i] of [$($UserPrincipalName.Count)] users"
							PercentComplete = (($i / $UserPrincipalName.Count) * 100)
							CurrentOperation = "Completed : [$UPN]"
						}
						Write-Progress @paramWriteProgress
					}
				}
				$i++
			}
		}
	}
	end
	{
		Write-Progress -Activity 'Doing Some Processing' -Completed
		$Global:ErrorActionPreference = [System.Management.Automation.ActionPreference]::Continue
	}
}

function Validate-QHGeneric
{
<#
	.SYNOPSIS
		A brief description of the Validate-QHGeneric function.
	
	.DESCRIPTION
		A detailed description of the Validate-QHGeneric function.
	
	.PARAMETER UserPrincipalName
		A description of the UserPrincipalName parameter.
	
	.PARAMETER ShowProgress
		A description of the ShowProgress parameter.
	
	.EXAMPLE
		PS C:\> Validate-QHGeneric -UserPrincipalName 'value1'
	
	.NOTES
		Additional information about the function.
#>
	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory = $true)]
		[String[]]$UserPrincipalName,
		[switch]$ShowProgress
	)
	
	begin
	{
		$i = 1
		$Global:ErrorActionPreference = [System.Management.Automation.ActionPreference]::Stop
		
		$genericGroups = @('ACT-SaaS-O365-Non-Employee', 'ACT-SaaS-O365-GenericAccountMailboxOnly', 'ACT-SaaS-O365-GenericAccountFullLicence', 'ACT-SaaS-O365-GenericAccountOfficeProPlusAndMailbox', 'ACT-SaaS-O365-GenericAccountOfficeProPlusOnly')
	}
	process
	{
		foreach ($UPN in $UserPrincipalName)
		{
			try
			{
				$recipient = $null
				$ad = $null
				$genOpsdetails = $null
				$GenOps = @()
				$PlansToDisable = $null
				$Allgroups = $null
				$recipient = Get-Recipient $UPN -ErrorAction Stop
				
				if ($recipient.RecipientTypeDetails -match 'UserMailbox')
				{
					#$generic = $true
					
					$ad = get-AdUser $recipient.SamAccountName -properties * -ErrorAction Stop
					$Allgroups = $ad.MemberOf.ForEach({ $_.Split(',')[0].Split('=')[1] })
					$groups = @()
					
					foreach ($g in $Allgroups)
					{
						if ($genericGroups -icontains $g)
						{
							$groups += $g
						}
					}
					
					if ($groups -ne $null)
					{
						$generic = $true
						
						if ($groups -contains 'ACT-SaaS-O365-GenericAccountOfficeProPlusOnly')
						{
							$paramRemoveQHE3ADGrpmembershipGen = @{
								AdGroup	       = 'ACT-SaaS-O365-GenericAccountOfficeProPlusOnly'
								SamAccountName = $recipient.SamAccountName
							}
							
							$GenOps += Remove-QHE3ADGrpmembership @paramRemoveQHE3ADGrpmembershipGen
							
							if ($groups -notcontains 'ACT-SaaS-O365-GenericAccountOfficeProPlusAndMailbox')
							{
								$paramAddQHGenericGroupMember = @{
									SamAccountName = $recipient.SamAccountName
									AdGroup	       = 'ACT-SaaS-O365-GenericAccountOfficeProPlusAndMailbox'
								}
								
								$GenOps += Add-QHGenericGroupMember @paramAddQHGenericGroupMember
							}
							
							if ($Allgroups -icontains 'ADM-O365-LIC-E3-OfficeProPlusOnly')
							{
								$paramRemoveQHE3ADGrpmembership = @{
									AdGroup	       = 'ADM-O365-LIC-E3-OfficeProPlusOnly'
									SamAccountName = $recipient.SamAccountName
								}
								
								$GenOps += Remove-QHE3ADGrpmembership @paramRemoveQHE3ADGrpmembership
							}
						}
						else
						{
							#ReportOnly
							$GenOps = $groups
						}
						
						$genOpsdetails = $($GenOps -join ',').ToString()
					}
					else
					{
						$generic = $false
						$genOpsdetails = 'None'
					}
					
				}
				elseif ($recipient.RecipientTypeDetails -match 'SharedMailbox')
				{
					
					$ad = get-AdUser $recipient.SamAccountName -properties * -ErrorAction Stop
					$Allgroups = $ad.MemberOf.ForEach({ $_.Split(',')[0].Split('=')[1] })
					$groups = @()
					
					foreach ($g in $Allgroups)
					{
						if ($genericGroups -icontains $g)
						{
							$groups += $g
						}
					}
					if ($groups -ne $null)
					{
						$generic = $true
						
						if ($groups -contains 'ACT-SaaS-O365-GenericAccountOfficeProPlusOnly')
						{
							$paramRemoveQHE3ADGrpmembershipGen = @{
								AdGroup	       = 'ACT-SaaS-O365-GenericAccountOfficeProPlusOnly'
								SamAccountName = $recipient.SamAccountName
							}
							
							$GenOps += Remove-QHE3ADGrpmembership @paramRemoveQHE3ADGrpmembershipGen
							
							if ($groups -notcontains 'ACT-SaaS-O365-GenericAccountOfficeProPlusAndMailbox')
							{
								$paramAddQHGenericGroupMember = @{
									SamAccountName = $recipient.SamAccountName
									AdGroup	       = 'ACT-SaaS-O365-GenericAccountOfficeProPlusAndMailbox'
								}
								
								$GenOps += Add-QHGenericGroupMember @paramAddQHGenericGroupMember
							}
							
							if ($Allgroups -icontains 'ADM-O365-LIC-E3-OfficeProPlusOnly')
							{
								$paramRemoveQHE3ADGrpmembership = @{
									AdGroup	       = 'ADM-O365-LIC-E3-OfficeProPlusOnly'
									SamAccountName = $recipient.SamAccountName
								}
								
								$GenOps += Remove-QHE3ADGrpmembership @paramRemoveQHE3ADGrpmembership
							}
							
						}
						else
						{
							$GenOps = $groups
						}
						
						$genOpsdetails = $($GenOps -join ',').ToString()
					}
					else
					{
						if ($ad.Enabled)
						{
							$generic = $true
							
							if ($Allgroups -icontains 'ADM-O365-LIC-E3-OfficeProPlusOnly')
							{
								$paramRemoveQHE3ADGrpmembership = @{
									AdGroup	       = 'ADM-O365-LIC-E3-OfficeProPlusOnly'
									SamAccountName = $recipient.SamAccountName
								}
								
								$GenOps += Remove-QHE3ADGrpmembership @paramRemoveQHE3ADGrpmembership
							}
							
							if ($groups -notcontains 'ACT-SaaS-O365-GenericAccountMailboxOnly')
							{
								$paramAddQHGenericGroupMember = @{
									SamAccountName = $recipient.SamAccountName
									AdGroup	       = 'ACT-SaaS-O365-GenericAccountMailboxOnly'
								}
								
								$GenOps += Add-QHGenericGroupMember @paramAddQHGenericGroupMember
							}
							
							$genOpsdetails = $($GenOps -join ',').ToString()
							
						}
						else
						{
							$generic = $false
							$genOpsdetails = 'None'
						}
					}
				}
				else
				{
					$generic = $false
					# not a user or Shared Mailbox
					$genOpsdetails = 'None'
				}
				
				if ($generic)
				{
					$lic = Get-iQHLicenseStatus -UserPrincipalName $UPN
					
					$directLicenses = $lic | Get-Member -MemberType NoteProperty | Where-Object { $_.Definition -match 'Direct' } |
					Select-Object -ExpandProperty Name
					
					$PlansToDisable = switch ($directLicenses)
					{
						'E3' { 'healthqld:ENTERPRISEPACK' }
						'E1' { 'healthqld:STANDARDPACK' }
						'MFA'{ 'healthqld:MFA_STANDALONE' }
						'PowerBi Free' { 'healthqld:POWER_BI_STANDARD' }
						'ExchangeOnlinePlan2' { 'healthqld:EXCHANGEENTERPRISE' }
						'Enterprise Mobility Security' { 'healthqld:EMS' }
						'Microsoft Flow Free' { 'healthqld:FLOW_FREE' }
						'PowerApps and Logic Flows' { 'healthqld:POWERAPPS_INDIVIDUAL_USER' }
						'Phone System' { 'healthqld:MCOEV' }
						'PowerBi Pro' { 'healthqld:POWER_BI_PRO' }
						'Power BI for Office 365 Add-On' { 'healthqld:POWER_BI_ADDON' }
						'Power BI Individual User' { 'healthqld:POWER_BI_INDIVIDUAL_USER' }
						'Enterprise Plan E4' { 'healthqld:ENTERPRISEWITHSCAL' }
						'Project Online' { 'healthqld:PROJECTONLINE_PLAN_1' }
						'Project Pro for Office 365' { 'healthqld:PROJECTCLIENT' }
						'Visio Pro Online' { 'healthqld:VISIOCLIENT' }
						'Microsoft Stream' { 'healthqld:STREAM' }
						'Microsoft Power Apps & Flow' { 'healthqld:POWERAPPS_VIRAL' }
						'Project Lite' { 'healthqld:PROJECTESSENTIALS' }
						'Project Professional' { 'healthqld:PROJECTPROFESSIONAL' }
						'App Connect' { 'healthqld:SPZA_IW' }
						'Power Bi Premium' { 'healthqld:PBI_PREMIUM_P1_ADDON' }
						'Dynamics 365 P1 Trial for Information Workers' { 'healthqld:DYN365_ENTERPRISE_P1_IW' }
						default
						{
							#<code>
						}
					}
					
					if ($PlansToDisable -ne $null)
					{
						Set-MsolUserLicense -UserPrincipalName $UPN -RemoveLicenses $PlansToDisable -ErrorAction Stop
						$PlansToDisable = $PlansToDisable -join ','
					}
					else
					{
						$PlansToDisable = 'None'
					}
				}
				else
				{
					$PlansToDisable = "None"
				}
				
				$prop = [Ordered]@{
					UserPrincipalName    = $UPN
					RecipientTypeDetails = $recipient.RecipientTypeDetails
					SamAccountName	     = $ad.SamAccountName
					ADAccountEnabled	 = $ad.Enabled
					IsGeneric		     = $generic
					GenericDetails	     = $genOpsdetails
					DisabledDirectLicenses = $PlansToDisable
					Status			     = 'Success'
					Details			     = 'None'
				}
			}
			catch
			{
				$prop = [Ordered]@{
					UserPrincipalName    = $UPN
					RecipientTypeDetails = $recipient.RecipientTypeDetails
					SamAccountName	     = $ad.SamAccountName
					ADAccountEnabled	 = $ad.Enabled
					IsGeneric		     = $generic
					GenericDetails	     = $genOpsdetails
					DisabledDirectLicenses = $PlansToDisable
					Status			     = 'Failed'
					Details			     = "ERROR : $($_.Exception.Message)"
				}
			}
			finally
			{
				$obj = New-Object -TypeName System.Management.Automation.PSObject -Property $prop
				Write-Output $obj
				if ($UserPrincipalName.count -gt 1)
				{
					if ($ShowProgress)
					{
						$paramWriteProgress = @{
							Activity = 'Doing Some Processing'
							Status   = "Processing [$i] of [$($UserPrincipalName.Count)] users"
							PercentComplete = (($i / $UserPrincipalName.Count) * 100)
							CurrentOperation = "Completed : [$UPN]"
						}
						Write-Progress @paramWriteProgress
					}
				}
				$i++
			}
		}
	}
	end
	{
		Write-Progress -Activity 'Doing Some Processing' -Completed
		$Global:ErrorActionPreference = [System.Management.Automation.ActionPreference]::Continue
	}
}

function Get-QHUSerInfo
{
<#
	.SYNOPSIS
		A brief description of the Get-QHUSerInfo function.
	
	.DESCRIPTION
		A detailed description of the Get-QHUSerInfo function.
	
	.PARAMETER UserName
		A description of the UserName parameter.
	
	.PARAMETER ShowProgress
		A description of the ShowProgress parameter.
	
	.EXAMPLE
		PS C:\> Get-QHUSerInfo -UserName 'value1'
	
	.NOTES
		Additional information about the function.
#>
	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory = $true)]
		[String[]]$UserName,
		[Switch]$ShowProgress
	)
	
	begin
	{
		$i = 1
	}
	process
	{
		foreach ($user in $UserName)
		{
			try
			{
				$info = Get-aduser -Filter { UserPrincipalName -like $user } -Properties * -ErrorAction Stop
				if ($info -eq $null)
				{
					throw "Object not found"
				}
				$prop = [ordered]@{
					UserPrincipalName = $user
					SamAccountName    = $info.SamAccountName
					AccountEnabled    = $info.Enabled
					DisplayName	      = $info.DisplayName
					GivenName		  = $info.GivenName
					SurName		      = $info.SurName
					WhenCreated	      = $info.WhenCreated
					PasswordExpired   = $info.PasswordExpired
					PasswordLastSet   = $info.PasswordLastSet
					LastLogonDate	  = $info.LastLogonDate
					DN			      = $info.DistinguishedName
					LanCostCenter	  = $info.AUQHLANCX
					CostCenter	      = $info.AUQHCostCentreInternet
					EmployeeType	  = $info.employeeType
					EmployeeID	      = $info.employeeID
					Title			  = $info.title
					Details		      = 'None'
				}
			}
			catch
			{
				$prop = [ordered]@{
					UserPrincipalName = $user
					SamAccountName    = $info.SamAccountName
					AccountEnabled    = $info.Enabled
					DisplayName	      = $info.DisplayName
					GivenName		  = $info.GivenName
					SurName		      = $info.SurName
					WhenCreated	      = $info.WhenCreated
					PasswordExpired   = $info.PasswordExpired
					PasswordLastSet   = $info.PasswordLastSet
					LastLogonDate	  = $info.LastLogonDate
					DN			      = $info.DistinguishedName
					LanCostCenter	  = $info.AUQHLANCX
					CostCenter	      = $info.AUQHCostCentreInternet
					EmployeeType	  = $info.employeeType
					EmployeeID	      = $info.employeeID
					Title			  = $info.title
					Details		      = "ERROR : $($_.Exception.Message)"
				}
			}
			finally
			{
				$obj = New-Object -TypeName PSobject -Property $prop
				Write-Output $obj
				if ($ShowProgress)
				{
					if ($user.count -gt 5)
					{
						$paramWriteProgress = @{
							Activity = 'Getting User Informaiton'
							Status   = "Processing [$i] of [$($UserName.Count)] users"
							PercentComplete = (($i / $UserName.Count) * 100)
							CurrentOperation = "Completed : [$user]"
						}
						Write-Progress @paramWriteProgress
						
					}
				}
				
				$i++
			}
		}
		
	}
	end
	{
		Write-Progress -Activity 'Getting User Informaiton' -Completed
	}
}

function get-QhMsolUserErrorDetails
{
<#
	.SYNOPSIS
		A brief description of the get-QhMsolUserErrorDetails function.
	
	.DESCRIPTION
		A detailed description of the get-QhMsolUserErrorDetails function.
	
	.PARAMETER UserPrincipalName
		A description of the UserPrincipalName parameter.
	
	.EXAMPLE
		PS C:\> get-QhMsolUserErrorDetails -UserPrincipalName 'value1'
	
	.NOTES
		Additional information about the function.
#>
	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory = $true,
				   ValueFromPipeline = $true,
				   ValueFromPipelineByPropertyName = $true,
				   Position = 0)]
		[ValidateNotNullOrEmpty()]
		[String[]]$UserPrincipalName
	)
	
	begin
	{
		try
		{
			$null = Get-MsolAccountSku -ErrorAction Stop
		}
		catch
		{
			Write-Warning "Please Connect to MsOnline before running this command"
			break
		}
	}
	process
	{
		foreach ($UPN in $UserPrincipalName)
		{
			try
			{
				$msolUser = Get-MsolUser -UserPrincipalName $UPN -ErrorAction Stop
				$err = $msolUser.Errors
				if ($err -ne $null)
				{
					$errDetails = $msolUser.Errors.ErrorDetail.ObjectErrors.ErrorRecord.ErrorDescription
				}
				else
				{
					$errDetails = "NoErrors"
				}
				$Prop = [ordered] @{
					UserPrincipalName = $UPN
					UserErrorDetails  = $errDetails
					Details		      = 'None'
				}
			}
			catch
			{
				$Prop = [ordered] @{
					UserPrincipalName = $UPN
					UserErrorDetails  = 'None'
					Details		      = "ERROR : $($_.Exception.Message)"
				}
			}
			finally
			{
				$obj = New-Object -TypeName System.Management.Automation.PSObject -Property $Prop
				Write-Output $obj
			}
		}
	}
	end
	{
	}
}

function Get-QHLicenseStatusForPowerBi #ForPowerBiReporting
{
<#
	.SYNOPSIS
		A brief description of the Get-QHLicenseStatusForPowerBi function.
	
	.DESCRIPTION
		A detailed description of the Get-QHLicenseStatusForPowerBi function.
	
	.PARAMETER UserPrincipalName
		A description of the UserPrincipalName parameter.
	
	.PARAMETER ShowProgress
		A description of the ShowProgress parameter.
	
	.EXAMPLE
		PS C:\> Get-QHLicenseStatusForPowerBi -UserPrincipalName 'value1'
	
	.NOTES
		Additional information about the function.
#>
	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory = $true,
				   ValueFromPipeline = $true,
				   ValueFromPipelineByPropertyName = $true,
				   Position = 0,
				   HelpMessage = 'Enter UserPrincipal Name ?')]
		[String[]]$UserPrincipalName,
		[switch]$ShowProgress
	)
	
	begin
	{
		$i = 1
		$Global:ErrorActionPreference = [System.Management.Automation.ActionPreference]::Stop
		$Plans = Get-MsolAccountSku -ErrorAction Stop | Select-Object -ExpandProperty AccountSkuId
		$date = (Get-Date).ToString('dd-MM-yyyy')
	}
	process
	{
		foreach ($UPN in $UserPrincipalName)
		{
			try
			{
				$msol = Get-MsolUser -UserPrincipalName $UPN -ErrorAction Stop
				
				$prop = [ordered]@{
					ReportDate	      = $date
					UserPrincipalName = $UPN
				}
				
				foreach ($Plan in $Plans)
				{
					$lic = $msol.Licenses.Where({ $Plan -eq ($_ | Select-Object AccountSkuId | Select-Object -ExpandProperty AccountSkuId) })
					if ($lic -ne $null)
					{
						$licobj = $lic | Select-Object @{
							n = 'AccountSku'; e = {
								switch ($lic | Select-Object AccountSkuid | Select-Object -ExpandProperty AccountSkuid)
								{
									'healthqld:ENTERPRISEPACK' { 'E3' }
									'healthqld:STANDARDPACK' { 'E1' }
									'healthqld:MFA_STANDALONE' { 'MFA' }
									'healthqld:POWER_BI_STANDARD' { 'PowerBi Free' }
									'healthqld:EXCHANGEENTERPRISE' { 'ExchangeOnlinePlan2' }
									'healthqld:EMS' { 'Enterprise Mobility Security' }
									'healthqld:FLOW_FREE' { 'Microsoft Flow Free' }
									'healthqld:POWERAPPS_INDIVIDUAL_USER' { 'PowerApps and Logic Flows' }
									'healthqld:MCOEV' { 'Phone System' }
									'healthqld:POWER_BI_PRO' { 'PowerBi Pro' }
									'healthqld:POWER_BI_ADDON' { 'Power BI for Office 365 Add-On' }
									'healthqld:POWER_BI_INDIVIDUAL_USER' { 'Power BI Individual User' }
									'healthqld:ENTERPRISEWITHSCAL' { 'Enterprise Plan E4' }
									'healthqld:PROJECTONLINE_PLAN_1' { 'Project Online' }
									'healthqld:PROJECTCLIENT' { 'Project Pro for Office 365' }
									'healthqld:VISIOCLIENT' { 'Visio Pro Online' }
									'healthqld:STREAM' { 'Microsoft Stream' }
									'healthqld:POWERAPPS_VIRAL' { 'Microsoft Power Apps & Flow' }
									'healthqld:PROJECTESSENTIALS' { 'Project Lite' }
									'healthqld:PROJECTPROFESSIONAL' { 'Project Professional' }
									'healthqld:SPZA_IW' { 'App Connect' }
									'healthqld:PBI_PREMIUM_P1_ADDON' { 'Power Bi Premium' }
									'healthqld:DYN365_ENTERPRISE_P1_IW' { 'Dynamics 365 P1 Trial for Information Workers' }
									default { "$_" }
								}
							}
						},
													   @{
							n = 'Assignment'; e = {
								1
							}
						}
					}
					else
					{
						$lic = switch ($Plan)
						{
							'healthqld:ENTERPRISEPACK' { 'E3' }
							'healthqld:STANDARDPACK' { 'E1' }
							'healthqld:MFA_STANDALONE' { 'MFA' }
							'healthqld:POWER_BI_STANDARD' { 'PowerBi Free' }
							'healthqld:EXCHANGEENTERPRISE' { 'ExchangeOnlinePlan2' }
							'healthqld:EMS' { 'Enterprise Mobility Security' }
							'healthqld:FLOW_FREE' { 'Microsoft Flow Free' }
							'healthqld:POWERAPPS_INDIVIDUAL_USER' { 'PowerApps and Logic Flows' }
							'healthqld:MCOEV' { 'Phone System' }
							'healthqld:POWER_BI_PRO' { 'PowerBi Pro' }
							'healthqld:POWER_BI_ADDON' { 'Power BI for Office 365 Add-On' }
							'healthqld:POWER_BI_INDIVIDUAL_USER' { 'Power BI Individual User' }
							'healthqld:ENTERPRISEWITHSCAL' { 'Enterprise Plan E4' }
							'healthqld:PROJECTONLINE_PLAN_1' { 'Project Online' }
							'healthqld:PROJECTCLIENT' { 'Project Pro for Office 365' }
							'healthqld:VISIOCLIENT' { 'Visio Pro Online' }
							'healthqld:STREAM' { 'Microsoft Stream' }
							'healthqld:POWERAPPS_VIRAL' { 'Microsoft Power Apps & Flow' }
							'healthqld:PROJECTESSENTIALS' { 'Project Lite' }
							'healthqld:PROJECTPROFESSIONAL' { 'Project Professional' }
							'healthqld:SPZA_IW' { 'App Connect' }
							'healthqld:PBI_PREMIUM_P1_ADDON' { 'Power Bi Premium' }
							'healthqld:DYN365_ENTERPRISE_P1_IW' { 'Dynamics 365 P1 Trial for Information Workers' }
							default { "$_" }
						}
						
						$licobj = [PScustomobject]@{
							AccountSku = $lic
							Assignment = 0
						}
					}
					
					$prop.Add($licobj.AccountSku, $licobj.Assignment)
					
				}
				$prop.Add('Details', 'None')
			}
			catch
			{
				$prop = [ordered]@{
					ReportDate	      = $date
					UserPrincipalName = $UPN
					Details		      = "ERROR : $($_.Exception.Message)"
				}
			}
			finally
			{
				$obj = New-Object -TypeName System.Management.Automation.PSObject -Property $prop
				Write-Output $obj
				
				if ($UserPrincipalName.count -gt 1)
				{
					if ($ShowProgress)
					{
						$paramWriteProgress = @{
							Activity = 'Getting MFA License Status'
							Status   = "Processing [$i] of [$($UserPrincipalName.Count)] users"
							PercentComplete = (($i / $UserPrincipalName.Count) * 100)
							CurrentOperation = "Completed : [$UPN]"
						}
						Write-Progress @paramWriteProgress
					}
				}
				$i++
			}
		}
	}
	end
	{
		Write-Progress -Activity 'Getting MFA License Status' -Completed
		$Global:ErrorActionPreference = [System.Management.Automation.ActionPreference]::Continue
	}
}

function Enable-QHMFA
{
<#
	.SYNOPSIS
		A brief description of the Enable-QHMFA function.
	
	.DESCRIPTION
		A detailed description of the Enable-QHMFA function.
	
	.PARAMETER UserPrincipalName
		A description of the UserPrincipalName parameter.
	
	.PARAMETER ShowProgress
		A description of the ShowProgress parameter.
	
	.EXAMPLE
		PS C:\> Enable-QHMFA -UserPrincipalName 'value1'
	
	.NOTES
		Additional information about the function.
#>
	
	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory = $true,
				   ValueFromPipeline = $true,
				   ValueFromPipelineByPropertyName = $true,
				   Position = 0)]
		[ValidateNotNullOrEmpty()]
		[String[]]$UserPrincipalName,
		[Switch]$ShowProgress
	)
	
	begin
	{
		$global:ErrorActionPreference = [System.Management.Automation.ActionPreference]::Stop
		$i = 1
	}
	process
	{
		foreach ($UPN in $UserPrincipalName)
		{
			try
			{
				$recipient = Get-Recipient $UPN -ErrorAction Stop
				if ($recipient.RecipientTypeDetails -match 'UserMailbox')
				{
					$auth = New-Object -TypeName Microsoft.Online.Administration.StrongAuthenticationRequirement
					$auth.RelyingParty = "*"
					$auth.State = "Enabled"
					$auth.RememberDevicesNotIssuedBefore = (Get-Date)
					
					Set-MsolUser -UserPrincipalName $UPN -StrongAuthenticationRequirements $auth -ErrorAction Stop
					$prop = [ordered]@{
						EmailAddress		 = $UPN
						RecipientTypeDetails = $recipient.RecipientTypeDetails
						Status			     = 'SUCCESS'
						Details			     = 'None'
					}
				}
				else
				{
					$prop = [ordered]@{
						EmailAddress		 = $UPN
						RecipientTypeDetails = $recipient.RecipientTypeDetails
						Status			     = 'SKIPPED'
						Details			     = "SKIPPED : Not a User Mailbox"
					}
				}
			}
			catch
			{
				$prop = [ordered]@{
					EmailAddress		 = $UPN
					RecipientTypeDetails = 'ERROR'
					Status			     = 'FAILED'
					Details			     = "ERROR : $($_.Exception.Message)"
				}
			}
			finally
			{
				$obj = New-Object -TypeName System.Management.Automation.PSObject -Property $prop
				Write-Output $obj
				
				if ($ShowProgress)
				{
					if ($UserPrincipalName.count -gt 1)
					{
						$paramWriteProgress = @{
							Activity = 'Enabling MFA'
							Status   = "Processing [$i] of [$($UserPrincipalName.Count)] users"
							PercentComplete = (($i / $UserPrincipalName.Count) * 100)
							CurrentOperation = "Completed : [$UPN]"
						}
						Write-Progress @paramWriteProgress
					}
				}
				$i++
			}
		}
	}
	end
	{
		Write-Progress -Activity 'Enabling MFA' -Completed
		$global:ErrorActionPreference = [System.Management.Automation.ActionPreference]::Continue
	}
}

function Reset-QHMFA
{
<#
	.SYNOPSIS
		A brief description of the Reset-QhMFASettings function.
	
	.DESCRIPTION
		A detailed description of the Reset-QhMFASettings function.
	
	.PARAMETER UserPrincipalName
		A description of the UserPrincipalName parameter.
	
	.PARAMETER ShowProgress
		A description of the ShowProgress parameter.
	
	.EXAMPLE
		PS C:\> Reset-QhMFASettings -UserPrincipalName 'value1'
	
	.NOTES
		Additional information about the function.
#>
	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory = $true,
				   ValueFromPipeline = $true,
				   ValueFromPipelineByPropertyName = $true,
				   Position = 0)]
		[ValidateNotNullOrEmpty()]
		[String[]]$UserPrincipalName,
		[Parameter(Position = 1)]
		[switch]$ShowProgress
	)
	
	begin
	{
		$global:ErrorActionPreference = [System.Management.Automation.ActionPreference]::Stop
		$i = 1
	}
	process
	{
		foreach ($UPN in $UserPrincipalName)
		{
			try
			{
				$mfa = @()
				$paramSetMsolUser = @{
					UserPrincipalName		    = $UPN
					StrongAuthenticationMethods = $mfa
					ErrorAction				    = 'Stop'
				}
				
				Set-MsolUser @paramSetMsolUser
				
				$prop = [Ordered]@{
					UserPrincipalName = $UPN
					MFAReset		  = 'Success'
					Details		      = 'None'
				}
			}
			catch
			{
				$prop = [Ordered]@{
					UserPrincipalName = $UPN
					MFAReset		  = 'Failed'
					Details		      = "Error : $($_.Exception.Message)"
				}
			}
			finally
			{
				$obj = New-Object -TypeName System.Management.Automation.PSObject -Property $prop
				Write-Output $obj
				if ($ShowProgress)
				{
					if ($UserPrincipalName.count -gt 1)
					{
						$paramWriteProgress = @{
							Activity = 'Resetting MFA'
							Status   = "Processing [$i] of [$($UserPrincipalName.Count)] users"
							PercentComplete = (($i / $UserPrincipalName.Count) * 100)
							CurrentOperation = "Completed : [$UPN]"
						}
						Write-Progress @paramWriteProgress
					}
				}
				$i++
			}
		}
	}
	end
	{
		$global:ErrorActionPreference = [System.Management.Automation.ActionPreference]::Continue
		Write-Progress -Activity 'Resetting MFA' -Completed
	}
}

function Disable-QHMFA
{
<#
	.SYNOPSIS
		A brief description of the Disable-QHMFA function.
	
	.DESCRIPTION
		A detailed description of the Disable-QHMFA function.
	
	.PARAMETER UserPrincipalName
		A description of the UserPrincipalName parameter.
	
	.PARAMETER ShowProgress
		A description of the ShowProgress parameter.
	
	.EXAMPLE
		PS C:\> Disable-QHMFA -UserPrincipalName 'value1'
	
	.NOTES
		Additional information about the function.
#>
	
	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory = $true,
				   ValueFromPipeline = $true,
				   ValueFromPipelineByPropertyName = $true,
				   Position = 0)]
		[ValidateNotNullOrEmpty()]
		[String[]]$UserPrincipalName,
		[Switch]$ShowProgress
	)
	begin
	{
		$global:ErrorActionPreference = [System.Management.Automation.ActionPreference]::Stop
		$i = 1
	}
	process
	{
		foreach ($UPN in $UserPrincipalName)
		{
			try
			{
				$recipient = Get-Recipient $UPN -ErrorAction Stop
				if ($recipient.RecipientTypeDetails -match 'UserMailbox')
				{
					$auth = @()
					
					$paramSetMsolUser = @{
						UserPrincipalName			     = $UPN
						StrongAuthenticationRequirements = $auth
						ErrorAction					     = 'Stop'
					}
					
					Set-MsolUser @paramSetMsolUser
					$prop = [ordered]@{
						EmailAddress		 = $UPN
						RecipientTypeDetails = $recipient.RecipientTypeDetails
						Status			     = 'SUCCESS'
						Details			     = 'None'
					}
				}
				else
				{
					$prop = [ordered]@{
						EmailAddress		 = $UPN
						RecipientTypeDetails = $recipient.RecipientTypeDetails
						Status			     = 'SKIPPED'
						Details			     = "SKIPPED : Not a User Mailbox"
					}
				}
			}
			catch
			{
				$prop = [ordered]@{
					EmailAddress		 = $UPN
					RecipientTypeDetails = 'ERROR'
					Status			     = 'FAILED'
					Details			     = "ERROR : $($_.Exception.Message)"
				}
			}
			finally
			{
				$obj = New-Object -TypeName System.Management.Automation.PSObject -Property $prop
				Write-Output $obj
				
				if ($ShowProgress)
				{
					if ($UserPrincipalName.count -gt 1)
					{
						$paramWriteProgress = @{
							Activity = 'Enabling MFA'
							Status   = "Processing [$i] of [$($UserPrincipalName.Count)] users"
							PercentComplete = (($i / $UserPrincipalName.Count) * 100)
							CurrentOperation = "Completed : [$UPN]"
						}
						Write-Progress @paramWriteProgress
					}
				}
				$i++
			}
		}
	}
	end
	{
		Write-Progress -Activity 'Enabling MFA' -Completed
		$global:ErrorActionPreference = [System.Management.Automation.ActionPreference]::Continue
	}
}

function Get-QHMFAStatus
{
<#
	.SYNOPSIS
		A brief description of the Get-QHMFAStatus function.
	
	.DESCRIPTION
		A detailed description of the Get-QHMFAStatus function.
	
	.PARAMETER UserPrincipalName
		A description of the UserPrincipalName parameter.
	
	.PARAMETER ShowProgress
		A description of the ShowProgress parameter.
	
	.EXAMPLE
		PS C:\> Get-QHMFAStatus -UserPrincipalName 'value1'
	
	.NOTES
		Additional information about the function.
#>
	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory = $true,
				   ValueFromPipeline = $true,
				   ValueFromPipelineByPropertyName = $true,
				   Position = 0,
				   HelpMessage = 'Enter UserPrincipal Name ?')]
		[ValidateNotNullOrEmpty()]
		[String[]]$UserPrincipalName,
		[switch]$ShowProgress
	)
	
	begin
	{
		$i = 1
		$Global:ErrorActionPreference = [System.Management.Automation.ActionPreference]::Stop
	}
	process
	{
		foreach ($UPN in $UserPrincipalName)
		{
			try
			{
				$msol = Get-MsolUser -UserPrincipalName $UPN -ErrorAction Stop
				if ($MSOL.StrongAuthenticationRequirements.State -eq $null)
				{
					$state = 'PotentiallyUnlicensed'
				}
				else
				{
					$state = $MSOL.StrongAuthenticationRequirements.State
				}
				
				if ($msol.StrongAuthenticationMethods.Count -eq 0)
				{
					$MFASetup = 'NotRegistered'
					$details = 'None'
				}
				else
				{
					$MFASetup = 'Registered'
					$details = "Default MFA Method : $(($msol.StrongAuthenticationMethods.Where({ $_.'IsDefault' })).MethodType)"
				}
				
				$prop = [ordered]@{
					UserPrincipalName = $UPN
					MFAState		  = $state
					MFARegistration   = $MFASetup
					Details		      = $details
				}
			}
			catch
			{
				$prop = [ordered]@{
					UserPrincipalName = $UPN
					MFAState		  = 'ERROR'
					MFARegistration   = 'ERROR'
					Details		      = "ERROR : $($_.Exception.Message)"
				}
			}
			finally
			{
				$obj = New-Object -TypeName System.Management.Automation.PSObject -Property $prop
				Write-Output $obj
				
				if ($UserPrincipalName.count -gt 1)
				{
					if ($ShowProgress)
					{
						$paramWriteProgress = @{
							Activity = 'Getting MFA License Status'
							Status   = "Processing [$i] of [$($UserPrincipalName.Count)] users"
							PercentComplete = (($i / $UserPrincipalName.Count) * 100)
							CurrentOperation = "Completed : [$UPN]"
						}
						Write-Progress @paramWriteProgress
					}
				}
				$i++
			}
		}
	}
	end
	{
		Write-Progress -Activity 'Getting MFA License Status' -Completed
		$Global:ErrorActionPreference = [System.Management.Automation.ActionPreference]::Continue
	}
}

function Write-QHLog
{
<#
	.SYNOPSIS
		A brief description of the Write-QHLog function.
	
	.DESCRIPTION
		A detailed description of the Write-QHLog function.
	
	.PARAMETER Type
		A description of the Type parameter.
	
	.PARAMETER Message
		A description of the Message parameter.
	
	.PARAMETER OnScreen
		A description of the OnScreen parameter.
	
	.PARAMETER Function
		A description of the Function parameter.
	
	.PARAMETER seperator
		A description of the seperator parameter.
	
	.EXAMPLE
		PS C:\> Write-QHLog -Type INFO -Message 'value2'
	
	.NOTES
		Additional information about the function.
#>
	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory = $true,
				   Position = 0)]
		[ValidateSet('INFO', 'ERROR', 'SUCCESS')]
		[string]$Type,
		[Parameter(Mandatory = $true,
				   Position = 1)]
		[String]$Message,
		[Parameter(Position = 2)]
		[switch]$OnScreen,
		[String]$Function = $($MyInvocation.InvocationName),
		[String]$seperator = '::'
	)
	begin
	{
		function seperator
		{
			param
			(
				[Parameter(Mandatory = $true)]
				[ValidateNotNullOrEmpty()]
				[String]$char
			)
			
			Write-Host -NoNewline " $char " -ForegroundColor Magenta
		}
	}
	process
	{
		$time = (Get-date).ToString('dd-MM-yyyy HH:mm:ss')
		
		$prop = [ordered]@{
			DateTime = $time
			type	 = $Type
			Function = $Function
			Details  = $Message
		}
		
		if ($OnScreen)
		{
			switch ($Type)
			{
				'INFO' {
					$col = 'Yellow'
				}
				'ERROR' {
					$col = 'Red'
				}
				'SUCCESS' {
					$col = 'Green'
				}
				default
				{
					#<code>
				}
			}
			
			#$StringMsg = "$time :: $Type :: $message"
			Write-Host "$time" -ForegroundColor Gray -NoNewline
			seperator -char $seperator
			Write-Host "$Type" -ForegroundColor $col -NoNewline
			seperator -char $seperator
			Write-Host "$Function" -ForegroundColor White -NoNewline
			seperator -char $seperator
			Write-Host $Message -ForegroundColor Cyan
		}
		$obj = New-Object -TypeName System.Management.Automation.PSObject -Property $prop
		Write-Output $obj
	}
	end
	{
	}
}

function Set-QHDistributionGroupToSMBX
{
<#
	.SYNOPSIS
		A brief description of the Set-QHDistributionGroupToSMBX function.
	
	.DESCRIPTION
		A detailed description of the Set-QHDistributionGroupToSMBX function.
	
	.PARAMETER DLName
		A description of the DLName parameter.
	
	.PARAMETER DomainController
		A description of the DomainController parameter.
	
	.EXAMPLE
		PS C:\> Set-QHDistributionGroupToSMBX -DLName 'value1' -DomainController EAD-WDCBTPP01
	
	.NOTES
		Additional information about the function.
#>
	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory = $true)]
		[String]$DLName,
		[Parameter(Mandatory = $true)]
		[ValidateSet('EAD-WDCBTPP01', 'EAD-WDCBK7P01', 'EAD-WDCBK7P03', 'EAD-WDCBK7P02', 'EAD-WDCBTPP02', 'EAD-WDCBTPP03', 'EAD-WDCBTPP04', 'EAD-WDCBK7P04', 'EAD-WDCBK7P05', 'EAD-WDCBTPP05', 'EAD-WDCBK7P06', 'EAD-WDCBTPP06', 'EAD-WDCBK7P07', 'EAD-WDCBTPP07', 'EAD-WDCBK7P08', 'EAD-WDCBTPP08')]
		[ValidateNotNullOrEmpty()]
		[String]$DomainController
	)
	
	begin
	{
		$Global:ErrorActionPreference = 'Stop'
		Write-Verbose "Testing Connection to Domain Controller $DomainController"
		if (!(Test-Connection $DomainController -Count 1 -Quiet))
		{
			Write-Verbose "Connection to Domain Controller $DomainController Failed, Exiting function"
			Write-Warning "$DomainController not Reachable, please try again with a different domain controller"
			break
		}
		Write-Verbose "Connection to Domain controller $DomainController OK"
		$SMPassword = ConvertTo-SecureString "P@ssW0rD!" -AsPlainText -Force
	}
	process
	{
		try
		{
			$group = $null
			Write-Verbose "Trying to get $DL"
			$group = Get-DistributionGroup $DLName -ErrorAction Stop
			Write-Verbose "Found DL $DL"
			Write-Verbose "Trying to Get Dynamic Distribution group members of $DL"
			$members = @()
			$members += Get-DistributionGroupMember -Identity $group.Name -ErrorAction Stop | Where-Object RecipientTypeDetails -eq DynamicDistributionGroup
			
			if ($members.count -gt 0)
			{
				Write-Verbose "Got $($members.count) members"
				foreach ($mem in $members)
				{
					$DDSInfo = Get-DynamicDistributionGroup $mem.Name
					$DDLForwardAddress = "qh.health.qld.gov.au/Queensland Health/Enterprise Groups/Distribution Lists/" + $mem.Name
					$smName = $DDSInfo.Name + "-SM"
					$smUpn = $smName + "@health.qld.gov.au"
					$ou = "qh.health.qld.gov.au/Queensland Health/User Accounts/resources"
					
					Write-Verbose "Trying to Check if the Shared Mailbox Already exists"
					$Checkmbx = get-Mailbox $smName -ErrorAction SilentlyContinue
					
					if ($Checkmbx -eq $null)
					{
						Write-Verbose "No Shared mailbox Exists for $mem"
						try
						{
							$paramNewMailbox = @{
								Name			   = $SMName
								Alias			   = $SMName
								UserPrincipalName  = $SMUPN
								Shared			   = $true
								Password		   = $SMPassword
								OrganizationalUnit = $ou
								Database		   = "Database120"
								DomainController   = $DomainController
								ErrorAction	       = 'Stop'
							}
							Write-Verbose "Trying to Create Shared Mailbox for $smName"
							New-Mailbox @paramNewMailbox | Out-Null
							Write-Verbose "Created Shared Mailbox with $smName"
							$mbxCreation = "Created:$smName"
						}
						catch
						{
							Write-Verbose "Failed to create mailbox with $smName"
							$mbxCreation = "Failed:$smName Error: $($_.Exception.Message)"
						}
					}
					else
					{
						Write-Verbose "Mailbox already exists with $smName"
						$mbxCreation = "Existed:$smName"
					}
					
					if ($mbxCreation -notmatch 'Failed')
					{
						try
						{
							Write-Verbose "Setting Forwarding Address to $DDLForwardAddress and max receive size to $($DDSInfo.MaxReceiveSize)"
							$paramSetMailbox = @{
								Identity		  = $smName
								ForwardingAddress = $DDLForwardAddress
								domainController  = $DomainController
								MaxReceiveSize    = $DDSInfo.MaxReceiveSize
							}
							
							Set-Mailbox @paramSetMailbox
							Write-Verbose "Added forwarding address : $DDLForwardAddress"
							$fwdAdd = "Added:$DDLForwardAddress"
						}
						catch
						{
							Write-Verbose "Failed to add forwarding address : $DDLForwardAddress"
							$fwdAdd = "Failed:$($_.Exception.Message)"
						}
						
						try
						{
							Write-Verbose "Trying to add $smUpn to DL $DLName"
							$paramAddDistributiongroupMember = @{
								identity					    = $DLName
								member						    = $smUpn
								BypassSecurityGroupManagerCheck = $true
								ErrorAction					    = 'Stop'
								DomainController			    = $DomainController
							}
							
							Add-DistributiongroupMember @paramAddDistributiongroupMember
							
							Write-Verbose "Added $smUpn to $DLName"
							
							$DLMem = "Added:$smUpn to $DLName"
						}
						catch
						{
							Write-Verbose "Error Adding $smUpn to $DLName"
							$DLMem = "Failed:$($_.Exception.Message)"
						}
						
						try
						{
							Write-Verbose "Trying to mark $($mem.name) as hidden"
							
							$paramSetDynamicDistributionGroup = @{
								Identity = $DDSInfo.PrimarySmtpAddress.ToString()
								HiddenFromAddressListsEnabled = $true
								ErrorAction = 'Stop'
								DomainController = $DomainController
							}
							
							Set-DynamicDistributionGroup @paramSetDynamicDistributionGroup
							Write-Verbose "$($DDSInfo.PrimarySmtpAddress.ToString()) marked as hidden"
							$DDLHidden = "$($DDSInfo.PrimarySmtpAddress.ToString()):HiddenFromGAL"
						}
						catch
						{
							Write-Verbose "$($mem.Name) errored while marking as hidden"
							$DDLHidden = "Failed:$($_.Exception.Message)"
						}
						
						try
						{
							Write-Verbose "tryin to remove $($mem.name) from $DLName"
							$paramRemoveDistributionGroupMember = @{
								Identity = $DLName
								Member   = $DDSInfo.PrimarySmtpAddress.ToString()
								Confirm  = $false
								BypassSecurityGroupManagerCheck = $true
								DomainController = $DomainController
								ErrorAction = 'Stop'
							}
							
							Remove-DistributionGroupMember @paramRemoveDistributionGroupMember
							Write-Verbose "removed $($DDSInfo.PrimarySmtpAddress.ToString()) from $DL"
							$removeDL = "Removed:$($DDSInfo.PrimarySmtpAddress.ToString()) from:$DLName"
						}
						catch
						{
							Write-Verbose "Error While removing $($mem.name) from $DL"
							$removeDL = "Failed:$($_.Exception.Message)"
						}
					}
					else
					{
						#error Creating mailbox
						$fwdAdd = "Error"
						$DLMem = "Error"
						$DDLHidden = "Error"
						$removeDL = "Error"
					}
					Write-Verbose "Trying to generate properties for the output PsObject"
					$prop = [Ordered]@{
						DLName	    = $DLName
						Member	    = $SMName
						SMBXState   = $mbxCreation
						ForwardAdd  = $fwdAdd
						AddDLGrp    = $DLMem
						SetHidden   = $DDLHidden
						RemoveDLMem = $removeDL
						Details	    = 'None'
					}
					
					$obj = New-Object -TypeName System.Management.Automation.PSObject -Property $prop
					Write-Output $obj
				}
			}
			else
			{
				Write-Verbose "No members found for $($group.Name)"
				#No members to process
				
				$fwdAdd = "NothingToProcess"
				$DLMem = "NothingToProcess"
				$DDLHidden = "NothingToProcess"
				$removeDL = "NothingToProcess"
				
				$prop = [Ordered]@{
					DLName	    = $DLName
					Member	    = $SMName
					SMBXState   = $mbxCreation
					ForwardAdd  = $fwdAdd
					AddDLGrp    = $DLMem
					SetHidden   = $DDLHidden
					RemoveDLMem = $removeDL
					Details	    = 'None'
				}
				
				$obj = New-Object -TypeName System.Management.Automation.PSObject -Property $prop
				Write-Output $obj
			}
			
			Write-Verbose "generated properties for the Properties"
		}
		catch
		{
			Write-Verbose "errors encountered in the main try block"
			$prop = [Ordered]@{
				DLName	    = $DLName
				Member	    = 'Error'
				SMBXState   = 'Error'
				ForwardAdd  = 'Error'
				AddDLGrp    = 'Error'
				SetHidden   = 'Error'
				RemoveDLMem = 'Error'
				Details	    = "ERROR : $($_.Exception.Message)"
			}
			Write-Verbose "trying to create Psobject"
			$obj = New-Object -TypeName System.Management.Automation.PSObject -Property $prop
			Write-Output $obj
			Write-Verbose "Psobject successfully written to Pipeline"
			
		}
		finally
		{
			
			#$obj = New-Object -TypeName System.Management.Automation.PSObject -Property $prop
			#Write-Output $obj
		}
	}
	end
	{
		$Global:ErrorActionPreference = 'Continue'
	}
}

function Get-QhRoomCalendarSettings
{
<#
	.SYNOPSIS
		A brief description of the Get-QhRoomCalendarSettings function.
	
	.DESCRIPTION
		A detailed description of the Get-QhRoomCalendarSettings function.
	
	.PARAMETER UserPrincipalName
		A description of the UserPrincipalName parameter.
	
	.PARAMETER DownloadDirectory
		A description of the DownloadDirectory parameter.
	
	.PARAMETER showProgress
		A description of the showProgress parameter.
	
	.EXAMPLE
		PS C:\> Get-QhRoomCalendarSettings -UserPrincipalName 'user1@health.qld.gov.au'
	
	.NOTES
		Additional information about the function.
#>
	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory = $true)]
		[String[]]$UserPrincipalName,
		[String]$DownloadDirectory = $(Get-Location | Select-Object -ExpandProperty path),
		[Switch]$showProgress
	)
	
	begin
	{
		$global:ErrorActionPreference = [System.Management.Automation.ActionPreference]::Stop
		$FormatEnumerationLimit = -1
		$targetDir = Set-Qhdir -path "$DownloadDirectory\CalendarSettings"
		$date = get-Date
		$i = 1
	}
	process
	{
		foreach ($Upn in $UserPrincipalName)
		{
			$fileusr = $null
			$CalprocessingSetings = $null
			$AuditFile = $null
			$AuditFileXML = $null
			
			try
			{
				$recipient = Get-Recipient $Upn -ErrorAction Stop
				
				if ($recipient.RecipientTypeDetails -eq 'RoomMailbox')
				{
					$fileusr = $Upn
					$csv = "$DownloadDirectory\CalendarProcessingsetting.csv"
					
					$CalprocessingSetings = Get-CalendarProcessing -identity $Upn -ErrorAction Stop
					$CalprocessingSetings | Export-Csv $csv -NoTypeInformation -Append -ErrorAction SilentlyContinue
					
					[System.IO.Path]::GetInvalidFileNameChars() | ForEach-Object { $fileusr = $fileusr.replace($_, '_') }
					
					$AuditFile = "$targetDir\$($fileusr)_Cal_$($date.ToString('yyyyMMdd-HHmmss')).txt"
					$AuditFileXML = "$targetDir\$($fileusr)_Cal_$($date.ToString('yyyyMMdd-HHmmss')).xml"
					"$UPN Calendar Processing Settings : $date" >> $AuditFile
					$CalprocessingSetings | Format-List >> $AuditFile
					$CalprocessingSetings | Export-Clixml -Depth 9000 $AuditFileXML
					$status = 'SettingsExported'
				}
				else
				{
					$status = 'NotApplicable'
					$csv = 'NotApplicable'
					$AuditFile = 'NotApplicable'
					$AuditFileXML = 'NotApplicable'
				}
				
				$prop = [ordered]@{
					UserPrincipalName    = $Upn
					RecipientTypeDetails = $recipient.RecipientTypeDetails
					CSVPath			     = $csv
					SettingsFileTxt	     = $AuditFile
					SettingsFileXML	     = $AuditFileXML
					Status			     = $status
					Details			     = 'None'
				}
			}
			catch
			{
				$status = 'Error'
				$csv = 'Error'
				$AuditFile = 'Error'
				$AuditFileXML = 'Error'
				
				$prop = [ordered]@{
					UserPrincipalName    = $Upn
					RecipientTypeDetails = $recipient.RecipientTypeDetails
					CSVPath			     = $csv
					SettingsFileTxt	     = $AuditFile
					SettingsFileXML	     = $AuditFileXML
					Status			     = $status
					Details			     = "ERROR : $($_.Exception.Message)"
				}
			}
			finally
			{
				$obj = New-Object -TypeName System.Management.Automation.PSObject -Property $prop
				Write-Output $obj
				if ($ShowProgress)
				{
					if ($UserPrincipalName.Count -gt 1)
					{
						$paramWriteProgress = @{
							Activity = 'Exporting Calendar Processing Settings'
							Status   = "Processing [$i] of [$($UserPrincipalName.Count)] users"
							PercentComplete = (($i / $UserPrincipalName.Count) * 100)
							CurrentOperation = "Completed : [$UPN]"
						}
						
						Write-Progress @paramWriteProgress
					}
				}
				$i++
			}
		}
	}
	end
	{
		Write-Progress -Activity 'Exporting Calendar Processing Settings' -Completed
		$global:ErrorActionPreference = [System.Management.Automation.ActionPreference]::Continue
	}
}

function Compare-QhCalProcessing
{
	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory = $true,
				   Position = 0)]
		$UserPrincipalName,
		[Parameter(Mandatory = $true)]
		$OldSettings
		
	)
	
	begin
	{
		$CalProp = @(
			'AddAdditionalResponse'
			'AdditionalResponse'
			'AddNewRequestsTentatively'
			'AddOrganizerToSubject'
			'AllBookInPolicy'
			'AllowConflicts'
			'AllowRecurringMeetings'
			'AllRequestInPolicy'
			'AllRequestOutOfPolicy'
			'AutomateProcessing'
			'BookingWindowInDays'
			'BookInPolicy'
			'ConflictPercentageAllowed'
			'DeleteAttachments'
			'DeleteComments'
			'DeleteNonCalendarItems'
			'DeleteSubject'
			'EnableResponseDetails'
			'EnforceSchedulingHorizon'
			'ForwardRequestsToDelegates'
			'MaximumConflictInstances'
			'MaximumDurationInMinutes'
			'OrganizerInfo'
			'ProcessExternalMeetingMessages'
			'RemoveForwardedMeetingNotifications'
			'RemoveOldMeetingMessages'
			'RemovePrivateProperty'
			'RequestInPolicy'
			'RequestOutOfPolicy'
			'ResourceDelegates'
			'ScheduleOnlyDuringWorkHours'
			'TentativePendingApproval'
		)
		$toFix = @('BookInPolicy', 'RequestInPolicy', 'RequestOutOfPolicy')
	}
	process
	{
		try
		{
			$prop = $null
			$NewSettings = $null
			$er = $null
			$prop = [ordered]@{
				UserPrincipalName = $UserPrincipalName
			}
			$paramToFix = @{ }
			
			$NewSettings = Get-EXOCalendarProcessing -identity $UserPrincipalName -ErrorAction 'Stop'
			
			foreach ($conf in $CalProp)
			{
				$old = $null
				$new = $null
				$old = $($OldSettings.$conf)
				$new = $($NewSettings.$conf)
				
				if ($New -ne $Old)
				{
					if ($toFix -icontains $conf)
					{
						$xold = $null
						$xold = $old | ForEach-Object { Get-ExoRecipient $($_.Split('/')[-1]) | Select-Object -ExpandProperty PrimarySmtpAddress }
						$paramToFix.Add($conf, $xold)
					}
					$prop.Add($conf, @("OLD:$old", "NEW:$New"))
				}
				else
				{
					$prop.Add($conf, "OK")
				}
			}
			if ($paramToFix -ne $null)
			{
				try
				{
					Set-EXOCalendarProcessing -Identity $UserPrincipalName @paramToFix -ErrorAction Stop
					$fixed = 'Success'
				}
				catch
				{
					$fixed = 'Failed'
				}
			}
			else
			{
				$filxed = 'NotNeeded'
			}
			$prop.Add('Fixed', $fixed)
			$prop.Add('Details', "None")
		}
		catch
		{
			
			$er = $_.Exception.Message
			Write-Host "$er" -ForegroundColor Magenta
			foreach ($conf in $CalProp)
			{
				$prop.Add($conf, "Error")
			}
			$prop.Add('Fixed', 'Error')
			$prop.Add('Details', "Error : $Er")
		}
	}
	end
	{
		$obj = New-Object -TypeName System.Management.Automation.PSObject -Property $prop
		Write-Output $obj
	}
}

function Move-xxQHUsersToSFB-Parallel
{
	[CmdletBinding()]
	param
	(
		[String]$batchName,
		[String[]]$BulkUsers,
		[ValidateRange(1, 20)]
		[int]$ParellelSessions,
		[String]$credCsv
	)
	
	begin
	{
		try
		{
			$Creds = Import-Csv $credCsv -ErrorAction Stop | Select-Object -first $ParellelSessions
		}
		catch
		{
			Write-Host "Error : $($_.Exception.Message)" -ForegroundColor Magenta
			break
		}
	}
	process
	{
		$PreScript = {
			<#$Exservers = @(
				'EXC-CHMDC2P002'
				'EXC-CHMDC1P004'
				'exc-casbk7p003'
				'EXC-CHMNDCP001'
				'EXC-CHMDC2P001'
				'EXC-JRNDC1P001'
				'exc-casbk7p004'
				'exc-casbk7p005'
				'exc-casbk7p002'
				'EXC-CHMDC1P012'
				'EXC-CHMDC1P010'
				'EXC-JRNDC2P001'
				'EXC-CHMDC2P007'
				'exc-casbtpp006'
				'EXC-CHMDC2P006'
				'EXC-CHMDC2P009'
				'EXC-CHMDC2P004'
				'EXC-CHMDC2P010'
				'EXC-JRNDC2P002'
				'EXC-CHMDC2P003'
				'EXC-CHMDC1P002'
				'EXC-CHMDC2P011'
				'EXC-CHMDC2P008'
				'EXC-CHMDC2P005'
				'EXC-JRNDC1P002'
				'EXC-CHMDC2P012'
				'EXC-CHMDC1P008'
				'EXC-CHMDC1P003'
				'EXC-CHMDC1P005'
				'EXC-CHMDC1P009'
			)#>
			$ExServers = ${D:\Office365\Migrations\Batch\WorkingExchangeServers.txt}
			Connect-QHOnpremExchange -Server ($Exservers | Get-Random)
			Import-Module QHO365MigrationOps -WarningAction SilentlyContinue # module which contains the functions.
			Import-Module SkypeForBusiness
		}
		
		$ScriptBlock = {
			param (
				[String[]]$users,
				[String]$batchName,
				[pscredential]$cred
			)
			
			
			
			Connect-QhO365 -Credential $cred
			#Connect-QhSkypeOnline -Credential $cred
			$Exservers = @(
				'EXC-CHMDC2P002'
				'EXC-CHMDC1P004'
				'exc-casbk7p003'
				'EXC-CHMNDCP001'
				'EXC-CHMDC2P001'
				'EXC-JRNDC1P001'
				'exc-casbk7p004'
				'exc-casbk7p005'
				'exc-casbk7p002'
				'EXC-CHMDC1P012'
				'EXC-CHMDC1P010'
				'EXC-JRNDC2P001'
				'EXC-CHMDC2P007'
				'exc-casbtpp006'
				'EXC-CHMDC2P006'
				'EXC-CHMDC2P009'
				'EXC-CHMDC2P004'
				'EXC-CHMDC2P010'
				'EXC-JRNDC2P002'
				'EXC-CHMDC2P003'
				'EXC-CHMDC1P002'
				'EXC-CHMDC2P011'
				'EXC-CHMDC2P008'
				'EXC-CHMDC2P005'
				'EXC-JRNDC1P002'
				'EXC-CHMDC2P012'
				'EXC-CHMDC1P008'
				'EXC-CHMDC1P003'
				'EXC-CHMDC1P005'
				'EXC-CHMDC1P009'
			)
			Connect-QHOnpremExchange -Server ($Exservers | Get-Random)
			Import-Module QHO365MigrationOps -WarningAction SilentlyContinue # module which contains the functions.
			Import-Module SkypeForBusiness
			New-PSDrive –Name "E" –PSProvider FileSystem –Root '\\exc-mgtbk7p001\D$' –Persist -ErrorAction SilentlyContinue >> $null
			
			Move-zQHSkypeUserToOkypeOnline -UserPrincipalName $users -BatchName $batchName -Credential $cred
			
		}
		if ($ParellelSessions -ne $null -and $BulkUsers.Count -gt 2)
		{
			$dataSet = Split-Array $BulkUsers -parts $ParellelSessions
			$Sub = 0
			foreach ($set in $dataSet)
			{
				#$user = $creds[$Sub].UPN
				$user = $creds[$Sub].UserPrincipalName
				$onpremUser = $Creds[$Sub].UserName
				$AppPassword = $creds[$Sub].AppPassword
				$password = $creds[$Sub].Password
				
				#$cred = Get-QHCreds -UserName $user -Password $AppPassword
				#$onpremcred = Get-QHCreds -UserName $user -Password $password
				
				$cred = Get-QHCreds -UserName $user -Password $AppPassword
				$loginCred = Get-QHCreds -UserName $user -Password $password
				
				
				$users = $set
				#Start-Job -Name "$($batchName)_MoveSkypeUsers_Sub$($Sub)" -InitializationScript $PreScript -ScriptBlock $scriptBlock -ArgumentList $users, $batchName, $cred
				Invoke-Command -Credential $LoginCred -Authentication Credssp -ComputerName 'EXC-HBDBK7P001' -AsJob -JobName "$($batchName)_MoveSkypeUsers_Sub$($Sub)" -ScriptBlock $ScriptBlock -ArgumentList $Users, $batchName, $cred
				
				
				#Invoke-Command -Credential $onpremcred -ScriptBlock $ScriptBlock -AsJob -ArgumentList $users, $batchName, $cred
				$sub++
			}
			#$completed = $null
			$stopwatch = [system.diagnostics.stopwatch]::StartNew()
			while (@(Get-Job -Name "$($batchName)*" | Where-Object {
						$_.State -eq "Running"
					}).Count -ne 0)
			{
				#Clear-Host
				Write-Host "Please Wait While Jobs Complete : Completed - $((Get-job | Receive-job -keep).count) ElapsedTime: $($stopwatch.Elapsed.Hours):$($stopwatch.Elapsed.Minutes):$($stopwatch.Elapsed.Seconds)" -ForegroundColor Yellow
				$jobStatus = Get-job | Out-String
				Write-Host $jobStatus -ForegroundColor Cyan
				Start-Sleep -Seconds 5
			}
			Start-Sleep -Seconds 3
			Write-Host "All Jobs Completed : Completed - $((Get-job | Receive-job -keep).count) ElapsedTime: $($stopwatch.Elapsed.Hours):$($stopwatch.Elapsed.Minutes):$($stopwatch.Elapsed.Seconds)" -ForegroundColor Yellow
			$jobStatus = Get-job | Out-String
			Write-Host $jobStatus -ForegroundColor Green
			$stopwatch.Stop()
			$data = Get-job | Receive-Job -Keep
			
			Navigate-QHMigrationFolder $batchName
			
			#Write-Output $data | Export-Csv "$($batchName)_MovetoSkypeOnline.csv" -NoTypeInformation
			
		}
		
	}
	end
	{
		
	}
}

function Get-IPrangeStartEnd
{
    <#  
      .SYNOPSIS   
        Get the IP addresses in a range  
      .EXAMPLE  
       Get-IPrangeStartEnd -start 192.168.8.2 -end 192.168.8.20  
      .EXAMPLE  
       Get-IPrangeStartEnd -ip 192.168.8.2 -mask 255.255.255.0  
      .EXAMPLE  
       Get-IPrangeStartEnd -ip 192.168.8.3 -cidr 24  
    #>	
	
	param (
		[string]$start,
		[string]$end,
		[string]$ip,
		[string]$mask,
		[int]$cidr
	)
	
	function IP-toINT64 ()
	{
		param ($ip)
		
		$octets = $ip.split(".")
		return [int64]([int64]$octets[0] * 16777216 + [int64]$octets[1] * 65536 + [int64]$octets[2] * 256 + [int64]$octets[3])
	}
	
	function INT64-toIP()
	{
		param ([int64]$int)
		
		return (([math]::truncate($int/16777216)).tostring() + "." + ([math]::truncate(($int % 16777216)/65536)).tostring() + "." + ([math]::truncate(($int % 65536)/256)).tostring() + "." + ([math]::truncate($int % 256)).tostring())
	}
	
	if ($ip) { $ipaddr = [Net.IPAddress]::Parse($ip) }
	if ($cidr) { $maskaddr = [Net.IPAddress]::Parse((INT64-toIP -int ([convert]::ToInt64(("1" * $cidr + "0" * (32 - $cidr)), 2)))) }
	if ($mask) { $maskaddr = [Net.IPAddress]::Parse($mask) }
	if ($ip) { $networkaddr = new-object net.ipaddress ($maskaddr.address -band $ipaddr.address) }
	if ($ip) { $broadcastaddr = new-object net.ipaddress (([system.net.ipaddress]::parse("255.255.255.255").address -bxor $maskaddr.address -bor $networkaddr.address)) }
	
	if ($ip)
	{
		$startaddr = IP-toINT64 -ip $networkaddr.ipaddresstostring
		$endaddr = IP-toINT64 -ip $broadcastaddr.ipaddresstostring
	}
	else
	{
		$startaddr = IP-toINT64 -ip $start
		$endaddr = IP-toINT64 -ip $end
	}
	
	$temp = "" | Select-Object start, end
	$temp.start = INT64-toIP -int $startaddr
	$temp.end = INT64-toIP -int $endaddr
	return $temp
}

function Get-zHtmlTable
{
	param (
		[Parameter(Mandatory = $true)]
		$url,
		$tableIndex = 0,
		$Header,
		[int]$FirstDataRow = 0
	)
	
	$r = Invoke-WebRequest $url -Proxy "http://proxy.health.qld.gov.au:8080" -ProxyUseDefaultCredentials
	$table = $r.ParsedHtml.getElementsByTagName("table")[$tableIndex]
	$propertyNames = $Header
	$totalRows = @($table.rows).count
	
	for ($idx = $FirstDataRow; $idx -lt $totalRows; $idx++)
	{
		
		$row = $table.rows[$idx]
		$cells = @($row.cells)
		
		if (!$propertyNames)
		{
			if ($cells[0].tagName -eq 'th')
			{
				$propertyNames = @($cells | ForEach-Object { $_.innertext -replace ' ', '' })
			}
			else
			{
				$propertyNames = @(1 .. ($cells.Count + 2) | ForEach-Object { "P$_" })
			}
			continue
		}
		
		$result = [ordered]@{ }
		
		for ($counter = 0; $counter -lt $cells.Count; $counter++)
		{
			$propertyName = $propertyNames[$counter]
			
			if (!$propertyName) { $propertyName = '[missing]' }
			$result.$propertyName = $cells[$counter].InnerText
		}
		
		[PSCustomObject]$result
	}
}

function Get-PublicDomainCerts
{
	param (
		$domain
	)
	try
	{
		$url = "https://crt.sh/?q=%25$($domain)"
		$ErrorActionPreference = 'Stop'
		
		[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
		$data = Get-zHtmlTable -url $url -tableIndex 2
		return $data
		
	}
	catch
	{
		Write-Warning "$($_.exception.Message)"
		
	}
}

function Test-QHGenericGroupMemberShips
{
<#
	.SYNOPSIS
		A brief description of the Test-QHGenericGroupMemberShips function.
	
	.DESCRIPTION
		A detailed description of the Test-QHGenericGroupMemberShips function.
	
	.PARAMETER UserPrincipalName
		A description of the UserPrincipalName parameter.
	
	.PARAMETER ShowProgress
		A description of the ShowProgress parameter.
	
	.EXAMPLE
		PS C:\> Test-QHGenericGroupMemberShips
	
	.NOTES
		Additional information about the function.
#>
	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory = $true,
				   ValueFromPipeline = $true,
				   ValueFromPipelineByPropertyName = $true,
				   Position = 0)]
		[ValidateNotNullOrEmpty()]
		[String[]]$UserPrincipalName,
		[switch]$ShowProgress
	)
	begin
	{
		$i = 1
		$Global:ErrorActionPreference = [System.Management.Automation.ActionPreference]::Stop
		
		$genericGroups = @(
			'ACT-SaaS-O365-Non-Employee',
			'ACT-SaaS-O365-GenericAccountMailboxOnly',
			'ACT-SaaS-O365-GenericAccountFullLicence',
			'ACT-SaaS-O365-GenericAccountOfficeProPlusAndMailbox',
			'ACT-SaaS-O365-GenericAccountOfficeProPlusOnly'
		)
		
	}
	process
	{
		foreach ($upn in $UserPrincipalName)
		{
			$recipient = $null
			$ad = $null
			$genOpsdetails = $null
			$Allgroups = $null
			$recipient = $null
			$groups = @()
			
			try
			{
				$recipient = Get-Recipient $UPN -ErrorAction Stop
				$ad = get-AdUser $recipient.SamAccountName -properties * -ErrorAction Stop
				$Allgroups = $ad.MemberOf.ForEach({ $_.Split(',')[0].Split('=')[1] })
				
				
				foreach ($g in $Allgroups)
				{
					if ($genericGroups -icontains $g)
					{
						$groups += $g
					}
				}
				
				if ($groups -ne $null)
				{
					$generic = $true
					$GroupDetails = $($groups -join ',' | Out-String).Trim()
				}
				else
				{
					$generic = $false
					$GroupDetails = 'None'
				}
				
				$prop = [Ordered]@{
					UserPrincipalName    = $upn
					RecipientTypeDetails = $recipient.RecipientTypeDetails
					ADEnabled		     = $ad.Enabled
					IsGeneric		     = $generic
					GenericGroups	     = $GroupDetails
					Details			     = 'None'
				}
			}
			catch
			{
				$prop = [Ordered]@{
					UserPrincipalName    = $upn
					RecipientTypeDetails = 'Error'
					ADEnabled		     = 'Error'
					IsGeneric		     = 'Error'
					GenericGroups	     = 'Error'
					Details			     = "ERROR : $($_.Exception.Message)"
				}
			}
			finally
			{
				$obj = New-Object -TypeName System.Management.Automation.PSObject -Property $prop
				Write-Output $obj
				if ($ShowProgress)
				{
					if ($UserPrincipalName.Count -gt 1)
					{
						$paramWriteProgress = @{
							Activity = 'Testing Generic Group Membership'
							Status   = "Processing [$i] of [$($UserPrincipalName.Count)] users"
							PercentComplete = (($i / $UserPrincipalName.Count) * 100)
							CurrentOperation = "Completed : [$UPN]"
						}
						Write-Progress @paramWriteProgress
					}
				}
				$i++
			}
		}
	}
	end
	{
		Write-Progress -Activity 'Testing Generic Group Membership' -Completed
	}
	#TODO: Place script here
}

function Get-QHForwarders
{
<#
	.SYNOPSIS
		A brief description of the Set-QHFunctionTemplate function.
	
	.DESCRIPTION
		A detailed description of the Set-QHFunctionTemplate function.
	
	.PARAMETER UserPrincipalName
		A description of the UserPrincipalName parameter.
	
	.PARAMETER ShowProgress
		A description of the ShowProgress parameter.
	
	.EXAMPLE
		PS C:\> Set-QHFunctionTemplate -UserPrincipalName 'value1'
	
	.NOTES
		Additional information about the function.
#>
	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory = $true,
				   ValueFromPipeline = $true,
				   ValueFromPipelineByPropertyName = $true,
				   Position = 0)]
		[ValidateNotNullOrEmpty()]
		[String[]]$UserPrincipalName,
		[switch]$ShowProgress
	)
	
	begin
	{
		$i = 1
		$Global:ErrorActionPreference = [System.Management.Automation.ActionPreference]::Stop
	}
	process
	{
		foreach ($UPN in $UserPrincipalName)
		{
			try
			{
				$mbx = Get-Mailbox $UPN -ErrorAction 'Stop'
				if ($mbx.RecipientTypeDetails -match 'mailbox')
				{
					if ($null -ne $mbx.ForwardingAddress)
					{
						$hasFwd = $true
						try
						{
							$fd = get-Recipient $mbx.ForwardingAddress -ErrorAction 'Stop'
							$FwdAdd = $fd.PrimarySmtpAddress.ToString()
							$fwdRec = $fd.RecipientTypeDetails
						}
						catch
						{
							$FwdAdd = "Error : $($mbx.ForwardingAddress) could not be retrived"
							$fwdRec = "Error : $($mbx.ForwardingAddress) could not be retrived"
						}
					}
					else
					{
						$hasFwd = $false
						$FwdAdd = 'None'
						$fwdRec = 'None'
						
					}
				}
				else
				{
					$hasFwd = 'NotApplicable'
					$FwdAdd = 'NotApplicable'
					$fwdRec = 'NotApplicable'
				}
				
				$prop = [Ordered] @{
					UserPrincipalName    = $UPN
					RecipientTypeDetails = $mbx.RecipientTypeDetails
					HasForwarder		 = $hasFwd
					ForwardedTo		     = $FwdAdd
					ForwarderType	     = $fwdRec
					Details			     = 'None'
				}
			}
			catch
			{
				$prop = [Ordered] @{
					UserPrincipalName    = $UPN
					RecipientTypeDetails = 'Error'
					HasForwarder		 = 'Error'
					ForwardedTo		     = 'Error'
					ForwarderType	     = 'Error'
					Details			     = "ERROR : $($_.Exception.Message)"
				}
			}
			finally
			{
				$obj = New-Object -TypeName System.Management.Automation.PSObject -Property $prop
				Write-Output $obj
				
				if ($UserPrincipalName.count -gt 1)
				{
					if ($ShowProgress)
					{
						$paramWriteProgress = @{
							Activity = 'Retriving Forwarders'
							Status   = "Processing [$i] of [$($UserPrincipalName.Count)] users"
							PercentComplete = (($i / $UserPrincipalName.Count) * 100)
							CurrentOperation = "Completed : [$UPN]"
						}
						Write-Progress @paramWriteProgress
					}
				}
				$i++
			}
		}
	}
	end
	{
		Write-Progress -Activity 'Retriving Forwarders' -Completed
		$Global:ErrorActionPreference = [System.Management.Automation.ActionPreference]::Continue
	}
}

function Get-QHForwardersO365
{
<#
	.SYNOPSIS
		A brief description of the Set-QHFunctionTemplate function.
	
	.DESCRIPTION
		A detailed description of the Set-QHFunctionTemplate function.
	
	.PARAMETER UserPrincipalName
		A description of the UserPrincipalName parameter.
	
	.PARAMETER ShowProgress
		A description of the ShowProgress parameter.
	
	.EXAMPLE
		PS C:\> Set-QHFunctionTemplate -UserPrincipalName 'value1'
	
	.NOTES
		Additional information about the function.
#>
	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory = $true,
				   ValueFromPipeline = $true,
				   ValueFromPipelineByPropertyName = $true,
				   Position = 0)]
		[ValidateNotNullOrEmpty()]
		[String[]]$UserPrincipalName,
		[switch]$ShowProgress
	)
	
	begin
	{
		$i = 1
		$Global:ErrorActionPreference = [System.Management.Automation.ActionPreference]::Stop
	}
	process
	{
		foreach ($UPN in $UserPrincipalName)
		{
			try
			{
				$mbx = Get-ExoMailbox $UPN -ErrorAction 'Stop'
				if ($mbx.RecipientTypeDetails -match 'mailbox')
				{
					if ($null -ne $mbx.ForwardingAddress)
					{
						$hasFwd = $true
						
						try
						{
							$fd = get-ExoRecipient $mbx.ForwardingAddress -ErrorAction 'Stop'
							$FwdAdd = $fd.PrimarySmtpAddress.ToString()
							$fwdRec = $fd.RecipientTypeDetails
						}
						catch
						{
							$FwdAdd = "Error : $($mbx.ForwardingAddress) could not be retrived"
							$fwdRec = "Error : $($mbx.ForwardingAddress) could not be retrived"
						}
					}
					else
					{
						$hasFwd = $false
						$FwdAdd = 'None'
						$fwdRec = 'None'
					}
				}
				else
				{
					$hasFwd = 'NotApplicable'
					$FwdAdd = 'NotApplicable'
					$fwdRec = 'NotApplicable'
				}
				
				$prop = [Ordered] @{
					UserPrincipalName    = $UPN
					RecipientTypeDetails = $mbx.RecipientTypeDetails
					HasForwarder		 = $hasFwd
					ForwardedTo		     = $FwdAdd
					ForwarderType	     = $fwdRec
					Details			     = 'None'
				}
			}
			catch
			{
				$prop = [Ordered] @{
					UserPrincipalName    = $UPN
					RecipientTypeDetails = 'Error'
					HasForwarder		 = 'Error'
					ForwardedTo		     = 'Error'
					ForwarderType	     = 'Error'
					Details			     = "ERROR : $($_.Exception.Message)"
				}
			}
			finally
			{
				$obj = New-Object -TypeName System.Management.Automation.PSObject -Property $prop
				Write-Output $obj
				
				if ($UserPrincipalName.count -gt 1)
				{
					if ($ShowProgress)
					{
						$paramWriteProgress = @{
							Activity = 'Retriving Forwarders'
							Status   = "Processing [$i] of [$($UserPrincipalName.Count)] users"
							PercentComplete = (($i / $UserPrincipalName.Count) * 100)
							CurrentOperation = "Completed : [$UPN]"
						}
						Write-Progress @paramWriteProgress
					}
				}
				$i++
			}
		}
	}
	end
	{
		Write-Progress -Activity 'Retriving Forwarders' -Completed
		$Global:ErrorActionPreference = [System.Management.Automation.ActionPreference]::Continue
	}
}

function Get-ValidatedUsers
{
	param (
		[String]$BatchName
		
	)
	
	if (Navigate-QHMigrationFolder -BatchName $BatchName)
	{
		try
		{
			[array]$users = Import-Csv ".\$($BatchName)_Validation.csv" |
			Where-Object lookup -eq passed |
			Select-Object -ExpandProperty EmailAddress
			
			Write-Host "Retrived : $($users.Count) Users" -ForegroundColor Green
			return $users
		}
		catch
		{
			Write-Host "Error : $($_.Exception.Message)" -ForegroundColor Magenta
			Write-Host "Retrived : $($users.Count) Users" -ForegroundColor magenta
			return $null
		}
	}
	else
	{
		Write-Host "Retrived : $($users.Count) Users" -ForegroundColor magenta
		return $null
	}
	
}

function Get-QHO365MbxAudit
{
	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory = $true,
				   Position = 0)]
		[String[]]$UserPrincipalName,
		[Parameter(Mandatory = $false,
				   Position = 1)]
		[ValidateRange(1, 30)]
		[ValidateNotNullOrEmpty()]
		[int]$DaysOld,
		[Parameter(Position = 2)]
		[switch]$AllIps
	)
	
	begin
	{
		$i = 1
		$start = ((Get-date).AddDays(- $($DaysOld))).ToString('MM/dd/yyyy')
		$Global:ErrorActionPreference = [System.Management.Automation.ActionPreference]::Stop
	}
	process
	{
		foreach ($Upn in $UserPrincipalName)
		{
			$obj = $null
			$RawAuditLogs = $null
			$xRawAuditLogs = $null
			$errmsg = $null
			
			try
			{
				if ($DaysOld -eq 0)
				{
					$paramSearchEXOMailboxAuditLog = @{
						Identity    = $Upn
						ShowDetails = $true
						LogonTypes  = 'Owner'
						#StartDate   = $start
						ErrorAction = 'Stop'
					}
				}
				else
				{
					$paramSearchEXOMailboxAuditLog = @{
						Identity    = $Upn
						ShowDetails = $true
						LogonTypes  = 'Owner'
						StartDate   = $start
						ErrorAction = 'Stop'
					}
				}
				
				
				$xRawAuditLogs = Search-EXOMailboxAuditLog @paramSearchEXOMailboxAuditLog |
				Where-Object { $_.ClientInfoString -Like "*MSExchangeRPC" }
				
				if ($AllIps)
				{
					$RawAuditLogs = $xRawAuditLogs | Select-Object *
				}
				else
				{
					$RawAuditLogs = $xRawAuditLogs | Where-Object { $_.ClientipAddress -notlike "*165.86*" } | Select-Object *
				}
				
				if ($RawAuditLogs -ne $null)
				{
					$obj = $RawAuditLogs | Sort-Object -Descending LastAccessed |
					Select-Object @{ n = 'UserPrincipalName'; e = { $_.MailboxOwnerUpn } },
								  LastAccessed,
								  ClientipAddress,
								  ClientProcessName,
								  ClientVersion,
								  @{ n = 'DaysOld'; e = { "$DaysOld" } },
								  @{ n = 'Details'; e = { 'None' } }
				}
				else
				{
					$obj = [PSCustomObject][ordered]@{
						UserPrincipalName = $Upn
						LastAccessed	  = 'NotAvailable'
						ClientipAddress   = 'NotAvailable'
						ClientProcessName = 'NotAvailable'
						ClientVersion	  = 'NotAvailable'
						DaysOld		      = 'NotAvailable'
						Details		      = "WARNING : No Audit Data Available"
					}
				}
				
			}
			catch
			{
				$errmsg = $_.Exception.Message
				
				if ($errmsg -like '"An error occurred while trying to access the audit log*')
				{
					Get-Qh
				}
				else
				{
					$obj = [PSCustomObject][ordered]@{
						UserPrincipalName = $Upn
						LastAccessed	  = 'Error'
						ClientipAddress   = 'Error'
						ClientProcessName = 'Error'
						ClientVersion	  = 'Error'
						DaysOld		      = 'Error'
						Details		      = "ERROR : $errmsg"
					}
				}
				
			}
			finally
			{
				Write-Output $obj
				$i++
			}
		}
	}
	end
	{
		$Global:ErrorActionPreference = [System.Management.Automation.ActionPreference]::Stop
	}
}

function Get-QHO365MbxAudit-Parallel
{
<#
	.SYNOPSIS
		A brief description of the Start-QHPostMigrationTasks-Parallel function.
	
	.DESCRIPTION
		A detailed description of the Start-QHPostMigrationTasks-Parallel function.
	
	.PARAMETER batchName
		A description of the batchName parameter.
	
	.PARAMETER BulkUsers
		A description of the BulkUsers parameter.
	
	.PARAMETER ParellelSessions
		A description of the ParellelSessions parameter.
	
	.PARAMETER credCsv
		A description of the credCsv parameter.
	
	.EXAMPLE
		PS C:\> Start-QHPostMigrationTasks-Parallel
	
	.NOTES
		Additional information about the function.
#>
	[CmdletBinding()]
	param
	(
		[String[]]$BulkUsers,
		[ValidateRange(1, 50)]
		[int]$ParellelSessions,
		[String]$credCsv
	)
	
	begin
	{
		try
		{
			$Creds = Import-Csv $credCsv -ErrorAction Stop | Select-Object -first $ParellelSessions
		}
		catch
		{
			Write-Host "Error : $($_.Exception.Message)" -ForegroundColor Magenta
			break
		}
	}
	process
	{
		$PreScript = {
			Import-Module QHO365MigrationOps -WarningAction SilentlyContinue # module which contains the functions.
		}
		
		$ScriptBlock = {
			param (
				[String[]]$users,
				[pscredential]$cred
			)
			
			Connect-QhO365 -Credential $cred
			
			Get-QHO365MbxAudit -UserPrincipalName $users -DaysOld 30
			
			Get-PSSession | Remove-PSSession
		}
		if ($ParellelSessions -ne $null -and $BulkUsers.Count -gt 3)
		{
			$dataSet = Split-Array $BulkUsers -parts $ParellelSessions
			$Sub = 0
			$jobs = foreach ($set in $dataSet)
			{
				$user = $creds[$Sub].UserPrincipalName
				$password = $creds[$Sub].AppPassword
				
				$cred = Get-QHCreds -UserName $user -Password $password
				$users = $set
				Start-Job -Name "MBXAudit$($Sub)" -InitializationScript $PreScript -ScriptBlock $scriptBlock -ArgumentList $users, $cred
				$sub++
			}
			#$completed = $null
			while ($jobs.State -eq 'Running')
			{
				Get-Job | Receive-Job
				Start-Sleep -Seconds 2
			}
			Start-Sleep -Seconds 1
			Get-Job | Receive-Job
		}
	}
	end
	{
		#Get-job | Remove-job
	}
}

function Get-QHAcompliHostStatus
{
<#
	.SYNOPSIS
		A brief description of the Test-QHAcompliEndPoints function.
	
	.DESCRIPTION
		A detailed description of the Test-QHAcompliEndPoints function.
	
	.PARAMETER AcompliHostsCSV
		A description of the AcompliHostsCSV parameter.
	
	.EXAMPLE
		PS C:\> Test-QHAcompliEndPoints -AcompliHostsCSV 'value1'
	
	.NOTES
		Additional information about the function.
#>
	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory = $true)]
		[String]$AcompliHostsCSV
	)
	
	begin
	{
		try
		{
			$OldHosts = Import-CSV $AcompliHostsCSV -ErrorAction Stop
			
			if ($OldHosts -eq $null)
			{
				throw "$AcompliHostsCSV does not have any previous data"
			}
		}
		catch
		{
			Write-Host "ERROR : $($_.Exception.Message)" -ForegroundColor Magenta
			break
		}
	}
	process
	{
		foreach ($Computer in $OldHosts)
		{
			$conn = $null
			$oldIP = $null
			$newip = $null
			$detail = $null
			$connection = $null
			$state = $null
			$obj = $null
			
			$oldIP = $Computer.IPAddress
			
			try
			{
				$conn = Test-NetConnection -ComputerName $($Computer.HostName) -Port 443 -ErrorAction Stop -WarningAction SilentlyContinue
				
				if ($conn.TcpTestSucceeded)
				{
					$connection = 'Suceeded'
					$newip = $conn.RemoteAddress.ToString()
					
					if ($oldIP -ne $newip)
					{
						$state = 'Changed'
						$detail = 'WARNING : IPs Mismatch'
					}
					else
					{
						$state = 'NotChanged'
						$detail = 'None'
					}
				}
				else
				{
					$state = 'Failed'
					$connection = 'Failed'
					$newip = 'None'
					$oldIP = $Computer.IPAddress
					$detail = "FAILED : Could not connect to $($Computer.HostName)"
				}
				
				$prop = [Ordered]@{
					HostName = $Computer.HostName
					ConnectionTest = $connection
					OldIpAddress = $oldIP
					NewIPAddress = $newip
					IPAddressState = $state
					Details  = $detail
				}
			}
			catch
			{
				$NewIp = 'Error'
				$detail = "ERROR : $($_.Exception.Message)"
				$prop = [Ordered]@{
					HostName = $Computer.HostName
					ConnectionTest = $connection
					OldIpAddress = $oldIP
					NewIPAddress = $newip
					IPAddressState = $state
					Details  = $detail
				}
			}
			finally
			{
				$obj = New-Object -TypeName System.Management.Automation.PSObject -Property $prop
				Write-Output $obj
			}
		}
	}
	end
	{
		
	}
}

function Get-QHAcompliHostStatusReport
{
<#
	.SYNOPSIS
		A brief description of the Get-QHAcompliHostStatusReport function.
	
	.DESCRIPTION
		A detailed description of the Get-QHAcompliHostStatusReport function.
	
	.PARAMETER InObject
		A description of the InObject parameter.
	
	.PARAMETER To
		A description of the To parameter.
	
	.PARAMETER From
		A description of the From parameter.
	
	.PARAMETER SmtpServer
		A description of the SmtpServer parameter.
	
	.PARAMETER HTMLReport
		A description of the HTMLReport parameter.
	
	.EXAMPLE
		PS C:\> Get-QHAcompliHostStatusReport -InObject $InObject -To 'value2' -From 'value3' -SmtpServer 'value4'
	
	.NOTES
		Additional information about the function.
#>
	
	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory = $true,
				   Position = 0)]
		[psobject]$InObject,
		[Parameter(Mandatory = $true,
				   Position = 1)]
		[String[]]$To,
		[Parameter(Mandatory = $false,
				   Position = 2)]
		[String]$From = "Acompli.Tests@health.qld.gov.au",
		[Parameter(Mandatory = $false,
				   Position = 3)]
		[String]$SmtpServer = "qhsmtp.health.qld.gov.au",
		[Parameter(Position = 4)]
		[Switch]$ShowHTMLReport
	)
	
	begin
	{
		$Style = @"
	    @charset "UTF-8";
	    body {
	    color:Black;
	    font-family: arial, sans-serif;
	    font-size: 11pt;
	    }
	    h1 {
	    text-align:center;
	    font-family:"Comic Sans MS", cursive, sans-serif;
	    color: Gray
	    }
	    h2 {
	    border-top:1px solid black;
	    font-family:"Comic Sans MS", cursive, sans-serif;
	    color: #435099
	    }

	    h3 {
		    font-family: arial, sans-serif;
		    color: Black;
	    }

	    h4 {
	    border-top:1px solid black;
	    font-family: arial, sans-serif;
	    color: Gray
	    }

	    h5 {
	    font-family:"Comic Sans MS", cursive, sans-serif;
	    color: Gray
	    }

	    table {
	      font-family: arial, sans-serif;
	      border-collapse: collapse;
	      margin-left: 35px;
	    }

	    th {
	        font-weight:Bold;
	        color:#ffffff;
	        background-color:#474545;
	        font-family: "Comic Sans MS", cursive, sans-serif;
	        font-size: 14px;
	        letter-spacing: 1.5px;
	        word-spacing: 0px;
	        font-weight: normal;
	        border: 1px solid #AAAAAA;
	        padding: 5px 10px;
	    }

	    TD{
		    Font color:Black;
		    font-family:Verdana, Geneva, sans-serif;
		    font-size: 11px;
		    border: 1px solid #AAAAAA;
		    border-style: solid;
	         padding: 3px 6px;
	        Text-Align: Center;
	    }

	    TR:Nth-Child(Even) {Background-Color: #dddddd;}
	    TR:Hover TD {Background-Color: #b8efff;}

	    .odd { background-color:#ffffff; }
	    .even { background-color:#dddddd; }


	    .paginate_disabled_previous, .paginate_disabled_next {
	    margin:4px;
	    color:#666666;
	    cursor:Default;
	    background-color:#dddddd;
	    font-family:Calibri,Tahoma;
	    padding:2px;
	    border-radius:2px;
	    }

	    .paginate_enabled_next, .paginate_enabled_previous {
	    cursor:pointer;
	    border:1px solid #222222;
	    background-color:#b8efff;
	    font-family:Calibri,Tahoma;
	    padding:2px;
	    margin:4px;
	    border-radius:2px;
	    }

	    .dataTables_info { margin-bottom:4px; }
	    .sectionheader { cursor:pointer; }
	    .sectionheader:hover { color:Blue; }


	    .Red {
	    color: #ffffff;
	    background-color:#FD625E;

	    }

	    .Green {
	    color:#ffffff;
	    background-color:#00b359;

	    }

	    .Yellow{
	    color:Black;
	    background-color:#F2C80F;
	    }

	    .Blue{
	    color:#ffffff;
	    background-color:#28738A;

	    }

	    .Gray {
	    color:#999696;

	    font-style:italic;
	    }
	    .Default {
	    color:Black;
	    font-family:Calibri,Tahoma;
	    }
"@
		$msg = @"
    <P>Hello Team,</P>
    <P>
    The below tests will confirm if the hostnames/IPs and ports for the Outlook Mobile App are open and functioning.<br>
    Connections Failed should be addressed promptly with a ticket logged to networks for the new IP Address to be added. There is a standard change for this.<br>
    Warnings can be ignored if the connection succeeds port is still open and has probably just failed over to another one of the already allowed IP Addressed.<br>
    </P>
    <br>
"@
	}
	process
	{
		$MailFrag = $InObject |
		ConvertTo-EnhancedHTMLFragment -As Table -TableCssClass Grid -PreContent "<h3>Details:</h3>" -Properties Hostname,
									   @{
			n   = 'Connection Test'; e = { $_.ConnectionTest };
			css = {
				if ($_.Connectiontest -eq 'Suceeded') { 'Green' }
				else { 'Red' }
			}
		},
									   @{
			n   = 'Old IPAddress'; e = { $_.OldIpAddress };
			css = { if ($_.OldIpAddress -eq $_.NewIPAddress) { 'Gray' } }
		},
									   @{
			n   = 'New IPAddress'; e = { $_.NewIPAddress };
			css = { if ($_.OldIpAddress -eq $_.NewIPAddress) { 'Gray' } }
		},
									   @{
			n   = 'Ip Address State'; e = { $_.IpAddressState };
			css = {
				if ($_.IpaddressState -eq 'Changed') { 'Yellow' }
				elseif ($_.IpaddressState -eq 'Failed') { 'Red' }
				else { 'Gray' }
			}
		},
									   @{
			n   = 'Details'; e = { $_.Details };
			css = { if ($_.Details -eq 'None') { 'Gray' } }
		}
		
		$htmlbody = ConvertTo-EnhancedHTML -HTMLFragments $MailFrag -CssStyleSheet $style -PreContent $msg | Out-String
		
		$paramSendMailMessage = @{
			To		   = $To
			From	   = $From
			Subject    = "Outlook Mobile App Daily Connection Test"
			Body	   = $htmlbody
			BodyAsHtml = $true
			SmtpServer = $SmtpServer
		}
		
		Send-MailMessage @paramSendMailMessage
		
		if ($ShowHTMLReport)
		{
			$HtmlFrag = $InObject |
			ConvertTo-EnhancedHTMLFragment -As Table -TableCssClass Grid -PreContent "<h3>Details:</h3>" -MakeTableDynamic -Properties Hostname,
										   @{
				n   = 'Connection Test'; e = { $_.ConnectionTest };
				css = {
					if ($_.Connectiontest -eq 'Suceeded') { 'Green' }
					else { 'Red' }
				}
			},
										   @{
				n   = 'Old IPAddress'; e = { $_.OldIpAddress };
				css = { if ($_.OldIpAddress -eq $_.NewIPAddress) { 'Gray' } }
			},
										   @{
				n   = 'New IPAddress'; e = { $_.NewIPAddress };
				css = { if ($_.OldIpAddress -eq $_.NewIPAddress) { 'Gray' } }
			},
										   @{
				n   = 'Ip Address State'; e = { $_.IpAddressState };
				css = {
					if ($_.IpaddressState -eq 'Changed') { 'Yellow' }
					elseif ($_.IpaddressState -eq 'Failed') { 'Red' }
					else { 'Gray' }
				}
			},
										   @{
				n   = 'Details'; e = { $_.Details };
				css = { if ($_.Details -eq 'None') { 'Gray' } }
			}
			
			$paramConvertToEnhancedHTML = @{
				HTMLFragments = $HtmlFrag
				CssStyleSheet = $style
				PreContent    = $msg
			}
			
			ConvertTo-EnhancedHTML @paramConvertToEnhancedHTML | Out-File .\AcompliHostReport.html
			Invoke-Item .\AcompliHostReport.html
		}
	}
	end
	{
		
	}
	#TODO: Place script here
}
