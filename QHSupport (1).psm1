<# 
    ********************************************************************************  
    *                                                                              *  
    *        This script Loads all my common f u n c t i o n s                     *
    *                                                                              *  
    ********************************************************************************    
    Note.
    All the standard Functions I use loaded.
      
    *******************
    Copyright Notice.
    *******************
    This Program is free software: you can redistribute it and/or modify
    it under the terms of the GNU General Public License as published by
    the Free Software Foundation, either version 3 of the License, or
    (at your option) any later version.

    This program is distributed in the hope that it will be useful,
    but WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
    GNU General Public License for more details.

    You should have received a copy of the GNU General Public License
    along with this program.  If not, see <http://www.gnu.org/licenses/>.

    test-exchangeonlineconnected

    Author:
    Lawrence McKay
    Lawrence@mckayit.com
    McKayIT Solutions Pty Ltd
    
    Date:   26  March 2018
   

    ******* Update Version number below when a change is done.*******

    History
    Version         Date                Name           Detail
    ---------------------------------------------------------------------------------------
    0.0.1           26 March 2018       Lawrence       Initial Coding
    0.0.2            5 April 2018       Lawrence       Added Check connection (test-msolserviceconnected, test-exchangeonlineconnected, test-exchangePremconnected, test-skypeonlineClixml)
    0.0.3           11 May   2018       Lawrence       Added Help function that displays the Functions list with the Descriptions it reads in that start with #... 
    0.0.4           11 May   2018       Lawrence       Added Check to make sure the script is run in ISE
    0.0.5           11 May   2018       Lawrence       Added Sharepoint Connections eg Connect-SpoService
    0.0.6           25 June  2018       Lawrence       Added Connect-all function to connect to all services
    0.0.7            4 July  2018       LAwrence       Added Logit function to logoutput to file. 
    0.0.8            5 July  2018       Lawrence       Fixed Connection / Check connection display issues.
    0.0.9           11 Sept  2018       Lawrence       Converted to a Module (QHSupport)
    0.0.12          14 Sept  2018       Lawrence       Added Function Sync-github
    0.0.13          14 Sept  2018       Lawrence       Added Function Remove-2ndMailboxFromOnLine2FixMigrationIssue
    0.0.14          17 Sept  2018       Lawrence       Added Function Publish-MailBoxMoveStats 
    0.0.15          17 Sept  2018       Lawrence       Added Function Compare-ImmutableID
    0.0.16          17 Sept  2018       Lawrence       Added Function Template
    0.0.17          24 Sept  2018       Lawrence       Fixed Function Get-FileComparesForInsentraAddRemovedusers

    0.0.19          15 Oct   2018       Lawrence       Added Prefix of QH to most functions.

    0.0.25          29 Oct   2018       Lawrence       Added Connect to MS Teams.
    0.0.26         3rd Dec   2018       Lawrence       Added Remove-usersFromValadationFile
    0.0.27         7rd Dec   2018       Lawrence       Added Get-QHLicenseFileErrors
    0.0.28         7rd Dec   2018       Lawrence       Added Get-QHMFAFileErrors
    0.0.29        10th Dec   2018       Lawrence       Added Get-FindWhenMoved
    0.0.30        11th Dec   2018       Lawrence       Added get-QHOnpremdatatoSync
    0.0.31        14th Dec   2018       Lawrence       Updated confirm-qhmigrationMatches


    #>

$version = "Version 0.0.31"
Write-Host $version -ForegroundColor Green


#$global:Functisloaded = "YES"

#Write-host "`nLoading $PSCommandPath"  -ForegroundColor green -BackgroundColor Red
$FormatEnumerationLimit = -1

#
#Displays Loaded Functions

Function get-help1 
{
<#
    #... Displays the Functions in QHSupport.
    Syntax   Get-help1

    This will generate the report for all of the batchname starting with Batch14*
#>

	param
	(
	
	)
   	
	BEGIN
    	{
         Write-Host $version -ForegroundColor Green
		}

      PROCESS
        {
            #... Display Functions with the comments
            $DDSPLAY = ""
            #Reads in the Current Powershell script file
            $DDSPLAY = get-content $PSCommandPath 

            foreach ($line in $DDSPLAY) {
                if ($line.Trim().StartsWith('Function', "CurrentCultureIgnoreCase") -or $line.Trim().Startswith('#...', "CurrentCultureIgnoreCase")) {
                    $1 = $line

                    if ($1.Trim().StartsWith('#...', "CurrentCultureIgnoreCase")) {
                
                        #Removes the first 4 char's Eg "#... "
                        $linedes = $line.trim().substring(4)
                        Write-Host $linedes -f Gray -NoNewline
                    }
                    Elseif (!($1.Trim().Startswith('#...'))) {
                        Write-host ''
                    }
                    if ($1.Trim().StartsWith('Function', "CurrentCultureIgnoreCase")) {
                        $linelong = $line + "                                               "
                
                        #makes the Line length to be 50 so the comments all line up. Fills it up with a space.
                        $line = $linelong.substring(0, 50)
                        Write-host "  $line" -f green -NoNewline
                    }
                }
            }
        
        }
   END
   {
    Write-Output  ""
   }
}


Function Set-qhonPremFullPermissions 
{
    #... Grants Full Mailbox Permissions to use on local Exchange
    test-exchangePremconnected
    [System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null
    $mailboxname = "" 
    $UserUPN = "" 
    $mailboxname = [Microsoft.VisualBasic.Interaction]::InputBox("Enter the name of the mailbox you want full controll of.", "Enter MB name", "")
    $UserUPN = [Microsoft.VisualBasic.Interaction]::InputBox("Enter the name of the user you want to add.`n E.G. lawrence.mckay@health.qld.gov.au", "Enter name", "")

    if ($mailboxname -eq "") {
        write-host   "You have not entered a name.." -ForegroundColor green
        break
    }
    
    if ($UserUPN -eq "") {
        write-host   "You have not entered a name.." -ForegroundColor green
        break
    }
   
    add-MailboxPermission -Identity Tester65O -User lawrence.mckay@health.qld.gov.au -AccessRights FullAccess -InheritanceType All -Automapping $false
}


function Remove-qhonPremFullPermissions 
{
    #... Removes Full Mailbox Permissions to use on local Exchange
    test-exchangePremconnected
    [System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null
    $mailboxname = "" 
    $UserUPN = "" 
    $mailboxname = [Microsoft.VisualBasic.Interaction]::InputBox("Enter the name of the mailbox you want full control of.", "Enter MB name", "")
    $UserUPN = [Microsoft.VisualBasic.Interaction]::InputBox("Enter the name of the user you want to Remove.`n E.G. lawrence.mckay@health.qld.gov.au", "Enter name", "")
  
    if ($mailboxname -eq "") {
        write-host   "You have not entered a name.." -ForegroundColor green
        break
    }
    
    if ($UserUPN -eq "") {
        write-host   "You have not entered a name.." -ForegroundColor green
        break
    }
  
    
    remove-MailboxPermission -Identity $mailboxname -User $UserUPN -AccessRights FullAccess -Confirm:$false
}


Function get-qhInputList 
{
    #... Reads in the CSV file used for most Functions (See inside for Comments)
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.Title = "Please select the file to use."
    $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.filter = "CSV (*.csv)| *.csv|All Files (*.*)| *.*"
    $OpenFileDialog.ShowDialog() | Out-Null
    $global:openfiledialogfilename = $OpenFileDialog.filename
  
    if ( $(Try { Test-Path $OpenFileDialog.filename } Catch { $false }) ) {
        write-host "Path OK"
        Write-host -ForegroundColor GREEN "Using the following file:"$OpenFileDialog.filename
  
        # Inports the CSV fiel that is used by most functions that need $inputlist  
        $global:inputlist = Import-Csv $OpenFileDialog.filename
 
        # Sets up the Batchname from the File name less Extension
        $global:batchname = [io.path]::GetFileNameWithoutExtension($OpenFileDialog.filename)
   
        # Setus up the Output File for where it is Proceessed and needs to write out
        $global:OutputFilepath = split-path $OpenFileDialog.filename -Leaf
        $global:OutputFile = $global:OutputFilepath + "\_" + $OpenFileDialog.SafeFileName
    }
    Else {
        write-host -ForegroundColor red -BackgroundColor yellow "$line70`n               You did not Select a File. Please try again..          `n$line70"
    }       
} 


function Connect-ExchangeOnline() 
{
    #... Connects to Exchange Online and Prompts for user name and Password.  
    $UserCredentials = Get-Credential x-$env:USERNAME@health.qld.gov.au -Message "Enter Exchange On-Line credentials"
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredentials -Authentication Basic -AllowRedirection 
    IF ($Session) {
        Import-PSSession $Session -prefix "exo" -AllowClobber
        #Import-PSSession $Session  -AllowClobber
        write-host "`n`n`n`n"
        show-line -color white
        Write-host "To disconnect this session, use Disconnect-ExchangeOnline cmdlet"
        show-line -color white
        write-host "Cmdlets are prefixed with 365"
        show-line -color white
    } 
    Else {
        Write-Error "Session failed to connect"
    }

} 


function Disconnect-ExchangeOnline 
{
    #... Disconnects from Exchange On Line eg removes PS Session
    Get-PSSession | where-object {$_.ComputerName -match "outlook.office"} | Remove-PSSession
}


function Disconnect-ExchangeOnPrem 
{
    #... Disconnects from Exchange OnPrem eg removes PS Session
    Remove-PSSnapin -name 'Microsoft.Exchange.Management.PowerShell.E2010'
    Get-PSSession | Where-Object {$_.ComputerName -match ".qh.health.qld.gov.au"} | Remove-PSSession
}
         
function Remove-2ndMailboxFromOnLine2FixMigrationIssue
{
    <# 
    ********************************************************************************  
    *                                                                              *  
    *              This script Removes 2nd Mailbox From OnLine                     *
    *                 2 Fix Migration Issue with Dup Mailbox                       *  
    *                                                                              *  
    ********************************************************************************    
    Note.
        Needs connection to AD and MSOL with no Prefix.
        
    SYNTAX.
        Remove-2ndMailboxFromOnLine2FixMigrationIssue -userUPN Fred.smith@health.qld.gov.au

      
    *******************
    Copyright Notice.
    *******************
    Copyright (c) 2018   McKayIT Solutions Pty Ltd.

    Permission is hereby granted, free of charge, to any person obtaining a copy
    of this software and associated documentation files (the "Software"), to deal
    in the Software without restriction, including without limitation the rights
    to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
    copies of the Software, and to permit persons to whom the Software is
    furnished to do so, subject to the following conditions:

    The above copyright notice and this permission notice shall be included in all
    copies or substantial portions of the Software.

    THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
    IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
    FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
    AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
    LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
    OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
    SOFTWARE.

    

    Author:
    Lawrence McKay
    Lawrence@mckayit.com
    McKayIT Solutions Pty Ltd
    
    Date:    14 Sept 2018
   

    ******* Update Version number below when a change is done.*******

    History
    Version         Date                Name           Detail
    ---------------------------------------------------------------------------------------
    0.0.1           14 Sept 2018       Lawrence       Initial Coding
    
   #>


#
#   Log into https://protection.office.com/?rfr=AdminCenter
#   and do a content Search and save results to PST and Download.
#   The PST is required to reinject the email back to the user.



    param
    (
        [Parameter(Mandatory = $true,
            HelpMessage = 'Enter UPN')]
        [ValidateNotNullOrEmpty()]
        $userUPN
    )
  
  Begin
  {
        show-line
        Write-host "   Log into https://protection.office.com/?rfr=AdminCenter"
        Write-host "   and do a content Search and save results to PST and Download."
        Write-host "   The PST is required to reinject the email back to the user."
        show-line
        Start-Sleep -Seconds 30
  }

  PROCESS
          {
 
  #  Change UPN so it does not sync to Azure
         $oldSuffix = "health.qld.gov.au"
         $newSuffix = "qh.health.qld.gov.au"
         Get-ADUser -Filter "UserPrincipalName -eq '$userUPN'" | ForEach-Object {
             $newUpn = $_.UserPrincipalName.Replace($oldSuffix,$newSuffix)
                $_ | Set-ADUser -UserPrincipalName $newUpn }



               #Remove Mailbox from O365 
                Remove-Msoluser -UserPrincipalName $userUPN -Force 
                Remove-Msoluser -UserPrincipalName $userUPN -RemoveFromRecycleBin

  #  Then after AD Sync has ran run the following
          wait-ToNextADSync


 #  Changes the UPN back to Sync Via AADSync
          $newSuffix = "health.qld.gov.au"
          $oldSuffix = "qh.health.qld.gov.au"
          Get-ADUser -Filter "UserPrincipalName -eq '$newupn'"| ForEach-Object {
             $updatedUpn = $_.UserPrincipalName.Replace($oldSuffix,$newSuffix)
                $_ | Set-ADUser -UserPrincipalName $updatedUpn  }

        Start-sleep -Seconds 60
        get-ADUser -Filter "UserPrincipalName  -eq '$updatedUpn'" |select userpr*
        #User will appear un licensed and with no Mailbox.

        wait-ToNextADSync

        Write-host 'Licenseing User' 
        Set-QHLicenseWithGeneric  -UserPrincipalName $updatedUpn
        }

        end
        {
        
        }

}

function Reset-Credentials 
{
    #... Clears all Variables that is used for Auth to conect
    $Global:OfficeAdminCredentials = $null
    $Global:ProxyServerCredentials = $null
    $Global:OnPremCredentials = $null
    $msolCredentials = $null
}


function Connect-ExchangeOnPrem 
{
    #... Connects to Exchange OnPrem and Prompts for user name and Password.      
    # Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010;
    Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010
  
    #     write-host "`n`n"
    show-line -color white
    Write-host "To disconnect this session, use Disconnect-ExchangeOnPrem cmdlet"
    show-line -color white
    test-exchangePremconnected
}


function Connect-msolserviceClixml 
{
    #... Connects to MSOL Service using Presaved Username and Password via CLIXML 
    $Global:msolCredFilePath = "$env:USERPROFILE\cred\msolCred.clixml"	
  
    if (-not (Test-Path $msolCredFilePath.Trim() -ErrorAction SilentlyContinue)) {
        Save-msolCred
    }

  
  
    #Get-PSSession | Remove-PSSession -ErrorAction SilentlyContinue
    $msolCredentials = Import-Clixml $msolCredFilePath -ErrorAction Stop
    If ($msolCredentials) {
        Connect-QhMSOLService -Credential $msolCredentials
        $Connectedas = Get-MsolUser -UserPrincipalName  ($msolCredentials.UserName)
        Write-host "Connected to MSOLService as: "$Connectedas.userPrincipalName
    
    }

} 


Function Connect-sharepointOnlineClixml 
{
    #... Connects to SPO Service using Presaved Username and Password via CLIXML 
    $Global:SPOLCredFilePath = "$env:USERPROFILE\Cred\SPOLCred.clixml"
    if (-not (Test-Path $SPOLCredFilePath.Trim() -ErrorAction SilentlyContinue)) {
        Save-SPOlCred
    }

    #Get-PSSession | Remove-PSSession -ErrorAction SilentlyContinue
    $SPOLCredentials = Import-Clixml $Global:SPOLCredFilePath -ErrorAction Stop
    If ($SPOLCredentials) {
        Connect-SPOService -Url https://healthqld-admin.sharepoint.com -Credential $SPOLCredentials
        Write-host " You are now connected to SPOService" -ForegroundColor Green
    }
}


function connect-365msolservice 
{
    #... Connects to MSOL Service and Prompts for user name and Password.
    If (-not $msolCredentials) {
        write-host "passing username x-$($env:USERNAME)@health.qld.gov.au" 
        $Global:msolCredentials = Get-Credential "x-$($env:USERNAME)@health.qld.gov.au" -Message "Enter on Exchange O365 credentials"
    }

    #$msolCredentials = Get-Credential
    If ($msolCredentials) {
        Connect-QhMSOLService -Credential $msolCredentials
        $Connectedas = Get-MsolUser -UserPrincipalName  ($msolCredentials.UserName)
        Write-host "Connected to MSOLService as: "$Connectedas.userPrincipalName
    }
    
}


function Connect-exchangeonlineO365Clixml 
{
    #... Connects to Exchange ON-Line using Presaved Username and Password via CLIXML 
  
    $Global:CredFilePath = "$env:USERPROFILE\cred\O365Cred.clixml"	

    if (-not (Test-Path $CredFilePath.Trim() -ErrorAction SilentlyContinue )) {
        Save-O365Cred
    }

    Get-PSSession | where-object {$_.ComputerName -eq "outlook.office.com"} | Remove-PSSession -ErrorAction SilentlyContinue
    $UserCredentials = Import-Clixml $CredFilePath -ErrorAction Stop
  		
      Connect-QhO365 -Credential $UserCredentials 
       # Import-PSSession $Session1 -prefix "EXO" -AllowClobber 
        #Import-PSSession $Session  -AllowClobber
        write-host "`n"
        show-line 
        Write-host "To disconnect this session, use Disconnect-ExchangeOnline cmdlet"
        show-line 
        write-host "Cmdlets are prefixed with EXO"
        show-line 
  
  
} 


function Connect-ExchangeOnPremclixml 
{
    #... Connects to Exchange ONPrem using Presaved Username and Password via CLIXML 
    param (
        $global:OnPremCredentials = $(import-clixml "$env:USERPROFILE\cred\OnPremCred.clixml")
    )

    Try {
        Connect-QHOnpremExchange -Server exc-casbk7p003
        
        <## $CredFilePathonprem = "$env:USERPROFILE\cred\OnPremCred.clixml"
        Write-Host "INFO : Trying to Connect to OnPrem Exchange using $($OnPremCredentials.Username)" -ForegroundColor Black -BackgroundColor Yellow
        Get-PSSession | Where-Object {$_.ComputerName -eq "exc-chmndcp001.qh.health.qld.gov.au"} | Remove-PSSession -ErrorAction SilentlyContinue
        #$OnPremCredentials = Import-Clixml $credFilePathonprem -ErrorAction Stop
        IF ($OnPremCredentials -eq $null) {
            $OnPremCredentials = get-credential -ErrorAction Stop -Message "Enter your Credentials"
        }
        
        #If ($OnPremCredentials -ne $null) {
        $sessionOptions = New-PSSessionOption -SkipCNCheck
        $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "http://exc-chmndcp001.qh.health.qld.gov.au/PowerShell/" -Credential $OnPremCredentials -Authentication Kerberos -SessionOption $sessionOptions -ErrorAction Stop
        Import-PSSession $Session -AllowClobber -ErrorAction Stop -WarningAction SilentlyContinue |Out-Null
        # Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010
            

        #write-host "`n`n`n"
        show-line -color White
        Write-host "To disconnect this session, use Disconnect-ExchangeOnPrem cmdlet"
        
        #>
        show-line -color White
        Write-Host "INFO : Successfully Connected to OnPrem Exchange using $($OnPremCredentials.Username)" -ForegroundColor White -BackgroundColor green


    }
    Catch {
        Write-Host "ERROR : $($error[0].Exception.Message)" -ForegroundColor White -BackgroundColor Red
    }
}


function Connect-skypeonlineClixml 
{
    #... Connects to Skype ON-Line using Presaved Username and Password via CLIXML 
    $CsolCredFilePath = "$env:USERPROFILE\cred\CsoCred.clixml"	
    if (-not (Test-Path $CsolCredFilePath.Trim() )) {
        Save-CSoCred
    }

    $CSonlineCredentials = Import-Clixml $CsolCredFilePath -ErrorAction Stop

    If (-not $CSonlineCredentials) {
        Save-csolCred
    }
    If ($CSonlineCredentials) {
        Connect-QhSkypeOnline -Credential $CSonlineCredentials
    }  
}

function Save-O365Cred 
{
    #... Save / Update Exchange On-Line Creds to the CLIXML file
    param (
        $Global:O365CredFilePath = "$env:USERPROFILE\cred\O365Cred.clixml"		)
    Try {
        $cred = $Host.ui.PromptForCredential("Office 365 credentials", "Please enter your Office 365 user name and password.", "$env:USERNAME@health.qld.gov.au", "O365 Creds")
			
        $cred | Export-Clixml $Global:O365CredFilePath
			
    }
    Catch {
        $errormsg = $($Error[0].exception.message)
    }
		
}


function Save-onpremCred 
{
    #... Save / Update Exchange On-Prem Creds to the CLIXML file
    param (
        $onpremCredFilePath = "$env:USERPROFILE\cred\OnPremCred.clixml"		)
    Try {
        $cred = $Host.ui.PromptForCredential("Local Exchange credentials", "Please enter your Local Exchange Admin user name and password.", "QH\$env:USERNAME", "Loc Creds")
         			
        $cred | Export-Clixml  $onpremCredFilePath
			
    }
    Catch {
        $errormsg = $($Error[0].exception.message)
    }
		
}


function Save-msolCred 
{
    #... Save / Update MSOL Creds to the CLIXML file
    param (
        $Global:msolCredFilePath = "$env:USERPROFILE\cred\MsolCred.clixml"		)
    Try {
        $cred = $Host.ui.PromptForCredential(" MSOL User Credentials", "Enter MSOL (Azure on line) User Credentials", "$env:USERNAME@health.qld.gov.au", "MSOL Creds")
    			
        $cred | Export-Clixml $Global:msolCredFilePath
			
    }
    Catch {
        $errormsg = $($Error[0].exception.message)
    }
		
}   


Function Save-csolCred 
{
    #... Save / Update Skype On-Line Creds to the CLIXML file
    param (
        $Global:CSOCredFilePath = "$env:USERPROFILE\cred\csoCred.clixml"		)
    Try {
        $cred = Get-Credential -UserName $env:USERNAME@health.qld.gov.au -Message "Enter O365 Skype User Credentials :" -ErrorAction Stop
			
        $cred | Export-Clixml $Global:csoCredFilePath
			
    }
    Catch {
        $errormsg = $($Error[0].exception.message)
    }


}  


Function Save-SPOlCred 
{
    #... Save / Update Sharepoint On-Line Creds to the CLIXML file
    param (
        $Global:SPOLCredFilePath = "$env:USERPROFILE\cred\SPOLCred.clixml"		)
    Try {
        $spolcred = $Host.ui.PromptForCredential(" SPOL User Credentials", "Enter SPOL (SharePoint online) User Credentials", "$env:USERNAME@health.qld.gov.au", "SPO Creds")
			
        $spolcred | Export-Clixml $Global:SPOLCredFilePath
			
    }
    Catch {
        $errormsg = $($Error[0].exception.message)
    }


}  


Function test-msolserviceconnected 
{
    #... Checks Connection Status and Reconnects if required for MSOL Service
    Get-MsolDomain -ErrorAction SilentlyContinue |Out-Null

    if ($?) {
        write-host   "You have a Valid connection to MsolService.." -ForegroundColor green
    }
    else {
        show-line -color Red
        write-host '         You Do not have a connection to MsolService'-ForegroundColor Yellow
        Write-host '         Connecting... '-ForegroundColor Yellow
        show-line -color Red
        Connect-msolserviceClixml
    }
  
}


Function test-exchangeonlineconnected 
{
    #... Checks Connection Status and Reconnects if required for Exchange On-Line
 
    if (Get-PSSession | where-object {$_.ComputerName -match "outlook.office"}) {
        write-host   "You have a Valid connection to Office365.." -ForegroundColor green
    }
    else {
        show-line -color Red
        write-host  "You Do not have a connection to Exchange Online" -ForegroundColor yellow
        Write-host '        Connecting... '-ForegroundColor Yellow
        show-line -color Red
      
        Connect-exchangeonlineO365Clixml  
    }
}


Function test-exchangePremconnected 
{
    #... Checks Connection Status and Reconnects if required for Exchange On-Prem
    #Uncomment this line and comment out the next   this is bebecause using PSSNAPIN 
    #if(Get-PSSession | where-object {$_.ComputerName -match ".qh.health.qld.gov.a"})
    if (Get-PSSnapin | where-object {$_.Name -match "Microsoft.Exchange.Management.PowerShell.E2010"}) {
        write-host   "You have a Valid connection to Exchange On Premise.." -ForegroundColor green
    }
    else {
        show-line -color Red
        write-host  'You Do not have a connection to Exchange On-Premise' -ForegroundColor Yellow
        Write-host '        Connecting... '-ForegroundColor Yellow
        show-line -color Red

        Connect-ExchangeOnPrem
        #   Connect-ExchangeOnPremclixml
    }

}


Function test-skypeonlineClixml 
{
    #... Checks Connection Status and Reconnects if required for Skype Online  
    if (Get-PSSession | where-object {$_.ComputerName -match "online.lync.com"}) {
        write-host   "You have a Valid connection to Skype Online.." -ForegroundColor green
    }
    else {


        show-line -color Red
        write-host 'You Do not have a connection to Skype Online' -ForegroundColor yellow
        Write-host '        Connecting... '-ForegroundColor Yellow
        show-line -color Red
      
        Connect-skypeonlineClixml
    }

}


Function test-sharepointonlineClixml 
{
    Try {
    
        $spol = get-sposite https://healthqld.sharepoint.com -ErrorAction SilentlyContinue
        if ($spol -match 'Microsoft.Online.SharePoint.PowerShell.SPOSite') {
            write-host 'You have a Valid connection to Sharepoint Online' -ForegroundColor green
        } 
    }

    Catch {
        # Write-Host "ERROR : $($error[0].Exception.Message)" -ForegroundColor White -BackgroundColor Red
 
        show-line -color Red
        write-host '         You do not have a connection'-ForegroundColor Yellow
        Write-host '         Connecting... '-ForegroundColor Yellow
        show-line -color Red
        Connect-sharepointOnlineClixml
    } 
   
}


Function Connect-all 
{
    #... Connects to all services at once using CLIXML
    test-exchangePremconnected
    test-msolserviceconnected
    test-exchangeonlineconnected
    test-skypeonlineClixml
    test-sharepointonlineClixml

}


function get-smileyface 
{
  #1... Displays a Smiley Face for the hell of it.
  $out="
    |||||      
    , ; ,   .-VVVVV-.   , ; ,
    \\|/  .'         '.  \|//
    \-;-/   ()   ()   \-;-/
    // ;               ; \\
    //__; :.         .; ;__\\
    `-----\'.'-.....-'.'/-----'
           '.'.-.-,_.'.'
             '(  (..-'
               '-'
                "
  write-host $out -ForegroundColor yellow
}


Function show-line 
{
    #... Draws a Line 70 Char in Yellow.  See Func.  Will take input to change color / Char /Length
    Param (
        [Parameter(Mandatory = $false)]  [String]$Char2use = "═", #default
        [Parameter(Mandatory = $false)]  [String]$numofchar = "70", #default
        [Parameter(Mandatory = $false)]  [String]$color = "yellow"  #default
    )
    # the dedault is Double Line 70 long and yellow 
  
    $xline = $Char2use * $NumofChar
    Write-host $xline -ForegroundColor $color

}


Function get-help2 
{
<#
    #... Displays the Functions in QHSupport.
    
#>

	param
	(
	
	)
   	
	BEGIN
    	{
         Write-Host $version -ForegroundColor Green
		}

      PROCESS
        {
            #... Display Functions with the comments
            $DDSPLAY = ""
            #Reads in the Current Powershell script file
            $DDSPLAY = get-content 'C:\Windows\System32\WindowsPowerShell\v1.0\Modules\QHO365MigrationOps\QHO365MigrationOps.psm1'

            foreach ($line in $DDSPLAY) {
                if ($line.Trim().StartsWith('Function', "CurrentCultureIgnoreCase") -or $line.Trim().Startswith('#...', "CurrentCultureIgnoreCase")) {
                    $1 = $line

                    if ($1.Trim().StartsWith('#...', "CurrentCultureIgnoreCase")) {
                
                        #Removes the first 4 char's Eg "#... "
                        $linedes = $line.trim().substring(4)
                        Write-Host $linedes -f Gray -NoNewline
                    }
                    Elseif (!($1.Trim().Startswith('#...'))) {
                        Write-host ''
                    }
                    if ($1.Trim().StartsWith('Function', "CurrentCultureIgnoreCase")) {
                        $linelong = $line + "                                               "
                
                        #makes the Line length to be 50 so the comments all line up. Fills it up with a space.
                        $line = $linelong.substring(0, 50)
                        Write-host "  $line" -f green -NoNewline
                    }
                }
            }
        
        }
   END
   {
    Write-Output  ""
   }
}


function Sync-github
{
copy-item 'C:\Windows\System32\WindowsPowerShell\v1.0\Modules\QHSupport\QHSupport.psm1' 'C:\Users\x-mckaylaw\OneDrive - Queensland Health(1)\Documents\GitHub\QH' -Force

}


Function sync-qhmodule 
{
    #... copies QHO365MigrationOps.psm1 module to 
    [string]$sourcefiles = '\\exc-mgtbk7p001\c$\Windows\System32\WindowsPowerShell\v1.0\Modules\QHO365MigrationOps\*'
    [string]$destinationDir = 'c:\Windows\System32\WindowsPowerShell\v1.0\Modules\QHO365MigrationOps\'
    copy-item -force -Recurse $sourcefiles -Destination $destinationDir


    [string]$sourcefiles = '\\exc-mgtbk7p001\c$\Windows\System32\WindowsPowerShell\v1.0\Modules\QHSupport\*'
    [string]$destinationDir = 'c:\Windows\System32\WindowsPowerShell\v1.0\Modules\QHSupport\'
    copy-item -Recurse $sourcefiles -Destination $destinationDir -force 
    Remove-Module QHSupport -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 10
    import-module QHSupport -Verbose
}


function Get-userfromValadationFile 
{
    #... This will get the list of emailaddress's from the Valadated file that are NOT Excluded.'
    <#

Syntax 
    Get-userfromValadationFile -pathToFileToImport 'D:\Office365\Migrations\Batch\Batch13\Batch13_Validation.csv' -inputVar qwerty


    This will get the list of emailaddress's from the Valadated file that are not Excluded.
#>

    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [String[]]$batchname,
        [Parameter(Mandatory = $true)]
        $outputVar
        
    )
	
    BEGIN {
	
    }

    PROCESS {
        Navigate-QHMigrationFolder -BatchName $Batchname
        new-Variable -name $outputVar -Scope Global  -force -value (import-csv ".\$($batchname)_Validation.csv"|where-object {$_.lookup -eq 'PASSED' } |select-object  -ExpandProperty emailaddress ) 
            
    }

}

function Get-qhAllusersnotPassedFromValadationFile 
{
    <#
#... This will get the list of emailaddress's from the Valadated file that are Excluded.'
Syntax 
    Get-ExcludeduserfromValadationFile -pathToFileToImport 'D:\Office365\Migrations\Batch\Batch13\Batch13_Validation.csv' -inputVar qwerty

#>

    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [String[]]$batchname,
        [Parameter(Mandatory = $true)]
        $ouputVarname
        
    )
	
    BEGIN {
	

    }

    PROCESS {
        Navigate-QHMigrationFolder -BatchName $Batchname
        new-Variable -name $ouputVarname -Scope Global  -force -value (import-csv ".\$($batchname)_Validation.csv"|where-object {$_.lookup -notmatch 'passed' } |select-object  -ExpandProperty emailaddress  ) 
            
    }

}


function Get-ExcludeduserfromValadationFile 
{
    <#
#... This will get the list of emailaddress's from the Valadated file that are Excluded.'
Syntax 
    Get-ExcludeduserfromValadationFile -pathToFileToImport 'D:\Office365\Migrations\Batch\Batch13\Batch13_Validation.csv' -inputVar qwerty

#>

    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [String[]]$batchname,
        [Parameter(Mandatory = $true)]
        $ouputVarname
        
    )
	
    BEGIN {
	

    }

    PROCESS {
        Navigate-QHMigrationFolder -BatchName $Batchname
        new-Variable -name $ouputVarname -Scope Global  -force -value (import-csv ".\$($batchname)_Validation.csv"|where-object {$_.lookup -eq 'Excluded' } |select-object  -ExpandProperty emailaddress  ) 
            
    }

}


Function Get-qhmigrationBatchFailure 
{
    #... This will get the list of emailaddress's from the MigrationBatch that has failed.'
    param
    (
        [Parameter(Mandatory = $true,
            HelpMessage = 'Please Enter the Batchname.')]
        [ValidateNotNullOrEmpty()]
        $batchname
    )

    Begin {
        $batches = Get-EXOMigrationBatch |Where-Object {$_.identity -match $batchname + '_'}

    }

    process {
        try {
            foreach ($line in $batches) {
                Get-EXOMigrationUser -BatchId $line.Identity.ToString() -ResultSize unlimited |Where-Object { ($_.Status -notlike "Synced" -and $_.Status -notlike "Syncing"-and $_.Status -notlike "STARTING")}  |Select-Object Identity, BatchID, Status, ErrorSummary 
            }    
    
        }
                
        catch {
        
        }
    }
    end {
        
    }
}


Function Get-qhmigrationBatchDups 
{
    #... This will get the list of emailaddress's from the MigrationBatch that has failed.'
    param
    (
        [Parameter(Mandatory = $true,
            HelpMessage = 'Please Enter the Batchname.')]
        [ValidateNotNullOrEmpty()]
        $batchname
    )
	
	
	
    $batches = Get-EXOMigrationBatch | Where-Object { $_.identity -match $batchname }
	
    foreach ($line in $batches) {
        #Get-EXOMigrationUser -BatchId $line.Identity.ToString()  |where {$_.Status -eq "failed"}  |Select-Object Identity, BatchID, Status,ErrorSummary |sort ErrorSummary  
        Get-EXOMigrationUser -BatchId $line.Identity.ToString() -ResultSize unlimited -WarningAction SilentlyContinue |
            Where-Object { $_.Status -eq "failed" } | Where-Object { $_.ErrorSummary -match "is already included in migration batch" } |
            Select-Object Identity, BatchID, Status, ErrorSummary | Sort-Object ErrorSummary
    }
}


function remove-qhexcludeduser 
{
    <#
Syntax 
    remove-excludeduser -BulkUsers 

#>

    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [String[]]$bulkUsers
        
    )
	
    BEGIN {
	
    }

    PROCESS {
   
        foreach ($line in $bulkusers) {
         
            ifexist (Get-EXOMigrationUser -Identity $line -ErrorAction Silentlycontinue)
            {Wtite-host 'exist'}
            else {
                write-host ' not htere'
            }

            Remove-EXOMigrationUser -Identity $line -Confirm:$false
        }           
    }

}


Function Get-qhmigrationUsersStatus 
{
    #... Gets the Migration Status from a Batch.
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [String[]]$Batch
        
    )
	

    BEGIN {
    Get-userfromValadationFile -batchname $batch -outputVar bulkusers

    }
    PROCESS {


        foreach ($line in $bulkUsers) {


        Try {
            $name = Get-EXOMigrationUser -Identity $line 
            #  Write-host $line" is in Batchname " $name -ForegroundColor green
        
            $prop = [ordered] @{
  'UserPrincipalName                     '  = $line
                  'Batchname             ' = $name.batchID
                                  Status   = $name.status

          }


        }
        Catch {
            $prop = [ordered] @{
                UserPrincipalName = $Line
                Batchname         = 'No migration request found' 
            }
        }
        Finally {
        
            $obj = New-Object -TypeName System.Management.Automation.PSObject -Property $prop
            Write-Output $obj
    
        	if ($bulkUsers.count -gt 1)
				{
					$paramWriteProgress = @{
						Activity = 'Validating Migration Request'
						Status = "Processing [$i] of [$($bulkUsers.Count)] users"
						PercentComplete = (($i / $bulkUsers.Count) * 100)
						CurrentOperation = "Checked : [$line]"
					}
					
					Write-Progress @paramWriteProgress
				}
				$i++

      }
            
    }

  }
}
    

Function Get-FileComparesForInsentraAddRemovedusers 
{
    #
    <#
  #... This will create the files for 'File compares for insentra add/removed users .'
  .Syntax
   
 Get-FileComparesForInsentraAddRemovedusers -file2oldusers D:\Office365\Migrations\Batch\FTC_Insentra_Combined_Files\Batch13_Batch13-HSQ\Batch13_Validation.csv  -file2newusers D:\Office365\Migrations\Batch\FTC_Insentra_Combined_Files\Batch13_Batch13-HSQ\Batch13_Batch13-HSQ.csv -fileremovesname D:\Office365\Migrations\Batch\FTC_Insentra_Combined_Files\Batch13_Batch13-HSQ\Batch13removed.csv -fileaddsname D:\Office365\Migrations\Batch\FTC_Insentra_Combined_Files\Batch13_Batch13-HSQ\Batch13added.csv
 

  #>
    #
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        $file2oldusers,
        [Parameter(Mandatory = $true)]
        $file2newusers,
        [Parameter(Mandatory = $true)]
        $fileaddsname,
        [Parameter(Mandatory = $true)]
        $fileremovesname
        
    )
 
    BEGIN {
        Write-host 'File compares for insentra add/removed users ' -ForegroundColor green
    }
    PROCESS {
        $oldusers = import-csv $file2oldusers #eg.  D:\Office365\Migrations\Batch\FTC_Insentra_Combined_Files\Batch13_Batch13-HSQ\Batch13_Validation.csv
        $newusers = import-csv $file2newusers #er.  D:\Office365\Migrations\Batch\FTC_Insentra_Combined_Files\Batch13_Batch13-HSQ\Batch13_Batch13-HSQ.csv
        Compare-object -ReferenceObject $oldusers.samaccountname -DifferenceObject $newusers.samaccountname |  Where-Object -FilterScript { $_.SideIndicator -eq '<='} | forEach-object { $_.InputObject }  | Out-File $fileremovesname
        Compare-object -ReferenceObject $oldusers.samaccountname -DifferenceObject $newusers.samaccountname |  Where-Object -FilterScript { $_.SideIndicator -eq '=>'} | forEach-object { $_.InputObject }  | Out-File $fileaddsname
            
    }   
    
}


Function publish-qhvalidationfile 
{
    #
    <#
      #... Created Valadation File and Reference File for Existing the batchname.
   
      .Syntax
   
      create-validationfile -batchname <Batchname>
 

  #>
    #
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        $Batchname
       
    )
	
        
    BEGIN {
       $WindowTitle =$host.ui.rawui.WindowTitle  # getting Window title name
        $host.ui.rawui.WindowTitle = 'Creating a Valadation File '   # Setting window title
        Write-host 'Running Process to create Validation file for ' $batchname  -ForegroundColor green
               
    }
    PROCESS {
        
        Navigate-QHMigrationFolder $batchname
        $filename = $($batchname + '_Validation.csv')
         
        if ( -not ($($batchname + '_Validation.csv')|test-path)) {
            Write-host 'file doe not  exist'
            Write-host 'Creating a Validation file for: '$batchname -ForegroundColor green  
            $users = Get-Content $batchname'.txt' | Sort-Object -unique
            Validate-QHMigrationUsers -UserPrincipalName $users -BatchName $batchname | export-csv $($batchname + '_Validation.csv') -NoTypeInformation
          <#     
            Write-host 'Creating a ReferenceList file for ' $batchname  -ForegroundColor green  
            Create-QHReferenceList -UserCSV $batchname'_Validation.csv' |export-csv  $($batchname + '_ReferenceList.csv') -NoTypeInformation
            #>            
        } 
         
        else {
            Write-host 'Valadation file ' $($batchname + '_Validation.csv')  ' exist' -foregroundColor yellow
            Write-host '     Please Delete and rerun if you want top overwrite it.' -foregroundColor yellow
        
        }   
         
         #this excludes all non approved mailbox moves
         revoke-qhexcludedfrommigration -batchname $batchname
        
        <#
        if ( -not ($($batchname + '_ReferenceList.csv')|test-path)) {
            Write-host 'Creating a ReferenceList file for ' $batchname  -ForegroundColor green  
            Create-QHReferenceList -UserCSV $batchname'_Validation.csv' |export-csv  $($batchname + '_ReferenceList.csv') -NoTypeInformation
        }
       
         
        else {
            Write-host 'ReferenceList file ' $($batchname + '_ReferenceList.csv')  ' exist' -foregroundColor yellow
            Write-host 'Please Delete and rerun if you want top overwrite it.' -foregroundColor yellow
        
        } 
         #>  
    }
    end
    {
        
        $host.ui.rawui.WindowTitle =  $WindowTitle  # Setting window title back to default
    }
}


Function publish-qhvalidationfilelockdown 
{
    #
    <#
      #... Created Valadation LockDown file for Existing batchname.
   
      .Syntax
   
      create-validationfilelockdown -batchname <Batchname>
 

  #>
    #
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        $Batchname
    )
	   
    BEGIN {
       
        Write-host 'Running Process to create Validation Lockdown file for ' $batchname  -ForegroundColor green
               
    }
    PROCESS {
        
        Navigate-QHMigrationFolder $batchname
        $filename = $($batchname + '__LockDown.csv')
         
        if ( -not ($($batchname + '_Validation.csv')|test-path)) {
            Write-host 'Validation file doe not exist' -ForegroundColor green 
            Write-host 'Please Create a Validation file for: '$batchname ' First' -ForegroundColor green  
        } 
         
        else {
            if ( -not ($($batchname + '_LockDown.csv')|test-path)) {
                Write-host 'Creating a Reference file for Batch: '$batchname -ForegroundColor green  
                Copy-Item $batchname'_Validation.csv' $batchname'_LockDown.csv'
            } 
            else {
                Write-host $($batchname + '_LockDown.csv')' file Exists, If you want to OverWrite the file please delete and Re-Run '-ForegroundColor yellow
                
            }
        
        }   

    }
}


Function publish-qhPremigrationEmailReport 
{
    #
    <#
      #... Created and Email the Premigration report for a Batch
   
      .Syntax
   
      create-validationfile -batchname <Batchname>
 

  #>
    #
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        $Batchname,
        [Parameter(Mandatory = $true)]
        $Recipients,
        $DeliveryName,
        $when = 'Saturday'
       
    )
	
        
    BEGIN {
       
       
               
    }
    PROCESS {


    foreach ($Batch in  $($batchname -split ",")){
         Write-host 'Generating the Pre Email Migration Report for: ' $batch  -ForegroundColor green
        Navigate-QHMigrationFolder -BatchName $Batch |out-null
        # get users that are not Excluded
        new-Variable -name ScheduledForCutover  -force -value (import-csv ".\$($batch)_Validation.csv"|where-object {$_.lookup -match 'Passed' } |select-object  -ExpandProperty emailaddress ).count
        new-Variable -name Totalmailboxes   -force -value (import-csv ".\$($batch)_Validation.csv").count
        new-Variable -name ExcludeForCutover  -force -value (import-csv ".\$($batch)_Validation.csv"|where-object {$_.lookup -notmatch 'Passed' } |select-object  -ExpandProperty emailaddress ).count

        


        #EMAIL sTUFF
        $batnameCAPS = $batch.ToUpper()
        $from = (get-aduser $env:username).UserPrincipalName
        #$Recipients = 'lawrence.mckay@health.qld.gov.au' 
        $displayname = (get-aduser $env:username).givenname + " "+ (get-aduser $env:username).surname
        $body = "Hi All,
         <STYLE>
  table {border: thick outset; padding 3px}
  td {border: thin inset; margin: 3px}
</STYLE>
         <br><br>
         </br><b>&nbsp;&nbsp;&nbsp;&nbsp;<u>   Batch Name:</u> $batnameCAPS  <br>
         <U>Delivery Name:</U> $deliveryname </b><br><br><br> 
         Cutover numbers for $When are below:</BR>
         </BR> 
        
        <table> 
        <tr> 
        <td><b>$batnameCAPS </b></td> 
        <td>$Totalmailboxes </td> 
        <td>(Original number before Pre Processing)  </td>
        </tr> 
        <tr> 
        <td><b>Already Excluded:</b></td> 
        <td>$ExcludeForCutover </TD>
        <td>(Due to account disabled, already moved or other errors from the original list provided) </td> 
        </tr> 
          
        <tr>
         <td><b>Scheduled for Cutover $when :</b></td> 
         <td>$ScheduledForCutover </td> 
         </td>
         <td>
         </TD>
         </tr> 
         </table></br>
         </br>
         </br>
         Regards </br>
         </br>
         <b>$Displayname  </B>    

        "

         
        Write-host 'Sending Email.  Please Wait..'  -ForegroundColor Darkgreen
        Send-MailMessage -Body $body -BodyAsHtml -From $from -Subject "$batnameCAPS Migrations for this $when . ( $deliveryName )" -To $Recipients -SmtpServer qhsmtp.qh.health.qld.gov.au
        
        }
    }
}


Function clear-VoipNumber 
{
    #... Removes a Voip Licence and clears the Phone number
    <# 
    ********************************************************************************  
    *                                                                              *  
    *       This script Removes a Voip Licence and clears the Phone number         *
    *                 for a user that has one assigned                             *  
    *                                                                              *  
    ********************************************************************************    
    Note.
        Needs connection to Office Skype with a prefix of 365 and connection to MSOL with no prefix.
        
    SYNTAX.
        Set-VoipTeams_cleanup_number.ps1 -number "tel:+61730822860"

      
    *******************
    Copyright Notice.
    *******************
    Copyright (c) 2018   McKayIT Solutions Pty Ltd.

    Permission is hereby granted, free of charge, to any person obtaining a copy
    of this software and associated documentation files (the "Software"), to deal
    in the Software without restriction, including without limitation the rights
    to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
    copies of the Software, and to permit persons to whom the Software is
    furnished to do so, subject to the following conditions:

    The above copyright notice and this permission notice shall be included in all
    copies or substantial portions of the Software.

    THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
    IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
    FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
    AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
    LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
    OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
    SOFTWARE.

    

    Author:
    Lawrence McKay
    Lawrence@mckayit.com
    McKayIT Solutions Pty Ltd
    
    Date:    10 Sept 2018
   

    ******* Update Version number below when a change is done.*******

    History
    Version         Date                Name           Detail
    ---------------------------------------------------------------------------------------
    0.0.1           10 Sept 2018       Lawrence       Initial Coding
    
   



#>


    param
    (
        [Parameter(Mandatory = $true,
            HelpMessage = 'Please Enter the Phone number.')]
        [ValidateNotNullOrEmpty()]
        $number
    )
  
    BEGIN {
        Write-host ' This Script Removes a Voip Licence and clears the Phone number ' -ForegroundColor Green
    }

    PROCESS {
        $info = Get-365CsOnlineUser -filter "lineuri -like '*$number'" |Select-Object Displayname, UserPrincipalName, LineURI -ErrorAction stop

        if ($info -eq $null) {

            write-host 'No Number found in use for that number.' -ForegroundColor yellow -BackgroundColor DarkRed
                
        }
        else {
                
            write-host "`n`nNumber Current setup for this number:" $info.LineURI  -ForegroundColor Magenta
            Write-host "________________________________________________________________________________" -ForegroundColor Cyan
            write-host $info.DisplayName "   " $info.UserPrincipalName "   " $info.LineURI -ForegroundColor Cyan
            
            if (get-MSOLUser -UserPrincipalName $info.UserPrincipalName | Where-Object { $_.Licenses.AccountSKUID -like "healthqld:MCOEV"}  ) {
                    
                write-host 'Removing Phone License for User: '$info.DisplayName '   ' $info.UserPrincipalName  -ForegroundColor green
                Set-MsolUserLicense -UserPrincipalName $info.UserPrincipalName  -removeLicenses healthqld:MCOEV 
            }
            else {
                    
                Write-host 'No Phone Licence found for for USER:'$info.DisplayName '   ' $info.UserPrincipalName  -ForegroundColor green
                    
            }

            
            # Remove old LineURI
                
            write-host 'Removing Enterprise Voice and onpremlineuri from user ' $old -ForegroundColor green
            Set-365CsUser -identity $info.UserPrincipalName -onpremlineuri $null -EnterpriseVoiceEnabled $false
                    
            $info = Get-365CsOnlineUser -filter "lineuri -like '$number'" |Select-Object Displayname, UserPrincipalName, LineURI 
            if ($info -eq $null) {
                write-host 'This number is not in use now'  $info.LineURI -ForegroundColor yellow 
                                                
            }          

        }
    }
     
    END {
                
    }
   
}


function set-VoIPTeams_Setup_NewNumber 
{
#... Sets up a new Voip Phone Number
    <# 
    ********************************************************************************  
    *                                                                              *  
    *             This script Assigns a ViOP Licence and set therm up              *
    *                                                                              *  
    ********************************************************************************    
    Note.
        Needs connection to Office Skype with a prefix of 365 and connection to MSOL with no prefix.
        
    SYNTAX.
        Set-NewViOPforTeams.ps1 -user Paul.Ng3@health.qld.gov.au -number "tel:+61730822860" 

      
    *******************
    Copyright Notice.
    *******************
    Copyright (c) 2018   McKayIT Solutions Pty Ltd.

    Permission is hereby granted, free of charge, to any person obtaining a copy
    of this software and associated documentation files (the "Software"), to deal
    in the Software without restriction, including without limitation the rights
    to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
    copies of the Software, and to permit persons to whom the Software is
    furnished to do so, subject to the following conditions:

    The above copyright notice and this permission notice shall be included in all
    copies or substantial portions of the Software.

    THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
    IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
    FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
    AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
    LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
    OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
    SOFTWARE.

    

    Author:
    Lawrence McKay
    Lawrence@mckayit.com
    McKayIT Solutions Pty Ltd
    
    Date:    07 Aug 2018
   

    ******* Update Version number below when a change is done.*******

    History
    Version         Date                Name           Detail
    ---------------------------------------------------------------------------------------
    0.0.1           07 Aug 2018       Lawrence       Initial Coding
    0.0.2           16 Aug 2018       Lawrence       Displaying Numer and user 
    0.0.3           11 Sep 2018       Lawrence       Updated to Function and did a recheck for Licenses befor Continuing 
   



#>


    param
    (
        [Parameter(Mandatory = $true,
            HelpMessage = 'Please Enter the USer as a UPN.')]
        [ValidateNotNullOrEmpty()]
        $user,
        [Parameter(Mandatory = $true)]
        $number
    
    )

    BEGIN {
        Write-host 'Setting up New Voip for User ' -ForegroundColor green
    }

    PROCESS {
        $info = Get-365CsOnlineUser -filter "lineuri -like '$number'" |Select-Object Displayname, UserPrincipalName, LineURI 
        write-host "`n`nNumber Current setup for this number:" -ForegroundColor Magenta
        Write-host "________________________________________________________________________________" -ForegroundColor Cyan
        write-host $info.DisplayName "   " $info.UserPrincipalName "   " $info.LineURI -ForegroundColor Cyan
            
        if (get-MSOLUser -UserPrincipalName $user | Where-Object { $_.Licenses.AccountSKUID -like "healthqld:MCOEV"}  ) {
            write-host 'Phone License is already Applied to User: ' $user -ForegroundColor green
        }
        else {
            Write-host 'Allocate Phone System Licence for USER:' $user  -ForegroundColor Green
            Set-MsolUserLicense -UserPrincipalName $user  -AddLicenses healthqld:MCOEV
            Write-host  'Waiting a couple of Min to make sure License is applied and for the Old Record to flush through So it can be used before Continuing.' -ForegroundColor Cyan
            Start-Sleep -Seconds 120

            if (get-MSOLUser -UserPrincipalName $user | Where-Object { $_.Licenses.AccountSKUID -notlike "healthqld:MCOEV"}  ) {
                Write-host ' License is still not applied waiting another 60 Sec'  -ForegroundColor Cyan
                Start-Sleep -Seconds 60
            }
               
        }


        # Set up new user
        Write-host 'Enabling Enterprise Voice and assigning the Number' -ForegroundColor Green
        Set-365CsUser -Identity $user -EnterpriseVoiceEnabled $true -HostedVoiceMail $true -OnPremLineURI $number
        
        Write-host 'Assigning Policyname National' -ForegroundColor Green   
        Grant-365CsOnlineVoiceRoutingPolicy -Identity $user -Policyname National

        Write-host 'Assigning Policyname DisallowOverrideCallingTeamsChatTeams' -ForegroundColor Green   
        #Grant-365CsTeamsInteropPolicy -Identity $user -PolicyName DisallowOverrideCallingTeamsChatTeams
        Grant-365CsTeamsUpgradePolicy -Identity $user -PolicyName UpgradeToTeams        
        # This adds the Call option to Teams.  
        Write-host 'Applying the PolicyName FederationOnly ' -ForegroundColor Green
        grant-365CsExternalAccessPolicy  -Identity $user -PolicyName FederationAndPICDefault 

        # This adds the Voice Routing policy.
        Write-host 'Applying the PolicyName InternationalCallsDisallowed2 ' -ForegroundColor Green
        Grant-365CsVoiceRoutingPolicy -Identity $user -PolicyName InternationalCallsDisallowed2

    }  

    End { 
        $info = Get-365CsOnlineUser -filter "lineuri -like '$number'" |Select-Object Displayname, UserPrincipalName, LineURI, RegistrarPool
        write-host "`n`nNumber Current setup for this number now:" -ForegroundColor Magenta
        Write-host "_________________________________________________________________________________________" -ForegroundColor Cyan
        write-host $info.DisplayName "   " $info.UserPrincipalName "   " $info.LineURI "   "  $info.RegistrarPool -ForegroundColor Cyan

    }
}


Function test-VoipNumberinuse 
{
    #... Checks VoipPhone Number is not used
    <# 
    ********************************************************************************  
    *                                                                              *  
    *              This script Checks the VoIP number status                       *
    *                 for a user that has one assigned                             *  
    *                                                                              *  
    ********************************************************************************    
    Note.
        Needs connection to Office Skype with a prefix of 365 and connection to MSOL with no prefix.
        
    SYNTAX.
        test-VoipNumberinuse -number "tel:+61730822860"

      
    *******************
    Copyright Notice.
    *******************
    Copyright (c) 2018   McKayIT Solutions Pty Ltd.

    Permission is hereby granted, free of charge, to any person obtaining a copy
    of this software and associated documentation files (the "Software"), to deal
    in the Software without restriction, including without limitation the rights
    to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
    copies of the Software, and to permit persons to whom the Software is
    furnished to do so, subject to the following conditions:

    The above copyright notice and this permission notice shall be included in all
    copies or substantial portions of the Software.

    THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
    IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
    FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
    AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
    LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
    OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
    SOFTWARE.

    

    Author:
    Lawrence McKay
    Lawrence@mckayit.com
    McKayIT Solutions Pty Ltd
    
    Date:    14 Sept 2018
   

    ******* Update Version number below when a change is done.*******

    History
    Version         Date                Name           Detail
    ---------------------------------------------------------------------------------------
    0.0.1           14 Sept 2018       Lawrence       Initial Coding
    
   



#>


    param
    (
        [Parameter(Mandatory = $true,
            HelpMessage = 'Please Enter the Phone number.')]
        [ValidateNotNullOrEmpty()]
        $number
    )
  
    BEGIN {
        Write-host ' This Script test Voip Number in use ' -ForegroundColor Green
    }

    PROCESS {
        $info = Get-365CsOnlineUser -filter "lineuri -like '*$number'" |Select-Object Displayname, UserPrincipalName, LineURI -ErrorAction stop

        if ($info -eq $null) 
            {

            write-host 'No Number found in use for that number.' -ForegroundColor yellow -BackgroundColor DarkRed
                
            }
        else 
            {
                
            write-host "`n`nNumber Current setup for this number:" $info.LineURI  -ForegroundColor Magenta
            Write-host "________________________________________________________________________________" -ForegroundColor Cyan
            write-host $info.DisplayName "   " $info.UserPrincipalName "   " $info.LineURI -ForegroundColor Cyan
            
                                
            }          

        }
    
    END 
    {
                
    }
    
}


Function wait-ToNextADSync
{
    #... Checks the AADSync time and Waits until the AADSync has run. Needs Connect-msolservice
    param (
    )
    BEGIN
    {
        $WindowTitle =$host.ui.rawui.WindowTitle  # getting Window title name
		$host.ui.rawui.WindowTitle = "Waiting for next AD Sync"
    }
    PROCESS
    {
        try
        {
            $strQuit = "Not yet"
            Do
            {
                $MsolInfo = Get-MsolCompanyInformation -ErrorAction Stop
                $lastSync = $msolInfo.LastDirSyncTime.ToLocalTime()
                $now = (get-date -ErrorAction Stop).ToLocalTime()
                $Duration = $now - $LastSync
                $durationMinsRaw = $duration.TotalMinutes
                $durationMins = [math]::Round($durationMinsRaw)
                $Timeleft = $([math]::Round(($lastSync.AddMinutes(30) - $now).TotalMinutes))
			
 
                if ($timeleft -eq '28') 
                {
                    write-host ' AAD has synced' -ForegroundColor Green
                    $strQuit = "n"
                }
                else
                {
                    Write-host 'Time to next AADSync Sync' $Timeleft -ForegroundColor YELLOW
                    start-sleep -Seconds 60
                }
            } # End of 'Do'

            Until($strQuit -eq 'N')

        }
        Catch
        {
            #$ErrorMsg = "ERROR : $($MyInvocation.InvocationName) `t`t$($error[0].Exception.Message)"
            Write-Host "ERROR : $($_.Exception.Message)" -ForegroundColor Magenta
        }
    }
    END
    {
        show-line
        Write-host '        AAD Synce has happened.' -ForegroundColor green
        Get-Date
        show-line
    $host.ui.rawui.WindowTitle =$WindowTitle # Setting Window title name back to what it was
    
    }
}


Function Publish-qhMailBoxMoveStats
{	
    #... Generates the MoverequestStats for the Batch.
<#
    Syntax   Publish-MailBoxMoveStats -batchname batch14

    This will generate the report for all of the batchname starting with Batch14*
#>


[CmdletBinding()]
	param
	(
		[Parameter(Mandatory = $true)]
		[String]$batchname
	)
	
	BEGIN
    	{
            Write-host 'Generating the Move Request Stats Report'
            $i = 1
		}

    PROCESS
        {

        Navigate-QHMigrationFolder $batchname

        $moveReqStats = Get-EXOMoveRequest -ResultSize Unlimited | Where-Object {($_.BatchName -like "*$batchname*")} | select-object -Expand Identity | Get-EXOMoveRequestStatistics
        #$moveReqStats = Get-MoveRequest -ResultSize Unlimited | select -Expand Identity | Get-MoveRequestStatistics

        $moveReqStats | Select-object MailboxIdentity,
                                DisplayName,
                                ExchangeGUID,
                                Status,
                                Flags,
                                Direction,
                                WorkLoadType,
                                RecipientTypeDetails,
                                SourceServer,
                                RemoteHostName,
                                BatchName,
                                RemoteCredentialUserName,
                                TargetDeliveryDomain,
                                BadItemLimit,
                                BadItemsEncountered,
                                LargeItemLimit,
                                LargeItemsEncountered,
                                QueuedTimestamp,
                                StartTimestamp,
                                LastUpdateTimestamp,
                                LastSuccessfulSyncTimestamp,
                                InitialSeedingCompletedTimestamp,
                                FinalSyncTimestamp,
                                CompletionTimestamp,
                                OverallDuration,
                                TotalSuspendedDuration,
                                TotalFailedDuration,
                                TotalQueuedDuration,
                                TotalInProgressDuration,
                                TotalMailboxSize,
                                @{n='TotalMailboxSize(MB)';e={[math]::Round(($_.TotalMailboxSize.ToString().Split("(")[1].Split(" ")[0].Replace(",","")/1MB),2)}},
                                TotalMailboxItemCount,
                                BytesTransferred,
                                @{n='BytesTransferred(MB)';e={[math]::Round(($_.BytesTransferred.ToString().Split("(")[1].Split(" ")[0].Replace(",","")/1MB),2)}},
                                ItemsTransferred,
                                PercentComplete,
                                Identity,
                                ObjectState |
                      Export-Csv .\AllMoveReqStats-$batchname.csv -NoTypeInformation -Force

        }

        END
        {
        
        }
	
}


Function Compare-ImmutableID
 {
   #... searches inSide Batch validation Passed users that have "PayRoll" in the name.
    <# 
    ********************************************************************************  
    *                                                                              *  
    *              This script Compares the  online and OnPremise.ImmutableID      *
    *                                                                              *  
    ********************************************************************************    
    Note.
        Needs connection to AD and MSOL with no Prefix.
        
    SYNTAX.
        Compare-ImmutableID -bulkUsers $users
        Compare-ImmutableID -bulkUsers $users |out-gridview
      
    *******************
    Copyright Notice.
    *******************
    Copyright (c) 2018   McKayIT Solutions Pty Ltd.

    Permission is hereby granted, free of charge, to any person obtaining a copy
    of this software and associated documentation files (the "Software"), to deal
    in the Software without restriction, including without limitation the rights
    to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
    copies of the Software, and to permit persons to whom the Software is
    furnished to do so, subject to the following conditions:

    The above copyright notice and this permission notice shall be included in all
    copies or substantial portions of the Software.

    THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
    IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
    FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
    AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
    LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
    OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
    SOFTWARE.

    

    Author:
    Lawrence McKay
    Lawrence@mckayit.com
    McKayIT Solutions Pty Ltd
    
    Date:    14 Sept 2018
   

    ******* Update Version number below when a change is done.*******

    History
    Version         Date                Name           Detail
    ---------------------------------------------------------------------------------------
    0.0.1           18 Sept 2018       Lawrence       Initial Coding
    
   #>

    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [String[]]$bulkUsers
        
    )

    begin
    {

    }

    Process 
    {
    
        foreach ($line in $bulkUsers){
            $onprem =get-aduser -Filter "UserPrincipalName  -eq '$line'"  -Properties msExchMailboxGuid |Select-Object userprincipalName, @{e={[system.convert]::ToBase64String($_.objectGuid.toByteArray())};l="ImmutableID"}

            $online =Get-MsolUser -UserPrincipalName $line -ErrorAction SilentlyContinue | Select-Object userprincipalName, ImmutableID 

                if ($online.ImmutableID -eq $onprem.ImmutableID){
        #        write-host $online.userprincipalName ' is the same both on Prem and Online' -ForegroundColor green
                $compared = 'Same'

                }
            else
                {
                $compared = 'Different'
                }

                    [PSCustomObject] @{ user = $line
            Matched = $compared
            onpremImmutableID = $onprem.ImmutableID
            onlineImmutableID = $online.ImmutableID
                }

        }

    }
    END
    {
        
    }
}


Function Compare-exchangewhenmailboxcreated
 {
   #... Compares the ExchangewhenmailboxCreated for a list of users onPremise and OnLine.
    <# 
    ********************************************************************************  
    *                                                                              *  
    *              This script Compares the  online and OnPremise.ImmutableID      *
    *                                                                              *  
    ********************************************************************************    
    Note.
        Needs connection to AD and MSOL with no Prefix.
        
    SYNTAX.
        Compare-exchangewhenmailboxcreated -Users $users
        Compare-exchangewhenmailboxcreated -Users $users |out-gridview
      
    *******************
    Copyright Notice.
    *******************
    Copyright (c) 2018   McKayIT Solutions Pty Ltd.

    Permission is hereby granted, free of charge, to any person obtaining a copy
    of this software and associated documentation files (the "Software"), to deal
    in the Software without restriction, including without limitation the rights
    to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
    copies of the Software, and to permit persons to whom the Software is
    furnished to do so, subject to the following conditions:

    The above copyright notice and this permission notice shall be included in all
    copies or substantial portions of the Software.

    THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
    IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
    FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
    AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
    LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
    OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
    SOFTWARE.

    

    Author:
    Lawrence McKay
    Lawrence@mckayit.com
    McKayIT Solutions Pty Ltd
    
    Date:    14 Sept 2018
   

    ******* Update Version number below when a change is done.*******

    History
    Version         Date                Name           Detail
    ---------------------------------------------------------------------------------------
    0.0.1           10 Oct 2018       Lawrence       Initial Coding
    
   #>

    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [String[]]$Users
        
    )

    begin
    {

    }

    Process 
    {
    
        foreach ($line in $Users){
               $onprem =Get-mailbox $line -erroraction SilentlyContinue
   $onpremgetrecipient = Get-Recipient -Identity $line -ErrorAction silentlycontinue

               $online =Get-exomailbox $line -erroraction SilentlyContinue
   $onlinegetrecipient = Get-exoRecipient -Identity $line  -erroraction SilentlyContinue
         Write-host 'getting Aduser'
              $enabled = Get-ADUser -Filter "UserPrincipalName -eq '$line'" -Property Enabled| Select Enabled
            if ($online.WhenMailboxCreated   -eq $onprem.WhenMailboxCreated ){
        #        write-host $online.userprincipalName ' is the same both on Prem and Online' -ForegroundColor green
                $compared = 'Same'

                }
            else
                {
                $compared = 'Different'
                }

            [PSCustomObject] @{ user = $line
            Matched = $compared
            AccountEnabled     = $enabled.enabled
            OnpremWhenmailboxCreated = $onprem.WhenmailboxCreated 
            
            onlineWhenmailboxCreated  = $online.WhenMailboxCreated 

            OnpremGetRecipient_Type = $onpremgetrecipient.RecipientTypeDetails 

            OnlineGetRecipient_Type = $onlinegetrecipient.RecipientTypeDetails 

                       
            
                }

        }

    }
    END
    {
        
    }
}


function Get-qhMailboxCountForBatches 
{
    #... This will get the Mailbox Number from the Valadated file that are passed For Migration.'
    <#

Syntax 
    Get-qhMailboxCountForBatches -batchname batch24,batch23


    *******************
    Copyright Notice.
    *******************
    Copyright (c) 2018   McKayIT Solutions Pty Ltd.

    Permission is hereby granted, free of charge, to any person obtaining a copy
    of this software and associated documentation files (the "Software"), to deal
    in the Software without restriction, including without limitation the rights
    to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
    copies of the Software, and to permit persons to whom the Software is
    furnished to do so, subject to the following conditions:

    The above copyright notice and this permission notice shall be included in all
    copies or substantial portions of the Software.

    THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
    IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
    FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
    AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
    LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
    OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
    SOFTWARE.

    

    Author:
    Lawrence McKay
    Lawrence@mckayit.com
    McKayIT Solutions Pty Ltd
    
    Date:    12 Oct 2018
   

    ******* Update Version number below when a change is done.*******

    History
    Version         Date                Name           Detail
    ---------------------------------------------------------------------------------------
    0.0.1           12 Oct 2018       Lawrence       Initial Coding
    
   #>


    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [String[]]$batchname        
    )
	
    BEGIN {
	
    Write-host 'Getting Mailbox Count for all Passed valadition for the Following Batches.' -ForegroundColor Green 
    Write-host '    '$batchname  -ForegroundColor green 
    
    }

    PROCESS {
    $total=0
    
foreach ($batch in $($batchname -split ",")){

        Navigate-QHMigrationFolder -BatchName $Batch |out-null
        $count = (import-csv ".\$($batch)_Validation.csv"|where-object {$_.lookup -eq 'PASSED' }).count
      
      $total = $total + $count

    [PSCustomObject] @{ 'Batch Name          ' = $Batch
                         'Mailbox Count' = $count

                }
      

            }
    Write-output "______________________________________"
    Write-output ""
    Write-Output "                      Total: $total"       
    Write-output "______________________________________"
        }
        END
        
        {
        
       
        }


}


Function search-QHBatchforPayRollinName
 {
    #... This Displays the Description when get-help1 is Run.
    <# 
    ********************************************************************************  
    *                                                                              *  
    *              This script IS a Blank Template                                 *
    *                                                                              *  
    ********************************************************************************    
    Note.
        Needs connection to What?? AD and MSOL with no Prefix.
        
    SYNTAX.
        <FunctionName -bulkUsers $users
        Compare-ImmutableID -bulkUsers $users |out-gridview
      
    *******************
    Copyright Notice.
    *******************
    Copyright (c) 2018   McKayIT Solutions Pty Ltd.

    Permission is hereby granted, free of charge, to any person obtaining a copy
    of this software and associated documentation files (the "Software"), to deal
    in the Software without restriction, including without limitation the rights
    to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
    copies of the Software, and to permit persons to whom the Software is
    furnished to do so, subject to the following conditions:

    The above copyright notice and this permission notice shall be included in all
    copies or substantial portions of the Software.

    THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
    IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
    FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
    AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
    LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
    OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
    SOFTWARE.

    

    Author:
    Lawrence McKay
    Lawrence@mckayit.com
    McKayIT Solutions Pty Ltd
    
    Date:    17 Oct 2018
   

    ******* Update Version number below when a change is done.*******

    History
    Version         Date                Name           Detail
    ---------------------------------------------------------------------------------------
    0.0.1           17  Oct 2018       Lawrence       Initial Coding
    
   #>

    [CmdletBinding()]
    param 
    (
        $Batchname
    )
    BEGIN {
    $result=$null
        Get-userfromValadationFile -batchname $Batchname -outputVar Process
          }
    PROCESS {
        
        Try {
		show-line 
        Write-Host '  Mailboxes that have Payroll in their Name for Batch:' $batchname -ForegroundColor Green
        show-line
        $result = $Process |where {$_ -match 'payroll'}

        if ($result -eq $null){
        Write-host 'No mailboxs found'
		    }
$result

        }
            
        Catch {
            # Catches error from Try 
            Write-Host "ERROR : $($_.Exception.Message)" -ForegroundColor Magenta
        }
    }
    END {
		
    }
}


Function remove-LicencesuserfromsharedMailboxes
{
 #... removeLicencesuserfromsharedMailboxes.
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [String[]]$Mailbox        
    )


    foreach ($Line1 in $mailbox)
    {
      #setting Usage Location
      
     $line = get-recipient $line1
     
      
        if ($line.recipienttypeDetails -match 'sharedmailbox'){

          #getting the USer info
          $lic=Get-MsolUser -UserPrincipalName $line.PrimarySmtpAddress

          
          #Getting the Licence Info 
          $licsku=$lic.Licenses.accountSku.SkuPartNumber

          Write-host  'Processing Licences for:'$lic.DisplayName -ForegroundColor Cyan


          #Checking and removeing E1 Licence    
          if ($licsku -match 'STANDARDPACK') {
    
            Write-host  'E1 STANDARDPACK Licence applied and Removing'  -ForegroundColor Green
                set-MsolUserLicense  -UserPrincipalName $lic.UserPrincipalName  -RemoveLicenses healthqld:STANDARDPACK -ErrorAction SilentlyContinue
       
          } 
          #Checking and removeing E2 Licence
          if ($licsku -match 'EXCHANGEENTERPRISE') {
    
            Write-host  'E2 EXCHANGEENTERPRISE Licence applied and Removing'  -ForegroundColor Green  -BackgroundColor darkRed
                set-MsolUserLicense  -UserPrincipalName $lic.UserPrincipalName  -RemoveLicenses healthqld:EXCHANGEENTERPRISE -ErrorAction SilentlyContinue
       
          } 

       
                 
       
       
    
       
       
        
        }

             #remove membership of ADM-O365-LIC-E3-OfficeProPlusOnly
            
                   try{
                      remove-adgroupmember -Identity 'ADM-O365-LIC-E3-OfficeProPlusOnly' -Members $line.UserPrimarySMTP  -ErrorAction SilentlyContinue 
                     }
                     
                     catch{}
            }                
    }


function publish-BatchDetailsforInsentra
{

    #... This will get the Mailbox Number from the Valadated file that are passed For Migration.'
    <#

Syntax 
    publish-BatchDetailsforInsentra -batchname batch24,batch23


    *******************
    Copyright Notice.
    *******************
    Copyright (c) 2018   McKayIT Solutions Pty Ltd.

    Permission is hereby granted, free of charge, to any person obtaining a copy
    of this software and associated documentation files (the "Software"), to deal
    in the Software without restriction, including without limitation the rights
    to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
    copies of the Software, and to permit persons to whom the Software is
    furnished to do so, subject to the following conditions:

    The above copyright notice and this permission notice shall be included in all
    copies or substantial portions of the Software.

    THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
    IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
    FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
    AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
    LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
    OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
    SOFTWARE.

    

    Author:
    Lawrence McKay
    Lawrence@mckayit.com
    McKayIT Solutions Pty Ltd
    
    Date:    20 Oct 2018
   

    ******* Update Version number below when a change is done.*******

    History
    Version         Date                Name           Detail
    ---------------------------------------------------------------------------------------
    0.0.1           20 Oct 2018       Lawrence       Initial Coding
    
   #>


	param (
		[String[]]$BatchNames
	)
	BEGIN 
	{
		
	}
	PROCESS
	{
		foreach ($line in $BatchNames)
		{
			
			if (Navigate-QHMigrationFolder -batchname $line)
			{
				$users = Import-Csv ".\$($line)_Validation.csv" | Where-Object { $_.lookup -eq 'Passed' } 
				
        foreach ($line in $users) {
                 [PSCustomObject] @{ EmailAddress = $line.EmailAddress
                                     SamAccountName = $line.SamAccountName
                                    }
                                }
			}
		}

	}
}
	

function merge-qhexcludedfrommigration
{
  #... This will find and mark as excluded users from a Valadation CSV file.'
    <#

Syntax 
    merge-qhexcludedfrommigration -file <Caladationfile -user <Users to exclude> -bywho < "Reason and by who in"> Note: the Quotes

    *******************
    Copyright Notice.
    *******************
    Copyright (c) 2018   McKayIT Solutions Pty Ltd.

    Permission is hereby granted, free of charge, to any person obtaining a copy
    of this software and associated documentation files (the "Software"), to deal
    in the Software without restriction, including without limitation the rights
    to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
    copies of the Software, and to permit persons to whom the Software is
    furnished to do so, subject to the following conditions:

    The above copyright notice and this permission notice shall be included in all
    copies or substantial portions of the Software.

    THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
    IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
    FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
    AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
    LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
    OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
    SOFTWARE.

    

    Author:
    Lawrence McKay
    Lawrence@mckayit.com
    McKayIT Solutions Pty Ltd
    
    Date:    27 Oct 2018
   

    ******* Update Version number below when a change is done.*******

    History
    Version         Date                Name           Detail
    ---------------------------------------------------------------------------------------
    0.0.1           27 Oct 2018       Lawrence       Initial Coding
    
   #>
       [CmdletBinding()]
    param
    (   [Parameter(Mandatory = $true)]
        [string]$batchname,
      #  [Parameter(Mandatory = $true)]
      #  [string]$file,
        [Parameter(Mandatory = $true)]
        [String[]]$users,
        [Parameter(Mandatory = $true)]
        [string]$bywho
    )
    
    BEGIN
    {
        Try
        {
         #   import-csv $($batchname + '_Validation1.csv') -ErrorAction Stop | Out-Null
        }
        Catch
        {
            
            Write-Host "ERROR : $($_.Exception.Message)" -ForegroundColor Magenta
            Break   
        }
    }
    PROCESS
    {
        Try
        {

              Navigate-QHMigrationFolder -BatchName $batchname |Out-Null
              $csv = import-csv $($batchname + '_Validation.csv') 
      
              write-host 'Processing the valadition file for :' $batchname -ForegroundColor Green
              Get-userfromValadationFile -batchname $batchname -outputVar users |Out-Null


            foreach ($user in $users)
            {
                foreach ($row in $csv)
                {
                    if ($row.emailaddress -eq $user)
                    {
                    write-host $user
                        $row.lookup = "Excluded"
                        $row.Details = $bywho
                    }
                }
            }
            $csv | Export-Csv $($batchname + '_Validation.csv') -NoTypeInformation
        }
        Catch
        {
            # Catches error from Try 
            Write-Host "ERROR : $($_.Exception.Message)" -ForegroundColor Magenta
        }
    }
    END
    {
        
    }
}



function revoke-qhExcludedFromMigration
{
  #... This will find and mark as excluded users that are not approved to be migrated.'
    <#

Syntax 
    revoke-qhexcludedfrommigration -batchname <batchname>


Note:  List below.

    *******************
    Copyright Notice.
    *******************
    Copyright (c) 2018   McKayIT Solutions Pty Ltd.

    Permission is hereby granted, free of charge, to any person obtaining a copy
    of this software and associated documentation files (the "Software"), to deal
    in the Software without restriction, including without limitation the rights
    to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
    copies of the Software, and to permit persons to whom the Software is
    furnished to do so, subject to the following conditions:

    The above copyright notice and this permission notice shall be included in all
    copies or substantial portions of the Software.

    THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
    IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
    FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
    AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
    LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
    OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
    SOFTWARE.

    

    Author:
    Lawrence McKay
    Lawrence@mckayit.com
    McKayIT Solutions Pty Ltd
    
    Date:    31 Oct 2018
   

    ******* Update Version number below when a change is done.*******

    History
    Version         Date                Name           Detail
    ---------------------------------------------------------------------------------------
    0.0.1           31 Oct 2018       Lawrence       Initial Coding
    
   #>
 [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [string]$batchname
        
    )
    
    BEGIN
    {
    Write-host 'Processing the Valadition file to make sure Mailboxes Excluded by Project are' -ForegroundColor yellow -NoNewline
    Write-host '  EXCLUDED' -ForegroundColor Red

   
   $users =import-csv D:\Office365\Migrations\Batch\Exclusions.csv |select -ExpandProperty email
      
      Navigate-QHMigrationFolder -BatchName $batchname |Out-Null
      $csv = import-csv $($batchname + '_Validation.csv') 
      
      write-host 'Processing the valadition file for :' $batchname -ForegroundColor Green
      Get-userfromValadationFile -batchname $batchname -outputVar csv1 |Out-Null

        }
       
    
    PROCESS
    {
        Try
        {
            foreach ($user in $users)
            {
                foreach ($row in $csv)
                {
                    if ($row.emailaddress -eq $user)
                    {
                    Write-host 'The following User is not approved to be migrated.' $user -ForegroundColor Yellow
                        $row.lookup = "Excluded"
                        $row.Details = "Not Approved to be migrated by Project. "

                    Write-host 'Checking and Removing any Migration requests if exist.' -ForegroundColor green
                    remove-exomigrationuser $user -confirm:$false -erroraction silentlycontinue

                    }
                }
            }
            $csv | Export-Csv $($batchname + '_Validation.csv' ) -NoTypeInformation
        }
        Catch
        {
            # Catches error from Try 
            Write-Host "ERROR : $($_.Exception.Message)" -ForegroundColor Magenta
        }
    }
    END
    {
        
    }
}


#Displaying functions and Descriptions
#get-help1

function get-StressReleif 
{
#... This will help you relieve stress (Written in Powershell).'
D:\Office365\Scripts\stressrelief.ps1
}


Function add-DL_For_Rollout
{
    #... This will empty and Add users into the DL for Rollout.

  [CmdletBinding()]
    param 
    (
    	[Parameter(Mandatory = $true,
				  HelpMessage = 'Enter The Distrubution Group name to update?')]
        $DistrubutionGroupname,
        $batchnames

    )
    BEGIN {
        $title = 'Updating the Distrubution Group '+$DistrubutionGroupname + ' With users from ' + $batchnames
        $host.ui.rawui.WindowTitle = $Title

          }
    PROCESS {
        
        Try {
       $i = 0
        $users =(publish-BatchDetailsforInsentra -BatchNames $Batchnames |Select-Object emailaddress).emailaddress
            foreach ($line  in $users) 
            {
                write-host 'Adding: ' -ForegroundColor Green -NoNewline
                write-host  $line
				    
                Add-exoDistributionGroupMember -Identity $DistrubutionGroupname -Member $line -ErrorAction silentlycontinue
                
                if ($users.count -gt 1)
                        {
                            $paramWriteProgress = @{
                                Activity = 'Aadding users To DL'
                                Status = "Processing $line  [$i] of [$($users.Count)] users"
                                PercentComplete = (($i / $users.Count) * 100)
                                
                            }
                            
                            Write-Progress @paramWriteProgress
                        }
                        $i++
  
            }

				
        }
            
        Catch {
            # Catches error from Try 
            Write-Host "ERROR : $($_.Exception.Message)" -ForegroundColor Magenta
        }
    }
    END {
		
     }


}

Function update-DL_For_Rollout
{
    #... This will empty and Add users into the DL for Rollout.

  [CmdletBinding()]
    param 
    (
    	[Parameter(Mandatory = $true,
				  HelpMessage = 'Enter The Distrubution Group name to update?')]
        $DistrubutionGroupname,
        $batchnames

    )
    BEGIN {
        $title = 'Updating the Distrubution Group '+$DistrubutionGroupname + ' With users from ' + $batchnames
        $host.ui.rawui.WindowTitle = $Title

          }
    PROCESS {
        
        Try {
        Write-host 'Removing existing DL members' -ForegroundColor Green
        $users = Get-exoDistributionGroupMember -Identity $DistrubutionGroupname -resultsize unlimited
            foreach ($PrimarySmtpAddress in $Users) 
            {
               if ($users.count -gt 1)
    {
        $paramWriteProgress = @{
            Activity        = 'Removing user from the DL'
            Status          = "Processing $PrimarySmtpAddress.PrimarySmtpAddress  [$i] of [$($users.Count)] users"
            PercentComplete = (($i / $users.Count) * 100)
                                
        }
                            
        Write-Progress @paramWriteProgress
    }
    $i++

		   
            Remove-exoDistributionGroupMember -Identity $DistrubutionGroupname  -Member $PrimarySmtpAddress.PrimarySmtpAddress -Confirm:$False
            

            }
              clear-Xline 1   

        $users =(publish-BatchDetailsforInsentra -BatchNames $Batchnames |Select-Object emailaddress).emailaddress
            $i=""
            foreach ($line  in $users) 
            {
                write-host 'Adding: ' -ForegroundColor Green -NoNewline
                write-host  $line
				    
                Add-exoDistributionGroupMember -Identity $DistrubutionGroupname -Member $line
                clear-Xline 1 

                if ($users.count -gt 1)
                        {
                            $paramWriteProgress = @{
                                Activity = 're-adding user from DL'
                                Status = "Processing $user  [$i] of [$($users.Count)] users"
                                PercentComplete = (($i / $users.Count) * 100)
                                
                            }
                            
                            Write-Progress @paramWriteProgress
                        }
                        $i++
  
            }

				
        }
            
        Catch {
            # Catches error from Try 
            Write-Host "ERROR : $($_.Exception.Message)" -ForegroundColor Magenta
        }
    }
    END {
		
     }


}


function remove-qhdupMailbox
{
        #... Removes the Exchange Attrib for Duplicate mailbox.n.
    <# 
    ********************************************************************************  
    *                                                                              *  
    *              This script IS a Blank Template                                 *
    *                                                                              *  
    ********************************************************************************    
    Note.
        Need to add as exclusion to retention Policy before running.

        
    SYNTAX.
        <remove-qhdupMailbox -mailbox <emailaddress>


      
    *******************
    Copyright Notice.
    *******************
    Copyright (c) 2018   McKayIT Solutions Pty Ltd.

    Permission is hereby granted, free of charge, to any person obtaining a copy
    of this software and associated documentation files (the "Software"), to deal
    in the Software without restriction, including without limitation the rights
    to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
    copies of the Software, and to permit persons to whom the Software is
    furnished to do so, subject to the following conditions:

    The above copyright notice and this permission notice shall be included in all
    copies or substantial portions of the Software.

    THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
    IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
    FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
    AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
    LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
    OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
    SOFTWARE.

    

    Author:
    Lawrence McKay
    Lawrence@mckayit.com
    McKayIT Solutions Pty Ltd
    
    Date:    6 Nov 2018
   

    ******* Update Version number below when a change is done.*******

    History
    Version         Date                Name           Detail
    ---------------------------------------------------------------------------------------
    0.0.1           6 Nov 2018       Lawrence       Initial Coding
    
   #>

    [CmdletBinding()]
    param 
    (
        $MAilbox
    )
    BEGIN {
        
          }
    PROCESS {
        
        Try {
				Set-EXOMailbox $mailbox -removeDelayHoldApplied
write-host 'Now Manually Remove E2 and Skype license'
#pause

Set-EXOUser  $mailbox -PermanentlyClearPreviousMailboxInfo -confirm:$false

Start-Sleep 30
Write-host  ' The RecipientTypeDetails shoud now be showing as "User"' -ForegroundColor Green
Write-host  ' If not then wait 10Min and re run Get-EXOUser <UPN>  to check it has changed.' -ForegroundColor Green
Get-EXOUser $mailbox

        }
            
        Catch {
            # Catches error from Try 
            Write-Host "ERROR : $($_.Exception.Message)" -ForegroundColor Magenta
        }
    }
    END {
		
    }
}


Function new-qhbatch 
{
    #... Create new Migration Batch.
    <#

Syntax 
    new-qhbatch -Batchname batch28  -MigrationBatchName 1111  -migrationCSVfile batch28_fix.csv -MRSServers mrs4.health.qld.gov.au

Notes

    The MRS Servers are pre coded to use ('mrs.health.qld.gov.au', 'mrs1.health.qld.gov.au', 'mrs2.health.qld.gov.au', 'mrs3.health.qld.gov.au', 'mrs4.health.qld.gov.au')



    *******************
    Copyright Notice.
    *******************
    Copyright (c) 2018   McKayIT Solutions Pty Ltd.

    Permission is hereby granted, free of charge, to any person obtaining a copy
    of this software and associated documentation files (the "Software"), to deal
    in the Software without restriction, including without limitation the rights
    to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
    copies of the Software, and to permit persons to whom the Software is
    furnished to do so, subject to the following conditions:

    The above copyright notice and this permission notice shall be included in all
    copies or substantial portions of the Software.

    THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
    IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
    FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
    AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
    LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
    OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
    SOFTWARE.

    

    Author:
    Lawrence McKay
    Lawrence@mckayit.com
    McKayIT Solutions Pty Ltd
    
    Date:    7 Nov 2018
   

    ******* Update Version number below when a change is done.*******

    History
    Version         Date                Name           Detail
    ---------------------------------------------------------------------------------------
    0.0.1           7 Nov 2018       Lawrence       Initial Coding
    
   #>
    [CmdletBinding()]
    param 
    (
         $Batchname,
        [Parameter(Mandatory = $true)]
        $MigrationBatchName,
        [Parameter(Mandatory = $true)]
        $migrationCSVfile,
        [Parameter(Mandatory = $true)]
        [ValidateSet('mrs.health.qld.gov.au', 'mrs1.health.qld.gov.au', 'mrs2.health.qld.gov.au', 'mrs3.health.qld.gov.au', 'mrs4.health.qld.gov.au')]
        $MRSServers

    )
    BEGIN {
    Navigate-QHMigrationFolder $batchname
        
    }
    PROCESS {
        if (Navigate-QHMigrationFolder $batchname) {
            Try {
            $location =(Get-Item -Path ".\").FullName+'\'+$migrationCSVfile
             New-exoMigrationBatch -Name $MigrationBatchName -SourceEndpoint $MRSServers -TargetDeliveryDomain healthqld.mail.onmicrosoft.com -CSVData ([System.IO.File]::ReadAllBytes($location))  -AutoStart -LargeItemLimit 20 -BadItemLimit 200 -MoveOptions skipFolderRestrictions # (USE only if requested by M$ )-MoveOptions skipFolderRestrictions
            			
            }
            
            Catch {
                # Catches error from Try 
                Write-Host "ERROR : $($_.Exception.Message)" -ForegroundColor Magenta
            }
        }
    }
    END {
		
    }


}


Function new-qhbatchNoAutostart 
{
    #... Create new Migration Batch with AutoStart Disabled..
    <#

Syntax 
    new-qhbatchNoAutostart  -Batchname batch28  -MigrationBatchName 1111  -migrationCSVfile batch28_fix.csv -MRSServers mrs4.health.qld.gov.au

Notes

    The MRS Servers are pre coded to use ('mrs.health.qld.gov.au', 'mrs1.health.qld.gov.au', 'mrs2.health.qld.gov.au', 'mrs3.health.qld.gov.au', 'mrs4.health.qld.gov.au')



    *******************
    Copyright Notice.
    *******************
    Copyright (c) 2018   McKayIT Solutions Pty Ltd.

    Permission is hereby granted, free of charge, to any person obtaining a copy
    of this software and associated documentation files (the "Software"), to deal
    in the Software without restriction, including without limitation the rights
    to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
    copies of the Software, and to permit persons to whom the Software is
    furnished to do so, subject to the following conditions:

    The above copyright notice and this permission notice shall be included in all
    copies or substantial portions of the Software.

    THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
    IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
    FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
    AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
    LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
    OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
    SOFTWARE.

    

    Author:
    Lawrence McKay
    Lawrence@mckayit.com
    McKayIT Solutions Pty Ltd
    
    Date:    7 Nov 2018
   

    ******* Update Version number below when a change is done.*******

    History
    Version         Date                Name           Detail
    ---------------------------------------------------------------------------------------
    0.0.1           7 Nov 2018       Lawrence       Initial Coding
    
   #>
    [CmdletBinding()]
    param 
    (
         $Batchname,
        [Parameter(Mandatory = $true)]
        $MigrationBatchName,
        [Parameter(Mandatory = $true)]
        $migrationCSVfile,
        [Parameter(Mandatory = $true)]
        [ValidateSet('mrs.health.qld.gov.au', 'mrs1.health.qld.gov.au', 'mrs2.health.qld.gov.au', 'mrs3.health.qld.gov.au', 'mrs4.health.qld.gov.au')]
        $MRSServers

    )
    BEGIN {
    Navigate-QHMigrationFolder $batchname
        
    }
    PROCESS {
        if (Navigate-QHMigrationFolder $batchname) {
            Try {
            $location =(Get-Item -Path ".\").FullName+'\'+$migrationCSVfile
             New-exoMigrationBatch -Name $MigrationBatchName -SourceEndpoint $MRSServers -TargetDeliveryDomain healthqld.mail.onmicrosoft.com -CSVData ([System.IO.File]::ReadAllBytes($location))   -LargeItemLimit 20 -BadItemLimit 200 -MoveOptions skipFolderRestrictions # (USE only if requested by M$ )-MoveOptions skipFolderRestrictions
            			
            }
            
            Catch {
                # Catches error from Try 
                Write-Host "ERROR : $($_.Exception.Message)" -ForegroundColor Magenta
            }
        }
    }
    END {
		
    }


}

Function confirm-qhmigrationMatches
 {#... This will compare the batch to show what is missing from moves and the List

    [CmdletBinding()]
    param 
    (
        [Parameter(Mandatory = $true)]
        $Batchname,
        [Parameter(Mandatory = $true)]
        $batchnameonline


    )
    BEGIN {
    $WindowTitle =$host.ui.rawui.WindowTitle 
    $title = 'Checking Migration Matches '+$batchname
    $host.ui.rawui.WindowTitle = $title
    }
    PROCESS {

        if (Navigate-QHMigrationFolder $batchname) {
            Try {

                Write-host '      Getting users from Valadition file for: ' $batchname -ForegroundColor green
               
                get-userfromvaladationfile -batchname $batchname -outputVar Validation


                Write-host '      Getting users from the Migration batches that match the Batchname: ' $batchname -ForegroundColor Green
                Write-host '      NOTE:  ' -ForegroundColor red -NoNewline
                Write-host 'This will take some time..' -ForegroundColor cyan
                #$batchfix = $batchname + '_'
                $batchfix = $batchnameonline
                $online = Get-EXOMigrationUser -ResultSize unlimited |Where-Object {$_.batchid -match "$batchfix" } |Select-Object -ExpandProperty identity


                $notinValadationfile = $Online.Where( {$Validation -inotcontains $_})

                Write-host '   Count in Validationfile is:' $Validation.count
                Write-host 'Count in Online migration  is:' $online.count
                show-line
                Write-host 'Not in Valadation File'

                show-line
                $notinValadationfile

                $notmigrationrequest = $Validation.Where( {$Online -inotcontains $_})
                show-line
                Write-host 'No moverequest'
                show-line
                $notmigrationrequest
            }
            
            Catch {
                # Catches error from Try 
                Write-Host "ERROR : $($_.Exception.Message)" -ForegroundColor Magenta
            }
        }
    }
    END {
		$host.ui.rawui.WindowTitle = $WindowTitle #Setting Window title bact to what it was.
    }
}


Function remove-qhUserFromDl
 {
    #... Remove Excluded Users from DL
    <# 
    ********************************************************************************  
    *                                                                              *  
    *              This script IS a Blank Template                                 *
    *                                                                              *  
    ********************************************************************************    
    Note.
        Needs connection to What?? AD and MSOL with no Prefix.
        
    SYNTAX.
        remove-QHUserFromDl -batchname <Batchname> -DLName <DL NAme>
        remove-QHUserFromDl -batchname batch30-hsq -DLName DL-Office365-MetroNorth
      
    *******************
    Copyright Notice.
    *******************
    Copyright (c) 2018   McKayIT Solutions Pty Ltd.

    Permission is hereby granted, free of charge, to any person obtaining a copy
    of this software and associated documentation files (the "Software"), to deal
    in the Software without restriction, including without limitation the rights
    to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
    copies of the Software, and to permit persons to whom the Software is
    furnished to do so, subject to the following conditions:

    The above copyright notice and this permission notice shall be included in all
    copies or substantial portions of the Software.

    THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
    IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
    FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
    AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
    LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
    OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
    SOFTWARE.

    

    Author:
    Lawrence McKay
    Lawrence@mckayit.com
    McKayIT Solutions Pty Ltd
    
    Date:    9 Nov 2018
   

    ******* Update Version number below when a change is done.*******

    History
    Version         Date                Name           Detail
    ---------------------------------------------------------------------------------------
    0.0.1           9 Nov 2018        Lawrence       Initial Coding
    
   #>

    [CmdletBinding()]
    param 
    (
        $batchname,
        $DLName
    )
    BEGIN {
        
    }
    PROCESS {
        
        Try {
            Get-qhAllusersnotPassedFromValadationFile  -batchname $batchname -ouputVarname ToProcess 

            foreach ($line in $toProcess) {

            clear-Xline 1   
                Write-host 'Checking and removing if required :' -ForegroundColor Green -NoNewline
                Write-host $line -ForegroundColor Yellow -NoNewline
            write-host '                                                                                               '
                
                Remove-DistributionGroupMember -Identity $DLName -Member $line  -confirm:$false -erroraction SilentlyContinue
Start-sleep -Milliseconds 40                
            }


        }
            
        Catch {
            # Catches error from Try 
            Write-Host "ERROR : $($_.Exception.Message)" -ForegroundColor Magenta
        }
    }
    END {
		Write-host 'Updated'
    }
}


Function restart-qhModule
 {
     #... ReLoads the Modules QHSupport or QHO365MigrationOps

    [CmdletBinding()]
    param 
    (
        [Parameter(Mandatory = $true)]
        [ValidateSet('QHSupport', 'QHO365MigrationOps')]
        $ModuleName

    )
    BEGIN {
        
    }
    PROCESS {
        
        Try {
            Write-host 'Reloading Module '$ModuleName -ForegroundColor green
            Remove-module $ModuleName
            Start-Sleep -Seconds 1
            Import-Module $ModuleName
            }


        
            
        Catch {
            # Catches error from Try 
            Write-Host "ERROR : $($_.Exception.Message)" -ForegroundColor Magenta
        }
    }
    END {
		
    }
}

Function clear-Xline
{
 #... Removes x number of lines from the Display
 #Syntax clear-xline <number of lines>
    Param (

	        [Parameter(Position=1)]
	        [int32]$Count=1

        )

        $CurrentLine  = $Host.UI.RawUI.CursorPosition.Y
        $ConsoleWidth = $Host.UI.RawUI.BufferSize.Width

        $i = 1
        for ($i; $i -le $Count; $i++) {
	
	        [Console]::SetCursorPosition(0,($CurrentLine - $i))
	        [Console]::Write("{0,-$ConsoleWidth}" -f " ")

        }

        [Console]::SetCursorPosition(0,($CurrentLine - $Count))
}

Function get-qhmigrationbatch 
{
    #... This will get the users from a migration batch online.
    <# 
    ********************************************************************************  
    *                                                                              *  
    *  This scriptwill get all users from a migration batch Online not the file    *
    *                            in the migration batch                            *
    *  get-exomigrationuser -resultsize unlimited |where {$_.batchid -match <name> *
    *                                                                              *  
    ********************************************************************************    
    Note.
        Needs connection to What?? AD and MSOL with no Prefix.
        
    SYNTAX.
        <FunctionName -bulkUsers $users
        get-qhmigrationbatch -Migrationgroupname <migration batch name>
      
    *******************
    Copyright Notice.
    *******************
    Copyright (c) 2018   McKayIT Solutions Pty Ltd.

    Permission is hereby granted, free of charge, to any person obtaining a copy
    of this software and associated documentation files (the "Software"), to deal
    in the Software without restriction, including without limitation the rights
    to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
    copies of the Software, and to permit persons to whom the Software is
    furnished to do so, subject to the following conditions:

    The above copyright notice and this permission notice shall be included in all
    copies or substantial portions of the Software.

    THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
    IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
    FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
    AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
    LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
    OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
    SOFTWARE.

    

    Author:
    Lawrence McKay
    Lawrence@mckayit.com
    McKayIT Solutions Pty Ltd
    
    Date:   11 Nov 2018
   

    ******* Update Version number below when a change is done.*******

    History
    Version         Date                Name           Detail
    ---------------------------------------------------------------------------------------
    0.0.1           11 Nov 2018       Lawrence       Initial Coding
    
   #>

    [CmdletBinding()]
    param 
    (
        $Migrationgroupname
    )
    BEGIN {
        
          }
    PROCESS {
        
        Try {
				get-exomigrationuser -resultsize unlimited |where {$_.batchid -match $Migrationgroupname }
        }
            
        Catch {
            # Catches error from Try 
            Write-Host "ERROR : $($_.Exception.Message)" -ForegroundColor Magenta
        }
    }
    END {
		
    }
}


function connect-qhMSTeams 
{

    #... Connects to MS Teams  using Presaved Username and Password via CLIXML 

    <# 
    ********************************************************************************  
    *                                                                              *  
    *  This scriptwill get all users from a migration batch Online not the file    *
    *                            in the migration batch                            *
    *  get-exomigrationuser -resultsize unlimited |where {$_.batchid -match <name> *
    *                                                                              *  
    ********************************************************************************    
    Note.
        Needs connection to What?? nothing.
        
    SYNTAX.
        connect-qhMSTeams 
      
      
    *******************
    Copyright Notice.
    *******************
    Copyright (c) 2018   McKayIT Solutions Pty Ltd.

    Permission is hereby granted, free of charge, to any person obtaining a copy
    of this software and associated documentation files (the "Software"), to deal
    in the Software without restriction, including without limitation the rights
    to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
    copies of the Software, and to permit persons to whom the Software is
    furnished to do so, subject to the following conditions:

    The above copyright notice and this permission notice shall be included in all
    copies or substantial portions of the Software.

    THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
    IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
    FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
    AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
    LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
    OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
    SOFTWARE.

    

    Author:
    Lawrence McKay
    Lawrence@mckayit.com
    McKayIT Solutions Pty Ltd
    
    Date:   29 Nov 2018
   

    ******* Update Version number below when a change is done.*******

    History
    Version         Date                Name           Detail
    ---------------------------------------------------------------------------------------
    0.0.1           29 Nov 2018       Lawrence       Initial Coding
    
   #>


    [CmdletBinding()]
    param 
    (
        
    )
    BEGIN {
        
    }
    PROCESS {
        
        Try {
            $CsolCredFilePath = "$env:USERPROFILE\cred\CsoCred.clixml"	
            if (-not (Test-Path $CsolCredFilePath.Trim() )) {
                Save-CSoCred
            }

            $CSonlineCredentials = Import-Clixml $CsolCredFilePath -ErrorAction Stop

            If (-not $CSonlineCredentials) {
                Save-csolCred
            }
            If ($CSonlineCredentials) {
                Connect-MicrosoftTeams -Credential $CSonlineCredentials
            }  

            $host.ui.rawui.WindowTitle = 'Connection to MS Teams'
        }
            
        Catch {
            # Catches error from Try 
            Write-Host "ERROR : $($_.Exception.Message)" -ForegroundColor Magenta
        }
    }
    END {
		
    }
}


function Remove-usersFromValadationFile
{
        #... This will get the users from a migration batch online.
    <# 
    ********************************************************************************  
    *                                                                              *  
    *     This script will Mark as Excluded from a Valadation file with the        *
    *                  Reason entered  into the migration file                     *
    *                                                                              *  
    ********************************************************************************    
    Note.
        Needs connection to What?? None
        
    SYNTAX.
       
        Remove-usersFromValadationFile -Batchname -users -reason <In Quotes'>
      
    *******************
    Copyright Notice.
    *******************
    Copyright (c) 2018   McKayIT Solutions Pty Ltd.

    Permission is hereby granted, free of charge, to any person obtaining a copy
    of this software and associated documentation files (the "Software"), to deal
    in the Software without restriction, including without limitation the rights
    to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
    copies of the Software, and to permit persons to whom the Software is
    furnished to do so, subject to the following conditions:

    The above copyright notice and this permission notice shall be included in all
    copies or substantial portions of the Software.

    THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
    IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
    FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
    AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
    LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
    OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
    SOFTWARE.

    

    Author:
    Lawrence McKay
    Lawrence@mckayit.com
    McKayIT Solutions Pty Ltd
    
    Date:   11 Nov 2018
   

    ******* Update Version number below when a change is done.*******

    History
    Version         Date                Name           Detail
    ---------------------------------------------------------------------------------------
    0.0.1           3 Dec 2018       Lawrence       Initial Coding
    
   #>

 [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [string]$batchname,
         [Parameter(Mandatory = $true)]
        $users,
         [Parameter(Mandatory = $true)]
        [string]$Reason
        
    )
    
    BEGIN
    {
    Write-host 'Processing the Valadition file to make sure Mailboxes Excluded for: ' $reason  ' are' -ForegroundColor yellow -NoNewline
    Write-host '  EXCLUDED' -ForegroundColor Red

      
      Navigate-QHMigrationFolder -BatchName $batchname |Out-Null
      $csv = import-csv $($batchname + '_Validation.csv') 
      
      write-host 'Processing the valadition file for :' $batchname -ForegroundColor Green
      Get-userfromValadationFile -batchname $batchname -outputVar csv1 |Out-Null

        }
       
    
    PROCESS
    {
        Try
        {
            foreach ($user in $users)
            {
                foreach ($row in $csv)
                {
                    if ($row.emailaddress -eq $user)
                    {
                    Write-host 'The following User ' -ForegroundColor Yellow -NoNewline
                     Write-host  $user.padright(50,[char]$null) -ForegroundColor white -NoNewline
                  
                     Write-host  ' is getting excluded because ' -ForegroundColor Yellow -NoNewline
                     write-host $reason -ForegroundColor green
                        $row.lookup = "Excluded"
                        $row.Details = $Reason



                    }
                }
            }
            $csv | Export-Csv $($batchname + '_Validation.csv' ) -NoTypeInformation
        }
        Catch
        {
            # Catches error from Try 
            Write-Host "ERROR : $($_.Exception.Message)" -ForegroundColor Magenta
        }
    }
    END
    {
        
    }
}

Function get-QHOnlineDataSynced
{
    #...  This script to show how the total GB synced to online for a batch .
    <# 
    ********************************************************************************  
    *                                                                              *  
    *     This script to show how the total GB synced to online for a batch        *
    *                                                                              *  
    ********************************************************************************    
    Note.
        Needs connection to What?? Exchange Online with a Prefix of EXO
        
    SYNTAX.
        <get-QHOnlineDataSynced -Batchname xxx>
       get-QHOnlineDataSynced -Batchname Batch33
      
    *******************
    Copyright Notice.
    *******************
    Copyright (c) 2018   McKayIT Solutions Pty Ltd.

    Permission is hereby granted, free of charge, to any person obtaining a copy
    of this software and associated documentation files (the "Software"), to deal
    in the Software without restriction, including without limitation the rights
    to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
    copies of the Software, and to permit persons to whom the Software is
    furnished to do so, subject to the following conditions:

    The above copyright notice and this permission notice shall be included in all
    copies or substantial portions of the Software.

    THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
    IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
    FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
    AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
    LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
    OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
    SOFTWARE.

    

    Author:
    Lawrence McKay
    Lawrence@mckayit.com
    McKayIT Solutions Pty Ltd
    
    Date:    6th Dec 2018
   

    ******* Update Version number below when a change is done.*******

    History
    Version         Date                Name           Detail
    ---------------------------------------------------------------------------------------
    0.0.1           6th Dec 2018       Lawrence       Initial Coding
    
   #>

    [CmdletBinding()]
    param 
    (
        [Parameter(Mandatory = $true)]
        $Batchname
    )
    BEGIN {
        $Totaltransferd = 0
        $ESTTotaltransferd = 0
        $WindowTitle =$host.ui.rawui.WindowTitle  # getting Window title name
        $host.ui.rawui.WindowTitle = 'Getting total GB synced to online for a batch  '   # Setting window title
          }
    PROCESS {
        
            Get-userfromValadationFile -batchname $batchname -outputVar fromlist
        try{
                    foreach ($user in $fromlist){
                        
                        clear-Xline
                        Write-host 'Processing ' -ForegroundColor green -NoNewline
                        write-host $user  -foregroundColor White

                        $migrationmbs = Get-exoMigrationUserStatistics $user -erroraction silentlycontinue
                        $BytesTransferred =$migrationmbs | Select @{n='BytesTransferred(MB)';e={[math]::Round(($_.BytesTransferred.ToString().Split("(")[1].Split(" ")[0].Replace(",","")/1MB),2)}}
                        $EstimatedTotalTransferSize=  $migrationmbs | Select @{n='EstimatedTotalTransferSize(MB)';e={[math]::Round(($_.EstimatedTotalTransferSize.ToString().Split("(")[1].Split(" ")[0].Replace(",","")/1MB),2)}}

                        $Totaltransferd +=$($BytesTransferred.'BytesTransferred(MB)')
                        $ESTTotaltransferd +=$($EstimatedTotalTransferSize.'EstimatedTotalTransferSize(MB)')
            
                        #$obj = New-Object -TypeName System.Management.Automation.PSObject -Property $prop
                    
            
                    if ($fromlist.count -gt 1)
                        {
                            $paramWriteProgress = @{
                                Activity = 'Adding up data Transfered'
                                Status = "Processing [$i] of [$($fromlist.Count)] users"
                                PercentComplete = (($i / $fromlist.Count) * 100)
                                
                            }
                            
                            Write-Progress @paramWriteProgress
                        }
                        $i++
                } 
            }

    catch {
            # Catches error from Try 
            Write-Host "ERROR : $($_.Exception.Message)" -ForegroundColor Magenta
          }
    
    
    Finally {

            }

        show-line
            write-host "        Estimated Total Transfer Size MB: $ESTTotaltransferd"

            write-host "             Actual Total Transferred MB: $Totaltransferd"
        show-line    
        $host.ui.rawui.WindowTitle = $WindowTitle #Setting Window title bact to what it was.
 }   

}


Function get-qhDupEmailaddressInValadationFile 
{
    #... This will display the Duplicate emailaddress's in a Valadation File.
    <# 
    ********************************************************************************  
    *                                                                              *  
    *      This will display the Duplicate emailaddress's in a Valadation File.    *
    *                                                                              *  
    ********************************************************************************    
    Note.
        Needs connection to What?? AD and MSOL with no Prefix.
        
    SYNTAX.
        <qet-qhDuplicateEmailaddressInValadationFile -Batchname >
       qet-qhDuplicateEmailaddressInValadationFile -batchname batch33
      
    *******************
    Copyright Notice.
    *******************
    Copyright (c) 2018   McKayIT Solutions Pty Ltd.

    Permission is hereby granted, free of charge, to any person obtaining a copy
    of this software and associated documentation files (the "Software"), to deal
    in the Software without restriction, including without limitation the rights
    to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
    copies of the Software, and to permit persons to whom the Software is
    furnished to do so, subject to the following conditions:

    The above copyright notice and this permission notice shall be included in all
    copies or substantial portions of the Software.

    THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
    IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
    FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
    AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
    LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
    OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
    SOFTWARE.

    

    Author:
    Lawrence McKay
    Lawrence@mckayit.com
    McKayIT Solutions Pty Ltd
    
    Date:    6th Dec 2018
   

    ******* Update Version number below when a change is done.*******

    History
    Version         Date                Name           Detail
    ---------------------------------------------------------------------------------------
    0.0.1           6th Dec 2018       Lawrence       Initial Coding
    
   #>

    [CmdletBinding()]
    param 
    (
        $batchname
    )
    BEGIN {
        Get-userfromValadationFile -batchname $batchname -outputVar batch
        $WindowTitle =$host.ui.rawui.WindowTitle  # getting Window title name
        $host.ui.rawui.WindowTitle = "Geting Duplicate emailaddress's in a Valadation File"   # Setting window title        
          }
    PROCESS {
        
        Try {
        
            $batch |group |where {$_.count -ne 1}



        }
            
        Catch {
            # Catches error from Try 
            Write-Host "ERROR : $($_.Exception.Message)" -ForegroundColor Magenta
        }
    }
    END {
		$host.ui.rawui.WindowTitle = $WindowTitle #Setting Window title bact to what it was.
    }
}


Function Get-QHLicenseFileErrors
 {
    #... This Displays all users that have an error in the License file 
    <# 
       ********************************************************************************  
       *                                                                              *  
       *       gets all users that have an error in the License file                  *
       *                                                                              *  
       ********************************************************************************    
       Note.
        Needs connection to What?? None
        
       SYNTAX.
        <Get-QHLicenseFileErrors -Batchname >
        Get-QHLicenseFileErrors -Batchname batch32
      
       *******************
       Copyright Notice.
       *******************
       Copyright (c) 2018   McKayIT Solutions Pty Ltd.

       Permission is hereby granted, free of charge, to any person obtaining a copy
       of this software and associated documentation files (the "Software"), to deal
       in the Software without restriction, including without limitation the rights
       to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
       copies of the Software, and to permit persons to whom the Software is
       furnished to do so, subject to the following conditions:

       The above copyright notice and this permission notice shall be included in all
       copies or substantial portions of the Software.

       THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
       IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
       FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
       AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
       LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
       OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
       SOFTWARE.

    

       Author:
       Lawrence McKay
       Lawrence@mckayit.com
       McKayIT Solutions Pty Ltd
    
       Date:    7th Dec 2018
   

       ******* Update Version number below when a change is done.*******

       History
       Version         Date                Name           Detail
       ---------------------------------------------------------------------------------------
       0.0.1           7th Dec 2018        Lawrence       Initial Coding
    
   #>

    [CmdletBinding()]
    param 
    (
      [Parameter(Mandatory = $true)]
        $Batchname
    )
    BEGIN {
        show-line
        Write-host "   Below are the errors in the $($batchname + '_License.csv')  file if exist." -ForegroundColor green
        show-line
          }
    PROCESS {
        
        Try {
              Navigate-QHMigrationFolder -BatchName $batchname |Out-Null
       $csv = import-csv $($batchname + '_License.csv') 

       foreach ($user in $CSV){ 
         if ( $user.Status -notmatch 'Success' -and $user.details -notmatch 'ERROR : User Not Found')
         {
        [PSCustomObject] @{ 
                         'UserPrincipalName' = $user.UserPrincipalName
                         'ErrorDetails' = $user.Details

                }
         }
       }
				
        }
            
        Catch {
            # Catches error from Try 
            Write-Host "ERROR : $($_.Exception.Message)" -ForegroundColor Magenta
        }
    }
    END {
		
    }
 }

 Function Get-QHMFAFileErrors
 {
    #... This Displays all users that have an error in the MFA file 
    <# 
       ********************************************************************************  
       *                                                                              *  
       *          gets all users that have an error in the MFA file                   *
       *                                                                              *  
       ********************************************************************************    
       Note.
        Needs connection to What?? None
        
       SYNTAX.
        <Get-QHMFAFileErrors -Batchname >
        Get-QHMFAFileErrors -Batchname batch32
      
       *******************
       Copyright Notice.
       *******************
       Copyright (c) 2018   McKayIT Solutions Pty Ltd.

       Permission is hereby granted, free of charge, to any person obtaining a copy
       of this software and associated documentation files (the "Software"), to deal
       in the Software without restriction, including without limitation the rights
       to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
       copies of the Software, and to permit persons to whom the Software is
       furnished to do so, subject to the following conditions:

       The above copyright notice and this permission notice shall be included in all
       copies or substantial portions of the Software.

       THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
       IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
       FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
       AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
       LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
       OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
       SOFTWARE.

    

       Author:
       Lawrence McKay
       Lawrence@mckayit.com
       McKayIT Solutions Pty Ltd
    
       Date:    7th Dec 2018
   

       ******* Update Version number below when a change is done.*******

       History
       Version         Date                Name           Detail
       ---------------------------------------------------------------------------------------
       0.0.1           7th Dec 2018        Lawrence       Initial Coding
    
   #>

    [CmdletBinding()]
    param 
    (
      [Parameter(Mandatory = $true)]
        $Batchname
    )
    BEGIN {
        show-line
        Write-host "   Below are the errors in the $($batchname + '_MFA.csv')  file if exist." -ForegroundColor green
        show-line
          }
    PROCESS {
        
        Try {
              Navigate-QHMigrationFolder -BatchName $batchname |Out-Null
       $csv = import-csv $($batchname + '_MFA.csv') 

       foreach ($user in $CSV){ 
         if ( $user.Status -notmatch 'Success' -and $User.details -notmatch 'skipped' )
         {



           [PSCustomObject] @{ 
                         'UserPrincipalName' = $user.EmailAddress
                         'Status' = $user.status
                         'ErrorDetails' = $user.Details

                }

         }
       }
				
        }
            
        Catch {
            # Catches error from Try 
            Write-Host "ERROR : $($_.Exception.Message)" -ForegroundColor Magenta
        }
    }
    END {
		
    }
 }


 
 Function find-whenmoved
 {
    #... Find when a user was moved from Batch processed files.
    <# 
       ********************************************************************************  
       *                                                                              *  
       *              This script will find when user moved                           *
       *                                                                              *  
       ********************************************************************************    
       Note.
        Needs connection to What?? None
        
       SYNTAX.
        <find-whenmoved -emailaddress
        find-whenmoved -emailaddress lawrence@health.qld.gov.au
      
       *******************
       Copyright Notice.
       *******************
       Copyright (c) 2018   McKayIT Solutions Pty Ltd.

       Permission is hereby granted, free of charge, to any person obtaining a copy
       of this software and associated documentation files (the "Software"), to deal
       in the Software without restriction, including without limitation the rights
       to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
       copies of the Software, and to permit persons to whom the Software is
       furnished to do so, subject to the following conditions:

       The above copyright notice and this permission notice shall be included in all
       copies or substantial portions of the Software.

       THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
       IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
       FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
       AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
       LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
       OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
       SOFTWARE.

    

       Author:
       Lawrence McKay
       Lawrence@mckayit.com
       McKayIT Solutions Pty Ltd
    
       Date:    14 Sept 2018
   

       ******* Update Version number below when a change is done.*******

       History
       Version         Date                Name           Detail
       ---------------------------------------------------------------------------------------
       0.0.1           18 Sept 2018       Lawrence       Initial Coding
    
   #>

    [CmdletBinding()]
    param 
    (
      [Parameter(Mandatory = $true)]
        $emailaddress
    )
    BEGIN {
            $WindowTitle =$host.ui.rawui.WindowTitle  # getting Window title name
            $host.ui.rawui.WindowTitle = "Finding when user was moved"
            $host.ui.rawui.WindowTitle = 'finding when a user Was moved'   # Setting window title
          show-line -numofchar 90
          write-host " See Below for where $emailaddress was found " -ForegroundColor Green
          show-line -numofchar 90

          }
    PROCESS {
        
        Try {
        

          dir D:\Office365\Migrations\Batch\*validation.csv -Recurse |select-string -pattern $emailaddress |select-object Filename, Path
				
           }
            
        Catch {
            # Catches error from Try 
            Write-Host "ERROR : $($_.Exception.Message)" -ForegroundColor Magenta
        }
    }
    END {
		$host.ui.rawui.WindowTitle = $WindowTitle #Setting Window title bact to what it was.
        }
 }

Function find-Duplicatemaiboxes
 {
    #... checks for Duplicate MAilboxes in Crowd and on Premise..
    <# 
       ********************************************************************************  
       *                                                                              *  
       *              This script find Duplicate mailboxes                            *
       *                                                                              *  
       ********************************************************************************    
       Note.
        Needs connection to What?? Connection to Exchange on Premise and Exchange Online with a EXO prefix.
        
       SYNTAX.
        <find-Duplicatemaiboxes -uses
        find-Duplicatemaiboxes -users John.Doe@Contoso.com or Array($VARS )
      
       *******************
       Copyright Notice.
       *******************
       Copyright (c) 2018   McKayIT Solutions Pty Ltd.

       Permission is hereby granted, free of charge, to any person obtaining a copy
       of this software and associated documentation files (the "Software"), to deal
       in the Software without restriction, including without limitation the rights
       to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
       copies of the Software, and to permit persons to whom the Software is
       furnished to do so, subject to the following conditions:

       The above copyright notice and this permission notice shall be included in all
       copies or substantial portions of the Software.

       THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
       IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
       FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
       AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
       LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
       OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
       SOFTWARE.

    

       Author:
       Lawrence McKay
       Lawrence@mckayit.com
       McKayIT Solutions Pty Ltd
    
       Date:    10th Dec 2018
   

       ******* Update Version number below when a change is done.*******

       History
       Version         Date                Name           Detail
       ---------------------------------------------------------------------------------------
       0.0.1           10th Dec 2018       Lawrence       Initial Coding
    
   #>

    [CmdletBinding()]
    param 
    (
      [Parameter(Mandatory = $true)]
        $Users
    )
    BEGIN {
        
        $WindowTitle =$host.ui.rawui.WindowTitle  # getting Window title name
        $host.ui.rawui.WindowTitle = 'Checking for Mailbox Duplicates'   # Setting window title
      

          }
    PROCESS {
        
        Try {
        foreach ($user in $users){
            $user | ForEach { $onpremexist = [bool](Get-mailbox $_ -erroraction SilentlyContinue)}
            $user | ForEach { $365exist = [bool](Get-exomailbox $_ -erroraction SilentlyContinue)}
 

        If ($onpremexist -eq $365Exist){
            
             [PSCustomObject] @{ 
                         'UserPrincipalName' = $user
                         'Status' = "Duplicate MAilbox found.   ***** WARNING ***"
                         }
            $duplicates += $user 
         }
         Else {


                }

   if ($users.count -gt 1)
                        {
                            $paramWriteProgress = @{
                                Activity = 'Checking for Duplicate Mailboxes'
                                Status = "Processing $user  [$i] of [$($users.Count)] users"
                                PercentComplete = (($i / $users.Count) * 100)
                                
                            }
                            
                            Write-Progress @paramWriteProgress
                        }
                        $i++

           }

     }            
        Catch {
            # Catches error from Try 
            Write-Host "ERROR : $($_.Exception.Message)" -ForegroundColor Magenta
        }
    }
    END {
    show-line 
	write-host "Processed $($users.count) users and the following users have a Duplicate Mailbox."
    Show-line
   

	  
    $host.ui.rawui.WindowTitle = $WindowTitle #Setting Window title bact to what it was.
    }
 }

function get-QHOnpremdatatoSync
{
    #... This Displays the Description when get-help1 is Run.
    <# 
    ********************************************************************************  
    *                                                                              *  
    *        This script gets the Onpremise Mailbox Sizes for a Batch              *
    *                                                                              *  
    ********************************************************************************    
    Note.
        Needs connection to What?? AD and MSOL with no Prefix.
        
    SYNTAX.
        <get-QHOnpremdatatoSync -Batchname
        get-QHOnpremdatatoSync -batchname batch33
      
    *******************
    Copyright Notice.
    *******************
    Copyright (c) 2018   McKayIT Solutions Pty Ltd.

    Permission is hereby granted, free of charge, to any person obtaining a copy
    of this software and associated documentation files (the "Software"), to deal
    in the Software without restriction, including without limitation the rights
    to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
    copies of the Software, and to permit persons to whom the Software is
    furnished to do so, subject to the following conditions:

    The above copyright notice and this permission notice shall be included in all
    copies or substantial portions of the Software.

    THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
    IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
    FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
    AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
    LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
    OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
    SOFTWARE.

    

    Author:
    Lawrence McKay
    Lawrence@mckayit.com
    McKayIT Solutions Pty Ltd
    
    Date:    10th Dec 2018 
   

    ******* Update Version number below when a change is done.*******

    History
    Version         Date                Name           Detail
    ---------------------------------------------------------------------------------------
    0.0.1          10th Dec 2018       Lawrence       Initial Coding
    
   #>

    [CmdletBinding()]
    param 
    (
        [Parameter(Mandatory = $true)]
        $Batchname
    )
    BEGIN {
        $WindowTitle =$host.ui.rawui.WindowTitle  # getting Window title name
        $host.ui.rawui.WindowTitle = 'Geting the Onpremise Mailbox Sizes for a Batch '   # Setting window title
        $TotalMBSizes = 0
          }
    PROCESS {
        
        Try {
            Get-userfromValadationFile -batchname $batchname -outputVar fromlist

            foreach ($user in $fromlist){
                
                Write-host 'Processing ' $user -ForegroundColor green
                

                $mbs = Get-mailboxStatistics $user
                 $TotalMBSize =$mbs | Select @{n='TotalItemSize';e={[math]::Round(($_.TotalItemSize.ToString().Split("(")[1].Split(" ")[0].Replace(",","")/1MB),2)}}
            
                   $TotalMBSizes +=$($TotalMBSize.TotalItemSize)
            
                }

       }
            
        Catch {
            # Catches error from Try 
            Write-Host "ERROR : $($_.Exception.Message)" -ForegroundColor Magenta
        }
    }
    END {
        write-host "Total data to transfer is: $($TotalMBSizes/1024) (GB's)"
        write-host "From a total number of mailboxes: " $fromlist.count
        $host.ui.rawui.WindowTitle = $WindowTitle #Setting Window title bact to what it was.
    }
    
    
}

function Start-QhMailboxsizeO365_Parallel
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
		[ValidateRange(1, 20)]
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
            Import-Module QHSupport  -WarningAction SilentlyContinue

			
		}
		
		$ScriptBlock = {
			param (
				$users
			)
            Connect-exchangeonlineO365Clixml			
			get-QhMailboxsizeO365 -UserPrincipalName $users
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
		#Get-Job | Remove-Job -Force
	}
}


function get-QhMailboxsizeO365
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

				$mbxStats = Get-ExoMailboxStatistics $UPN
$tsize =$mbxStats.TotalItemSize.ToString().Split("(")[1].Split(" ")[0].Replace(",","")/1MB

				$prop = [ordered]@{
					UserPrincipalName = $UPN
					Mailbox		      = 'O365'
					TotalMBSize	  = $tsize
					MailboxTypeDetail = $mbxStats.MailboxTypeDetail
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

Function start-QHFailedmigrationbatches
{
    #... checks and restarts Migration batches that have an error while syncing
    <#

Syntax 
    start-QHFailedmigrationbatches

Notes

    Need connection to O365 Exchange



    *******************
    Copyright Notice.
    *******************
    Copyright (c) 2018   McKayIT Solutions Pty Ltd.

    Permission is hereby granted, free of charge, to any person obtaining a copy
    of this software and associated documentation files (the "Software"), to deal
    in the Software without restriction, including without limitation the rights
    to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
    copies of the Software, and to permit persons to whom the Software is
    furnished to do so, subject to the following conditions:

    The above copyright notice and this permission notice shall be included in all
    copies or substantial portions of the Software.

    THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
    IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
    FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
    AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
    LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
    OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
    SOFTWARE.

    

    Author:
    Lawrence McKay
    Lawrence@mckayit.com
    McKayIT Solutions Pty Ltd
    
    Date:    7 Nov 2018
   

    ******* Update Version number below when a change is done.*******

    History
    Version         Date                Name           Detail
    ---------------------------------------------------------------------------------------
    0.0.1           7 Nov 2018       Lawrence       Initial Coding
    
   #>


    BEGIN
    {
        $i = $null
    }

    process
    {

        try
        {
            do
            {
                #making sure Exchange on connection is kept open
                Connect-exchangeonlineO365Clixml |out-null
                clear-Xline 6


                Write-host 'checking to see if there are any failed users with MRS issues.' -ForegroundColor Cyan
        
                $failed = Get-EXOMigrationuser -ResultSize Unlimited  |Where-Object {$_.status -match 'fail' -and $_.ErrorSummary -match "switch the mailbox into Sync Source mode"}
                
                FOREACH ($user in $failed.identity)
                {
                    Write-host 'Fixing ' $user -ForegroundColor green
                    start-exomigrationuser $user 
                }
                $Othererror =Get-EXOMigrationuser -ResultSize Unlimited  |Where-Object {$_.status -match 'fail'} 
                if ($othererror -ne $null){
                    Write-host   'Here are other errors' -ForegroundColor Yellow
                    $othererror  |ft identity,batchid ,status ,errors*
                }
                
                $i = $i + 1
                $nextruntime = (Get-Date).AddSeconds(1800).ToString(" hh:mm")
                Write-host 'Waiting for ' $nextruntime  ' before it rechecks' -ForegroundColor Cyan
                Start-Sleep 1800
            }  until ($i -eq 1200)
         
          
        }
        Catch
        {
            # Catches error from Try 
            Write-Host "ERROR : $($_.Exception.Message)" -ForegroundColor Magenta
        }
    }  
        
        
}    

 
Function Resize-Console
{
<#  .Synopsis Resize PowerShell console window .Description Resize PowerShell console window. Make it bigger, smaller or increase / reduce the width and height by a specified number .Parameter -Bigger Increase the window's both width and height by 10. .Parameter -Smaller Reduce the window's both width and height by 10. .Parameter Width Resize the window's width by passing in an integer. .Parameter Height Resize the window's height by passing in an integer. .Example # Make the window bigger. Resize-Console -bigger  .Example # Make the window smaller. Resize-Console -smaller  .Example # Increase the width by 15. Resize-Console -Width 15  .Example # Reduce the Height by 10. Resize-Console -Height -10  .Example # Reduce the Width by 5 and Increase Height by 10. Resize-Console -Width -5 -Height 10 #>
 
[CmdletBinding()]
PARAM (
[Parameter(Mandatory=$false,HelpMessage="Increase Width and Height by 10")][Switch] $Bigger,
[Parameter(Mandatory=$false,HelpMessage="Reduce Width and Height by 10")][Switch] $Smaller,
[Parameter(Mandatory=$false,HelpMessage="Increase / Reduce Width" )][Int32] $Width,
[Parameter(Mandatory=$false,HelpMessage="Increase / Reduce Height" )][Int32] $Height
)
 
#Get Current Buffer Size and Window Size
$bufferSize = $Host.UI.RawUI.BufferSize
$WindowSize = $host.UI.RawUI.WindowSize
If ($Bigger -and $Smaller)
{
Write-Error "Please make up your mind, you can't go bigger and smaller at the same time!"
} else {
if ($Bigger)
{
$NewWindowWidth = $WindowSize.Width + 10
$NewWindowHeight = $WindowSize.Height + 10
 
#Buffer size cannot be smaller than Window size
If ($bufferSize.Width -lt $NewWindowWidth)
{
$bufferSize.Width = $NewWindowWidth
}
if ($bufferSize.Height -lt $NewWindowHeight)
{
$bufferSize.Height = $NewWindowHeight
}
$WindowSize.Width = $NewWindowWidth
$WindowSize.Height = $NewWindowHeight
 
} elseif ($Smaller)
{
$NewWindowWidth = $WindowSize.Width - 10
$NewWindowHeight = $WindowSize.Height - 10
$WindowSize.Width = $NewWindowWidth
$WindowSize.Height = $NewWindowHeight
}
 
if ($Width)
{
#Resize Width
$NewWindowWidth = $WindowSize.Width + $Width
If ($bufferSize.Width -lt $NewWindowWidth)
{
$bufferSize.Width = $NewWindowWidth
}
$WindowSize.Width = $NewWindowWidth
}
if ($Height)
{
#Resize Height
$NewWindowHeight = $WindowSize.Height + $Height
If ($bufferSize.Height -lt $NewWindowHeight)
{
$bufferSize.Height = $NewWindowHeight
}
$WindowSize.Height = $NewWindowHeight
 
}
#commit resize
$host.UI.RawUI.BufferSize = $buffersize
$host.UI.RawUI.WindowSize = $WindowSize
}
 
}


Function new-blankFunction 
{
    #... This Displays the Description when get-help1 is Run.
    <# 
    ********************************************************************************  
    *                                                                              *  
    *              This script IS a Blank Template                                 *
    *                                                                              *  
    ********************************************************************************    
    Note.
        Needs connection to What?? AD and MSOL with no Prefix.
        
    SYNTAX.
        <FunctionName -bulkUsers $users
        Compare-ImmutableID -bulkUsers $users |out-gridview
      
    *******************
    Copyright Notice.
    *******************
    Copyright (c) 2018   McKayIT Solutions Pty Ltd.

    Permission is hereby granted, free of charge, to any person obtaining a copy
    of this software and associated documentation files (the "Software"), to deal
    in the Software without restriction, including without limitation the rights
    to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
    copies of the Software, and to permit persons to whom the Software is
    furnished to do so, subject to the following conditions:

    The above copyright notice and this permission notice shall be included in all
    copies or substantial portions of the Software.

    THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
    IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
    FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
    AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
    LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
    OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
    SOFTWARE.

    

    Author:
    Lawrence McKay
    Lawrence@mckayit.com
    McKayIT Solutions Pty Ltd
    
    Date:    14 Sept 2018
   

    ******* Update Version number below when a change is done.*******

    History
    Version         Date                Name           Detail
    ---------------------------------------------------------------------------------------
    0.0.1           18 Sept 2018       Lawrence       Initial Coding
    
   #>

    [CmdletBinding()]
    param 
    (
        $BulkUsers
    )
    BEGIN {
        
          }
    PROCESS {
        
        Try {
				
        }
            
        Catch {
            # Catches error from Try 
            Write-Host "ERROR : $($_.Exception.Message)" -ForegroundColor Magenta
        }
    }
    END {
		
    }
}
