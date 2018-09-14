function PresentExchangeOptions {
    Write-Host "--------------------------------"
    Write-Host "Exchange Options"
    Write-Host "1. View Mailbox"
    Write-Host "2. Edit Mailbox"
    Write-Host "3. View Quarantine"
    Write-Host "4. Release Quarantine"
    Write-Host "5. Exit"

    $option = Read-Host "What would you like to do?"

    if($option -eq "1") {
        $user = Read-Host "Office365 Email"
        Write-Host "=================MAILBOX=================="
        Get-Mailbox -Identity $user | fl
        Write-Host "=================STATISTICS=================="
        $stats = Get-MailboxStatistics $user
        $stats | fl
        Write-Host "======SIZE======"
        $stats | ft displayname, totalitemsize
    }
    if($option -eq "2") {
        Write-Host "Not implimented"
    }
    if($option -eq "3") {
        $user = Read-Host "Office365 Email"
        $startdate = Read-Host "Start Date (01/01/0001)"
        $enddate = Read-Host "End Date (01/01/0001)"
        Get-QuarantineMessage -RecipientAddress $user -StartReceivedDate $startdate -EndReceivedDate $enddate | fl
    }
    if($option -eq "4") {
        $messageid = Read-Host "Quarantine MessageId"
        Write-Host "Releasing quarantined message..."
        Release-QuarantineMessage -Identity $messageid
    }
    if($option -ne "5") {
        PresentExchangeOptions
    }
}

function EndPSSession {
    Remove-PSSession $Session
}

$run = $true
$Session

while($run) {
    Write-Host "--------------------------------"
    Write-Host "Gary's Office365 Exchange Powerscript"
    Write-Host "--------------------------------"
    Write-Host ""
    Write-Host "Options"
    Write-Host "1. Connect To Server"
    Write-Host "2. Quit"

    $option = Read-Host "What would you like to do? "

    if($option -eq "1") {
        $adminEmail = Read-Host "Office365 Admin"
        $adminPassword = Read-Host "Password" -AsSecureString
        $credentials = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $adminEmail, $adminPassword
        $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $credentials -Authentication Basic -AllowRedirection
        try {
            Import-PSSession $Session -DisableNameChecking
        }
        catch{
            "Error: "
            Write-Host $_
            EndPSSession
        }

        PresentExchangeOptions
    }
    if($option -eq "2") {
        $run = $false
    }

    EndPSSession
}