function PresentExchangeOptions {
    Write-Host "--------------------------------"
    Write-Host "Exchange Options"
    Write-Host "1. View Mailbox"
    Write-Host "2. Edit Mailbox"
    Write-Host "3. View Quarantine"
    Write-Host "4. Release Quarantine"
    Write-Host "5. Run My Command"
    Write-Host "6. Import Commands"
    Write-Host "7. Exit"

    $option = Read-Host "What would you like to do?"

    switch($option) {
        "1" {
            $user = Read-Host "Office365 Email"
            Write-Host "=================MAILBOX=================="
            Get-Mailbox -Identity $user | Format-List
            Write-Host "=================STATISTICS=================="
            $stats = Get-MailboxStatistics $user
            $stats | Format-List
            Write-Host "======SIZE======"
            $stats | Format-Table displayname, totalitemsize
        }
        "2" {
            Write-Host "Not implimented..."
        }
        "3" {
            $user = Read-Host "Office365 Email"
            $startdate = Read-Host "Start Date (01/01/0001)"
            $enddate = Read-Host "End Date (01/01/0001)"
            Get-QuarantineMessage -RecipientAddress $user -StartReceivedDate $startdate -EndReceivedDate $enddate | Format-List
        }
        "4" {
            $messageid = Read-Host "Quarantine MessageId"
            Write-Host "Releasing quarantined message..."
            Release-QuarantineMessage -Identity $messageid
        }
        "5" {
            Write-Host "Check Here For Commands: https://docs.microsoft.com/en-us/powershell/module/exchange/mailboxes/get-mailbox?view=exchange-ps"
            $command = Read-Host "Command"
            Write-Host "==========COMMAND RESULTS============"
            try {
                Invoke-Expression $command
            }
            catch {
                Write-Host "Error invoking your command..."
                $_
            }
        }
        "6" {
            $file = Read-Host "Command File Location"
            $commands = [IO.File]::ReadAllText($file)
            Write-Host "===========IMPORTED COMMANDS============"
            $commands
            Write-Host "===========COMMAND RESULTS=============="
            try {
                Invoke-Expression $commands
            }
            catch {
                Write-Host "Error invoking commands..."
                $_
            }
        }
    }

    if($option -ne "7") {
        PresentExchangeOptions
    }
}

function EndPSSession {
    if($Session) {
        Remove-PSSession $Session
    }
}

$run = $true
$Session

while($run) {
    Write-Host "--------------------------------"
    Write-Host "Gary's Office365 Exchange Powerscript"
    Write-Host "--------------------------------"
    Write-Host ""
    Write-Host "Options"
    Write-Host "1. Connect To Office365 Exchange"
    Write-Host "2. Quit"

    $option = Read-Host "What would you like to do? "

    if($option -eq "1") {
        $adminEmail = Read-Host "Office365 Admin"
        $adminPassword = Read-Host "Password" -AsSecureString
        try {
            $credentials = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $adminEmail, $adminPassword
            $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $credentials -Authentication Basic -AllowRedirection -ErrorAction Stop

            Import-PSSession $Session -DisableNameChecking
            PresentExchangeOptions
        }
        catch {
            Write-Host "====================================="
            Write-Host "Error: " $_
            Write-Host "====================================="
        }

        EndPSSession
    }
    if($option -eq "2") {
        $run = $false
    }

    EndPSSession
}