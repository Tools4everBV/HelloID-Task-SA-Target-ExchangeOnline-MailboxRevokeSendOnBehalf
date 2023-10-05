# HelloID-Task-SA-Target-ExchangeOnline-MailboxRevokeSendOnBehalf
#################################################################
# Form mapping
$formObject = @{
    MailboxIdentity = $form.MailboxIdentity
    UsersToRemove   = $form.UsersToRemove.id
}

try {
    Write-Information "Executing ExchangeOnline action: [MailboxRevokeSendOnBehalf] for: [$($formObject.MailboxIdentity)]"

    $null = Import-Module ExchangeOnlineManagement

    $securePassword = ConvertTo-SecureString $ExchangeOnlineAdminPassword -AsPlainText -Force
    $credential = [System.Management.Automation.PSCredential]::new($ExchangeOnlineAdminUsername, $securePassword)
    $null = Connect-ExchangeOnline -Credential $credential -ShowBanner:$false -ShowProgress:$false -ErrorAction Stop -Verbose:$false -CommandName 'Set-Mailbox', 'Disconnect-ExchangeOnline'
    $IsConnected = $true

    foreach ($user in $formObject.UsersToRemove) {
        $null = Set-Mailbox -Identity $formObject.MailboxIdentity -GrantSendOnBehalfTo @{remove = "$user" } -Confirm:$false -ErrorAction Stop

        $auditLog = @{
            Action            = 'UpdateResource'
            System            = 'ExchangeOnline'
            TargetIdentifier  = $formObject.MailboxIdentity
            TargetDisplayName = $formObject.MailboxIdentity
            Message           = "ExchangeOnline action: [MailboxRevokeSendOnBehalf] Revoke [$($user)] from mailbox [$($formObject.MailboxIdentity)] executed successfully"
            IsError           = $false
        }
        Write-Information -Tags 'Audit' -MessageData $auditLog
        Write-Information "ExchangeOnline action: [MailboxRevokeSendOnBehalf] Revoke [$($user)] from mailbox [$($formObject.MailboxIdentity)] executed successfully"
    }
} catch {
    $ex = $_
    $auditLog = @{
        Action            = 'UpdateResource'
        System            = 'ExchangeOnline'
        TargetIdentifier  = $formObject.MailboxIdentity
        TargetDisplayName = $formObject.MailboxIdentity
        Message           = "Could not execute ExchangeOnline action: [MailboxRevokeSendOnBehalf] for: [$($formObject.MailboxIdentity)], error: $($ex.Exception.Message)"
        IsError           = $true
    }
    Write-Information -Tags 'Audit' -MessageData $auditLog
    Write-Error "Could not execute ExchangeOnline action: [MailboxRevokeSendOnBehalf] for: [$($formObject.MailboxIdentity)], error: $($ex.Exception.Message)"
} finally {
    if ($IsConnected) {
        $null = Disconnect-ExchangeOnline -Confirm:$false -Verbose:$false
    }
}
#################################################################
