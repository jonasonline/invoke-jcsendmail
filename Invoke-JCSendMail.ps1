[CmdletBinding()]
Param(
    [Parameter(Mandatory = $True)]
    [string]$MailServer,

    [Parameter(Mandatory = $False)]
    [string]$EmailSmtpUser,

    [Parameter(Mandatory = $False)]
    [string]$EmailSmtpPass,
	
    [Parameter(Mandatory = $True)]
    [string]$EmailAdressFrom,

    [Parameter(Mandatory = $True)]
    [string[]]$EmailAdressesTo,

    [Parameter(Mandatory = $False)]
    [string[]]$ReplyToEmailAdresses,

    [Parameter(Mandatory = $True)]
    [string]$Subject,

    [Parameter(Mandatory = $True)]
    [string]$InputFile,

    [Parameter(Mandatory = $False)]
    [int]$Delay,

    [Parameter(Mandatory = $False)]
    [switch]$UseTLS,

    [Parameter(Mandatory = $False)]
    [switch]$IsHTML
)

function Invoke-JCSendMail {
    param(
        [CmdletBinding()]
        [Parameter(Mandatory = $True)]
        [string]$MailServer,

        [Parameter(Mandatory = $False)]
        [string]$EmailSmtpUser,

        [Parameter(Mandatory = $False)]
        [string]$EmailSmtpPass,
	
        [Parameter(Mandatory = $True)]
        [string]$EmailAdressFrom,

        [Parameter(Mandatory = $True)]
        [string]$EmailAdressTo,

        [Parameter(Mandatory = $False)]
        [string[]]$ReplyToEmailAdresses,

        [Parameter(Mandatory = $True)]
        [string]$Subject,

        [Parameter(Mandatory = $True)]
        [string]$InputFile,

        [Parameter(Mandatory = $False)]
        [bool]$UseTLS,

        [Parameter(Mandatory = $False)]
        [bool]$IsHTML
    )

    $SmtpClient = New-object system.net.mail.smtpClient 
    $MailMessage = New-Object system.net.mail.mailmessage
    if ($EmailSmtpUser -and $EmailSmtpPass) {
        $SMTPClient.Credentials = New-Object System.Net.NetworkCredential($EmailSmtpUser , $EmailSmtpPass );    
    }

    $SmtpClient.Host = $MailServer
    $mailmessage.From = $EmailAdressFrom
    if ($ReplyToEmailAdresses) {
        foreach ($emailAdress in $ReplyToEmailAdresses) {
            $mailmessage.ReplyToList.add($emailAdress)    
        }
    }

    if ($UseTLS) {
        $SmtpClient.Port = 587
    }
    
    $mailmessage.To.add($EmailAdressTo)
    if ($IsHTML) {
        $mailmessage.IsBodyHTML = $true
    }
    $mailmessage.Subject = $Subject 

    $message = Get-Content -Path $InputFile

    $mailmessage.Body += $message
    $Smtpclient.Send($mailmessage)

}

foreach ($emailAdress in $EmailAdressesTo) {

    Write-Host "Sending mail to $emailAdress" -ForegroundColor Green
    Invoke-JCSendMail -MailServer $MailServer -EmailSmtpUser $EmailSmtpUser -EmailSmtpPass $EmailSmtpPass -EmailAdressFrom $EmailAdressFrom -EmailAdressTo $emailAdress -ReplyToEmailAdresses  $ReplyToEmailAdresses -Subject $Subject -InputFile $InputFile -UseTLS $UseTLS -IsHTML $IsHTML
    if ($Delay) {
        Start-Sleep -Seconds $Delay
    }
}



