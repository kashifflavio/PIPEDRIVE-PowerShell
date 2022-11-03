$PipeDriveAPI = "your API Key"

$Pipedrive_domain = "company"

$Pipedrive_BaseURL = "https://$Pipedrive_domain.pipedrive.com/v1/"

$PipedrivePersonalMailHoster = @("hotmail.com","gmail.com","gmx.at","outlook.com","live.de","gmx.net","yahoo.de","me.com","online.de","googlemail.com","msn.com","gmx.net","yahoo.com","freemail.hu","web.de")


[string]$LogPath = "C:\Users\seimi\1 - TECHGUY\GitHub\PIPEDRIVE-PowerShell" #Path to store the Lofgile
[string]$LogfileName = "FindContactsWithMails" #FileName of the Logfile
[int]$DeleteAfterDays = 10 #Time Period in Days when older Files will be deleted


$MailContent = @()
$MailSender = "your Mail"


$clientID = "your Client ID"
$Clientsecret = "your Secret"
$tenantID = "your TenantID"

$UpdateCount=0



#endregion Parameters

#region Function
function Write-TechguyLog {
    [CmdletBinding()]
    param
    (
        [ValidateSet('DEBUG', 'INFO', 'WARNING', 'ERROR')]
        [string]$Type,
        [string]$Text
    )

    # Set logging path
    if (!(Test-Path -Path $logPath)) {
        try {
            $null = New-Item -Path $logPath -ItemType Directory
            Write-Verbose ("Path: ""{0}"" was created." -f $logPath)
        }
        catch {
            Write-Verbose ("Path: ""{0}"" couldn't be created." -f $logPath)
        }
    }
    else {
        Write-Verbose ("Path: ""{0}"" already exists." -f $logPath)
    }
    [string]$logFile = '{0}\{1}_{2}.log' -f $logPath, $(Get-Date -Format 'yyyyMMdd'), $LogfileName
    $logEntry = '{0}: <{1}> {2}' -f $(Get-Date -Format dd.MM.yyyy-HH:mm:ss), $Type, $Text
    Add-Content -Path $logFile -Value $logEntry
}
#endregion Function
Write-TechguyLog -Type INFO -Text "START SCRIPT"


Write-TechguyLog -Type INFO -Text "Get all Contacts from Pipedrive"

$start = 0
$Contacts = @()
do {

    $Result = Invoke-RestMethod -Uri "$($Pipedrive_BaseURL)persons?start=$start&limit=500&api_token=$PipeDriveAPI"
    $start = $start + 500
    $Contacts += $Result
    #$Contacts.data.count
}
while
(
    $Result.data.count -eq 500
)
#return $Contacts 

Write-TechguyLog -Type INFO -Text "Found this amount of Contacts: $($Contacts.data.count)"

foreach ($Con in $Contacts.data) {
    Write-TechguyLog -Type INFO -Text "Work with Contact ID: $($Con.id)"
    Write-TechguyLog -Type INFO -Text "Work with Contact name: $($Con.name)"


    if ($Con.email.count -gt 1) {
        Write-TechguyLog -Type INFO -Text "Contact has more than 1 mail: $($Con.email.count)"

        #"https://au2mator.pipedrive.com/person/$($Con.id)"

        $preJson = @()
        $IsTrue = $false
        $i=0
        foreach ($Email in $Con.email) {
            $i++
            Write-TechguyLog -Type INFO -Text "Work with that mail: $($email.value)"

            $Hoster = $email.value.split("@")[1].trim()

            Write-TechguyLog -Type INFO -Text "That is the Hoster: $Hoster"

            if ($PipedrivePersonalMailHoster -contains $Hoster) {
                Write-TechguyLog -Type INFO -Text "Mail is a personal Hoster Mail"
                $item = New-Object PSObject
                $item | Add-Member -type NoteProperty -Name 'label' -Value 'personal'
                $item | Add-Member -type NoteProperty -Name 'value' -Value $email.value
                if ($IsTrue -eq $false -and $i -eq $Con.email.count)
                {
                    Write-TechguyLog -Type INFO -Text "There is still not Primary, so mark the last entry as Primary"
                    $item | Add-Member -type NoteProperty -Name 'primary' -Value $true
                    $IsTrue = $true
                }
                else {
                    $item | Add-Member -type NoteProperty -Name 'primary' -Value $false
                }


            }
            else {
                Write-TechguyLog -Type INFO -Text "Mail is a business Mail"
                $item = New-Object PSObject
                $item | Add-Member -type NoteProperty -Name 'label' -Value 'work'
                $item | Add-Member -type NoteProperty -Name 'value' -Value $email.value

                If ($IsTrue) {
                    Write-TechguyLog -Type INFO -Text "IsTrue Value is true so set all to false"
                    Write-TechguyLog -Type INFO -Text "Mail Primary"
                    $item | Add-Member -type NoteProperty -Name 'primary' -Value $false
                }
                else {
                    Write-TechguyLog -Type INFO -Text "IsTrue else reached"
                    $item | Add-Member -type NoteProperty -Name 'primary' -Value $true
                    $IsTrue = $true
                }
            }

            $preJson += $item
        }



        $JsonDestString=$preJson | ConvertTo-Json -Compress
        $JsonOrigString= $Con.email | ConvertTo-Json -Compress
        Write-TechguyLog -Type INFO -Text "Thats our JSON: $JsonBody"
        if ($JsonDestString -ne $JsonOrigString)
        {
            Write-TechguyLog -Type INFO -Text "Values are different, so do update"
            $JsonBody = $preJson | ConvertTo-Json
            $BodyJsonTeam = @"
            {
               "email":$JsonBody
            }
"@

        Invoke-RestMethod -Uri "https://api.pipedrive.com/v1/persons/$($Con.id)?api_token=$PipeDrive_API" -Method PUT -Body $BodyJsonTeam  -ContentType 'application/json'
        $UpdateCount++
        $MailContent += "ID:  $($Con.id), Name: $($Con.name) <br>" 


        Write-TechguyLog -Type INFO -Text "Add a Note to Contact"

            $bodyNote = @{
                "person_id" = "$($Con.id)"
                "content" = "ReLabled Mail and Primary: $JsonDestString"
            }
            $bodyNoteJSON=$bodyNote | ConvertTo-Json
            Invoke-RestMethod -Uri "https://api.pipedrive.com/v1/notes?api_token=$PipeDrive_API" -Method POST -Body $bodyNoteJSON  -ContentType 'application/json'

        }



    }
    else {
        Write-TechguyLog -Type INFO -Text "Contact has only one Mail: $($Con.email.value)"
    }



}





if ($UpdateCount -gt 0)
{
    Write-TechguyLog -Type INFO -Text "Try to connect to Graph API"

    $tokenBody = @{
        Grant_Type    = "client_credentials"
        Scope         = "https://graph.microsoft.com/.default"
        Client_Id     = $clientId
        Client_Secret = $clientSecret
    }
    $tokenResponse = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$tenantID/oauth2/v2.0/token" -Method POST -Body $tokenBody
    $headers = @{
        "Authorization" = "Bearer $($tokenResponse.access_token)"
        "Content-type"  = "application/json"
    }
    Write-TechguyLog -Type INFO -Text "Connected to Graph API"

    Write-TechguyLog -Type INFO -Text "Send Mail with result"

    $URLsend = "https://graph.microsoft.com/v1.0/users/$MailSender/sendMail"

    $BodyJsonsend = @"
    {
        "message": {
          "subject": "Pipedrive mails updated related to Label and Primary",
          "body": {
            "contentType": "HTML",
            "content": "We have update that amount of Contacts: $UpdateCount <br>
            Here is the List: <br>
            $MailContent <br>
            
            "
          },
          "toRecipients": [
            {
              "emailAddress": {
                "address": "michael.seidl@au2mator.com"
              }
            }
          ]
        },
        "saveToSentItems": "false"
      }
"@

    Invoke-RestMethod -Method POST -Uri $URLsend -Headers $headers -Body $BodyJsonsend


}





#Clean Logs
Write-TechguyLog -Type INFO -Text "Clean Log Files"
$limit = (Get-Date).AddDays(-$DeleteAfterDays)
Get-ChildItem -Path $LogPath -Filter "*$LogfileName.log" | Where-Object { !$_.PSIsContainer -and $_.CreationTime -lt $limit } | Remove-Item -Force
Write-TechguyLog -Type INFO -Text "END Script"



