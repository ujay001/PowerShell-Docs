# Define parameters
$csvFilePath = "SalesforceData.csv"
$logFilePath = "SalesforceDataExtraction.log"
$teamsWebhookUrl = "https://unfi.webhook.office.com/webhookb2/3496e607-6a03-4a69-8986-857ab2609aef@59f75f15-85cb-43ae-9985-f4c2e2f846cb/IncomingWebhook/ddaf77b9ca73460c8db3f3ef12dfe7fb/a5a8a627-de31-4ff2-8b82-f02a1545f1e1"

# Salesforce credentials
$sfInstanceUrl = "https://myunfi--uat.sandbox.my.salesforce.com"
$sfUsername = "ulysses.jayakar@unfi.com.mybiz.uat"
$sfPassword = "Sadmin@12345"
$sfSecurityToken = ""
$sfClientId = "3MVG9t9ADM8gwXaDkcJ_Obi18If5VNgowZw270FjyWL30XsUeZRcqEKG4Pf3A.NKhDn1yIXOlu6dGrHz14zHC"
$sfClientSecret = "CF5344FCFD8A1A6D6F373FD62E690B7BA6E673E085CD86140EBC2D8692E80D75"

$salesforceInstanceUrl = "https://myunfi--uat.sandbox.my.salesforce.com"
$accessToken = "YOUR_SALESFORCE_ACCESS_TOKEN"

# Define SFTP parameters
$sftpHost = "ft.unfi.com"
$sftpPort = 2022
$sftpUser = "UnifiedGrocersCID01"
$sftpPassword = "di5p&sT_S3ePH*bRaJ#s"
$sftpRemotePath = "/Outbound"

# Define status variable
$status = "Failed"

# Define archive parameters
$archiveDir = "Archive"
$timestamp = (Get-Date).ToString("yyyyMMddHHmmss")
$archiveFilePath = "$archiveDir\SalesforceData_$timestamp.zip"

# Logging function
function Log-Message {
    param (
        [string]$message
    )
    $timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    $logMessage = "$timestamp - $message"
    Add-Content -Path $logFilePath -Value $logMessage
    Write-Output $logMessage
}

# Send Teams notification function
function Send-TeamsNotification {
    param (
        [string]$status,
        [string]$message
    )
    $payload = @{
        title = "Salesforce Data Extraction Status"
        text = "$status- $message"
    }
    $payloadJson = $payload | ConvertTo-Json

    try {
        Invoke-RestMethod -Uri $teamsWebhookUrl -Method Post -Body $payloadJson -ContentType "application/json"
    } catch {
        Log-Message "Failed to send Teams notification: $_"
    }
}


function Get-SalesforceAccessToken {
    param (
        [string]$instanceUrl,
        [string]$username,
        [string]$password,
        [string]$securityToken,
        [string]$clientId,
        [string]$clientSecret
    )

    $body = @{
        grant_type    = "password"
        client_id     = $clientId
        client_secret = $clientSecret
        username      = $username
        password      = "$password$securityToken"
    }

    try {
        $response = Invoke-RestMethod -Uri "$instanceUrl/services/oauth2/token" -Method Post -Body $body
        return $response.access_token
    } catch {
        Log-Message "Failed to get Salesforce access token: $_" -logFile $logFilePath
        throw
    }
}

# Extract data from Salesforce
function Extract-SalesforceData {
    try {
        
        # Calculate dates for last week
        $today = Get-Date
        $startDate = ($today.AddDays(-7)).Date
        $endDate = $today.Date

        $query = "SELECT Id, CaseNumber, Subject, Status, CreatedDate FROM Case WHERE CreatedDate >= $($startDate.ToString('yyyy-MM-ddTHH:mm:ssZ')) AND CreatedDate <= $($endDate.ToString('yyyy-MM-ddTHH:mm:ssZ'))"
        $requestUrl = "$salesforceInstanceUrl/services/data/v54.0/query/?q=$([uri]::EscapeDataString($query))"
        $response = Invoke-RestMethod -Uri $requestUrl -Headers @{Authorization = "Bearer $accessToken"} -Method Get

        if ($response.records.Count -eq 0) {
            Log-Message "No cases found to export."
            $global:status = "Completed"
            return
        }

        # Export data to CSV
        $response.records | Select-Object Id, CaseNumber, Subject, Status, CreatedDate |
            Export-Csv -Path $csvFilePath -NoTypeInformation

        # Modify CSV header
        $csvContent = Get-Content $csvFilePath
        $csvContent[0] = "CaseID,CaseNumber,Subject,Status,CreationDate" # Modify as needed
        $csvContent | Set-Content $csvFilePath

        Log-Message "Data extraction completed successfully."
        $global:status = "Completed"

    } catch {
        Log-Message "Error during data extraction: $_"
        $global:status = "Failed"
    }
}

# Upload CSV to SFTP
function Upload-ToSFTP {
    if ($status -eq "Completed") {
        $session = $null
        try {
            Import-Module Posh-SSH

            # Establish the SFTP session
            $credential = New-Object PSCredential ($sftpUser, (ConvertTo-SecureString $sftpPassword -AsPlainText -Force))
            $session = New-SFTPSession -ComputerName $sftpHost -Port $sftpPort -Credential $credential

            # Upload the file
            Set-SFTPFile -SessionId $session.SessionId -LocalFile $csvFilePath -RemotePath $sftpRemotePath

            Log-Message "CSV file uploaded to SFTP server successfully."
        } catch {
            Log-Message "Failed to upload CSV file to SFTP server: $_"
            $global:status = "Failed"
        } finally {
            if ($session) {
                Remove-SFTPSession -SessionId $session.SessionId
                Log-Message "SFTP session closed."
            }
        }
    }
}

# Archive exported data and logs
function Archive-Files {
    try {
        if (Test-Path $csvFilePath -or Test-Path $logFilePath) {
            # Create a ZIP file
            Compress-Archive -Path $csvFilePath, $logFilePath -DestinationPath $archiveFilePath -Update
            Log-Message "Files archived successfully to $archiveFilePath."
        }
    } catch {
        Log-Message "Failed to archive files: $_"
    }
}

# Main script execution
try {
    $accessToken = Get-SalesforceAccessToken -instanceUrl $sfInstanceUrl -username $sfUsername -password $sfPassword -securityToken $sfSecurityToken -clientId $sfClientId -clientSecret $sfClientSecret

    Extract-SalesforceData
    Upload-ToSFTP
} finally {
    # Archive files and send final notification
    #Archive-Files

    if ($status -eq "Completed") {
        Send-TeamsNotification -status "Completed" -message "Data extraction and upload completed successfully."
        exit 0
    } else {
        Send-TeamsNotification -status "Failed" -message "Error occurred during data extraction or upload. Please check the logs."
        exit 1
    }
}
