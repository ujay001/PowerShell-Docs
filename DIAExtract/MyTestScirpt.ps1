# Paths
$logFilePath = "logfile.log"
$csvFilePath = "file.csv"
$cleanedCsvFilePath = "cleaned_file.csv"
$modifiedCsvFilePath = "modified_file.csv"

# Salesforce credentials
$sfInstanceUrl = "https://myunfi--uat.sandbox.my.salesforce.com"
$sfUsername = "ulysses.jayakar@unfi.com.mybiz.uat"
$sfPassword = "Sadmin@12345"
$sfSecurityToken = ""
$sfClientId = "3MVG9t9ADM8gwXaDkcJ_Obi18If5VNgowZw270FjyWL30XsUeZRcqEKG4Pf3A.NKhDn1yIXOlu6dGrHz14zHC"
$sfClientSecret = "CF5344FCFD8A1A6D6F373FD62E690B7BA6E673E085CD86140EBC2D8692E80D75"

# FTP server credentials
$ftpServer = "ft.unfi.com"
$ftpUsername = "UnifiedGrocersCID01"
$ftpPassword = "di5p&sT_S3ePH*bRaJ#s"
$ftpPath = "/Outbound"
$port = 2022

# Teams webhook URL
$teamsWebhookUrl = "https://unfi.webhook.office.com/webhookb2/3496e607-6a03-4a69-8986-857ab2609aef@59f75f15-85cb-43ae-9985-f4c2e2f846cb/IncomingWebhook/ddaf77b9ca73460c8db3f3ef12dfe7fb/a5a8a627-de31-4ff2-8b82-f02a1545f1e1"

function Log-Message {
    param (
        [string]$message,
        [string]$logFile
    )
    
    $timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    "$timestamp - $message" | Out-File -Append -FilePath $logFile
}

function Send-TeamsNotification {
    param (
        [string]$message
    )

    $payload = @{
        text = $message
    }

    try {
        $response = Invoke-RestMethod -Uri $teamsWebhookUrl -Method Post -Body ($payload | ConvertTo-Json) -ContentType "application/json"
        Log-Message "Teams notification sent: $message" -logFile $logFilePath
    } catch {
        Log-Message "Failed to send Teams notification: $_" -logFile $logFilePath
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

function Export-DataFromSalesforce {
    param (
        [string]$accessToken,
        [string]$instanceUrl,
        [string]$query,
        [string]$filePath
    )

    $headers = @{
        Authorization = "Bearer $accessToken"
    }
    
    try {
        $response = Invoke-RestMethod -Uri "$instanceUrl/services/data/v56.0/query?q=$($query)" -Method Get -Headers $headers
        $data = $response.records
        $data | ConvertTo-Csv -NoTypeInformation | Out-File -FilePath $filePath -Encoding UTF8
        Log-Message "Data exported to CSV successfully." -logFile $logFilePath
    } catch {
        Log-Message "Failed to export data from Salesforce: $_" -logFile $logFilePath
        throw
    }
}

function Drop-ColumnsFromCsv {
    param (
        [string]$inputFilePath,
        [string]$outputFilePath,
        [string[]]$columnsToDrop
    )

    try {
        $csv = Import-Csv -Path $inputFilePath

        # Drop specified columns
        $csv | Select-Object -Property * -ExcludeProperty $columnsToDrop | Export-Csv -Path $outputFilePath -NoTypeInformation -Encoding UTF8
        Log-Message "Columns dropped and CSV file cleaned successfully." -logFile $logFilePath
    } catch {
        Log-Message "Failed to drop columns from CSV: $_" -logFile $logFilePath
        throw
    }
}

function Modify-CsvHeaders {
    param (
        [string]$inputFilePath,
        [string]$outputFilePath,
        [hashtable]$headerMapping
    )

    try {
        $csvContent = Get-Content -Path $inputFilePath
        $header = $csvContent[0] -replace '(?<=^)(.*?)(?=,)', { $headerMapping[$matches[0]] }
        $csvContent[0] = $header
        $csvContent | Set-Content -Path $outputFilePath -Encoding UTF8
        Log-Message "CSV headers modified successfully." -logFile $logFilePath
    } catch {
        Log-Message "Failed to modify CSV headers: $_" -logFile $logFilePath
        throw
    }
}

function Upload-FileToFtp {
    param (
        [string]$ftpUrl,
        [string]$filePath,
        [string]$ftpUsername,
        [string]$ftpPassword
    )

    try {
        $fileInfo = Get-Item $filePath
        $ftpRequest = [System.Net.FtpWebRequest]::Create($ftpUrl)
        $ftpRequest.Method = [System.Net.WebRequestMethods+Ftp]::UploadFile
        $ftpRequest.Credentials = New-Object System.Net.NetworkCredential($ftpUsername, $ftpPassword)
        $ftpRequest.UseBinary = $true
        $ftpRequest.UsePassive = $true
        $ftpRequest.EnableSsl = $false

        $fileStream = $fileInfo.OpenRead()
        $ftpStream = $ftpRequest.GetRequestStream()
        $buffer = New-Object byte[] 1024
        $bytesRead = 0

        while (($bytesRead = $fileStream.Read($buffer, 0, $buffer.Length)) -gt 0) {
            $ftpStream.Write($buffer, 0, $bytesRead)
        }

        $fileStream.Close()
        $ftpStream.Close()
        Log-Message "Upload successful." -logFile $logFilePath
    } catch {
        Log-Message "Failed to upload file to FTP: $_" -logFile $logFilePath
        throw
    }
}

# Main script
try {
    $accessToken = Get-SalesforceAccessToken -instanceUrl $sfInstanceUrl -username $sfUsername -password $sfPassword -securityToken $sfSecurityToken -clientId $sfClientId -clientSecret $sfClientSecret

    # Query to get cases created in the last week
    $salesforceQuery = "SELECT Id, CaseNumber, Subject, Status, CreatedDate FROM Case WHERE CreatedDate = last_week"

    Export-DataFromSalesforce -accessToken $accessToken -instanceUrl $sfInstanceUrl -query $salesforceQuery -filePath $csvFilePath

    # Define columns to drop
    $columnsToDrop = @("attributes")

    # Drop unwanted columns
    Drop-ColumnsFromCsv -inputFilePath $csvFilePath -outputFilePath $cleanedCsvFilePath -columnsToDrop $columnsToDrop

    $headerMapping = @{
        "Id"          = "Case_ID"
        "CaseNumber"  = "Case_Number"
        "Subject"     = "Case_Subject"
        "Status"      = "Case_Status"
        "CreatedDate" = "Case_CreatedDate"
    }

    # Modify CSV headers
    Modify-CsvHeaders -inputFilePath $cleanedCsvFilePath -outputFilePath $modifiedCsvFilePath -headerMapping $headerMapping

    # Upload modified CSV file to FTP server
    Upload-FileToFtp -ftpUrl $ftpServer -filePath $modifiedCsvFilePath -ftpUsername $ftpUsername -ftpPassword $ftpPassword

    # Notify Teams
    Send-TeamsNotification -message "Process completed successfully."

    Log-Message "Process completed successfully." -logFile $logFilePath
} catch {
    Send-TeamsNotification -message "Script failed: $_"
    Log-Message "Script failed: $_" -logFile $logFilePath
    exit 1
}

exit 0
