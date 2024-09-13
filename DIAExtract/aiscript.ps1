# Import required modules
Import-Module Posh-SSH

# File paths and variables
$csvFilePath = "output.csv"
$logFilePath = "logfile.log"
$sftpHost = "ft.unfi.com"
$sftpPort = 2022
$sftpUser = "your-username"
$sftpPass = "your-password"
$remoteFilePath = "/remote/path/output.csv"
$teamsWebhookUrl = "https://outlook.office.com/webhook/YOUR-TEAMS-WEBHOOK-URL"
$jobStartTime = Get-Date
$errorMessage = $null
$retryCount = 0

# Function to log errors
function Log-Error {
    param ($message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Add-Content -Path $logFilePath -Value "$timestamp - ERROR: $message"
}

# Function to notify Teams with job details
function Notify-Teams {
    param ($status, $lineCount, $elapsedTime, $errorMessage, $retryCount)

    $message = @{
        title = "Salesforce Data Extract Job: $status"
        text = "Job Details:
        - Status: $status
        - Elapsed Time: $elapsedTime
        - CSV Line Count: $lineCount
        - Error: $errorMessage
        - Retry Count: $retryCount"
    }

    $jsonPayload = $message | ConvertTo-Json
    Invoke-RestMethod -Uri $teamsWebhookUrl -Method Post -Body $jsonPayload -ContentType 'application/json'
}

# Step 1: Fetch data from Salesforce using SFDX CLI
function Get-SalesforceData {
    try {
        $query = "SELECT Id, CaseNumber, Subject FROM Case"
        $sfdxCommand = "sfdx force:data:soql:query -q '$query' -r csv -u your-sfdx-org-alias"

        # Execute the SFDX command and save to CSV
        Invoke-Expression "$sfdxCommand | Out-File -FilePath $csvFilePath -Encoding utf8"

        if (!(Test-Path $csvFilePath)) {
            throw "Failed to create CSV file from Salesforce data"
        }
    } catch {
        Log-Error "Failed to fetch Salesforce data: $_"
        throw $_
    }
}

# Step 2: Modify CSV header
function Modify-CsvHeader {
    try {
        if (!(Test-Path $csvFilePath)) {
            throw "CSV file does not exist"
        }

        $csvContent = Get-Content -Path $csvFilePath
        $newHeader = "CaseId,TicketNumber,Subject"
        $csvContent[0] = $newHeader
        $csvContent | Set-Content -Path $csvFilePath
    } catch {
        Log-Error "Failed to modify CSV header: $_"
        throw $_
    }
}

# Step 3: Upload file to SFTP using Posh-SSH with error handling for network and permission issues
function Upload-ToSftp {
    try {
        $sftpSession = New-SFTPSession -ComputerName $sftpHost -Port $sftpPort -Credential (New-Object System.Management.Automation.PSCredential ($sftpUser, (ConvertTo-SecureString $sftpPass -AsPlainText -Force))) -ErrorAction Stop
        
        # Check if the file already exists on the remote server
        $existingFile = Get-SFTPItem -SessionId $sftpSession.SessionId -Path $remoteFilePath -ErrorAction SilentlyContinue
        
        if ($existingFile) {
            # If the file exists, remove it before uploading the new one
            Remove-SFTPItem -SessionId $sftpSession.SessionId -Path $remoteFilePath -ErrorAction Stop
        }
        
        # Upload the file
        Set-SFTPFile -SessionId $sftpSession.SessionId -LocalFile $csvFilePath -RemotePath $remoteFilePath -ErrorAction Stop
        
        # Close the SFTP session
        Remove-SFTPSession -SessionId $sftpSession.SessionId
    } catch [System.Net.Sockets.SocketException] {
        Log-Error "Network issue: $_"
        throw "Network issue occurred while uploading file"
    } catch [PoshSSH.SFTPPermissionDeniedException] {
        Log-Error "Permission issue: $_"
        throw "Permission denied during file upload"
    } catch {
        Log-Error "Failed to upload file to SFTP: $_"
        throw $_  # Rethrow to be caught by main try-catch block
    }
}

# Retry Logic with Exponential Backoff
function Retry-Operation {
    param (
        [ScriptBlock]$operation,
        [int]$maxRetries,
        [int]$initialDelay
    )

    $attempt = 0
    $delay = $initialDelay

    while ($attempt -lt $maxRetries) {
        try {
            # Run the operation
            &$operation
            return  # Exit if successful
        } catch {
            Log-Error "Attempt $($attempt + 1) failed: $_"
            $attempt++
            $retryCount = $attempt  # Update retry count
            if ($attempt -ge $maxRetries) {
                throw "Maximum retry attempts ($maxRetries) reached."
            }
            Start-Sleep -Seconds $delay
            $delay = $delay * 2  # Exponential backoff
        }
    }
}

# Main job logic with edge case handling for retries
try {
    Retry-Operation -operation {
        Get-SalesforceData
        Modify-CsvHeader
        Upload-ToSftp
    } -maxRetries 5 -initialDelay 3600
} catch {
    $errorMessage = $_
    Log-Error "Script failed after multiple retries: $_"
    # Exit with failure status for Task Scheduler
    exit 1
}

# Calculate elapsed time
$elapsedTime = (Get-Date) - $jobStartTime

# Count lines in the CSV file
$lineCount = (Get-Content -Path $csvFilePath | Measure-Object -Line).Lines

# Send notification to Teams
$jobStatus = if ($errorMessage) { "Failed" } else { "Succeeded" }
Notify-Teams -status $jobStatus -lineCount $lineCount -elapsedTime $elapsedTime -errorMessage $errorMessage -retryCount $retryCount

# Exit with success status
exit 0
