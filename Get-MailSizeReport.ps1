<#
.SYNOPSIS
    Audits Microsoft 365 mailbox message sizes for all users in a tenant.
.DESCRIPTION
    This script connects to Microsoft Graph using app-only authentication (client secret)
    to gather statistics on the size of messages in each user's Inbox and Sent Items folders.
    It is designed for large environments and includes resilience features to allow for
    restarting the script without losing progress.
    The script generates a CSV file to track progress and store the final results. Upon starting,
    it checks for this CSV to resume where it left off.
    Core Features:
    - PowerShell 5.1 and ISE compatible.
    - Uses native Invoke-RestMethod to minimize dependencies.
    - Idempotent design using a CSV for state management.
    - Granular error handling to skip problematic mailboxes.
    - Optimized API calls to minimize data transfer and avoid throttling.
    - Dynamic throttling delay with exponential backoff and decay.
    - Filters messages by age (e.g., last 180 days).
.PARAMETER TenantId
    The Microsoft Entra ID (Tenant) ID for your organization.
.PARAMETER ClientId
    The Application (Client) ID of your Entra ID registered application.
.PARAMETER ClientSecret
    The client secret for your Entra ID registered application.
    Note: For production use, avoid passing this as a plain-text parameter.
    Consider using Read-Host -AsSecureString or retrieving it from a secure vault.
.PARAMETER CsvPath
    The full path for the output/progress CSV file (e.g., 'C:\Temp\MailboxAudit.csv').
.PARAMETER MaxMessageAgeInDays
    Only analyze messages received or sent within this many days (default: 180).
.EXAMPLE
   $secret = Read-Host "Enter client secret" -AsSecureString
   .\MailSizeReport.ps1 -TenantId 'xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx' `
                        -ClientId 'xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx' `
                        -ClientSecret $secret `
                        -CsvPath 'C:\Temp\MailboxAudit.csv' `
                        -MaxMessageAgeInDays 90
.PERMISSIONS_REQUIRED
    Microsoft Graph Application Permissions:
    - User.Read.All: Required to get the list of all users in the tenant.
    - Mail.ReadBasic.All: Required to read basic message properties, including the
      extended property for message size, from all mailboxes.
.NOTES
    Author:  Tydorius
    Version: 1.1
    Date:    July 28, 2025
    License: CC BY-NC 4.0
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$TenantId,

    [Parameter(Mandatory = $true)]
    [string]$ClientId,

    [Parameter(Mandatory = $true)]
    [Security.SecureString]$ClientSecret,

    [Parameter(Mandatory = $true)]
    [string]$CsvPath,

    [Parameter(Mandatory = $false)]
    [int]$MaxMessageAgeInDays = 180
)

#region Helper Functions

# Determine log file path based on CSV path
$script:LogPath = $CsvPath -replace '\.csv$', "_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"

function Write-Log {
    param(
        [string]$Message,
        [string]$Level = 'INFO'
    )
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $logLine = "[$timestamp][$Level] $Message"
    Write-Host $logLine
    # Append to log file
    Add-Content -Path $script:LogPath -Value $logLine
}

function Get-GraphToken {
    param(
        [string]$TenantId,
        [string]$ClientId,
        [Security.SecureString]$ClientSecret
    )
    Write-Log "Requesting new authentication token..."
    
    # Convert SecureString to plain text (in memory only)
    $bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($ClientSecret)
    try {
        $plainClientSecret = [Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr)
    }
    finally {
        [Runtime.InteropServices.Marshal]::FreeBSTR($bstr)
    }

    $tokenEndpoint = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
    $body = @{
        grant_type    = 'client_credentials'
        scope         = 'https://graph.microsoft.com/.default'
        client_id     = $ClientId
        client_secret = $plainClientSecret
    }

    try {
        Write-Log "Headers: $($script:GraphTokenHeader | ConvertTo-Json -Compress)" -Level 'DEBUG'
        $tokenResponse = Invoke-RestMethod -Uri $tokenEndpoint -Method Post -Body $body -ErrorAction Stop
        $script:GraphTokenHeader = @{
            'Authorization' = "Bearer $($tokenResponse.access_token)"
            'ConsistencyLevel' = 'eventual'
        }
        Write-Log "Successfully obtained authentication token."
    }
    catch {
        Write-Log "Failed to obtain authentication token. Error: $($_.Exception.Message)" -Level 'ERROR'
        throw "Authentication failed. Cannot continue."
    }
}

function Start-ThrottlingDelay {
    if ($script:ThrottleDelayMs -gt 0) {
        Write-Log "Applying delay of $($script:ThrottleDelayMs)ms before next request..." -Level 'WARN'
        Start-Sleep -Milliseconds $script:ThrottleDelayMs
    }
}

function Update-ThrottleDelayAfter429 {
    $script:LastThrottleEvent = Get-Date
    $script:LastActivity = Get-Date

    if ($script:ThrottleDelayMs -eq 0) {
        $script:ThrottleDelayMs = $script:MinDelayMs
    }
    else {
        $script:ThrottleDelayMs = [Math]::Min($script:ThrottleDelayMs * 2, $script:MaxDelayMs)
    }

    Write-Log "Throttling detected. Increasing delay to $($script:ThrottleDelayMs)ms." -Level 'WARN'
}

function Decay-ThrottleDelay {
    if ($script:ThrottleDelayMs -le $script:MinDelayMs) { return }

    $now = Get-Date
    $lastEventAge = $now - $script:LastThrottleEvent
    $decayThreshold = $script:DecayWindowMultiplier * ($script:ThrottleDelayMs / 1000)

    if ($lastEventAge.TotalSeconds -ge $decayThreshold) {
        $newDelay = [Math]::Max($script:ThrottleDelayMs / 2, $script:MinDelayMs)
        Write-Log "No throttling for $($lastEventAge.TotalSeconds.ToString("N1"))s. Reducing delay from $($script:ThrottleDelayMs)ms to $($newDelay)ms." -Level 'INFO'
        $script:ThrottleDelayMs = $newDelay
        $script:LastThrottleEvent = $now
    }
}

function Get-MessageStats {
    param(
        [string]$UserId,
        [string]$FolderWellKnownName
    )
    $allMessageSizes = [System.Collections.Generic.List[long]]::new()
    $dateFilterField = if ($FolderWellKnownName -eq 'SentItems') { 'sentDateTime' } else { 'receivedDateTime' }
    $cutoffDate = (Get-Date).AddDays(-$MaxMessageAgeInDays).ToString("yyyy-MM-ddTHH:mm:ssZ")
    $uri = "https://graph.microsoft.com/v1.0/users/$UserId/mailFolders/$FolderWellKnownName/messages"
    $filter = "$dateFilterField ge $cutoffDate"
    $expand = "singleValueExtendedProperties(`$filter=id eq 'Integer 0x0E08')"
    $queryParams = @(
        "`$filter=$filter"
        "`$select=id,$dateFilterField"
        "`$expand=$expand"
        "`$top=1000"
        "`$count=true"
    )
    $nextLink = "$uri`?$($queryParams -join '&')"
    $retryOn401 = $true  # Only retry once after token refresh

    do {
        try {
            Decay-ThrottleDelay
            Start-ThrottlingDelay
            Write-Log "Requesting: $nextLink" -Level 'DEBUG'

            $response = Invoke-RestMethod -Uri $nextLink -Headers $script:GraphTokenHeader -Method Get -ErrorAction Stop
            $script:LastActivity = Get-Date

            if ($response.value) {
                foreach ($message in $response.value) {
                    $sizeProp = $message.singleValueExtendedProperties | Where-Object { $_.id -eq 'Integer 0xe08' }
                    if ($sizeProp -and $sizeProp.value -and [long]::TryParse($sizeProp.value, [ref]0)) {
                        $allMessageSizes.Add([long]$sizeProp.value)
                    }
                }
            }
            $nextLink = $response.'@odata.nextLink'
        }
        catch {
            $script:LastActivity = Get-Date
            $statusCode = $_.Exception.Response.StatusCode.value__

            if ($statusCode -eq 401 -and $retryOn401) {
                Write-Log "Received 401 Unauthorized for user $UserId. Refreshing token and retrying..." -Level 'WARN'
                Get-GraphToken -TenantId $TenantId -ClientId $ClientId -ClientSecret $ClientSecret
                $script:LastTokenAcquired = Get-Date
                $retryOn401 = $false  # Prevent infinite loop
                continue  # Retry the same page with new token
            }
            elseif ($statusCode -eq 429) {
                Update-ThrottleDelayAfter429
                $retryAfter = $_.Exception.Response.Headers['Retry-After']
                if ($retryAfter) {
                    $script:ThrottleDelayMs = [int]$retryAfter * 1000
                }
            }
            elseif ($statusCode -eq 403) {
                Write-Log "Access denied for user $UserId. Skipping folder $FolderWellKnownName. Error: $($_.Exception.Message)" -Level 'ERROR'
                $nextLink = $null
            }
            elseif ($statusCode -eq 400) {
                try {
                    $stream = $_.Exception.Response.GetResponseStream()
                    $reader = New-Object System.IO.StreamReader($stream)
                    $errorBody = $reader.ReadToEnd()
                    $reader.Close(); $stream.Close()
                    Write-Log "Bad request for user $UserId. URI: $nextLink" -Level 'ERROR'
                    Write-Log "API Error (400): $errorBody" -Level 'ERROR'
                    try {
                        $errorObject = $errorBody | ConvertFrom-Json
                        Write-Log ("Detailed error: " + ($errorObject | Format-List | Out-String).Trim()) -Level 'ERROR'
                    } catch { }
                } catch {
                    Write-Log "Could not read error response body." -Level 'ERROR'
                }
                $nextLink = $null
            }
            else {
                Write-Log "Error retrieving messages for user $UserId from folder $FolderWellKnownName. Status: $statusCode. Error: $($_.Exception.Message)" -Level 'WARN'
                $nextLink = $null
            }
        }
    } while ($nextLink)

    if ($allMessageSizes.Count -gt 0) {
        $stats = $allMessageSizes | Measure-Object -Minimum -Maximum -Average -Sum
        return @{
            TotalMessages = $stats.Count
            MinSize       = $stats.Minimum
            MaxSize       = $stats.Maximum
            AvgSize       = [math]::Round($stats.Average)
            MeanSize      = [math]::Round($stats.Average)
            TotalBytes    = $stats.Sum
        }
    }
    else {
        return @{
            TotalMessages = 0
            MinSize       = 0
            MaxSize       = 0
            AvgSize       = 0
            MeanSize      = 0
            TotalBytes    = 0
        }
    }
}

#endregion Helper Functions

# --- Throttling Control Variables ---
$script:ThrottleDelayMs = 0
$script:LastThrottleEvent = Get-Date
$script:LastActivity = Get-Date
$script:MinDelayMs = 10
$script:MaxDelayMs = 30000
$script:DecayWindowMultiplier = 10

# --- Main Script Body ---

if ($host.Name -like '*ISE*') {
    Write-Log "Running in PowerShell ISE. Critical errors will use 'throw' to halt script without closing."
}

# --- Initialization and State Loading ---
$allUsersData = @()

if (Test-Path $CsvPath) {
    Write-Log "Existing CSV file found at '$CsvPath'. Loading data to resume."
    try {
        $allUsersData = Import-Csv -Path $CsvPath | ForEach-Object {
            [PSCustomObject]@{
                UserPrincipalName               = $_.UserPrincipalName
                UserCreationDate                = $_.UserCreationDate
                Processed                       = [int]$_.Processed
                MinReceivedMessageSize          = [long]$_.MinReceivedMessageSize
                AverageReceivedMessageSize      = [long]$_.AverageReceivedMessageSize
                MeanReceivedMessageSize         = [long]$_.MeanReceivedMessageSize
                MaximumReceivedMessageSize      = [long]$_.MaximumReceivedMessageSize
                TotalReceivedMessages           = [int]$_.TotalReceivedMessages
                MinSentMessageSize              = [long]$_.MinSentMessageSize
                AverageSentMessageSize          = [long]$_.AverageSentMessageSize
                MeanSentMessageSize             = [long]$_.MeanSentMessageSize
                MaximumSentMessageSize          = [long]$_.MaximumSentMessageSize
                TotalSentMessages               = [int]$_.TotalSentMessages
            }
        }
        Write-Log "Successfully loaded $($allUsersData.Count) user records."
    }
    catch {
        Write-Log "Failed to read or parse CSV file at '$CsvPath'. Error: $($_.Exception.Message)" -Level 'ERROR'
        throw "Cannot proceed due to corrupted or unreadable CSV file."
    }
}
else {
    Write-Log "No existing CSV file found. This is a first run. Populating user list from tenant."
    if (-not $script:GraphTokenHeader) {
        Get-GraphToken -TenantId $TenantId -ClientId $ClientId -ClientSecret $ClientSecret
        $script:LastTokenAcquired = Get-Date
    }

    $users = [System.Collections.ArrayList]::new()
    $nextLink = "https://graph.microsoft.com/v1.0/users?`$select=id,userPrincipalName,createdDateTime,accountEnabled,signInActivity&`$filter=accountEnabled eq true and (userType eq 'Member')&`$top=999"

    do {
        try {
            Write-Log "Headers: $($script:GraphTokenHeader | ConvertTo-Json -Compress)" -Level 'DEBUG'
            $response = Invoke-RestMethod -Uri $nextLink -Headers $script:GraphTokenHeader -Method Get -ErrorAction Stop
            [void]$users.AddRange(@($response.value))
            $nextLink = $response.'@odata.nextLink'
            Write-Log "Fetched $($users.Count) users so far..."
        } catch {
            Write-Log "Failed during user pagination. Error: $($_.Exception.Message)" -Level 'ERROR'
            throw "User population failed."
        }
    } while ($nextLink)

    Write-Log "Found $($users.Count) enabled member users. Filtering for those with active mailboxes and recent sign-ins..."

    $filteredUsers = @()
    $cutoffDate = (Get-Date).AddDays(-$MaxMessageAgeInDays)

    foreach ($user in $users) {
        # Skip guest users, external, or invalid UPNs
        if ($user.userPrincipalName -like "*#EXT#*" -or $user.userPrincipalName.EndsWith("onmicrosoft.com")) {
            continue  # Adjust domain as needed
        }

        # Optional: Filter by last sign-in (if signInActivity is available)
        $lastSignIn = $null
        if ($user.signInActivity.lastSignInDateTime) {
            $lastSignIn = [DateTime]::Parse($user.signInActivity.lastSignInDateTime)
        }

        if ($lastSignIn -and $lastSignIn -lt $cutoffDate) {
            Write-Log "Skipping user $($user.userPrincipalName): last sign-in $($lastSignIn) is older than $cutoffDate"
            continue
        }

        # Heuristic: Assume user has mailbox if account is enabled and not a service account
        $filteredUsers += [PSCustomObject]@{
            UserPrincipalName               = $user.userPrincipalName
            UserCreationDate                = $user.createdDateTime
            Processed                       = 0
            MinReceivedMessageSize          = 0
            AverageReceivedMessageSize      = 0
            MeanReceivedMessageSize         = 0
            MaximumReceivedMessageSize      = 0
            TotalReceivedMessages           = 0
            MinSentMessageSize              = 0
            AverageSentMessageSize          = 0
            MeanSentMessageSize             = 0
            MaximumSentMessageSize          = 0
            TotalSentMessages               = 0
        }
    }

    Write-Log "Filtered down to $($filteredUsers.Count) users with active accounts and recent sign-ins."
    $allUsersData = $filteredUsers
    $allUsersData | Export-Csv -Path $CsvPath -NoTypeInformation
    Write-Log "Initial filtered user list saved to '$CsvPath'."
}

# --- Main Processing Loop ---
$unprocessedUsers = $allUsersData | Where-Object { $_.Processed -eq 0 }
$totalToProcess = $unprocessedUsers.Count
$processedCount = 0

if ($totalToProcess -eq 0) {
    Write-Log "All users have already been processed. Script finished."
    return
}

Write-Log "Starting processing for $totalToProcess users."
if (-not $script:GraphTokenHeader) {
    Get-GraphToken -TenantId $TenantId -ClientId $ClientId -ClientSecret $ClientSecret
    $script:LastTokenAcquired = Get-Date
}

Write-Log "Headers: $($script:GraphTokenHeader | ConvertTo-Json -Compress)" -Level 'DEBUG'

foreach ($user in $unprocessedUsers) {
    $processedCount++
    
    # Check token age
    if (((Get-Date) - $script:LastTokenAcquired).TotalMinutes -gt 50) {
        Write-Log "Token is over 50 minutes old. Refreshing to prevent 401s..."
        Get-GraphToken -TenantId $TenantId -ClientId $ClientId -ClientSecret $ClientSecret
        $script:LastTokenAcquired = Get-Date
    }
    
    Write-Log "Processing user $($processedCount) of $($totalToProcess): $($user.UserPrincipalName)"

    $processingSuccess = $false
    try {
        Write-Log "--> Getting stats for Inbox..."
        $receivedStats = Get-MessageStats -UserId $user.UserPrincipalName -FolderWellKnownName 'Inbox'
        $user.TotalReceivedMessages = $receivedStats.TotalMessages
        $user.MinReceivedMessageSize = $receivedStats.MinSize
        $user.AverageReceivedMessageSize = $receivedStats.AvgSize
        $user.MeanReceivedMessageSize = $receivedStats.MeanSize
        $user.MaximumReceivedMessageSize = $receivedStats.MaxSize

        Write-Log "--> Getting stats for SentItems..."
        $sentStats = Get-MessageStats -UserId $user.UserPrincipalName -FolderWellKnownName 'SentItems'
        $user.TotalSentMessages = $sentStats.TotalMessages
        $user.MinSentMessageSize = $sentStats.MinSize
        $user.AverageSentMessageSize = $sentStats.AvgSize
        $user.MeanSentMessageSize = $sentStats.MeanSize
        $user.MaximumSentMessageSize = $sentStats.MaxSize

        $totalBytes = $receivedStats.TotalBytes + $sentStats.TotalBytes
        Write-Log "--> Successfully processed user $($user.UserPrincipalName). Inbox: $($receivedStats.TotalMessages) messages ($($receivedStats.TotalBytes) bytes), Sent: $($sentStats.TotalMessages) messages ($($sentStats.TotalBytes) bytes)"

        $processingSuccess = $true
    }
    catch {
        Write-Log "Failed to fully process user $($user.UserPrincipalName). Error: $($_.Exception.Message)" -Level 'ERROR'
    }
    finally {
        if ($processingSuccess) {
            $user.Processed = 1
        }
        else {
            # Ensure user is NOT marked as processed on failure
            $user.Processed = 0
        }
        $allUsersData | Export-Csv -Path $CsvPath -NoTypeInformation -Force
    }
}

Write-Log "Script finished. All users have been processed."
Write-Log "Final report is available at: $CsvPath"
