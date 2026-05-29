# ================================================================
# Add-UsersToGroups.ps1
# Bulk add users to Azure AD groups via Microsoft Graph API
# Auth: Service Principal | Input: CSV from repo
# ================================================================

param(
    [Parameter(Mandatory)] [string] $TenantId,
    [Parameter(Mandatory)] [string] $ClientId,
    [Parameter(Mandatory)] [string] $ClientSecret,
    [Parameter(Mandatory)] [string] $CsvPath,
    [string] $ResultsPath = "./results_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
)

# ----------------------------------------------------------------
# STEP 1: Authenticate — Get Bearer Token via SPN
# ----------------------------------------------------------------
Write-Host "`n[1/5] Authenticating with Service Principal..." -ForegroundColor Cyan

$tokenUrl  = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
$tokenBody = @{
    grant_type    = "client_credentials"
    client_id     = $ClientId
    client_secret = $ClientSecret
    scope         = "https://graph.microsoft.com/.default"
}

try {
    $tokenResponse = Invoke-RestMethod -Method Post -Uri $tokenUrl `
        -Body $tokenBody -ContentType "application/x-www-form-urlencoded" -ErrorAction Stop
    $accessToken = $tokenResponse.access_token
    $headers = @{
        Authorization  = "Bearer $accessToken"
        "Content-Type" = "application/json"
    }
    Write-Host "   ✅ Authentication successful." -ForegroundColor Green
} catch {
    Write-Error "   ❌ Authentication failed: $($_.Exception.Message)"
    exit 1
}

# ----------------------------------------------------------------
# STEP 2: Load and Validate CSV
# ----------------------------------------------------------------
Write-Host "`n[2/5] Loading CSV from: $CsvPath" -ForegroundColor Cyan

if (-not (Test-Path $CsvPath)) {
    Write-Error "   ❌ CSV file not found at path: $CsvPath"
    exit 1
}

$csvData = Import-Csv -Path $CsvPath

# Validate required columns
$requiredCols = @("UserPrincipalName", "GroupName")
foreach ($col in $requiredCols) {
    if ($csvData[0].PSObject.Properties.Name -notcontains $col) {
        Write-Error "   ❌ Missing required column '$col' in CSV."
        exit 1
    }
}

# Filter out blank rows
$csvData = $csvData | Where-Object {
    -not [string]::IsNullOrWhiteSpace($_.UserPrincipalName) -and
    -not [string]::IsNullOrWhiteSpace($_.GroupName)
}

Write-Host "   ✅ Loaded $($csvData.Count) valid rows." -ForegroundColor Green

# ----------------------------------------------------------------
# STEP 3: Cache Group Object IDs (avoid duplicate lookups)
# ----------------------------------------------------------------
Write-Host "`n[3/5] Resolving Group Object IDs..." -ForegroundColor Cyan

$groupCache = @{}
$uniqueGroups = $csvData.GroupName | Sort-Object -Unique

foreach ($groupName in $uniqueGroups) {
    # FIX 1: Use -f operator to build URL so & is never parsed as PS operator
    $uri = "https://graph.microsoft.com/v1.0/groups?`$filter=displayName eq '{0}'`&`$select=id,displayName" -f $groupName

    try {
        $response = Invoke-RestMethod -Method Get -Uri $uri -Headers $headers -ErrorAction Stop
        if ($response.value.Count -eq 0) {
            Write-Warning "   ⚠️  Group not found: '$groupName'"
            $groupCache[$groupName] = $null
        } else {
            $groupCache[$groupName] = $response.value[0].id
            Write-Host "   ✅ Resolved group '$groupName' → $($response.value[0].id)" -ForegroundColor Green
        }
    } catch {
        Write-Warning "   ⚠️  Error resolving group '$groupName': $($_.Exception.Message)"
        $groupCache[$groupName] = $null
    }
}

# ----------------------------------------------------------------
# STEP 4: Cache User Object IDs + Add to Groups
# ----------------------------------------------------------------
Write-Host "`n[4/5] Processing user-group assignments..." -ForegroundColor Cyan

$userCache = @{}
$results   = @()
$counter   = 0

foreach ($row in $csvData) {
    $counter++
    $upn       = $row.UserPrincipalName.Trim()
    $groupName = $row.GroupName.Trim()

    Write-Host "`n   [$counter/$($csvData.Count)] $upn → $groupName" -ForegroundColor White

    # -- Resolve User (use cache if already looked up)
    if (-not $userCache.ContainsKey($upn)) {
        try {
            $userResp = Invoke-RestMethod -Method Get `
                -Uri "https://graph.microsoft.com/v1.0/users/$upn`?`$select=id,displayName" `
                -Headers $headers -ErrorAction Stop
            $userCache[$upn] = $userResp.id
        } catch {
            Write-Warning "      ⚠️  User not found: $upn — skipping."
            $userCache[$upn] = $null
        }
    }

    $userId  = $userCache[$upn]
    $groupId = $groupCache[$groupName]

    # -- Skip if user or group unresolvable
    if (-not $userId) {
        $results += [PSCustomObject]@{
            UPN       = $upn
            Group     = $groupName
            Status    = "Skipped"
            Reason    = "User not found in Azure AD"
        }
        continue
    }

    if (-not $groupId) {
        $results += [PSCustomObject]@{
            UPN       = $upn
            Group     = $groupName
            Status    = "Skipped"
            Reason    = "Group not found in Azure AD"
        }
        continue
    }

    # -- Check if user is already a member (avoid duplicate error)
    # FIX 2: Use -f operator to build URL so & is never parsed as PS operator
    $memberCheckUri = "https://graph.microsoft.com/v1.0/groups/{0}/members?`$filter=id eq '{1}'`&`$select=id" -f $groupId, $userId
    try {
        $memberCheck = Invoke-RestMethod -Method Get -Uri $memberCheckUri -Headers $headers -ErrorAction Stop
        if ($memberCheck.value.Count -gt 0) {
            Write-Host "      ℹ️  Already a member — skipping." -ForegroundColor Yellow
            $results += [PSCustomObject]@{
                UPN    = $upn
                Group  = $groupName
                Status = "AlreadyMember"
                Reason = "User is already in this group"
            }
            continue
        }
    } catch {
        # Non-fatal — proceed to add anyway
    }

    # -- Add user to group
    $addBody = @{
        "@odata.id" = "https://graph.microsoft.com/v1.0/directoryObjects/$userId"
    } | ConvertTo-Json

    try {
        Invoke-RestMethod -Method Post `
            -Uri "https://graph.microsoft.com/v1.0/groups/$groupId/members/`$ref" `
            -Headers $headers -Body $addBody -ErrorAction Stop

        Write-Host "      ✅ Added successfully." -ForegroundColor Green
        $results += [PSCustomObject]@{
            UPN    = $upn
            Group  = $groupName
            Status = "Success"
            Reason = ""
        }
    } catch {
        $errDetail = ""
        try   { $errDetail = ($_.ErrorDetails.Message | ConvertFrom-Json).error.message }
        catch { $errDetail = $_.Exception.Message }

        Write-Warning "      ❌ Failed: $errDetail"
        $results += [PSCustomObject]@{
            UPN    = $upn
            Group  = $groupName
            Status = "Failed"
            Reason = $errDetail
        }
    }

    Start-Sleep -Milliseconds 200   # Avoid Graph throttling
}

# ----------------------------------------------------------------
# STEP 5: Export Results + Print Summary
# ----------------------------------------------------------------
Write-Host "`n[5/5] Exporting results to: $ResultsPath" -ForegroundColor Cyan
$results | Export-Csv -Path $ResultsPath -NoTypeInformation

$success       = ($results | Where-Object Status -eq "Success").Count
$failed        = ($results | Where-Object Status -eq "Failed").Count
$skipped       = ($results | Where-Object Status -eq "Skipped").Count
$alreadyMember = ($results | Where-Object Status -eq "AlreadyMember").Count

Write-Host "`n============== SUMMARY ==============" -ForegroundColor Yellow
Write-Host "  ✅ Success        : $success"        -ForegroundColor Green
Write-Host "  ❌ Failed         : $failed"         -ForegroundColor Red
Write-Host "  ⚠️  Skipped        : $skipped"        -ForegroundColor Yellow
Write-Host "  ℹ️  Already Member : $alreadyMember"  -ForegroundColor Cyan
Write-Host "  📋 Total Processed: $($results.Count)"
Write-Host "======================================`n" -ForegroundColor Yellow

# Exit with error code if any hard failures (for Harness step status)
if ($failed -gt 0) {
    Write-Warning "Some assignments failed. Check results CSV."
    exit 2   # Harness will mark step as failed but pipeline continues
}

exit 0