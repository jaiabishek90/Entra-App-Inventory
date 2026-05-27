#Requires -Version 5.1
<#
.SYNOPSIS
Generates an Entra enterprise application inventory report.

.DESCRIPTION
Collects integrated enterprise applications from Microsoft Entra ID (Azure AD),
enriches them with owners, memberships, assignments, application permissions,
delegated permissions, credential metadata, sign-in activity, and Exchange
Application Access Policy scoping. Exports the result to CSV and optionally
sends an HTML summary email.

.PARAMETER ConfigPath
Path to the JSON configuration file.

.PARAMETER OutputDirectory
Directory where the CSV export will be written.

.PARAMETER SkipEmail
Skips the HTML email summary.

.PARAMETER WhatIf
Shows what would happen if the script runs.

.EXAMPLE
pwsh .\src\Get-EntraEnterpriseAppInventory.ps1 -ConfigPath .\config\settings.json

.EXAMPLE
pwsh .\src\Get-EntraEnterpriseAppInventory.ps1 -ConfigPath .\config\settings.json -SkipEmail

.NOTES
- Uses Microsoft Graph Beta cmdlets.
- Replace Send-MailMessage if your organization mandates a supported mail API.
- Existing high-priority logic is preserved from the source implementation.
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$ConfigPath,

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$OutputDirectory = (Get-Location).Path,

    [Parameter()]
    [switch]$SkipEmail
)

Set-StrictMode -Version Latest
$ProgressPreference    = 'Continue'
$VerbosePreference     = 'SilentlyContinue'
$ErrorActionPreference = 'Stop'

function Import-RequiredModule {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    if (-not (Get-Module -ListAvailable -Name $Name)) {
        throw "Required module '$Name' is not installed."
    }

    Import-Module $Name -Verbose:$false -ErrorAction Stop
}

function Format-DateYMD {
    param([datetime]$Date)
    if ($Date) { Get-Date $Date -Format 'yyyy-MM-dd' } else { 'N/A' }
}

function Convert-Base64ThumbprintToHex {
    param([string]$Base64Thumbprint)
    if ([string]::IsNullOrEmpty($Base64Thumbprint)) { return $null }
    try {
        ([BitConverter]::ToString([Convert]::FromBase64String($Base64Thumbprint))).Replace('-', '')
    }
    catch {
        $null
    }
}

function Get-DaysLeft {
    param([datetime]$EndDate)
    if ($EndDate) { [math]::Floor(($EndDate - (Get-Date)).TotalDays) } else { $null }
}

function Is-Expired {
    param([datetime]$EndDate)
    if ($EndDate) { $EndDate -lt (Get-Date) } else { $false }
}

function Split-CsvValues {
    param([string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) {
        @()
    }
    else {
        $Value -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ }
    }
}

function Invoke-WithRetry {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ScriptBlock]$Script,

        [int]$MaxAttempts = 5,
        [int]$InitialDelayMs = 500
    )

    $attempt = 0
    $delay   = $InitialDelayMs

    while ($true) {
        try {
            return & $Script
        }
        catch {
            $attempt++
            $message = $_.Exception.Message

            if (
                $attempt -lt $MaxAttempts -and
                ($message -match '429' -or
                 $message -match '5\d\d' -or
                 $message -match 'timeout' -or
                 $message -match 'temporar')
            ) {
                Start-Sleep -Milliseconds $delay
                $delay = [Math]::Min($delay * 2, 8000)
                continue
            }

            throw
        }
    }
}

function ConvertTo-HtmlTable {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        $InputObjects
    )

    if (-not $InputObjects -or $InputObjects.Count -eq 0) {
        return "<p class='meta'>None</p>"
    }

    $columns = $InputObjects[0].PSObject.Properties.Name
    $header  = ($columns | ForEach-Object { "<th>$_</th>" }) -join ''

    $rowsHtml = foreach ($row in $InputObjects) {
        $cells = foreach ($column in $columns) {
            $value = $row.$column
            if ($null -eq $value -or $value -eq '') { $value = '&nbsp;' }
            "<td>" + [System.Web.HttpUtility]::HtmlEncode([string]$value) + "</td>"
        }
        "<tr>$($cells -join '')</tr>"
    }

    "<table><tr>$header</tr>$($rowsHtml -join '')</table>"
}

function Get-Configuration {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    if (-not (Test-Path -Path $Path -PathType Leaf)) {
        throw "Configuration file not found: $Path"
    }

    $config = Get-Content -Path $Path -Raw | ConvertFrom-Json

    $required = @(
        'TenantId',
        'AppId',
        'CertThumbprint',
        'ExchangeOrganization',
        'SmtpServer',
        'MailFrom',
        'Recipients',
        'HighPriorityPermissions',
        'EmailDisplayExceptions'
    )

    foreach ($property in $required) {
        if (-not $config.PSObject.Properties.Name.Contains($property)) {
            throw "Missing required configuration property: $property"
        }
    }

    return $config
}

Add-Type -AssemblyName System.Web | Out-Null

$config = Get-Configuration -Path $ConfigPath

$IncludeOwnersAndMemberships = if ($config.PSObject.Properties.Name.Contains('IncludeOwnersAndMemberships')) { [bool]$config.IncludeOwnersAndMemberships } else { $true }
$IncludeAssignedPrincipals   = if ($config.PSObject.Properties.Name.Contains('IncludeAssignedPrincipals'))   { [bool]$config.IncludeAssignedPrincipals }   else { $true }
$ResolveResourceNames        = if ($config.PSObject.Properties.Name.Contains('ResolveResourceNames'))        { [bool]$config.ResolveResourceNames }        else { $true }
$PreloadAllResourceSPs       = if ($config.PSObject.Properties.Name.Contains('PreloadAllResourceSPs'))       { [bool]$config.PreloadAllResourceSPs }       else { $true }

Import-RequiredModule -Name 'Microsoft.Graph.Beta.Applications'
Import-RequiredModule -Name 'ExchangeOnlineManagement'

$connectParams = @{
    TenantId              = $config.TenantId
    ApplicationId         = $config.AppId
    CertificateThumbprint = $config.CertThumbprint
}

Connect-MgGraph @connectParams -NoWelcome
Connect-ExchangeOnline -CertificateThumbprint $config.CertThumbprint -AppId $config.AppId -Organization $config.ExchangeOrganization

[array]$HighPriorityPermissions = @($config.HighPriorityPermissions)
$HighPrioritySet = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
$null = $HighPriorityPermissions | ForEach-Object { $HighPrioritySet.Add($_) }

[array]$EmailDisplayExceptions = @($config.EmailDisplayExceptions)
$EmailExceptionsSet = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
$null = $EmailDisplayExceptions | ForEach-Object {
    $value = $_.ToString().Trim()
    if ($value) { $EmailExceptionsSet.Add($value) }
}

$TenantOwnerOrgCache = @{}
$ResourceSpCache     = @{}
$ApplicationsByAppId = @{}

function Get-SPOwnerOrg {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$TenantObjectId
    )

    if ($TenantOwnerOrgCache.ContainsKey($TenantObjectId)) {
        return $TenantOwnerOrgCache[$TenantObjectId]
    }

    try {
        $response = Invoke-WithRetry {
            Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/beta/tenantRelationships/findTenantInformationByTenantId(tenantId='$TenantObjectId')"
        }

        $TenantOwnerOrgCache[$TenantObjectId] = if ($response -and $response.defaultDomainName) {
            $response.defaultDomainName
        }
        else {
            $null
        }
    }
    catch {
        $TenantOwnerOrgCache[$TenantObjectId] = $null
    }

    return $TenantOwnerOrgCache[$TenantObjectId]
}

function Cache-ResourceSP {
    [CmdletBinding()]
    param($ServicePrincipal)

    if ($ServicePrincipal -and $ServicePrincipal.Id -and -not $ResourceSpCache.ContainsKey($ServicePrincipal.Id)) {
        $ResourceSpCache[$ServicePrincipal.Id] = [pscustomobject]@{
            Id             = $ServicePrincipal.Id
            AppDisplayName = $ServicePrincipal.AppDisplayName
            AppRoles       = $ServicePrincipal.AppRoles
        }
    }
}

function Get-ResourceSpCached {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Id
    )

    if ($ResourceSpCache.ContainsKey($Id)) {
        return $ResourceSpCache[$Id]
    }

    try {
        $sp = Invoke-WithRetry {
            Get-MgBetaServicePrincipal -ServicePrincipalId $Id -Property id,appDisplayName,appRoles -Verbose:$false -ErrorAction Stop
        }

        Cache-ResourceSP $sp
    }
    catch {
        $ResourceSpCache[$Id] = $null
    }

    return $ResourceSpCache[$Id]
}

$AppScopeByAppId = @{}
try {
    $policies = Get-ApplicationAccessPolicy | Select-Object AppId, ScopeName
    foreach ($policy in $policies) {
        if (-not $policy.AppId) { continue }

        if (-not $AppScopeByAppId.ContainsKey($policy.AppId)) {
            $AppScopeByAppId[$policy.AppId] = New-Object 'System.Collections.Generic.List[string]'
        }

        if ($policy.ScopeName -and -not $AppScopeByAppId[$policy.AppId].Contains($policy.ScopeName)) {
            $AppScopeByAppId[$policy.AppId].Add($policy.ScopeName)
        }
    }
}
catch {
    Write-Warning "Failed to load ApplicationAccessPolicy: $($_.Exception.Message)"
}

$servicePrincipalProperties = @(
    'id',
    'appId',
    'displayName',
    'appDisplayName',
    'servicePrincipalType',
    'accountEnabled',
    'publisherName',
    'verifiedPublisher',
    'homepage',
    'tags',
    'appOwnerOrganizationId',
    'createdDateTime',
    'passwordCredentials',
    'keyCredentials',
    'tokenEncryptionKeyId',
    'customSecurityAttributes',
    'appRoleAssignmentRequired'
) -join ','

$ServicePrincipals = Invoke-WithRetry {
    Get-MgBetaServicePrincipal -All -PageSize 999 `
        -Filter "tags/any(t:t eq 'WindowsAzureActiveDirectoryIntegratedApp')" `
        -Property $servicePrincipalProperties `
        -Verbose:$false
}

$ServicePrincipalStats = @{}
Invoke-WithRetry { Get-MgBetaReportServicePrincipalSignInActivity -All -Verbose:$false } | ForEach-Object {
    if (-not $ServicePrincipalStats[$_.AppId]) {
        $ServicePrincipalStats[$_.AppId] = @{
            LastSignIn                 = $_.LastSignInActivity.LastSignInDateTime
            LastDelegateClientSignIn   = $_.DelegatedClientSignInActivity.LastSignInDateTime
            LastDelegateResourceSignIn = $_.DelegatedResourceSignInActivity.LastSignInDateTime
            LastAppClientSignIn        = $_.ApplicationAuthenticationClientSignInActivity.LastSignInDateTime
            LastAppResourceSignIn      = $_.ApplicationAuthenticationResourceSignInActivity.LastSignInDateTime
        }
    }
}

$ServicePrincipalSummaryStats = @{}
Invoke-WithRetry { Get-MgBetaReportAzureAdApplicationSignInSummary -Period D30 -Verbose:$false } | ForEach-Object {
    if (-not $ServicePrincipalSummaryStats[$_.Id]) {
        $ServicePrincipalSummaryStats[$_.Id] = @{
            SignInSuccessCount = $_.SuccessfulSignInCount
            SignInFailureCount = $_.FailedSignInCount
        }
    }
}

if ($ResolveResourceNames -and $PreloadAllResourceSPs) {
    Invoke-WithRetry { Get-MgBetaServicePrincipal -All -Property id,appDisplayName,appRoles -Verbose:$false } |
        ForEach-Object { Cache-ResourceSP $_ }
}

Invoke-WithRetry {
    Get-MgBetaApplication -All -Property id,appId,displayName,createdDateTime,passwordCredentials,keyCredentials -Verbose:$false
} | ForEach-Object {
    if ($_.AppId -and -not $ApplicationsByAppId.ContainsKey($_.AppId)) {
        $ApplicationsByAppId[$_.AppId] = $_
    }
}

if (-not (Test-Path -Path $OutputDirectory -PathType Container)) {
    New-Item -Path $OutputDirectory -ItemType Directory -Force | Out-Null
}

try {
    Get-ChildItem -Path $OutputDirectory -Filter '*_EntraAppInventory.csv' -File -ErrorAction SilentlyContinue |
        Remove-Item -Force -ErrorAction SilentlyContinue
}
catch {
}

$rows   = New-Object 'System.Collections.Generic.List[object]'
$total  = @($ServicePrincipals).Count
$timer  = [System.Diagnostics.Stopwatch]::StartNew()
$now    = Get-Date
$since  = $now.AddDays(-7)
$index  = 0

foreach ($sp in $ServicePrincipals) {
    $index++

    $elapsedSeconds = [Math]::Max(1, [int][Math]::Round($timer.Elapsed.TotalSeconds, 0))
    $rate           = if ($index -gt 0) { [Math]::Round($index / $elapsedSeconds, 2) } else { 0 }
    $remaining      = if ($index -gt 0 -and $rate -gt 0) { [int][Math]::Round((($total - $index) / $rate), 0) } else { 0 }
    $percent        = if ($total -gt 0) { [int][Math]::Round((($index * 100.0) / $total), 0) } else { 0 }

    if ($percent -lt 0)   { $percent = 0 }
    if ($percent -gt 100) { $percent = 100 }

    Write-Progress -Id 1 `
        -Activity "Entra App Inventory: processing $index of $total" `
        -Status ("Rate: {0}/s | Elapsed: {1:mm\:ss} | ETA ~ {2}s" -f $rate, $timer.Elapsed, $remaining) `
        -PercentComplete $percent

    $scopesForApp = $null
    $isScopedApp  = $false
    if ($sp.AppId -and $AppScopeByAppId.ContainsKey($sp.AppId)) {
        $isScopedApp = $true
        $scopesForApp = ($AppScopeByAppId[$sp.AppId] | Select-Object -Unique) -join ';'
    }

    $ownersJoined   = $null
    $memberOfGroups = $null
    $memberOfRoles  = $null

    if ($IncludeOwnersAndMemberships) {
        try {
            $owners = @()
            Invoke-WithRetry { Get-MgBetaServicePrincipalOwner -ServicePrincipalId $sp.Id -All -Verbose:$false } |
                ForEach-Object {
                    $upn  = $null
                    $name = $null

                    if ($_.PSObject.Properties.Match('UserPrincipalName').Count) { $upn = $_.UserPrincipalName }
                    if ($_.PSObject.Properties.Match('DisplayName').Count)       { $name = $_.DisplayName }
                    if ($_.PSObject.Properties.Match('AdditionalProperties').Count -and $_.AdditionalProperties) {
                        if (-not $upn  -and $_.AdditionalProperties.ContainsKey('userPrincipalName')) { $upn  = $_.AdditionalProperties['userPrincipalName'] }
                        if (-not $name -and $_.AdditionalProperties.ContainsKey('displayName'))       { $name = $_.AdditionalProperties['displayName'] }
                    }

                    if ($upn) { $owners += $upn }
                    elseif ($name) { $owners += $name }
                }

            if ($owners.Count) { $ownersJoined = ($owners | Select-Object -Unique) -join ',' }
        }
        catch {}

        try {
            $memberOf = Invoke-WithRetry { Get-MgBetaServicePrincipalMemberOf -All -ServicePrincipalId $sp.Id -Verbose:$false }
            $memberOfGroups = ($memberOf.AdditionalProperties | Where-Object { $_.'@odata.type' -eq '#microsoft.graph.group' }).DisplayName -join ';'
            $memberOfRoles  = ($memberOf.AdditionalProperties | Where-Object { $_.'@odata.type' -eq '#microsoft.graph.directoryRole' }).DisplayName -join ';'
        }
        catch {}
    }

    $assignedUsers  = @()
    $assignedGroups = @()
    if ($IncludeAssignedPrincipals) {
        try {
            $assigned = Invoke-WithRetry { Get-MgBetaServicePrincipalAppRoleAssignedTo -All -ServicePrincipalId $sp.Id -Verbose:$false }
            if ($assigned) {
                $assignedUsers  = ($assigned | Where-Object { $_.PrincipalType -eq 'User'  } | Select-Object -ExpandProperty PrincipalDisplayName) 2>$null
                $assignedGroups = ($assigned | Where-Object { $_.PrincipalType -eq 'Group' } | Select-Object -ExpandProperty PrincipalDisplayName) 2>$null
            }
        }
        catch {}
    }

    $appPermByResource = @{}
    try {
        $appRoleAssignments = Invoke-WithRetry { Get-MgBetaServicePrincipalAppRoleAssignment -All -ServicePrincipalId $sp.Id -Verbose:$false }
        if ($appRoleAssignments) {
            foreach ($assignment in $appRoleAssignments) {
                $resourceName = $assignment.ResourceDisplayName
                if (-not $resourceName) {
                    if ($ResolveResourceNames -and $assignment.ResourceId) {
                        $resourceSp = if ($PreloadAllResourceSPs) { $ResourceSpCache[$assignment.ResourceId] } else { Get-ResourceSpCached $assignment.ResourceId }
                        if ($resourceSp -and $resourceSp.AppDisplayName) { $resourceName = $resourceSp.AppDisplayName }
                    }
                    if (-not $resourceName) { $resourceName = $assignment.ResourceId }
                }

                $roleValue = $null
                if ($assignment.AppRoleId -and $assignment.ResourceId -and $ResolveResourceNames) {
                    $resourceSpForRole = if ($PreloadAllResourceSPs) { $ResourceSpCache[$assignment.ResourceId] } else { Get-ResourceSpCached $assignment.ResourceId }
                    if ($resourceSpForRole -and $resourceSpForRole.AppRoles) {
                        $role = $resourceSpForRole.AppRoles | Where-Object { $_.Id -eq $assignment.AppRoleId }
                        if ($role) { $roleValue = $role.Value }
                    }
                }
                if (-not $roleValue) { $roleValue = "Orphaned ($($assignment.AppRoleId))" }

                if ($appPermByResource.ContainsKey($resourceName)) {
                    if ($appPermByResource[$resourceName] -notmatch [regex]::Escape($roleValue)) {
                        $appPermByResource[$resourceName] += ",$roleValue"
                    }
                }
                else {
                    $appPermByResource[$resourceName] = $roleValue
                }
            }
        }
    }
    catch {}

    $delegatedPermByResource = @{}
    $delegatedValidUntil = $null
    try {
        $oauthGrants = Invoke-WithRetry { Get-MgBetaServicePrincipalOAuth2PermissionGrant -All -ServicePrincipalId $sp.Id -Verbose:$false }
        if ($oauthGrants) {
            foreach ($grant in $oauthGrants) {
                $resourceName = $null
                if ($grant.ResourceId) {
                    if ($ResolveResourceNames) {
                        $resourceSp = if ($PreloadAllResourceSPs) { $ResourceSpCache[$grant.ResourceId] } else { Get-ResourceSpCached $grant.ResourceId }
                        if ($resourceSp -and $resourceSp.AppDisplayName) { $resourceName = $resourceSp.AppDisplayName }
                    }
                    if (-not $resourceName) { $resourceName = $grant.ResourceId }
                }
                else {
                    $resourceName = '(Unknown Resource)'
                }

                $scopes = if ($grant.Scope) { $grant.Scope -split ' ' } else { @('Orphaned scope') }
                foreach ($scope in $scopes) {
                    if ($delegatedPermByResource.ContainsKey($resourceName)) {
                        if ($delegatedPermByResource[$resourceName] -notmatch [regex]::Escape($scope)) {
                            $delegatedPermByResource[$resourceName] += ",$scope"
                        }
                    }
                    else {
                        $delegatedPermByResource[$resourceName] = $scope
                    }
                }

                if ($grant.ExpiryTime) {
                    $expiry = [datetime]$grant.ExpiryTime
                    if (-not $delegatedValidUntil -or $expiry -gt $delegatedValidUntil) {
                        $delegatedValidUntil = $expiry
                    }
                }
            }
        }
    }
    catch {}

    $applicationPermissionsJoined = $null
    if ($appPermByResource.Keys.Count -gt 0) {
        $parts = foreach ($resource in ($appPermByResource.Keys | Sort-Object)) {
            "[{0}]: {1}" -f $resource, $appPermByResource[$resource]
        }
        $applicationPermissionsJoined = $parts -join '; '
    }

    $delegatedPermissionsJoined = $null
    if ($delegatedPermByResource.Keys.Count -gt 0) {
        $parts = foreach ($resource in ($delegatedPermByResource.Keys | Sort-Object)) {
            "[{0}]: {1}" -f $resource, $delegatedPermByResource[$resource]
        }
        $delegatedPermissionsJoined = $parts -join '; '
    }

    $highPriorityAppMatches = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    foreach ($resource in $appPermByResource.Keys) {
        foreach ($value in (Split-CsvValues $appPermByResource[$resource])) {
            if ($HighPrioritySet.Contains($value)) { $null = $highPriorityAppMatches.Add($value) }
        }
    }

    $highPriorityDelegatedMatches = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    foreach ($resource in $delegatedPermByResource.Keys) {
        foreach ($value in (Split-CsvValues $delegatedPermByResource[$resource])) {
            if ($HighPrioritySet.Contains($value)) { $null = $highPriorityDelegatedMatches.Add($value) }
        }
    }

    $hasHighPriorityAppRaw = $highPriorityAppMatches.Count -gt 0
    $hasHighPriorityDelRaw = $highPriorityDelegatedMatches.Count -gt 0
    $excludeFromHighPriority = $isScopedApp

    $hasHighPriorityApp = $hasHighPriorityAppRaw -and -not $excludeFromHighPriority
    $hasHighPriorityDel = $hasHighPriorityDelRaw -and -not $excludeFromHighPriority
    $hasHighPriorityAny = $hasHighPriorityApp

    $highPriorityAppJoined = if ($hasHighPriorityAppRaw) { (@($highPriorityAppMatches) | Sort-Object) -join ',' } else { $null }
    $highPriorityDelegatedJoined = if ($hasHighPriorityDelRaw) { (@($highPriorityDelegatedMatches) | Sort-Object) -join ',' } else { $null }
    $highPriorityAllJoined = if ($hasHighPriorityAppRaw -or $hasHighPriorityDelRaw) { (@($highPriorityAppMatches) + @($highPriorityDelegatedMatches) | Sort-Object -Unique) -join ',' } else { $null }

    $signInStats  = $ServicePrincipalStats[$sp.AppId]
    $summaryStats = $ServicePrincipalSummaryStats[$sp.AppId]

    $createdRaw = $sp.CreatedDateTime
    if (-not $createdRaw -and $sp.AdditionalProperties.CreatedDateTime) { $createdRaw = [datetime]$sp.AdditionalProperties.CreatedDateTime }
    $isNewInLast7Days = $false
    if ($createdRaw) { $isNewInLast7Days = [datetime]$createdRaw -ge $since }

    $allPasswordCredentials = @()
    $allKeyCredentials      = @()
    $secretExpiredAny       = $false
    $certExpiredAny         = $false
    $credentialSourceSecrets = @()
    $credentialSourceCerts   = @()

    $appObject = $null
    if ($sp.AppId -and $ApplicationsByAppId.ContainsKey($sp.AppId)) { $appObject = $ApplicationsByAppId[$sp.AppId] }

    if ($appObject -and $appObject.PasswordCredentials) {
        $allPasswordCredentials += $appObject.PasswordCredentials
        $credentialSourceSecrets += 'Application'
    }
    if ($sp.PasswordCredentials) {
        $allPasswordCredentials += $sp.PasswordCredentials
        $credentialSourceSecrets += 'ServicePrincipal'
    }
    if ($appObject -and $appObject.KeyCredentials) {
        $allKeyCredentials += $appObject.KeyCredentials
        $credentialSourceCerts += 'Application'
    }
    if ($sp.KeyCredentials) {
        $allKeyCredentials += $sp.KeyCredentials
        $credentialSourceCerts += 'ServicePrincipal'
    }

    $secretSummaries = @()
    $earliestSecret  = $null
    foreach ($credential in $allPasswordCredentials) {
        $name     = if ($credential.DisplayName) { $credential.DisplayName } else { $credential.KeyId }
        $start    = $credential.StartDateTime
        $end      = $credential.EndDateTime
        $expired  = Is-Expired $end
        $daysLeft = Get-DaysLeft $end
        if ($expired) { $secretExpiredAny = $true }
        if ($end -and (-not $earliestSecret -or $end -lt $earliestSecret)) { $earliestSecret = $end }
        $secretSummaries += "{0}|{1}|{2}|{3}|{4}" -f $name, (Format-DateYMD $start), (Format-DateYMD $end), $expired, $daysLeft
    }

    $certSummaries = @()
    $earliestCert  = $null
    foreach ($credential in $allKeyCredentials) {
        $name      = if ($credential.DisplayName) { $credential.DisplayName } else { $credential.KeyId }
        $start     = $credential.StartDateTime
        $end       = $credential.EndDateTime
        $type      = $credential.Type
        $thumb     = Convert-Base64ThumbprintToHex $credential.CustomKeyIdentifier
        $expired   = Is-Expired $end
        $daysLeft  = Get-DaysLeft $end
        if ($expired) { $certExpiredAny = $true }
        if ($end -and (-not $earliestCert -or $end -lt $earliestCert)) { $earliestCert = $end }
        $certSummaries += "{0}|{1}|{2}|{3}|{4}|{5}|{6}" -f $name, (Format-DateYMD $start), (Format-DateYMD $end), $type, $thumb, $expired, $daysLeft
    }

    $credentialSourceSecretsString = if ($credentialSourceSecrets.Count) { ($credentialSourceSecrets | Select-Object -Unique) -join '+' } else { 'None' }
    $credentialSourceCertsString   = if ($credentialSourceCerts.Count)   { ($credentialSourceCerts   | Select-Object -Unique) -join '+' } else { 'None' }

    $customSecurityAttributesString = if ($sp.CustomSecurityAttributes) {
        $attributeSets = @()
        foreach ($attributeSet in $sp.CustomSecurityAttributes.AdditionalProperties.GetEnumerator()) {
            $setName = $attributeSet.Key
            $attributes = @()
            foreach ($property in $attributeSet.Value.GetEnumerator()) {
                if ($property.Key -eq '@odata.type') { continue }
                $attributes += "$($property.Key):$($property.Value)"
            }
            $attributeSets += "[$setName]$($attributes -join '|')"
        }
        $attributeSets -join ';'
    }
    else { $null }

    $rows.Add([pscustomobject]@{
        'Application Name' = if ($sp.AppDisplayName) { $sp.AppDisplayName } else { $sp.DisplayName }
        'ApplicationId'    = $sp.AppId
        'ObjectId'         = $sp.Id
        'Type'             = $sp.ServicePrincipalType
        'Enabled'          = $sp.AccountEnabled
        'Visible to users?'    = (-not ($sp.Tags -contains 'HideApp'))
        'Assignment required?' = $sp.AppRoleAssignmentRequired
        'Publisher' = $sp.PublisherName
        'Owned by org' = if ($sp.AppOwnerOrganizationId) {
            $ownerDomain = Get-SPOwnerOrg $sp.AppOwnerOrganizationId
            if ($ownerDomain) { "$($sp.AppOwnerOrganizationId) ($ownerDomain)" } else { $sp.AppOwnerOrganizationId }
        }
        else { $null }
        'Verified'   = if ($sp.VerifiedPublisher.VerifiedPublisherId) { $sp.VerifiedPublisher.DisplayName } else { 'Not verified' }
        'Homepage'   = $sp.Homepage
        'Created on' = if ($createdRaw) { (Get-Date $createdRaw -Format g) } else { 'N/A' }
        'Owners'             = $ownersJoined
        'Member of (groups)' = $memberOfGroups
        'Member of (roles)'  = $memberOfRoles
        'Assigned users (count)'  = ($assignedUsers | Measure-Object).Count
        'Assigned users (names)'  = if ($assignedUsers.Count)  { ($assignedUsers  | Select-Object -Unique) -join ';' } else { $null }
        'Assigned groups (count)' = ($assignedGroups | Measure-Object).Count
        'Assigned groups (names)' = if ($assignedGroups.Count) { ($assignedGroups | Select-Object -Unique) -join ';' } else { $null }
        'Client secrets (name|start|end|expired|daysLeft)' = if ($secretSummaries.Count) { $secretSummaries -join ';' } else { $null }
        'Any secret expired?'           = $secretExpiredAny
        'Earliest secret expiry (date)' = if ($earliestSecret) { Format-DateYMD $earliestSecret } else { $null }
        'Earliest secret expiry (days)' = if ($earliestSecret) { Get-DaysLeft $earliestSecret } else { $null }
        'Certificates (name|start|end|type|thumbprint|expired|daysLeft)' = if ($certSummaries.Count) { $certSummaries -join ';' } else { $null }
        'Any cert expired?'           = $certExpiredAny
        'Earliest cert expiry (date)' = if ($earliestCert) { Format-DateYMD $earliestCert } else { $null }
        'Earliest cert expiry (days)' = if ($earliestCert) { Get-DaysLeft $earliestCert } else { $null }
        'Credentials source (secrets)' = $credentialSourceSecretsString
        'Credentials source (certs)'   = $credentialSourceCertsString
        'App access policy scopes'             = $scopesForApp
        'Excluded from high-priority (scoped)' = $isScopedApp
        'Permissions (application)' = $applicationPermissionsJoined
        'Permissions (delegate)'    = $delegatedPermissionsJoined
        'Valid until (delegate)'    = if ($delegatedValidUntil) { (Get-Date $delegatedValidUntil -Format g) } else { $null }
        'Last sign-in'                    = if ($signInStats.LastSignIn)                 { (Get-Date($signInStats.LastSignIn) -Format g) } else { $null }
        'Last delegate client sign-in'   = if ($signInStats.LastDelegateClientSignIn)   { (Get-Date($signInStats.LastDelegateClientSignIn) -Format g) } else { $null }
        'Last delegate resource sign-in' = if ($signInStats.LastDelegateResourceSignIn) { (Get-Date($signInStats.LastDelegateResourceSignIn) -Format g) } else { $null }
        'Last app client sign-in'        = if ($signInStats.LastAppClientSignIn)        { (Get-Date($signInStats.LastAppClientSignIn) -Format g) } else { $null }
        'Last app resource sign-in'      = if ($signInStats.LastAppResourceSignIn)      { (Get-Date($signInStats.LastAppResourceSignIn) -Format g) } else { $null }
        'Sign-in success count (30 days)' = $summaryStats.SignInSuccessCount
        'Sign-in failure count (30 days)' = $summaryStats.SignInFailureCount
        'Has high-priority permissions (any)'         = $hasHighPriorityAny
        'Has high-priority permissions (application)' = $hasHighPriorityApp
        'Has high-priority permissions (delegated)'   = $hasHighPriorityDel
        'High-priority permissions (application)'     = $highPriorityAppJoined
        'High-priority permissions (delegated)'       = $highPriorityDelegatedJoined
        'High-priority permissions (all)'             = $highPriorityAllJoined
        'New in last 7 days?' = $isNewInLast7Days
        'CustomSecurityAttributes' = $customSecurityAttributesString
    }) | Out-Null
}

$timestamp     = (Get-Date).ToString('yyyy-MM-dd_HH-mm-ss')
$CsvOutputFile = Join-Path -Path $OutputDirectory -ChildPath "${timestamp}_EntraAppInventory.csv"

$sortedRows = $rows | Sort-Object `
    @{ Expression = 'Application Name'; Ascending = $true },
    @{ Expression = 'ApplicationId'    ; Ascending = $true }

$sortedRows | Export-Csv -NoTypeInformation -Path $CsvOutputFile -Encoding UTF8

$rowsForEmail = $rows | Where-Object {
    $objectId      = $_.'ObjectId'
    $applicationId = $_.'ApplicationId'
    $objectIdMatch = ($objectId -and $EmailExceptionsSet.Contains($objectId.ToString().Trim()))
    $appIdMatch    = ($applicationId -and $EmailExceptionsSet.Contains($applicationId.ToString().Trim()))
    -not ($objectIdMatch -or $appIdMatch)
}

$totalCount  = $rowsForEmail.Count
$new7Count   = ($rowsForEmail | Where-Object { $_.'New in last 7 days?' }).Count
$hpAnyCount  = ($rowsForEmail | Where-Object { $_.'Has high-priority permissions (any)' }).Count
$hpNew7Count = ($rowsForEmail | Where-Object { $_.'Has high-priority permissions (any)' -and $_.'New in last 7 days?' }).Count

$TopNewApps = $rowsForEmail |
    Where-Object { $_.'New in last 7 days?' } |
    Sort-Object 'Created on' -Descending |
    Select-Object -First 15 'Application Name', 'ApplicationId', 'Created on', 'Owners'

$TopHighPriorityApps = $rowsForEmail |
    Where-Object { $_.'Has high-priority permissions (any)' } |
    Sort-Object 'Application Name' |
    Select-Object -First 15 'Application Name', 'ApplicationId', 'High-priority permissions (all)', 'Owners'

$reportWindow = "{0:yyyy-MM-dd} to {1:yyyy-MM-dd}" -f $since, (Get-Date)

$style = @'
<style>
  body { font-family: "Segoe UI", Arial, sans-serif; font-size: 12px; color: #24292e; }
  .card { border: 1px solid #e1e4e8; border-radius: 8px; padding: 16px; margin-bottom: 16px; }
  h2 { margin: 0 0 8px 0; font-weight: 600; }
  .meta { color: #57606a; margin: 0 0 12px 0; }
  table { border-collapse: collapse; width: 100%; }
  th, td { border: 1px solid #d0d7de; padding: 8px; text-align: left; vertical-align: top; }
  th { background: #0f6cbd; color: #ffffff; }
  tr:nth-child(even) { background: #f6f8fa; }
  .kpi { font-weight: 600; }
</style>
'@

$summaryTable = @"
<div class='card'>
  <h2>Entra App Summary</h2>
  <p class='meta'>Window: $reportWindow</p>
  <table>
    <tr><th>Metric</th><th>Count</th></tr>
    <tr><td>Total apps scanned</td><td class='kpi'>$totalCount</td></tr>
    <tr><td>New apps (last 7 days)</td><td class='kpi'>$new7Count</td></tr>
    <tr><td>Apps with high-priority permissions</td><td class='kpi'>$hpAnyCount</td></tr>
    <tr><td>New apps with high-priority permissions</td><td class='kpi'>$hpNew7Count</td></tr>
  </table>
</div>
"@

$newAppsTable = @"
<div class='card'>
  <h2>Newest apps (top 15)</h2>
  $(ConvertTo-HtmlTable -InputObjects $TopNewApps)
</div>
"@

$highPriorityAppsTable = @"
<div class='card'>
  <h2>High-priority apps (top 15)</h2>
  $(ConvertTo-HtmlTable -InputObjects $TopHighPriorityApps)
</div>
"@

$htmlBody = $style + $summaryTable + $newAppsTable + $highPriorityAppsTable
$mailSubject = "Entra_APP_Report_$((Get-Date -Format 'yyyy-MMM-dd-ddd hh-mm tt').ToString())"

try {
    if (-not $SkipEmail) {
        if ($PSCmdlet.ShouldProcess(($config.Recipients -join ','), 'Send summary email')) {
            Send-MailMessage `
                -From $config.MailFrom `
                -To $config.Recipients `
                -Subject $mailSubject `
                -Attachments $CsvOutputFile `
                -SmtpServer $config.SmtpServer `
                -Body $htmlBody `
                -BodyAsHtml `
                -WarningAction SilentlyContinue
        }
    }
}
finally {
    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
}

Write-Host "CSV export completed: $CsvOutputFile" -ForegroundColor Green
if ($SkipEmail) {
    Write-Host 'Email summary skipped.' -ForegroundColor Yellow
}
else {
    Write-Host 'Email summary completed.' -ForegroundColor Green
}
