#region --- Config ---

$ThrottleSeconds = 2
$ApiVersion = "2016-11-01"

#endregion

#region --- Auth ---

$TenantId     = "<TENANT_ID>"
$ClientId     = "<CLIENT_ID>"
$ClientSecret = "<CLIENT_SECRET>"

Write-Output "Authenticating (PowerApps cmdlets)..."

Import-Module Microsoft.PowerApps.PowerShell -ErrorAction Stop
Import-Module Microsoft.PowerApps.Administration.PowerShell -ErrorAction Stop

Add-PowerAppsAccount `
    -TenantID $TenantId `
    -ApplicationId $ClientId `
    -ClientSecret $ClientSecret

Write-Output "Authenticating (Az token for REST)..."

Import-Module Az.Accounts -ErrorAction Stop

$secureSecret = ConvertTo-SecureString $ClientSecret -AsPlainText -Force
$cred = New-Object System.Management.Automation.PSCredential ($ClientId, $secureSecret)

Connect-AzAccount `
    -ServicePrincipal `
    -Tenant $TenantId `
    -Credential $cred `
    -ErrorAction Stop

$AccessToken = (Get-AzAccessToken -ResourceUrl "https://service.flow.microsoft.com/").Token

if (-not $AccessToken) {
    throw "Failed to acquire Power Automate access token"
}

$Headers = @{
    Authorization  = "Bearer $AccessToken"
    "Content-Type" = "application/json"
}

#endregion

#region --- Get environments ---

Write-Output "Fetching Power Platform environments..."

$environments = Get-AdminPowerAppEnvironment

if (-not $environments -or $environments.Count -eq 0) {
    Write-Output "No environments found. Exiting ❌"
    return
}

Write-Output "Found $($environments.Count) environment(s)"

#endregion

#region --- Processing ---

foreach ($env in $environments) {

    $envName = $env.EnvironmentName

    if ([string]::IsNullOrWhiteSpace($envName)) {
        continue
    }

    Write-Output "============================================="
    Write-Output "Processing Environment: $envName"

    try {
        $riskyFlows = Get-AdminFlowAtRiskOfSuspension -EnvironmentName $envName
    }
    catch {
        Write-Output "Failed to fetch risky flows for environment $envName"
        Write-Output $_.Exception.Message
        continue
    }

    if (-not $riskyFlows -or $riskyFlows.Count -eq 0) {
        Write-Output "No risky flows in environment $envName ✅"
        continue
    }

    Write-Output "Found $($riskyFlows.Count) risky flow(s)"

    foreach ($flow in $riskyFlows) {

        $flowName = $flow.FlowName
        $dispName = $flow.DisplayName

        Write-Output "---------------------------------------------"
        Write-Output "Processing flow: $dispName"

        if ([string]::IsNullOrWhiteSpace($flowName)) {
            Write-Output "Missing FlowName. Skipping ❌"
            continue
        }

        $flowUri = "https://api.flow.microsoft.com/providers/Microsoft.ProcessSimple/environments/$envName/flows/$flowName?api-version=$ApiVersion"

        try {
            # 1️⃣ GET flow definition
            $flowObj = Invoke-RestMethod `
                -Method GET `
                -Uri $flowUri `
                -Headers $Headers `
                -ErrorAction Stop

            if (-not $flowObj.properties -or -not $flowObj.properties.definition) {
                Write-Output "Definition missing. Skipping ❌"
                continue
            }

            if ($flowObj.properties.state -eq "Stopped") {
                Write-Output "Flow is stopped. Skipping ⚠️"
                continue
            }

            # 2️⃣ PATCH same definition back (re-save)
            $patchBody = @{
                properties = @{
                    definition = $flowObj.properties.definition
                }
            }

            Invoke-RestMethod `
                -Method PATCH `
                -Uri $flowUri `
                -Headers $Headers `
                -Body ($patchBody | ConvertTo-Json -Depth 100) `
                -ErrorAction Stop

            Write-Output "Re-save successful ✅ (risk timer reset)"

        }
        catch {
            Write-Output "FAILED for flow: $dispName"
            Write-Output $_.Exception.Message
        }

        Start-Sleep -Seconds $ThrottleSeconds
    }
}

#endregion

Write-Output "============================================="
Write-Output "Run completed 🎯"