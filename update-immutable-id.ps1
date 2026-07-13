 # Define required Microsoft Graph scopes
$RequiredScopes = @('User.Read.All', 'User.ReadWrite.All')
$IntermediateDomain = "{your-domain}.onmicrosoft.com"

Write-Host "Starting Azure AD User Update Process..."
Write-Host "Connecting to Microsoft Graph with scopes: $($RequiredScopes -join ', ')..."

# region Connect to Microsoft Graph
try {
    Connect-MgGraph -Scopes $RequiredScopes -UseDeviceCode -ErrorAction Stop
    Write-Host "Successfully connected to Microsoft Graph."
}
catch {
    Write-Error "Failed to connect to Microsoft Graph. Please ensure you have the necessary permissions and the Microsoft.Graph module is installed. Error: $($_.Exception.Message)"
    exit 1 # Exit script on connection failure
}
# endregion

# region Request User Inputs
Write-Host "`n--- Please provide the required user details ---"

$InitialUserId = Read-Host -Prompt "Enter the current UserPrincipalName of the user to update (e.g., user@yourdomain.com)"
# Extract username part for constructing intermediate UPN
if ($InitialUserId -like "*@*") {
    $UsernamePart = $InitialUserId.Split('@')[0]
} else {
    # If ObjectId is entered, we'll try to get UPN later to derive username
    Write-Warning "User ID provided is not a UPN. Assuming it's an ObjectId. UPN will be resolved to derive intermediate UPN."
    $UsernamePart = $null # Will attempt to resolve later
}

$NewImmutableId = Read-Host -Prompt "Enter the NEW On-Premises Immutable ID for the user (e.g., 'newuser@yourdomain.com' or a GUID). USE WITH EXTREME CAUTION"
$FinalUpn = Read-Host -Prompt "Enter the FINAL desired UserPrincipalName for the user (e.g., newuser@yourdomain.com)"

Write-Host "----------------------------------------------`n"
# endregion

# region Retrieve Initial User Information and Derive Intermediate UPN
Write-Host "Attempting to retrieve user with identifier: '$InitialUserId'..."
$user = $null
try {
    # Using -Property to explicitly request OnPremisesImmutableId
    $user = Get-MgUser -UserId $InitialUserId -Property Id,UserPrincipalName,DisplayName,OnPremisesImmutableId,Mail -ErrorAction Stop |
            Select-Object Id, UserPrincipalName, DisplayName, OnPremisesImmutableId, Mail

    if (-not $user) {
        Write-Error "User with identifier '$InitialUserId' not found in Azure AD. Please verify the input."
        Disconnect-MgGraph
        exit 1
    }

    # If UsernamePart was not derived from InitialUserId (e.g., if ObjectId was given)
    if (-not $UsernamePart) {
        $UsernamePart = $user.UserPrincipalName.Split('@')[0]
        if (-not $UsernamePart) {
            Write-Error "Could not derive username part from user's UPN ($($user.UserPrincipalName)). Cannot construct intermediate UPN."
            Disconnect-MgGraph
            exit 1
        }
    }

    $IntermediateUpn = "$UsernamePart@$IntermediateDomain"

    Write-Host "Found user: " -NoNewline
    Write-Host "$($user.DisplayName)" -ForegroundColor Yellow
    Write-Host "  Current UPN: " -NoNewline
    Write-Host "$($user.UserPrincipalName)" -ForegroundColor Yellow
    Write-Host "  Azure AD ObjectId: " -NoNewline
    Write-Host "$($user.Id)" -ForegroundColor Yellow
    Write-Host "  Current On-Premises Immutable ID: " -NoNewline
    Write-Host "$($user.OnPremisesImmutableId -replace '^$','<Not Set>')" -ForegroundColor Yellow # Handle null/empty for display

    Write-Host "`nPlanned Changes:"
    Write-Host "  Initial UPN: " -NoNewline
    Write-Host "$($user.UserPrincipalName)" -ForegroundColor Yellow
    Write-Host "  Step 1 (Intermediate) UPN: " -NoNewline
    Write-Host "$($IntermediateUpn)" -ForegroundColor Yellow
    Write-Host "  Step 2 (Part 2) Final UPN: " -NoNewline
    Write-Host "$($FinalUpn)" -ForegroundColor Yellow
    Write-Host "  Step 2 (Part 1) New Immutable ID: " -NoNewline
    Write-Host "$($NewImmutableId)" -ForegroundColor Yellow

}
catch {
    Write-Error "An error occurred while retrieving user '$InitialUserId' or deriving UPNs: $($_.Exception.Message)"
    Disconnect-MgGraph
    exit 1
}
# endregion

# region Confirmation Prompt for All Steps
# Updated warning message to include SSO impact
Write-Host "`nWARNING: CHANGING THE IMMUTABLE ID CAN BREAK SINGLE SIGN ON FOR THE USER!" -ForegroundColor Red
Write-Host "WARNING: ENSURE YOU UNDERSTAND THE IMPLICATIONS BEFORE PROCEEDING." -ForegroundColor Red
Write-Host ""

$confirm = Read-Host "Are you absolutely sure you want to proceed with this multi-step update? Type 'YES' to confirm:"
if ($confirm -ne 'YES') {
    Write-Host "Operation cancelled by user."
    Disconnect-MgGraph
    exit 0 # Exit gracefully
}
# endregion

# region Step 1: Change UPN to Intermediate .onmicrosoft.com
Write-Host "`n--- Starting Step 1: Changing UPN to '$IntermediateUpn' ---"
try {
    $step1UpdateParams = @{
        UserPrincipalName = $IntermediateUpn
    }
    Update-MgUser -UserId $user.Id -BodyParameter $step1UpdateParams -ErrorAction Stop
    Write-Host "Successfully initiated Step 1 UPN change for '$($user.DisplayName)'."
    Write-Host "Waiting 10 seconds for UPN propagation before proceeding..." -ForegroundColor DarkYellow
    Start-Sleep -Seconds 10 # Allow time for UPN to propagate

    # Using -Property to explicitly request OnPremisesImmutableId
    $user = Get-MgUser -UserId $user.Id -Property Id,UserPrincipalName,DisplayName,OnPremisesImmutableId,Mail -ErrorAction Stop | Select-Object Id, UserPrincipalName, DisplayName, OnPremisesImmutableId, Mail
    Write-Host "Current UPN after Step 1: $($user.UserPrincipalName)"
    if ($user.UserPrincipalName -ne $IntermediateUpn) {
        Write-Host "WARNING: UPN did not immediately update to '$IntermediateUpn' after Step 1. Proceeding anyway, but be aware." -ForegroundColor DarkYellow
    }
}
catch {
    Write-Error "Failed during Step 1 (UPN change to '$IntermediateUpn') for user '$($user.DisplayName)'. Error: $($_.Exception.Message)"
    if ($_.Exception.InnerException) {
        Write-Error "Inner Exception: $($_.Exception.InnerException.Message)"
    }
    Disconnect-MgGraph
    exit 1
}
# endregion

# region Step 2 (Part 1): Update Immutable ID
Write-Host "`n--- Starting Step 2 (Part 1): Setting On-Premises Immutable ID to '$NewImmutableId' ---"
try {
    $step2Part1UpdateParams = @{
        OnPremisesImmutableId = $NewImmutableId
    }
    Update-MgUser -UserId $user.Id -BodyParameter $step2Part1UpdateParams -ErrorAction Stop
    Write-Host "Successfully initiated Immutable ID update for '$($user.DisplayName)'."
    Write-Host "Waiting 10 seconds to allow for Immutable ID propagation in Entra ID..." -ForegroundColor DarkYellow
    Start-Sleep -Seconds 10

    # Re-fetch user to confirm Immutable ID has changed on Entra
    Write-Host "Verifying Immutable ID update on Entra ID..."
    # Using -Property to explicitly request OnPremisesImmutableId
    $userAfterImmutableIdUpdate = Get-MgUser -UserId $user.Id -Property Id,UserPrincipalName,DisplayName,OnPremisesImmutableId,Mail -ErrorAction Stop | Select-Object Id, UserPrincipalName, DisplayName, OnPremisesImmutableId, Mail

    Write-Host "Current On-Premises Immutable ID on Entra ID: $($userAfterImmutableIdUpdate.OnPremisesImmutableId)"

    if ($userAfterImmutableIdUpdate.OnPremisesImmutableId -eq $NewImmutableId) {
        Write-Host "SUCCESS: Immutable ID verified to be updated to '$NewImmutableId' on Entra ID."
        # Update the $user variable for subsequent steps with the latest data
        $user = $userAfterImmutableIdUpdate
    } else {
        Write-Host "WARNING: Immutable ID on Entra ID is still '$($userAfterImmutableIdUpdate.OnPremisesImmutableId)' which does not match '$NewImmutableId' after 10 seconds. This might be due to propagation delays or an issue. Proceeding to next step, but manual verification is advised." -ForegroundColor DarkYellow
    }
}
catch {
    Write-Error "Failed during Step 2 (Part 1: Immutable ID update) for user '$($user.DisplayName)'. Error: $($_.Exception.Message)"
    if ($_.Exception.InnerException) {
        Write-Error "Inner Exception: $($_.Exception.InnerException.Message)"
    }
    Disconnect-MgGraph
    exit 1
}
# endregion

# region Step 2 (Part 2): Change UPN to Final UPN
Write-Host "`n--- Starting Step 2 (Part 2): Changing UPN to final '$FinalUpn' ---"
try {
    $step2Part2UpdateParams = @{
        UserPrincipalName = $FinalUpn
    }
    Update-MgUser -UserId $user.Id -BodyParameter $step2Part2UpdateParams -ErrorAction Stop
    Write-Host "Successfully initiated final UPN change for '$($user.DisplayName)'."
    Write-Host "Waiting 10 seconds for final UPN propagation..." -ForegroundColor DarkYellow
    Start-Sleep -Seconds 10

    # Final re-fetch to confirm all changes
    # Using -Property to explicitly request OnPremisesImmutableId
    $finalUser = Get-MgUser -UserId $user.Id -Property UserPrincipalName,OnPremisesImmutableId -ErrorAction Stop | Select-Object UserPrincipalName, OnPremisesImmutableId
    Write-Host "Final UPN (verified): $($finalUser.UserPrincipalName)" -ForegroundColor Green
    Write-Host "Final On-Premises Immutable ID (verified): $($finalUser.OnPremisesImmutableId)" -ForegroundColor Green

    if ($finalUser.UserPrincipalName -eq $FinalUpn -and $finalUser.OnPremisesImmutableId -eq $NewImmutableId) {
        Write-Host "SUCCESS: User '$($user.DisplayName)' UPN and Immutable ID updated as specified." -ForegroundColor Green
        Write-Host "All changes should now be reflected in Azure AD (subject to propagation)." -ForegroundColor Green
    } else {
        Write-Host "WARNING: Verification shows UPN or Immutable ID might not have fully propagated yet. Check Azure AD manually." -ForegroundColor DarkYellow
    }
}
catch {
    Write-Error "Failed during Step 2 (Part 2: Final UPN change) for user '$($user.DisplayName)'. Error: $($_.Exception.Message)"
    if ($_.Exception.InnerException) {
        Write-Error "Inner Exception: $($_.Exception.InnerException.Message)"
    }
    Disconnect-MgGraph
    exit 1
}
# endregion

# region Disconnect from Microsoft Graph
finally {
    # Disconnect-MgGraph is now silent
    Write-Host "`nDisconnecting from Microsoft Graph."
    Disconnect-MgGraph | Out-Null # Redirect output to null to make it silent
}
# endregion

Write-Host "`nScript execution completed." 
