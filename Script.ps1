# Connect to Microsoft Graph
Connect-MgGraph -Scopes User.ReadWrite.All, Group.ReadWrite.All, Device.Read.All

# Connect to Exchange Online
try {
    Connect-ExchangeOnline -ErrorAction Stop
    $exoAvailable = $true
    Write-Host "Connected to Exchange Online"
} catch {
    $exoAvailable = $false
    Write-Host "Exchange Online not available. Only Graph-supported groups will be processed."
}

# Input CSV file structure: nameofgroup,name2
$csvPath = "groups.csv"

# Logging setup
$logPath = "group_script_log_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
function Log-Action($message) {
    Write-Host $message
    $message | Out-File -FilePath $logPath -Append
}

# License SKUs to filter humans only
$humanLicenses = @("STANDARDPACK", "ENTERPRISEPACK", "EMS")

# Validate CSV
if (-not (Test-Path $csvPath)) {
    Log-Action "CSV file not found at path: $csvPath"
    exit
}

$csvData = Import-Csv -Path $csvPath
if ($csvData.Count -eq 0) {
    Log-Action "CSV file is empty or improperly formatted."
    exit
}

foreach ($row in $csvData) {
    $sourceGroupName = $row.nameofgroup
    $newGroupName = $row.name2

    Log-Action "\n--- START PROCESSING GROUP: $sourceGroupName ---"

    # Fetch source group
    $sourceGroup = Get-MgGroup -All | Where-Object { $_.DisplayName -eq $sourceGroupName } | Select-Object -First 1

    if (!$sourceGroup) {
        Log-Action "Group '$sourceGroupName' not found. Skipping."
        continue
    }

    Log-Action "Group '$sourceGroupName' found. ID: $($sourceGroup.Id)"

    # Attempt to get members via Graph
    $groupMembers = Get-MgGroupMember -GroupId $sourceGroup.Id -All | Where-Object { $_.'@odata.type' -eq "#microsoft.graph.user" }

    # Fallback via Exchange if necessary
    if ($groupMembers.Count -eq 0 -and $exoAvailable -and $sourceGroup.Mail) {
        try {
            Log-Action "Attempting Exchange fallback for mail-enabled group '$($sourceGroup.Mail)'"
            $exchangeMembers = Get-DistributionGroupMember -Identity $sourceGroup.Mail -ResultSize Unlimited | Where-Object { $_.RecipientType -eq "UserMailbox" -or $_.RecipientType -eq "MailUser" }
            $groupMembers = foreach ($m in $exchangeMembers) {
                Get-MgUser -UserId $m.PrimarySmtpAddress -ErrorAction SilentlyContinue
            }
            Log-Action "Exchange fallback retrieved $($groupMembers.Count) users"
        } catch {
            Log-Action "Exchange fallback failed: $_"
        }
    } else {
        Log-Action "Found $($groupMembers.Count) users via Graph"
    }

    # Filter members by license
    $filteredUsers = @()
    foreach ($member in $groupMembers) {
        try {
            $userLicenses = Get-MgUserLicenseDetail -UserId $member.Id
            $licenseParts = $userLicenses | ForEach-Object { (Get-MgSubscribedSku | Where-Object { $_.SkuId -eq $_.SkuId }).SkuPartNumber }

            if ($licenseParts | Where-Object { $humanLicenses -contains $_ }) {
                $filteredUsers += $member
            }
        } catch {
            Log-Action "Error retrieving license info for user ID: $($member.Id)"
        }
    }

    Log-Action "Filtered users with valid licenses: $($filteredUsers.Count)"

    # Create user group
    $newUserGroup = New-MgGroup -DisplayName $newGroupName -MailEnabled:$false -MailNickname $newGroupName -SecurityEnabled:$true
    Log-Action "Created user group: '$newGroupName' (ID: $($newUserGroup.Id))"

    # Add users to new group
    foreach ($user in $filteredUsers) {
        try {
            New-MgGroupMember -GroupId $newUserGroup.Id -DirectoryObjectId $user.Id
            Log-Action "Added user ID: $($user.Id) to group '$newGroupName'"
        } catch {
            Log-Action "Failed to add user ID: $($user.Id)"
        }
    }

    # Create devices group
    $deviceGroupName = "${newGroupName}_devices"
    $newDeviceGroup = New-MgGroup -DisplayName $deviceGroupName -MailEnabled:$false -MailNickname $deviceGroupName -SecurityEnabled:$true
    Log-Action "Created device group: '$deviceGroupName' (ID: $($newDeviceGroup.Id))"

    $deviceCount = 0
    # Add devices belonging to filtered users (fallback-aware)
    foreach ($user in $filteredUsers) {
        try {
            $devices = Get-MgUserOwnedDevice -UserId $user.Id -All | Where-Object { $_.AccountEnabled -eq $true }
            if ($devices.Count -eq 0) {
                $devices = Get-MgDevice -All | Where-Object {
                    $_.RegisteredOwners -ne $null -and
                    $_.AccountEnabled -eq $true -and
                    $_.RegisteredOwners.Id -contains $user.Id
                }
                Log-Action "Fallback method found $($devices.Count) device(s) for user ID: $($user.Id)"
            }
            foreach ($device in $devices) {
                try {
                    New-MgGroupMember -GroupId $newDeviceGroup.Id -DirectoryObjectId $device.Id
                    Log-Action "Added device '$($device.DisplayName)' (ID: $($device.Id))"
                    $deviceCount++
                } catch {
                    Log-Action "Failed to add device ID: $($device.Id)"
                }
            }
        } catch {
            Log-Action "Error retrieving devices for user ID: $($user.Id)"
        }
    }

    Log-Action "Total devices added: $deviceCount"
    Log-Action "--- END PROCESSING GROUP: $sourceGroupName ---\n"
}

Log-Action "\nScript execution completed."
