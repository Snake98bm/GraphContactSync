<#
.SYNOPSIS
    Sync contacts from the GAL to 1:many mailboxes.
    Version adapted for Azure Automation Runbook.

.DESCRIPTION
    Original version: https://github.com/tardispilot/GraphContactSync
    Changes for Azure Automation:
    - Removed PoShLog, replaced with Write-Output/Write-Warning/Write-Error
    - Removed Set-ExecutionPolicy (not needed in Azure Automation)
    - Authentication via Get-AutomationCertificate (Automation Account asset)
    - Certificate password is read from Get-AutomationVariable (secure variable)
    - Photos are saved to $env:TEMP (temporary file system available in the sandbox)
    - Sensitive parameters (ClientID, ExchangeOrg) can be passed as
      Automation Variables or directly as Runbook parameters

.PARAMETER ExchangeOrg
    The .onmicrosoft.com tenant of the Exchange organisation.
    If empty, read from the Automation Variable "GraphContactSync-ExchangeOrg".

.PARAMETER ClientID
    The Client ID of the Azure AD app registration.
    If empty, read from the Automation Variable "GraphContactSync-ClientID".

.PARAMETER CertificateAssetName
    Name of the certificate uploaded to the Automation Account Assets.
    Default: "GraphContactSyncCert"

.PARAMETER MailboxList
    Comma-separated list of mailboxes, or "DIRECTORY" for all.

.PARAMETER ManagedContactFolderName
    Name of the contact folder to manage in each mailbox.

.PARAMETER FileAsFormat
    Format for the FileAs field: "FirstLast" (default) or "LastFirst".

.PARAMETER Categories
    Array of categories to assign to contacts.

.NOTES
    Prerequisites in the Automation Account:
    1. Microsoft.Graph.Users module imported (or at least the required sub-modules)
    2. Microsoft.Graph.PersonalContacts module imported
    3. Certificate uploaded in Assets > Certificates with the name specified in CertificateAssetName
    4. (Optional) Automation Variables:
       - "GraphContactSync-ExchangeOrg" (encryption recommended)
       - "GraphContactSync-ClientID"
#>

param(
    [Parameter(Mandatory = $false)][string]$ExchangeOrg = "",
    [Parameter(Mandatory = $false)][string]$ClientID = "",
    [Parameter(Mandatory = $false)][string]$CertificateAssetName = "GraphContactSyncCert",
    [Parameter(Mandatory = $true)][string]$MailboxList,
    [Parameter(Mandatory = $true)][string]$ManagedContactFolderName,
    [Parameter(Mandatory = $false)][ValidateSet("FirstLast", "LastFirst")][string]$FileAsFormat = "FirstLast",
    [Parameter(Mandatory = $false)][string[]]$Categories = @()
)

# ---------------------------------------------------------------------------
# Simplified logging functions (replacement for PoShLog)
# In Azure Automation output goes to Job Streams visible in the portal
# ---------------------------------------------------------------------------
function Write-InfoLog    { param([string]$msg) Write-Output  "[INFO]    $msg" }
function Write-VerboseLog { param([string]$msg) Write-Verbose "[VERBOSE] $msg" }
function Write-DebugLog   { param([string]$msg) Write-Output  "[DEBUG]   $msg" }
function Write-WarningLog { param([string]$msg) Write-Warning "[WARN]    $msg" }
function Write-ErrorLog   { param([string]$msg) Write-Error   "[ERROR]   $msg" }

# ---------------------------------------------------------------------------
# Read parameters from Automation Variables if not passed explicitly
# ---------------------------------------------------------------------------
if ([string]::IsNullOrEmpty($ExchangeOrg)) {
    try {
        $ExchangeOrg = Get-AutomationVariable -Name "GraphContactSync-ExchangeOrg"
        Write-InfoLog "ExchangeOrg read from Automation Variable."
    } catch {
        throw "ExchangeOrg was not specified as a parameter and the variable 'GraphContactSync-ExchangeOrg' does not exist in the Automation Account."
    }
}

if ([string]::IsNullOrEmpty($ClientID)) {
    try {
        $ClientID = Get-AutomationVariable -Name "GraphContactSync-ClientID"
        Write-InfoLog "ClientID read from Automation Variable."
    } catch {
        throw "ClientID was not specified as a parameter and the variable 'GraphContactSync-ClientID' does not exist in the Automation Account."
    }
}

# ---------------------------------------------------------------------------
# Temporary folder for photos (in Azure Automation the file system is
# ephemeral: it only exists for the duration of the job execution)
# ---------------------------------------------------------------------------
$PhotosDir = Join-Path $env:TEMP "GraphContactSync_Photos"
if (!(Test-Path -Path $PhotosDir)) {
    New-Item -ItemType Directory -Path $PhotosDir -Force | Out-Null
    Write-InfoLog "Temporary photo folder created: $PhotosDir"
}

# ---------------------------------------------------------------------------
# Main contact synchronisation function
# (logic unchanged from the original, only the photo path has been updated)
# ---------------------------------------------------------------------------
function Sync-ManagedContacts {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)][string]$Mailbox,
        [Parameter(Mandatory = $true)][string]$ManagedContactFolderName,
        [Parameter(Mandatory = $true)]$ManagedContacts,
        [Parameter(Mandatory = $false)][ValidateSet("FirstLast", "LastFirst")][string]$FileAsFormat = "FirstLast",
        [Parameter(Mandatory = $false)][string[]]$Categories = @()
    )

    Write-InfoLog "[$Mailbox] Locating Managed Contact folder"
    $ManagedContactFolder = Get-MgUserContactFolder -UserId $Mailbox -Filter "DisplayName eq '$ManagedContactFolderName'"

    if ($null -eq $ManagedContactFolder) {
        Write-InfoLog "[$Mailbox] Creating Managed Contact folder ($ManagedContactFolderName)"
        $ManagedContactFolder = New-MgUserContactFolder -UserId $Mailbox -DisplayName $ManagedContactFolderName
    }

    $ExistingManagedContacts = Get-MgUserContactFolderContact `
        -UserId $Mailbox `
        -ContactFolderId $ManagedContactFolder.Id `
        -All `
        -ExpandProperty "extensions(`$filter=id eq 'ManagedContactCorrelation'`)"

    $ContactsToAdd    = @()
    $ContactsToDelete = @()
    $ContactsToAdd    += $ManagedContacts | Where-Object { $ExistingManagedContacts.Extensions.AdditionalProperties.CorrelationId -notcontains $_.Id }
    $ContactsToDelete += $ExistingManagedContacts | Where-Object { $ManagedContacts.Id -notcontains $_.Extensions.AdditionalProperties.CorrelationId }
    $ContactsToChecksum = $ExistingManagedContacts | Where-Object { $ManagedContacts.Id -contains $_.Extensions.AdditionalProperties.CorrelationId }

    $md5  = New-Object -TypeName System.Security.Cryptography.MD5CryptoServiceProvider
    $utf8 = New-Object -TypeName System.Text.UTF8Encoding

    foreach ($ExistingContact in $ContactsToChecksum) {
        $ManagedContact = $ManagedContacts | Where-Object { $_.Id -eq $ExistingContact.Extensions.AdditionalProperties.CorrelationId }

        $ManagedContactChecksumFields = ""
        $ExistingContactChecksumFields = ""

        if ($ManagedContact.EntryType -eq 'User') {
            $ManagedContactChecksumFields = ($ManagedContact | Select-Object -Property `
                @{Name='DisplayName';     Expression={$_.DisplayName     ?? ""}}, `
                @{Name='GivenName';       Expression={$_.GivenName       ?? ""}}, `
                @{Name='Surname';         Expression={$_.Surname         ?? ""}}, `
                @{Name='CompanyName';     Expression={$_.CompanyName     ?? ""}}, `
                @{Name='JobTitle';        Expression={$_.JobTitle        ?? ""}}, `
                @{Name='Department';      Expression={$_.Department      ?? ""}}, `
                @{Name='OfficeLocation';  Expression={$_.OfficeLocation  ?? ""}}, `
                @{Name='Mail';            Expression={$_.Mail            ?? ""}}, `
                @{Name='BusinessPhones';  Expression={$_.BusinessPhones  ?? ""}}, `
                @{Name='MobilePhone';     Expression={$_.MobilePhone     ?? ""}}, `
                @{Name='StreetAddress';   Expression={$_.StreetAddress   ?? ""}}, `
                @{Name='City';            Expression={$_.City            ?? ""}}, `
                @{Name='State';           Expression={$_.State           ?? ""}}, `
                @{Name='PostalCode';      Expression={$_.PostalCode      ?? ""}}, `
                @{Name='Country';         Expression={$_.Country         ?? ""}} `
            | ConvertTo-Json -Depth 10)

            $ExistingContactChecksumFields = ($ExistingContact | Select-Object -Property `
                @{Name='DisplayName';    Expression={$_.DisplayName                      ?? ""}}, `
                @{Name='GivenName';      Expression={$_.GivenName                        ?? ""}}, `
                @{Name='Surname';        Expression={$_.Surname                          ?? ""}}, `
                @{Name='CompanyName';    Expression={$_.CompanyName                      ?? ""}}, `
                @{Name='JobTitle';       Expression={$_.JobTitle                         ?? ""}}, `
                @{Name='Department';     Expression={$_.Department                       ?? ""}}, `
                @{Name='OfficeLocation'; Expression={$_.OfficeLocation                   ?? ""}}, `
                @{Name='Mail';           Expression={$_.EmailAddresses[0].Address        ?? ""}}, `
                @{Name='BusinessPhones'; Expression={$_.BusinessPhones                   ?? ""}}, `
                @{Name='MobilePhone';    Expression={$_.MobilePhone                      ?? ""}}, `
                @{Name='StreetAddress';  Expression={$_.BusinessAddress.Street           ?? ""}}, `
                @{Name='City';           Expression={$_.BusinessAddress.City             ?? ""}}, `
                @{Name='State';          Expression={$_.BusinessAddress.State            ?? ""}}, `
                @{Name='PostalCode';     Expression={$_.BusinessAddress.PostalCode       ?? ""}}, `
                @{Name='Country';        Expression={$_.BusinessAddress.CountryOrRegion  ?? ""}} `
            | ConvertTo-Json -Depth 10)
        }
        elseif ($ManagedContact.EntryType -eq 'Contact') {
            $ManagedContactChecksumFields = ($ManagedContact | Select-Object -Property `
                @{Name='DisplayName';    Expression={$_.DisplayName           ?? ""}}, `
                @{Name='GivenName';      Expression={$_.GivenName             ?? ""}}, `
                @{Name='Surname';        Expression={$_.Surname               ?? ""}}, `
                @{Name='CompanyName';    Expression={$_.CompanyName           ?? ""}}, `
                @{Name='JobTitle';       Expression={$_.JobTitle              ?? ""}}, `
                @{Name='Mail';           Expression={$_.Mail                  ?? ""}}, `
                @{Name='Mobile';         Expression={$_.Phones[1].Number      ?? ""}}, `
                @{Name='BusinessPhone';  Expression={$_.Phones[2].Number      ?? ""}}, `
                @{Name='StreetAddress';  Expression={$_.Addresses[0].Street   ?? ""}}, `
                @{Name='City';           Expression={$_.Addresses[0].City     ?? ""}}, `
                @{Name='State';          Expression={$_.Addresses[0].State    ?? ""}}, `
                @{Name='PostalCode';     Expression={$_.Addresses[0].PostalCode ?? ""}}, `
                @{Name='Country';        Expression={$_.Addresses[0].Country  ?? ""}} `
            | ConvertTo-Json -Depth 10)

            $ExistingContactChecksumFields = ($ExistingContact | Select-Object -Property `
                @{Name='DisplayName';    Expression={$_.DisplayName                      ?? ""}}, `
                @{Name='GivenName';      Expression={$_.GivenName                        ?? ""}}, `
                @{Name='Surname';        Expression={$_.Surname                          ?? ""}}, `
                @{Name='CompanyName';    Expression={$_.CompanyName                      ?? ""}}, `
                @{Name='JobTitle';       Expression={$_.JobTitle                         ?? ""}}, `
                @{Name='Mail';           Expression={$_.EmailAddresses[0].Address        ?? ""}}, `
                @{Name='Mobile';         Expression={$_.MobilePhone                      ?? ""}}, `
                @{Name='BusinessPhone';  Expression={$_.BusinessPhones[0]               ?? ""}}, `
                @{Name='StreetAddress';  Expression={$_.BusinessAddress.Street           ?? ""}}, `
                @{Name='City';           Expression={$_.BusinessAddress.City             ?? ""}}, `
                @{Name='State';          Expression={$_.BusinessAddress.State            ?? ""}}, `
                @{Name='PostalCode';     Expression={$_.BusinessAddress.PostalCode       ?? ""}}, `
                @{Name='Country';        Expression={$_.BusinessAddress.CountryOrRegion  ?? ""}} `
            | ConvertTo-Json -Depth 10)
        }

        $ManagedContactChecksum  = [System.BitConverter]::ToString($md5.ComputeHash($utf8.GetBytes($ManagedContactChecksumFields)))
        $ExistingContactChecksum = [System.BitConverter]::ToString($md5.ComputeHash($utf8.GetBytes($ExistingContactChecksumFields)))

        $PhotoChanged = $false
        if ($ManagedContact.EntryType -eq 'User') {
            $CurrentPhotoMeta = Get-MgUserPhoto -UserId $ManagedContact.UserPrincipalName -ProfilePhotoId 120x120 -ErrorAction SilentlyContinue
            $CurrentPhotoChecksum = ""
            if ($CurrentPhotoMeta) {
                $PhotoFingerprint = $CurrentPhotoMeta.AdditionalProperties["@odata.mediaEtag"]
                if (-not $PhotoFingerprint) {
                    $PhotoFingerprint = "$($CurrentPhotoMeta.Id)_$($CurrentPhotoMeta.Height)x$($CurrentPhotoMeta.Width)"
                }
                $CurrentPhotoChecksum = [System.BitConverter]::ToString($md5.ComputeHash($utf8.GetBytes($PhotoFingerprint)))
            }
            $StoredPhotoChecksum = $ExistingContact.Extensions.AdditionalProperties.PhotoChecksum ?? ""
            if ($CurrentPhotoChecksum -ne $StoredPhotoChecksum) {
                $PhotoChanged = $true
                Write-VerboseLog "[$Mailbox] $($ManagedContact.DisplayName) photo changed. Old:[$StoredPhotoChecksum] New:[$CurrentPhotoChecksum]"
            }
        }

        if ($ExistingContactChecksum -ne $ManagedContactChecksum -or $PhotoChanged) {
            Write-DebugLog "[$Mailbox] $($ManagedContact.DisplayName) contact changed. Marking for update."
            $ContactsToDelete += $ExistingContact
            $ContactsToAdd    += $ManagedContact
        }
    }

    foreach ($Contact in $ContactsToDelete) {
        Write-VerboseLog "[$Mailbox] Deleting contact: $($Contact.DisplayName)"
        Remove-MgUserContactFolderContact -UserId $Mailbox -ContactFolderId $ManagedContactFolder.Id -ContactId $Contact.Id
    }

    foreach ($Contact in $ContactsToAdd) {
        $ContactPhotoFile = $null
        $PhotoChecksum    = ""

        if ($Contact.EntryType -eq 'User') {
            $PhotoMeta = Get-MgUserPhoto -UserId $Contact.UserPrincipalName -ProfilePhotoId 120x120 -ErrorAction SilentlyContinue
            if ($PhotoMeta) {
                $PhotoFingerprint = $PhotoMeta.AdditionalProperties["@odata.mediaEtag"]
                if (-not $PhotoFingerprint) {
                    $PhotoFingerprint = "$($PhotoMeta.Id)_$($PhotoMeta.Height)x$($PhotoMeta.Width)"
                }
                $PhotoChecksum = [System.BitConverter]::ToString($md5.ComputeHash($utf8.GetBytes($PhotoFingerprint)))
            }

            # Use $PhotosDir instead of a relative "Photos\" path
            $PhotoFilePath = Join-Path $PhotosDir "$($Contact.UserPrincipalName)-$PhotoChecksum.jpg"
            if (!(Test-Path -PathType Leaf -Path $PhotoFilePath)) {
                Write-VerboseLog "Downloading photo for contact: $($Contact.DisplayName)"
                Get-MgUserPhotoContent -UserId $Contact.UserPrincipalName -ProfilePhotoId 120x120 -OutFile $PhotoFilePath -ErrorAction SilentlyContinue
            }
            if (Test-Path -PathType Leaf -Path $PhotoFilePath) {
                $ContactPhotoFile = $PhotoFilePath
            }
        }

        if ($Contact.EntryType -eq 'User') {
            $ManagedContactString = ($Contact | Select-Object -Property DisplayName, GivenName, Surname, CompanyName, JobTitle, Department, OfficeLocation, Mail, BusinessPhones, MobilePhone, StreetAddress, City, State, PostalCode, Country | ConvertTo-Json -Depth 10)
        }
        elseif ($Contact.EntryType -eq 'Contact') {
            $ManagedContactString = ($Contact | Select-Object -Property DisplayName, GivenName, Surname, CompanyName, JobTitle, Department, Mail, Phones, Addresses | ConvertTo-Json -Depth 10)
        }

        $ManagedContactChecksum = [System.BitConverter]::ToString($md5.ComputeHash($utf8.GetBytes($ManagedContactString)))
        Write-VerboseLog "[$Mailbox] Adding contact $($Contact.DisplayName). Checksum:$ManagedContactChecksum"

        $fileAsValue = ""
        if ($FileAsFormat -eq "LastFirst" -and $Contact.Surname -and $Contact.GivenName) {
            $fileAsValue = "$($Contact.Surname), $($Contact.GivenName)"
        } elseif ($FileAsFormat -eq "FirstLast" -and $Contact.GivenName -and $Contact.Surname) {
            $fileAsValue = "$($Contact.GivenName) $($Contact.Surname)"
        } elseif ($Contact.DisplayName) {
            $fileAsValue = $Contact.DisplayName
        }

        $newContact = @{
            extensions = @(
                @{
                    "@odata.type" = "microsoft.graph.openTypeExtension"
                    ExtensionName = "ManagedContactCorrelation"
                    CorrelationId = $Contact.Id.ToString()
                    PhotoChecksum = $PhotoChecksum
                }
            )
            displayName    = $Contact.DisplayName
            givenName      = $Contact.GivenName
            surname        = $Contact.Surname
            companyName    = $Contact.CompanyName
            jobTitle       = $Contact.JobTitle
            department     = $Contact.Department
            officeLocation = $Contact.OfficeLocation
            fileAs         = $fileAsValue
            emailAddresses = @(
                @{
                    name    = $Contact.DisplayName
                    address = $Contact.Mail
                }
            )
        }

        if ($Categories.Count -gt 0) {
            $newContact.categories = $Categories
        }

        if ($Contact.EntryType -eq 'User') {
            $newContact.businessAddress = @{
                street          = $Contact.StreetAddress
                city            = $Contact.City
                state           = $Contact.State
                postalCode      = $Contact.PostalCode
                countryOrRegion = $Contact.Country
            }
            $newContact.businessPhones = $Contact.BusinessPhones
            $newContact.mobilePhone    = $Contact.MobilePhone
        }
        elseif ($Contact.EntryType -eq 'Contact') {
            $newContact.businessAddress = @{
                street          = $Contact.Addresses[0].Street
                city            = $Contact.Addresses[0].City
                state           = $Contact.Addresses[0].State
                postalCode      = $Contact.Addresses[0].PostalCode
                countryOrRegion = $Contact.Addresses[0].Country
            }
            $newContact.businessPhones = @($Contact.Phones[2].Number)
            $newContact.mobilePhone    = $Contact.Phones[1].Number
        }

        $newContactObject = New-MgUserContactFolderContact -UserId $Mailbox -ContactFolderId $ManagedContactFolder.Id -BodyParameter $newContact

        if ($null -ne $ContactPhotoFile) {
            Write-VerboseLog "[$Mailbox] Adding photo to contact: $($Contact.DisplayName)"
            Set-MgUserContactFolderContactPhotoContent -UserId $Mailbox -ContactFolderId $ManagedContactFolder.Id -ContactId $newContactObject.Id -InFile $ContactPhotoFile
        }
    }
}

# ---------------------------------------------------------------------------
# MAIN
# ---------------------------------------------------------------------------

Write-InfoLog "=== Starting Graph Contact Sync (Azure Runbook) ==="

$ErrorActionPreference = "Stop"

# Force TLS 1.2
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# ---------------------------------------------------------------------------
# AUTHENTICATION — Azure Automation Certificate Asset
# The certificate must be uploaded under:
#   Automation Account > Assets > Certificates > [CertificateAssetName]
#
# IMPORTANT: the certificate in the asset must include the private key (PFX).
# ---------------------------------------------------------------------------
Write-InfoLog "Loading certificate from Automation Account asset: '$CertificateAssetName'"
try {
    $Certificate = Get-AutomationCertificate -Name $CertificateAssetName
}
catch {
    throw "Unable to load certificate '$CertificateAssetName' from the Automation Account. " +
          "Make sure it has been uploaded under Assets > Certificates. Error: $_"
}

if ($null -eq $Certificate) {
    throw "Certificate '$CertificateAssetName' was not found in the Automation Account Assets."
}

Write-InfoLog "Certificate loaded: $($Certificate.Subject)"

# ---------------------------------------------------------------------------
# Connect to Microsoft Graph
# TenantId accepts both the .onmicrosoft.com name and the tenant GUID
# ---------------------------------------------------------------------------
Write-InfoLog "Connecting to Microsoft Graph (Tenant: $ExchangeOrg, ClientID: $ClientID)"
Connect-MgGraph -Certificate $Certificate -ClientId $ClientID -TenantId $ExchangeOrg -NoWelcome

# ---------------------------------------------------------------------------
# Retrieve users and contacts
# ---------------------------------------------------------------------------
Write-InfoLog "Retrieving user list from directory..."
$UserList = Get-MgUser -Filter '(AccountEnabled eq true)' -All -Property `
    Id, UserType, UserPrincipalName, ShowInAddressList, EmployeeId, DisplayName, GivenName, Surname, `
    CompanyName, JobTitle, Department, OfficeLocation, Mail, BusinessPhones, MobilePhone, `
    StreetAddress, City, State, PostalCode, Country `
| Select-Object @{Name='EntryType'; Expression={'User'}}, `
    Id, UserType, UserPrincipalName, ShowInAddressList, EmployeeId, DisplayName, GivenName, Surname, `
    CompanyName, JobTitle, Department, OfficeLocation, Mail, BusinessPhones, MobilePhone, `
    StreetAddress, City, State, PostalCode, Country

# Minimum filter: only Member users visible in the address book
$UserList = $UserList | Where-Object UserType -eq 'Member' `
                      | Where-Object { $_.ShowInAddressList -ne $false } `
                      | Where-Object { $_.CompanyName -like 'Erion*' } `
                      | Where-Object { $_.UserPrincipalName -notlike 'adm_*' } `
                      | Where-Object { $_.JobTitle -notlike '*Stage*' }

Write-InfoLog "Found $($UserList.Count) users after filtering."

$CombinedContactList = $UserList
Write-InfoLog "Total contacts to sync: $($CombinedContactList.Count)"

# ---------------------------------------------------------------------------
# Determine target mailboxes
# ---------------------------------------------------------------------------
if ($MailboxList -eq "DIRECTORY") {
    $MailboxTargets = ($UserList | Select-Object UserPrincipalName).UserPrincipalName
    Write-InfoLog "Target: DIRECTORY ($($MailboxTargets.Count) mailboxes)"
} else {
    $MailboxTargets = $MailboxList -split ","
    Write-InfoLog "Target: $($MailboxTargets.Count) specified mailbox(es)"
}

# ---------------------------------------------------------------------------
# Synchronisation
# ---------------------------------------------------------------------------
$SuccessCount = 0
$ErrorCount   = 0

foreach ($MailboxTarget in $MailboxTargets) {
    try {
        Write-DebugLog "[$MailboxTarget] Syncing Managed Contacts"
        Sync-ManagedContacts `
            -Mailbox $MailboxTarget `
            -ManagedContactFolderName $ManagedContactFolderName `
            -ManagedContacts $CombinedContactList `
            -FileAsFormat $FileAsFormat `
            -Categories $Categories
        $SuccessCount++
    }
    catch {
        Write-ErrorLog "Error syncing Managed Contacts for Mailbox: $MailboxTarget Exception: $($_.Exception.Message)"
        $ErrorCount++
    }
}

# ---------------------------------------------------------------------------
# Cleanup and summary
# ---------------------------------------------------------------------------
Disconnect-MgGraph

Write-InfoLog "=== Graph Contact Sync completed. Success: $SuccessCount | Errors: $ErrorCount ==="

if ($ErrorCount -gt 0) {
    Write-Warning "Completed with $ErrorCount error(s). Check the job logs for details."
}
