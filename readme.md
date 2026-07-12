# Graph Contact Sync

Synchronizes Global Address List and Organizational Contacts from a M365 environment to selected/all mailboxes in the directory. Uses the MS Graph API PowerShell module to perform all operations.

Forked from [tardispilot/GraphContactSync](https://github.com/tardispilot/GraphContactSync).  
Heavily inspired by the excellent [EWS-Office365-Contact-Sync](https://github.com/grahamr975/EWS-Office365-Contact-Sync) code by grahamr975.

---

## What's new in this fork

- **`GraphContactSync-Runbook.ps1`** — A version of the script adapted to run as an **Azure Automation Runbook**, with no local infrastructure required.  
  Key differences from the original:
  - Authentication via `Get-AutomationCertificate` (certificate stored as an Automation Account asset — no PFX files on disk)
  - Sensitive parameters (`ClientID`, `ExchangeOrg`) can be stored as encrypted Automation Variables
  - Removed PoShLog dependency; logging uses native `Write-Output` / `Write-Warning` / `Write-Error` so output appears in Azure Automation Job Streams
  - `Set-ExecutionPolicy` removed (not applicable in the Automation sandbox)
  - Photos saved to `$env:TEMP` (ephemeral file system available during job execution)

---

## Features

- Includes org-level contacts from M365 → Users → Contacts (useful for non-person entries such as office/branch information)
- Compares old and new field values and only replaces a contact if a change is detected
- **Automatic photo change detection**: tracks profile photo metadata and updates contacts only when photos actually change
- **FileAs field formatting**: configure how contacts are filed ("First Last" or "Last, First")
- **Categories support**: assign categories to contacts, useful when syncing to the main Contacts folder

---

## Security

### Certificate Authentication Methods

The original script (`GraphContactSync.ps1`) supports three authentication methods, listed from most to least secure:

1. **Certificate Thumbprint (Recommended)** — uses a certificate installed in the Windows Certificate Store; no password storage required
2. **Encrypted Password File** — stores the PFX password in a file that can only be decrypted by the same user on the same machine
3. **Plaintext Password** — not recommended for production use

The Runbook variant (`GraphContactSync-Runbook.ps1`) uses the **Automation Account certificate asset**, which is the equivalent of method 1 in a serverless context.

### Security Best Practices

- Use Certificate Thumbprint authentication (or the Automation Account asset equivalent) whenever possible
- Never commit plaintext passwords or PFX files to source control
- Regularly rotate certificates
- Apply least-privilege principles when assigning Azure application permissions

### Creating Encrypted Password Files (original script only)

```powershell
.\Getting Started\Create-EncryptedPassword.ps1 -OutputPath "C:\Certs\certificate.cred"
```

---

## Getting Started — Original Script (`GraphContactSync.ps1`)

### Prerequisites

```powershell
# Install the Microsoft Graph PowerShell module
Install-Module Microsoft.Graph

# Install PoShLog for structured console/file logging
Install-Module PoShLog
```

You will also need your Office 365 organisation URL (format: `mycompany.onmicrosoft.com`).  
Find it under **Microsoft 365 Admin Center → Settings → Domains**.

### Step 1 — Create Certificates

```powershell
cd "Getting Started"

# Recommended: create an encrypted password file alongside the certificate
.\Create-Certificates.ps1 -CertificateName contactsync.mydomain.com -CertificatePassword 'myPassword!' -CreatePasswordFile
```

This produces:
- `contactsync.mydomain.com.pfx` — private key + certificate (keep secure)
- `contactsync.mydomain.com.cer` — public certificate for Azure upload
- `contactsync.mydomain.com.cred` — encrypted password file (if `-CreatePasswordFile` was used)

Note the **certificate thumbprint** printed at the end — you will need it.

### Step 2 — Create the Azure App Registration

#### 2.1 Register the application

1. Go to **Azure Portal → Azure Active Directory → App registrations → New registration**
2. Name: `GraphContactSync` (or your preferred name)
3. Supported account types: *Accounts in this organizational directory only*
4. Click **Register** and note the **Application (client) ID**

#### 2.2 Configure Authentication

1. Go to **Authentication → Add a platform → Mobile and desktop applications**
2. Add redirect URI: `https://login.microsoftonline.com/common/oauth2/nativeclient`
3. Under **Advanced settings**, set **Allow public client flows** to **Yes**
4. Click **Save**

#### 2.3 Upload the Certificate

1. Go to **Certificates & secrets → Upload certificate**
2. Select the `.cer` file from Step 1
3. Add a description (e.g. "GraphContactSync Certificate") and click **Add**

#### 2.4 Configure API Permissions

1. Go to **API permissions → Add a permission → Microsoft Graph → Application permissions**
2. Add:
   - `Contacts.ReadWrite`
   - `User.Read.All`
3. Click **Grant admin consent for [Your Organisation]** and confirm

![API Permissions](images/api_permissions.png)

#### 2.5 Alternative: Configure via Manifest

```json
"requiredResourceAccess": [
    {
        "resourceAppId": "00000003-0000-0000-c000-000000000000",
        "resourceAccess": [
            { "id": "6918b873-d17a-4dc1-b314-35f528134491", "type": "Role" },
            { "id": "df021288-bdef-4463-88db-98f22de89214", "type": "Role" }
        ]
    }
]
```

### Step 3 — Test with a Single Mailbox

#### Method 1: Certificate Thumbprint (Recommended)

```powershell
.\GraphContactSync.ps1 `
    -ExchangeOrg        "mycompany.onmicrosoft.com" `
    -ClientID           "your-application-client-id" `
    -CertificateThumbprint "your-certificate-thumbprint" `
    -MailboxList        "testuser@mycompany.com" `
    -ManagedContactFolderName "Company Contacts - Test" `
    -LogPath            "$PSScriptRoot\Logs" `
    -FileAsFormat       "LastFirst" `
    -Categories         @("Business Contacts", "Company Directory")
```

#### Method 2: Encrypted Password File

```powershell
.\GraphContactSync.ps1 `
    -ExchangeOrg        "mycompany.onmicrosoft.com" `
    -ClientID           "your-application-client-id" `
    -CertificatePath    "C:\Certs\contactsync.mydomain.com.pfx" `
    -CertificatePasswordFile "C:\Certs\contactsync.mydomain.com.cred" `
    -MailboxList        "testuser@mycompany.com" `
    -ManagedContactFolderName "Company Contacts - Test" `
    -LogPath            "$PSScriptRoot\Logs"
```

#### Method 3: Plaintext Password (not recommended for production)

```powershell
.\GraphContactSync.ps1 `
    -ExchangeOrg        "mycompany.onmicrosoft.com" `
    -ClientID           "your-application-client-id" `
    -CertificatePath    "C:\Certs\contactsync.mydomain.com.pfx" `
    -CertificatePassword "YourCertificatePassword" `
    -MailboxList        "testuser@mycompany.com" `
    -ManagedContactFolderName "Company Contacts - Test" `
    -LogPath            "$PSScriptRoot\Logs"
```

### Step 4 — Deploy to All Mailboxes

```powershell
.\GraphContactSync.ps1 `
    -ExchangeOrg        "mycompany.onmicrosoft.com" `
    -ClientID           "your-application-client-id" `
    -CertificateThumbprint "your-certificate-thumbprint" `
    -MailboxList        "DIRECTORY" `
    -ManagedContactFolderName "Company Contacts" `
    -LogPath            "$PSScriptRoot\Logs" `
    -FileAsFormat       "LastFirst" `
    -Categories         @("Business Contacts")
```

### Step 5 — Schedule Automated Runs (Optional)

1. Open **Task Scheduler** and create a new task
2. Set it to run whether the user is logged on or not
3. Add a trigger for your desired schedule (e.g. daily at 06:00)
4. Action: `PowerShell.exe -ExecutionPolicy Bypass -File "C:\Path\To\ProductionRun.ps1"`
5. Use a service account with the appropriate permissions

---

## Getting Started — Azure Automation Runbook (`GraphContactSync-Runbook.ps1`)

This variant runs entirely in Azure with no local server or scheduled task required.

### Prerequisites in the Automation Account

Before importing the Runbook, ensure the following are configured in your Azure Automation Account:

1. **PowerShell modules** imported under *Modules*:
   - `Microsoft.Graph.Authentication`
   - `Microsoft.Graph.Users`
   - `Microsoft.Graph.PersonalContacts`

2. **Certificate** uploaded under *Assets → Certificates* (must include the private key — upload the `.pfx`).  
   Default expected name: `GraphContactSyncCert` (configurable via the `CertificateAssetName` parameter).

3. **Automation Variables** (optional but recommended to avoid hardcoding values in the Runbook parameters):

   | Variable name | Type | Notes |
   |---|---|---|
   | `GraphContactSync-ExchangeOrg` | String (encrypted recommended) | Your `.onmicrosoft.com` tenant |
   | `GraphContactSync-ClientID` | String | App Registration Client ID |

   If these variables are absent, the values must be passed as Runbook parameters at runtime.

### Azure App Registration

The same app registration used for the original script works here. Ensure it has:
- `Contacts.ReadWrite` (Application permission)
- `User.Read.All` (Application permission)
- `OrgContact.Read.All` (Application permission)
- The certificate uploaded under **Certificates & secrets**

### Importing the Runbook

1. In the Azure Portal, navigate to your **Automation Account → Runbooks → Import a runbook**
2. Upload `GraphContactSync-Runbook.ps1`
3. Select **PowerShell** as the Runbook type
4. Click **Create**, then **Publish**

### Runbook Parameters

| Parameter | Mandatory | Default | Description |
|---|---|---|---|
| `ExchangeOrg` | No* | `""` | `.onmicrosoft.com` tenant. Read from Automation Variable if empty. |
| `ClientID` | No* | `""` | App Registration Client ID. Read from Automation Variable if empty. |
| `CertificateAssetName` | No | `GraphContactSyncCert` | Name of the certificate asset in the Automation Account. |
| `MailboxList` | **Yes** | — | Comma-separated mailbox list, or `"DIRECTORY"` for all users. |
| `ManagedContactFolderName` | **Yes** | — | Name of the contact folder to manage in each mailbox. |
| `FileAsFormat` | No | `FirstLast` | `"FirstLast"` or `"LastFirst"`. |
| `Categories` | No | `@()` | Array of categories to assign to synced contacts. |

\* Mandatory if the corresponding Automation Variable is not set.

### Running the Runbook

#### From the Azure Portal

1. Open the Runbook and click **Start**
2. Fill in the parameters in the panel that appears
3. Click **OK** — the job will start and output will appear in the **Job Streams** view

#### Example: sync to a single test mailbox

| Parameter | Value |
|---|---|
| `MailboxList` | `testuser@mycompany.com` |
| `ManagedContactFolderName` | `Company Contacts - Test` |
| `FileAsFormat` | `LastFirst` |

#### Example: sync to all mailboxes

| Parameter | Value |
|---|---|
| `MailboxList` | `DIRECTORY` |
| `ManagedContactFolderName` | `Company Contacts` |
| `FileAsFormat` | `LastFirst` |

### Scheduling the Runbook

1. Open the Runbook and go to **Schedules → Add a schedule**
2. Create a new schedule (e.g. daily at 06:00 UTC) or link an existing one
3. Fill in the Runbook parameters for that schedule
4. Click **OK**

From that point on the Runbook will run automatically. Results and any errors will be visible under **Jobs** in the Automation Account.

### Monitoring and Logs

In Azure Automation, output is split into Job Streams:
- **Output** — `[INFO]` and `[DEBUG]` messages
- **Warning** — `[WARN]` messages
- **Error** — `[ERROR]` messages

To inspect a run, go to **Automation Account → Jobs**, select a job, and open the relevant stream tab.

### Troubleshooting

| Symptom | Likely cause | Fix |
|---|---|---|
| `Cannot load certificate` | Certificate not uploaded or wrong name | Check *Assets → Certificates*; verify `CertificateAssetName` |
| `GraphContactSync-ExchangeOrg variable not found` | Automation Variable missing | Create it under *Assets → Variables* or pass the value as a parameter |
| `Insufficient privileges` | Missing Graph permissions or admin consent not granted | Verify permissions in the App Registration and re-grant admin consent |
| Module not found errors | Required Graph modules not imported | Import `Microsoft.Graph.Users` and `Microsoft.Graph.PersonalContacts` under *Modules* |
| Photos not syncing | `$env:TEMP` path issue | Check the Output stream for photo-related log lines; the temp folder is created automatically |

---

## Parameters — Original Script (`GraphContactSync.ps1`)

### Required

- `ExchangeOrg` — the Exchange organisation (`.onmicrosoft.com` tenant)
- `ClientID` — the App Registration Client ID
- `MailboxList` — comma-separated mailbox list, or `"DIRECTORY"` for all
- `ManagedContactFolderName` — name of the contact folder to sync to
- `LogPath` — path for log file output

### Certificate Authentication (choose one method)

- **Method 1** (recommended): `-CertificateThumbprint`
- **Method 2**: `-CertificatePath` + `-CertificatePasswordFile`
- **Method 3** (not recommended): `-CertificatePath` + `-CertificatePassword`

### Optional

- `FileAsFormat` — `"FirstLast"` (default) or `"LastFirst"`
- `Categories` — array of category strings, e.g. `@("Business Contacts")`

---

## Photo Handling

- Photos are downloaded to a local `Photos\` subdirectory (original script) or `$env:TEMP\GraphContactSync_Photos\` (Runbook)
- The script tracks photo metadata and only re-downloads when a change is detected
- No manual intervention is required

---

## Notes and Disclaimers

- Test thoroughly before deploying in any automated or unconstrained way. This application has permissions to **DELETE** contacts from mailboxes.
- This works for the maintainer's own environment; your mileage may vary. Adapt filters in the user query to match your organisation.
- PRs are welcome. Please be respectful.

---

## License

MIT — see [LICENSE.md](LICENSE.md) for details.

---

## Acknowledgements

- **tardispilot** for the original [GraphContactSync](https://github.com/tardispilot/GraphContactSync)
- **Ryan Graham** for [EWS-Office365-Contact-Sync](https://github.com/grahamr975/EWS-Office365-Contact-Sync) and the concept behind this project
