# Exchange Online DL Automation via Azure Automation + Webhook

## ✅ Overview

This guide sets up **certificate‑based automation** for managing **classic Exchange Online distribution lists (DLs)** using **Azure Automation**, a **PowerShell 7 runbook**, and a **webhook**.

The runbook relies on a single `$WebhookData` parameter (auto‑injected by Azure Automation) and extracts `Action`, `DLName`, and `MemberUPNs` from the webhook body.

---

## 🔹 1  Create a Self‑Signed Certificate (CSP)

```powershell
$cert = New-SelfSignedCertificate `
    -Subject "CN=ExO-Automation" `
    -CertStoreLocation Cert:\CurrentUser\My `
    -KeyAlgorithm RSA -KeyLength 2048 `
    -KeyExportPolicy Exportable `
    -Provider "Microsoft Enhanced RSA and AES Cryptographic Provider" `
    -NotAfter (Get-Date).AddYears(3)

Export-Certificate  -Cert $cert -FilePath "$env:TEMP\ExOAutomation.cer"
Export-PfxCertificate -Cert $cert -FilePath "$env:TEMP\ExOAutomation.pfx" `
    -Password (Read-Host -AsSecureString "PFX password")
```

---

## 🔹 2  Register an App in Entra ID (Azure AD)

1. **Entra ID → App registrations → New registration**  
2. Record **Application (client) ID** and **Directory (tenant) ID**  
3. Upload the `.cer` in **Certificates & secrets → Certificates**  
4. Add **API permission** `Exchange.ManageAsApp` (Application) → **Grant admin consent**  
5. Assign the service‑principal to **Exchange Administrator** (or **Exchange Recipient Administrator**)

---

## 🔹 3  Import the Certificate into Azure Automation

* Automation Account → **Certificates**  
* Upload the `.pfx`, name it `ExOAutomationCert`, mark **Exportable**

---

## 🔹 4  Install `ExchangeOnlineManagement` (netCore)

### Option A – From Gallery *(simplest)*

1. Automation → **Modules (preview)**  
2. **Add → From Gallery**  
3. Search `ExchangeOnlineManagement`, choose **Version 3.5.0**  
4. Wait until status = **Available**

### Option B – Manual ZIP upload *(if Gallery import fails)*

1. On a workstation with PowerShellGet:

    ```powershell
    Save-Module ExchangeOnlineManagement -RequiredVersion 3.5.0 -Path .\ExoTemp
    Compress-Archive `
        -Path          .\ExoTemp\ExchangeOnlineManagement `
        -DestinationPath .\ExchangeOnlineManagement.zip
    ```

    > Resulting ZIP folder layout must be:<br>
    > `ExchangeOnlineManagement\3.5.0\netCore\ExchangeOnlineManagement.psd1`

2. Return to **Modules (preview)** → **Add → Local package** → upload `ExchangeOnlineManagement.zip`.  
3. Wait for **Available** status.

---

## 🔹 5  Create the PowerShell 7 Runbook

Runtime: **PowerShell 7.2 (preview)**

```powershell
$VerbosePreference = 'Continue'

param (
    [Parameter(Mandatory=$false)]
    [object] $WebhookData
)

Write-Verbose "▶️ Raw WebhookData:`n$($WebhookData | ConvertTo-Json -Depth 5)"

function Unwrap-Json ($input) {
    $value = $input
    while ($value -is [string]) { $value = $value | ConvertFrom-Json }
    $value
}

if (-not $WebhookData) { throw "Runbook expects to be triggered via webhook." }

$payload     = Unwrap-Json $WebhookData.RequestBody
$Action      = $payload.Action
$DLName      = $payload.DLName
$MemberUPNs  = $payload.MemberUPNs

Write-Verbose "🧩 Action=$Action | DLName=$DLName | MemberUPNs=$MemberUPNs"

if ($Action -notin @('Add','Remove') -or -not $DLName -or -not $MemberUPNs) {
    throw "Action (Add/Remove), DLName, and MemberUPNs are required in the webhook body."
}

# Connect to Exchange Online
$AppId   = 'YOUR-APP-ID'
$OrgName = 'yourtenant.onmicrosoft.com'
$Cert    = Get-AutomationCertificate -Name 'ExOAutomationCert'

Import-Module ExchangeOnlineManagement -ErrorAction Stop
Connect-ExchangeOnline -AppId $AppId -Certificate $Cert -Organization $OrgName -ShowBanner:$false

try {
    $members = $MemberUPNs -split '[,;]' | ForEach-Object { $_.Trim() } | Where-Object { $_ }

    foreach ($m in $members) {
        Write-Verbose "$Action → $m"

        if ($Action -eq 'Add') {
            Add-DistributionGroupMember -Identity $DLName -Member $m -BypassSecurityGroupManagerCheck
        }
        else {
            Remove-DistributionGroupMember -Identity $DLName -Member $m -BypassSecurityGroupManagerCheck -Confirm:$false
        }
    }
}
finally {
    Disconnect-ExchangeOnline -Confirm:$false
}
```

---

## 🔹 6  Enable & Invoke the Webhook

1. **Publish** the runbook  
2. Runbook → **Add webhook → Create new** (no parameters)  
3. Copy the secret URL  

**Example JSON body**

```json
{
  "Action": "Add",
  "DLName": "team-dl@contoso.com",
  "MemberUPNs": "user1@contoso.com,user2@contoso.com"
}
```

**Call it**

```bash
curl -X POST -H "Content-Type: application/json" \
     -d '{ "Action":"Remove", "DLName":"team-dl@contoso.com", "MemberUPNs":"u1@contoso.com;u2@contoso.com" }' \
     "https://<webhook-URL>"
```

### 🚀 Trigger from **Power Automate**

1. **Create a new flow**  
   *Type*: **Instant cloud flow** → trigger **Manually trigger a flow** (or any other event).

2. **Add action** → **HTTP**  
   | Field | Value |
   |-------|-------|
   | **Method** | `POST` |
   | **URI** | *your webhook URL* |
   | **Headers** | `Content-Type : application/json` |
   | **Body** | ``` {  "Action": "Add", "DLName": "team-dl@contoso.com",  "MemberUPNs": "user1@contoso.com;user2@contoso.com"}``` |

3. **Save & Test** the flow.  
   The HTTP action’s response code should be **202 Accepted** and the runbook job will appear in **Azure Automation → Jobs**.


---

## 🔐 Best Practices

| Area | Guidance |
|------|----------|
| **Webhook URL** | Treat as a secret; store in Key Vault / secure config. |
| **Certificate rotation** | Upload new `.cer` to app reg + replace `.pfx` asset. |
| **Logging** | `$VerbosePreference='Continue'`, `Write-Verbose`, enable *Log verbose records*. |
| **Diagnostics** | Forward **JobLogs** + **JobStreams** to Log Analytics or Storage via Diagnostic settings. |

---

You now have a **webhook‑enabled, certificate‑authenticated runbook** that can bulk **Add** or **Remove** members from classic Exchange Online DLs.
