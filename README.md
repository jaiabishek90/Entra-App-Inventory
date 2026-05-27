# Entra Enterprise App Inventory Report

PowerShell automation to inventory integrated enterprise applications in Microsoft Entra ID, enrich them with permissions, sign-in activity, credential metadata, assignments, ownership, and Exchange Application Access Policy scoping, then export a CSV and send an HTML summary email.

> **GitHub-ready package**: this repository version is sanitized and parameterized for safe sharing. Update the example configuration before running.

---

## Features

- Collects integrated enterprise applications from Entra ID
- Resolves application owners
- Resolves group and role memberships
- Resolves assigned users and groups
- Collects application and delegated permissions
- Flags high-priority permissions
- Collects sign-in activity and 30-day sign-in summary
- Merges credentials from:
  - Application object
  - Service Principal object
- Tracks:
  - expired secrets
  - expired certificates
  - earliest expiry dates
- Includes Exchange Application Access Policy scoping
- Exports timestamped CSV report
- Sends HTML email summary with KPI metrics

---

## Repository Structure

```text
Entra-App-Inventory-GitHub-Package/
│
├── src/
│   └── Get-EntraEnterpriseAppInventory.ps1
│
├── config/
│   └── settings.example.json
│
├── docs/
│   └── sample-output.md
│
├── .github/
│   └── workflows/
│       └── ci.yml
│
├── .gitignore
├── CHANGELOG.md
├── LICENSE
└── README.md
```

---

## Requirements

### PowerShell
- Windows PowerShell 5.1 or PowerShell 7+

### Modules
- Microsoft.Graph.Beta.Applications
- ExchangeOnlineManagement

### Access / Authentication
The script uses:
- App registration
- Certificate-based authentication
- Microsoft Graph
- Exchange Online

---

## Configuration

1. Copy the example configuration.
2. Replace the placeholder values.
3. Pass the file to the script with `-ConfigPath`.

### Example

```powershell
Copy-Item .\config\settings.example.json .\config\settings.json
notepad .\config\settings.json
```

### Run

```powershell
pwsh .\src\Get-EntraEnterpriseAppInventory.ps1 -ConfigPath .\config\settings.json
```

---

## Output

### CSV Export
The script exports a timestamped CSV file:

```text
yyyy-MM-dd_HH-mm-ss_EntraAppInventory.csv
```

### HTML Email Summary
The HTML report includes:
- Total apps scanned
- New apps in last 7 days
- Apps with high-priority permissions
- New apps with high-priority permissions
- Top 15 newest apps
- Top 15 high-priority apps

---

## Security Recommendations

Before pushing to GitHub or sharing internally:

- Do **not** commit tenant-specific values
- Do **not** commit certificate files
- Use placeholders in examples
- Store sensitive values in:
  - pipeline variables
  - environment variables
  - secure automation accounts
  - Azure Key Vault (if applicable)

---

## Suggested Improvements

Future enhancements you may consider:

- Replace `Send-MailMessage` with Microsoft Graph mail sending
- Add Pester tests
- Add Excel export in addition to CSV
- Add risk scoring model
- Publish as a PowerShell module

---

## License

This package includes the MIT License by default. Change it if your organization requires a different licensing model.

---

## Maintainer

- Team: Infrastructure Engineering
- Owner: Update with your team/contact
