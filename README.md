# MFASweep

MFASweep is a PowerShell script that attempts to log in to various Microsoft services using a provided set of credentials and will attempt to identify if MFA is enabled. Depending on how conditional access policies and other multi-factor authentication settings are configured some protocols may end up being left single factor. It also has an additional check for ADFS configurations and can attempt to log in to the on-prem ADFS server if detected.

Currently MFASweep has the ability to login to the following services:

- Microsoft Graph API
- Azure Service Management API
- Microsoft 365 Exchange Web Services
- Microsoft 365 Web Portal w/ 6 device types (Windows, Linux, MacOS, Android Phone, iPhone, Windows Phone)
- Microsoft 365 Active Sync
- ADFS

**WARNING: This script attempts to login to the provided account TEN (10) different times (11 if you include ADFS). If you entered an incorrect password this may lock the account out.**

For more information check out the blog post here: [Exploiting MFA Inconsistencies on Microsoft Services](https://www.blackhillsinfosec.com/exploiting-mfa-inconsistencies-on-microsoft-services/)

![MFASweep Example](/example.jpg?raw=true)

![Single Factor Access Results Example](https://user-images.githubusercontent.com/2296229/204374571-0b299177-a5ab-4e05-a313-d9fe5151d1d6.png)

## Usage

This command will use the provided credentials and attempt to authenticate to the Microsoft Graph API, Azure Service Management API, Microsoft 365 Exchange Web Services, Microsoft 365 Web Portal with both a desktop browser and mobile, and Microsoft 365 Active Sync. If any authentication methods result in success, tokens and/or cookies will be written to AccessTokens.json. (Currently does not log cookies or tokens for EWS, ActiveSync, and ADFS)

```PowerShell
Invoke-MFASweep -Username targetuser@targetdomain.com -Password Winter2024 -WriteTokens
```

Key optional switches:

- `-WriteTokens` – saves any collected access/refresh tokens and cookies to `AccessTokens.json`.
- `-Recon` / `-IncludeADFS` – enumerate ADFS settings and test the detected on-prem ADFS endpoint.
- `-UseChromium` – forces Microsoft 365 web portal checks through a Playwright-driven Chromium instance to better mimic real browsers (see below for setup steps).

This command runs with the default auth methods and checks for ADFS as well.

```PowerShell
Invoke-MFASweep -Username targetuser@targetdomain.com -Password Winter2020 -Recon -IncludeADFS
```

### Chromium Automation Mode

Microsoft now blocks many scripted requests to the Microsoft 365 web portal. MFASweep can optionally emulate a real browser session by calling Playwright/Chromium.

1. Install [Node.js](https://nodejs.org/) (v18+ recommended).
2. From the repository root run:

```powershell
npm install
npx playwright install chromium
```

3. Execute MFASweep with the `-UseChromium` switch (other parameters remain the same):

```powershell
Invoke-MFASweep -Username targetuser@targetdomain.com -Password Winter2024 -UseChromium -WriteTokens
```

During Chromium mode the script launches a headless browser for every user agent check, captures cookies/tokens, and reports whether MFA was required.

## Individual Modules

Each individual module can be run separately if needed as well.

**Microsoft Graph API**

```PowerShell
Invoke-GraphAPIAuth -Username targetuser@targetdomain.com -Password Winter2020
```

**Azure Service Management API**

```PowerShell
Invoke-AzureManagementAPIAuth -Username targetuser@targetdomain.com -Password Winter2020
```

**Microsoft 365 Exchange Web Services**

```PowerShell
Invoke-EWSAuth -Username targetuser@targetdomain.com -Password Winter2020
```

**Microsoft 365 Web Portal**

```PowerShell
Invoke-O365WebPortalAuth -Username targetuser@targetdomain.com -Password Winter2020
```

**Microsoft 365 Web Portal w/ Mobile User Agent**

```PowerShell
Invoke-O365WebPortalAuthMobile -Username targetuser@targetdomain.com -Password Winter2020
```

**Microsoft 365 Active Sync**

```PowerShell
Invoke-O365ActiveSyncAuth -Username targetuser@targetdomain.com -Password Winter2020
```

**ADFS**

```PowerShell
Invoke-ADFSAuth -Username targetuser@targetdomain.com -Password Winter2020
```

### Brute Forcing Client IDs During ROPC Auth

The Invoke-BruteClientIDs function will loop through various resource types and client IDs during ROPC auth to find single factor access for various combinations of client IDs and resources. If any authentication methods result in success, tokens and/or cookies will be written to AccessTokens.json. (Currently does not log cookies or tokens for EWS, ActiveSync, and ADFS)

```PowerShell
Invoke-BruteClientIDs -Username targetuser@targetdomain.com -Password Winter2024 -VerboseOut
```
