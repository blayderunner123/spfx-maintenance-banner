# SharePoint Maintenance Banner v1.3.4

Copyright © 2026 Jonathan R. Adcox

Original author: **Jonathan R. Adcox (Blayderunner123)**

This is a forensic reconstruction from the deployed production SPPKG. It is not claimed to be the original TypeScript tree. See `FORENSIC-NOTES.md` for provenance and known limitations.

Version 1.3.4 is the publication-ready successor to the reconstructed v1.3.3 artifact. It removes organization-specific metadata and makes the configuration web explicit rather than hardcoded.

## Configuration

Create a SharePoint list named `Site Maintenance Banners` using the field inventory in `FORENSIC-NOTES.md`. By default, the extension reads that list from the current web. To use a central configuration web, set the Application Customizer property when registering the extension:

```json
{
  "configurationSiteUrl": "https://contoso.sharepoint.com/sites/configuration"
}
```

The requested Microsoft Graph `GroupMember.Read.All` permission must be approved by a tenant administrator for Graph-backed audience targeting. SharePoint-group and direct-user targeting do not depend on that Graph call.

## Build

Use Node.js 18 (the supported SPFx 1.20 toolchain), then run:

```powershell
npm ci
npm run package
```

The package is written to `sharepoint/solution/sharepoint-maintenance-banner.sppkg`.

No deployment is performed by these commands.

## License

GNU General Public License v3.0 or later (`GPL-3.0-or-later`). See [LICENSE](LICENSE).

This license permits use, modification, forking, improvement, and redistribution under its terms. Modified versions need not retain project branding or imply endorsement by the original author.
