# Microsoft 365 GRC Asset Auditing Template

This repository provides an automated workflow to audit and document governance, risk, and compliance (GRC) status across Microsoft 365, Entra ID, and Intune.

## Features
- Automated asset collection for Users, Groups, and Devices (Entra ID joined, Intune managed, Defender endpoints).
- Certificate-based authentication (CBA) or interactive browser logon.
- Daily automated runs via GitHub Actions.

## Setup
To set up this auditing flow, please use the onboarding portal at [ciso-onboarding-portal](https://ciso-onboarding-portal.pages.dev) and select **Option C: M365 GRC Asset Auditing Setup-Wizard**.

## Known Issues

- **Microsoft Purview Retention Labels (Graph API restriction):** When running in an unattended/headless automation context using App-Only (Application) permissions, Microsoft Graph queries to `/security/labels/retentionLabels` will fail with an HTTP 500/403 (Forbidden/DataInsightsRequestError). This is a known design restriction of Microsoft Graph where retention label endpoints do not support Application permissions. To resolve this, the collector automatically falls back to Security & Compliance Center PowerShell (`Get-ComplianceTag`) when the connection is active, which fully supports App-Only certificate authentication.

