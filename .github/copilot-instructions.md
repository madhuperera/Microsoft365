# Microsoft 365 Codebase Guide

## Repository overview

This repository contains administrator and security tooling for Microsoft 365 services.

The repository primarily contains PowerShell scripts that report on, validate, or remediate Microsoft 365 tenant configuration. It may also contain supporting policy, configuration, HTML, TXT, CSV, or documentation artefacts used by Microsoft 365 services.

Scripts are shared as-is without warranty. Do not add wording that implies formal support, guaranteed outcomes, or production readiness unless that is already stated in the repository.

## General working principles

- Follow the existing repository structure and naming conventions.
- Make the smallest practical change required to complete the requested task.
- Read the closest peer file before editing and preserve the existing style where practical.
- Do not invent technical details, configuration values, policy names, permissions, commands, dependencies, or usage steps that are not supported by repository content.
- When repository evidence is missing, state that clearly instead of filling the gap with assumptions.
- Use New Zealand English.
- Use a professional, consultant-friendly tone.
- Avoid emojis, decorative symbols, and overly casual wording in professional output.
- Do not commit sample output, screenshots, real UPNs, public IPs, internal IP ranges, tenant IDs, customer names, secrets, or credentials.

## Documentation standards

Documentation must be practical, structured, and easy to scan.

A reader should be able to quickly find how to set up, run, validate, and troubleshoot the repository without needing to read the full document from top to bottom.

Where relevant, README and other Markdown documentation should include:

- Purpose
- Scope
- Quick start
- Permissions or prerequisites
- Setup
- Usage
- Validation or testing
- Troubleshooting
- Known limitations

README files should place the most important operational information near the top, including:

- What the repository is for
- What is required before using it
- How to set it up
- How to run it
- How to confirm it worked

Use clear headings, short paragraphs, and tables where they improve readability.

## HTML usage in Markdown documentation

HTML may be used inside Markdown files where it improves readability, navigation, or visual structure.

Acceptable uses include:

- Summary cards
- Callout boxes
- Two-column layouts
- Anchored navigation sections
- Status or capability tables
- Collapsible sections using `<details>` and `<summary>`

HTML must remain simple, readable, and maintainable.

Do not use HTML that makes the documentation harder to edit, harder to read in raw Markdown, or dependent on external styling that may not render consistently in GitHub.

Avoid unnecessary visual decoration. The goal is clarity, not visual noise.

## Microsoft 365 repository boundaries

Do not convert content into a different Microsoft 365 workload or platform pattern unless the task explicitly asks for it.

Examples:

- Do not convert administrator-run scripts into automation runbooks unless requested.
- Do not convert Advanced Hunting queries into Microsoft Sentinel analytics rules unless requested.
- Do not introduce GitHub Actions, CI/CD, Bicep, ARM, Terraform, or deployment workflows unless requested.
- Do not add dependencies or modules that are not required for the requested change.

## Repository folder structure

The following folder rules are non-negotiable. Do not propose, plan, or apply changes that violate them, even when restructuring or "tidying up" the repository.

### Reports live in a single top-level folder

All reporting scripts for every Microsoft 365 workload (Entra ID, Exchange Online, Intune, Defender, licensing, DNS hygiene, and so on) must live in the top-level [Reports/](../Reports) folder.

- Do not create per-workload `Reports/` subfolders such as `EntraID/Reports/`, `Exchange/Reports/`, `Intune/Reports/`, or `DNS/Reports/`.
- Do not move reporting scripts out of `Reports/` into workload folders.
- Workload folders (`Exchange/`, `Configurations/`, `Incident_Response/`, etc.) are reserved for non-reporting assets such as configuration data, transport-rule HTML, remediation actions, and incident-response tooling.

If a reporting script needs grouping inside `Reports/`, achieve it through the filename prefix (for example `ReportEntraID*`, `ReportExo*`, `ReportIntune*`) rather than by adding subfolders.

## Pull request expectations

Pull requests should:

- Have a clear title.
- Summarise what changed.
- Call out any assumptions.
- Call out any files intentionally not changed.
- Include validation steps where practical.
- Identify whether the change affects documentation only, script behaviour, permissions, or operational usage.
