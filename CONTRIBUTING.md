# Contributing

Thanks for your interest in improving this project.

## Scope

This repository contains two localized editions of the same PowerShell solution:

- `en/` for English operator-facing usage
- `ru/` for Russian operator-facing usage

Please keep functional changes aligned across both editions unless the change is intentionally language-specific.

## Before You Open A Pull Request

1. Keep behavior changes minimal and reviewable.
2. Preserve compatibility with Windows Server 2008-2025 where practical.
3. Update both language editions when the underlying collection logic changes.
4. Update user-facing documentation when parameters, outputs, or report behavior change.
5. Do not commit real `servers.txt`, exported reports, credentials, or internal infrastructure details.

## Development Notes

- WMI/DCOM and remote-registry collection are intentional compatibility choices for mixed Windows Server environments.
- PowerShell `5.1` compatibility matters for worker execution and broad operator adoption.
- PowerShell `7+` should continue to work for orchestration and parallel collection scenarios.
- Excluding `WAN Miniport` and `Microsoft` adapters is part of the intended reporting behavior unless explicitly redesigned.

## Pull Request Guidelines

- Describe the problem and the operational impact.
- Summarize the change at a high level.
- Mention compatibility implications if any.
- Note whether both language editions were reviewed.
- Include manual validation notes when applicable.

## Issues

Use the issue templates when reporting bugs or requesting features.
