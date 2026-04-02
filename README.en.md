# Windows Server Network Adapter Report

Monorepo with two localized editions of the same PowerShell solution for auditing Windows Server network adapters,
default gateways, and legacy TCP/IP profiles.

## Repository Layout

- [en/](./en/): full English edition
- [ru/](./ru/): full Russian edition
- [LICENSE](./LICENSE): shared MIT license

## What Is Inside Each Edition

Each edition contains its own:

- localized entry script
- localized help text
- localized runtime messages
- localized documentation
- sample `servers.txt` file

The collection logic is equivalent across editions. The operator-facing language changes depending on the audience.

## Recommended Entry Points

English edition:

- [en/get_windows_server_network_adapter_report.ps1](./en/get_windows_server_network_adapter_report.ps1)

Russian edition:

- [ru/get_windows_server_network_adapter_report.ps1](./ru/get_windows_server_network_adapter_report.ps1)

## Publishing Strategy

This layout is intended for publishing one Git repository instead of maintaining separate repositories only because of
language differences.

Use this structure when you want:

- one issue tracker
- one release flow
- one shared license
- one star and fork history
- two operator-facing language editions

## License

Distributed under the MIT License. See [LICENSE](./LICENSE).
