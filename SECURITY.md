# Security Policy

## Supported Versions

Security fixes are applied to the current `main` branch of this repository.

## Reporting A Vulnerability

Please do not publish potential security issues as public GitHub issues before they are reviewed.

If you believe you found a vulnerability:

1. prepare a short description of the issue
2. include affected files or execution paths
3. include safe reproduction steps if possible
4. describe the potential impact

For sensitive reports, contact the maintainer privately first and wait for confirmation before public disclosure.

## Operational Notes

This project is an administrative reporting tool. Many real-world failures are caused by:

- insufficient permissions
- unreachable servers
- blocked WMI/DCOM or remote-registry access
- stale or inaccurate source inventories
- unsupported local execution environments

Please distinguish operational environment problems from code-level security issues when reporting.
