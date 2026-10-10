# Security policy

## Supported versions

| Version | Supported |
|---|---|
| 2.2.x | Yes |
| < 2.2 | No |

## Reporting a vulnerability

Please **do not** open a public issue for security problems.

Use GitHub's private reporting: **Security** tab → **Report a vulnerability** on this repository. Include the script name and version (`-Help` prints it), the parameters used (with tenant names, IDs and URLs removed), and what you observed. You should get an acknowledgement within 5 business days.

## Handling your data when reporting

The scripts never log passwords, tokens or pre-authenticated download/upload URLs. Work-folder logs (`deployment.csv`, `deployment.jsonl`, `publish-home-*.csv`, `remove-*.csv`) can still contain tenant hostnames, site URLs and account names; redact them before sharing.

## Security model

The controls built into the kit are listed in [README.md – Security design](README.md#security-design). In short: delegated browser sign-in only, tenant and host verification before changes, new private sites with external sharing disabled, SHA-256 verified uploads, validated content packages, sanitized errors, and no execution-policy bypass.
