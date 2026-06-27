# Setup Guide

This folder is the operator runbook for development sideloading and tenant setup.

Start here:

1. [Prerequisites and configuration](./01-prerequisites-and-config.md)
2. [Development server and Cloudflare tunnel](./02-dev-server-and-tunnel.md)
3. [Manifest generation and packages](./03-manifests-and-packages.md)
4. [Sideloading by host and client](./04-sideloading.md)
5. [Debugging and troubleshooting](./05-debugging.md)

Do not commit real tenant IDs, app IDs, domains, or secrets in these docs. Put real values in
`packages/web-shell/.env` or in your tenant/admin consoles.
