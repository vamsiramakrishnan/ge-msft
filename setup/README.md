# Setup Guide

This folder is the operator runbook for development sideloading and tenant setup.

Start here:

0. [Readiness and guided setup](./00-readiness.md)
1. [Prerequisites and configuration](./01-prerequisites-and-config.md)
2. [Development server and Cloudflare tunnel](./02-dev-server-and-tunnel.md)
3. [Manifest generation and packages](./03-manifests-and-packages.md)
4. [Sideloading by host and client](./04-sideloading.md)
5. [Debugging and troubleshooting](./05-debugging.md)
6. [Tenant deployment automation](./06-tenant-deployment.md)
7. [Deployment methods matrix](./07-deployment-methods-matrix.md)
8. [Hosting origin and release flow](./08-hosting-origin-and-release.md)

Use [Deployment methods matrix](./07-deployment-methods-matrix.md) before deciding whether
`bun bootstrap` should use `--deployment-lane catalog`, `--deployment-lane xml`, or sideloading.
Use [Hosting origin and release flow](./08-hosting-origin-and-release.md) before deciding whether
the add-in should load from a Cloudflare tunnel, GCS/CDN, App Engine, or Cloud Run.

Do not commit real tenant IDs, app IDs, domains, or secrets in these docs. Put real values in
`packages/web-shell/.env` or in your tenant/admin consoles.
