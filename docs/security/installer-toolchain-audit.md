# Installer toolchain audit

Checked on 2026-07-15 with Node 22.23.1 and npm 10.9.8.

Local Excel-for-Mac installation uses an add-in-only XML manifest and the pinned `office-addin-dev-certs` 2.0.10 package. It does not include Microsoft 365 Agents Toolkit, a cloud title client, or a manifest packager in the production installer toolchain.

A clean production-only install contains 19 packages, occupies approximately 3.6 MB, and reports zero vulnerabilities from `npm audit --omit=dev`.

Current controls:

- Installation is reproduced from an exact lockfile with offline `npm ci --ignore-scripts --omit=dev`.
- Lifecycle status verifies a digest of the complete locked certificate-tool tree, package metadata, lockfile, and direct executable before using the persisted tool.
- The XML manifest registration is an exact hard link owned by Analyst Lens; repair and uninstall refuse mismatched paths, hashes, or inodes.
- Lifecycle commands have no permitted non-loopback network destination.
- The installer never receives workbook content.

This removes the 422 MB Agents Toolkit dependency and its six moderate transitive advisories. The original cloud-title limitation remains documented as the reason for the Mac XML pivot.
