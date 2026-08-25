# Security Policy

## Supported versions

Until the unified application reaches its first production release, security and data-safety fixes are made on `main` only.

## Reporting a vulnerability

Please do not publish exploit details, data-loss reproduction details, signing secrets, private keys, credentials, or user media in a public issue.

Use GitHub's private security reporting/advisory flow when available. If private reporting is not available, contact the repository owner privately before public disclosure.

## High-priority reports

Treat any issue that can cause source-media deletion, source-media modification, destination overwrite, false duplicate detection, false SD-reuse approval, path traversal outside the selected storage, signing bypass, or arbitrary code execution as release-blocking until investigated.

Never include real customer photos, Apple signing material, Windows signing keys, passwords, tokens, or `.p12` contents in a report.
