# Architecture

## Goal

Photo Organizer is one product with one shared implementation wherever behavior is not inherently OS-specific.

The repository uses C#/.NET 10 and Avalonia for the cross-platform desktop application.

## Layers

```text
PhotoOrganizer.App (Avalonia UI)
        |
        v
PhotoOrganizer.Core
  - media classification
  - scanning semantics
  - duplicate detection
  - verified copy transaction
  - destination byte verification
  - import orchestration (migration target)
        |
        v
Platform services
  - Windows storage/card detection and identity
  - macOS storage/card detection and identity
  - login/startup integration
  - native lifecycle hooks
  - signing/package integration
```

`PhotoOrganizer.Core` must not depend on WinForms, AppKit, SwiftUI, Windows Management APIs, or Avalonia UI types.

## Source repositories

`PeachGumi/PhotoOrganizer-win` is the original Windows implementation and remains a reference for Windows-specific behavior.

`PeachGumi/PhotoOrganizer-mac` contains the most mature data-safety implementation and is the reference for safety semantics during consolidation.

Neither legacy repository is the long-term source of truth after migration acceptance. New cross-platform product behavior belongs here.

## Dependency direction

Platform and UI projects may depend on Core. Core must never depend on a platform or UI project.

Platform-specific code should implement interfaces owned by Core/Application boundaries, rather than adding `OperatingSystem.IsWindows()` or `OperatingSystem.IsMacOS()` branches throughout business logic.

## Versioning

Windows and macOS ship the same product version. A release is not complete unless the shared Core tests pass on both operating systems and each platform artifact passes its signing/package acceptance gates.
