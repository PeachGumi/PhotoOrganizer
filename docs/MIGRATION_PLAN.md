# Consolidation Plan

## Principle

Migration is behavior-first, not file-copy-first. The existing Windows and macOS applications remain available as reference implementations until the unified application satisfies the same safety and workflow acceptance criteria.

## Phase 1 — shared safety foundation

- establish .NET 10 / Avalonia repository structure;
- define normative data-safety contract;
- port media classification, fail-closed scanning, safe copy transaction, and fresh destination byte verification to shared Core;
- run Core tests on Windows and macOS CI.

## Phase 2 — storage and camera-media platform adapters

Windows:
- detect inserted/removable camera media;
- resolve a selected nested DCIM/PRIVATE directory to the safe card root;
- capture mount-session volume identity that changes on removal/replacement;
- detect source/destination removal.

macOS:
- reproduce the current safe card-root resolution;
- reproduce mount-session volume identity semantics currently provided by Foundation `volumeIdentifier`;
- reproduce mounted-card insertion/removal handling and queued second-card behavior.

No adapter may fall back to path equality as storage identity for reuse approval.

## Phase 3 — shared import orchestrator

Port the macOS fail-closed state machine into shared application logic:

- scan -> ready;
- processing;
- post-copy complete scan;
- fresh real-byte destination verification;
- final storage-identity recheck;
- only then `safe to reuse`.

Destination-inside-source/card rejection and capacity preflight are required before copy begins.

## Phase 4 — unified UI, settings and resident workflow

Reproduce the useful workflows from both applications in Avalonia:

- detected-card summary and media counts;
- destination selection and destination-layout preview;
- event naming;
- progress and explicit copy-complete vs verified-safe states;
- queued second-card indication;
- RAW extension configuration;
- login auto-start through platform adapters;
- a shared Avalonia tray/menu-bar resident mode on Windows and macOS;
- idle window close hides the window while camera-card monitoring remains active;
- tray/menu-bar Show and explicit Quit actions;
- a `StartInBackground` preference independent from login auto-start;
- hidden startup still scans already-mounted camera cards and automatically shows the workflow when a card is detected;
- normal quit is rejected while an import is active.

Safety approval is never persisted with UI preferences. Restarting the process always returns reuse approval to an unverified state.

### Deliberate differences from legacy applications

The unified product does not preserve legacy behavior when that behavior is weaker or ambiguous compared with the shared safety contract.

- The old Windows implementation could overwrite an existing same-named destination file. The unified implementation never deletes or overwrites an existing destination file first; byte-identical content is skipped and different content receives `_2`, `_3`, and so on.
- The old Windows duplicate rule used size plus timestamp. The unified product requires real-byte SHA-256 equality for duplicate and reuse proof.
- The old Windows implementation automatically retried a failed file once. The unified implementation intentionally does not silently retry a failed copy transaction. Any failed supported file blocks the run. The user may run the import again; already completed byte-identical files are safely recognized and skipped, while the failed/new files are attempted again under a fresh transaction and final verification.
- The hardened macOS distinction between `copy complete` and `safe to reuse` is retained and is authoritative for both operating systems.
- Unsupported sidecars and miscellaneous files remain outside the supported import/reuse-verification scope; only JPG/JPEG, configured RAW, MOV and MP4 participate in the safety gate.

## Phase 5 — release engineering

Windows:
- production publish target;
- code signing;
- installer/package strategy;
- SmartScreen/reputation acceptance.

macOS:
- supported architecture artifacts;
- Developer ID signing;
- Hardened Runtime;
- notarization and stapling;
- Gatekeeper acceptance.

The production release workflow is all-or-nothing: unsigned or one-platform-only output must never be substituted for the required signed Windows and macOS artifacts.

## Phase 6 — migration acceptance and legacy retirement

Before marking the legacy repositories read-only:

- all shared Core safety tests pass on Windows and macOS;
- platform identity/removal tests pass;
- tray/menu-bar and background-start behavior passes on both operating systems;
- real SD-card repeated-import acceptance passes on both operating systems;
- filename reuse, same-name/different-bytes, unplug, destination removal, forced-exit, and same-path/storage-replacement cases pass;
- signed production artifacts pass installation and runtime acceptance;
- documentation and customer distribution path point to this repository/product.

Only then archive `PhotoOrganizer-win` and `PhotoOrganizer-mac` with README links to this repository. Do not delete or rewrite their history.
