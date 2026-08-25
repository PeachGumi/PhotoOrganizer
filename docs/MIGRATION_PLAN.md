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

## Phase 4 — unified UI and settings

Reproduce the useful workflows from both applications in Avalonia:

- detected-card summary and media counts;
- destination selection;
- event naming;
- progress and explicit copy-complete vs verified-safe states;
- queued second-card indication;
- RAW extension configuration;
- startup/login behavior through platform adapters.

## Phase 5 — release engineering

Windows:
- production publish target;
- code signing;
- installer/package strategy;
- SmartScreen/reputation acceptance.

macOS:
- universal or supported architecture artifacts;
- Developer ID signing;
- Hardened Runtime;
- notarization and stapling;
- Gatekeeper acceptance.

## Phase 6 — migration acceptance and legacy retirement

Before marking the legacy repositories read-only:

- all shared Core safety tests pass on Windows and macOS;
- platform identity/removal tests pass;
- real SD-card repeated-import acceptance passes on both operating systems;
- filename reuse, same-name/different-bytes, unplug, destination removal, forced-exit, and same-path/storage-replacement cases pass;
- signed production artifacts pass installation and runtime acceptance;
- documentation and customer distribution path point to this repository/product.

Only then archive `PhotoOrganizer-win` and `PhotoOrganizer-mac` with README links to this repository. Do not delete their history.
