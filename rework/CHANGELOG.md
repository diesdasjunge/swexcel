# Changelog

All notable changes to the SWExcel rework will be documented here.

The format follows [Keep a Changelog 1.1.0](https://keepachangelog.com/en/1.1.0/), and this project follows [Semantic Versioning 2.0.0](https://semver.org/spec/v2.0.0.html). The legacy project is preserved separately.

## [Unreleased]

## [0.1.0-dev.4] - 2026-10-06

### Added

- Safe interfaces for all 106 pinned exports: 95 worksheet functions and 11 VBA commands, reviewed buffer capacities, ownership, units, status handling, and indexed output documentation.
- Editable examples for every worksheet interface, command recipes, and convenient UTC-offset, house-cusp and date-series helpers.
- Independent Windows C caller with buffer canaries, all-interface numerical comparison, real worksheet execution, and state/calendar/error regressions.
- Explicit full-project compilation through the user-authorized Excel Compile command, with temporary VBA-project access restored afterward.

### Fixed

- Reserve ten native output slots for heliacal event searches even though only three dates are exposed.
- Retain scalar dates from node-crossing calls alongside their coordinates; keep updated input metadata out of compact numeric results.
- Bind worksheet formulas after importing VBA so option ranges calculate correctly from the first saved workbook; add live recipe spill checks.
- Decode the pinned degree formatter's UTF-8 symbol correctly and preserve the distinct ANSI path/string contract.

Public-release download/onboarding acceptance is explicitly deferred. Runtime fixtures establish interface behavior on the recorded Windows host, not every possible astronomical input.

## [0.1.0-dev.3] - 2026-10-05

### Fixed

- Resolve the workbook builder's default source directory after PowerShell parameter binding.
- Preserve Excel's installed general number format instead of assuming the English `General` token.
- Accept VBE's case-insensitive identifier recasing in source parity while preserving string and comment contents; add focused Windows regression checks.
- Store integration diagnostics as text so regional decimal commas and timestamps are not silently reinterpreted.

### Added

- Real Excel workbook generation and all 18 smoke checks, plus 25 passing extended checks and a repeatable desktop acceptance suite for numerical parity, state isolation, spills, recalculation, relocation and loader failures.

Explicit full-project VBE compilation is unverified due to unreliable Parallels UI control. Full API coverage and public/downloaded-release acceptance remain pending.

## [0.1.0-dev.2] - 2026-10-05

### Fixed

- Windows MSVC reference build: the pinned `swetest.c` Windows output branch referenced an undeclared `fp`. Build a copy with that one statement corrected to `fputs(info, stdout)` and record both source hashes and the correction in build provenance. Vendored files and engine calculations remain unchanged.

### Added

- Real Windows ARM64/MSVC x64 build evidence, 106-export verification, and a successful J2000 Sun calculation from the freshly built reference executable. All 22 host checks passed with the fresh binary evidence.

Excel/VBA and worksheet acceptance remain pending; this is not a public workbook release.

## [0.1.0-dev.1] - 2026-10-01

### Added

- Full V1 specification, compatibility/platform/astrology/data roadmap, AGPL license and retained upstream notices.
- All 106 native VBA declarations generated from the pinned public C header and independently checked against the published x64 export table, with an explicit incomplete semantic/runtime catalog.
- Package-relative loader with verified module path, pointer identity and bounded byte strings, native path limits, conflict diagnostics and readiness checks.
- Seven initial calculation/name/calendar worksheet helpers and three diagnostic functions, six-value and labeled-detail spills, explicit state resets and per-result warnings/model provenance.
- Styled guided workbook specification and Windows COM builder with source import/export parity, hash evidence and an explicit integration smoke suite.
- Latest pinned upstream C source, fresh x64 MSVC DLL/swetest build script and source-build provenance gate before workbook construction.
- Exact three-file binary ephemeris bundle, three supporting text catalogs, provenance/coverage/SHA-256 manifest and source-checkpoint ZIP preparation.
- Static ABI, source, packaging and original-preservation checks. Native source compilation and a Sun reference calculation were also checked on macOS.

### Changed

- Raised the rework development version from `0.1.0-dev.0` to `0.1.0-dev.1` after available source checks passed.
- Amended the original 2.10.03 engine pin to the latest official release `v2.10.3bfinal`. Its source version remains `2.10.03`; the unchanged published Windows DLL is an inspection reference until the fresh Windows source build is performed.

Windows/MSVC compilation, VBA compilation, live Excel calculations/spills and public release acceptance are pending. This version is the source checkpoint immediately before the user-requested Parallels test pause.
