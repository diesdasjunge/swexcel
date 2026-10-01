# Changelog

All notable changes to the SWExcel rework will be documented here.

The format follows [Keep a Changelog 1.1.0](https://keepachangelog.com/en/1.1.0/), and this project follows [Semantic Versioning 2.0.0](https://semver.org/spec/v2.0.0.html). The legacy project is preserved separately.

## [Unreleased]

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
