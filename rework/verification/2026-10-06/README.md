# Full API development verification — 2026-10-06

Version **0.1.0-dev.4** implements the full 106-export interface scope. Public-release download/extraction/macro-onboarding acceptance is **deferred by the user**. These are local development-package results.

## Executed checks

| Check | Result | Evidence |
| --- | --- | --- |
| Full VBA project compile | Passed; Compile command enabled before invocation, disabled afterward | `compile-full-api.json` |
| Imported/exported VBA source parity | All 17 modules match the authoritative source | `build.json` |
| Original integration smoke suite | 18 passed, 0 failed | `integration.json` |
| All native interfaces | 106 passed, 0 failed, including actual VBA command calls | `full-api.json` |
| Independent typed Windows C caller | 106 exports, correct return/output comparisons and intact buffer guards | `windows-native-reference.json`; generator in `tools/generate_native_probe.py` |
| API/helper/state/error examples | 136 passed, 0 failed | `api-regression.json` |
| Desktop acceptance | 25 passed, 0 failed | `desktop-acceptance.json` |
| Host ABI/generation/source/package checks | 27 passed | `host-tests.txt` |

The 136 checks include all 95 live worksheet examples, five Recipes spills, 95 working catalog links, Gregorian/Julian leap dates, a UTC leap second, positive/negative fractional offsets, angle wrapping, Gauquelin's 36 sectors, Sunshine I/i and Savard J houses, polar unavailability, custom sidereal epochs, topocentric/delta-T reset, two-workbook isolation, buffer/string validation, missing JPL/Eros data, date boundaries, no-eclipse and circumpolar no-rise results.

The desktop suite additionally checks fresh reopening, the 200-formula recalculation batch, blocked/unblocked spills, dependent inputs, package relocation, spaces and ANSI-compatible non-ASCII paths, explicit rejection of unsupported Unicode paths, missing files, an x86 DLL, and conflicting package copies. The intentional blocked-spill demonstration is retained in Integration.

The presentation review found that formulas written before VBA import retained bad option-range bindings despite full recalculation. The builder now imports VBA before assigning formulas. The rebuilt Recipes sheet and full native example set are covered by regression checks. Detail blanks remain blank, input cells are identified, long spilled text is fitted, and the catalog links directly to native examples.

## Host, engine and source identity

Windows 11 Pro ARM64 in Parallels; Microsoft 365 Excel version 16.0, build 20430 (installation 16.0.20430.20092, x64-compatible ARM64X). The actual freshly built x64 native engine was loaded and executed. Full environment fields are in `build.json`; the initial engine/compiler provenance is preserved under `../2026-10-05/` and the package's `runtime/engine/build-provenance.json`.

- Pinned release: `v2.10.3bfinal`.
- Source commit: `f4dcd18e8005dde95fd8a8d2312ed12f9accd1b0`.
- Runtime string: `2.10.03`.
- Engine SHA-256: `16a7eedf061f5133a999a72ef87105f3af75fb0efd92532e6a94736191f51736`.

`../current-api.json` ties the 106 verified catalog entries to hashes of the exact 17 source modules and the comparison report. If source hashes change, the generator drops their runtime-verified status until fresh evidence is supplied. Build reports retain their original pre-compilation workbook hashes; `package-attestation.json` records the delivered workbook after saving and verification; `archive-sha256.txt` records the final ZIP hash.

The user authorized scripted compilation and temporary VBA-project access. `AccessVBOM` was restored to 0 afterward; `VBAWarnings` remains 2. Normal workbook users do not need VBA-project access. No global enable-all-macros setting was applied.

## Limits

Coverage means every exported entry has a safe interface and an executed fixture; it does not prove every body/date/flag/house combination or independently validate the astronomical model. The Windows C caller checks VBA marshaling against the same pinned DLL through separately typed native calls. The existing swetest comparison provides another calling path for the six-coordinate Sun fixture.

An exploratory macOS native comparison found differences in fixed-star radial speeds and one occultation azimuth. These also occur in the independently compiled native engines; the Windows VBA/C comparison passes without widening its tolerances. See `cross-platform-notes.json`. Windows is the supported runtime for this package.

Downloaded-file markings, clean-profile macro onboarding and public release are deliberately untested. Excel 2024, older editions, Mac/web, named time zones/DST and optional external data packs remain deferred. The original workbook and vendored sources/data are preserved.
