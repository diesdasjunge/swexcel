# SWExcel Rework: Public Windows Toolkit

## Summary

Create a new toolkit under `/Users/kiryll/Development/swexcel/rework/`, preserving the existing project as a reference.

V1 delivers a downloadable `.xlsm` workbook, Swiss Ephemeris DLL, data, readable VBA source, and documentation. It targets **64-bit Microsoft 365 Excel on Windows** and exposes the full Swiss Ephemeris API through native VBA wrappers and appropriate worksheet functions.

## Calculation and API design

- Pin the latest official release checked on 2026-10-01, **Swiss Ephemeris v2.10.3bfinal** (commit `f4dcd18e8005dde95fd8a8d2312ed12f9accd1b0`). Build correct bindings for all **106 exports** from the [matching upstream header](https://raw.githubusercontent.com/aloistr/swisseph/f4dcd18e8005dde95fd8a8d2312ed12f9accd1b0/swephexp.h), rather than copying the old declarations.
- Maintain an API catalog recording each function’s parameters, units, outputs, buffer sizes, required files, example, and test coverage. Include upstream’s experimental exports, clearly labeled.
- Separate native bindings, calculation wrappers, worksheet functions, and workbook presentation. Use explicit types, correct pointer handling, and checked native return values.
- Expose calculation and conversion functions in cells with an `SW_` prefix. Keep configuration setters and cleanup operations as documented VBA commands.
- Provide both Swiss-style functions and convenient helpers for positions, houses, time conversion, and ephemeris tables.
- Offer scalar numeric results, compact numeric arrays, and optional labeled detail tables. Detail tables report that calculation’s inputs, outputs, units, flags, actual ephemeris, and warnings.
- Keep full precision internally; apply rounding only through cell formatting.
- Friendly defaults: geocentric, tropical, apparent positions; Swiss ephemeris with speed; degrees and AU. Other modes remain explicitly selectable.
- Support UTC and fractional UTC offsets. Distinguish UTC, UT1, and TT; leave named time zones and automatic DST resolution for later.
- Apply formula options consistently so results do not depend on calculation order or previous DLL settings. Surface missing data, fallback models, invalid inputs, and unavailable house systems explicitly.

## Workbook and downloadable package

- Create a guided reference workbook with **Welcome/Setup, Function Catalog, Bodies, Options, Diagnostics**, and worked examples grouped by API family.
- Examples cover positions and date ranges, houses, sidereal calculations, time/calendar conversions, fixed stars, eclipses/occultations, rise/set, heliacal calculations, and coordinate utilities. Advanced examples identify additional file requirements.
- Bundle exactly three binary ephemeris files: `sepl_18.se1`, `semo_18.se1`, and `seas_18.se1`. They provide the selected planets/Moon data and Ceres, Pallas, Juno, Vesta, Chiron, and Pholus; lunar nodes require no separate pack.
- Include supporting catalogs: `sefstars.txt`, `seasnam.txt`, and `seorbel.txt`. Record provenance, hashes, and coverage in a package manifest.
- **Eros and other numbered-body files remain outside the v1 bundle.** Document their required files and directory placement.
- Load the DLL from the package’s location without requiring a fixed `C:\sweph` installation. Use a package-specific DLL filename, verify the loaded engine, and handle conflicting copies explicitly.
- Keep exported VBA source authoritative. Provide a Windows build process that imports it into the workbook, generates the reference/examples, and checks source/workbook synchronization.
- Publish the rework source under **AGPL-3.0-or-later**, retaining upstream notices and corresponding-source instructions. [Swiss Ephemeris licensing](https://github.com/aloistr/swisseph/blob/master/LICENSE)

## Implementation and acceptance

1. **Prove integration first:** In your Parallels Windows Excel installation, verify DLL loading, string/pointer handling, one scalar formula, and one spilling array formula. Record Excel build and process architecture.
2. **Complete API coverage:** Implement and verify every catalog entry, including configuration, diagnostics, and special output buffers.
3. **Build the reference workbook:** Connect documented examples to the tested functions and add setup/data diagnostics.
4. **Validate the packaged release:** Test the actual ZIP after extraction and reopening.

Required checks:

- Compile VBA and exercise all native bindings; compare numerical cases against matching upstream reference calculations.
- Test fractional offsets, leap dates, calendar conversions, angle wraparound, and supported date boundaries.
- Test changing sidereal/topocentric settings, interleaved calculations, repeated recalculation, and native error/no-event results.
- Test missing files, incorrect DLLs, package relocation, spaces/non-ASCII paths, duplicate package copies, blocked spills, and save/reopen behavior.
- Verify catalog coverage, workbook/source parity, and packaged hashes.
- Test downloaded-workbook macro onboarding without requiring users to disable macro protection globally. [Microsoft guidance](https://learn.microsoft.com/en-us/microsoft-365-apps/security/internet-macros-blocked)

Windows Excel acceptance is required before calling v1 ready for public distribution.

## Documentation, versions, and roadmap

During implementation, save the specification and future work in `rework/PLAN.md` and `rework/ROADMAP.md`. Establish project versioning, a Keep a Changelog changelog, and Conventional Commits. At each completed code checkpoint: run relevant tests, fix failures, update version/documentation, then commit. Use development versions before the first public `1.0.0`.

Record these future phases explicitly:

- **V2:** Excel 2024 compatibility; investigate IANA time zones and historical DST support.
- **V3:** Older Excel compatibility, with named supported editions and appropriate array behavior.
- **Later platform path:** Hosted Office add-in with a WebAssembly engine for Mac/web; separately assess offline operation and distribution.
- **Later astrology workspace:** Plan the feature set and implementation for chart wheels, profiles, aspects, comparisons, transit workflows, and Cosmodynes.
- **Later data packs:** Broader date ranges, Eros and additional asteroids/centaurs, planetary moons, and raw JPL files.

V1/V2/V3 describe roadmap phases; actual version numbers follow Semantic Versioning.

## Implementation stop requested by the user

Implement only until real testing via Parallels is needed, then pause and report. The first checkpoint prepares integration step 1: source, native ABI inventory, loader, scalar/spill probes, guided workbook build and pinned package assets. Complete the remaining API/workbook work after the foundational Windows test passes. Do not start or resume the VM during this checkpoint.

## Latest-engine amendment

The user requested the latest Swiss Ephemeris version during implementation. The latest official release is `v2.10.3bfinal`; its source still declares runtime version `2.10.03`. The published Windows archive DLL is byte-identical to the original repository DLL. Therefore build a fresh x64 DLL and matching swetest from the pinned latest source at the Windows gate, retain source/compiler/export/hash evidence, and distinguish release/source identity from the runtime version string. Published prebuilts in the checkpoint are inspection references, not proof of a latest-source build.
