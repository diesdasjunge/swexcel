# SWExcel roadmap

V1, V2 and V3 are product phases. Releases use Semantic Versioning independently of those labels.

## V1 — public Windows toolkit

64-bit Microsoft 365 Excel on Windows; downloadable XLSM package; full native Swiss Ephemeris coverage, Swiss-style worksheet interfaces and friendly helpers; guided reference sheets. Exactly three binary data files plus `sefstars.txt`, `seasnam.txt` and `seorbel.txt`.

Version 0.1.0-dev.4 implements and exercises all 106 interfaces, fractional-offset UTC/UT1/TT conversion, houses, date-series helpers and worked examples. Downloaded-package onboarding and desktop checks passed on the existing Windows profile on 2026-10-07. Publish the tested development preview with this limitation; clean-profile acceptance remains required before stable 1.0. `PLAN.md` records the complete specification.

## V2 — Excel 2024

Verify and document Excel 2024 compatibility. Investigate IANA named time zones, historical DST, database provenance and updates. Decide whether named-zone support belongs in V2 or V3 after the investigation.

## V3 — older Excel

Name the exact supported editions/builds. Design compatible array entry and output behavior for editions without modern spilling arrays; verify each supported version.

## Later — Mac and web

Investigate a hosted Office add-in: Excel loads a web-hosted extension that adds formulas and reference UI. Explore a WebAssembly Swiss Ephemeris engine for Mac and browser Excel, deployment and updates, licensing/source delivery, data downloads and offline behavior. A downloadable VBA workbook and a hosted add-in have different build and distribution requirements.

## Later — astrology workspace

Create a separate feature and implementation plan for chart wheels, saved profiles, aspects, comparisons, transit workflows and Cosmodynes. Use the legacy workbook as behavioral evidence; do not silently transfer its mutable-global state, indexing quirks or calculation bugs into the toolkit.

## Later — optional data packs

Broader date ranges, Eros and other numbered asteroids/centaurs, planetary moons and raw JPL files. Document per-file/body coverage, size, naming and directories. Eros is not bundled in V1. Numbered asteroid files use `astN/` where N is the MPC number divided by 1000 with the remainder discarded; body ID is 10000 plus MPC number. These packs require their own manifest and verification.
