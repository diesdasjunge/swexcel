# Checkpoint and stop boundary

This checkpoint prepares plan step 1, proving integration. The next required work executes Windows tools and real Microsoft 365 Excel in Parallels. The user requested a pause before that work; the VM must not be resumed or used during checkpoint preparation.

## Present

- Complete preserved V1 specification and future roadmap.
- Latest upstream release source pinned to `v2.10.3bfinal`, with headers, notices and Windows source-build inputs. Runtime version string remains `2.10.03`.
- 106 generated native declarations matching the pinned C ABI and published x64 DLL exports; four experimental entries clearly identified.
- Package-relative loader, checked byte buffers/borrowed pointers, explicit calculation state and per-result warnings/model provenance.
- Seven initial calculation/name/calendar functions plus three diagnostic worksheet functions, numeric and detail spills, Windows COM workbook construction and integration smoke source.
- Guided workbook specification, pinned V1 data/catalogs, SHA-256 manifests, source staging and static tests.

## Required Windows gate

Follow `WINDOWS-TESTING.md`: compile the latest engine with x64 MSVC, build/import the VBA workbook, explicitly compile VBA, execute the integration checks, inspect scalar/spill/detail formulas, record Excel build and actual process architecture, compare with freshly built swetest, and save/reopen the package. Neither architecture compatibility nor successful VBA compilation can be inferred from static tests on macOS.

The official latest-release Windows DLL is unchanged from the original binary. A fresh source build is therefore required to verify that the engine update actually uses the latest source. Retain build evidence separately from the runtime version string.

## V1 still pending after the gate

Complete semantic buffer/units/file contracts for all exports; safe wrappers and every appropriate worksheet function; configuration command interfaces; UTC/UT1/TT and fractional offsets; houses and range helpers; all API-family examples; all required numerical/error/state/coverage tests; data/body-specific boundaries; source/workbook parity; downloaded-ZIP macro onboarding, path and conflict cases; final public release acceptance. Do not expand worksheet coverage before resolving foundational loader/ABI/spill failures.

## Verification evidence

Version `0.1.0-dev.1`, verified on 2026-10-01:

- `python3 rework/tools/generate_bindings.py --check`: passed. Pinned header/type/license hashes, PE exports, all 106 generated declarations and catalog agree.
- `python3 rework/tools/prepare.py --verify-only`: passed. All seven initial runtime/data assets match their authoritative hash/size records.
- `python3 -m unittest discover -s rework/tests -v`: 22 tests, 21 passed after package preparation; one explicitly skipped because fresh Windows/MSVC build evidence does not yet exist. The passing package check verifies complete file hashes and runtime assets. Tests also preserve the original workbook, DLL/data and legacy declarations by hash.
- Pinned native C source compiled and linked with macOS Clang into `rework/build/macos/swetest`, using all nine engine compilation units plus upstream `swetest.c`. A J2000 Swiss/DE441 Sun calculation produced longitude `280.3689187` degrees, latitude `0.0002274` degrees, distance `0.983327625` AU and longitude speed `1.0194342` degrees/day. This is source completeness and a native numerical reference, not Windows DLL or Excel validation.
- `git diff --check`: passed before commit. Source-only package and ZIP contents were checked against `package-files.json`; no XLSM or Windows pass report is fabricated.

Windows engine build, VBA compilation, live calculations, rendered workbook inspection and packaged release acceptance remain **pending**. Generated source-only packages carry `checkpoint-status.json` with that distinction. The VM was not started or resumed.
