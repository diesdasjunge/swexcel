# Notices and corresponding source

The new SWExcel rework source is licensed under GNU Affero General Public License version 3 or, at your option, any later version (SPDX `AGPL-3.0-or-later`). The full license is in `LICENSE`.

Swiss Ephemeris is copyright Astrodienst AG. Its authors are Dieter Koch and Alois Treindl. Upstream provides an AGPL/professional dual licensing model; this project selects the AGPL route. Original notices are preserved in `vendor/swisseph/LICENSE` and in each vendored source file. See [upstream licensing](https://github.com/aloistr/swisseph/blob/f4dcd18e8005dde95fd8a8d2312ed12f9accd1b0/LICENSE).

The engine source, public headers, supporting catalogs and Windows source-build recipe are pinned to release `v2.10.3bfinal`, commit `f4dcd18e8005dde95fd8a8d2312ed12f9accd1b0`. `vendor/swisseph/source/`, `tools/build-engine.ps1` and their provenance metadata accompany the staged package. Generated worksheet VBA, workbook specification and import/build instructions are also included.

The published upstream prebuilt DLL and swetest executables are retained for static export inspection and reference only. The prebuilt DLL is byte-identical to the legacy binary and is not evidence of a fresh build from the latest source. Before public release, build and verify the engine from the included pinned source, retain its build report/hashes and distribute the corresponding source with it. This checkpoint ZIP is not a public binary release.

The three binary `.se1` data files are copied without modification from the existing `ephem/` assets. Their embedded headers identify DE441 data created on 2026-05-26. The manifest preserves hashes and source provenance separately from the engine source version. Body-specific date limits can be narrower than the file interval; the API/data coverage audit remains part of V1 acceptance.

This notice applies to `rework/`. It does not claim ownership of or relicense the legacy workbook and original VBA. Upstream notices and license choices should accompany all later downloadable releases.
