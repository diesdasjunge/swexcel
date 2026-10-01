# SWExcel rework

Development checkpoint for a modern Swiss Ephemeris toolkit in 64-bit Microsoft 365 Excel on Windows. The original workbook, VBA and `ephem/` files remain the legacy reference.

**Status: ready to prepare the first Windows integration test; not a public-ready workbook.** No Windows DLL execution, VBA compilation or Excel spill acceptance has been performed at this checkpoint. The actual XLSM must be created with Windows Excel.

The source targets the latest official Swiss Ephemeris release checked on 2026-10-01: [v2.10.3bfinal](https://github.com/aloistr/swisseph/releases/tag/v2.10.3bfinal), commit `f4dcd18e8005dde95fd8a8d2312ed12f9accd1b0`. Its C source version remains `2.10.03`. Its published Windows DLL is identical to the old repository DLL, so the Windows gate first builds an x64 engine from the pinned latest source. `runtimeVersion` alone cannot identify this upstream release.

Read these files in order:

1. [PLAN.md](PLAN.md): complete agreed V1 specification and the latest-engine amendment.
2. [docs/CHECKPOINT.md](docs/CHECKPOINT.md): implemented and pending work, checks and the stop boundary.
3. [docs/WINDOWS-TESTING.md](docs/WINDOWS-TESTING.md): Windows build, compilation and integration procedure.
4. [docs/API.md](docs/API.md): seven initial formulas and native inventory.
5. [ROADMAP.md](ROADMAP.md): compatibility, platform, astrology and data-pack work.

## Prepare on macOS or Windows

From the repository root, using Python 3.10 or newer:

```sh
python3 rework/tools/generate_bindings.py --check
python3 -m unittest discover -s rework/tests -v
python3 rework/tools/prepare.py --verify-only
python3 rework/tools/prepare.py
```

The final command creates `rework/dist/SWExcel-<VERSION>/` and a source-checkpoint ZIP. It contains source, build scripts, documentation, hashes, the published reference binary and the exact V1 data assets. It does not fabricate an XLSM or cached numerical results. Existing generated package directories are never overwritten automatically.

## Layout

| Directory/file | Purpose |
|---|---|
| `src/vba/` | Authoritative native bindings, loader, calculation layer, worksheet functions and smoke checks |
| `api/catalog.json` | All 106 exports with ABI metadata and explicit semantic/runtime completion status |
| `workbook/seed.json` | Guided workbook presentation and initial formulas |
| `tools/` | Binding generation, source packaging and Windows engine/workbook builds |
| `vendor/swisseph/` | Pinned upstream header, C source and notices |
| `vendor/engine/` | Published binaries used for static inspection/reference, with archive provenance |
| `vendor/ephe/` | Pinned text catalogs |
| `package-manifest.json` | Source asset paths, provenance, coverage and SHA-256 hashes |
| `tests/` | Static ABI and packaging checks; these do not prove Excel integration |

The V1 binary data set is exactly `sepl_18.se1`, `semo_18.se1` and `seas_18.se1`; text catalogs are `sefstars.txt`, `seasnam.txt` and `seorbel.txt`. No Eros file is included. Supporting-file sizes are recorded exactly in the manifest; `seasnam.txt` is approximately 16 MB.

Source is licensed AGPL-3.0-or-later; retain [NOTICE.md](NOTICE.md), upstream notices and [LICENSE](LICENSE). The source/build inputs are included in the checkpoint package. Do not publish a source-build DLL without its corresponding pinned source and build instructions.

Versioning follows [Semantic Versioning 2.0.0](https://semver.org/spec/v2.0.0.html), changes follow [Keep a Changelog 1.1.0](https://keepachangelog.com/en/1.1.0/), and commits follow [Conventional Commits 1.0.0](https://www.conventionalcommits.org/en/v1.0.0/). Run appropriate checks, fix failures, update the version and documentation, then commit each completed code checkpoint.
