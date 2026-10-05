# Windows verification — 2026-10-05

Status: **native source build passed; Excel acceptance pending**.

The user authorized real Parallels testing, lifting the previous stop. The initial
source checkpoint was `7ca5408a88a66ff59670c9f30f3b38573d80c4f1`, version
`0.1.0-dev.1`. The build correction is version `0.1.0-dev.2`. Original workbook,
legacy files and hash-pinned vendor sources remain unchanged.

## Environment

- Apple Silicon host; Parallels Desktop/guest tools 27.0.1 (58670).
- Windows 11 Pro 25H2, build 26200.9168, ARM64; interactive session 1.
- Microsoft 365 Family Excel installed and opened: 16.0.20430.20092,
  Click-to-Run platform x64, Current Channel GUID
  `492350f6-3a01-4f97-b9c0-c7c6ddf67d60`. Microsoft's `dumpbin` identifies the
  executable as `8664 machine (x64) (ARM64X)`. This does not itself prove
  compatibility with the SWExcel DLL. See `windows-build-session.json` and
  `excel-pe-headers.txt`.
- Visual Studio Build Tools 2022, 17.14.41, MSVC 14.44.35229.0, ARM64 host / x64
  target, Windows SDK 10.0.26100.7705. Installation completed in the UI.
- Official Microsoft bootstrapper SHA-256:
  `985969f472caad75d993a5cb4c35a6a4271460cc12b343e2433b994d173aa990`;
  Authenticode signature was valid and identified Microsoft Corporation.
- Original source ZIP SHA-256:
  `3b92c0a35a631b6270a9051402f74c2abbcc0ec94f1be468aa8eb85ddcba284b`.
  Its 77 inventory files passed hash/size verification after extraction into
  `C:\SWExcel\SWExcel-0.1.0-dev.1`. The corrected build script was then transferred
  there separately. `windows-package-check.json` describes the original transfer,
  not the subsequently built directory.

## Native build and checks

The first MSVC build compiled the DLL but failed compiling `swetest.c:3959`:
`error C2065: 'fp': undeclared identifier`. The pinned Windows output branch
contains `fprintf(fp, info);`. The correction applies exactly once in a generated
build-directory copy, replacing it with `fputs(info, stdout);`. Both hashes and
this exact change are in `build-provenance.json`; no engine calculation source
was patched.

The corrected build completed with exit 0. Both DLL and standalone reference
are AMD64 PE32+ files. All **106** DLL exports match the catalog. See
`engine-build-log.txt` and `build-provenance.json` for compiler, source commit,
arguments, hashes and the output-only correction. The source reports version
`2.10.03`, pinned release `v2.10.3bfinal`, commit
`f4dcd18e8005dde95fd8a8d2312ed12f9accd1b0`.

Fresh `swetest64.exe`, Swiss data, 2000-01-01 12:00 UT:

| Value | Result |
| --- | ---: |
| Sun longitude | 280.3689187 degrees |
| Sun latitude | 0.0002274 degrees |
| Distance | 0.983327625 AU |
| Longitude speed | 1.0194342 degrees/day |

The command exited 0; see `swetest-j2000.txt`. These match the earlier macOS
reference at printed precision. This is standalone engine execution, not a DLL
call from Excel. All **22 host checks passed**, including the formerly skipped
fresh Windows build evidence check. Original-source and legacy hashes passed.

## Pending Excel gate

Workbook generation/import, explicit VBE compilation, the 18 integration checks,
Excel-to-swetest comparison, save/reopen, relocation, failure recovery, conflicting
copies, interleaved options, recalculation and spill tests have **not run**.

Excel currently has `Disable VBA macros with notification` selected and
`Trust access to the VBA project object model` unchecked. Temporary permission
for the latter was requested for source import and remains pending. No trust
setting has been changed. Build commands used process-only `RemoteSigned`;
persistent execution policy remains unchanged. A temporary browser clock error
cleared after time synchronization without bypassing a security interstitial.

Continue with the already installed compiler and the corrected source package.
After VBA import permission is granted, follow `../../docs/WINDOWS-TESTING.md`.
Restore that developer-only permission after testing. Public release acceptance
remains pending. Credentials and account screenshots are excluded from evidence.
