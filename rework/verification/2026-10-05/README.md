**Follow-up, 2026-10-06:** The earlier compile gate passed through the user-authorized script (`compile-dev3.json`). Full API completion and current evidence are in [2026-10-06](../2026-10-06/README.md). The record below describes the earlier checkpoint.

# Windows verification — 2026-10-05

Status: **version 0.1.0-dev.3 built and execution-tested; explicit VBE compile remains unverified**.

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

## Excel execution

The initial development build exposed and fixed four issues: PowerShell default
parameter initialization, the COM `Standard` versus `General` format token,
VBE identifier recasing, and regional interpretation of diagnostic strings.
The fixes are version `0.1.0-dev.3`; pre-version-bump testing used the dev.2
package with corrected source.

`SWExcel-verified.xlsm` passed normalized import/export parity and
`PASS=18;FAIL=0`. The first report (`workbook-build-dev2.json`) predates the
text-format correction: some decimal strings in that historical report were
reinterpreted as large integers. The numerical assertions passed, but those
report strings must not be used as numerical evidence. The later acceptance
report verifies the corrected string `280,3689186699` against the direct Double.

`acceptance-dev2.json` records 24 passing checks and zero failures, including all
six direct Excel values within 1e-7 of the freshly compiled swetest's printed
output, interleaved sidereal modes and observers, restoration of default
options, spill unblock/reblock, input/dependent recalculation, a 200-formula
batch (about 0.51 seconds), fresh-session save/reopen, space and München paths,
explicit rejection of a path outside the ANSI code page, missing DLL/data,
rejection of an actual x86 DLL, and different-package collision handling.
These tests do not cover all native functions or all data/date ranges.

The final `SWExcel.xlsm` was rebuilt as **0.1.0-dev.3**. Its normalized source
parity and all 18 integration checks passed (`workbook-build-dev3.json`). The
final `acceptance-dev3.json` records **25 passed, 0 failed**, adding simultaneous
same-package workbooks with different observers to the preceding checks. The
200-formula batch took about 0.67 seconds in this VM; this is a single observed
run, not a performance guarantee. Tracked report copies normalize line endings and trailing whitespace only.
JSON is the authoritative numerical report;
the console log may replace non-ASCII characters during host decoding.

The Welcome sheet was inspected in real Excel. VBE was opened through
Parallels Coherence and Debug/Compile was attempted through the keyboard.
No error dialog was observed, but subsequent inputs did not reliably reach
the guest, and neither the command result nor its disabled state could be
confirmed. **Explicit full-project VBE compilation remains unverified.** The
executed smoke macro proves only its compiled/executed paths. The workbook
was saved and closed; no compile success is inferred from that save.

The user approved temporary VBA-project access. On resumption the VM already
had AccessVBOM=1 and VBAWarnings=1 (all macros enabled), changed during manual
setup. After testing, both settings were restored to AccessVBOM=0 and
VBAWarnings=2 (disable VBA macros with notification), verified by registry
readback in `security-restoration.json`. The pre-existing blank Book1 session
was left open; new Excel sessions use the restored defaults.
Persistent PowerShell policy remains Restricted, with every policy scope
Undefined. Builds used process-only RemoteSigned. Credentials and account
screenshots are excluded. No trusted location or permanent macro exception
was created by this verification work.

The local Windows checkpoint package includes the tested workbook, freshly
built engine, source, data and evidence. `package-files.json` identifies the
final packaged bytes; `evidence/build.json` identifies the workbook at build
time, before acceptance edits/save. Test peer workbooks are kept outside the
deliverable. The source-only ZIP remains a historical preparation artifact.
This checkpoint is not a public release. Full API and downloaded-release
macro onboarding remain pending.
