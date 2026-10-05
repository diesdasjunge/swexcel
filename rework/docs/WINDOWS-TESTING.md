# Windows workbook build and first integration checkpoint

This checkpoint targets **Windows 64-bit Excel for Microsoft 365**. It generates a macro-enabled workbook from the reference sheet specification and imports the current VBA modules. It does not validate all 106 native functions or complete the full API workbook product.

The original spreadsheet is preserved. This workflow creates `SWExcel.xlsm` inside a prepared `rework/dist/SWExcel-<VERSION>` package. No Windows VM, Parallels session or Excel process is launched by the macOS preparation step.

## Requirements

- A logged-on Windows desktop session with installed, licensed **64-bit Excel Microsoft 365**. Record its actual version, build and update channel.
- Windows PowerShell 5.1 or later. The PowerShell process can be 32-bit or 64-bit; Excel's architecture is checked by the executed VBA smoke macro.
- The complete prepared package, including `VERSION`, `package-manifest.json`, `workbook/seed.json`, `api/catalog.json`, `src/vba/*.bas`, `runtime/ephe`, `tools/build-engine.ps1`, and `tools/build-workbook.ps1`.
- The latest pinned upstream engine compiled first with `tools/build-engine.ps1`, using MSVC's **x64 Native Tools** developer environment. The workbook loader's package-relative DLL is recorded in `package-manifest.json` at `engine.file`, currently `runtime/engine/swexcel-se-2.10.3b-x64.dll`.
- For **developer source import only**, Excel Trust Center must already allow **Trust access to the VBA project object model**. The script never changes this setting or the registry. Normal use of the generated workbook does not need VBA project access. An administrator may control this setting.

Use a clean test user profile or a dedicated test Excel process. The builder creates its own Excel COM instance, closes that workbook and quits that instance; it never kills all Excel processes. Microsoft's desktop Office automation requires an interactive user session; running as a service/SYSTEM task is not an equivalent acceptance environment. [Office automation limitations](https://support.microsoft.com/en-us/visio/considerations-for-server-side-automation-of-office).

## Prepare and build

From the repository root, prepare the package using the project Python command:

```sh
python3 rework/tools/prepare.py
```

For this checkpoint `VERSION` is `0.1.0-dev.2`, so the prepared folder is `rework/dist/SWExcel-0.1.0-dev.2`. The prepared package is source and runtime assets, not a Windows-built `.xlsm`.

On Windows, open the **x64 Native Tools Command Prompt for Visual Studio** with the C++ build tools installed. Change to the **prepared package folder**, start PowerShell from that prompt so it inherits the compiler environment, and build the pinned DLL before building the workbook:

```powershell
Get-Content .\VERSION
.\tools\build-engine.ps1 -OutputRoot .\runtime\engine
.\tools\build-workbook.ps1 -PackageDirectory .
```

The engine build compiles release `v2.10.3bfinal`, source commit `f4dcd18e8005dde95fd8a8d2312ed12f9accd1b0`, and writes the DLL, `swetest64.exe` and `runtime/engine/build-provenance.json`. It requires `cl.exe` and `dumpbin.exe`; it does not install a compiler. The workbook builder refuses to proceed without source-build evidence matching the pinned commit, version and source-manifest hash, the actual DLL hash/size, and the exact 106-symbol catalog. A staged published DLL alone cannot pass this gate. Source compilation remains separate from DLL call and Excel acceptance.

The pinned `swetest.c` has an undeclared `fp` in its Windows printing branch. The builder corrects exactly that statement to `fputs(info, stdout)` in `build/engine-x64/swetest-console.c`, preserving the vendor tree. `swetestConsoleFix` in build provenance records the original and compiled source hashes; engine compilation units are unchanged.

The script writes `SWExcel.xlsm`, exports imported modules into a fresh `evidence/<UTC-run-id>/vba-export` directory, and writes `evidence/<UTC-run-id>/build.json`. `evidence/build.json` is the latest build report. VBA source and export hashes are recorded; parity ignores VBE attribute lines, line-ending differences and trailing whitespace only. Unexpected source changes fail the build.

The generated workbook starts on **Welcome** and includes Setup, Function Catalog, Bodies, Options, Diagnostics, Integration, and clearly labeled pending example/roadmap sheets. All 106 catalog symbols are listed as native inventory; unimplemented worksheet interfaces remain pending. The workbook stores formulas with `Formula2`. Ordinary UDF calculations during construction/opening are **not a recorded smoke test**. Without `-RunIntegration`, the check table remains `PENDING`, the report says `not_run`, and compilation/runtime acceptance remains unverified.

The builder refuses to overwrite an existing workbook. Preserve a previous package or choose a different `.xlsm` filename in the same package root:

```powershell
.\tools\build-workbook.ps1 -PackageDirectory . -OutputWorkbook SWExcel-review.xlsm
```

If a script is blocked by the Windows execution policy, use your approved developer policy or administrator guidance; this workflow does not silently change the policy.

## Explicit integration run

To create a new workbook and execute the first smoke checkpoint, use:

```powershell
.\tools\build-workbook.ps1 -PackageDirectory . -OutputWorkbook SWExcel-tested.xlsm -RunIntegration
```

Alternatively, after opening a generated workbook in Excel with trusted macros enabled, invoke `SW_RunIntegrationChecks` from the VBA Immediate window:

```vb
? Application.Run("'SWExcel.xlsm'!SW_RunIntegrationChecks")
```

This is a Function, so it may not appear in Excel's ordinary **Macros** list. A successful summary is `PASS=<count>;FAIL=0`. The macro populates **Integration** and records Excel version/build/OS, the architecture probe, workbook path, runtime data folder and the runtime's actual verified native-module path. A fatal error produces a failed check. It does not run automatically when the workbook opens.

The command-line run additionally records these actuals in `evidence/<UTC-run-id>/integration.json` and in the build report, including failed summaries. A failed run retains the saved workbook and build evidence and returns a failing command. A failure before the macro produces a summary is recorded in `build.json`. A manually executed macro changes the workbook sheet; it does not create an external JSON report, so preserve/save that workbook as evidence.

Checkpoint assertions cover:

- `SW_JULDAY(2000,1,1,12)` returns `2451545` within `1e-9`; `SW_DEGNORM(-1)` returns `359` within `1e-12`.
- `SW_VERSION()` matches **Setup!B8**, populated from `package-manifest.json` at `engine.runtimeVersion`; `SW_BODY_NAME(0)` returns `Sun`. These exercise native string-pointer and buffer handling.
- `SW_CALC_UT(2451545,0)` returns one row with six `Double` values. `SW_LONGITUDE` matches its first value within `1e-9`.
- Sun longitude at J2000 is within `0.01` degrees of the original README's broad `280.369` degree reference. This independent check is not exact pinned-data numerical acceptance.
- February 30 returns an Excel error. An invalid native body (`1000`) also returns an Excel error, followed immediately by a valid Sun calculation; the latter checks recovery after the native error path acquired the calculation guard.
- `SW_POSITION_DETAIL` includes Longitude, Actual flags, Engine path, Warning and Status. Labels are located by name, not a fixed table row count. With the packaged Swiss data, actual flags must be `258`, Status must be `OK`, and longitude/path must match the scalar and verified engine path.
- Worksheet scalar formulas and every calculation-spill value match direct calls. Detail output is read through the actual `SpillingToRange`.
- **B63** deliberately demonstrates `#SPILL!`: **C63** contains a blocker. Error `2045` is an expected passing case. Clearing C63 should allow six-column output; restore the blocker before rerunning the unchanged smoke suite.

The first checkpoint covers seven calculation/name/calendar helpers plus three diagnostic worksheet functions (`SW_RUNTIME_STATUS`, `SW_ENGINE_PATH`, `SW_DATA_PATH`). It is not full API runtime coverage. Successful import/export does not mean VBA compiled; an executed smoke macro establishes only its compiled/executed path. In the VBE, separately run **Debug → Compile VBAProject**, record its result, then save and reopen the workbook in a fresh Excel process.

## Desktop acceptance beyond the smoke suite

Record the package version, source/DLL hashes, expected upstream runtime version, Excel build/channel, host architecture and test date. These checks remain separate, explicit acceptance work:

1. Open/recalculate/save/reopen the complete package in a fresh Excel process. Compare worksheet output, actual flags and numerical reference fixtures.
2. Move the complete package to a new local folder, including a path with spaces. Confirm actual loaded-module/data paths move with it. Test non-ASCII paths separately; Unicode DLL loading does not prove the engine's data-path encoding behavior.
3. Test missing/wrong-architecture DLLs, missing data, and a different same-named DLL already loaded in the **test** Excel process. Errors must be actionable; do not replace modules in a production Excel session.
4. Test simultaneous workbooks with different options and copies in different folders. Verify collision handling and no silent cross-workbook engine/options reuse.
5. Edit formula inputs and options, test ordinary recalculation and full rebuilding, check dependent formulas, block/unblock a spill and measure a representative batch calculation. VBA UDF execution is single-threaded. [Excel performance guidance](https://learn.microsoft.com/en-us/office/vba/excel/concepts/excel-performance/excel-tips-for-optimizing-performance-obstructions).
6. Download the actual release archive into a clean Windows profile and test extraction/onboarding. A local developer copy without internet markings is not download acceptance.

Spilled formulas belong in the worksheet grid outside Excel Tables. `Formula2` avoids implicit-intersection behavior associated with older formula APIs. [Formula2](https://learn.microsoft.com/en-us/office/vba/api/excel.range.formula2), [spill rules](https://support.microsoft.com/en-us/excel/dynamic-array-formulas-and-spilled-array-behavior).

## Downloaded macro onboarding

Windows Microsoft 365 may block downloaded macros because the workbook has **Mark of the Web**. The red risk banner has no **Enable Content** button. After verifying that the package and source are trusted, Microsoft documents **file Properties → General → Unblock**, then reopen the workbook. Organizational policy or Trust Center settings can still block macros. Do not instruct users to enable all macros or make their Downloads folder globally trusted.

Signed VBA with an already trusted publisher is another deployment path; signatures and publisher trust must be actually established and tested before advertising that experience. The ordinary user need not enable **Trust access to the VBA project object model**. [Microsoft's documented macro-download behavior](https://learn.microsoft.com/en-us/microsoft-365-apps/security/internet-macros-blocked).

## Evidence boundaries

- Prepared source/assets: package generation completed; workbook absent until Windows build.
- Saved workbook plus normalized source parity: generation/import verified; full compilation and runtime acceptance unverified.
- Explicit smoke summary with zero failures: first checkpoint passed on the recorded Windows Excel host.
- Manual compile, fresh-session reopening, relocation, collision and downloaded-archive acceptance: separately recorded checks.
- Full API examples/tests, public release, Excel 2024, older Excel, Mac/web and the full astrology workspace: pending or deferred.

After closing Excel and writing its reports, the workbook builder refreshes `package-files.json` with hashes and sizes of the complete current package, including workbook, engine build artifacts and evidence; the inventory excludes itself. `package-manifest.json` retains provenance for the initially staged published reference DLL. For the fresh Windows build, `runtime/engine/build-provenance.json` and the current file inventory record the authoritative DLL hash. A current inventory is integrity evidence, not public-release acceptance.
