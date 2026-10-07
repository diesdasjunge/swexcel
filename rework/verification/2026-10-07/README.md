# Download onboarding verification — 2026-10-07

Status: **downloaded development preview verified on the existing Windows profile; stable 1.0 acceptance incomplete**. The user authorized computer-use continuation of the download and first-use checks. No product code changed.

## Observed through computer use

- Resumed the existing Windows 11 VM and used Windows Edge, File Explorer and Microsoft 365 Excel through Parallels.
- Staged the unchanged `0.1.0-dev.4` Windows development ZIP as an **unpublished GitHub draft**. GitHub reports the uploaded SHA-256 as `c3604d01cb360a809d77c259faa37f21231d3f9047928b06c95b2f1893ba39dd`, matching the local archive. Size: 16,493,542 bytes.
- Downloaded through Edge. Its configured default directory is the shared Mac Downloads folder, so a second download used **Save link as** to `C:\Users\kiryll\Downloads\SWExcel-onboarding.zip` on the Windows disk. Browser settings were not changed.
- Windows archive Properties displayed “This file came from another computer and might be blocked to help protect this computer” and an unchecked **Unblock** checkbox.
- Extracted through Explorer's **Extract All** into `C:\Users\kiryll\Downloads\SWExcel-onboarding`.
- Excel initially rejected opening a second workbook with the same name. Closed the two previously open, saved development workbooks, then opened the newly extracted `SWExcel-0.1.0-dev.4\SWExcel.xlsm`.
- The downloaded workbook displayed the `0.1.0-dev.4` Welcome sheet and the **Protected View** internet-files banner with **Enable Editing**.

## Completed follow-up checks

The user granted file-specific trust confirmation for enabling editing, unblocking the downloaded workbook if needed, and enabling its macros. After the user reported enabling editing, computer use observed Excel's red banner stating that macros were blocked because the file source was untrusted. The user subsequently reported completing the file Properties Unblock and reopen steps. Computer use then observed the dev.4 Welcome sheet without either the Protected View or macro-blocking banner.

Computer-use guest input remained unresponsive. The user authorized alternative automation, and Parallels guest commands connected to the signed-in Windows account in session 1. Excel COM confirmed that the active Recipes sheet belonged to the downloaded workbook, and the workbook no longer had a Zone.Identifier stream.

- The Windows-downloaded ZIP hash independently matched the uploaded/local SHA-256 above.
- Changed Recipes!B10 from 1 to 2: Recipes!B16 changed from 2451545.0007428704 to 2451546.0007428704, exactly one day. Restoring the input restored the original result.
- Explicit smoke execution returned `PASS=18;FAIL=0` with the DLL loaded from this downloaded package's runtime directory.
- Saved and closed the workbook, then ran the full desktop acceptance suite in dedicated Excel instances: **25 passed, 0 failed**, including fresh reopening and calculation persistence. See `downloaded-desktop-acceptance.json`. Its legacy `downloadOnboarding: not tested` field describes that automated suite's scope; the preceding browser/onboarding observations are recorded here separately.
- Full API comparison against the regenerated independent native reference: **106 passed, 0 failed**. API/helper regression: **136 passed, 0 failed**. See `downloaded-full-api.json`, `downloaded-api-regression.json` and `downloaded-native-reference.json`.
- Host checks: **27 passed**; pinned-asset verification passed.
- Excel security remained `AccessVBOM=0`, `VBAWarnings=2`. The PowerShell test process used an execution-policy override; machine/user execution policy was not changed.

This is the existing Windows user profile, not a clean profile. Clean-profile acceptance and stable 1.0 versioning/package publication remain pending. The unchanged dev.4 artifact is suitable for an explicitly labeled development prerelease with these limits. Prior local development checks are recorded separately under `../2026-10-06/`.

An initial all-interface rerun used the old reference JSON: two path outputs differed because the package moved, and two strings were misread by PowerShell's default encoding. The independent native probe was then rerun against the downloaded DLL/data to obtain an ASCII-escaped reference with current paths; numerical tolerances and product source were unchanged.


Release target: https://github.com/diesdasjunge/swexcel/releases/tag/v0.1.0-dev.4
