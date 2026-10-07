# Download onboarding verification — 2026-10-07

Status: **partial; not public-release acceptance**. The user authorized computer-use continuation of the download and first-use checks. No product code changed.

## Observed through computer use

- Resumed the existing Windows 11 VM and used Windows Edge, File Explorer and Microsoft 365 Excel through Parallels.
- Staged the unchanged `0.1.0-dev.4` Windows development ZIP as an **unpublished GitHub draft**. GitHub reports the uploaded SHA-256 as `c3604d01cb360a809d77c259faa37f21231d3f9047928b06c95b2f1893ba39dd`, matching the local archive. Size: 16,493,542 bytes.
- Downloaded through Edge. Its configured default directory is the shared Mac Downloads folder, so a second download used **Save link as** to `C:\Users\kiryll\Downloads\SWExcel-onboarding.zip` on the Windows disk. Browser settings were not changed.
- Windows archive Properties displayed “This file came from another computer and might be blocked to help protect this computer” and an unchecked **Unblock** checkbox.
- Extracted through Explorer's **Extract All** into `C:\Users\kiryll\Downloads\SWExcel-onboarding`.
- Excel initially rejected opening a second workbook with the same name. Closed the two previously open, saved development workbooks, then opened the newly extracted `SWExcel-0.1.0-dev.4\SWExcel.xlsm`.
- The downloaded workbook displayed the `0.1.0-dev.4` Welcome sheet and the **Protected View** internet-files banner with **Enable Editing**.

## Pending boundary

The user granted file-specific trust confirmation for enabling editing, unblocking the downloaded workbook if needed, and enabling its macros. Computer-use clicks and keyboard shortcuts did not visibly reach the Windows guest after this approval, including after reconnecting and resetting the computer-use session. The workbook remained in Protected View. A manual click on Enable Editing was requested to restore progress. This downloaded copy has not yet executed its macros or calculations. Global Excel security settings and VBA-project access were not changed.

This is the existing Windows user profile, not a clean profile. Clean-profile acceptance, the downloaded archive's independently measured hash, post-onboarding live calculations, final release packaging/versioning and public publication remain pending. Prior local development checks are recorded separately under `../2026-10-06/` and do not substitute for these steps.

Draft: https://github.com/diesdasjunge/swexcel/releases/tag/untagged-e68a9e723ff963b213f7
