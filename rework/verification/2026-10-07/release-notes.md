Windows development preview of the new Swiss Ephemeris toolkit: **95 worksheet functions and 11 VBA commands covering all 106 pinned native exports**, with editable examples, UTC-offset conversion, positions and houses.

Download **SWExcel-0.1.0-dev.4-windows-development.zip**, extract the complete folder, then open `SWExcel.xlsm`. Read **QUICKSTART.md** below for the downloaded-file macro steps. The package includes readable VBA, the corresponding pinned native source and build instructions, ephemeris data, license and notices.

**Supported preview target:** 64-bit Microsoft 365 Excel on Windows. Verified on Windows 11 ARM64 in Parallels with Excel build 20430. Mac/web, 32-bit Excel and other Excel editions are not supported by this preview.

**Verified on the actual browser-downloaded package:** archive hash, extraction, Protected View and macro blocking, file-specific unblocking, live input recalculation, save/reopen, 18 smoke checks, 25 desktop checks, 106 native-interface comparisons and 136 API/helper regression checks. No global macro protections were disabled. [Dated evidence](https://github.com/diesdasjunge/swexcel/tree/main/rework/verification/2026-10-07).

**Development prerelease, not stable 1.0:** testing used the existing Windows user profile; clean-profile onboarding remains pending. The archive is unchanged from the tested dev.4 build, so its Welcome sheet and older evidence retain the original checkpoint wording. The quickstart and dated report provide the current status. This is the API/reference workbook, not the planned full chart-wheel/astrology workspace.

ZIP SHA-256:
`c3604d01cb360a809d77c259faa37f21231d3f9047928b06c95b2f1893ba39dd`

ZIP size: 16,493,542 bytes. Engine source: Swiss Ephemeris `v2.10.3bfinal` (runtime `2.10.03`). AGPL-3.0-or-later; retain the included upstream notices.
