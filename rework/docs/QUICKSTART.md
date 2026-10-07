# SWExcel 0.1.0-dev.4: Windows preview

Download **SWExcel-0.1.0-dev.4-windows-development.zip** from the [prerelease](https://github.com/diesdasjunge/swexcel/releases/tag/v0.1.0-dev.4). This is a development preview, not stable 1.0. It provides 95 worksheet functions and 11 VBA commands for the pinned Swiss Ephemeris engine.

## Open the workbook

1. Use **64-bit Microsoft 365 Excel on Windows**. Excel for Mac, Excel on the web, 32-bit Excel and other Excel editions are not supported by this preview.
2. Download the ZIP and choose **Extract All**. Keep the entire extracted folder together. Open `SWExcel.xlsm` beside its `runtime` folder; opening the workbook directly inside the ZIP will not work.
3. If Excel shows **Protected View**, choose **Enable Editing** only after checking that you trust this download.
4. If Excel shows the red **Microsoft has blocked macros** banner, close the workbook. In File Explorer, right-click the extracted `SWExcel.xlsm`, choose **Properties**, check **Unblock**, then **Apply** and **OK**. Reopen the workbook and enable its content if prompted. Organizational policy may require your administrator's assistance.
5. Open **Diagnostics** and **Recipes**. Recipes contains editable inputs and live examples for UTC conversion, positions and houses. Changing the day in `Recipes!B10` from 1 to 2 should increase `Recipes!B16` by exactly one day; restore it to 1 afterward.
6. Use **Function Catalog** to find examples and `docs/API-REFERENCE.md` for parameter and output details. Save and reopen your working copy normally.

Keep Excel's global macro protection enabled. Normal workbook use does **not** require Trust access to the VBA project object model or a globally trusted Downloads folder. Macros execute VBA and load the bundled native DLL, so trust only a package whose source you accept.

## Integrity

ZIP size: **16,493,542 bytes**. SHA-256:

```
c3604d01cb360a809d77c259faa37f21231d3f9047928b06c95b2f1893ba39dd
```

The Windows download matched this digest. The archive includes the readable VBA, pinned native source, AGPL license and upstream notices. Preserve those files when redistributing.

## What has been tested

On Windows 11 ARM64 in Parallels with 64-bit-compatible Microsoft 365 Excel build 20430: all 106 native interfaces, 136 API/helper regression cases, 18 smoke checks and 25 desktop checks passed in the development verification. The real browser download, extraction, Excel security banners, file-specific unblocking, live input recalculation and all 25 desktop, 106 native-interface and 136 API/helper checks subsequently passed on the downloaded package.

This used an existing Windows user profile. Clean-profile onboarding remains a stable-release gate. The archive's Welcome sheet and older evidence describe the original development checkpoint; the dated [download verification report](https://github.com/diesdasjunge/swexcel/tree/main/rework/verification/2026-10-07) records the later results.

## Troubleshooting

- **Missing engine/data:** extract the complete archive and retain its folder structure.
- **Conflicting package loaded:** close SWExcel workbooks, exit that Excel instance and reopen the intended package.
- **Path cannot be represented:** move the complete folder to a simple local path containing characters supported by the Windows system code page.
- **Integration!B63 shows #SPILL!:** this is an intentional blocked-spill demonstration. It is not a failed calculation.
- **Eros or additional asteroid data unavailable:** those files are outside this preview's bundled data set.

The preview does not include the future chart-wheel/astrology workspace or automatic named-time-zone/DST conversion.
