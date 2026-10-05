# Worksheet and VBA API

The pinned engine has **106 exports**. Each has a reviewed safe interface: **95 `SW_SWE_*` worksheet functions and 11 `SW_CMD_*` VBA commands**. [API-REFERENCE.md](API-REFERENCE.md) lists every signature, input example, buffer contract and indexed result meaning. The family sheets contain editable examples for all 95 worksheet functions. [Verification](../verification/2026-10-06/README.md) separates native comparison, compilation, live formulas and regressions.

## Calling conventions

Supply arguments in the documented order, followed by optional `options` and `detail`. For example:

```excel
=SW_SWE_CALC_UT(2451545,0,258)
=SW_SWE_CALC_UT(2451545,0,258,,TRUE)
=SW_SWE_HOUSES_EX2(2451545,0,52.52,13.405,"P",,TRUE)
=SW_SWE_FIXSTAR2_UT("Spica",2451545,258)
```

Formula separators may be semicolons in your Excel locale. Pass a native input vector as a row or column range with exactly the documented number of elements. Long integers must be exact signed 32-bit values; strings must fit the reviewed NUL-terminated buffer and round-trip through the Windows ANSI code page. The upstream degree formatter alone emits UTF-8, which is decoded explicitly.

Compact results contain the scalar result or one numeric row in documented native output order. Essential scalar return values, such as a node-crossing Julian date, precede output vectors. Updated star names and atmospheric/observer inputs appear only in detail. House cusps start at native index 1; Gauquelin system G returns 36 sectors. Reserved arrays retain their full native allocation even when fewer meaningful values are exposed.

`detail=TRUE` returns field/value/unit rows. It includes status, warning, native return, engine version and absolute path, effective options, inputs and indexed outputs. Functions returning calculation flags report requested and actual ephemeris. Event functions return event flags instead: those flags do **not** identify the actual ephemeris, and the table does not infer an unreported model. Native warnings remain visible. Consult the indexed reference for zero/missing contacts and reserved fields.

| Status | Meaning |
| --- | --- |
| `OK` | Native call succeeded without a warning |
| `WARNING` | Native call returned diagnostic text; inspect detail |
| `FALLBACK` | Returned calculation flags identify another ephemeris model |
| `NO_EVENT` | No event, or circumpolar rise/set result |
| `BELOW_HORIZON` | Visibility calculation cannot observe the object |
| `ERROR` | Validation, loader or native failure |

Compact outputs return `#N/A` for non-OK statuses. Detail retains valid warning/fallback values; failed or unavailable event outputs are `#N/A`, never a successful-looking zero. Pure convenience input errors may return `#VALUE!`. Native zero contact times still mean an unavailable individual contact, as documented upstream.

## Convenience functions

| Formula | Result |
| --- | --- |
| `SW_UTC_JD(year,month,day,hour,minute,second,[offsetHours=0],[detail=FALSE])` | Gregorian local clock plus fractional hours east of UTC -> TT JD, UT1 JD; accepts valid leap seconds |
| `SW_HOUSES(jdUT1,latitude,longitude,[system="P"],[flags=0],[options],[detail=FALSE])` | 12 cusps, or 36 Gauquelin sectors; detail includes angles and daily speeds |
| `SW_POSITIONS(startJD,stepDays,count,body,[flags=258],[options])` | Date series: JD UT1 plus six coordinates/speeds; 1..10000 rows; any failed row fails the compact table |
| `SW_CALC_UT(jdUT1,body,[flags=258],[siderealMode=0],[longitudeEast],[latitudeNorth],[altitudeMetres=0])` | Original six-value convenience result, built-in sidereal modes; use `SW_SWE_CALC_UT` for custom epochs or JPL |
| `SW_LONGITUDE(...)` | Ecliptic longitude in degrees using the original convenience arguments |
| `SW_POSITION_DETAIL(...)` | Original labeled 22-row position result |
| `SW_JULDAY(year,month,day,[hour=0],[calendar=1])` | Validated civil date to JD in the same input time scale; 1 Gregorian, 0 Julian |
| `SW_DEGNORM(value)` | Degrees wrapped into [0,360) |
| `SW_BODY_NAME(body)` | Body name |
| `SW_VERSION()`, `SW_RUNTIME_STATUS()`, `SW_ENGINE_PATH()`, `SW_DATA_PATH()` | Runtime diagnostics |

The Recipes sheet demonstrates the new helpers and explicit options. `SW_JULDAY` does not convert UTC into UT1; use `SW_UTC_JD` or native UTC functions. Native calendar functions support Julian/Gregorian astronomical year numbering. Named time zones and automatic DST are deferred. Array formulas belong in empty grid areas outside Excel Tables; blocking cells produce Excel's normal `#SPILL!`.

## Explicit options

Pass a two-column range/array of key/value pairs. Keys are case-insensitive; unknown or duplicate keys are errors. Every worksheet call reapplies omitted defaults, so previous formulas and VBA setters do not silently configure the next calculation.

| Key | Default and meaning |
| --- | --- |
| `sidereal_mode` | 0 Fagan/Bradley; built-in 0..46 or 255 user-defined, with upstream modifier bits |
| `sidereal_epoch` | 0; JD TT for a user-defined epoch |
| `ayanamsa_epoch` | 0 degrees at that epoch |
| `longitude`, `latitude`, `altitude` | 0; degrees east, north, metres; use topocentric flag 32768 |
| `delta_t` | -1e-10 automatic sentinel; otherwise TT minus UT1 in days |
| `tidal_acceleration` | 999999 automatic sentinel; otherwise arcseconds/century squared |
| `lapse_rate` | 0.0065 K/m |
| `interpolate_nutation` | 0; boolean 0 or 1 |
| `astro_models` | `0,0,0,0,0,0,0,0`; eight pinned engine model IDs, or supported `SE...` version specification; experimental |
| `ephe_path` | Package data directory; one absolute ANSI-compatible directory |
| `jpl_file` | `de431.eph`; filename relative to the chosen data directory, not bundled |
| `sun_declination` | Required for Sunshine `SW_SWE_HOUSE_POS` with system I/i; degrees |

For example, Options!A5:B10 supplies observer and Lahiri settings. Flags 258 request Swiss plus speed; 65794 adds sidereal and 33026 adds topocentric. Defaults produce apparent geocentric tropical longitude/latitude in degrees, distance AU and daily speeds. Equatorial, radians and XYZ flags change the coordinate meaning; retain the flags and consult the native reference. ARMC Sunshine calls take Sun declination in input `ascmc[9]`; the full ten-element vector is required.

## VBA commands and state

`SW_CMD_SET_*` and `SW_CMD_CLOSE` are Subs in a private module and cannot be entered as worksheet formulas. They safely validate/marshal arguments for direct VBA native callers. `SW_CMD_GET_CURRENT_FILE_DATA(ifno)` is an intentionally stateful VBA query: call it immediately after the relevant calculation; slots are 0..4. Worksheet calls reset their explicit options. Closing native data files never unloads the DLL because VBA caches function addresses.

`Native_swe_*` remains the low-level typed ABI. It expects correctly sized caller storage, is hidden from worksheet names, and should only be used by callers that implement the contracts themselves. Borrowed returned strings are bounded copies and never freed. Four upstream experimental exports are labeled in the catalog.

## Data and engine boundaries

The package bundles exactly `sepl_18.se1`, `semo_18.se1`, `seas_18.se1`, plus `sefstars.txt`, `seasnam.txt`, `seorbel.txt`. The binary segment covers the selected 1800..2399 era; individual bodies and calculations need their applicable files and sufficient date margins. Native errors/fallbacks are authoritative at boundaries. Fixed stars use the star catalog; orbital-element fictitious bodies use `seorbel.txt`. Time-only and coordinate utility calls generally need no binary ephemeris, although UT1/TT uses the engine's delta-T/leap-second model.

Eros (ID 10433) needs an external `ast0/se00433.se1`. Other numbered asteroids use `astN/`, N=floor(MPC number/1000), and body ID=10000+MPC number. Raw JPL files are optional external inputs. Place them under the selected ephemeris directory; do not advertise unbundled data as available.

The loader uses a package-relative absolute DLL path, restricted Windows dependency search and version/export/path checks. Keep `SE_EPHE_PATH` unset. Different package DLL copies require separate Excel processes; same-package workbooks share the DLL but reapply options. Paths must fit 242 encoded bytes; lossy ANSI conversion and path-list separators are rejected. Build/package SHA-256 manifests establish file integrity separately from loader checks.
