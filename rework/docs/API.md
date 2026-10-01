# Integration checkpoint API

Seven worksheet functions are implemented as source. Their Windows execution and VBA compilation are pending. All 106 raw native exports have ABI declarations; this is not 106 completed worksheet wrappers. See `api/catalog.json` for reviewed ABI metadata and explicitly pending units, capacities, examples and runtime tests.

| Formula | Result |
|---|---|
| `SW_VERSION()` | Runtime version decoded from a bounded byte buffer after checking the returned pointer identifies that buffer |
| `SW_BODY_NAME(body)` | Name decoded from a bounded byte buffer after checking pointer identity, with package data configured |
| `SW_JULDAY(year,month,day,[hour=0],[calendar=1])` | Civil date to Julian day, checking validity by reverse conversion; Gregorian 1, Julian 0 |
| `SW_DEGNORM(value)` | Degrees normalized to [0,360) |
| `SW_CALC_UT(jdUT1,body,[flags=258],[siderealMode=0],[longitudeEast],[latitudeNorth],[altitudeMetres=0])` | 1 × 6 numeric array |
| `SW_LONGITUDE(...)` | First coordinate, restricted to spherical ecliptic degrees |
| `SW_POSITION_DETAIL(...)` | 22 × 3 table: fields, full-precision values, units, options, actual engine/model and warnings for this calculation |

Default flags 258 mean Swiss ephemeris plus speed, with geocentric tropical apparent positions. Default outputs are longitude/latitude in degrees, distance in AU, angular speeds in degrees/day and distance speed in AU/day. Equatorial, radians and Cartesian flags change coordinate labels and units in detail results. Built-in sidereal mode IDs 0..46 and explicit topocentric observer coordinates are accepted. Custom sidereal epochs and raw JPL configuration are pending.

Each position call reapplies package data path, default astronomical models, automatic delta T/tidal settings, interpolation, sidereal mode and observer coordinates before calling the DLL. VBA UDFs are single-threaded; the calculation layer also rejects nested calls. No worksheet function mutates cells or exposes a recalc-dependent global last error.

Native errors or a model fallback return `#N/A` in numeric helpers. `SW_POSITION_DETAIL` retains the warning and distinguishes `ERROR`, `FALLBACK` and `OK`. A fallback detail table contains its actual values/model; failed native outputs are `#N/A`, rather than misleading zeroes. Pure scalar invalid-input/setup failures return `#VALUE!`; `SW_RUNTIME_STATUS()` and Diagnostics provide setup details. Three additional diagnostic worksheet functions are `SW_RUNTIME_STATUS()`, `SW_ENGINE_PATH()` and `SW_DATA_PATH()`; the internal runtime module stays private.

`SW_JULDAY` preserves the time scale of its input. Decimal hour 12 is not a UTC-to-UT1 conversion. `SW_CALC_UT` takes a UT1 Julian day. Dedicated UTC → UT1/TT conversion and fractional UTC-offset helpers remain required after the first integration gate. Named time zones and DST are later roadmap work.

Use spilling formulas in empty worksheet ranges outside Excel Tables. The Windows builder writes `Formula2`. A blocker produces Excel's normal `#SPILL!` error; clearing the blocker should restore the spill. Numerical values are never rounded in code.

`Native_swe_*` declarations belong to `Option Private Module` and are low-level VBA interfaces; callers must supply valid first elements of correctly sized arrays/byte buffers. They are not safe public worksheet functions. Commands such as `SW_CLOSE` are VBA-only; closing engine data does not unload the DLL, because VBA caches procedure addresses. Further configuration commands and full API family wrappers are pending.

The loader uses an absolute package path, a package-specific DLL filename and restricted Windows dependency search directories. It verifies the actual module path, required bootstrap exports and runtime version. `SE_EPHE_PATH` must be unset. A different package DLL already loaded with the same name requires a separate Excel process. C file paths use the active Windows ANSI code page; lossy encodings, semicolon separators or paths exceeding 242 encoded bytes produce an actionable setup error. Windows UTF-8 ACP uses its required conversion flags. Hash checks are provided by package/build tooling; the loader does not claim to cryptographically authenticate a same-version DLL.
