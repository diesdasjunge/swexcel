# Full native API interfaces

Generated from reviewed contracts and the pinned header. `SW_SWE_*` are worksheet functions.
`SW_CMD_*` are VBA commands; the current-file query is VBA-only in intended use.
Pass input vectors as a row/column range or VBA array. Every worksheet function accepts
optional `options` (a two-column key/value array) and `detail` (TRUE for labeled diagnostics).
Configuration commands affect direct native VBA callers; worksheet calls reapply their explicit options.

Units and error conventions follow the [upstream programming manual](https://www.astro.com/swisseph/swephprg.htm).

## Functions

### SW_SWE_AZALT

Native: `void swe_azalt(double tjd_ut, int32 calc_flag, double *geopos, double atpress, double attemp, double *xin, double *xaz);`

Family: Utilities. Return handling: `value`.

Data: No binary data normally required; delta-T/leap-second model applies to time conversions.

| Argument | Direction | Capacity | Units |
| --- | --- | ---: | --- |
| `tjd_ut` | in | 1 | Julian day UT1 |
| `calc_flag` | in | 1 | coordinate/refraction mode code |
| `geopos` | in | 3 | degrees east; degrees north; metres above sea level |
| `atpress` | in | 1 | hPa (0 estimates pressure) |
| `attemp` | in | 1 | degrees Celsius |
| `xin` | in | 3 | coordinates selected by calc_flag, degrees; optional radius |
| `xaz` | out | 3 | azimuth from south toward west; true altitude; apparent altitude, degrees |

Output `xaz` indices:

- `0`: azimuth south through west degrees
- `1`: true altitude degrees
- `2`: apparent altitude degrees

Example inputs in parameter order: `[2451545.0, 0, [13.405, 52.52, 34.0], 1013.25, 15.0, [280.0, 1.0, 1.0]]`.

### SW_SWE_AZALT_REV

Native: `void swe_azalt_rev(double tjd_ut, int32 calc_flag, double *geopos, double *xin, double *xout);`

Family: Utilities. Return handling: `value`.

Data: No binary data normally required; delta-T/leap-second model applies to time conversions.

| Argument | Direction | Capacity | Units |
| --- | --- | ---: | --- |
| `tjd_ut` | in | 1 | Julian day UT1 |
| `calc_flag` | in | 1 | coordinate/refraction mode code |
| `geopos` | in | 3 | degrees east; degrees north; metres above sea level |
| `xin` | in | 2 | coordinates selected by calc_flag, degrees; optional radius |
| `xout` | out | 2 | longitude/right ascension; latitude/declination, degrees |

Example inputs in parameter order: `[2451545.0, 0, [13.405, 52.52, 34.0], [90.0, 15.0]]`.

### SW_SWE_CALC

Native: `int32 swe_calc(double tjd, int ipl, int32 iflag, double *xx, char *serr);`

Family: Positions. Return handling: `flags`.

Data: Body/date/flags determine files: sepl_18.se1, semo_18.se1, seas_18.se1; external asteroid/JPL files where requested.

| Argument | Direction | Capacity | Units |
| --- | --- | ---: | --- |
| `tjd` | in | 1 | Julian day TT or explicitly named input time scale |
| `ipl` | in | 1 | Swiss body ID |
| `iflag` | in | 1 | Swiss flag bit mask |
| `xx` | out | 6 | coordinate units selected by flags: degrees/radians or AU, followed by per-day speeds |
| `serr` | out | 256 | diagnostic text |

Example inputs in parameter order: `[2451545.0, 0, 258]`.

### SW_SWE_CALC_PCTR

Native: `int32 swe_calc_pctr(double tjd, int32 ipl, int32 iplctr, int32 iflag, double *xxret, char *serr);`

Family: Positions. Return handling: `flags`.

Data: Body/date/flags determine files: sepl_18.se1, semo_18.se1, seas_18.se1; external asteroid/JPL files where requested.

| Argument | Direction | Capacity | Units |
| --- | --- | ---: | --- |
| `tjd` | in | 1 | Julian day TT or explicitly named input time scale |
| `ipl` | in | 1 | Swiss body ID |
| `iplctr` | in | 1 | Swiss centre body ID |
| `iflag` | in | 1 | Swiss flag bit mask |
| `xxret` | out | 6 | coordinate units selected by flags: degrees/radians or AU, followed by per-day speeds |
| `serr` | out | 256 | diagnostic text |

Example inputs in parameter order: `[2451545.0, 0, 4, 258]`.

### SW_SWE_CALC_UT

Native: `int32 swe_calc_ut(double tjd_ut, int32 ipl, int32 iflag, double *xx, char *serr);`

Family: Positions. Return handling: `flags`.

Data: Body/date/flags determine files: sepl_18.se1, semo_18.se1, seas_18.se1; external asteroid/JPL files where requested.

| Argument | Direction | Capacity | Units |
| --- | --- | ---: | --- |
| `tjd_ut` | in | 1 | Julian day UT1 |
| `ipl` | in | 1 | Swiss body ID |
| `iflag` | in | 1 | Swiss flag bit mask |
| `xx` | out | 6 | coordinate units selected by flags: degrees/radians or AU, followed by per-day speeds |
| `serr` | out | 256 | diagnostic text |

Example inputs in parameter order: `[2451545.0, 0, 258]`.

### SW_CMD_CLOSE

Native: `void swe_close(void);`

Family: Configuration. Return handling: `value`.

Data: Configuration/state query; supplied paths/files must exist for subsequent calculations.

| Argument | Direction | Capacity | Units |
| --- | --- | ---: | --- |

Example inputs in parameter order: `[]`.

### SW_SWE_COTRANS

Native: `void swe_cotrans(double *xpo, double *xpn, double eps);`

Family: Utilities. Return handling: `value`.

Data: No binary data normally required; delta-T/leap-second model applies to time conversions.

| Argument | Direction | Capacity | Units |
| --- | --- | ---: | --- |
| `xpo` | in | 3 | longitude/latitude degrees; radius; optional matching daily speeds |
| `xpn` | out | 3 | longitude/latitude degrees; radius; optional matching daily speeds |
| `eps` | in | 1 | degrees |

Example inputs in parameter order: `[[280.0, 1.0, 1.0], 23.4392911]`.

### SW_SWE_COTRANS_SP

Native: `void swe_cotrans_sp(double *xpo, double *xpn, double eps);`

Family: Utilities. Return handling: `value`.

Data: No binary data normally required; delta-T/leap-second model applies to time conversions.

| Argument | Direction | Capacity | Units |
| --- | --- | ---: | --- |
| `xpo` | in | 6 | longitude/latitude degrees; radius; optional matching daily speeds |
| `xpn` | out | 6 | longitude/latitude degrees; radius; optional matching daily speeds |
| `eps` | in | 1 | degrees |

Example inputs in parameter order: `[[280.0, 1.0, 1.0, 1.0, 0.0, 0.0], 23.4392911]`.

### SW_SWE_CS2DEGSTR

Native: `char * swe_cs2degstr(CSEC t, char *a);`

Family: Utilities. Return handling: `value`.

Data: No binary data normally required; delta-T/leap-second model applies to time conversions.

| Argument | Direction | Capacity | Units |
| --- | --- | ---: | --- |
| `t` | in | 1 | centiseconds (1/360000 degree) |
| `a` | out | 256 | formatted text |

Example inputs in parameter order: `[1234567]`.

### SW_SWE_CS2LONLATSTR

Native: `char * swe_cs2lonlatstr(CSEC t, char pchar, char mchar, char *s);`

Family: Utilities. Return handling: `value`.

Data: No binary data normally required; delta-T/leap-second model applies to time conversions.

| Argument | Direction | Capacity | Units |
| --- | --- | ---: | --- |
| `t` | in | 1 | centiseconds (1/360000 degree) |
| `pchar` | in | 1 | positive hemisphere character |
| `mchar` | in | 1 | negative hemisphere character |
| `s` | out | 256 | formatted text |

Example inputs in parameter order: `[1234567, "N", "S"]`.

### SW_SWE_CS2TIMESTR

Native: `char * swe_cs2timestr(CSEC t, int sep, AS_BOOL suppressZero, char *a);`

Family: Time. Return handling: `value`.

Data: No binary data normally required; delta-T/leap-second model applies to time conversions.

| Argument | Direction | Capacity | Units |
| --- | --- | ---: | --- |
| `t` | in | 1 | centiseconds (1/360000 degree) |
| `sep` | in | 1 | separator character code |
| `suppressZero` | in | 1 | C boolean 0 or 1 |
| `a` | out | 256 | formatted text |

Example inputs in parameter order: `[1234567, 58, 0]`.

### SW_SWE_CSNORM

Native: `centisec swe_csnorm(centisec p);`

Family: Utilities. Return handling: `signed`.

Data: No binary data normally required; delta-T/leap-second model applies to time conversions.

| Argument | Direction | Capacity | Units |
| --- | --- | ---: | --- |
| `p` | in | 1 | centiseconds (1/360000 degree) |

Example inputs in parameter order: `[-360000]`.

### SW_SWE_CSROUNDSEC

Native: `centisec swe_csroundsec(centisec x);`

Family: Utilities. Return handling: `signed`.

Data: No binary data normally required; delta-T/leap-second model applies to time conversions.

| Argument | Direction | Capacity | Units |
| --- | --- | ---: | --- |
| `x` | in | 1 | centiseconds (1/360000 degree) |

Example inputs in parameter order: `[1234567]`.

### SW_SWE_D2L

Native: `int32 swe_d2l(double x);`

Family: Utilities. Return handling: `signed`.

Data: No binary data normally required; delta-T/leap-second model applies to time conversions.

| Argument | Direction | Capacity | Units |
| --- | --- | ---: | --- |
| `x` | in | 1 | degrees |

Example inputs in parameter order: `[12.345]`.

### SW_SWE_DATE_CONVERSION

Native: `int swe_date_conversion(int y, int m, int d, double utime, char c, double *tjd);`

Family: Time. Return handling: `status`.

Data: No binary data normally required; delta-T/leap-second model applies to time conversions.

| Argument | Direction | Capacity | Units |
| --- | --- | ---: | --- |
| `y` | in | 1 | astronomical year |
| `m` | in | 1 | month 1..12 |
| `d` | in | 1 | day of month |
| `utime` | in | 1 | decimal hours |
| `c` | in | 1 | g Gregorian or j Julian |
| `tjd` | out | 1 | Julian day TT or explicitly named input time scale |

Example inputs in parameter order: `[2000, 1, 1, 12.0, "g"]`.

### SW_SWE_DAY_OF_WEEK

Native: `int swe_day_of_week(double jd);`

Family: Time. Return handling: `signed`.

Data: No binary data normally required; delta-T/leap-second model applies to time conversions.

| Argument | Direction | Capacity | Units |
| --- | --- | ---: | --- |
| `jd` | in | 1 | Julian day TT or explicitly named input time scale |

Example inputs in parameter order: `[2451545.0]`.

### SW_SWE_DEG_MIDP

Native: `double swe_deg_midp(double x1, double x0);`

Family: Utilities. Return handling: `value`.

Data: No binary data normally required; delta-T/leap-second model applies to time conversions.

| Argument | Direction | Capacity | Units |
| --- | --- | ---: | --- |
| `x1` | in | 1 | degrees |
| `x0` | in | 1 | degrees |

Example inputs in parameter order: `[350.0, 10.0]`.

### SW_SWE_DEGNORM

Native: `double swe_degnorm(double x);`

Family: Utilities. Return handling: `value`.

Data: No binary data normally required; delta-T/leap-second model applies to time conversions.

| Argument | Direction | Capacity | Units |
| --- | --- | ---: | --- |
| `x` | in | 1 | degrees |

Example inputs in parameter order: `[12.345]`.

### SW_SWE_DELTAT

Native: `double swe_deltat(double tjd);`

Family: Time. Return handling: `value`.

Data: No binary data normally required; delta-T/leap-second model applies to time conversions.

| Argument | Direction | Capacity | Units |
| --- | --- | ---: | --- |
| `tjd` | in | 1 | Julian day TT or explicitly named input time scale |

Example inputs in parameter order: `[2451545.0]`.

### SW_SWE_DELTAT_EX

Native: `double swe_deltat_ex(double tjd, int32 iflag, char *serr);`

Family: Time. Return handling: `value`.

Data: No binary data normally required; delta-T/leap-second model applies to time conversions.

| Argument | Direction | Capacity | Units |
| --- | --- | ---: | --- |
| `tjd` | in | 1 | Julian day TT or explicitly named input time scale |
| `iflag` | in | 1 | Swiss flag bit mask |
| `serr` | out | 256 | diagnostic text |

Example inputs in parameter order: `[2451545.0, 258]`.

### SW_SWE_DIFCS2N

Native: `centisec swe_difcs2n(centisec p1, centisec p2);`

Family: Utilities. Return handling: `signed`.

Data: No binary data normally required; delta-T/leap-second model applies to time conversions.

| Argument | Direction | Capacity | Units |
| --- | --- | ---: | --- |
| `p1` | in | 1 | centiseconds (1/360000 degree) |
| `p2` | in | 1 | centiseconds (1/360000 degree) |

Example inputs in parameter order: `[129600000, 360000]`.

### SW_SWE_DIFCSN

Native: `centisec swe_difcsn(centisec p1, centisec p2);`

Family: Utilities. Return handling: `signed`.

Data: No binary data normally required; delta-T/leap-second model applies to time conversions.

| Argument | Direction | Capacity | Units |
| --- | --- | ---: | --- |
| `p1` | in | 1 | centiseconds (1/360000 degree) |
| `p2` | in | 1 | centiseconds (1/360000 degree) |

Example inputs in parameter order: `[129600000, 360000]`.

### SW_SWE_DIFDEG2N

Native: `double swe_difdeg2n(double p1, double p2);`

Family: Utilities. Return handling: `value`.

Data: No binary data normally required; delta-T/leap-second model applies to time conversions.

| Argument | Direction | Capacity | Units |
| --- | --- | ---: | --- |
| `p1` | in | 1 | degrees |
| `p2` | in | 1 | degrees |

Example inputs in parameter order: `[350.0, 10.0]`.

### SW_SWE_DIFDEGN

Native: `double swe_difdegn(double p1, double p2);`

Family: Utilities. Return handling: `value`.

Data: No binary data normally required; delta-T/leap-second model applies to time conversions.

| Argument | Direction | Capacity | Units |
| --- | --- | ---: | --- |
| `p1` | in | 1 | degrees |
| `p2` | in | 1 | degrees |

Example inputs in parameter order: `[350.0, 10.0]`.

### SW_SWE_DIFRAD2N

Native: `double swe_difrad2n(double p1, double p2);`

Family: Utilities. Return handling: `value`.

Data: No binary data normally required; delta-T/leap-second model applies to time conversions.

| Argument | Direction | Capacity | Units |
| --- | --- | ---: | --- |
| `p1` | in | 1 | radians |
| `p2` | in | 1 | radians |

Example inputs in parameter order: `[3.0, -3.0]`.

### SW_SWE_FIXSTAR

Native: `int32 swe_fixstar(char *star, double tjd, int32 iflag, double *xx, char *serr);`

Family: Stars. Return handling: `flags`.

Data: sefstars.txt; selected planetary ephemeris for apparent corrections.

| Argument | Direction | Capacity | Units |
| --- | --- | ---: | --- |
| `star` | inout | 256 | star name or catalog index; canonical name is returned |
| `tjd` | in | 1 | Julian day TT or explicitly named input time scale |
| `iflag` | in | 1 | Swiss flag bit mask |
| `xx` | out | 6 | coordinate units selected by flags: degrees/radians or AU, followed by per-day speeds |
| `serr` | out | 256 | diagnostic text |

Example inputs in parameter order: `["Spica", 2451545.0, 258]`.

### SW_SWE_FIXSTAR2

Native: `int32 swe_fixstar2(char *star, double tjd, int32 iflag, double *xx, char *serr);`

Family: Stars. Return handling: `flags`.

Data: sefstars.txt; selected planetary ephemeris for apparent corrections.

| Argument | Direction | Capacity | Units |
| --- | --- | ---: | --- |
| `star` | inout | 256 | star name or catalog index; canonical name is returned |
| `tjd` | in | 1 | Julian day TT or explicitly named input time scale |
| `iflag` | in | 1 | Swiss flag bit mask |
| `xx` | out | 6 | coordinate units selected by flags: degrees/radians or AU, followed by per-day speeds |
| `serr` | out | 256 | diagnostic text |

Example inputs in parameter order: `["Spica", 2451545.0, 258]`.

### SW_SWE_FIXSTAR2_MAG

Native: `int32 swe_fixstar2_mag(char *star, double *mag, char *serr);`

Family: Stars. Return handling: `status`.

Data: sefstars.txt; selected planetary ephemeris for apparent corrections.

| Argument | Direction | Capacity | Units |
| --- | --- | ---: | --- |
| `star` | inout | 256 | star name or catalog index; canonical name is returned |
| `mag` | out | 1 | astronomical magnitude |
| `serr` | out | 256 | diagnostic text |

Example inputs in parameter order: `["Spica"]`.

### SW_SWE_FIXSTAR2_UT

Native: `int32 swe_fixstar2_ut(char *star, double tjd_ut, int32 iflag, double *xx, char *serr);`

Family: Stars. Return handling: `flags`.

Data: sefstars.txt; selected planetary ephemeris for apparent corrections.

| Argument | Direction | Capacity | Units |
| --- | --- | ---: | --- |
| `star` | inout | 256 | star name or catalog index; canonical name is returned |
| `tjd_ut` | in | 1 | Julian day UT1 |
| `iflag` | in | 1 | Swiss flag bit mask |
| `xx` | out | 6 | coordinate units selected by flags: degrees/radians or AU, followed by per-day speeds |
| `serr` | out | 256 | diagnostic text |

Example inputs in parameter order: `["Spica", 2451545.0, 258]`.

### SW_SWE_FIXSTAR_MAG

Native: `int32 swe_fixstar_mag(char *star, double *mag, char *serr);`

Family: Stars. Return handling: `status`.

Data: sefstars.txt; selected planetary ephemeris for apparent corrections.

| Argument | Direction | Capacity | Units |
| --- | --- | ---: | --- |
| `star` | inout | 256 | star name or catalog index; canonical name is returned |
| `mag` | out | 1 | astronomical magnitude |
| `serr` | out | 256 | diagnostic text |

Example inputs in parameter order: `["Spica"]`.

### SW_SWE_FIXSTAR_UT

Native: `int32 swe_fixstar_ut(char *star, double tjd_ut, int32 iflag, double *xx, char *serr);`

Family: Stars. Return handling: `flags`.

Data: sefstars.txt; selected planetary ephemeris for apparent corrections.

| Argument | Direction | Capacity | Units |
| --- | --- | ---: | --- |
| `star` | inout | 256 | star name or catalog index; canonical name is returned |
| `tjd_ut` | in | 1 | Julian day UT1 |
| `iflag` | in | 1 | Swiss flag bit mask |
| `xx` | out | 6 | coordinate units selected by flags: degrees/radians or AU, followed by per-day speeds |
| `serr` | out | 256 | diagnostic text |

Example inputs in parameter order: `["Spica", 2451545.0, 258]`.

### SW_SWE_GAUQUELIN_SECTOR

Native: `int32 swe_gauquelin_sector(double t_ut, int32 ipl, char *starname, int32 iflag, int32 imeth, double *geopos, double atpress, double attemp, double *dgsect, char *serr);`

Family: Houses. Return handling: `status`.

Data: Date-based houses use obliquity/nutation; ARMC houses need no binary data; Gauquelin body calculations need the selected body files.

| Argument | Direction | Capacity | Units |
| --- | --- | ---: | --- |
| `t_ut` | in | 1 | Julian day UT1 |
| `ipl` | in | 1 | Swiss body ID |
| `starname` | in | 256 | star name, or empty to use ipl |
| `iflag` | in | 1 | Swiss flag bit mask |
| `imeth` | in | 1 | Gauquelin method 0..5 |
| `geopos` | in | 3 | degrees east; degrees north; metres above sea level |
| `atpress` | in | 1 | hPa (0 estimates pressure) |
| `attemp` | in | 1 | degrees Celsius |
| `dgsect` | out | 1 | Gauquelin sector 1..36 |
| `serr` | out | 256 | diagnostic text |

Example inputs in parameter order: `[2451545.0, 0, "", 258, 0, [13.405, 52.52, 34.0], 1013.25, 15.0]`.

### SW_SWE_GET_ASTRO_MODELS

Native: `void swe_get_astro_models(char *samod, char *sdet, int32 iflag);`

Family: Utilities. Return handling: `value`. **Upstream experimental.**

Data: No binary data normally required; delta-T/leap-second model applies to time conversions.

| Argument | Direction | Capacity | Units |
| --- | --- | ---: | --- |
| `samod` | in | 256 | experimental astronomical model/version specification |
| `sdet` | out | 16384 | experimental model-description text |
| `iflag` | in | 1 | Swiss flag bit mask |

Example inputs in parameter order: `["0,0,0,0,0,0,0,0", 258]`.

### SW_SWE_GET_AYANAMSA

Native: `double swe_get_ayanamsa(double tjd_et);`

Family: Positions. Return handling: `value`.

Data: Body/date/flags determine files: sepl_18.se1, semo_18.se1, seas_18.se1; external asteroid/JPL files where requested.

| Argument | Direction | Capacity | Units |
| --- | --- | ---: | --- |
| `tjd_et` | in | 1 | Julian day TT or explicitly named input time scale |

Example inputs in parameter order: `[2451545.0]`.

### SW_SWE_GET_AYANAMSA_EX

Native: `int32 swe_get_ayanamsa_ex(double tjd_et, int32 iflag, double *daya, char *serr);`

Family: Positions. Return handling: `flags`.

Data: Body/date/flags determine files: sepl_18.se1, semo_18.se1, seas_18.se1; external asteroid/JPL files where requested.

| Argument | Direction | Capacity | Units |
| --- | --- | ---: | --- |
| `tjd_et` | in | 1 | Julian day TT or explicitly named input time scale |
| `iflag` | in | 1 | Swiss flag bit mask |
| `daya` | out | 1 | degrees |
| `serr` | out | 256 | diagnostic text |

Example inputs in parameter order: `[2451545.0, 258]`.

### SW_SWE_GET_AYANAMSA_EX_UT

Native: `int32 swe_get_ayanamsa_ex_ut(double tjd_ut, int32 iflag, double *daya, char *serr);`

Family: Positions. Return handling: `flags`.

Data: Body/date/flags determine files: sepl_18.se1, semo_18.se1, seas_18.se1; external asteroid/JPL files where requested.

| Argument | Direction | Capacity | Units |
| --- | --- | ---: | --- |
| `tjd_ut` | in | 1 | Julian day UT1 |
| `iflag` | in | 1 | Swiss flag bit mask |
| `daya` | out | 1 | degrees |
| `serr` | out | 256 | diagnostic text |

Example inputs in parameter order: `[2451545.0, 258]`.

### SW_SWE_GET_AYANAMSA_NAME

Native: `const char * swe_get_ayanamsa_name(int32 isidmode);`

Family: Positions. Return handling: `value`.

Data: Body/date/flags determine files: sepl_18.se1, semo_18.se1, seas_18.se1; external asteroid/JPL files where requested.

| Argument | Direction | Capacity | Units |
| --- | --- | ---: | --- |
| `isidmode` | in | 1 | built-in sidereal ID 0..46 |

Example inputs in parameter order: `[1]`.

### SW_SWE_GET_AYANAMSA_UT

Native: `double swe_get_ayanamsa_ut(double tjd_ut);`

Family: Positions. Return handling: `value`.

Data: Body/date/flags determine files: sepl_18.se1, semo_18.se1, seas_18.se1; external asteroid/JPL files where requested.

| Argument | Direction | Capacity | Units |
| --- | --- | ---: | --- |
| `tjd_ut` | in | 1 | Julian day UT1 |

Example inputs in parameter order: `[2451545.0]`.

### SW_CMD_GET_CURRENT_FILE_DATA

Native: `const char * swe_get_current_file_data(int ifno, double *tfstart, double *tfend, int *denum);`

Family: Configuration. Return handling: `value`.

Data: Configuration/state query; supplied paths/files must exist for subsequent calculations.

| Argument | Direction | Capacity | Units |
| --- | --- | ---: | --- |
| `ifno` | in | 1 | file slot 0..4 |
| `tfstart` | out | 1 | Julian day TT |
| `tfend` | out | 1 | Julian day TT |
| `denum` | out | 1 | JPL DE number |

Example inputs in parameter order: `[0]`.

### SW_SWE_GET_LIBRARY_PATH

Native: `char * swe_get_library_path(char *);`

Family: Utilities. Return handling: `value`.

Data: No binary data normally required; delta-T/leap-second model applies to time conversions.

| Argument | Direction | Capacity | Units |
| --- | --- | ---: | --- |
| `arg1` | out | 256 | NUL-terminated string |

Example inputs in parameter order: `[]`.

### SW_SWE_GET_ORBITAL_ELEMENTS

Native: `int32 swe_get_orbital_elements(double tjd_et, int32 ipl, int32 iflag, double *dret, char *serr);`

Family: Positions. Return handling: `status`.

Data: Body/date/flags determine files: sepl_18.se1, semo_18.se1, seas_18.se1; external asteroid/JPL files where requested.

| Argument | Direction | Capacity | Units |
| --- | --- | ---: | --- |
| `tjd_et` | in | 1 | Julian day TT or explicitly named input time scale |
| `ipl` | in | 1 | Swiss body ID |
| `iflag` | in | 1 | Swiss flag bit mask |
| `dret` | out | 50 | a AU; e; inclination/node/periapsis/anomalies/longitudes degrees; daily motion degrees/day; sidereal/tropical periods years; synodic period days; perihelion epoch JD TT; perihelion/aphelion AU |
| `serr` | out | 256 | diagnostic text |

Output `dret` indices:

- `0`: semimajor axis AU
- `1`: eccentricity
- `2`: inclination degrees
- `3`: ascending node degrees
- `4`: argument of perihelion degrees
- `5`: longitude of perihelion degrees
- `6`: mean anomaly degrees
- `7`: true anomaly degrees
- `8`: eccentric anomaly degrees
- `9`: mean longitude degrees
- `10`: sidereal period tropical years
- `11`: mean daily motion degrees/day
- `12`: tropical period years
- `13`: synodic period days
- `14`: perihelion epoch JD TT
- `15`: perihelion distance AU
- `16`: aphelion distance AU

Example inputs in parameter order: `[2451545.0, 4, 258]`.

### SW_SWE_GET_PLANET_NAME

Native: `char * swe_get_planet_name(int ipl, char *spname);`

Family: Positions. Return handling: `value`.

Data: Body/date/flags determine files: sepl_18.se1, semo_18.se1, seas_18.se1; external asteroid/JPL files where requested.

| Argument | Direction | Capacity | Units |
| --- | --- | ---: | --- |
| `ipl` | in | 1 | Swiss body ID |
| `spname` | out | 256 | body name |

Example inputs in parameter order: `[0]`.

### SW_SWE_GET_TID_ACC

Native: `double swe_get_tid_acc(void);`

Family: Time. Return handling: `value`.

Data: No binary data normally required; delta-T/leap-second model applies to time conversions.

| Argument | Direction | Capacity | Units |
| --- | --- | ---: | --- |

Example inputs in parameter order: `[]`.

### SW_SWE_HELIACAL_ANGLE

Native: `int32 swe_heliacal_angle(double tjdut, double *dgeo, double *datm, double *dobs, int32 helflag, double mag, double azi_obj, double azi_sun, double azi_moon, double alt_moon, double *dret, char *serr);`

Family: Heliacal. Return handling: `status`. **Upstream experimental.**

Data: Body/date/flags determine files: sepl_18.se1, semo_18.se1, seas_18.se1; external asteroid/JPL files where requested.

| Argument | Direction | Capacity | Units |
| --- | --- | ---: | --- |
| `tjdut` | in | 1 | Julian day UT1 |
| `dgeo` | in | 3 | degrees east; degrees north; metres above sea level |
| `datm` | inout | 4 | hPa; Celsius; relative humidity percent; visibility km or extinction coefficient |
| `dobs` | inout | 6 | age years; Snellen ratio; binocular flag; magnification; aperture mm; transmission fraction |
| `helflag` | in | 1 | heliacal flag bit mask |
| `mag` | in | 1 | astronomical magnitude |
| `azi_obj` | in | 1 | azimuth degrees |
| `azi_sun` | in | 1 | azimuth degrees |
| `azi_moon` | in | 1 | azimuth degrees |
| `alt_moon` | in | 1 | altitude degrees |
| `dret` | out | 3 | optimum object altitude; minimum arcus visionis; solar altitude, degrees |
| `serr` | out | 256 | diagnostic text |

Output `dret` indices:

- `0`: optimum object altitude degrees
- `1`: minimum arcus visionis degrees
- `2`: Sun altitude degrees

Example inputs in parameter order: `[2451545.0, [13.405, 52.52, 34.0], [1013.25, 15.0, 40.0, 0.0], [36.0, 1.0, 0.0, 1.0, 0.0, 0.0], 0, 0.0, 90.0, 100.0, 110.0, -10.0]`.

### SW_SWE_HELIACAL_PHENO_UT

Native: `int32 swe_heliacal_pheno_ut(double tjd_ut, double *geopos, double *datm, double *dobs, char *ObjectName, int32 TypeEvent, int32 helflag, double *darr, char *serr);`

Family: Heliacal. Return handling: `status`.

Data: Body/date/flags determine files: sepl_18.se1, semo_18.se1, seas_18.se1; external asteroid/JPL files where requested.

| Argument | Direction | Capacity | Units |
| --- | --- | ---: | --- |
| `tjd_ut` | in | 1 | Julian day UT1 |
| `geopos` | in | 3 | degrees east; degrees north; metres above sea level |
| `datm` | inout | 4 | hPa; Celsius; relative humidity percent; visibility km or extinction coefficient |
| `dobs` | inout | 6 | age years; Snellen ratio; binocular flag; magnification; aperture mm; transmission fraction |
| `ObjectName` | in | 256 | planet or fixed-star name |
| `TypeEvent` | in | 1 | heliacal event type 1..6 |
| `helflag` | in | 1 | heliacal flag bit mask |
| `darr` | out | 50 | indexed heliacal circumstances: degrees, days, magnitudes and dimensionless visibility measures; see field reference |
| `serr` | out | 256 | diagnostic text |

Output `darr` indices:

- `0`: object true altitude degrees
- `1`: object apparent altitude degrees
- `2`: geocentric altitude degrees
- `3`: object azimuth degrees
- `4`: Sun altitude degrees
- `5`: Sun azimuth degrees
- `6`: topocentric arcus visionis degrees
- `7`: geocentric arcus visionis degrees
- `8`: azimuth difference degrees
- `9`: longitude difference degrees
- `10`: extinction coefficient
- `11`: minimum arcus visionis degrees
- `12`: first visibility JD UT1
- `13`: optimum visibility JD UT1
- `14`: last visibility JD UT1
- `15`: Yallop best time JD UT1
- `16`: crescent width degrees
- `17`: Yallop q
- `18`: Yallop criterion
- `19`: parallax degrees
- `20`: magnitude
- `21`: object rise/set JD UT1
- `22`: Sun rise/set JD UT1
- `23`: lag days
- `24`: visibility duration days
- `25`: crescent length degrees
- `26`: elongation degrees
- `27`: illuminated percent

Example inputs in parameter order: `[2451545.0, [13.405, 52.52, 34.0], [1013.25, 15.0, 40.0, 0.0], [36.0, 1.0, 0.0, 1.0, 0.0, 0.0], "venus", 1, 0]`.

### SW_SWE_HELIACAL_UT

Native: `int32 swe_heliacal_ut(double tjdstart_ut, double *geopos, double *datm, double *dobs, char *ObjectName, int32 TypeEvent, int32 iflag, double *dret, char *serr);`

Family: Heliacal. Return handling: `status`.

Data: Body/date/flags determine files: sepl_18.se1, semo_18.se1, seas_18.se1; external asteroid/JPL files where requested.

| Argument | Direction | Capacity | Units |
| --- | --- | ---: | --- |
| `tjdstart_ut` | in | 1 | Julian day UT1 |
| `geopos` | in | 3 | degrees east; degrees north; metres above sea level |
| `datm` | inout | 4 | hPa; Celsius; relative humidity percent; visibility km or extinction coefficient |
| `dobs` | inout | 6 | age years; Snellen ratio; binocular flag; magnification; aperture mm; transmission fraction |
| `ObjectName` | in | 256 | planet or fixed-star name |
| `TypeEvent` | in | 1 | heliacal event type 1..6 |
| `iflag` | in | 1 | Swiss flag bit mask |
| `dret` | out | 10 | Julian day UT1 beginning/optimum/end; unavailable dates are zero |
| `serr` | out | 256 | diagnostic text |

Output `dret` indices:

- `0`: first visibility JD UT1
- `1`: optimum visibility JD UT1
- `2`: last visibility JD UT1

Example inputs in parameter order: `[2451545.0, [13.405, 52.52, 34.0], [1013.25, 15.0, 40.0, 0.0], [36.0, 1.0, 0.0, 1.0, 0.0, 0.0], "venus", 1, 258]`.

### SW_SWE_HELIO_CROSS

Native: `int32 swe_helio_cross(int32 ipl, double x2cross, double jd_et, int32 iflag, int32 dir, double *jd_cross, char *serr);`

Family: Events. Return handling: `status`.

Data: Body/date/flags determine files: sepl_18.se1, semo_18.se1, seas_18.se1; external asteroid/JPL files where requested.

| Argument | Direction | Capacity | Units |
| --- | --- | ---: | --- |
| `ipl` | in | 1 | Swiss body ID |
| `x2cross` | in | 1 | target ecliptic longitude, degrees |
| `jd_et` | in | 1 | Julian day TT or explicitly named input time scale |
| `iflag` | in | 1 | Swiss flag bit mask |
| `dir` | in | 1 | positive forward or negative backward |
| `jd_cross` | out | 1 | Julian day TT or explicitly named input time scale |
| `serr` | out | 256 | diagnostic text |

Example inputs in parameter order: `[4, 0.0, 2451545.0, 258, 1]`.

### SW_SWE_HELIO_CROSS_UT

Native: `int32 swe_helio_cross_ut(int32 ipl, double x2cross, double jd_ut, int32 iflag, int32 dir, double *jd_cross, char *serr);`

Family: Events. Return handling: `status`.

Data: Body/date/flags determine files: sepl_18.se1, semo_18.se1, seas_18.se1; external asteroid/JPL files where requested.

| Argument | Direction | Capacity | Units |
| --- | --- | ---: | --- |
| `ipl` | in | 1 | Swiss body ID |
| `x2cross` | in | 1 | target ecliptic longitude, degrees |
| `jd_ut` | in | 1 | Julian day UT1 |
| `iflag` | in | 1 | Swiss flag bit mask |
| `dir` | in | 1 | positive forward or negative backward |
| `jd_cross` | out | 1 | Julian day UT1 |
| `serr` | out | 256 | diagnostic text |

Example inputs in parameter order: `[4, 0.0, 2451545.0, 258, 1]`.

### SW_SWE_HOUSE_NAME

Native: `const char * swe_house_name(int hsys);`

Family: Houses. Return handling: `value`.

Data: Date-based houses use obliquity/nutation; ARMC houses need no binary data; Gauquelin body calculations need the selected body files.

| Argument | Direction | Capacity | Units |
| --- | --- | ---: | --- |
| `hsys` | in | 1 | one-character house system |

Example inputs in parameter order: `["P"]`.

### SW_SWE_HOUSE_POS

Native: `double swe_house_pos(double armc, double geolat, double eps, int hsys, double *xpin, char *serr);`

Family: Houses. Return handling: `house_position`.

Data: Date-based houses use obliquity/nutation; ARMC houses need no binary data; Gauquelin body calculations need the selected body files.

| Argument | Direction | Capacity | Units |
| --- | --- | ---: | --- |
| `armc` | in | 1 | degrees |
| `geolat` | in | 1 | degrees north |
| `eps` | in | 1 | degrees |
| `hsys` | in | 1 | one-character house system |
| `xpin` | in | 2 | ecliptic longitude and latitude, degrees |
| `serr` | out | 256 | diagnostic text |

Example inputs in parameter order: `[280.0, 52.52, 23.4392911, "P", [280.0, 1.0]]`.

### SW_SWE_HOUSES

Native: `int swe_houses(double tjd_ut, double geolat, double geolon, int hsys, double *cusps, double *ascmc);`

Family: Houses. Return handling: `status`.

Data: Date-based houses use obliquity/nutation; ARMC houses need no binary data; Gauquelin body calculations need the selected body files.

| Argument | Direction | Capacity | Units |
| --- | --- | ---: | --- |
| `tjd_ut` | in | 1 | Julian day UT1 |
| `geolat` | in | 1 | degrees north |
| `geolon` | in | 1 | degrees east |
| `hsys` | in | 1 | one-character house system |
| `cusps` | out | 37 | degrees (radians with SEFLG_RADIANS); index 1..12 or 1..36 for G |
| `ascmc` | out | 10 | degrees (radians with SEFLG_RADIANS) |

Output `ascmc` indices:

- `0`: Ascendant; degrees (radians with SEFLG_RADIANS)
- `1`: Midheaven; degrees (radians with SEFLG_RADIANS)
- `2`: ARMC; degrees (radians with SEFLG_RADIANS)
- `3`: Vertex; degrees (radians with SEFLG_RADIANS)
- `4`: equatorial Ascendant; degrees (radians with SEFLG_RADIANS)
- `5`: Koch co-Ascendant; degrees (radians with SEFLG_RADIANS)
- `6`: Munkasey co-Ascendant; degrees (radians with SEFLG_RADIANS)
- `7`: polar Ascendant; degrees (radians with SEFLG_RADIANS)

Example inputs in parameter order: `[2451545.0, 52.52, 13.405, "P"]`.

### SW_SWE_HOUSES_ARMC

Native: `int swe_houses_armc(double armc, double geolat, double eps, int hsys, double *cusps, double *ascmc);`

Family: Houses. Return handling: `status`.

Data: Date-based houses use obliquity/nutation; ARMC houses need no binary data; Gauquelin body calculations need the selected body files.

| Argument | Direction | Capacity | Units |
| --- | --- | ---: | --- |
| `armc` | in | 1 | degrees |
| `geolat` | in | 1 | degrees north |
| `eps` | in | 1 | degrees |
| `hsys` | in | 1 | one-character house system |
| `cusps` | out | 37 | degrees (radians with SEFLG_RADIANS); index 1..12 or 1..36 for G |
| `ascmc` | inout | 10 | degrees (radians with SEFLG_RADIANS) |

Output `ascmc` indices:

- `0`: Ascendant; degrees (radians with SEFLG_RADIANS)
- `1`: Midheaven; degrees (radians with SEFLG_RADIANS)
- `2`: ARMC; degrees (radians with SEFLG_RADIANS)
- `3`: Vertex; degrees (radians with SEFLG_RADIANS)
- `4`: equatorial Ascendant; degrees (radians with SEFLG_RADIANS)
- `5`: Koch co-Ascendant; degrees (radians with SEFLG_RADIANS)
- `6`: Munkasey co-Ascendant; degrees (radians with SEFLG_RADIANS)
- `7`: polar Ascendant; degrees (radians with SEFLG_RADIANS)
- `8`: reserved; degrees (radians with SEFLG_RADIANS)
- `9`: Sun declination (Sunshine ARMC input); degrees (radians with SEFLG_RADIANS)

Example inputs in parameter order: `[280.0, 52.52, 23.4392911, "P", [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0]]`.

### SW_SWE_HOUSES_ARMC_EX2

Native: `int swe_houses_armc_ex2(double armc, double geolat, double eps, int hsys, double *cusps, double *ascmc, double *cusp_speed, double *ascmc_speed, char *serr);`

Family: Houses. Return handling: `status`.

Data: Date-based houses use obliquity/nutation; ARMC houses need no binary data; Gauquelin body calculations need the selected body files.

| Argument | Direction | Capacity | Units |
| --- | --- | ---: | --- |
| `armc` | in | 1 | degrees |
| `geolat` | in | 1 | degrees north |
| `eps` | in | 1 | degrees |
| `hsys` | in | 1 | one-character house system |
| `cusps` | out | 37 | degrees (radians with SEFLG_RADIANS); index 1..12 or 1..36 for G |
| `ascmc` | inout | 10 | degrees (radians with SEFLG_RADIANS) |
| `cusp_speed` | out | 37 | degrees/day (radians/day with SEFLG_RADIANS) |
| `ascmc_speed` | out | 10 | degrees/day (radians/day with SEFLG_RADIANS) |
| `serr` | out | 256 | diagnostic text |

Output `ascmc` indices:

- `0`: Ascendant; degrees (radians with SEFLG_RADIANS)
- `1`: Midheaven; degrees (radians with SEFLG_RADIANS)
- `2`: ARMC; degrees (radians with SEFLG_RADIANS)
- `3`: Vertex; degrees (radians with SEFLG_RADIANS)
- `4`: equatorial Ascendant; degrees (radians with SEFLG_RADIANS)
- `5`: Koch co-Ascendant; degrees (radians with SEFLG_RADIANS)
- `6`: Munkasey co-Ascendant; degrees (radians with SEFLG_RADIANS)
- `7`: polar Ascendant; degrees (radians with SEFLG_RADIANS)
- `8`: reserved; degrees (radians with SEFLG_RADIANS)
- `9`: Sun declination (Sunshine ARMC input); degrees (radians with SEFLG_RADIANS)

Output `ascmc_speed` indices:

- `0`: Ascendant daily speed; degrees/day (radians/day with SEFLG_RADIANS)
- `1`: Midheaven daily speed; degrees/day (radians/day with SEFLG_RADIANS)
- `2`: ARMC daily speed; degrees/day (radians/day with SEFLG_RADIANS)
- `3`: Vertex daily speed; degrees/day (radians/day with SEFLG_RADIANS)
- `4`: equatorial Ascendant daily speed; degrees/day (radians/day with SEFLG_RADIANS)
- `5`: Koch co-Ascendant daily speed; degrees/day (radians/day with SEFLG_RADIANS)
- `6`: Munkasey co-Ascendant daily speed; degrees/day (radians/day with SEFLG_RADIANS)
- `7`: polar Ascendant daily speed; degrees/day (radians/day with SEFLG_RADIANS)

Example inputs in parameter order: `[280.0, 52.52, 23.4392911, "P", [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0]]`.

### SW_SWE_HOUSES_EX

Native: `int swe_houses_ex(double tjd_ut, int32 iflag, double geolat, double geolon, int hsys, double *cusps, double *ascmc);`

Family: Houses. Return handling: `status`.

Data: Date-based houses use obliquity/nutation; ARMC houses need no binary data; Gauquelin body calculations need the selected body files.

| Argument | Direction | Capacity | Units |
| --- | --- | ---: | --- |
| `tjd_ut` | in | 1 | Julian day UT1 |
| `iflag` | in | 1 | Swiss flag bit mask |
| `geolat` | in | 1 | degrees north |
| `geolon` | in | 1 | degrees east |
| `hsys` | in | 1 | one-character house system |
| `cusps` | out | 37 | degrees (radians with SEFLG_RADIANS); index 1..12 or 1..36 for G |
| `ascmc` | out | 10 | degrees (radians with SEFLG_RADIANS) |

Output `ascmc` indices:

- `0`: Ascendant; degrees (radians with SEFLG_RADIANS)
- `1`: Midheaven; degrees (radians with SEFLG_RADIANS)
- `2`: ARMC; degrees (radians with SEFLG_RADIANS)
- `3`: Vertex; degrees (radians with SEFLG_RADIANS)
- `4`: equatorial Ascendant; degrees (radians with SEFLG_RADIANS)
- `5`: Koch co-Ascendant; degrees (radians with SEFLG_RADIANS)
- `6`: Munkasey co-Ascendant; degrees (radians with SEFLG_RADIANS)
- `7`: polar Ascendant; degrees (radians with SEFLG_RADIANS)

Example inputs in parameter order: `[2451545.0, 258, 52.52, 13.405, "P"]`.

### SW_SWE_HOUSES_EX2

Native: `int swe_houses_ex2(double tjd_ut, int32 iflag, double geolat, double geolon, int hsys, double *cusps, double *ascmc, double *cusp_speed, double *ascmc_speed, char *serr);`

Family: Houses. Return handling: `status`.

Data: Date-based houses use obliquity/nutation; ARMC houses need no binary data; Gauquelin body calculations need the selected body files.

| Argument | Direction | Capacity | Units |
| --- | --- | ---: | --- |
| `tjd_ut` | in | 1 | Julian day UT1 |
| `iflag` | in | 1 | Swiss flag bit mask |
| `geolat` | in | 1 | degrees north |
| `geolon` | in | 1 | degrees east |
| `hsys` | in | 1 | one-character house system |
| `cusps` | out | 37 | degrees (radians with SEFLG_RADIANS); index 1..12 or 1..36 for G |
| `ascmc` | out | 10 | degrees (radians with SEFLG_RADIANS) |
| `cusp_speed` | out | 37 | degrees/day (radians/day with SEFLG_RADIANS) |
| `ascmc_speed` | out | 10 | degrees/day (radians/day with SEFLG_RADIANS) |
| `serr` | out | 256 | diagnostic text |

Output `ascmc` indices:

- `0`: Ascendant; degrees (radians with SEFLG_RADIANS)
- `1`: Midheaven; degrees (radians with SEFLG_RADIANS)
- `2`: ARMC; degrees (radians with SEFLG_RADIANS)
- `3`: Vertex; degrees (radians with SEFLG_RADIANS)
- `4`: equatorial Ascendant; degrees (radians with SEFLG_RADIANS)
- `5`: Koch co-Ascendant; degrees (radians with SEFLG_RADIANS)
- `6`: Munkasey co-Ascendant; degrees (radians with SEFLG_RADIANS)
- `7`: polar Ascendant; degrees (radians with SEFLG_RADIANS)

Output `ascmc_speed` indices:

- `0`: Ascendant daily speed; degrees/day (radians/day with SEFLG_RADIANS)
- `1`: Midheaven daily speed; degrees/day (radians/day with SEFLG_RADIANS)
- `2`: ARMC daily speed; degrees/day (radians/day with SEFLG_RADIANS)
- `3`: Vertex daily speed; degrees/day (radians/day with SEFLG_RADIANS)
- `4`: equatorial Ascendant daily speed; degrees/day (radians/day with SEFLG_RADIANS)
- `5`: Koch co-Ascendant daily speed; degrees/day (radians/day with SEFLG_RADIANS)
- `6`: Munkasey co-Ascendant daily speed; degrees/day (radians/day with SEFLG_RADIANS)
- `7`: polar Ascendant daily speed; degrees/day (radians/day with SEFLG_RADIANS)

Example inputs in parameter order: `[2451545.0, 258, 52.52, 13.405, "P"]`.

### SW_SWE_JDET_TO_UTC

Native: `void swe_jdet_to_utc(double tjd_et, int32 gregflag, int32 *iyear, int32 *imonth, int32 *iday, int32 *ihour, int32 *imin, double *dsec);`

Family: Time. Return handling: `value`.

Data: No binary data normally required; delta-T/leap-second model applies to time conversions.

| Argument | Direction | Capacity | Units |
| --- | --- | ---: | --- |
| `tjd_et` | in | 1 | Julian day TT or explicitly named input time scale |
| `gregflag` | in | 1 | 0 Julian, 1 Gregorian |
| `iyear` | out | 1 | astronomical year |
| `imonth` | out | 1 | month 1..12 |
| `iday` | out | 1 | day of month |
| `ihour` | out | 1 | hour 0..23 |
| `imin` | out | 1 | minutes |
| `dsec` | out | 1 | seconds, leap second allowed where documented |

Example inputs in parameter order: `[2451545.0, 1]`.

### SW_SWE_JDUT1_TO_UTC

Native: `void swe_jdut1_to_utc(double tjd_ut, int32 gregflag, int32 *iyear, int32 *imonth, int32 *iday, int32 *ihour, int32 *imin, double *dsec);`

Family: Time. Return handling: `value`.

Data: No binary data normally required; delta-T/leap-second model applies to time conversions.

| Argument | Direction | Capacity | Units |
| --- | --- | ---: | --- |
| `tjd_ut` | in | 1 | Julian day UT1 |
| `gregflag` | in | 1 | 0 Julian, 1 Gregorian |
| `iyear` | out | 1 | astronomical year |
| `imonth` | out | 1 | month 1..12 |
| `iday` | out | 1 | day of month |
| `ihour` | out | 1 | hour 0..23 |
| `imin` | out | 1 | minutes |
| `dsec` | out | 1 | seconds, leap second allowed where documented |

Example inputs in parameter order: `[2451545.0, 1]`.

### SW_SWE_JULDAY

Native: `double swe_julday(int year, int month, int day, double hour, int gregflag);`

Family: Time. Return handling: `value`.

Data: No binary data normally required; delta-T/leap-second model applies to time conversions.

| Argument | Direction | Capacity | Units |
| --- | --- | ---: | --- |
| `year` | in | 1 | astronomical year |
| `month` | in | 1 | month 1..12 |
| `day` | in | 1 | day of month |
| `hour` | in | 1 | decimal hours |
| `gregflag` | in | 1 | 0 Julian, 1 Gregorian |

Example inputs in parameter order: `[2000, 1, 1, 12.0, 1]`.

### SW_SWE_LAT_TO_LMT

Native: `int32 swe_lat_to_lmt(double tjd_lat, double geolon, double *tjd_lmt, char *serr);`

Family: Time. Return handling: `status`.

Data: No binary data normally required; delta-T/leap-second model applies to time conversions.

| Argument | Direction | Capacity | Units |
| --- | --- | ---: | --- |
| `tjd_lat` | in | 1 | Julian day TT or explicitly named input time scale |
| `geolon` | in | 1 | degrees east |
| `tjd_lmt` | out | 1 | Julian day TT or explicitly named input time scale |
| `serr` | out | 256 | diagnostic text |

Example inputs in parameter order: `[2451545.0, 13.405]`.

### SW_SWE_LMT_TO_LAT

Native: `int32 swe_lmt_to_lat(double tjd_lmt, double geolon, double *tjd_lat, char *serr);`

Family: Time. Return handling: `status`.

Data: No binary data normally required; delta-T/leap-second model applies to time conversions.

| Argument | Direction | Capacity | Units |
| --- | --- | ---: | --- |
| `tjd_lmt` | in | 1 | Julian day TT or explicitly named input time scale |
| `geolon` | in | 1 | degrees east |
| `tjd_lat` | out | 1 | Julian day TT or explicitly named input time scale |
| `serr` | out | 256 | diagnostic text |

Example inputs in parameter order: `[2451545.0, 13.405]`.

### SW_SWE_LUN_ECLIPSE_HOW

Native: `int32 swe_lun_eclipse_how(double tjd_ut, int32 ifl, double *geopos, double *attr, char *serr);`

Family: Events. Return handling: `event`.

Data: Body/date/flags determine files: sepl_18.se1, semo_18.se1, seas_18.se1; external asteroid/JPL files where requested.

| Argument | Direction | Capacity | Units |
| --- | --- | ---: | --- |
| `tjd_ut` | in | 1 | Julian day UT1 |
| `ifl` | in | 1 | ephemeris flag bit mask |
| `geopos` | in | 3 | degrees east; degrees north; metres above sea level |
| `attr` | out | 20 | indexed event/phenomena attributes; see field reference |
| `serr` | out | 256 | diagnostic text |

Output `attr` indices:

- `0`: umbral magnitude
- `1`: penumbral magnitude
- `2`: reserved
- `3`: reserved
- `4`: Moon azimuth degrees
- `5`: true altitude degrees
- `6`: apparent altitude degrees
- `7`: opposition distance degrees
- `8`: eclipse magnitude
- `9`: Saros series
- `10`: Saros member

Example inputs in parameter order: `[2451545.0, 2, [13.405, 52.52, 34.0]]`.

### SW_SWE_LUN_ECLIPSE_WHEN

Native: `int32 swe_lun_eclipse_when(double tjd_start, int32 ifl, int32 ifltype, double *tret, int32 backward, char *serr);`

Family: Events. Return handling: `event`.

Data: Body/date/flags determine files: sepl_18.se1, semo_18.se1, seas_18.se1; external asteroid/JPL files where requested.

| Argument | Direction | Capacity | Units |
| --- | --- | ---: | --- |
| `tjd_start` | in | 1 | Julian day UT1 |
| `ifl` | in | 1 | ephemeris flag bit mask |
| `ifltype` | in | 1 | event type bit mask |
| `tret` | out | 10 | Julian day UT1; 0 means unavailable contact |
| `backward` | in | 1 | 0 forward, 1 backward; occultation one-try bit 32768 |
| `serr` | out | 256 | diagnostic text |

Output `tret` indices:

- `0`: maximum JD UT1
- `1`: reserved
- `2`: partial begins JD UT1
- `3`: partial ends JD UT1
- `4`: total begins JD UT1
- `5`: total ends JD UT1
- `6`: penumbral begins JD UT1
- `7`: penumbral ends JD UT1
- `8`: moonrise JD UT1 (local only)
- `9`: moonset JD UT1 (local only)

Example inputs in parameter order: `[2451545.0, 2, 0, 0]`.

### SW_SWE_LUN_ECLIPSE_WHEN_LOC

Native: `int32 swe_lun_eclipse_when_loc(double tjd_start, int32 ifl, double *geopos, double *tret, double *attr, int32 backward, char *serr);`

Family: Events. Return handling: `event`.

Data: Body/date/flags determine files: sepl_18.se1, semo_18.se1, seas_18.se1; external asteroid/JPL files where requested.

| Argument | Direction | Capacity | Units |
| --- | --- | ---: | --- |
| `tjd_start` | in | 1 | Julian day UT1 |
| `ifl` | in | 1 | ephemeris flag bit mask |
| `geopos` | in | 3 | degrees east; degrees north; metres above sea level |
| `tret` | out | 10 | Julian day UT1; 0 means unavailable contact |
| `attr` | out | 20 | indexed event/phenomena attributes; see field reference |
| `backward` | in | 1 | 0 forward, 1 backward; occultation one-try bit 32768 |
| `serr` | out | 256 | diagnostic text |

Output `tret` indices:

- `0`: maximum JD UT1
- `1`: reserved
- `2`: partial begins JD UT1
- `3`: partial ends JD UT1
- `4`: total begins JD UT1
- `5`: total ends JD UT1
- `6`: penumbral begins JD UT1
- `7`: penumbral ends JD UT1
- `8`: moonrise JD UT1 (local only)
- `9`: moonset JD UT1 (local only)

Output `attr` indices:

- `0`: umbral magnitude
- `1`: penumbral magnitude
- `2`: reserved
- `3`: reserved
- `4`: Moon azimuth degrees
- `5`: true altitude degrees
- `6`: apparent altitude degrees
- `7`: opposition distance degrees
- `8`: eclipse magnitude
- `9`: Saros series
- `10`: Saros member

Example inputs in parameter order: `[2451545.0, 2, [13.405, 52.52, 34.0], 0]`.

### SW_SWE_LUN_OCCULT_WHEN_GLOB

Native: `int32 swe_lun_occult_when_glob(double tjd_start, int32 ipl, char *starname, int32 ifl, int32 ifltype, double *tret, int32 backward, char *serr);`

Family: Events. Return handling: `event`.

Data: Body/date/flags determine files: sepl_18.se1, semo_18.se1, seas_18.se1; external asteroid/JPL files where requested.

| Argument | Direction | Capacity | Units |
| --- | --- | ---: | --- |
| `tjd_start` | in | 1 | Julian day UT1 |
| `ipl` | in | 1 | Swiss body ID |
| `starname` | in | 256 | star name, or empty to use ipl |
| `ifl` | in | 1 | ephemeris flag bit mask |
| `ifltype` | in | 1 | event type bit mask |
| `tret` | out | 10 | Julian day UT1; 0 means unavailable contact |
| `backward` | in | 1 | 0 forward, 1 backward; occultation one-try bit 32768 |
| `serr` | out | 256 | diagnostic text |

Output `tret` indices:

- `0`: maximum JD UT1
- `1`: local apparent noon JD UT1
- `2`: eclipse begins JD UT1
- `3`: eclipse ends JD UT1
- `4`: totality begins JD UT1
- `5`: totality ends JD UT1
- `6`: centre line begins JD UT1
- `7`: centre line ends JD UT1
- `8`: reserved unimplemented
- `9`: reserved unimplemented

Example inputs in parameter order: `[2451545.0, 3, "", 2, 0, 0]`.

### SW_SWE_LUN_OCCULT_WHEN_LOC

Native: `int32 swe_lun_occult_when_loc(double tjd_start, int32 ipl, char *starname, int32 ifl, double *geopos, double *tret, double *attr, int32 backward, char *serr);`

Family: Events. Return handling: `event`.

Data: Body/date/flags determine files: sepl_18.se1, semo_18.se1, seas_18.se1; external asteroid/JPL files where requested.

| Argument | Direction | Capacity | Units |
| --- | --- | ---: | --- |
| `tjd_start` | in | 1 | Julian day UT1 |
| `ipl` | in | 1 | Swiss body ID |
| `starname` | in | 256 | star name, or empty to use ipl |
| `ifl` | in | 1 | ephemeris flag bit mask |
| `geopos` | in | 3 | degrees east; degrees north; metres above sea level |
| `tret` | out | 10 | Julian day UT1; 0 means unavailable contact |
| `attr` | out | 20 | indexed event/phenomena attributes; see field reference |
| `backward` | in | 1 | 0 forward, 1 backward; occultation one-try bit 32768 |
| `serr` | out | 256 | diagnostic text |

Output `tret` indices:

- `0`: maximum JD UT1
- `1`: first contact JD UT1
- `2`: second contact JD UT1
- `3`: third contact JD UT1
- `4`: fourth contact JD UT1
- `5`: sunrise JD UT1
- `6`: sunset JD UT1
- `7`: reserved
- `8`: reserved
- `9`: reserved

Output `attr` indices:

- `0`: diameter fraction covered
- `1`: Moon/object diameter ratio
- `2`: disc obscuration fraction
- `3`: core shadow diameter km
- `4`: azimuth south through west degrees
- `5`: true altitude degrees
- `6`: apparent altitude degrees
- `7`: Moon/object separation degrees
- `8`: NASA magnitude
- `9`: Saros series (-99999999 unavailable)
- `10`: Saros member (-99999999 unavailable)

Example inputs in parameter order: `[2451545.0, 3, "", 2, [13.405, 52.52, 34.0], 0]`.

### SW_SWE_LUN_OCCULT_WHERE

Native: `int32 swe_lun_occult_where(double tjd, int32 ipl, char *starname, int32 ifl, double *geopos, double *attr, char *serr);`

Family: Events. Return handling: `event`.

Data: Body/date/flags determine files: sepl_18.se1, semo_18.se1, seas_18.se1; external asteroid/JPL files where requested.

| Argument | Direction | Capacity | Units |
| --- | --- | ---: | --- |
| `tjd` | in | 1 | Julian day UT1 |
| `ipl` | in | 1 | Swiss body ID |
| `starname` | in | 256 | star name, or empty to use ipl |
| `ifl` | in | 1 | ephemeris flag bit mask |
| `geopos` | out | 10 | degrees east; degrees north; metres above sea level |
| `attr` | out | 20 | indexed event/phenomena attributes; see field reference |
| `serr` | out | 256 | diagnostic text |

Output `geopos` indices:

- `0`: central longitude degrees east
- `1`: central latitude degrees north

Output `attr` indices:

- `0`: diameter fraction covered
- `1`: Moon/object diameter ratio
- `2`: disc obscuration fraction
- `3`: core shadow diameter km
- `4`: azimuth south through west degrees
- `5`: true altitude degrees
- `6`: apparent altitude degrees
- `7`: Moon/object separation degrees
- `8`: NASA magnitude
- `9`: Saros series (-99999999 unavailable)
- `10`: Saros member (-99999999 unavailable)

Example inputs in parameter order: `[2451545.0, 3, "", 2]`.

### SW_SWE_MOONCROSS

Native: `double swe_mooncross(double x2cross, double jd_et, int32 flag, char *serr);`

Family: Events. Return handling: `crossing`.

Data: Body/date/flags determine files: sepl_18.se1, semo_18.se1, seas_18.se1; external asteroid/JPL files where requested.

| Argument | Direction | Capacity | Units |
| --- | --- | ---: | --- |
| `x2cross` | in | 1 | target ecliptic longitude, degrees |
| `jd_et` | in | 1 | Julian day TT or explicitly named input time scale |
| `flag` | in | 1 | Swiss flag bit mask |
| `serr` | out | 256 | diagnostic text |

Example inputs in parameter order: `[0.0, 2451545.0, 258]`.

### SW_SWE_MOONCROSS_NODE

Native: `double swe_mooncross_node(double jd_et, int32 flag, double *xlon, double *xlat, char *serr);`

Family: Events. Return handling: `crossing`.

Data: Body/date/flags determine files: sepl_18.se1, semo_18.se1, seas_18.se1; external asteroid/JPL files where requested.

| Argument | Direction | Capacity | Units |
| --- | --- | ---: | --- |
| `jd_et` | in | 1 | Julian day TT or explicitly named input time scale |
| `flag` | in | 1 | Swiss flag bit mask |
| `xlon` | out | 1 | degrees |
| `xlat` | out | 1 | degrees |
| `serr` | out | 256 | diagnostic text |

Example inputs in parameter order: `[2451545.0, 258]`.

### SW_SWE_MOONCROSS_NODE_UT

Native: `double swe_mooncross_node_ut(double jd_ut, int32 flag, double *xlon, double *xlat, char *serr);`

Family: Events. Return handling: `crossing`.

Data: Body/date/flags determine files: sepl_18.se1, semo_18.se1, seas_18.se1; external asteroid/JPL files where requested.

| Argument | Direction | Capacity | Units |
| --- | --- | ---: | --- |
| `jd_ut` | in | 1 | Julian day UT1 |
| `flag` | in | 1 | Swiss flag bit mask |
| `xlon` | out | 1 | degrees |
| `xlat` | out | 1 | degrees |
| `serr` | out | 256 | diagnostic text |

Example inputs in parameter order: `[2451545.0, 258]`.

### SW_SWE_MOONCROSS_UT

Native: `double swe_mooncross_ut(double x2cross, double jd_ut, int32 flag, char *serr);`

Family: Events. Return handling: `crossing`.

Data: Body/date/flags determine files: sepl_18.se1, semo_18.se1, seas_18.se1; external asteroid/JPL files where requested.

| Argument | Direction | Capacity | Units |
| --- | --- | ---: | --- |
| `x2cross` | in | 1 | target ecliptic longitude, degrees |
| `jd_ut` | in | 1 | Julian day UT1 |
| `flag` | in | 1 | Swiss flag bit mask |
| `serr` | out | 256 | diagnostic text |

Example inputs in parameter order: `[0.0, 2451545.0, 258]`.

### SW_SWE_NOD_APS

Native: `int32 swe_nod_aps(double tjd_et, int32 ipl, int32 iflag, int32 method, double *xnasc, double *xndsc, double *xperi, double *xaphe, char *serr);`

Family: Positions. Return handling: `status`.

Data: Body/date/flags determine files: sepl_18.se1, semo_18.se1, seas_18.se1; external asteroid/JPL files where requested.

| Argument | Direction | Capacity | Units |
| --- | --- | ---: | --- |
| `tjd_et` | in | 1 | Julian day TT or explicitly named input time scale |
| `ipl` | in | 1 | Swiss body ID |
| `iflag` | in | 1 | Swiss flag bit mask |
| `method` | in | 1 | node/apsis method bit mask |
| `xnasc` | out | 6 | ascending node coordinates and speeds, selected by flags |
| `xndsc` | out | 6 | descending node coordinates and speeds, selected by flags |
| `xperi` | out | 6 | periapsis coordinates and speeds, selected by flags |
| `xaphe` | out | 6 | apoapsis/focal-point coordinates and speeds, selected by flags |
| `serr` | out | 256 | diagnostic text |

Example inputs in parameter order: `[2451545.0, 0, 258, 0]`.

### SW_SWE_NOD_APS_UT

Native: `int32 swe_nod_aps_ut(double tjd_ut, int32 ipl, int32 iflag, int32 method, double *xnasc, double *xndsc, double *xperi, double *xaphe, char *serr);`

Family: Positions. Return handling: `status`.

Data: Body/date/flags determine files: sepl_18.se1, semo_18.se1, seas_18.se1; external asteroid/JPL files where requested.

| Argument | Direction | Capacity | Units |
| --- | --- | ---: | --- |
| `tjd_ut` | in | 1 | Julian day UT1 |
| `ipl` | in | 1 | Swiss body ID |
| `iflag` | in | 1 | Swiss flag bit mask |
| `method` | in | 1 | node/apsis method bit mask |
| `xnasc` | out | 6 | ascending node coordinates and speeds, selected by flags |
| `xndsc` | out | 6 | descending node coordinates and speeds, selected by flags |
| `xperi` | out | 6 | periapsis coordinates and speeds, selected by flags |
| `xaphe` | out | 6 | apoapsis/focal-point coordinates and speeds, selected by flags |
| `serr` | out | 256 | diagnostic text |

Example inputs in parameter order: `[2451545.0, 0, 258, 0]`.

### SW_SWE_ORBIT_MAX_MIN_TRUE_DISTANCE

Native: `int32 swe_orbit_max_min_true_distance(double tjd_et, int32 ipl, int32 iflag, double *dmax, double *dmin, double *dtrue, char *serr);`

Family: Positions. Return handling: `status`.

Data: Body/date/flags determine files: sepl_18.se1, semo_18.se1, seas_18.se1; external asteroid/JPL files where requested.

| Argument | Direction | Capacity | Units |
| --- | --- | ---: | --- |
| `tjd_et` | in | 1 | Julian day TT or explicitly named input time scale |
| `ipl` | in | 1 | Swiss body ID |
| `iflag` | in | 1 | Swiss flag bit mask |
| `dmax` | out | 1 | AU |
| `dmin` | out | 1 | AU |
| `dtrue` | out | 1 | AU |
| `serr` | out | 256 | diagnostic text |

Example inputs in parameter order: `[2451545.0, 0, 258]`.

### SW_SWE_PHENO

Native: `int32 swe_pheno(double tjd, int32 ipl, int32 iflag, double *attr, char *serr);`

Family: Positions. Return handling: `status`.

Data: Body/date/flags determine files: sepl_18.se1, semo_18.se1, seas_18.se1; external asteroid/JPL files where requested.

| Argument | Direction | Capacity | Units |
| --- | --- | ---: | --- |
| `tjd` | in | 1 | Julian day TT or explicitly named input time scale |
| `ipl` | in | 1 | Swiss body ID |
| `iflag` | in | 1 | Swiss flag bit mask |
| `attr` | out | 20 | indexed event/phenomena attributes; see field reference |
| `serr` | out | 256 | diagnostic text |

Output `attr` indices:

- `0`: phase angle degrees
- `1`: illuminated fraction
- `2`: elongation degrees
- `3`: apparent diameter degrees
- `4`: apparent magnitude
- `5`: horizontal parallax degrees (Moon)

Example inputs in parameter order: `[2451545.0, 0, 258]`.

### SW_SWE_PHENO_UT

Native: `int32 swe_pheno_ut(double tjd_ut, int32 ipl, int32 iflag, double *attr, char *serr);`

Family: Positions. Return handling: `status`.

Data: Body/date/flags determine files: sepl_18.se1, semo_18.se1, seas_18.se1; external asteroid/JPL files where requested.

| Argument | Direction | Capacity | Units |
| --- | --- | ---: | --- |
| `tjd_ut` | in | 1 | Julian day UT1 |
| `ipl` | in | 1 | Swiss body ID |
| `iflag` | in | 1 | Swiss flag bit mask |
| `attr` | out | 20 | indexed event/phenomena attributes; see field reference |
| `serr` | out | 256 | diagnostic text |

Output `attr` indices:

- `0`: phase angle degrees
- `1`: illuminated fraction
- `2`: elongation degrees
- `3`: apparent diameter degrees
- `4`: apparent magnitude
- `5`: horizontal parallax degrees (Moon)

Example inputs in parameter order: `[2451545.0, 0, 258]`.

### SW_SWE_RAD_MIDP

Native: `double swe_rad_midp(double x1, double x0);`

Family: Utilities. Return handling: `value`.

Data: No binary data normally required; delta-T/leap-second model applies to time conversions.

| Argument | Direction | Capacity | Units |
| --- | --- | ---: | --- |
| `x1` | in | 1 | radians |
| `x0` | in | 1 | radians |

Example inputs in parameter order: `[350.0, 10.0]`.

### SW_SWE_RADNORM

Native: `double swe_radnorm(double x);`

Family: Utilities. Return handling: `value`.

Data: No binary data normally required; delta-T/leap-second model applies to time conversions.

| Argument | Direction | Capacity | Units |
| --- | --- | ---: | --- |
| `x` | in | 1 | radians |

Example inputs in parameter order: `[12.345]`.

### SW_SWE_REFRAC

Native: `double swe_refrac(double inalt, double atpress, double attemp, int32 calc_flag);`

Family: Utilities. Return handling: `value`.

Data: No binary data normally required; delta-T/leap-second model applies to time conversions.

| Argument | Direction | Capacity | Units |
| --- | --- | ---: | --- |
| `inalt` | in | 1 | altitude degrees |
| `atpress` | in | 1 | hPa (0 estimates pressure) |
| `attemp` | in | 1 | degrees Celsius |
| `calc_flag` | in | 1 | coordinate/refraction mode code |

Example inputs in parameter order: `[5.0, 1013.25, 15.0, 0]`.

### SW_SWE_REFRAC_EXTENDED

Native: `double swe_refrac_extended(double inalt, double geoalt, double atpress, double attemp, double lapse_rate, int32 calc_flag, double *dret);`

Family: Utilities. Return handling: `value`.

Data: No binary data normally required; delta-T/leap-second model applies to time conversions.

| Argument | Direction | Capacity | Units |
| --- | --- | ---: | --- |
| `inalt` | in | 1 | altitude degrees |
| `geoalt` | in | 1 | metres |
| `atpress` | in | 1 | hPa (0 estimates pressure) |
| `attemp` | in | 1 | degrees Celsius |
| `lapse_rate` | in | 1 | kelvin/metre |
| `calc_flag` | in | 1 | coordinate/refraction mode code |
| `dret` | out | 4 | true altitude deg; apparent altitude deg; refraction deg; dip deg |

Output `dret` indices:

- `0`: true altitude degrees
- `1`: apparent altitude degrees
- `2`: refraction degrees
- `3`: horizon dip degrees

Example inputs in parameter order: `[5.0, 34.0, 1013.25, 15.0, 0.0065, 0]`.

### SW_SWE_REVJUL

Native: `void swe_revjul(double jd, int gregflag, int *jyear, int *jmon, int *jday, double *jut);`

Family: Time. Return handling: `value`.

Data: No binary data normally required; delta-T/leap-second model applies to time conversions.

| Argument | Direction | Capacity | Units |
| --- | --- | ---: | --- |
| `jd` | in | 1 | Julian day TT or explicitly named input time scale |
| `gregflag` | in | 1 | 0 Julian, 1 Gregorian |
| `jyear` | out | 1 | astronomical year |
| `jmon` | out | 1 | month 1..12 |
| `jday` | out | 1 | day of month |
| `jut` | out | 1 | decimal hours |

Example inputs in parameter order: `[2451545.0, 1]`.

### SW_SWE_RISE_TRANS

Native: `int32 swe_rise_trans(double tjd_ut, int32 ipl, char *starname, int32 epheflag, int32 rsmi, double *geopos, double atpress, double attemp, double *tret, char *serr);`

Family: Events. Return handling: `rise`.

Data: Body/date/flags determine files: sepl_18.se1, semo_18.se1, seas_18.se1; external asteroid/JPL files where requested.

| Argument | Direction | Capacity | Units |
| --- | --- | ---: | --- |
| `tjd_ut` | in | 1 | Julian day UT1 |
| `ipl` | in | 1 | Swiss body ID |
| `starname` | in | 256 | star name, or empty to use ipl |
| `epheflag` | in | 1 | ephemeris flag bit mask |
| `rsmi` | in | 1 | rise/set/transit bit mask |
| `geopos` | in | 3 | degrees east; degrees north; metres above sea level |
| `atpress` | in | 1 | hPa (0 estimates pressure) |
| `attemp` | in | 1 | degrees Celsius |
| `tret` | out | 1 | Julian day UT1; 0 means unavailable contact |
| `serr` | out | 256 | diagnostic text |

Example inputs in parameter order: `[2451545.0, 0, "", 2, 1, [13.405, 52.52, 34.0], 1013.25, 15.0]`.

### SW_SWE_RISE_TRANS_TRUE_HOR

Native: `int32 swe_rise_trans_true_hor(double tjd_ut, int32 ipl, char *starname, int32 epheflag, int32 rsmi, double *geopos, double atpress, double attemp, double horhgt, double *tret, char *serr);`

Family: Events. Return handling: `rise`.

Data: Body/date/flags determine files: sepl_18.se1, semo_18.se1, seas_18.se1; external asteroid/JPL files where requested.

| Argument | Direction | Capacity | Units |
| --- | --- | ---: | --- |
| `tjd_ut` | in | 1 | Julian day UT1 |
| `ipl` | in | 1 | Swiss body ID |
| `starname` | in | 256 | star name, or empty to use ipl |
| `epheflag` | in | 1 | ephemeris flag bit mask |
| `rsmi` | in | 1 | rise/set/transit bit mask |
| `geopos` | in | 3 | degrees east; degrees north; metres above sea level |
| `atpress` | in | 1 | hPa (0 estimates pressure) |
| `attemp` | in | 1 | degrees Celsius |
| `horhgt` | in | 1 | degrees |
| `tret` | out | 1 | Julian day UT1; 0 means unavailable contact |
| `serr` | out | 256 | diagnostic text |

Example inputs in parameter order: `[2451545.0, 0, "", 2, 1, [13.405, 52.52, 34.0], 1013.25, 15.0, 0.0]`.

### SW_CMD_SET_ASTRO_MODELS

Native: `void swe_set_astro_models(char *samod, int32 iflag);`

Family: Configuration. Return handling: `value`. **Upstream experimental.**

Data: Configuration/state query; supplied paths/files must exist for subsequent calculations.

| Argument | Direction | Capacity | Units |
| --- | --- | ---: | --- |
| `samod` | in | 256 | experimental astronomical model/version specification |
| `iflag` | in | 1 | Swiss flag bit mask |

Example inputs in parameter order: `["0,0,0,0,0,0,0,0", 258]`.

### SW_CMD_SET_DELTA_T_USERDEF

Native: `void swe_set_delta_t_userdef(double dt);`

Family: Configuration. Return handling: `value`.

Data: Configuration/state query; supplied paths/files must exist for subsequent calculations.

| Argument | Direction | Capacity | Units |
| --- | --- | ---: | --- |
| `dt` | in | 1 | days; automatic sentinel -1e-10 |

Example inputs in parameter order: `[-1e-10]`.

### SW_CMD_SET_EPHE_PATH

Native: `void swe_set_ephe_path(const char *path);`

Family: Configuration. Return handling: `value`.

Data: Configuration/state query; supplied paths/files must exist for subsequent calculations.

| Argument | Direction | Capacity | Units |
| --- | --- | ---: | --- |
| `path` | in | 256 | absolute directory or path list, Windows ANSI encoding |

Example inputs in parameter order: `["@DATA@"]`.

### SW_CMD_SET_INTERPOLATE_NUT

Native: `void swe_set_interpolate_nut(AS_BOOL do_interpolate);`

Family: Configuration. Return handling: `value`.

Data: Configuration/state query; supplied paths/files must exist for subsequent calculations.

| Argument | Direction | Capacity | Units |
| --- | --- | ---: | --- |
| `do_interpolate` | in | 1 | C boolean 0 or 1 |

Example inputs in parameter order: `[0]`.

### SW_CMD_SET_JPL_FILE

Native: `void swe_set_jpl_file(const char *fname);`

Family: Configuration. Return handling: `value`.

Data: Configuration/state query; supplied paths/files must exist for subsequent calculations.

| Argument | Direction | Capacity | Units |
| --- | --- | ---: | --- |
| `fname` | in | 256 | JPL filename relative to configured data path |

Example inputs in parameter order: `["de431.eph"]`.

### SW_CMD_SET_LAPSE_RATE

Native: `void swe_set_lapse_rate(double lapse_rate);`

Family: Configuration. Return handling: `value`.

Data: Configuration/state query; supplied paths/files must exist for subsequent calculations.

| Argument | Direction | Capacity | Units |
| --- | --- | ---: | --- |
| `lapse_rate` | in | 1 | kelvin/metre |

Example inputs in parameter order: `[0.0065]`.

### SW_CMD_SET_SID_MODE

Native: `void swe_set_sid_mode(int32 sid_mode, double t0, double ayan_t0);`

Family: Configuration. Return handling: `value`.

Data: Configuration/state query; supplied paths/files must exist for subsequent calculations.

| Argument | Direction | Capacity | Units |
| --- | --- | ---: | --- |
| `sid_mode` | in | 1 | sidereal mode ID and modifier bits; 255 user defined |
| `t0` | in | 1 | reference Julian day TT unless sidereal modifier changes it |
| `ayan_t0` | in | 1 | degrees |

Example inputs in parameter order: `[1, 2451545.0, 24.0]`.

### SW_CMD_SET_TID_ACC

Native: `void swe_set_tid_acc(double t_acc);`

Family: Configuration. Return handling: `value`.

Data: Configuration/state query; supplied paths/files must exist for subsequent calculations.

| Argument | Direction | Capacity | Units |
| --- | --- | ---: | --- |
| `t_acc` | in | 1 | arcseconds/century squared; automatic sentinel 999999 |

Example inputs in parameter order: `[999999.0]`.

### SW_CMD_SET_TOPO

Native: `void swe_set_topo(double geolon, double geolat, double geoalt);`

Family: Configuration. Return handling: `value`.

Data: Configuration/state query; supplied paths/files must exist for subsequent calculations.

| Argument | Direction | Capacity | Units |
| --- | --- | ---: | --- |
| `geolon` | in | 1 | degrees east |
| `geolat` | in | 1 | degrees north |
| `geoalt` | in | 1 | metres |

Example inputs in parameter order: `[13.405, 52.52, 34.0]`.

### SW_SWE_SIDTIME

Native: `double swe_sidtime(double tjd_ut);`

Family: Time. Return handling: `value`.

Data: No binary data normally required; delta-T/leap-second model applies to time conversions.

| Argument | Direction | Capacity | Units |
| --- | --- | ---: | --- |
| `tjd_ut` | in | 1 | Julian day UT1 |

Example inputs in parameter order: `[2451545.0]`.

### SW_SWE_SIDTIME0

Native: `double swe_sidtime0(double tjd_ut, double eps, double nut);`

Family: Time. Return handling: `value`.

Data: No binary data normally required; delta-T/leap-second model applies to time conversions.

| Argument | Direction | Capacity | Units |
| --- | --- | ---: | --- |
| `tjd_ut` | in | 1 | Julian day UT1 |
| `eps` | in | 1 | degrees |
| `nut` | in | 1 | degrees |

Example inputs in parameter order: `[2451545.0, 23.4392911, 0.0]`.

### SW_SWE_SOL_ECLIPSE_HOW

Native: `int32 swe_sol_eclipse_how(double tjd, int32 ifl, double *geopos, double *attr, char *serr);`

Family: Events. Return handling: `event`.

Data: Body/date/flags determine files: sepl_18.se1, semo_18.se1, seas_18.se1; external asteroid/JPL files where requested.

| Argument | Direction | Capacity | Units |
| --- | --- | ---: | --- |
| `tjd` | in | 1 | Julian day UT1 |
| `ifl` | in | 1 | ephemeris flag bit mask |
| `geopos` | in | 3 | degrees east; degrees north; metres above sea level |
| `attr` | out | 20 | indexed event/phenomena attributes; see field reference |
| `serr` | out | 256 | diagnostic text |

Output `attr` indices:

- `0`: diameter fraction covered
- `1`: Moon/object diameter ratio
- `2`: disc obscuration fraction
- `3`: core shadow diameter km
- `4`: azimuth south through west degrees
- `5`: true altitude degrees
- `6`: apparent altitude degrees
- `7`: Moon/object separation degrees
- `8`: NASA magnitude
- `9`: Saros series (-99999999 unavailable)
- `10`: Saros member (-99999999 unavailable)

Example inputs in parameter order: `[2451545.0, 2, [13.405, 52.52, 34.0]]`.

### SW_SWE_SOL_ECLIPSE_WHEN_GLOB

Native: `int32 swe_sol_eclipse_when_glob(double tjd_start, int32 ifl, int32 ifltype, double *tret, int32 backward, char *serr);`

Family: Events. Return handling: `event`.

Data: Body/date/flags determine files: sepl_18.se1, semo_18.se1, seas_18.se1; external asteroid/JPL files where requested.

| Argument | Direction | Capacity | Units |
| --- | --- | ---: | --- |
| `tjd_start` | in | 1 | Julian day UT1 |
| `ifl` | in | 1 | ephemeris flag bit mask |
| `ifltype` | in | 1 | event type bit mask |
| `tret` | out | 10 | Julian day UT1; 0 means unavailable contact |
| `backward` | in | 1 | 0 forward, 1 backward; occultation one-try bit 32768 |
| `serr` | out | 256 | diagnostic text |

Output `tret` indices:

- `0`: maximum JD UT1
- `1`: local apparent noon JD UT1
- `2`: eclipse begins JD UT1
- `3`: eclipse ends JD UT1
- `4`: totality begins JD UT1
- `5`: totality ends JD UT1
- `6`: centre line begins JD UT1
- `7`: centre line ends JD UT1
- `8`: reserved unimplemented
- `9`: reserved unimplemented

Example inputs in parameter order: `[2451545.0, 2, 0, 0]`.

### SW_SWE_SOL_ECLIPSE_WHEN_LOC

Native: `int32 swe_sol_eclipse_when_loc(double tjd_start, int32 ifl, double *geopos, double *tret, double *attr, int32 backward, char *serr);`

Family: Events. Return handling: `event`.

Data: Body/date/flags determine files: sepl_18.se1, semo_18.se1, seas_18.se1; external asteroid/JPL files where requested.

| Argument | Direction | Capacity | Units |
| --- | --- | ---: | --- |
| `tjd_start` | in | 1 | Julian day UT1 |
| `ifl` | in | 1 | ephemeris flag bit mask |
| `geopos` | in | 3 | degrees east; degrees north; metres above sea level |
| `tret` | out | 10 | Julian day UT1; 0 means unavailable contact |
| `attr` | out | 20 | indexed event/phenomena attributes; see field reference |
| `backward` | in | 1 | 0 forward, 1 backward; occultation one-try bit 32768 |
| `serr` | out | 256 | diagnostic text |

Output `tret` indices:

- `0`: maximum JD UT1
- `1`: first contact JD UT1
- `2`: second contact JD UT1
- `3`: third contact JD UT1
- `4`: fourth contact JD UT1
- `5`: sunrise JD UT1
- `6`: sunset JD UT1
- `7`: reserved
- `8`: reserved
- `9`: reserved

Output `attr` indices:

- `0`: diameter fraction covered
- `1`: Moon/object diameter ratio
- `2`: disc obscuration fraction
- `3`: core shadow diameter km
- `4`: azimuth south through west degrees
- `5`: true altitude degrees
- `6`: apparent altitude degrees
- `7`: Moon/object separation degrees
- `8`: NASA magnitude
- `9`: Saros series (-99999999 unavailable)
- `10`: Saros member (-99999999 unavailable)

Example inputs in parameter order: `[2451545.0, 2, [13.405, 52.52, 34.0], 0]`.

### SW_SWE_SOL_ECLIPSE_WHERE

Native: `int32 swe_sol_eclipse_where(double tjd, int32 ifl, double *geopos, double *attr, char *serr);`

Family: Events. Return handling: `event`.

Data: Body/date/flags determine files: sepl_18.se1, semo_18.se1, seas_18.se1; external asteroid/JPL files where requested.

| Argument | Direction | Capacity | Units |
| --- | --- | ---: | --- |
| `tjd` | in | 1 | Julian day UT1 |
| `ifl` | in | 1 | ephemeris flag bit mask |
| `geopos` | out | 10 | degrees east; degrees north; metres above sea level |
| `attr` | out | 20 | indexed event/phenomena attributes; see field reference |
| `serr` | out | 256 | diagnostic text |

Output `geopos` indices:

- `0`: central longitude degrees east
- `1`: central latitude degrees north

Output `attr` indices:

- `0`: diameter fraction covered
- `1`: Moon/object diameter ratio
- `2`: disc obscuration fraction
- `3`: core shadow diameter km
- `4`: azimuth south through west degrees
- `5`: true altitude degrees
- `6`: apparent altitude degrees
- `7`: Moon/object separation degrees
- `8`: NASA magnitude
- `9`: Saros series (-99999999 unavailable)
- `10`: Saros member (-99999999 unavailable)

Example inputs in parameter order: `[2451545.0, 2]`.

### SW_SWE_SOLCROSS

Native: `double swe_solcross(double x2cross, double jd_et, int32 flag, char *serr);`

Family: Events. Return handling: `crossing`.

Data: Body/date/flags determine files: sepl_18.se1, semo_18.se1, seas_18.se1; external asteroid/JPL files where requested.

| Argument | Direction | Capacity | Units |
| --- | --- | ---: | --- |
| `x2cross` | in | 1 | target ecliptic longitude, degrees |
| `jd_et` | in | 1 | Julian day TT or explicitly named input time scale |
| `flag` | in | 1 | Swiss flag bit mask |
| `serr` | out | 256 | diagnostic text |

Example inputs in parameter order: `[0.0, 2451545.0, 258]`.

### SW_SWE_SOLCROSS_UT

Native: `double swe_solcross_ut(double x2cross, double jd_ut, int32 flag, char *serr);`

Family: Events. Return handling: `crossing`.

Data: Body/date/flags determine files: sepl_18.se1, semo_18.se1, seas_18.se1; external asteroid/JPL files where requested.

| Argument | Direction | Capacity | Units |
| --- | --- | ---: | --- |
| `x2cross` | in | 1 | target ecliptic longitude, degrees |
| `jd_ut` | in | 1 | Julian day UT1 |
| `flag` | in | 1 | Swiss flag bit mask |
| `serr` | out | 256 | diagnostic text |

Example inputs in parameter order: `[0.0, 2451545.0, 258]`.

### SW_SWE_SPLIT_DEG

Native: `void swe_split_deg(double ddeg, int32 roundflag, int32 *ideg, int32 *imin, int32 *isec, double *dsecfr, int32 *isgn);`

Family: Utilities. Return handling: `value`.

Data: No binary data normally required; delta-T/leap-second model applies to time conversions.

| Argument | Direction | Capacity | Units |
| --- | --- | ---: | --- |
| `ddeg` | in | 1 | degrees |
| `roundflag` | in | 1 | split-degree bit mask |
| `ideg` | out | 1 | whole degrees |
| `imin` | out | 1 | minutes |
| `isec` | out | 1 | arcseconds |
| `dsecfr` | out | 1 | fractional arcsecond |
| `isgn` | out | 1 | sign or zodiac index selected by roundflag |

Example inputs in parameter order: `[123.456, 0]`.

### SW_SWE_TIME_EQU

Native: `int32 swe_time_equ(double tjd, double *te, char *serr);`

Family: Time. Return handling: `status`.

Data: No binary data normally required; delta-T/leap-second model applies to time conversions.

| Argument | Direction | Capacity | Units |
| --- | --- | ---: | --- |
| `tjd` | in | 1 | Julian day UT1 |
| `te` | out | 1 | days (apparent minus mean solar time) |
| `serr` | out | 256 | diagnostic text |

Example inputs in parameter order: `[2451545.0]`.

### SW_SWE_TOPO_ARCUS_VISIONIS

Native: `int32 swe_topo_arcus_visionis(double tjdut, double *dgeo, double *datm, double *dobs, int32 helflag, double mag, double azi_obj, double alt_obj, double azi_sun, double azi_moon, double alt_moon, double *dret, char *serr);`

Family: Heliacal. Return handling: `status`. **Upstream experimental.**

Data: Body/date/flags determine files: sepl_18.se1, semo_18.se1, seas_18.se1; external asteroid/JPL files where requested.

| Argument | Direction | Capacity | Units |
| --- | --- | ---: | --- |
| `tjdut` | in | 1 | Julian day UT1 |
| `dgeo` | in | 3 | degrees east; degrees north; metres above sea level |
| `datm` | inout | 4 | hPa; Celsius; relative humidity percent; visibility km or extinction coefficient |
| `dobs` | inout | 6 | age years; Snellen ratio; binocular flag; magnification; aperture mm; transmission fraction |
| `helflag` | in | 1 | heliacal flag bit mask |
| `mag` | in | 1 | astronomical magnitude |
| `azi_obj` | in | 1 | azimuth degrees |
| `alt_obj` | in | 1 | altitude degrees |
| `azi_sun` | in | 1 | azimuth degrees |
| `azi_moon` | in | 1 | azimuth degrees |
| `alt_moon` | in | 1 | altitude degrees |
| `dret` | out | 1 | arcus visionis degrees |
| `serr` | out | 256 | diagnostic text |

Example inputs in parameter order: `[2451545.0, [13.405, 52.52, 34.0], [1013.25, 15.0, 40.0, 0.0], [36.0, 1.0, 0.0, 1.0, 0.0, 0.0], 0, 0.0, 90.0, 5.0, 100.0, 110.0, -10.0]`.

### SW_SWE_UTC_TIME_ZONE

Native: `void swe_utc_time_zone(int32 iyear, int32 imonth, int32 iday, int32 ihour, int32 imin, double dsec, double d_timezone, int32 *iyear_out, int32 *imonth_out, int32 *iday_out, int32 *ihour_out, int32 *imin_out, double *dsec_out);`

Family: Time. Return handling: `value`.

Data: No binary data normally required; delta-T/leap-second model applies to time conversions.

| Argument | Direction | Capacity | Units |
| --- | --- | ---: | --- |
| `iyear` | in | 1 | astronomical year |
| `imonth` | in | 1 | month 1..12 |
| `iday` | in | 1 | day of month |
| `ihour` | in | 1 | hour 0..23 |
| `imin` | in | 1 | minutes |
| `dsec` | in | 1 | seconds, leap second allowed where documented |
| `d_timezone` | in | 1 | hours east of UTC (fractional allowed) |
| `iyear_out` | out | 1 | astronomical year |
| `imonth_out` | out | 1 | month 1..12 |
| `iday_out` | out | 1 | day of month |
| `ihour_out` | out | 1 | hour 0..23 |
| `imin_out` | out | 1 | minutes |
| `dsec_out` | out | 1 | seconds |

Example inputs in parameter order: `[2000, 1, 1, 12, 0, 0.0, 5.5]`.

### SW_SWE_UTC_TO_JD

Native: `int32 swe_utc_to_jd(int32 iyear, int32 imonth, int32 iday, int32 ihour, int32 imin, double dsec, int32 gregflag, double *dret, char *serr);`

Family: Time. Return handling: `status`.

Data: No binary data normally required; delta-T/leap-second model applies to time conversions.

| Argument | Direction | Capacity | Units |
| --- | --- | ---: | --- |
| `iyear` | in | 1 | astronomical year |
| `imonth` | in | 1 | month 1..12 |
| `iday` | in | 1 | day of month |
| `ihour` | in | 1 | hour 0..23 |
| `imin` | in | 1 | minutes |
| `dsec` | in | 1 | seconds, leap second allowed where documented |
| `gregflag` | in | 1 | 0 Julian, 1 Gregorian |
| `dret` | out | 2 | Julian day TT, Julian day UT1 |
| `serr` | out | 256 | diagnostic text |

Output `dret` indices:

- `0`: Julian day TT
- `1`: Julian day UT1

Example inputs in parameter order: `[2000, 1, 1, 12, 0, 0.0, 1]`.

### SW_SWE_VERSION

Native: `char * swe_version(char *);`

Family: Utilities. Return handling: `value`.

Data: No binary data normally required; delta-T/leap-second model applies to time conversions.

| Argument | Direction | Capacity | Units |
| --- | --- | ---: | --- |
| `arg1` | out | 256 | NUL-terminated string |

Example inputs in parameter order: `[]`.

### SW_SWE_VIS_LIMIT_MAG

Native: `int32 swe_vis_limit_mag(double tjdut, double *geopos, double *datm, double *dobs, char *ObjectName, int32 helflag, double *dret, char *serr);`

Family: Heliacal. Return handling: `visibility`.

Data: Body/date/flags determine files: sepl_18.se1, semo_18.se1, seas_18.se1; external asteroid/JPL files where requested.

| Argument | Direction | Capacity | Units |
| --- | --- | ---: | --- |
| `tjdut` | in | 1 | Julian day UT1 |
| `geopos` | in | 3 | degrees east; degrees north; metres above sea level |
| `datm` | inout | 4 | hPa; Celsius; relative humidity percent; visibility km or extinction coefficient |
| `dobs` | inout | 6 | age years; Snellen ratio; binocular flag; magnification; aperture mm; transmission fraction |
| `ObjectName` | in | 256 | planet or fixed-star name |
| `helflag` | in | 1 | heliacal flag bit mask |
| `dret` | out | 10 | limiting magnitude; object alt/az deg; Sun alt/az deg; Moon alt/az deg; object magnitude |
| `serr` | out | 256 | diagnostic text |

Output `dret` indices:

- `0`: limiting magnitude
- `1`: object altitude degrees
- `2`: object azimuth degrees
- `3`: Sun altitude degrees
- `4`: Sun azimuth degrees
- `5`: Moon altitude degrees
- `6`: Moon azimuth degrees
- `7`: object magnitude

Example inputs in parameter order: `[2451545.0, [13.405, 52.52, 34.0], [1013.25, 15.0, 40.0, 0.0], [36.0, 1.0, 0.0, 1.0, 0.0, 0.0], "venus", 0]`.
