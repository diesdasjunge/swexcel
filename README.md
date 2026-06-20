![swexcel_logo](https://github.com/diesdasjunge/swexcel/assets/72585192/68ab345a-9a9a-4cad-9227-e073ecfc76d5)

---

# Swiss Ephemeris for 64-bit Excel & VBA

A VBA implementation of the [Swiss Ephemeris](https://www.astro.com/swisseph/swephinfo_e.htm) by Astrodienst AG, packaged as a ready-to-use 64-bit Excel spreadsheet. Calculate natal charts, Cosmodynes (planet/sign/house power scores), and precise transit dates — all from inside Excel.

> [!WARNING]
> **Windows only.** The calculation engine is a native Windows DLL (`swedll64.dll`, a 64-bit PE binary). Because of how DLLs work, it can **only** be loaded by **64-bit Microsoft Excel on Windows**. It will *not* run on Excel for macOS (Apple Silicon or Intel) or on 32-bit Excel. You can clone and edit this repo on any OS — including a Mac — but the spreadsheet only *executes* on Windows Excel, natively or inside a Windows VM (Parallels, VMware, UTM, etc.).

---

## Features

- **Natal chart calculation** — geocentric planetary positions (longitude and declination) for all classical planets plus Chiron, Moon's Node, Part of Fortune, Ascendant, and MC
- **House cusps** — Placidus house system via the Swiss Ephemeris `swe_houses` function
- **Cosmodynes** — fully automated Power, Harmony, and Discord scores for every planet, sign, and house following the Dobyns/Pottenger method
- **Transit date & time finder** — pinpoints the exact moment a transiting planet reaches a natal degree using a Secant-method numerical solver
- **Planetary dignities & receptions** — rulership, exaltation, detriment, fall, and mutual reception detection
- **Intercepted signs & houses** — automatic detection and flagging
- **Aspect power calculation** — weighted aspect scores per planet pair
- **64-bit clean** — all DLL declarations use `PtrSafe` and are compatible with 64-bit Office (Microsoft 365, Office 2016+)

---

## Requirements

| Component | Details |
|---|---|
| **Excel** | 64-bit, Microsoft 365 / Office 2016 or later (Windows only) |
| **Swiss Ephemeris DLL** | `swedll64.dll` (Swiss Ephemeris 2.10.03, Astrodienst prebuilt) — included in `ephem/` |
| **Ephemeris data files** | `SEPL_18.SE1`, `SEMO_18.SE1`, `SEAS_18.SE1` (JPL DE441) — included in `ephem/` |
| **Macros** | Must be enabled in Excel Trust Center |

> **Note:** This project is Windows-only. The 64-bit DLL cannot be loaded by 32-bit Excel or Excel on macOS.

---

## Repository Layout

```
swexcel/
├── Cosmodynes and Transits SE spreadsheet_64bit.xls   # Main workbook
├── ephem/
│   ├── swedll64.dll          # 64-bit Swiss Ephemeris DLL
│   ├── swedll64.lib          # Import library
│   ├── sweph_vb7_64.bas      # Original VB7 declarations reference
│   ├── SEPL_18.SE1           # Planet ephemeris data
│   ├── SEMO_18.SE1           # Moon ephemeris data
│   ├── SEAS_18.SE1           # Asteroid ephemeris data
│   └── asteroids/
│       └── SEAS_18.SE1       # Additional asteroid data
└── vba/                      # Extracted VBA module source files
    ├── GeneralDeclarations.vba
    ├── PrintCosmodynes.vba
    ├── Get_transit_date_and_time.vba
    ├── CalcPlanetAspectPower.vba
    ├── CalcPlanetHousePower.vba
    ├── CalcHousePower.vba
    ├── CalcSignPower.vba
    ├── Dignities.vba
    ├── Reception.vba
    ├── Get_natal_data.vba
    ├── Get_1_Planet.vba
    ├── Get_1_Planet_geo.vba
    ├── Get_1_Planet_decl.vba
    ├── Get_Ascendant.vba
    ├── Get_MC.vba
    ├── PLANETSINHOUSES.vba
    ├── FigureHouseInterceptions.vba
    ├── FigureSignInterceptions
    ├── SecantMethod.vba
    └── ... (utility modules)
```

---

## Installation

### 1. Place the DLL where Excel can find it

Copy `swedll64.dll` to a directory on your system `PATH`, such as:

```
C:\sweph\bin\swedll64.dll
```

The DLL path is hardcoded in `GeneralDeclarations.vba`. If you place the DLL elsewhere, open the VBA editor (`Alt + F11`) and update every `Lib "C:\sweph\bin\swedll64.dll"` declaration to match your chosen path.

### 2. Copy the ephemeris data files

Place the `.SE1` files from `ephem/` in the same directory as the DLL, or in a location that Swiss Ephemeris can resolve at runtime:

```
C:\sweph\bin\SEPL_18.SE1
C:\sweph\bin\SEMO_18.SE1
C:\sweph\bin\SEAS_18.SE1
```

### 3. Open the workbook

Open `Cosmodynes and Transits SE spreadsheet_64bit.xls` in **64-bit Excel**. When prompted, choose **Enable Macros**.

### 4. Enter birth data

Fill in the input cells on **Sheet1**:

| Cell | Value |
|------|-------|
| A2 | Month |
| B2 | Day |
| C2 | Year |
| D2 | Hour (local time) |
| E2 | Minute |
| F2 | Second |
| G2 | UTC offset (hours) |
| H2 | Longitude (decimal degrees, West negative) |
| I2 | Latitude (decimal degrees, South negative) |

### 5. Run the calculations

Click the **Cosmodynes** button (or run `PrintCosmodynes` from the Macros dialog) to populate the output tables.

---

## Upgrading the Swiss Ephemeris engine

This repo ships Swiss Ephemeris **2.10.03** (Astrodienst's prebuilt 64-bit DLL) with **JPL DE441** data files. Refreshing or re-applying the engine is a **binary + data swap — no VBA changes are required**, because the `Declare` signatures in `GeneralDeclarations.vba` (and the `ephem/sweph_vb7_64.bas` reference copy) already match the 2.10 API.

### Files to download

Everything comes from the official Astrodienst repository: <https://github.com/aloistr/swisseph>

| File | Where in the repo | Notes |
|---|---|---|
| `swedll64.dll` | `windows/sweph.zip` → unzip → `sweph/bin/swedll64.dll` | 999,936 bytes, CRC-32 `5bf1f794` — the engine |
| `swedll64.lib` | same zip → `sweph/bin/swedll64.lib` | 23,230 bytes, CRC-32 `142ad89d` — C-link import lib only (not needed at VBA runtime) |
| `sepl_18.se1` | [`ephe/sepl_18.se1`](https://github.com/aloistr/swisseph/blob/master/ephe/sepl_18.se1) | planets, JPL DE441 (rebuilt 14 Apr 2026), ~0.5 MB |
| `semo_18.se1` | [`ephe/semo_18.se1`](https://github.com/aloistr/swisseph/blob/master/ephe/semo_18.se1) | Moon, DE441, ~1.3 MB |
| `seas_18.se1` | [`ephe/seas_18.se1`](https://github.com/aloistr/swisseph/blob/master/ephe/seas_18.se1) | main asteroids incl. Chiron, DE441, ~218 KB |

For each data file, use GitHub's **Download raw** button, or fetch `https://github.com/aloistr/swisseph/raw/master/ephe/<name>`.

### Where to put them

1. **Runtime location — what Excel actually loads.** Per this project's install that is `C:\sweph\bin\`. Replace:
   ```
   C:\sweph\bin\swedll64.dll
   C:\sweph\bin\sepl_18.se1
   C:\sweph\bin\semo_18.se1
   C:\sweph\bin\seas_18.se1
   ```
2. **Repo copy — source of truth.** Replace the matching files in `ephem/` (and the duplicate `ephem/asteroids/SEAS_18.SE1`). Windows filenames are case-insensitive, so the existing `SEPL_18.SE1` etc. and the lowercase downloads are interchangeable.

> [!NOTE]
> The shipped prebuilt (`swedll64.dll`, CRC-32 `5bf1f794`) reports `2.10.03` via `swe_version()` — verified from the binary. To build your own from source instead, use the Visual Studio project `sweph/src/projects/swedll64.vcxproj` inside `sweph.zip`.

### Verify it worked

1. **Version** — in the VBA Immediate window (`Alt+F11`, then `Ctrl+G`):
   ```vba
   ?dll_version()
   ```
   It should return `2.10.03` (no longer `1.60`).
2. **Positions** — for **2000-01-01 12:00 UT** (Julian Day `2451545.0`), geocentric tropical longitudes. Your Swiss/DE441 result should match these to within ~0.01°:

   | Body | Longitude | Body | Longitude |
   |---|---|---|---|
   | Sun | 280.369° (10°22′ ♑) | Jupiter | 25.253° (25°15′ ♈) |
   | Moon | 223.324° (13°19′ ♏) | Saturn | 40.396° (10°24′ ♉) |
   | Mercury | 271.889° (1°53′ ♑) | Uranus | 314.809° (14°49′ ♒) |
   | Venus | 241.566° (1°34′ ♐) | Neptune | 303.193° (3°12′ ♒) |
   | Mars | 327.963° (27°58′ ♓) | Pluto | 251.455° (11°27′ ♐) |
   | Mean Node | 125.041° (5°02′ ♌) | True Node | 123.953° (3°57′ ♌) |

   Placidus houses at 40°N, 75°W for the same instant: **ASC 273.84°**, **MC 207.42°**.
   *(References computed with Swiss Ephemeris 2.10.03 via pyswisseph — Moshier model, which agrees with the DE441 `.se1` files to <0.01°.)*

---

## Usage

### Cosmodynes report

Running `PrintCosmodynes` calculates and prints:

- Geocentric positions and declinations for Sun, Moon, Mercury, Venus, Mars, Jupiter, Saturn, Uranus, Neptune, Pluto, Chiron, and the Moon's Node
- Placidus house cusps and the degree span of each house
- Planet-in-house and planet-in-sign power scores
- Aspect power scores for every planet pair
- Aggregated Power, Harmony, and Discord totals per planet, sign, and house

### Transit finder

`Get_transit_date_and_time(sm, sd, sy, em, ed, ey, p_num, start_pos, angle)` returns a date/time string for when planet `p_num` completes the given `angle` from `start_pos` within the search window `[sm/sd/sy … em/ed/ey]`.

```vba
' Find when Mars reaches 15° Aries between 1 Jan 2025 and 31 Dec 2025
Dim result As String
result = Get_transit_date_and_time(1, 1, 2025, 12, 31, 2025, SE_MARS, 0, 15)
MsgBox result
```

---

## Technology Stack

| Layer | Technology |
|---|---|
| Host application | Microsoft Excel (64-bit, Windows) |
| Calculation engine | [Swiss Ephemeris](https://www.astro.com/swisseph/swephinfo_e.htm) **2.10.03** (`swedll64.dll`, Astrodienst prebuilt) — JPL DE441 ephemeris |
| Programming language | VBA (Visual Basic for Applications) with `PtrSafe` 64-bit declarations |
| Ephemeris data format | Swiss Ephemeris binary `.SE1` files, JPL DE441 rebuild (1800–2400 CE coverage) |
| Numerical solver | Secant method (`SecantMethod.vba`) for transit time refinement |

---

## Swiss Ephemeris Planet Constants

The following constants from `GeneralDeclarations.vba` are available in all VBA modules:

| Constant | Value | Body |
|---|---|---|
| `SE_SUN` | 0 | Sun |
| `SE_MOON` | 1 | Moon |
| `SE_MERCURY` | 2 | Mercury |
| `SE_VENUS` | 3 | Venus |
| `SE_MARS` | 4 | Mars |
| `SE_JUPITER` | 5 | Jupiter |
| `SE_SATURN` | 6 | Saturn |
| `SE_URANUS` | 7 | Uranus |
| `SE_NEPTUNE` | 8 | Neptune |
| `SE_PLUTO` | 9 | Pluto |
| `SE_CHIRON` | 15 | Chiron |

---

## Contributing

Contributions are welcome. Please:

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/my-improvement`)
3. Export any changed VBA modules as `.vba` text files into the `vba/` directory so diffs are readable
4. Open a pull request describing what you changed and why

Bug reports and feature requests can be filed via GitHub Issues.

---

## License

This project is released under the **GNU General Public License v2.0 or later**, consistent with the Swiss Ephemeris library it wraps.

The Swiss Ephemeris itself is copyright © Astrodienst AG and is used here under its GPL license. See [https://www.astro.com/swisseph/swephinfo_e.htm](https://www.astro.com/swisseph/swephinfo_e.htm) for the upstream license terms.
