![swexcel_logo](https://github.com/diesdasjunge/swexcel/assets/72585192/68ab345a-9a9a-4cad-9227-e073ecfc76d5)

---

# Swiss Ephemeris for 64-bit Excel & VBA

A VBA implementation of the [Swiss Ephemeris](https://www.astro.com/swisseph/swephinfo_e.htm) by Astrodienst AG, packaged as a ready-to-use 64-bit Excel spreadsheet. Calculate natal charts, Cosmodynes (planet/sign/house power scores), and precise transit dates — all from inside Excel.

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
| **Swiss Ephemeris DLL** | `swedll64.dll` — included in `ephem/` |
| **Ephemeris data files** | `SEPL_18.SE1`, `SEMO_18.SE1`, `SEAS_18.SE1` — included in `ephem/` |
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
| Calculation engine | [Swiss Ephemeris](https://www.astro.com/swisseph/swephinfo_e.htm) 1.60+ (`swedll64.dll`) |
| Programming language | VBA (Visual Basic for Applications) with `PtrSafe` 64-bit declarations |
| Ephemeris data format | Swiss Ephemeris binary `.SE1` files (1800–2400 CE coverage) |
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
