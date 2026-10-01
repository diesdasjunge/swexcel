# Pinned Swiss Ephemeris API source

These files are exact upstream bytes from release **v2.10.3bfinal**, commit
`f4dcd18e8005dde95fd8a8d2312ed12f9accd1b0`:

- [`swephexp.h`](https://github.com/aloistr/swisseph/blob/f4dcd18e8005dde95fd8a8d2312ed12f9accd1b0/swephexp.h): official exported C prototypes and constants.
- [`sweodef.h`](https://github.com/aloistr/swisseph/blob/f4dcd18e8005dde95fd8a8d2312ed12f9accd1b0/sweodef.h): integer, boolean and centisecond type definitions.
- [`LICENSE`](https://github.com/aloistr/swisseph/blob/f4dcd18e8005dde95fd8a8d2312ed12f9accd1b0/LICENSE): upstream license notices.
- `LICENSE.TXT`: preserved upstream link-file content (`LICENSE`); read `LICENSE` for the notices.

`../../tools/generate_bindings.py` verifies their SHA-256 hashes and independently
reads the packaged AMD64 DLL's PE export table before generating
`../../src/vba/SWNative.bas` and `../../api/catalog.json`. Use `--check` to verify
the generated files without changing them. The generator has no network access.

The release tag and the DLL's reported runtime version are separate facts.
Raw ABI declaration coverage is not runtime verification or a complete worksheet
API. Pending buffer sizes, directions, ownership and units are explicit in the
catalogue. The four functions labelled "secret" in the upstream header are
identified as experimental/internal APIs, although the DLL exports them.

`source/` contains the unchanged minimal engine sources, headers, notices and
upstream Makefile needed to build the DLL and a reference `swetest64.exe` from
this exact commit. Its `provenance.json` records every download URL, size and
SHA-256. The nine engine translation units follow upstream `SWEOBJ`.

`../../tools/build-engine.ps1` is a prepared Windows MSVC build recipe, requiring
an x64 Native Tools environment and Windows SDK. It uses the upstream DLL project's
Release/x64 flags (Cdecl, `MAKE_DLL`, MultiByte, static `/MT` CRT). The build adds
local loader glue to initialize upstream `dllhandle` for `swe_get_library_path`;
the vendored upstream sources remain unchanged. It verifies AMD64 outputs and
the complete export list, and writes build provenance next to its output DLL.
This script has not yet been executed on Windows.

The latest release's published prebuilt DLL is byte-identical to the legacy
repository DLL. It is therefore only the static export-inspection reference;
its presence does not establish that latest source changes were compiled. A
fresh source build is required for the final engine-update acceptance gate.
Pinned `source/sweph.h` still defines `SE_VERSION` as `2.10.03`, despite the later
release tag. Check source/build provenance as well as the runtime version.

Retain the upstream notices when copying this folder. This folder does not change
the project's chosen license or claim Windows Excel execution has occurred.
