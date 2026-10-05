#!/usr/bin/env python3
"""Generate/check the raw VBA ABI layer from pinned headers and PE exports.

No network access or Windows execution is performed. --check is read-only.
This checks declaration coverage, not wrapper completeness or runtime safety.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import re
import struct
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
COMMIT = "f4dcd18e8005dde95fd8a8d2312ed12f9accd1b0"
RELEASE_TAG = "v2.10.3bfinal"
DLL_NAME = "swexcel-se-2.10.3b-x64.dll"
UPSTREAM = "https://github.com/aloistr/swisseph"
DOCUMENTATION = "https://www.astro.com/swisseph/swephprg.htm"
SOURCE_HASHES = {
    "swephexp.h": "53647b650504bbdec65166af1fbdb8580fa26a2ec821d5f50f032f5b0745b3cc",
    "sweodef.h": "aa763c98cfa62d9a643432b6908c820ea0759fcbdee378cf4203b0de9c6c34e0",
    "LICENSE": "b6345743292c516d07799dfad46f1a641e5270da2160cc4d6bd54594d6faa991",
    "LICENSE.TXT": "c693279643b8cd5d248172d9c22cb7cf4ed163a3c98c8a3f69c2717edd3eacb7",
}
PINNED_DLL_SHA256 = "1b1645c164cdbc073408df8cd50586cea67b9c8cbf2ec8b8cee78d4706774a35"
EXPERIMENTAL = {
    "swe_heliacal_angle": "Upstream labels this API 'secret, for Victor Reijs'.",
    "swe_topo_arcus_visionis": "Upstream labels this API 'secret, for Victor Reijs'.",
    "swe_set_astro_models": "Upstream labels this API 'secret, for Dieter' for model testing.",
    "swe_get_astro_models": "Upstream labels this API 'secret, for Dieter' for model testing.",
}
SCALAR_TYPES = {
    "int": "Long", "int32": "Long", "AS_BOOL": "Long",
    "centisec": "Long", "CSEC": "Long", "double": "Double", "char": "Byte",
}


def sha256(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def pe_exports(data: bytes) -> list[dict]:
    """Read the x64 PE export table without trusting declarations/catalogue."""
    if data[:2] != b"MZ":
        raise ValueError("Expected a PE DLL with an MZ header")
    pe = struct.unpack_from("<I", data, 0x3C)[0]
    if data[pe:pe + 4] != b"PE\0\0":
        raise ValueError("Invalid PE signature")
    machine, section_count = struct.unpack_from("<HH", data, pe + 4)
    optional_size = struct.unpack_from("<H", data, pe + 20)[0]
    optional = pe + 24
    if machine != 0x8664 or struct.unpack_from("<H", data, optional)[0] != 0x20B:
        raise ValueError("The bindings require an AMD64 PE32+ DLL")
    export_rva, export_size = struct.unpack_from("<II", data, optional + 112)
    sections = []
    for index in range(section_count):
        position = optional + optional_size + index * 40
        virtual_size, address, raw_size, raw = struct.unpack_from("<IIII", data, position + 8)
        sections.append((address, max(virtual_size, raw_size), raw, raw_size))

    def offset(rva: int) -> int:
        for address, size, raw, raw_size in sections:
            if address <= rva < address + size and rva - address < raw_size:
                return raw + rva - address
        raise ValueError(f"Export RVA {rva:#x} is outside file-backed sections")

    def text(rva: int) -> str:
        position = offset(rva)
        return data[position:data.index(0, position)].decode("ascii")

    values = struct.unpack_from("<IIHHIIIIIII", data, offset(export_rva))
    _, _, _, _, _, base, function_count, name_count, functions, names, ordinals = values
    if function_count != name_count:
        raise ValueError("Unnamed exports need explicit ABI review")
    found = []
    for index in range(name_count):
        name_rva = struct.unpack_from("<I", data, offset(names) + index * 4)[0]
        ordinal_index = struct.unpack_from("<H", data, offset(ordinals) + index * 2)[0]
        address = struct.unpack_from("<I", data, offset(functions) + ordinal_index * 4)[0]
        if export_rva <= address < export_rva + export_size:
            raise ValueError("Forwarded exports need explicit ABI review")
        found.append({"name": text(name_rva), "ordinal": base + ordinal_index, "rva": address})
    if len({item["name"] for item in found}) != len(found):
        raise ValueError("Duplicate PE export names")
    return sorted(found, key=lambda item: item["name"])


def strip_comments(source: str) -> str:
    # Retain newlines so catalogue line references address the untouched header.
    return re.sub(r"/\*.*?\*/|//[^\n]*", lambda m: re.sub(r"[^\n]", " ", m[0]), source, flags=re.S)


def canonical_type(value: str) -> str:
    value = " ".join(value.split())
    return re.sub(r"\s*\*\s*", " *", value)


def parameter(value: str, index: int) -> dict:
    value = " ".join(value.split())
    match = re.fullmatch(r"((?:const )?(?:double|int32|int|char|AS_BOOL|centisec|CSEC)(?:\s*\*)?)\s*(\w+)?", value)
    if not match:
        raise ValueError(f"Unreviewed C parameter: {value}")
    c_type = canonical_type(match[1])
    source_name = match[2]
    name = source_name or f"arg{index}"
    pointer = "*" in c_type
    base_type = c_type.replace("const ", "").replace(" *", "")
    return {
        "name": name, "sourceName": source_name, "vbaName": f"p_{name}",
        "cType": c_type, "vbaType": SCALAR_TYPES[base_type],
        "passing": "ByRef" if pointer else "ByVal",
        "pointer": pointer, "const": c_type.startswith("const "),
    }


def prototypes(source: str) -> dict[str, dict]:
    clean = strip_comments(source)
    pattern = r"ext_def\s*\(\s*([^)]*?)\s*\)\s+(swe_\w+)\s*\((.*?)\)\s*;"
    result = {}
    for match in re.finditer(pattern, clean, re.S):
        return_type, name, arguments = match.groups()
        return_type = canonical_type(return_type)
        args = [] if arguments.strip() == "void" else [parameter(arg, i + 1) for i, arg in enumerate(arguments.split(","))]
        if return_type == "void":
            returns = {"cType": "void", "vbaType": None, "kind": "Sub"}
        elif "*" in return_type:
            if return_type not in ("char *", "const char *"):
                raise ValueError(f"Unreviewed return pointer: {return_type}")
            returns = {"cType": return_type, "vbaType": "LongPtr", "kind": "Function"}
        else:
            returns = {"cType": return_type, "vbaType": SCALAR_TYPES[return_type], "kind": "Function"}
        if name in result:
            raise ValueError(f"Duplicate prototype: {name}")
        result[name] = {
            "name": name, "nativeName": f"Native_{name}",
            "prototype": f"{return_type} {name}({', '.join(' '.join(arg.split()) for arg in arguments.split(','))});",
            "headerLine": clean.count("\n", 0, match.start()) + 1,
            "parameters": args, "returns": returns,
        }
    if not result:
        raise ValueError("No Swiss Ephemeris prototypes found")
    return result


def contract(function: str, argument: dict) -> dict:
    name, c_type = argument["name"], argument["cType"]
    result = {"status": "pending_semantic_review", "direction": "pending", "units": "pending", "capacity": None}
    if argument["pointer"]:
        result["storage"] = "Address of first element in contiguous caller-owned storage."
    if c_type in ("char *", "const char *"):
        result["encoding"] = "NUL-terminated bytes; non-ASCII encoding policy pending wrapper review."
        if argument["const"]:
            result["direction"] = "in"
        if name == "serr":
            result.update(status="documented", direction="out", units="bytes", capacity={"minimumElements": 256, "elementType": "Byte", "includesTerminator": True}, source=DOCUMENTATION)
        if function in ("swe_version", "swe_get_library_path"):
            result.update(status="documented", direction="out", units="bytes", capacity={"minimumElements": 256, "elementType": "Byte", "includesTerminator": True}, source=DOCUMENTATION)
        return result
    if name in ("xx", "xxret") and function in ("swe_calc", "swe_calc_ut", "swe_calc_pctr", "swe_fixstar", "swe_fixstar_ut", "swe_fixstar2", "swe_fixstar2_ut"):
        result.update(status="documented", direction="out", units="Coordinates/speeds selected by flags: degrees or radians, AU; XYZ changes interpretation.", capacity={"minimumElements": 6, "elementType": "Double"}, source=DOCUMENTATION)
    elif name == "do_interpolate":
        result.update(status="documented", direction="in", units="C boolean 0/1; use Long, not VBA Boolean.", source="sweodef.h AS_BOOL typedef")
    elif c_type in ("centisec", "CSEC"):
        result.update(status="documented", direction="in", units="centiseconds (angle or time according to function)", source="sweodef.h centisec typedef and CSEC macro")
    elif c_type == "char":
        result.update(status="documented", direction="in", units="single 8-bit character code", source="swephexp.h")
    elif not argument["pointer"]:
        result["direction"] = "in"
    return result


def generate(dll_path: Path) -> tuple[str, str]:
    source_files = []
    for name, expected in SOURCE_HASHES.items():
        data = (ROOT / "vendor/swisseph" / name).read_bytes()
        if sha256(data) != expected:
            raise ValueError(f"Pinned upstream file changed: {name}")
        source_files.append({"path": f"vendor/swisseph/{name}", "sha256": expected, "url": f"https://raw.githubusercontent.com/aloistr/swisseph/{COMMIT}/{name}"})
    data = dll_path.read_bytes()
    if sha256(data) != PINNED_DLL_SHA256:
        raise ValueError("DLL hash differs from the inspected upstream release archive member")
    engine_source_path = ROOT / "vendor/swisseph/source"
    engine_source_manifest = json.loads((engine_source_path / "provenance.json").read_text())
    if engine_source_manifest["commit"] != COMMIT:
        raise ValueError("Engine source manifest does not match the pinned release commit")
    for entry in engine_source_manifest["files"]:
        if sha256((engine_source_path / entry["path"]).read_bytes()) != entry["sha256"]:
            raise ValueError("Vendored engine source changed: " + entry["path"])
    version_match = re.search(r'#define\s+SE_VERSION\s+"([^"]+)"', (engine_source_path / "sweph.h").read_text())
    if not version_match:
        raise ValueError("Missing version in pinned source")
    exports = pe_exports(data)
    source = (ROOT / "vendor/swisseph/swephexp.h").read_text()
    functions = prototypes(source)
    names = {item["name"] for item in exports}
    if names != set(functions):
        raise ValueError(f"Header/export mismatch: export-only={sorted(names - set(functions))}; header-only={sorted(set(functions) - names)}")

    lines = [
        'Attribute VB_Name = "SWNative"', "Option Explicit", "Option Private Module", "",
        "' Generated by tools/generate_bindings.py; do not hand-edit this module.",
        f"' Upstream {RELEASE_TAG}, commit {COMMIT}.",
        "' Retain vendor/swisseph/LICENSE and the notices in the pinned headers.",
        "' Raw Windows x64 ABI only: declarations do not imply tested wrappers.",
        "' Pass first elements of correctly sized numeric/Byte buffers ByRef.",
        "' C strings require a NUL terminator; LongPtr returns are pointers, not VBA Strings.",
        "' Long represents 32-bit C integers/AS_BOOL, not pointer-sized integers.",
        "' Windows x64 uses the unified Microsoft x64 calling convention.", "",
    ]
    evidence_path = ROOT / "verification/current-api.json"
    evidence = json.loads(evidence_path.read_text()) if evidence_path.exists() else {}
    sources = evidence.get("vbaSources", {})
    evidence_current = set(sources) == {p.relative_to(ROOT).as_posix() for p in (ROOT / "src/vba").glob("*.bas")} and all((ROOT / path).exists() and sha256((ROOT / path).read_bytes()) == digest for path, digest in sources.items())
    verified = set(evidence.get("verifiedFunctions", [])) if evidence_current else set()
    entries = []
    for export in exports:
        item = functions[export["name"]]
        item.update(status="native_declared", nativeDeclared=True, runtimeVerified=False, worksheetWrapped=False, experimental=item["name"] in EXPERIMENTAL, experimentalNote=EXPERIMENTAL.get(item["name"]), export=export)
        for arg in item["parameters"]:
            arg["contract"] = contract(item["name"], arg)
        item["returns"]["contract"] = {
            "status": "not_applicable" if item["returns"]["kind"] == "Sub" else "pending_semantic_review",
            "units": None if item["returns"]["kind"] == "Sub" else "pending",
            "ownership": "pending; never free the returned C string pointer" if item["returns"]["vbaType"] == "LongPtr" else None,
        }
        from api_contracts import describe
        reviewed = describe(item)
        item.update(worksheetWrapped=not reviewed["command"], worksheetName=reviewed["worksheetName"],
                    commandWrapped=reviewed["command"], family=reviewed["family"],
                    status="windows_fixture_verified" if item["name"] in verified else "safe_interface_implemented_runtime_pending", runtimeVerified=item["name"] in verified, resultKind=reviewed["resultKind"],
                    exampleInputs=[p["example"] for p in reviewed["parameters"] if p["direction"] != "out"],
                    semanticSource=reviewed["reference"], requiredFiles=reviewed["requiredFiles"], runtimeEvidence="verification/current-api.json" if item["name"] in verified else None)
        for parameter, semantic in zip(item["parameters"], reviewed["parameters"]):
            parameter["contract"] = {"status": "reviewed", "direction": semantic["direction"],
                "units": semantic["units"], "capacity": {"minimumElements": semantic["capacity"],
                "visibleElements": semantic["visible"], "elementType": parameter["vbaType"]} if parameter["pointer"] else None}
            if semantic["string"]:
                parameter["contract"].update(encoding="Windows ANSI; lossless conversion required", includesTerminator=True)
        item["returns"]["contract"] = {"status": "reviewed", "convention": reviewed["resultKind"],
            "ownership": "borrowed or caller-buffer alias; never freed; bounded copy" if item["returns"]["vbaType"] == "LongPtr" else None}
        entries.append(item)
        if item["experimental"]:
            lines.append("' EXPERIMENTAL / upstream internal: " + item["experimentalNote"])
        lines.append(f"' {item['prototype']}")
        kind = item["returns"]["kind"]
        lines.append(f'Public Declare PtrSafe {kind} {item["nativeName"]} Lib "{DLL_NAME}" Alias "{item["name"]}" _')
        suffix = "" if kind == "Sub" else " As " + item["returns"]["vbaType"]
        if item["parameters"]:
            lines.append("    ( _")
            for index, arg in enumerate(item["parameters"]):
                ending = ", _" if index < len(item["parameters"]) - 1 else " _"
                lines.append(f"        {arg['passing']} {arg['vbaName']} As {arg['vbaType']}{ending}")
            lines.append("    )" + suffix)
        else:
            lines.append("    ()" + suffix)
        lines.append("")

    catalogue = {
        "schemaVersion": 1,
        "checkpoint": "Native ABI plus reviewed safe interfaces; runtime evidence is recorded separately.",
        "engine": {
            "releaseTag": RELEASE_TAG, "sourceCommit": COMMIT,
            "releaseUrl": f"{UPSTREAM}/releases/tag/{RELEASE_TAG}",
            "sourceVersion": version_match[1],
            "runtimeVersion": evidence.get("engine", {}).get("runtimeVersion") if verified else None, "runtimeVersionStatus": "executed_on_recorded_windows_excel" if verified else "not_executed_for_current_sources",
            "latestSourceBuildStatus": "passed_recorded_windows_msvc_build" if verified else "see_source_build_evidence",
            "dllName": DLL_NAME,
            "binary": {"path": dll_path.resolve().relative_to(ROOT).as_posix(), "role": "upstream_prebuilt_export_inspection_reference", "latestSourceBuildVerified": False, "sha256": sha256(data), "sizeBytes": len(data), "machine": "AMD64", "peFormat": "PE32+", "exportCount": len(exports)},
        },
        "source": {"files": source_files, "engineSourceManifest": {"path": "vendor/swisseph/source/provenance.json", "sha256": sha256((engine_source_path / "provenance.json").read_bytes())}, "documentationUrl": DOCUMENTATION, "licenseNotice": "Preserved upstream dual-license notices; selected project license is recorded separately by the project."},
        "abi": {"target": "Windows x64 Excel VBA7", "callingConvention": "Microsoft x64 unified ABI", "integerBits": 32, "pointerBits": 64, "doubleBits": 64, "charBits": 8, "charPointerParameters": "ByRef Byte", "numericPointerParameters": "ByRef matching Long or Double", "charPointerReturns": "LongPtr", "voidReturns": "Sub", "nullPointerPolicy": "Raw ByRef buffer declarations expect valid storage; optional null-pointer use needs a separately reviewed declaration/wrapper."},
        "summary": {"nativeDeclared": len(entries), "runtimeVerified": len(verified), "worksheetWrapped": sum(e["worksheetWrapped"] for e in entries), "commandWrapped": sum(e["commandWrapped"] for e in entries), "experimental": len(EXPERIMENTAL), "headerPrototypesNotExported": [], "exportsWithoutHeaderPrototype": []},
        "functions": entries,
    }
    return "\n".join(lines), json.dumps(catalogue, indent=2, ensure_ascii=True) + "\n"


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--dll", type=Path, default=ROOT / "vendor/engine/swedll64.dll")
    parser.add_argument("--check", action="store_true", help="Verify generated content without writing")
    args = parser.parse_args()
    module, catalogue = generate(args.dll)
    outputs = [(ROOT / "src/vba/SWNative.bas", module), (ROOT / "api/catalog.json", catalogue)]
    if args.check:
        different = [str(path.relative_to(ROOT)) for path, content in outputs if not path.exists() or path.read_text() != content]
        if different:
            parser.error("Generated files differ: " + ", ".join(different))
        print("PASS: pinned header/DLL hashes, PE exports, C prototypes, VBA declarations and catalogue agree.")
    else:
        for path, content in outputs:
            path.parent.mkdir(parents=True, exist_ok=True)
            path.write_text(content)
        print(f"Generated {len(pe_exports(args.dll.read_bytes()))} native declarations and catalogue entries.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
