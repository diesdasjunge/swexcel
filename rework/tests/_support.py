"""Independent, standard-library parsers used by repository verification.

These inspect packaging and ABI declarations. They do not execute Excel or the
Windows DLL and must never be reported as Windows runtime acceptance tests.
"""

from dataclasses import dataclass
from pathlib import Path
import re
import struct


PROJECT = Path(__file__).resolve().parents[2]
REWORK = PROJECT / "rework"


@dataclass(frozen=True)
class PEExports:
    machine: int
    optional_magic: int
    characteristics: int
    function_count: int
    names: tuple[str, ...]
    ordinals: tuple[int, ...]


def pe_exports(data: bytes) -> PEExports:
    """Read named exports from an actual PE section mapping, not string scans."""

    def unpack(fmt: str, offset: int):
        size = struct.calcsize(fmt)
        if offset < 0 or offset + size > len(data):
            raise ValueError("Truncated PE structure")
        return struct.unpack_from(fmt, data, offset)

    if len(data) < 64 or data[:2] != b"MZ":
        raise ValueError("Not a DOS/PE image")
    pe = unpack("<I", 0x3C)[0]
    if data[pe : pe + 4] != b"PE\0\0":
        raise ValueError("Missing PE signature")
    machine, section_count, _, _, _, optional_size, characteristics = unpack(
        "<HHIIIHH", pe + 4
    )
    optional = pe + 24
    magic = unpack("<H", optional)[0]
    if magic == 0x20B:
        directory_offset, count_offset = 112, 108
    elif magic == 0x10B:
        directory_offset, count_offset = 96, 92
    else:
        raise ValueError("Unsupported PE optional-header magic")
    if optional_size < directory_offset + 8:
        raise ValueError("Missing export data directory")
    if unpack("<I", optional + count_offset)[0] < 1:
        raise ValueError("No PE data directories")
    export_rva, export_size = unpack("<II", optional + directory_offset)
    if not export_rva or export_size < 40:
        raise ValueError("Missing export table")
    sections = []
    for index in range(section_count):
        section = unpack("<8sIIIIIIHHI", optional + optional_size + index * 40)
        sections.append((section[2], section[1], section[4], section[3]))

    def mapped(rva: int, length: int = 1) -> int:
        for virtual_address, virtual_size, raw_pointer, raw_size in sections:
            delta = rva - virtual_address
            if 0 <= delta < max(virtual_size, raw_size):
                if delta + length > raw_size:
                    raise ValueError("Export RVA points outside initialized section data")
                offset = raw_pointer + delta
                if offset + length > len(data):
                    raise ValueError("Section data is truncated")
                return offset
        raise ValueError("Unmapped export RVA")

    def c_string(rva: int) -> str:
        chars = bytearray()
        for delta in range(4096):
            char = data[mapped(rva + delta)]
            if not char:
                return chars.decode("ascii")
            chars.append(char)
        raise ValueError("Unterminated export name")

    table = unpack("<IIHHIIIIIII", mapped(export_rva, 40))
    base, functions, named = table[5:8]
    function_table = mapped(table[8], functions * 4)
    name_table = mapped(table[9], named * 4)
    ordinal_table = mapped(table[10], named * 2)
    names, ordinals = [], []
    for index in range(named):
        name_rva = unpack("<I", name_table + index * 4)[0]
        ordinal = unpack("<H", ordinal_table + index * 2)[0]
        if ordinal >= functions:
            raise ValueError("Name ordinal exceeds function table")
        function_rva = unpack("<I", function_table + ordinal * 4)[0]
        if not function_rva:
            raise ValueError("Named export has no function address")
        mapped(function_rva)
        names.append(c_string(name_rva))
        ordinals.append(base + ordinal)
    if len(set(names)) != len(names):
        raise ValueError("Duplicate exported names")
    return PEExports(machine, magic, characteristics, functions, tuple(names), tuple(ordinals))


def without_c_comments(text: str) -> str:
    return re.sub(r"/\*.*?\*/|//[^\n]*", "", text, flags=re.S)


def c_prototypes(text: str) -> dict[str, tuple[str, list[str]]]:
    """Parse the pinned public header without executing the generator under test."""
    matches = re.findall(
        r"ext_def\s*\(\s*([^)]*?)\s*\)\s*(swe_\w+)\s*\((.*?)\)\s*;",
        without_c_comments(text),
        flags=re.S,
    )
    result = {}
    for returns, name, parameters in matches:
        if name in result:
            raise ValueError(f"Duplicate C declaration: {name}")
        params = [" ".join(param.split()) for param in parameters.split(",")]
        if params == ["void"] or params == [""]:
            params = []
        result[name] = (" ".join(returns.split()), params)
    if not result:
        raise ValueError("No public C prototypes found")
    return result


def without_vba_comments(text: str) -> str:
    # Apostrophes in quoted strings are data, not comment delimiters.
    return "\n".join(
        re.sub(r"('.*)$", "", re.sub(r'"(?:[^\"]|\"\")*"', '""', line))
        for line in text.splitlines()
    )


def vba_declarations(text: str) -> dict[str, dict]:
    """Read actual Declare statements, retaining alias strings and ABI widths."""
    lines = []
    for line in text.splitlines():
        # Native declarations use no quoted apostrophes; preserve Lib/Alias values.
        line = re.sub(r"\s+'[^\n]*$", "", line)
        lines.append(line)
    text = re.sub(r"_\s*\n\s*", " ", "\n".join(lines))
    pattern = re.compile(
        r"(?:Public|Private)\s+Declare\s+(PtrSafe\s+)?(Function|Sub)\s+"
        r"(Native_swe_\w+)\s+Lib\s+\"([^\"]+)\"\s+Alias\s+\"([^\"]+)\""
        r"\s*\((.*?)\)(?:\s+As\s+(\w+))?",
        re.I | re.S,
    )
    result = {}
    for ptrsafe, kind, name, library, alias, params, returns in pattern.findall(text):
        if alias in result:
            raise ValueError(f"Duplicate native alias: {alias}")
        parameters = []
        for parameter in params.split(","):
            if not parameter.strip():
                continue
            parsed = re.fullmatch(
                r"\s*(ByVal|ByRef)\s+(\w+)\s+As\s+(\w+)\s*",
                parameter,
                re.I,
            )
            if not parsed:
                raise ValueError(f"Unparsed native VBA parameter: {parameter}")
            parameters.append(parsed.groups())
        result[alias] = {
            "name": name,
            "ptrsafe": bool(ptrsafe),
            "kind": kind,
            "library": library,
            "parameters": parameters,
            "returns": returns,
        }
    return result
