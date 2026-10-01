"""Compare independently parsed PE, official C declarations, VBA, and catalog.

Passing these tests establishes static agreement, not that Excel can load or
call the DLL. The Windows integration/build acceptance remains mandatory.
"""

import json
import re
import unittest

from _support import REWORK, c_prototypes, pe_exports, vba_declarations


def c_type(parameter):
    return re.sub(r"\b[A-Za-z_]\w*\s*$", "", parameter).strip() if not parameter.rstrip().endswith("*") else parameter


def normalized_type(value):
    return re.sub(r"\s+", " ", value.replace("*", " * ")).strip()


def scalar_vba_type(value):
    value = value.replace("const", "").strip()
    return {
        "int": "Long",
        "int32": "Long",
        "AS_BOOL": "Long",
        "centisec": "Long",
        "CSEC": "Long",
        "double": "Double",
        "char": "Byte",
    }[value]


class NativeABITests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.prototypes = c_prototypes((REWORK / "vendor/swisseph/swephexp.h").read_text())
        cls.exports = pe_exports((REWORK / "vendor/engine/swedll64.dll").read_bytes())
        cls.native_text = (REWORK / "src/vba/SWNative.bas").read_text()
        cls.native = vba_declarations(cls.native_text)
        cls.catalog = json.loads((REWORK / "api/catalog.json").read_text())

    def test_pe_header_native_and_catalog_cover_same_exports(self):
        expected = set(self.exports.names)
        self.assertEqual(expected, set(self.prototypes), "C header differs from actual PE exports")
        self.assertEqual(expected, set(self.native), "VBA omits or invents native exports")
        functions = self.catalog["functions"]
        names = [function["name"] for function in functions]
        self.assertEqual(len(names), len(set(names)), "Duplicate catalog entries")
        self.assertEqual(expected, set(names), "Catalog omits or invents native exports")
        self.assertNotIn("swe_set_timeout", expected, "Commented-out C API was treated as an export")

    def test_every_native_signature_matches_c_abi(self):
        for name, (returns, parameters) in self.prototypes.items():
            with self.subTest(function=name):
                declaration = self.native[name]
                self.assertEqual("Native_" + name, declaration["name"])
                self.assertTrue(declaration["ptrsafe"])
                self.assertEqual("swexcel-se-2.10.3b-x64.dll", declaration["library"])
                self.assertEqual(len(parameters), len(declaration["parameters"]))
                if returns == "void":
                    self.assertEqual("sub", declaration["kind"].lower())
                    self.assertFalse(declaration["returns"])
                else:
                    expected = "LongPtr" if "*" in returns else scalar_vba_type(returns)
                    self.assertEqual("function", declaration["kind"].lower())
                    self.assertEqual(expected.lower(), declaration["returns"].lower())
                for c_parameter, vba_parameter in zip(parameters, declaration["parameters"]):
                    passing, _, vba_type = vba_parameter
                    native_type = c_type(c_parameter)
                    if "*" in native_type:
                        base = native_type.replace("*", "").replace("const", "").strip()
                        typed_reference = ("byref", scalar_vba_type(base).lower())
                        actual = (passing.lower(), vba_type.lower())
                        # ByVal LongPtr also represents a native pointer correctly.
                        self.assertIn(actual, (typed_reference, ("byval", "longptr")))
                    else:
                        self.assertEqual("byval", passing.lower())
                        self.assertEqual(scalar_vba_type(native_type).lower(), vba_type.lower())

    def test_critical_pointer_and_julian_return_contracts(self):
        for name in ("swe_version", "swe_get_library_path", "swe_get_planet_name", "swe_get_ayanamsa_name", "swe_get_current_file_data", "swe_house_name"):
            with self.subTest(function=name):
                self.assertEqual("longptr", self.native[name]["returns"].lower())
        self.assertEqual(("byval", "double"), (self.native["swe_revjul"]["parameters"][0][0].lower(), self.native["swe_revjul"]["parameters"][0][2].lower()))
        self.assertEqual("sub", self.native["swe_revjul"]["kind"].lower())
        self.assertEqual("long", self.native["swe_get_ayanamsa_name"]["parameters"][0][2].lower())
        self.assertNotIn("_swe_revjul", self.native)

    def test_catalog_parameters_describe_actual_signatures(self):
        for function in self.catalog["functions"]:
            with self.subTest(function=function["name"]):
                returns, parameters = self.prototypes[function["name"]]
                self.assertEqual(normalized_type(returns), normalized_type(function["returns"]["cType"]))
                self.assertEqual(len(parameters), len(function["parameters"]))
                for c_parameter, parameter in zip(parameters, function["parameters"]):
                    self.assertEqual(normalized_type(c_type(c_parameter)), normalized_type(parameter["cType"]))
                    self.assertEqual("*" in c_parameter, parameter["pointer"])

    def test_native_source_is_explicit_and_hidden_from_worksheet_names(self):
        self.assertRegex(self.native_text, r"(?im)^Option Explicit\s*$")
        self.assertRegex(self.native_text, r"(?im)^Option Private Module\s*$")
        self.assertNotRegex(self.native_text, r"(?i)\bAs\s+(?:Integer|Boolean|Single|String)\b")


if __name__ == "__main__":
    unittest.main()
