"""Static binary identity and original-preservation checks; no DLL execution."""

import hashlib
import json
from pathlib import Path
import unittest

from _support import PROJECT, REWORK, pe_exports


class PEAndOriginalAssetsTests(unittest.TestCase):
    def test_original_assets_remain_byte_identical(self):
        fixture = Path(__file__).with_name("fixtures") / "original-assets.json"
        for asset in json.loads(fixture.read_text())["assets"]:
            with self.subTest(path=asset["path"]):
                data = (PROJECT / asset["path"]).read_bytes()
                self.assertEqual(asset["size"], len(data))
                self.assertEqual(asset["sha256"], hashlib.sha256(data).hexdigest())

    def test_original_dll_is_windows_x64_with_106_real_named_exports(self):
        exports = pe_exports((PROJECT / "ephem/swedll64.dll").read_bytes())
        self.assertEqual(0x8664, exports.machine)
        self.assertEqual(0x20B, exports.optional_magic)
        self.assertTrue(exports.characteristics & 0x2000, "PE is not marked as a DLL")
        self.assertEqual(106, len(exports.names))
        self.assertEqual(106, exports.function_count)
        self.assertIn("swe_revjul", exports.names)
        self.assertNotIn("_swe_revjul", exports.names)

    def test_modern_engine_is_x64_and_exports_are_named(self):
        path = REWORK / "vendor/engine/swedll64.dll"
        self.assertTrue(path.is_file(), "Pinned modern engine has not been prepared")
        exports = pe_exports(path.read_bytes())
        self.assertEqual(0x8664, exports.machine)
        self.assertEqual(0x20B, exports.optional_magic)
        self.assertTrue(exports.characteristics & 0x2000)
        self.assertEqual(exports.function_count, len(exports.names), "Unexpected ordinal-only exports")
        self.assertGreaterEqual(len(exports.names), 106)
        self.assertTrue(all(name.startswith("swe_") for name in exports.names))

    def test_pe_parser_rejects_truncated_and_non_pe_inputs(self):
        real = (PROJECT / "ephem/swedll64.dll").read_bytes()
        for invalid in (b"", b"swe_revjul\0", real[:64], b"ZZ" + real[2:]):
            with self.subTest(size=len(invalid)):
                with self.assertRaises(ValueError):
                    pe_exports(invalid)


if __name__ == "__main__":
    unittest.main()
