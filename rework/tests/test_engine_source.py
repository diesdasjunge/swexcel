"""Pinned-source completeness and optional fresh Windows build evidence.

These checks inspect source and PE files; they never execute a Windows DLL.
"""

import hashlib
import json
from pathlib import Path
import re
import unittest

from _support import REWORK, pe_exports
from test_package_integrity import safe_relative


COMMIT = "f4dcd18e8005dde95fd8a8d2312ed12f9accd1b0"
SOURCE_MANIFEST_SHA256 = "4312a79b62fec145a8e4c359df8a1d72f35f075b0122bbdc5944ea9c7747f375"
LIBRARY = "swexcel-se-2.10.3b-x64.dll"


class EngineSourceTests(unittest.TestCase):
    def test_source_manifest_closes_over_every_pinned_file(self):
        directory = REWORK / "vendor/swisseph/source"
        manifest_path = directory / "provenance.json"
        data = manifest_path.read_bytes()
        self.assertEqual(SOURCE_MANIFEST_SHA256, hashlib.sha256(data).hexdigest())
        manifest = json.loads(data)
        self.assertEqual(COMMIT, manifest["commit"])
        self.assertEqual("v2.10.3bfinal", manifest["releaseTag"])
        declared = [record["path"] for record in manifest["files"]]
        self.assertEqual(len(declared), len(set(declared)))
        actual = {path.relative_to(directory).as_posix() for path in directory.rglob("*")
                  if path.is_file() and path != manifest_path}
        self.assertEqual(actual, set(declared), "Source files must be declared and hash-pinned")
        for record in manifest["files"]:
            with self.subTest(path=record["path"]):
                path = directory / safe_relative(record["path"])
                data = path.read_bytes()
                self.assertEqual(record["sizeBytes"], len(data))
                self.assertEqual(record["sha256"], hashlib.sha256(data).hexdigest())
                expected_url = f"https://raw.githubusercontent.com/aloistr/swisseph/{COMMIT}/{record['path']}"
                self.assertEqual(expected_url, record["url"])

    def test_all_local_include_dependencies_are_present(self):
        directory = REWORK / "vendor/swisseph/source"
        files = [path for path in directory.iterdir() if path.suffix in (".c", ".h")]
        self.assertGreater(len(files), 0)
        for path in files:
            # Comments may contain sample includes, so strip them before scanning.
            source = re.sub(r"/\*.*?\*/|//[^\n]*", "", path.read_text(), flags=re.S)
            dependencies = re.findall(r'^\s*#\s*include\s*"([^"]+)"', source, re.M)
            for dependency in dependencies:
                with self.subTest(file=path.name, dependency=dependency):
                    self.assertTrue((path.parent / safe_relative(dependency)).is_file(),
                                    "Pinned source is missing a transitive include")
        for name in ("swephexp.h", "sweodef.h"):
            self.assertEqual((directory / name).read_bytes(),
                             (REWORK / "vendor/swisseph" / name).read_bytes(),
                             "Binding and compiler headers must describe the same ABI")

    def test_optional_fresh_build_evidence_matches_binary_and_source(self):
        roots = [REWORK / "package/runtime/engine"]
        roots.extend(path / "runtime/engine" for path in (REWORK / "dist").glob("SWExcel-*") if path.is_dir())
        evidence_paths = [root / "build-provenance.json" for root in roots
                          if (root / "build-provenance.json").is_file()]
        if not evidence_paths:
            self.skipTest("Fresh Windows/MSVC build has not produced evidence")
        catalogue = json.loads((REWORK / "api/catalog.json").read_text())
        expected_exports = {function["name"] for function in catalogue["functions"]}
        source_manifest = REWORK / "vendor/swisseph/source/provenance.json"
        source_version = re.search(r'#define\s+SE_VERSION\s+"([^"]+)"',
                                   (source_manifest.parent / "sweph.h").read_text()).group(1)
        for evidence_path in evidence_paths:
            with self.subTest(evidence=evidence_path):
                evidence = json.loads(evidence_path.read_text(encoding="utf-8-sig"))
                self.assertEqual(1, evidence["schemaVersion"])
                self.assertEqual(COMMIT, evidence["sourceCommit"])
                self.assertEqual(hashlib.sha256(source_manifest.read_bytes()).hexdigest(),
                                 evidence["sourceManifestSha256"])
                self.assertEqual(source_version, evidence["sourceVersion"])
                self.assertEqual(expected_exports, set(evidence["exportedNames"]))
                assets = evidence["assets"]
                names = [asset["path"] for asset in assets]
                self.assertEqual(len(names), len(set(names)))
                self.assertIn(LIBRARY, names)
                for asset in assets:
                    binary = evidence_path.parent / safe_relative(asset["path"])
                    data = binary.read_bytes()
                    self.assertEqual(asset["sizeBytes"], len(data))
                    self.assertEqual(asset["sha256"], hashlib.sha256(data).hexdigest())
                    pe = pe_exports(data) if binary.suffix.lower() == ".dll" else None
                    if pe is not None:
                        self.assertEqual(0x8664, pe.machine)
                        self.assertEqual(0x20B, pe.optional_magic)
                        self.assertEqual(expected_exports, set(pe.names))
                self.assertTrue(evidence["compilerPath"])
                self.assertRegex(evidence["compilerSha256"], r"^[0-9a-f]{64}$")
                self.assertRegex(evidence["loaderGlueSha256"], r"^[0-9a-f]{64}$")
                self.assertTrue(evidence["completedAtUtc"])


if __name__ == "__main__":
    unittest.main()
