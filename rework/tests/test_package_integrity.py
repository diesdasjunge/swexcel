"""Source pinning and prepared-distribution integrity, independent of prepare.py."""

import hashlib
import json
from pathlib import Path, PurePosixPath
import unittest

from _support import PROJECT, REWORK


def safe_relative(value):
    path = PurePosixPath(value)
    if not value or path.is_absolute() or ".." in path.parts or "\\" in value or ":" in value:
        raise AssertionError(f"Unsafe package path: {value}")
    return Path(*path.parts)


def verify_record(test, root, record):
    path = root / safe_relative(record["path"])
    test.assertTrue(path.is_file(), f"Missing manifest file: {path}")
    data = path.read_bytes()
    test.assertEqual(record["size"], len(data), f"Size differs: {path}")
    test.assertEqual(record["sha256"], hashlib.sha256(data).hexdigest(), f"Hash differs: {path}")


class PackageIntegrityTests(unittest.TestCase):
    def test_pinned_vendor_inputs_match_independent_snapshot(self):
        fixture = Path(__file__).with_name("fixtures") / "pinned-vendor-assets.json"
        expected = json.loads(fixture.read_text())
        for asset in expected["assets"]:
            with self.subTest(path=asset["path"]):
                verify_record(self, REWORK, asset)
        provenance = json.loads((REWORK / "vendor/engine/provenance.json").read_text())
        for field in ("releaseTag", "commit", "archiveSha256"):
            self.assertEqual(expected[field], provenance[field])

    def test_authoritative_asset_manifest_has_exact_runtime_bundle(self):
        manifest = json.loads((REWORK / "package-manifest.json").read_text())
        assets = manifest["assets"]
        paths = [asset["path"] for asset in assets]
        self.assertEqual(len(paths), len(set(paths)), "Duplicate packaged asset paths")
        expected = {
            "runtime/engine/swexcel-se-2.10.3b-x64.dll",
            *("runtime/ephe/" + name for name in ("sepl_18.se1", "semo_18.se1", "seas_18.se1", "sefstars.txt", "seasnam.txt", "seorbel.txt")),
        }
        runtime = {path for path in paths if path.startswith("runtime/")}
        self.assertEqual(expected, runtime, "Unexpected or missing runtime engine/data asset")
        for asset in assets:
            with self.subTest(path=asset["path"]):
                safe_relative(asset["path"])
                source = PROJECT / safe_relative(asset["sourcePath"])
                self.assertTrue(source.is_file(), f"Missing asset source: {source}")
                data = source.read_bytes()
                self.assertEqual(asset["size"], len(data))
                self.assertEqual(asset["sha256"], hashlib.sha256(data).hexdigest())
                self.assertTrue(asset["license"])
                self.assertTrue(asset["kind"])
                self.assertTrue(asset["coverage"])
                self.assertTrue(asset["sourceUrl"].startswith("https://"))

    def test_prepared_packages_have_complete_hash_manifests(self):
        packages = sorted((REWORK / "dist").glob("SWExcel-*"))
        if not packages:
            self.skipTest("Run prepare.py to enable generated-package verification")
        for package in packages:
            if not package.is_dir():
                continue
            with self.subTest(package=package.name):
                manifest = json.loads((package / "package-files.json").read_text())
                self.assertEqual(1, manifest["schemaVersion"])
                version = (package / "VERSION").read_text().strip()
                self.assertEqual(version, manifest["projectVersion"])
                self.assertEqual(f"SWExcel-{version}", package.name)
                records = manifest["files"]
                declared = [record["path"] for record in records]
                self.assertEqual(sorted(declared), declared, "Manifest is not deterministic")
                self.assertEqual(len(declared), len(set(declared)))
                self.assertNotIn("package-files.json", declared)
                actual = {path.relative_to(package).as_posix() for path in package.rglob("*") if path.is_file() and path.name != "package-files.json"}
                self.assertEqual(actual, set(declared), "Unlisted or missing files in prepared package")
                for record in records:
                    verify_record(self, package, record)
                source_assets = json.loads((package / "package-manifest.json").read_text())["assets"]
                build_path = package / "runtime/engine/build-provenance.json"
                built_assets = {}
                if build_path.is_file():
                    build = json.loads(build_path.read_text(encoding="utf-8-sig"))
                    built_assets = {"runtime/engine/" + record["path"]: record for record in build["assets"]}
                for asset in source_assets:
                    # The initial DLL is an export-inspection reference. A fresh
                    # Windows build supersedes only that asset, with an explicit
                    # source/build provenance record checked by test_engine_source.
                    if asset["kind"] == "engine_reference" and build_path.is_file():
                        self.assertIn(asset["path"], built_assets)
                        built = built_assets[asset["path"]]
                        verify_record(self, package, {"path": asset["path"],
                                                     "size": built["sizeBytes"],
                                                     "sha256": built["sha256"]})
                    else:
                        verify_record(self, package, asset)


if __name__ == "__main__":
    unittest.main()
