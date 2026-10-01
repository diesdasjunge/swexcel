#!/usr/bin/env python3
"""Verify pinned assets and stage the source package; Excel creates the XLSM on Windows."""
from __future__ import annotations

import argparse
import hashlib
import json
from pathlib import Path, PurePosixPath
import shutil
import tempfile
import zipfile

ROOT = Path(__file__).resolve().parents[1]
REPO = ROOT.parent
SOURCE_DIRECTORIES = ("api", "docs", "src", "tools", "vendor", "workbook", "tests")
SOURCE_FILES = (".gitattributes", "VERSION", "README.md", "PLAN.md", "ROADMAP.md", "CHANGELOG.md", "LICENSE", "NOTICE.md", "package-manifest.json")


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def safe_relative(value: str) -> Path:
    path = PurePosixPath(value)
    if path.is_absolute() or not path.parts or any(part in (".", "..") for part in path.parts):
        raise ValueError(f"Unsafe manifest path: {value!r}")
    return Path(*path.parts)


def verify_assets() -> dict:
    manifest = json.loads((ROOT / "package-manifest.json").read_text())
    binary_names = sorted(Path(asset["path"]).name for asset in manifest["assets"] if asset["kind"] == "binary_ephemeris")
    if binary_names != ["seas_18.se1", "semo_18.se1", "sepl_18.se1"]:
        raise ValueError("V1 bundles exactly seas_18.se1, semo_18.se1 and sepl_18.se1")
    targets = set()
    for asset in manifest["assets"]:
        target = safe_relative(asset["path"])
        if target.as_posix() in targets:
            raise ValueError(f"Duplicate package asset: {target}")
        targets.add(target.as_posix())
        source = REPO / safe_relative(asset["sourcePath"])
        if source.stat().st_size != asset["size"] or sha256(source) != asset["sha256"]:
            raise ValueError(f"Pinned asset changed: {asset['sourcePath']}")
    return manifest


def file_manifest(directory: Path, version: str) -> dict:
    return {
        "schemaVersion": 1,
        "projectVersion": version,
        "files": [
            {"path": path.relative_to(directory).as_posix(), "sha256": sha256(path), "size": path.stat().st_size}
            for path in sorted(directory.rglob("*"))
            if path.is_file() and path.name != "package-files.json"
        ],
    }


def prepare() -> Path:
    manifest = verify_assets()
    version = (ROOT / "VERSION").read_text().strip()
    if not version or any(char not in "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ.-" for char in version):
        raise ValueError("Invalid VERSION")
    distribution = ROOT / "dist"
    distribution.mkdir(exist_ok=True)
    destination = distribution / f"SWExcel-{version}"
    # A Windows-tested package must not be silently replaced by a source-only stage.
    if destination.exists():
        raise FileExistsError(f"Package already exists: {destination}. Use a new version or remove that generated directory deliberately.")
    with tempfile.TemporaryDirectory(prefix=".stage-", dir=distribution) as temporary:
        stage = Path(temporary)
        for name in SOURCE_DIRECTORIES:
            shutil.copytree(ROOT / name, stage / name, ignore=shutil.ignore_patterns("__pycache__", "*.pyc"))
        for name in SOURCE_FILES:
            shutil.copy2(ROOT / name, stage / name)
        for asset in manifest["assets"]:
            target = stage / safe_relative(asset["path"])
            target.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(REPO / safe_relative(asset["sourcePath"]), target)
        pending = {
            "checkpoint": "integration_preparation",
            "publicReleaseReady": False,
            "engineSourceBuild": "pending_windows_msvc",
            "vbaCompilation": "pending_windows_excel",
            "scalarAndSpillAcceptance": "pending_windows_excel",
            "workbookFile": "pending_windows_build",
            "note": "This source package is not a tested XLSM release. See docs/WINDOWS-TESTING.md.",
        }
        (stage / "checkpoint-status.json").write_text(json.dumps(pending, indent=2) + "\n")
        files = file_manifest(stage, version)
        (stage / "package-files.json").write_text(json.dumps(files, indent=2) + "\n")
        stage.rename(destination)
    archive = distribution / f"SWExcel-{version}-source-checkpoint.zip"
    with zipfile.ZipFile(archive, "w", zipfile.ZIP_DEFLATED) as output:
        for path in sorted(destination.rglob("*")):
            if path.is_file():
                output.write(path, arcname=f"{destination.name}/{path.relative_to(destination).as_posix()}")
    return destination


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--verify-only", action="store_true")
    arguments = parser.parse_args()
    if arguments.verify_only:
        verify_assets()
        print("Pinned runtime/data assets verified; Windows acceptance remains pending.")
    else:
        destination = prepare()
        print(f"Source checkpoint prepared: {destination}")
        print("Next: build-engine.ps1, then build-workbook.ps1 in Windows. Do not treat this as a public release.")


if __name__ == "__main__":
    main()
