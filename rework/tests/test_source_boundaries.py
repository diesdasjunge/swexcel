"""Static policy regressions, not a VBA compiler or native execution substitute."""

import re
import unittest

from _support import REWORK, without_vba_comments


def procedure(source, name):
    match = re.search(
        rf"(?ims)^(?:Public|Private) (?:Function|Sub) {re.escape(name)}\b.*?^End (?:Function|Sub)\s*$",
        source,
    )
    if not match:
        raise AssertionError(f"Missing VBA procedure {name}")
    return match.group(0)


class SourceBoundaryTests(unittest.TestCase):
    def test_vba_modules_fit_the_editor_source_line_limits(self):
        modules = sorted((REWORK / "src/vba").glob("*.bas"))
        self.assertTrue(modules)
        for module in modules:
            continuations = 0
            for number, line in enumerate(module.read_text().splitlines(), 1):
                with self.subTest(module=module.name, line=number):
                    self.assertLessEqual(len(line), 1023, "VBA physical line exceeds its editor limit")
                    continuations = continuations + 1 if re.search(r"\s_\s*$", line) else 0
                    self.assertLessEqual(continuations, 24, "VBA statement has too many line continuations")

    def test_engine_validation_failure_cannot_become_a_ready_retry(self):
        source = (REWORK / "src/vba/SWRuntime.bas").read_text()
        ensure = procedure(source, "SWEnsureEngine")
        # LoadLibrary may succeed before export/version validation fails. A retry
        # must pass those checks rather than treating the nonzero handle as ready.
        code = without_vba_comments(ensure)
        ready_assignments = list(re.finditer(r"(?im)^\s*mReady\s*=\s*True\s*$", code))
        self.assertEqual(1, len(ready_assignments))
        ready = ready_assignments[0].start()
        for check in ("GetProcAddress", "VarPtr(versionBuffer(0))", "SWBufferText", "SW_ENGINE_VERSION", "Native_swe_set_ephe_path"):
            self.assertGreater(ready, code.index(check), f"Ready is set before {check}")
        cached_branch = code.split("If mEngine <> 0 Then", 1)[1].split("Else", 1)[0]
        exits = [line.strip() for line in cached_branch.splitlines() if re.search(r"(?i)\bExit Sub\b", line)]
        self.assertEqual(["If mReady Then Exit Sub"], exits,
                         "An engine handle alone must not bypass incomplete validation")

    def test_path_capacity_matches_engine_guard_and_strings_are_bounded(self):
        engine = (REWORK / "vendor/swisseph/source/sweph.c").read_text()
        header = (REWORK / "vendor/swisseph/source/sweodef.h").read_text()
        constants = (REWORK / "src/vba/SWConstants.bas").read_text()
        runtime = (REWORK / "src/vba/SWRuntime.bas").read_text()
        functions = (REWORK / "src/vba/SWFunctions.bas").read_text()
        as_maxch = int(re.search(r"#define\s+AS_MAXCH\s+(\d+)", header).group(1))
        guard = re.search(r"strlen\(path\)\s*<=\s*AS_MAXCH\s*-\s*(\d+)\s*-\s*(\d+)", engine)
        self.assertIsNotNone(guard, "Pinned engine's path capacity contract changed")
        maximum_chars = as_maxch - sum(map(int, guard.groups()))
        declared = int(re.search(r"SW_EPHE_PATH_BYTES\s+As\s+Long\s*=\s*(\d+)", constants).group(1))
        self.assertEqual(maximum_chars + 1, declared, "Byte capacity must include the terminal NUL")
        self.assertRegex(runtime, r"SWAnsiZ\(data,\s*SW_EPHE_PATH_BYTES\)")
        self.assertRegex(runtime, r"SWAnsiZ\(SWDataPath\(\),\s*SW_EPHE_PATH_BYTES\)")
        self.assertNotRegex(without_vba_comments(runtime), r"(?i)\b(?:lstrlenA|SWPointerText|CopyMemory)\b")
        self.assertRegex(runtime, r"versionPointer\s*<>\s*VarPtr\(versionBuffer\(0\)\)")
        name = procedure(functions, "SW_BODY_NAME")
        self.assertRegex(name, r"pointer\s*<>\s*VarPtr\(bytes\(0\)\)")
        self.assertRegex(name, r"SWBufferText\(bytes\)")

    def test_runtime_does_not_mutate_processwide_paths_or_unload_cached_native_code(self):
        source = (REWORK / "src/vba/SWRuntime.bas").read_text()
        code = without_vba_comments(source)
        forbidden = r"\b(?:ChDir|ChDrive|SetCurrentDirectory[AW]?|SetDllDirectory[AW]?|SetEnvironmentVariable[AW]?|FreeLibrary)\b"
        self.assertNotRegex(code, forbidden)
        self.assertIn("LoadLibraryExW", code)
        self.assertIn("GetModuleFileNameW", code)
        self.assertNotIn("C:\\sweph", source)

    def test_worksheet_functions_do_not_write_cells_or_show_ui(self):
        source = (REWORK / "src/vba/SWFunctions.bas").read_text()
        code = without_vba_comments(source)
        forbidden = r"(?i)\b(?:MsgBox|InputBox|Shell|DoEvents)\b|\.(?:Range|Cells|Value2|Formula2?)\b"
        self.assertNotRegex(code, forbidden)
        functions = set(re.findall(r"(?im)^Public Function (SW_\w+)\(", source))
        self.assertTrue({"SW_JULDAY", "SW_DEGNORM", "SW_VERSION", "SW_BODY_NAME", "SW_CALC_UT", "SW_LONGITUDE", "SW_POSITION_DETAIL", "SW_RUNTIME_STATUS", "SW_ENGINE_PATH", "SW_DATA_PATH"}.issubset(functions))

    def test_native_position_call_has_six_double_outputs_and_full_error_buffer(self):
        source = (REWORK / "src/vba/SWCalculation.bas").read_text()
        arrays = {
            name.lower(): (int(end) - int(start) + 1, kind.lower())
            for name, start, end, kind in re.findall(
                r"(\w+)\(\s*(\d+)\s+To\s+(\d+)\s*\)\s+As\s+(Byte|Double)", source, re.I
            )
        }
        self.assertEqual((6, "double"), arrays["values"])
        call = re.search(r"Native_swe_calc_ut\(([^\n]+)\)", source)
        self.assertIsNotNone(call)
        buffer = re.search(r",\s*(\w+)\(0\)\s*$", call.group(1))
        self.assertIsNotNone(buffer, "Native error argument is not an initialized Byte array")
        capacity, kind = arrays[buffer.group(1).lower()]
        self.assertEqual("byte", kind)
        self.assertGreaterEqual(capacity, 256)

    def test_failure_path_releases_calculation_guard_and_keeps_diagnostics_per_result(self):
        source = (REWORK / "src/vba/SWCalculation.bas").read_text()
        result_type = re.search(r"(?ims)^Public Type SWPositionResult\b.*?^End Type", source).group(0)
        for field in ("ActualFlags", "RequestedEphemeris", "ActualEphemeris", "Warning", "Status", "EnginePath", "EngineVersion", "DataPath"):
            self.assertRegex(result_type, rf"\b{field}\s+As\b")
        calculation = procedure(source, "SWCalculateUT")
        failure = calculation.split("Failed:", 1)[1]
        self.assertRegex(failure, r"(?i)If\s+acquired\s+Then\s+SWEndCalculation")
        self.assertRegex(failure, r'result\.Status\s*=\s*"ERROR"')
        self.assertIn("result.Warning", failure)
        code = without_vba_comments(source)
        self.assertNotRegex(code, r"(?im)^(?:Public|Private|Global)\s+(?:m)?(?:LastError|LastWarning|LastResult)\b")


if __name__ == "__main__":
    unittest.main()
