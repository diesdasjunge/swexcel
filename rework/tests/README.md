# Repository verification

Run from the repository root:

```sh
python3 -m unittest discover -s rework/tests -v
```

The tests use only the Python standard library. They independently parse the
Windows PE export table and pinned public C prototypes, compare the native VBA
ABI and catalogue, preserve original file hashes, verify vendor/source closure
and prepared-package integrity, and check selected runtime source regressions.
They do not compile VBA or execute Excel or the Windows DLL.

Generated-package verification skips until `tools/prepare.py` has produced a
distribution. Fresh-build verification skips until the Windows source compiler
has written `runtime/engine/build-provenance.json`. A fresh compiled DLL replaces
the prepared reference DLL only through that explicit provenance record; the
three binary ephemerides and three supporting catalogues remain hash-pinned.

The prepared package includes this harness for inspection. The complete suite
expects the repository layout, including preserved originals in `ephem/`,
`vba/`, and the original `.xls` workbook. Run it against the repository checkout,
not a standalone extracted runtime package. Actual Windows compilation,
numerical comparisons and Excel scalar/spill acceptance follow
`docs/WINDOWS-TESTING.md` and remain separate gates.
