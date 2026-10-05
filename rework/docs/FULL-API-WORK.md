# Full API implementation

Authorized on 2026-10-05 after the Windows integration checkpoint. The earlier pre-Parallels stop in PLAN.md is superseded. Public-release/download onboarding is explicitly deferred.

| Work package | Completion evidence |
| --- | --- |
| API-01: retry full-project compile | Passed on dev.3 and dev.4 through the authorized Excel Compile command; VBA-project access restored. |
| API-02: review all 106 signatures and contracts | Reviewed against pinned C source: buffer capacities, ownership, units, status and state. `api/contracts.json` and `docs/API-REFERENCE.md`. |
| API-03: safe interfaces | 95 worksheet functions, 11 VBA commands, UTC/house/date-series helpers. |
| API-04: examples and numerical fixtures | All 95 worksheet examples, command recipes, guided Recipes sheet, independent Windows C caller. |
| API-05: Windows verification | Explicit compile, 17-module source parity, 106 native comparisons, 136 API/example regressions, 18 smoke and 25 desktop checks. |
| API-06: version and delivery | 0.1.0-dev.4; documentation, delivered-package attestation and source commit. See verification/2026-10-06. |

Declaration coverage, interface coverage, actual execution and numerical agreement remain distinct evidence counts. Public-release acceptance is outside this work package.
