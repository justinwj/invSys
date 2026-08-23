# Plan 022 Slice 4x Designs Lifecycle/Graph RED Results

- Date: 2026-08-23
- Focused VBA range: 268-275 of 309
- Passed: 2
- Failed: 6
- Harness: compiled and executed successfully in an isolated Excel workbook

| Test | Result | Expected missing behavior |
|---|---|---|
| `TestProcessLifecycle_ReleasesObsoletesAndReusesVersions` | PASS | Existing lifecycle event replay already preserves immutable version reuse. |
| `TestRecipeRelease_RejectsMissingOrUnreleasedProcessVersion` | RED | Recipe release does not validate pinned Process release status. |
| `TestRecipeRelease_RejectsUnresolvedExternalRequirement` | RED | Recipe release does not reject a requirement with no upstream connection or acceptable alternative. |
| `TestRecipeRelease_RejectsIncompatibleConnection` | RED | Recipe release does not validate output/requirement UOM and item compatibility. |
| `TestRecipeRelease_RejectsOutputOverallocation` | RED | Recipe release does not reject routed quantity above output yield. |
| `TestRecipeRelease_RejectsContradictoryExecutionOrder` | RED | Recipe release does not reject an operator order that contradicts the graph. |
| `TestProcessObsolete_RejectsReleasedRecipeDependency` | RED | Process obsolete does not protect versions referenced by a released Recipe. |
| `TestRecipeLifecycle_ReleasesValidGraphAndThenObsoletes` | PASS | Existing event replay supports the valid lifecycle sequence. |

Command:

```powershell
& .\tools\run_phase6_excel_validation.ps1 -StartAt 268 -EndAt 275
```

The generated range report and harness workbook are ignored machine evidence;
this sanitized record contains no operational inventory, user, or credential
data.
