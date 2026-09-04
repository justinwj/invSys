# Plan 022 Slice 4x Designs Domain RED Results

- Date: 2026-08-23
- Focused VBA range: 264-267 of 301
- Passed: 0
- Failed: 4
- Harness: compiled and executed successfully in an isolated Excel workbook

| Test | Result | Expected missing behavior |
|---|---|---|
| `TestReusableProductionSchema_CreatesProcessRecipeProjections` | RED | Process, requirement, alternative, output, instruction, Recipe-node, and Recipe-connection projections do not exist. |
| `TestProcessSave_AppliesReusableMultiOutputDefinition` | RED | `PROCESS_SAVE` and its two-output projection application are unsupported. |
| `TestProcessSave_RejectsDefinitionWithoutOutput` | RED | The Domain does not yet return `PROCESS_OUTPUT_REQUIRED`. |
| `TestRecipeSave_RejectsCircularProcessGraph` | RED | The Domain does not yet return `RECIPE_CYCLE`. |

Command:

```powershell
& .\tools\run_phase6_excel_validation.ps1 -StartAt 264 -EndAt 267
```

The generated range report and harness workbook are ignored machine evidence;
this sanitized record contains no operational inventory, user, or credential
data.
