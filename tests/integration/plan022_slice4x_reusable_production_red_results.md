# Plan 022 Slice 4x Reusable Production RED Results

- Passed: 5
- Failed: 5

| Check | Result | Contract |
|---|---|---|
| Production.PublicLauncherPreserved | PASS | Slice 4x preserves the packaged Production launcher, PROD_POST gate, and captured-workbook form boundary. |
| Production.Form.FiveTargetPages | PASS | The public Production form exposes Process Designer, Recipe Designer, Ingredients Assignment, Run List, and experimental Run Tree, with no Recipe Builder page. |
| Production.ProcessDesigner.LifecycleHandlers | FAIL | Process Designer uses operator handlers for new/reuse, draft save, release, and obsolete lifecycle events. |
| Production.ProcessDesigner.RequiresOutput | PASS | The form and Designs Domain both reject a Process definition with no output. |
| Production.RecipeDesigner.GraphHandlers | FAIL | Recipe Designer and Designs Domain validate connections, execution order, unresolved inputs, quantities, and circular dependencies. |
| Production.IngredientsAssignment.ProcessRequirements | FAIL | Ingredients Assignment maps each exact Process-version requirement to acceptable managed item/SKU alternatives. |
| Production.RunSession.MultiOutput | FAIL | The typed run session carries correlated Process executions and multiple output allocations rather than one singular output key. |
| Production.Completion.ExactKeysAndCoProducts | FAIL | Completion serializes every fresh output key, exact routed-intermediate input keys, and finished/co-product balances. |
| DesignsDomain.ProcessRecipeSchemaAndEvents | PASS | The headless Designs Domain owns reusable Process/Recipe lifecycle events and rebuildable projections. |
| Viewer.ProductionOperatorEvents | PASS | Published Viewer Events expose correlated Production input consumption and output creation as operator actions. |
