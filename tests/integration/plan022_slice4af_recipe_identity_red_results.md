# Plan 022 Slice 4af Recipe identity RED

Date: 2026-08-26

The focused source contract was added before implementation and reported
`0/6`. Recipe Designer currently leaves Recipe ID and version blank when the
Production form initializes, locks both fields, and validates Save Draft and
Release before supplying missing identity values. Packaged evidence also still
asserts the superseded all-identity-controls-locked contract and does not prove
that an operator-edited Recipe version survives the actual Save/Release
handlers.

| Focused check | RED reason |
|---|---|
| Docs.RecipeIdentityContract | Architecture, Plan 022, and controls catalog still require a locked Recipe version. |
| Form.IdentityControlState | `txtReusableRecipeVersion` is locked. |
| Form.AutomaticIdentityHelper | No shared Recipe identity initializer exists. |
| Form.InitialAndHandlerPaths | Initial load and Save/Release do not guarantee Recipe identity first. |
| Packaged.OperatorEvidence | Public form evidence lacks generated/editable/retained-version assertions. |
| Validator.RequiresRecipeEvidence | Packaged validator does not require the revised contract. |

This is meaningful behavioral RED; the workbook and harness are available and
the failure describes the operator-visible defect.
