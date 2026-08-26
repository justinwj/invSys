# Plan 022 Slice 4af Recipe identity GREEN

Date: 2026-08-26

Recipe Designer now initializes a blank draft with the next collision-checked
three-character Base-36 Recipe ID and proposed version `1` when Production
opens, on **New Recipe**, and on **Clear**. Save Draft, Validate, and Release
defensively supply either missing generated value before graph validation.
Recipe ID remains locked; Recipe Version is editable and must be a positive
whole number.

| Gate | Result |
|---|---:|
| Focused Slice 4af source | 6/6 |
| Packaged Production public callback + clean restart | 2/2 |
| Historical Slice 4y through 4ae focused source | GREEN |
| Packaged XLAM | 74/74 |
| Ribbon/compile | 142/142 |
| Live role workflows | 47/47 |
| Ordered Release 1 chain | 30/30 |
| Launcher contracts | 24/24 |
| Dedicated NAS runtime | 16/16 |
| Deterministic static baseline | 19/19 |
| Reviewed cleanup/growth | 13/13 |

The packaged form report records `RecipeIdentityInitialized=True`,
`RecipeIdGenerated=True`, `RecipeVersionGenerated=True`,
`RecipeIdLocked=True`, and `RecipeVersionEditable=True`. The actual New Recipe,
Save Draft, and Release handlers change the proposed Version from `1` to `9`
and record `EditedRecipeVersionRetained=True`. Static evidence is 153
components, 5,081 procedures, and 1,038 candidates. Visible acceptance in the
user's saved Production workbook remains pending.
