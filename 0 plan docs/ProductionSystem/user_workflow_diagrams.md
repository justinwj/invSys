# Production System - User Workflow Diagrams (Draft)

## 1) Recipe Builder + Inventory Palette Builder (user-facing)

```mermaid
flowchart TD
    classDef btn fill:#2f4e9c,stroke:#1f2f5c,color:#ffffff;
    classDef list fill:#2c7a9b,stroke:#195a73,color:#ffffff;
    classDef picker fill:#b27600,stroke:#6f4300,color:#ffffff;
    classDef note fill:#fff2cc,stroke:#b99a33,color:#000000;
    classDef legend fill:#f5f5f5,stroke:#333333,color:#000000;

    Toggle["Toggle Recipe Builder"]:::btn
    TogglePalette["Toggle Inventory Palette Builder"]:::btn
    LoadRecipe["Load Recipe (edit)"]:::btn
    SaveRecipe["Save Recipe"]:::btn
    SavePalette["Save IngredientPalette"]:::btn

    RecipeHeader["RB_AddRecipeName\nRECIPE_NAME | DESCRIPTION | GUID | RECIPE_ID"]:::list
    RecipeLines["RecipeBuilder (lines)\nPROCESS | DIAGRAM_ID | INPUT/OUTPUT | INGREDIENT | PERCENT | UOM | AMOUNT | OOO | INSTRUCTION | RECIPE_LIST_ROW | INGREDIENT_ID | GUID"]:::list
    AddProc["Add Recipe Process Table"]:::btn
    ProcTables["proc_*_rbuilder\nPROCESS | DIAGRAM_ID | INPUT/OUTPUT | INGREDIENT | PERCENT | UOM | AMOUNT | OOO | INSTRUCTION | RECIPE_LIST_ROW | INGREDIENT_ID | GUID"]:::list

    TemplatesTable["TemplatesTable\n(GUID | TEMPLATE_SCOPE | RECIPE_ID | INGREDIENT_ID | PROCESS | TARGET_TABLE | TARGET_COLUMN | FORMULA | NOTES | ACTIVE | CREATED_AT | UPDATED_AT)"]:::list
    SaveFormulas["Save Formulas\n(per recipe)"]:::btn

    PalettePicker["Recipe picker (IP_ChooseRecipe)\nopens Recipes table list"]:::picker
    IngredientPicker["Ingredient picker\nshows INGREDIENTs for RECIPE_ID"]:::picker
    PaletteHeader["IP_ChooseRecipe\nRECIPE_NAME | DESCRIPTION | GUID | RECIPE_ID"]:::list
    PaletteLines["IP_ChooseIngredient\nINGREDIENT | UOM | QUANTITY | DESCRIPTION | GUID | RECIPE_ID | INGREDIENT_ID | PROCESS"]:::list
    PaletteItems["IP_ChooseItem\nITEMS | UOM | DESCRIPTION | ROW | RECIPE_ID | INGREDIENT_ID"]:::list
    IngredientPalette["IngredientPalette (sheet table)\nRECIPE_ID | INGREDIENT_ID | INPUT/OUTPUT | ITEM | PERCENT | UOM | AMOUNT | ROW | GUID"]:::list

    ItemPicker["Item search picker\nshows invSys items\n(+ add CATEGORY filters)"]:::picker

    Toggle --> RecipeHeader
    TogglePalette --> PaletteHeader
    LoadRecipe --> RecipeHeader
    RecipeHeader --> RecipeLines --> AddProc --> ProcTables
    ProcTables --> SaveRecipe
    ProcTables --> SaveFormulas --> TemplatesTable

    PalettePicker --> PaletteHeader
    PaletteHeader --> IngredientPicker --> PaletteLines
    PaletteLines --> ItemPicker --> PaletteItems
    PaletteItems --> SavePalette --> IngredientPalette
    PaletteItems --> SaveFormulas

    Note1["Ingredient picker pulls from Recipes sheet.\nItem picker starts unfiltered; user can add CATEGORY filters (+)."]:::note
    Note2["TemplatesTable stores FormulaR1C1 per recipe + scope.\nSave Formulas scans proc_*_rbuilder + IP_Choose* columns."]:::note
    ItemPicker -.-> Note1
    TemplatesTable -.-> Note2

    subgraph Legend
        L1["Button"]:::btn
        L2["ListObject / Table"]:::list
        L3["Picker / Search"]:::picker
        L4["Note"]:::note
        L5["Legend"]:::legend
    end
```

## 3) Recipe Chooser (System 3: user-facing)

```mermaid
flowchart TD
    classDef btn fill:#2f4e9c,stroke:#1f2f5c,color:#ffffff;
    classDef list fill:#2c7a9b,stroke:#195a73,color:#ffffff;
    classDef picker fill:#b27600,stroke:#6f4300,color:#ffffff;
    classDef note fill:#fff2cc,stroke:#b99a33,color:#000000;
    classDef chk fill:#2e7d32,stroke:#1b4f1f,color:#ffffff;
    classDef legend fill:#f5f5f5,stroke:#333333,color:#000000;

    RCHeader["RC_RecipeChoose (header)\nRECIPE | RECIPE_ID | DEPARTMENT | DESCRIPTION | PREDICTED OUTPUT | PROCESS"]:::list
    RecipePicker["Recipe picker\nopens Recipes table list"]:::picker
    BuildProcTables["Build process tables per PROCESS\n(RecipeChooser_generated)\nPROCESS | DIAGRAM_ID | INPUT/OUTPUT | INGREDIENT | PERCENT | UOM | AMOUNT NEEDED | INGREDIENT_ID | RECIPE_LIST_ROW"]:::list
    ProcessCheckboxes["Process selector checkboxes\n(checked = build palette tables)"]:::chk

    UsedIngredientScan["Scan process tables\n(INPUT/OUTPUT = USED only)"]:::note
    BuildPalettes["Build InventoryPalette tables\nper USED ingredient\n(proc_*_palette or InventoryPalette_generated)\nITEM_CODE | VENDORS | VENDOR_CODE | DESCRIPTION | ITEM | UOM | QUANTITY | PROCESS | LOCATION | ROW | INPUT/OUTPUT"]:::list

    PaletteFilter["Filter choices by IngredientPalette\n(RECIPE_ID + INGREDIENT_ID)\n(acceptable invSys items)"]:::note
    ItemPicker["Item search picker\ninvSys items filtered by palette\n+ optional CATEGORY filters"]:::picker

    RCHeader --> RecipePicker --> BuildProcTables
    BuildProcTables --> ProcessCheckboxes --> BuildPalettes
    BuildProcTables --> UsedIngredientScan --> BuildPalettes
    BuildPalettes --> PaletteFilter --> ItemPicker

    Note1["User picks a recipe once; process tables paste in.\nPalette tables are generated per USED ingredient in each PROCESS."]:::note
    BuildPalettes -.-> Note1

    subgraph Legend
        L1["Button"]:::btn
        L2["ListObject / Table"]:::list
        L3["Picker / Search"]:::picker
        L4["Checkbox"]:::chk
        L5["Note"]:::note
        L6["Legend"]:::legend
    end
```

## 2) Production Run (Recipe chooser → outputs → logs)

```mermaid
flowchart TD
    classDef btn fill:#2f4e9c,stroke:#1f2f5c,color:#ffffff;
    classDef list fill:#2c7a9b,stroke:#195a73,color:#ffffff;
    classDef picker fill:#b27600,stroke:#6f4300,color:#ffffff;
    classDef log fill:#8c6239,stroke:#4f341f,color:#ffffff;
    classDef note fill:#fff2cc,stroke:#b99a33,color:#000000;
    classDef legend fill:#f5f5f5,stroke:#333333,color:#000000;
    classDef chk fill:#2e7d32,stroke:#1b4f1f,color:#ffffff;

    ToggleProd["Toggle Production Run"]:::btn
    RecipeChooser["Recipe chooser table\nRECIPE | RECIPE_ID | DEPARTMENT | DESCRIPTION | PREDICTED OUTPUT | PROCESS"]:::list
    ProcessSelector["Process selector\n(checkbox list beside process tables)\n(checked = generate palette table per process)"]:::chk
    BatchCodesToggle["Batch codes required?\n(checkbox per PROCESS/BATCH)"]:::chk
    RecipePicker["Recipe picker\nopens Recipes table list"]:::picker
    BuildProcessTable["Generate Process Table\n(from Recipes + TemplatesTable)"]:::list

    ItemPicker2["Item search picker\ninvSys items filtered by\nIngredientPalette (RECIPE_ID+INGREDIENT_ID)\n+ optional CATEGORY filters"]:::picker
    InventoryInputs["Inventory item chooser (inputs)\n(per process, generated next to process table)\nITEM_CODE | VENDORS | VENDOR_CODE | DESCRIPTION | ITEM | UOM | QUANTITY | PROCESS | LOCATION | ROW | INPUT/OUTPUT"]:::list
    KeepInventory["Keep inventory selection\n(per PROCESS)\n(checked = keep on Next Batch)"]:::chk
    ProductionOutputs["ProductionOutput\nROW | PROCESS | OUTPUT | UOM | REAL OUTPUT | BATCH | RECALL CODE\n(recall checkbox per row)"]:::list

    BatchCodesLog["BatchCodesLog\nRECIPE | RECIPE_ID | PROCESS | OUTPUT | UOM | REAL OUTPUT | BATCH | RECALL CODE | TIMESTAMP | LOCATION | USER | GUID"]:::log
    PrintCodes["Print recall codes\n(format batch code sheet)"]:::btn
    CheckInv["Prod_invSys_Check\nUSED | MADE | TOTAL INV | ROW"]:::list
    ToUsed["To USED"]:::btn
    SendTotal["Send to TOTAL INV"]:::btn
    ToMade["Send to MADE"]:::btn
    ProductionLog["ProductionLog"]:::log

    NextBatch["Next Batch\n(clear Inventory item chooser)"]:::btn
    ClearProd["Clear Production sheet\n(when production complete)"]:::btn

    ToggleProd --> RecipeChooser --> RecipePicker --> BuildProcessTable
    BuildProcessTable --> ProcessSelector
    ProcessSelector --> ItemPicker2 --> InventoryInputs
    InventoryInputs --> ExpandRows["Auto-expand rows\n(push tables below down)\n(no left/right edits)"]:::note
    InvChooserCF["Conditional formatting\nQUANTITY green=exact red=over/under"]:::note
    InventoryInputs -.-> InvChooserCF
    InventoryInputs -.-> KeepInventory
    InventoryInputs --> ProductionOutputs
    ProductionOutputs --> BatchCodesToggle --> PrintCodes
    ProductionOutputs --> CheckInv

    InventoryInputs --> ToUsed --> CheckInv
    CheckInv --> ProductionOutputs
    ProductionOutputs --> ToMade --> CheckInv
    ToMade --> BatchCodesLog
    ToMade --> NextBatch
    CheckInv --> SendTotal --> CheckInv

    ToUsed --> ProductionLog
    ToMade --> ProductionLog
    SendTotal --> ProductionLog

    SendTotal --> ClearProd

    Note2["Process table formulas come from TemplatesTable.\nQuantities edited here drive outputs and inventory moves."]:::note
    Note4["Inventory item chooser highlights QUANTITY:\nGreen = matches required, Red = over/under."]:::note
    Note6["TOTAL INV in Prod_invSys_Check highlights red when insufficient."]:::note
    CheckInv -.-> Note6
    Note7["Batch = user-defined run.\nRecall code generated per PROCESS/BATCH when sent to MADE."]:::note
    BatchCodesLog -.-> Note7
    Note5["Process selector checkboxes control which\nInventory Palette tables are generated.\nIf none checked, no palette tables are created."]:::note
    ProcessSelector -.-> Note5
    BuildProcessTable -.-> Note2
    Note3["All PROCESS tables are generated on recipe load (Release 1).\nUser edits batch size / quantities in generated tables.\nLayout supports fixed bands or dynamic bands (test both)."]:::note
    BuildProcessTable -.-> Note3

    subgraph Legend
        L1["Button"]:::btn
        L2["ListObject / Table"]:::list
        L3["Picker / Search"]:::picker
        L7["Checkbox"]:::chk
        L4["Log"]:::log
        L5["Note"]:::note
        L6["Legend"]:::legend
    end
```

## 4) Formula Management (VBA-level draft)

```mermaid
flowchart TD
    classDef btn fill:#2f4e9c,stroke:#1f2f5c,color:#ffffff;
    classDef op fill:#2c7a9b,stroke:#195a73,color:#ffffff;
    classDef data fill:#8c6239,stroke:#4f341f,color:#ffffff;
    classDef event fill:#6a1b9a,stroke:#3f0f5c,color:#ffffff;
    classDef note fill:#fff2cc,stroke:#b99a33,color:#000000;

    SaveAll["Save Formulas\n(all systems per recipe)"]:::btn

    ScanRecipe["Scan recipe process tables\n(proc_*_rbuilder)"]:::op
    ScanPalette["Scan palette builder tables\n(IP_ChooseIngredient / IP_ChooseItem)"]:::op
    ScanProd["Scan production tables\n(proc_*_palette / ProductionOutput)"]:::op

    UpsertTemplates["Upsert TemplatesTable rows\n(RECIPE_ID + TEMPLATE_SCOPE + PROCESS + TARGET_TABLE + TARGET_COLUMN)"]:::op
    TemplatesTable["TemplatesTable"]:::data

    RecipeChosen["Recipe selected\n(RB_AddRecipeName / IP_ChooseRecipe / RC_RecipeChoose)"]:::event
    EnsureTables["Ensure tables exist\n(build proc tables, palettes, outputs)"]:::op
    ApplyTemplates["Apply templates by scope + process\n(cTemplateApplier)"]:::op
    ClearUnchecked["If process unchecked or table cleared:\nremove formulas in target columns"]:::op

    SaveAll --> ResolveRecipe --> ScanRecipe --> UpsertTemplates --> TemplatesTable
    ResolveRecipe --> ScanPalette --> UpsertTemplates
    ResolveRecipe --> ScanProd --> UpsertTemplates

    RecipeChosen --> EnsureTables --> ApplyTemplates --> ClearUnchecked
    TemplatesTable -.-> ApplyTemplates

    Note1["Template scopes (draft):\nRECIPE_PROCESS = builder + chooser tables\nPALETTE_BUILDER = IP_Choose* tables\nPROD_RUN = proc_*_palette + ProductionOutput"]:::note
    Note1 -.-> UpsertTemplates

    Note2["TemplatesTable.FORMULA stores FormulaR1C1"]:::note
    Note2 -.-> TemplatesTable
```
