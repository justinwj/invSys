# Plan 022 Slice 4o Packaged Results

- Passed: 5
- Failed: 0

| Check | Result | Detail |
|---|---|---|
| Packaged.ReceivingSurface | PASS | Generated operator surface accepts the expanded Receiving schema. |
| Packaged.ReturnsTabContract | PASS | OK\|Selected=Returns\|AddCaption=Add Return\|ConditionVisible=True\|ReturnReasonVisible=True\|ReceiptEventType=RECEIVE |
| Packaged.InboundReturnFormAction | PASS | OK\|StagedRows=1\|ReceiptType=RETURN\|Condition=BAD\|Reason=TEST RETURN |
| Packaged.ReturnProjection | PASS | Staging and aggregate retain return type, condition, and reason. |
| Packaged.DemoInventorySilentClose | PASS | OK\|Seed=True\|DeleteInventory=True\|UploadDataSet=True\|DeleteDataSet=True\|R1Protected=True\|Cancel=False\|CloseIsSilent=True |
