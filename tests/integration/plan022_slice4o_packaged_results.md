# Plan 022 Slice 4o Packaged Results

- Passed: 5
- Failed: 0

| Check | Result | Detail |
|---|---|---|
| Packaged.ReceivingSurface | PASS | Generated operator surface accepts the expanded Receiving schema. |
| Packaged.ReturnsTabContract | PASS | OK\|Selected=Returns\|AddCaption=Add Return\|ConditionVisible=True\|ReturnReasonVisible=True\|HistoryTitle=Return Entries History\|TallyTitle=Return Tally\|AggregateTitle=Aggregate Returns\|ItemConditionColumn=True\|ReceiptEventType=RECEIVE |
| Packaged.InboundReturnFormAction | PASS | OK\|StagedRows=1\|ReceiptType=RETURN\|Condition=BAD\|Reason=TEST RETURN |
| Packaged.ReturnProjection | PASS | Rebuilt=True; Rows=1; Ref=RETURN-TEST, RETURN-SECOND; Qty=3; Condition=BAD; Reason=TEST RETURN |
| Packaged.DemoInventorySilentClose | PASS | OK\|Seed=True\|DeleteInventory=True\|UploadDataSet=True\|DeleteDataSet=True\|R1Protected=True\|Cancel=False\|CloseIsSilent=True |
