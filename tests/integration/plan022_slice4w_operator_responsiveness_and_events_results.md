# Plan 022 Slice 4w Operator Responsiveness and Events Results

- Passed: 11
- Failed: 0

| Check | Result | Contract |
|---|---|---|
| ServerConnection.ProgressBeforeBlockingIO | PASS | Manual and ribbon Server Sign In render progress before the synchronous Windows SMB call and restore Excel UI afterward. |
| Receiving.AggregateReferenceDetail | PASS | The fixed-height aggregate list retains one-line rows while a dedicated multiline detail surface shows every concatenated reference and clears with staging. |
| Viewer.Events.ReadOnlyTab | PASS | Inventory Viewer exposes a read-only Events tab covering canonical inventory events plus current box-design and held-shipment activity. |
| Viewer.Tabs.ExactlyInventoryAndEvents | PASS | The runtime Viewer does not append duplicate placeholder pages, exposes exactly Inventory and Events, and its public Events action selects the operator-visible Events tab. |
| Viewer.Layout.GuardsNativeWindowState | PASS | Operations anchoring skips native form-size enforcement while minimized or maximized and contains residual run-time error 384 without disabling restored-state layout. |
| Viewer.Events.ReadableTimestampRefresh | PASS | Events renders readable timestamps and the public Events refresh reports the newly published first event rather than retaining stale rows. |
| Viewer.Events.RollingDateFilters | PASS | The operator-visible Events Refresh action combines text search with All, rolling Day/Week/Month, or a typed positive whole-number-of-days filter; Inventory remains unfiltered. |
| Viewer.Events.RemembersDateFilter | PASS | A valid applied Event range is stored as a per-Windows-user Operations preference and restored when the public Viewer action creates a new form instance. |
| Viewer.Events.PublishedProjection | PASS | Viewer event history is read from the published snapshot projection rather than making the form a canonical writer or authority. |
| Viewer.Events.RemoveRelease | PASS | Shipping Remove releases locked inventory through SHIP_RELEASE and the operator-facing Events view labels that event Remove. |
| OperatorPersistence.PendingStatus | PASS | Receiving/Returns and Shipping render their own saving-to-server status before required persistence begins; Office-native progress UI remains separate. |
