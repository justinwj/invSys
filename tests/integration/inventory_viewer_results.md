# Inventory Viewer Packaged Results

- Status: **PASS**
- Runtime: isolated generated test warehouse
- ConfigLoaded: True
- AuthLoaded: True
- TargetSelected: True
- TargetPathsSet: True
- SignedIn: True
- SnapshotCreated: True
- FirstActionRows: 3
- RepeatedLaunchReusedGeneration: True
- FilterVisibleRows: 1
- EventsVisibleRows: 2
- RefreshedEventsVisibleRows: 3
- NewestPublishedReference: BOL-VIEWER-NEW
- ReadableEventDates: True
- ViewerTabCount: 2
- ViewerTabCaptions: Inventory,Events
- SelectedViewerTab: Events
- RemoveEventsVisible: True
- EventsReadOnly: True
- SnapshotHashUnchanged: True
- NewPublicationChangedSnapshot: True

## Observed result

The public Operations Viewer action loaded readable Receipt and Shipping Remove events, refreshed the already-open Events page to show a newly published receipt first, reused one form generation, kept Events read-only, and left the new snapshot byte-for-byte unchanged.
