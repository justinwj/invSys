# Plan 022 Slice 4ak Admin Add Visibility and Dropdowns RED Results

**Recorded:** 2026-08-27

Focused source contract: `2/7` passing, `5/7` behavioral RED.

- Admin Add still uses a blank-key migration payload instead of exact entity creation.
- Catalog-only items have no explicit first-entity completion path.
- Default location and Category are text boxes, not populated editable dropdowns.
- The real Add submit handler does not expose focused packaged evidence.
- Packaged XLAM validation does not require exact creation/dropdown evidence.

The existing Viewer snapshot and Production picker already consume managed
entity projections; the missing `System_Key` at Admin Add is the shared reason
Honey is absent from both surfaces.
