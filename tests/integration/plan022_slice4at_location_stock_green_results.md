# Plan 022 Slice 4at Location Stock GREEN Results

**Date:** 2026-08-30

## Contract result

- Focused source contract: RED `0/7`, GREEN `7/7`.
- Packaged reusable Production actions/run/restart: `2/2` GREEN.
- Packaged Receiving durability: `1/1` GREEN.

The Production action used the same **Apply** handler as the operator. Its
authoritative fixture contained seven exact receipt entities for one
SKU/UOM/location/condition stock bucket. The run displayed one stock choice,
expanded a 5-LB allocation over multiple exact `System_Key` entities, retained
exact-key Check In and completion, completed two batches, and passed clean
restart. Receiving exposed `CapacityStub=True`; the column is blank and inert.

## Regression result

- Adjacent Production source contracts: `6/6`, `7/7`, `8/8`, and `6/6` GREEN.
- Receiving stabilization: `10/10` GREEN.
- Workflow readiness: `18/18` GREEN.
- Launcher contracts: `24/24` GREEN.
- Packaged XLAM: `81/81` GREEN.
- Ribbon/VBA compile: `142/142` GREEN.
- Live roles: `47/47` GREEN.
- Ordered Release 1 chain: `30/30` GREEN.
- Deterministic static baseline: `19/19` GREEN.
- Reviewed cleanup/growth: `13/13` GREEN.
- Static metrics: 154 components, 5,223 procedures, 1,050 scanner candidates.

The dedicated-NAS rerun did not start because the configured test root was
unavailable. The prior `16/16` result remains the last verified NAS evidence;
it is not claimed as current Slice 4at proof.

## Remaining gate

Visible operator confirmation is required for one Cassia Oil/CLEARVIEW stock
row, the readable Ingredients Assignment `System_Key`, and the blank Receiving
**Capacity (coming later)** column.
