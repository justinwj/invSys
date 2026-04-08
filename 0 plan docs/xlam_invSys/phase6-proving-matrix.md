# Phase 6 Proving Matrix

**Authority:** `invSys-Design-v4.7.md` lines 700–707, 1047, 1078  
**Purpose:** Re-anchor acceptance criteria to the four-rung proving ladder defined in the main design doc. Prevents any single-warehouse flow from being counted as Phase 6 completion.

---

## The Four Required Gates (from design doc §Phase 6)

| Gate | Label | Required Before |
|------|-------|-----------------|
| G1 | One-account local use | G2 |
| G2 | LAN use — multi-PC, one warehouse | G3 |
| G3 | LAN + WAN — multi-PC, two warehouses | G4 |
| G4 | Central aggregation — HQ view, published artifacts only | — |

No gate may be claimed on the basis of a lower gate's evidence.

---

## Current Evidence Map

### G1 — One-Account Local (PROVEN)

| Artifact | Status | Notes |
|----------|--------|-------|
| `tests/integration/test_ConfirmWrites_Tester.bas` | ✅ Evidence | Valid for single-account, saved-workbook scope only |
| `tests/integration/confirm-writes-results.md` | ✅ Evidence | Scoped to G1; not WAN acceptance |
| `docs/tester-distribution-contract.md` | ⚠️ Misframed | Currently written as if it covers all slices; scope must be narrowed to G1 |
| `docs/tester-guide-confirm-writes.md` | ⚠️ Misframed | Should be relabeled as a G1 quickstart, not Phase 6 tester guide |

**G1 verdict: PASS — but must not be promoted as G2/G3/G4 evidence.**

---

### G2 — LAN Multi-PC, One Warehouse (INCOMPLETE)

Required evidence bundle — none of these are fully on file yet:

- [ ] Package/load-path identity verified on each physical machine
- [ ] Shared runtime folder reachability confirmed from both stations
- [ ] Inbox routing confirmed across the LAN boundary
- [ ] Processor lock acquired, clean denial, and retry confirmed on second station
- [ ] Two stations converge to identical warehouse totals after concurrent edits

Existing partial coverage to audit:

| Artifact | Status |
|----------|--------|
| `0 plan docs/xlam_invSys/Phase-6.user-side-test-guide.md` | Partial — review for G2 applicability |
| `tests/unit/phase6_lan_wan_proving_results.md` | Partial — tag each row real-machine vs simulated |

**G2 verdict: OPEN — LAN gate must close before any WAN claim is valid.**

---

### G3 — LAN + WAN, Two Warehouses (INCOMPLETE)

**WH1 operator proof:**
- [ ] WH1 processes locally and publishes with delay
- [ ] Stale/cached operator refresh behavior confirmed during WAN delay
- [ ] Catch-up publish after connectivity returns

**WH2 operator proof:**
- [ ] WH2 processes locally and publishes with delay (independent of WH1)
- [ ] Same stale/refresh/catch-up behavior confirmed for WH2

**Cross-warehouse:**
- [ ] WH1 and WH2 publish artifacts are independently valid
- [ ] No cross-contamination of warehouse totals

Existing partial coverage to audit:

| Artifact | Status |
|----------|--------|
| `tests/integration/wan-smoke-results.md` | Partial — confirm whether scope is WH1-only or both |
| `0 plan docs/xlam_invSys/Phase-6.user-side-test-guide.md` | Review for G3 sections |

**G3 verdict: OPEN — both warehouse operator paths must be independently proven.**

---

### G4 — Central Aggregation / HQ View (BLOCKED)

Blocked on G3. When G3 is green, the HQ proof requires:

- [ ] Published WH1 artifact only (no live WH1 connection)
- [ ] Published WH2 artifact only (no live WH2 connection)
- [ ] Staggered republish: HQ re-reads after WH1 republishes, then after WH2 republishes
- [ ] Advisory global snapshot preserves per-warehouse rows
- [ ] Global snapshot confirmed non-authoritative (no write-back path)

**G4 verdict: BLOCKED — do not attempt until G2 and G3 are closed.**

---

## Artifact Reclassification

| Artifact | Was Framed As | Correct Scope |
|----------|---------------|---------------|
| `docs/tester-distribution-contract.md` | All-slice acceptance target | G1 only — narrow to single-account/saved-workbook |
| `docs/tester-guide-confirm-writes.md` | Phase 6 tester guide | G1 quickstart — relabel or replace |
| `tests/integration/confirm-writes-results.md` | WAN acceptance evidence | G1 evidence only |
| `test_ConfirmWrites_Tester.bas` | General Phase 6 tester | G1 tool — valid, keep, label scope |

---

## Sequenced Next Steps

1. **Now:** Narrow `tester-distribution-contract.md` scope to G1; add header disclaimer.
2. **Now:** Relabel `tester-guide-confirm-writes.md` as G1 quickstart.
3. **Next:** Audit `phase6_lan_wan_proving_results.md` and `wan-smoke-results.md` — tag each row G1/G2/G3.
4. **Then:** Run real-machine LAN bundle (G2 checklist above) on two physical stations.
5. **Then:** Run two-warehouse WAN operator proof (G3 checklist above).
6. **Finally:** Run HQ aggregation proof (G4 checklist above) only after G2 and G3 are green.

---

## Gate Closure Criteria

A gate is closed when:
- Every checklist item above has a corresponding result row in a dated results file.
- The results file name includes the gate label (e.g. `phase6-g2-lan-results.md`).
- No checklist item is marked "simulated" or "assumed" — all must be real-machine runs.

---

*Last updated: 2026-04-07. Authoritative source: `invSys-Design-v4.7.md` §Phase 6.*
