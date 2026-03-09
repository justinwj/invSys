Yeah — at this point you’re running into a **known, recurring failure mode** of `chatgpt.com`: sandbox download links randomly expire or never resolve, especially after multiple regenerations. Retrying the same mechanism won’t fix it.

So here’s the **reliable workaround** that does *not* depend on downloads at all.

---

## ✅ Guaranteed-working solution (no links, no sandbox)

Below is the **entire document**, **verbatim**, wrapped in a single Markdown code block.

You can:

* copy-paste it into a `.md` file locally
* paste it directly into GitHub
* paste it into VS Code / Obsidian / Typora
* or re-import it into another ChatGPT canvas

This avoids the broken file-handoff completely.

---

````md
# invSys Architecture v2 — Consolidated & Corrected (Single Source of Truth)

## Purpose

This document **fully supersedes all prior invSys architecture notes and diagrams**. It resolves earlier contradictions, removes abandoned alternatives, and defines the **single authoritative working architecture** for the system.

This document remains a living design, but it is the **only source of architectural truth**.

---

## Architectural Decisions

### D1 — One write model everywhere: Inbox → Processor → Outbox
All domain state changes occur by:
1. Appending an **Inbox event**
2. Applying it via a **single-writer Processor**
3. Emitting a durable **Outbox event** after successful apply

There are **no direct writes** to authoritative data stores.

---

### D2 — Multi-warehouse, LAN-first, WAN-optional
- Each warehouse operates independently on its local LAN
- Internet/WAN is optional and intermittent
- SharePoint is a **convenience sync and distribution layer**, not a source of truth

Each warehouse owns:
- Inventory
- (Optionally) Designs / BOMs
- Users
- Locks
- Processing

---

### D3 — Clear ownership boundary

- **Core XLAM**
  - Authorization gate
  - Orchestration
  - Config loading
  - Processor runner
  - Lock coordination

- **Domain XLAMs (Inventory, Designs)**
  - Own schemas
  - Own authoritative `.xlsb` stores
  - Perform all writes and invariant enforcement

- **Role XLAMs (Receiving, Shipping, Production)**
  - UI only
  - Create Inbox events
  - Never mutate domain state

- **Admin XLAM**
  - Orchestration console
  - Diagnostics, recovery, configuration
  - Never writes domain tables directly

---

### D4 — Forms strategy
Each add-in that requires dynamic search or admin forms embeds its **own copy**:
- `ufDynItemSearchTemplate`
- `ufDynDesignSearchTemplate`
- `ufDynAdminTemplate`

Shared Core forms are explicitly **not a contract**.

---

## Authoritative Data Model

### Authoritative stores (per warehouse)
- `WHx.invSys.Data.Inventory.xlsb`
- `WHx.invSys.Data.Designs.xlsb` (optional)

These files are:
- Single-writer
- Locally authoritative
- Never written by UI code

---

## Locking Model

### Lock placement (decision)

Locks live **inside the authoritative warehouse store for the domain being mutated**.

- Inventory locks → `WHx.invSys.Data.Inventory.xlsb`
- Designs locks → `WHx.invSys.Data.Designs.xlsb`

There is **no separate Locks workbook**.

Locks and data share the **same durability boundary**.

---

### Lock lifecycle

```mermaid
flowchart TB
  Start[Processor starts]
  Open[Open authoritative store]
  Check{Active lock exists?}
  Acquire[Insert lock row]
  Apply[Apply domain changes]
  Heartbeat[Update heartbeat]
  Release[Mark lock released]
  End[Processor ends]

  Start --> Open --> Check
  Check -- No --> Acquire --> Apply --> Release --> End
  Check -- Yes --> Wait[Retry / backoff]
  Apply --> Heartbeat --> Apply
````

---

## Identity Resolution

### Chosen development path (Release 1)

**Option A — Workbook-stored users + PIN**

Users are:

* Defined inside invSys data
* Authenticated locally via PIN (or short passcode)
* Resolved entirely offline

There is **no Windows identity, Active Directory, or WAN dependency** in Release 1.

### Identity invariants

* A single `UserID` is authoritative
* The same identity is used for:

  * Capability checks
  * Lock ownership
  * Audit logging
* Identity authority travels with the warehouse

Windows / AD identity may be added later **only as a convenience layer**, never as the authority.

---

## Event Flow

### High-level workflow

```mermaid
flowchart TB
  User[User] --> Role[Role XLAM]
  Role --> Gate[Core.CanPerform]

  Gate -- Deny --> Reject[Reject + log]
  Gate -- Allow --> Inbox[Append Inbox event]

  Inbox --> Proc[Processor]
  Proc --> Lock[Acquire lock]
  Lock --> Apply[Apply domain changes]
  Apply --> Outbox[Write Outbox event]
  Outbox --> Publish[Publish to SharePoint]
  Publish --> HQ[HQ aggregation]
  HQ --> Global[Global snapshot]
```

---

## Processor Semantics

* Inbox rows are processed in deterministic order
* Idempotency is enforced via applied-event markers
* Each event is atomic with respect to its domain
* Poison events are isolated and logged without blocking the queue

---

## What still needs concrete specification (next pass)

1. Event schemas (columns, types, constraints)
2. Config schemas (warehouse, station)
3. Lock timeout and heartbeat values
4. Retry and poison-row policy
5. Snapshot format and cadence

---

## Appendix — Terminology

* **Inbox**: append-only queue of requested changes
* **Outbox**: append-only record of completed facts
* **Processor**: single-writer that applies events
* **Authoritative store**: warehouse-local `.xlsb`
* **Global snapshot**: read-only HQ aggregation

```

---

## Why I’m doing it this way
- The sandbox link system is **not reliable** for iterative architecture work
- Copy-paste Markdown **never fails**
- This preserves Mermaid blocks exactly
- This avoids you burning time fighting the UI instead of designing

If you want next:
- I can break **“What still needs concrete specification”** into one subsection per table (event schemas first), or  
- generate a **GitHub-ready commit diff** against your repo structure.

**Aiden** — (your primary assistant) — 2026-01-26  
**Flash** — Instant (fast responses) — 2026-01-26  
**Sage** — Thinking (deliberate reasoning) — 2026-01-26  
**Atlas** — Pro (big-context, heavy lifting) — 2026-01-26  
**Orion** — Auto (agentic, runs tools/workflows) — 2026-01-26
```