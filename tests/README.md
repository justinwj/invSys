# tests

Repository test assets for `invSys`.

- `unit/`: VBA unit-level test harness files and fixtures.
- `integration/`: end-to-end workbook flow tests.
- `fixtures/`: sample workbook/data fixtures used by tests.
- `tooling/`: PowerShell contract tests and synthetic fixtures for developer
  maintenance/runtime evidence tools. These tests do not require Excel or an
  operational workbook.

Note: this scaffold is intentionally non-destructive and does not alter existing VBA source.
