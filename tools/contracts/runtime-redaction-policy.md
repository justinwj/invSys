# Runtime evidence redaction policy

Version: `1.0.0`

The default runtime-state report contains schemas, counts, identifiers,
statuses, versions, hashes, and freshness metadata. It never contains row-level
inventory, customer, shipment, recipe, credential, or custom-field values.

The extractor must redact a field when its normalized name contains or equals a
credential marker such as `password`, `pin`, `pinhash`, `token`, `secret`,
`credential`, or `apikey`. Redacted values are replaced with `[REDACTED]`, and
the report records only the field name and redaction count.

Machine and session identities are emitted only as one-way hashes. Windows,
Office, NAS, and service credentials are never emitted. A future row-value
diagnostic mode requires an explicit opt-in parameter, a separate documented
field policy, and tests proving that the default behavior remains unchanged.

Fixture values prefixed with `REDACTION_SENTINEL_` are deliberately inert test
markers, not credentials. They exist only to prove that redaction removes
values from generated evidence.
