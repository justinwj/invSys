# invSys VBA Maintenance Candidates

Synthetic fixture expectations:

- `DuplicateAlpha` and `DuplicateBeta` are a duplicate-body candidate.
- `UnreferencedCandidate` is review-only and is never auto-deleted.
- The unresolved dynamic `Application.Run` call is reported, not guessed.
- Dynamic and compatibility roots remain retained.
