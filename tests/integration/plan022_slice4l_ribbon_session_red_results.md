# Plan 022 Slice 4l Ribbon Session RED

- Date: 2026-08-17
- Active slice: Plan 022 Slice 4l — ribbon session-state controls
- Package: Core, consumed by Operations and Admin
- Protecting test: `Test-Plan022Slice4lRibbonSessionControls.ps1`
- Expected behavior: Send To immediately refreshes the live Operations status;
  server and invSys authentication have explicit independent toggles; Server
  Sign Out also signs out invSys, clears the target, disconnects SMB, and
  disables operator actions; disconnected invSys Sign In fails closed.
- Observed RED: 0/8 passed. A follow-up access-status assertion then failed at
  8/9 because the disconnected Operations label still said `Access: Ready`.

The behavioral cause was present in reachable source: target selection
invalidated retired role-specific Ribbon IDs but omitted the deployed
Operations dropdown/status IDs, while the only generic Sign Out action retained
the NAS target and disconnected invSys Sign In could revive remembered context.
