# Pitch Pulse account controls

These are local candidate changes, not a deployment. Back up the spreadsheet and Apps Script before deploying; test on a separate spreadsheet and deployment first.

## Implemented
- Administrators can reversibly deactivate players and send re-invitations. Accepting the invitation restores access to the same profile/history. Each invitation has a fresh token, expires after seven days, and can only be accepted once. Re-send outstanding pre-upgrade invitations after deploying this version. Only players initiate privacy removal; the legacy admin delete endpoint now deactivates rather than deleting.
- Notification preference `None` prevents unsolicited league emails and pushes, including queued jobs and fallback. User-requested security/verification codes continue.
- Change PIN requires authenticated account, actual current PIN and matching six-digit new PIN; all sessions and reset codes are invalidated.
- Email change verifies the new inbox, requires current PIN again, migrates structured identity/history references, expires old-address queued messages and invalidates sessions. Protected configured administrator addresses cannot use this flow.
- Player removal requires current PIN, saved recovery code acknowledgment and typed confirmation. A random 64-hex recovery code is shown before the final action. Only its keyed hash and random anonymous history ID remain on the server.
- Removal retains an inactive initials-only profile and historical games. Future signups/guest reservations are cleared. Structured identity fields, relevant human-text fields, and account-scoped properties are scrubbed for the old email/full name, including team JSON and queued message contents. Unrelated players and global secrets are preserved. Profile fields outside the known schema are cleared too.
- Recovery requires a fresh valid administrator invitation plus the saved code. Previous name/email and PIN are not restored; the returning player supplies new details and chooses a new PIN. Prior admin privileges are not restored. Notifications remain off until the player opts in.

## Important limitations
Initials and linked game history are pseudonymous, not guaranteed anonymous. This application does not erase Google Sheets revision history, external backups, copies/exports, already delivered emails, browser caches on other devices, or OneSignal provider records. Administrators must handle those retention surfaces separately. Exact-match scrubbing cannot identify arbitrary nicknames or misspelled references in free text. Legacy administrator logs record only names, not actor IDs; matching name-only attributions are reduced to initials, so historical entries belonging to another person with the identical name cannot be distinguished. Other player profiles and structured identities remain unchanged.

Apps Script and Sheets do not provide a multi-sheet database transaction. Handled write failures attempt rollback; an execution termination, timeout or failed rollback can require administrator repair. A staging migration test and backup are essential. Recovery mappings remain available if a deletion write outcome is uncertain. Do not blindly retry an uncertain account mutation.

Email changes require signing in again and reconnecting phone notifications to the new identity. Previously issued notification links to the old email stop resolving. Stored notification references are migrated too, so a subsequent privacy removal can scrub them; already delivered external copies remain outside the application.

## Verification
25 account/notification/privacy checks and 8 administrator lifecycle checks passed, along with both frontend unit suites and JavaScript syntax.
Node syntax checks and mocked behavioral tests cover current-PIN validation, session revocation, code expiry/attempt limits, email migration/rollback, notification opt-out after queue claim, removal confirmations, table/property scrubbing, escaped JSON names, clearing future signups, invitation-plus-code recovery and one-time consumption. These are not live Apps Script, spreadsheet or OneSignal integration tests.
