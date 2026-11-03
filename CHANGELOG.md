# Pitch Pulse League 5.4.0

Complete replacement frontend and Apps Script backend for the existing league. See `DEPLOYMENT.md` for the deployment order, upgrade functions, publication steps, post-deployment checks, and rollback.

## Signup and match display

- Match cards put the date, kickoff, field, match status, and the player's answer first. Roster previews can expand to show everyone.
- New matches default to **12:15 PM**, and new weekly series default to **Monday–Friday**. Both remain editable for exceptions.
- A match that is short of its minimum no longer claims to be playing simply because kickoff passed. The scheduled end is labeled as a time boundary rather than proof that everyone attended.
- The next active match appears ahead of matches whose scheduled time has already finished today.
- After a time or field change, existing places stay reserved. Only answers acknowledged for the revised schedule count toward the minimum needed for “Game on.” Pending reconfirmations remain visible after kickoff.
- Existing named guests can still be removed after guest privileges are revoked, the league guest limit is reduced, or organizer walk-ins put the match over capacity.
- A member's name cannot be newly added as a guest. A legacy guest reservation that conflicts with a member's own signup produces an explanation naming the host. The app never merges identities based on a name match.
- Whole-number score validation enforces the supported 0–99 range.

## Safer saves and forms

- The page serializes write requests and shows a saving state. Duplicate taps are blocked while the relevant write is pending.
- Responses from a previous account/session are discarded after logout or account switching.
- A timeout is treated as an uncertain outcome. The page reads the latest server state and warns before the user retries; it does not automatically repeat a potentially completed invitation or match creation.
- Profile appearance changes preserve other unsaved profile inputs.
- Unsaved league settings and rules survive redraws, and automatic refresh waits while a settings draft or modal is active.
- Sheets have a dialog name, keyboard focus containment, background isolation, Escape behavior, and focus restoration. Live announcements are scoped to status messages.
- User-supplied strings in inline JavaScript handlers are escaped for both the JavaScript string and HTML attribute layers.

## Email links and notifications

- Opening an RSVP email link is read-only. The player reviews the match and explicitly confirms the answer before it is saved.
- The legacy RSVP landing page no longer generates a valid answer signature from arbitrary email and match parameters.
- RSVP signatures use a new versioned namespace. Old signed RSVP emails must be replaced by a new reminder or answered through normal app login. Invitation tokens, PIN hashes, sessions, and the existing app secret are retained.
- Notifications are stored in an additive `Notifications` table with separate email and push states for each recipient.
- The dispatcher sends pending work, respects quota reserves, prioritizes important updates, retries safe failures, and expires stale notices.
- Accepted sends are not intentionally repeated. Interrupted email sends with an unknown outcome are held for review; OneSignal retries reuse the same idempotency key.
- Email preference is honored for meaningful notifications. Phone-first preferences have an email fallback when a subscribed device is unavailable or a phone notification cannot be sent.
- The admin status panel distinguishes pending, provider-accepted, expired, and review-needed channel records. Provider acceptance does not prove device delivery or that a player read the message.
- Morning reminders are eligible around **8:30 AM Pacific**, before the 9:15 AM three-hour cutoff for the regular 12:15 PM match. The shared dispatcher checks approximately every five minutes.
- Nudging a player creates one personalized notification job rather than separate overlapping push/email jobs.
- Invitations include the weekday lunch schedule and a signup URL in the plain-text body as well as the HTML email.

## Attendance and recurring schedules

- Organizer attendance is stored per player and match as `Attended` or `NoShow`, with the recorder and timestamp. A no-show is only available after kickoff.
- Recording a player as present corrects a prior no-show. Repeating an attendance action does not create duplicate penalties or notices.
- Reliability uses unique recorded match outcomes. One match cannot count as both a late-drop failure and a no-show failure. A recorded attendance correction takes precedence.
- A player with no recorded sample has no reliability percentage. Old aggregate late-drop counters remain preserved as legacy data rather than being presented as newly verified attendance.
- Suspending a player releases only commitments for matches that have not started; it preserves earlier matches on the same day.
- Weekly series store effective-dated configurations. An edit scheduled beyond the four-week generation window takes effect on its selected date.
- Stopping a series in the future preserves generation of eligible matches before the stop date. Removed future matches are canceled with their RSVP rows retained.
- Started matches and manually canceled exceptions are not silently rewritten or revived by series edits.
- Player-of-the-match voting begins after the scheduled end or after a final score is saved. Recorded no-shows are excluded.

## Time zones and compatibility

- Match instants use `America/Los_Angeles`, including daylight saving time. Google Calendar links and downloaded event files use absolute UTC start/end values.
- Calendar additions remain individual events. Subsequent changes in Pitch Pulse do not automatically update an event already copied into a player's calendar.
- The existing rule allowing RSVP edits throughout match day is retained. A final score or the end of the league day closes self-service answers; the scheduled 75-minute end alone does not.
- The upgrade adds schema without deleting or renaming existing sheets or columns. It does not infer old attendance, reset credentials, or replace saved league settings.
- `upgradeToV540()` performs the schema upgrade and installs the dispatcher. `applyRecommendedLaunchSettings()` is a separate explicit choice for the simpler launch switches.
- The existing API URL, site URL, brand assets, language support, and optional feature structure remain in use.

## Validation boundary

The package includes regression tests that exercise the actual code using synthetic data and local Google-service doubles. No deployment, live league writes, emails, or push messages were performed to build this release. Google authorization, production trigger execution, real message delivery, and physical-phone behavior must be checked after publication using the steps in `DEPLOYMENT.md`.
