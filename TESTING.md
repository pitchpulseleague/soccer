# Validation of Pitch Pulse 5.4.0

The supplied tests execute the actual HTML JavaScript and Apps Script function bodies with synthetic records and in-memory service doubles. They do not contact the production API, modify a Google spreadsheet, send email or push notifications, or deploy anything.

## Release result

**All 75 checks passed on November 3, 2026**, using the final files in the portable package layout:

| Suite | Passed |
| --- | ---: |
| Release and API contracts | 9 |
| Match rules and data integrity | 21 |
| Backend security, notifications, and setup | 19 |
| Frontend behavior | 22 |
| Generated HTML and escaping contracts | 4 |
| **Total** | **75** |

The rendering checks also parsed **26 generated interface fragments and 232 decoded event handlers** in the English and Spanish paths. These counts describe the included checks, not a percentage of all possible application behavior.

## Run the checks locally

Unzip the package, open a terminal in its folder, and run:

```sh
node tests/release-check.cjs
node tests/rules.test.cjs
node tests/backend.test.cjs
node tests/frontend.test.cjs
node tests/render-contract.test.cjs
```

The tests use standard Node modules. `render-contract.test.cjs` additionally uses `python3` and Python's standard HTML parser to decode generated attributes as HTML. No npm or pip packages are needed. This release was checked with Node 24.22.1 and Python 3.12.14.

The backend and frontend behavior suites also accept an explicit source-file path as their first argument. Without one, they find the `Code.gs` or `index.html` next to the `tests` folder.

## What is covered

| Suite | Coverage |
| --- | --- |
| `release-check.cjs` | All inline scripts and backend JavaScript parse; versions agree; original deployment and asset references remain; API handlers exist; every frontend API action is supported; no merge markers or duplicate global backend functions. |
| `rules.test.cjs` | Pacific/DST dates, reserved and reconfirmed counts, Rule2B, guest/member conflicts, guest removal with disabled privileges or over-capacity matches, waitlist handling, no-show correction, reliability samples, late drops, same-day suspension, effective-dated recurrence, future stop dates, canceled exceptions, and voting eligibility. |
| `backend.test.cjs` | Legacy signature protection, new RSVP namespace, read-only email preview, invitation links in both email formats, channel preferences/fallback, push acceptance and idempotency, uncertain email outcomes, quota priority, repeat-run deduplication, notification replacement rules, morning windows, private diagnostics, and quiet additive setup. |
| `frontend.test.cjs` | Strict score input, Pacific calendar exports, reconfirmation display, honest reliability labels, weekday/12:15 defaults, saved future-series dates, draft preservation, serialized writes, account-switch races, uncertain saves, email confirmation, guest removal, attendance correction controls, focus management, and match ordering. |
| `render-contract.test.cjs` | English/Spanish key parity; hostile quote/entity inputs cannot escape their inline JavaScript string; actual core renderers produce HTML whose decoded event handlers compile; a short match does not claim Game on after kickoff. |

## What these results establish

The changed logic passes the tested cases using the final replacement files. The test data includes capacity boundaries, date/time boundaries, interrupted requests, conflicting guest names, future changes beyond the current generation window, and both English and Spanish rendering paths. Tests use synthetic names, addresses, and account values.

The renderer checks inspect generated markup and execute selected UI behavior with a small DOM double. They are not visual screenshots or a full browser/device accessibility audit. Similarly, Google-service doubles establish the code's intended behavior but cannot prove that Google authorization, a production trigger, email delivery, or OneSignal configuration will work in a particular deployed environment.

## Finish validation after deployment

Follow the post-publication section of `DEPLOYMENT.md`. In particular, verify the new `/exec` version, the single dispatcher trigger, admin/player role separation, one intentional invitation, one intentional test-match signup flow, the notification status panel, and the calendar time on an actual phone.

The package includes `checkDeploymentV540()` for a read-only schema/time-zone/trigger diagnostic in the Apps Script editor. Its successful result confirms the checked installation structure; it does not send a test message or certify delivery.
