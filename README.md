# Sasa — JoblyHub Outreach Assistant (Cloudflare Edition)

This is a migration of the existing Sasa / JoblyHub outreach assistant to a Cloudflare-native runtime. It replaces the original Python web-app runtime and external scheduler/database setup with:

- **Cloudflare Workers** — API and scheduled automation.
- **Cloudflare D1** — organizations, settings, send ledger, notifications and daily history.
- **Cloudflare Workers Static Assets** — dashboard UI.
- **Resend HTTP API** — outbound email delivery (SMTP is not used from the Worker).
- **Native Cloudflare Cron Triggers** — no external scheduler service or always-on server.

The original JoblyHub colors, outreach copy direction, tracker concepts, test mode, daily cap, follow-up flow, activity log, campaign metrics and organization lifecycle fields are retained in this first Cloudflare migration. The existing Python files are not required at runtime.

## Important migration notes

1. **The old database is not embedded in the source ZIP.** Export your existing organization tracker and send-history data from the current installation before retiring it. Import the organization list as CSV in the dashboard. If you need historical send logs and statuses preserved exactly, export and transform those tables separately before cutover.
2. **Do not infer consent.** Existing tracker rows without explicit, documented written-consent evidence are imported as `consent_status=unknown` and are ineligible for automatic sending. `approved=1` alone is not enough. The automatic scheduler requires all of: `approved=1`, `consent_status='written'`, a non-empty `consent_record`, and `opted_out=0`.
3. **Do not ship `.env` files.** This package intentionally contains no live credentials. Configure all secrets using Wrangler.
4. **Email delivery is changed from SMTP to an HTTPS email API.** Cloudflare Workers do not run the old Python SMTP/IMAP code. Resend must approve and verify the sending domain before live sends. Replies are not currently auto-classified by IMAP; set status manually in the tracker for now. Provider delivery/bounce webhooks and mailbox reply processing can be added as a separate integration after the core migration is stable.
5. **Test mode is ON and sending is PAUSED by default.** Do not switch to live until you have tested sending, sender authentication, consent evidence, unsubscribe links and duplicate prevention.

## Requirements

- Node.js LTS and npm.
- A Cloudflare account with Workers and D1 enabled.
- A domain whose DNS you control. If `joblyhub.com` is already in your Cloudflare account, use a subdomain such as `sasa.joblyhub.com`; you do not need to purchase another domain. Otherwise purchase/connect the domain you want to use.
- A Resend account and verified sending domain/subdomain.

## Local setup

```bash
npm install
npx wrangler login
npx wrangler d1 create sasa-joblyhub-db
```

Copy the `database_id` returned by the D1 create command into `wrangler.toml`, replacing `REPLACE_WITH_D1_DATABASE_ID`.

Apply the schema locally and remotely:

```bash
npm run db:migrate:local
npm run db:migrate:remote
```

Create strong random secrets (use a password manager or a cryptographically secure generator):

```bash
npx wrangler secret put ADMIN_TOKEN
npx wrangler secret put RESEND_API_KEY
```

`ADMIN_TOKEN` protects all dashboard API operations. `RESEND_API_KEY` is the provider key. Unsubscribe uses a unique random token per contact stored in D1, so no separate unsubscribe secret is needed.

Set the production `APP_BASE_URL` in `wrangler.toml` to your final HTTPS domain, for example `https://sasa.joblyhub.com`. Configure `FROM_EMAIL` to an address on the domain verified with Resend.

Run locally:

```bash
npm run dev
```

The UI will load, but admin API requests require the `ADMIN_TOKEN` secret. For local testing, create `.dev.vars` (never commit it):

```dotenv
ADMIN_TOKEN=replace-with-a-long-random-secret
RESEND_API_KEY=re_your_test_key
```

## Deploy to Cloudflare

```bash
npm run db:migrate:remote
npm run deploy
```

Then open the Worker in Cloudflare, go to **Settings → Domains & Routes**, and attach your chosen custom domain. Configure DNS through Cloudflare. The `wrangler.toml` asset configuration deploys the dashboard alongside the Worker.

## Native schedule

`wrangler.toml` uses these Cloudflare Cron Triggers (Cloudflare cron schedules are UTC; Ghana is UTC+0):

- `0 8 * * 1-5` — weekdays at 08:00: check due follow-ups (up to three).
- `*/15 9-16 * * 1-5` — weekdays every 15 minutes through the original daytime sending window. The Worker guard stops sends after 16:30. Each invocation sends at most one initial email, and the configured daily maximum is enforced.

This is a Cloudflare-managed scheduled Worker, not an external cron-job service. If you want a single 08:00 batch instead of the original paced 09:00–16:30 send window, update the cron and sending logic deliberately; don't simply remove the schedule, or automatic sending will stop.

## CSV import

The dashboard's **Import CSV / Excel** accepts `.csv` and `.xlsx` files up to 5 MB. For the original JoblyHub tracker workbook, the importer searches worksheets for a header row containing organization/name and email columns, and attempts to preserve the legacy highlighted NGO marker where spreadsheet styles expose it. For reliable results, include an explicit `is_ngo` column. Minimum columns:

```csv
name,email
Example Organization,contact@example.org
```

Recommended columns (these preserve legacy tracker state such as Day 1–10 sent flags and follow-up history):

```csv
row_number,day_batch,day_number,batch_position,planned_send_date,planned_followup_date,name,sector,email,is_ngo,consent_status,consent_record,approved,opted_out,notes
1,Day 1,1,1,2026-10-12,2026-10-19,Example Organization,Education,contact@example.org,no,written,Consent form 2026-10-01,yes,no,Contact approved for outreach
```

Accepted consent values for `consent_status` include `written`, `yes`, `true`, `1`, or `consented`. `approved` is not treated as consent; approval to send and consent evidence are separate checks. Only use them when you have the corresponding written consent record. The `consent_record` should identify the evidence (for example, a consent form reference and date), not just say “yes”. `approved` accepts yes/true/1. Missing fields default to safe values and do not make the row eligible.

Import is an upsert by normalized email address. It updates organization details but will not clear an existing opt-out or downgrade an existing consent record based on a blank value.

## Admin dashboard

1. Open the dashboard.
2. Click **Admin access** and enter the value used for `ADMIN_TOKEN`.
3. Keep **Assistant active** OFF and **Test mode** ON initially.
4. Import a small CSV and confirm rows show as not eligible until both written consent evidence and approval are recorded.
5. Use **Send test email** to send a test to the configured test inbox.
6. Check provider logs and dashboard Activity Log.
7. Verify the unsubscribe link on a live-like approved test contact before switching to live. Do not send to real recipients unless you have a valid legal basis and verified sending configuration.
8. Enable the assistant only when ready.

## Security checklist

- Never commit `.dev.vars`, `.env`, provider keys, admin tokens, or database credentials.
- Use separate, strong random values for each secret.
- Rotate any credentials that were ever stored in an uploaded `.env` file or committed to a repository.
- Keep the assistant paused until sender-domain verification and a test send succeed.
- Restrict admin-token access to the owner; the token grants access to contact data and send controls.
- Consent and opt-out are enforced in the scheduler; opt-out disables approval immediately.
- Each send attempt has an idempotency key, and successful sends are logged before the next scheduled batch.
- Keep a backup/export of the existing tracker before migrating data.

## Runtime routes

- `GET /api/health` — public health check.
- `GET /api/dashboard` — dashboard KPIs (admin token required).
- `GET /api/organizations` — paginated tracker (admin token required).
- `POST /api/import` — CSV import (admin token required).
- `PATCH /api/organizations/:id` — update consent/approval/opt-out/status (admin token required).
- `GET/PUT /api/settings` — settings (admin token required).
- `POST /api/run-test` — send one test message to the test inbox (admin token required).
- `GET /api/activity` — send logs and notifications (admin token required).
- `GET /api/analytics` — campaign metrics (admin token required).
- `GET /unsubscribe?token=...` — public unsubscribe endpoint.

## Known migration scope / follow-up work

- The first Cloudflare version includes Dashboard, Email Tracker, CSV import, Daily Schedule, Analytics, Activity Log and Settings. The original template had a separate Reports page; weekly numbers are included in the analytics/funnel data, while downloadable PDF/XLSX export is not yet ported.
- Automatic IMAP reply classification is not included because the original IMAP client cannot be reused as-is in Workers. Until an inbound email/webhook integration is added, update replies/bounces in the tracker manually. Provider bounce webhooks should be configured before increasing volume.
- The old `import_tracker.py` expects an XLSX file and a Python database session. This Cloudflare edition uses the dashboard's CSV importer instead; export Excel as CSV UTF-8 first.
- This is a migration starter. Run a local build/typecheck and test in a staging Worker before connecting production DNS or enabling live sends.
