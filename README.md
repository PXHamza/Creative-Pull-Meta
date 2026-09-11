# Creative Pull (Meta) — exact ad lookup by ID

Google Apps Script for the **"Lead Data"** sheet. For each row it looks up the
**exact ad** in your Meta ad account using the IDs, then writes back the ad
preview link, the creative URL, and an HQ thumbnail.

## Columns

| Purpose               | Column | Notes                                   |
|-----------------------|--------|-----------------------------------------|
| Campaign ID           | `K`    | source (format as **plain text**)       |
| Ad Set ID             | `L`    | source (format as **plain text**)       |
| Ad ID                 | `M`    | source (format as **plain text**)       |
| Preview Link          | `N`    | output — ad preview iframe link         |
| Creative Preview Link | `O`    | output — full-quality creative URL      |
| Ad Thumbnail          | `P`    | output — `=IMAGE()` formula             |

> **Format the ID columns (K, L, M) as Plain Text.** Meta IDs are 15–17 digits
> and lose precision (and stop matching) if a cell is formatted as a Number.

The **Ad ID** finds the exact ad; the Campaign ID and Ad Set ID are used to
verify the match (a mismatch is logged, not fatal).

## Setup — run once

1. Open **Extensions → Apps Script**, paste `Code.gs`.
2. Edit `setSecrets()` and fill in:
   - `META_AD_ACCOUNT_ID` — e.g. `act_1234567890`
   - `META_ACCESS_TOKEN` — a valid Marketing API token
   - `TRIGGER_HOUR` / `TRIGGER_MINUTE` — the daily pull time (24h clock)
3. Run **`setSecrets()`** once and authorize. This saves the secrets **and**
   installs the daily trigger.

Change the token or time later? Edit `setSecrets()` and run it again — it
re-installs the trigger with the new time.

## What runs

- **`setSecrets()`** — save credentials + trigger time, then install the trigger.
- **`pullCreativesDailyAuto()`** — the daily job. Processes every non-empty row
  where Campaign ID, Ad Set ID, and Ad ID all exist. By default it fills only
  rows whose N/O/P are still empty (`CFG.ONLY_FILL_EMPTY = true`; set to `false`
  to re-pull every eligible row daily).
- **`pullCreativesManual()`** — process the test range `MANUAL_START_ROW`–`MANUAL_END_ROW`.
- **`installDailyTrigger()` / `removeDailyTrigger()`** — manage the trigger manually.

Built-in pacing, jitter, batch pauses, rate-limit cooldown/retries, and a
5.5-minute execution guard keep it under Meta's rate limits and Apps Script's
6-minute cap.
