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
| Preview Link          | `N`    | output — **permanent Ads Manager deep link** |
| Creative Preview Link | `O`    | output — full-quality creative URL      |
| Ad Thumbnail          | `P`    | output — `=IMAGE()` formula             |

> **Format the ID columns (K, L, M) as Plain Text.** Meta IDs are 15–17 digits
> and lose precision (and stop matching) if a cell is formatted as a Number.

The **Ad ID** finds the exact ad; the Campaign ID and Ad Set ID are used to
verify the match (a mismatch is logged, not fatal).

## Link expiry

Meta's URLs are deliberately temporary — there's no permanent version from Meta:

- **Preview Link (N)** is now a **permanent Ads Manager deep link**
  (`…/manage/ads?act=<acct>&selected_ad_ids=<adId>`). It opens the exact ad,
  needs no API call, and never expires.
- **Creative / Thumbnail (O, P)** use Meta's signed `scontent.*.fbcdn.net`
  image URLs, which expire after a few days. Video links
  (`facebook.com/watch`) are already permanent. To keep the image links working
  **without re-hosting**, the daily trigger runs with `CFG.ONLY_FILL_EMPTY =
  false`, so it **refreshes every eligible row each day** — the stored URL is
  never more than ~1 day old, well inside Meta's signature lifetime. Identical
  ads across multiple rows are resolved once per run (per-Ad-ID cache).
- Want images that are permanent without a daily refresh? That requires
  re-hosting the image bytes somewhere public (e.g. Google Drive shared
  "anyone with the link"). That path is intentionally **not** used here.

> To refresh rows that currently hold expired links, just let the daily trigger
> run (it overwrites them), or run `pullCreativesDailyAuto()` manually once.

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
  where Campaign ID, Ad Set ID, and Ad ID all exist. Default `CFG.ONLY_FILL_EMPTY
  = false` re-pulls every eligible row so image links stay fresh; set it to
  `true` to fill only rows whose N/O/P are still empty (image links will then
  eventually expire).
- **`pullCreativesManual()`** — process the test range `MANUAL_START_ROW`–`MANUAL_END_ROW`.
- **`installDailyTrigger()` / `removeDailyTrigger()`** — manage the trigger manually.

Built-in pacing, jitter, batch pauses, rate-limit cooldown/retries, and a
5.5-minute execution guard keep it under Meta's rate limits and Apps Script's
6-minute cap.
