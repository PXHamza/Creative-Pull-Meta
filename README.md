# Creative Pull (Meta) — ID-based lookup

Google Apps Script that pulls the ad **preview link**, **creative URL**, and an
**HQ thumbnail** for each row of a Google Sheet, using the Meta Marketing API.

## What changed

The lookup now searches Ads Manager by **Campaign ID, Ad Set ID, and Ad ID**
instead of by Campaign / Ad Set / Ad **name**. Everything else — output columns,
pacing, rate-limit handling, the daily trigger, and the creative-resolution
fallbacks — is unchanged.

### Lookup priority

1. **Ad ID** — fetched directly (`GET /{ad_id}`); an Ad ID uniquely identifies the ad.
   When present, the Campaign ID / Ad Set ID are used only as a sanity check
   (a mismatch is logged, not fatal).
2. **Ad Set ID** — falls back to the first ad on `GET /{adset_id}/ads`.
3. **Campaign ID** — falls back to the first ad on `GET /{campaign_id}/ads`.

## Configuration

All settings live in `AD_CREATIVE_CONFIG` at the top of `Code.gs`.

### Source columns (update to match your sheet)

| Config key        | Default | Meaning                                  |
|-------------------|---------|------------------------------------------|
| `COL_CAMPAIGN_ID` | `I`     | Campaign ID (optional — scope/verify)    |
| `COL_ADSET_ID`    | `J`     | Ad Set ID (optional — scope/verify)      |
| `COL_AD_ID`       | `K`     | Ad ID (primary lookup)                   |

> **Format the ID columns as PLAIN TEXT.** Meta IDs are 15–17 digits and lose
> precision (and stop matching) if a cell is formatted as a Number.

### Output columns (unchanged)

| Config key          | Default | Meaning                              |
|---------------------|---------|--------------------------------------|
| `COL_AD_PREVIEW`    | `V`     | Ad preview iframe link               |
| `COL_CREATIVE_LINK` | `W`     | Full-quality creative URL            |
| `COL_AD_THUMBNAIL`  | `X`     | `=IMAGE()` formula for the thumbnail |

### Script Properties (Project Settings → Script Properties)

- `META_AD_ACCOUNT_ID` — e.g. `act_1234567890`
- `META_ACCESS_TOKEN`  — a valid Marketing API access token
- `API_VERSION` *(optional)* — defaults to `v19.0`

## Usage

- `pullAdCreativeManual()` — process the hardcoded test range (`MANUAL_START_ROW`–`MANUAL_END_ROW`).
- `pullAdCreativeDailyAuto()` — resume from the last filled output row up to the last Ad ID row.
- `installDailyTrigger()` / `removeDailyTrigger()` — manage the daily time trigger.
