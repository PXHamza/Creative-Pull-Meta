// =============================================================================
//
//
//
// Testing Creative Pulling  —  ID-BASED LOOKUP
//   Looks up the creative in Ads Manager by Campaign ID, Ad Set ID, and Ad ID
//   (instead of by Campaign / Ad Set / Ad NAME).
//
//
//
// =============================================================================

/****************************************************
 * CONFIG — Adjust columns, schedule, sizing, pacing here
 ****************************************************/
const AD_CREATIVE_CONFIG = {
  SHEET_NAME: 'Lead Data',

  // Column letters (numeric equivalents computed automatically)
  //
  // ---- SOURCE: which ad to look up. We now search Ads Manager by ID. ----
  // Update these letters to wherever the IDs live in your sheet.
  COL_CAMPAIGN_ID: 'I',  // Source: Campaign ID (optional — used to scope / verify the match)
  COL_ADSET_ID:    'J',  // Source: Ad Set ID  (optional — used to scope / verify the match)
  COL_AD_ID:       'K',  // Source: Ad ID (primary — an Ad ID uniquely identifies the ad)
  //
  // IMPORTANT: Format these ID columns as PLAIN TEXT in the sheet. Meta IDs are
  // 15–17 digits long and will lose precision (and stop matching) if the cell is
  // formatted as a Number.
  //
  // ---- OUTPUT: where results are written (unchanged) ----
  COL_AD_PREVIEW:     'V',  // Output: ad preview iframe link
  COL_CREATIVE_LINK:  'W',  // Output: full-quality creative URL (image or video watch link)
  COL_AD_THUMBNAIL:   'X',  // Output: =IMAGE() formula for HQ thumbnail

  // Image sizing (in pixels)
  IMAGE_WIDTH_PX:  100,
  IMAGE_HEIGHT_PX: 100,
  ROW_HEIGHT_PX:   100,

  // Daily trigger schedule (24h clock, script timezone)
  TRIGGER_HOUR:   15,
  TRIGGER_MINUTE: 0,

  // Manual test range
  MANUAL_START_ROW: 2,
  MANUAL_END_ROW:   5,

  // ===== PACING & RATE LIMIT CONFIG =====
  SLEEP_MS_BETWEEN_ROWS:        1500,   // base delay between each row (was 400)
  JITTER_MS:                    500,    // random extra 0–500ms added each row
  BATCH_PAUSE_EVERY_N_ROWS:     20,     // every N rows, take a longer pause
  BATCH_PAUSE_MS:               30000,  // 30s long pause every batch
  RATE_LIMIT_COOLDOWN_MS:       300000, // 5 min pause when rate-limited
  MAX_RATE_LIMIT_RETRIES:       3,      // how many times to retry a rate-limited row
  MAX_EXECUTION_MS:             330000, // 5.5 min — Apps Script kills at 6 min, so bail before that
};

/****************************************************
 * MANUAL FUNCTION — Test on a hardcoded row range
 ****************************************************/
function pullAdCreativeManual() {
  const startRow = AD_CREATIVE_CONFIG.MANUAL_START_ROW;
  const endRow   = AD_CREATIVE_CONFIG.MANUAL_END_ROW;
  Logger.log(`▶️  MANUAL run for rows ${startRow}–${endRow}`);
  processRows_(startRow, endRow);
}

/****************************************************
 * AUTOMATED FUNCTION — Runs daily at TRIGGER_HOUR
 ****************************************************/
function pullAdCreativeDailyAuto() {
  const cfg = AD_CREATIVE_CONFIG;
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(cfg.SHEET_NAME);
  if (!sheet) throw new Error(`Sheet "${cfg.SHEET_NAME}" not found.`);

  const colAdId = colLetterToNum_(cfg.COL_AD_ID);
  const colV = colLetterToNum_(cfg.COL_AD_PREVIEW);
  const colW = colLetterToNum_(cfg.COL_CREATIVE_LINK);
  const colX = colLetterToNum_(cfg.COL_AD_THUMBNAIL);

  const lastDataRow = sheet.getLastRow();
  if (lastDataRow < 2) {
    Logger.log('ℹ️  No data rows to process.');
    return;
  }

  // Last non-empty row in the Ad ID column
  const idValues = sheet.getRange(2, colAdId, lastDataRow - 1, 1).getValues();
  let lastIdRow = 1;
  for (let i = idValues.length - 1; i >= 0; i--) {
    if (String(idValues[i][0] || '').trim() !== '') {
      lastIdRow = i + 2;
      break;
    }
  }

  if (lastIdRow < 2) {
    Logger.log('ℹ️  No Ad IDs found in source column.');
    return;
  }

  // Last filled row across V, W, X
  const minCol = Math.min(colV, colW, colX);
  const maxCol = Math.max(colV, colW, colX);
  const outputRange = sheet.getRange(2, minCol, lastIdRow - 1, maxCol - minCol + 1).getValues();
  const vIdx = colV - minCol;
  const wIdx = colW - minCol;
  const xIdx = colX - minCol;

  let lastFilledRow = 1;
  for (let i = outputRange.length - 1; i >= 0; i--) {
    const row = outputRange[i];
    const filled =
      String(row[vIdx] || '').trim() !== '' ||
      String(row[wIdx] || '').trim() !== '' ||
      String(row[xIdx] || '').trim() !== '';
    if (filled) {
      lastFilledRow = i + 2;
      break;
    }
  }

  const startRow = lastFilledRow + 1;
  const endRow = lastIdRow;

  if (startRow > endRow) {
    Logger.log(`✅ Nothing to do. Last filled row: ${lastFilledRow}, last Ad ID row: ${lastIdRow}.`);
    return;
  }

  Logger.log(`▶️  AUTO run for rows ${startRow}–${endRow} (last filled: ${lastFilledRow}, last Ad ID: ${lastIdRow})`);
  processRows_(startRow, endRow);
}

/****************************************************
 * SHARED — Loop through a row range with rate-limit handling
 ****************************************************/
function processRows_(startRow, endRow) {
  const cfg = AD_CREATIVE_CONFIG;

  if (startRow < 2) {
    Logger.log(`⚠️  startRow ${startRow} is below 2, bumping to 2 to protect header.`);
    startRow = 2;
  }

  const runStartTime = Date.now();
  let success = 0, failed = 0, skipped = 0;
  let processedInBatch = 0;

  for (let row = startRow; row <= endRow; row++) {
    // Bail if we're approaching the 6-min Apps Script execution cap
    if (Date.now() - runStartTime > cfg.MAX_EXECUTION_MS) {
      Logger.log(`⏰ Approaching execution time limit. Stopping at row ${row - 1}. Next trigger run will resume.`);
      break;
    }

    let attempts = 0;
    let done = false;

    while (!done && attempts <= cfg.MAX_RATE_LIMIT_RETRIES) {
      try {
        const result = pullAdCreative_(row);
        if (result === 'SUCCESS') success++;
        else if (result === 'SKIPPED') skipped++;
        else failed++;
        done = true;
      } catch (err) {
        const isRateLimit = isRateLimitError_(err.message);
        if (isRateLimit && attempts < cfg.MAX_RATE_LIMIT_RETRIES) {
          attempts++;
          Logger.log(`🛑 Rate limited on row ${row}. Cooldown ${cfg.RATE_LIMIT_COOLDOWN_MS / 1000}s (attempt ${attempts}/${cfg.MAX_RATE_LIMIT_RETRIES})…`);
          Utilities.sleep(cfg.RATE_LIMIT_COOLDOWN_MS);
        } else {
          Logger.log(`❌ Row ${row} threw: ${err.message}`);
          failed++;
          done = true;
        }
      }
    }

    processedInBatch++;

    // Base delay + random jitter between every row
    const jitter = Math.floor(Math.random() * cfg.JITTER_MS);
    Utilities.sleep(cfg.SLEEP_MS_BETWEEN_ROWS + jitter);

    // Longer pause every N rows
    if (processedInBatch >= cfg.BATCH_PAUSE_EVERY_N_ROWS && row < endRow) {
      Logger.log(`⏸️  Batch pause: ${cfg.BATCH_PAUSE_MS / 1000}s after ${processedInBatch} rows…`);
      Utilities.sleep(cfg.BATCH_PAUSE_MS);
      processedInBatch = 0;
    }
  }

  Logger.log(`✅ Done. Rows ${startRow}–${endRow} | Success: ${success} | Failed: ${failed} | Skipped: ${skipped}`);
}

/****************************************************
 * HELPER — Detect Meta rate-limit error messages
 ****************************************************/
function isRateLimitError_(message) {
  if (!message) return false;
  const m = message.toLowerCase();
  return (
    m.includes('user request limit reached') ||
    m.includes('too many api calls') ||
    m.includes('"code":17') ||
    m.includes('"code":4') ||         // app-level rate limit
    m.includes('"code":32') ||        // page-level rate limit
    m.includes('"code":613') ||       // custom limit
    m.includes('rate limit')
  );
}

/****************************************************
 * CORE — Fetch ad preview + creative URL + thumbnail for one row
 ****************************************************/
function pullAdCreative_(rowNum) {
  const cfg = AD_CREATIVE_CONFIG;

  if (rowNum < 2) {
    Logger.log(`⛔ Refusing to write to row ${rowNum} (header row).`);
    return 'FAILED';
  }

  const props = PropertiesService.getScriptProperties();
  const ACCOUNT_ID   = props.getProperty('META_AD_ACCOUNT_ID');
  const ACCESS_TOKEN = props.getProperty('META_ACCESS_TOKEN');
  const API_VERSION  = props.getProperty('API_VERSION') || 'v19.0';

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(cfg.SHEET_NAME);
  if (!sheet) throw new Error(`Sheet "${cfg.SHEET_NAME}" not found.`);

  const colCampaignId = colLetterToNum_(cfg.COL_CAMPAIGN_ID);
  const colAdsetId    = colLetterToNum_(cfg.COL_ADSET_ID);
  const colAdId       = colLetterToNum_(cfg.COL_AD_ID);
  const colV = colLetterToNum_(cfg.COL_AD_PREVIEW);
  const colW = colLetterToNum_(cfg.COL_CREATIVE_LINK);
  const colX = colLetterToNum_(cfg.COL_AD_THUMBNAIL);

  // Read the IDs. getDisplayValue() preserves the cell text so large IDs are not
  // mangled into scientific notation when read.
  const campaignId = cleanId_(sheet.getRange(rowNum, colCampaignId).getDisplayValue());
  const adsetId    = cleanId_(sheet.getRange(rowNum, colAdsetId).getDisplayValue());
  const adId       = cleanId_(sheet.getRange(rowNum, colAdId).getDisplayValue());

  if (!adId && !adsetId && !campaignId) {
    Logger.log(`⏭️  Row ${rowNum}: no Campaign/Ad Set/Ad ID, skipping.`);
    return 'SKIPPED';
  }

  Logger.log(`🔎 Row ${rowNum}: searching by IDs (campaign=${campaignId || '-'}, adset=${adsetId || '-'}, ad=${adId || '-'})`);

  const ad = findAdByIds_(API_VERSION, ACCESS_TOKEN, campaignId, adsetId, adId);

  if (!ad) {
    sheet.getRange(rowNum, colW).setValue('NOT FOUND');
    Logger.log(`❌ Row ${rowNum}: no ad found for the provided ID(s).`);
    return 'FAILED';
  }

  const creative = ad.creative || {};
  Logger.log(`✅ Row ${rowNum}: ad ID ${ad.id}`);

  // ===== AD PREVIEW LINK (Column V) =====
  let adPreviewLink = '';
  try {
    const previewUrl = buildUrl_(
      `https://graph.facebook.com/${API_VERSION}/${ad.id}/previews`,
      { ad_format: 'DESKTOP_FEED_STANDARD', access_token: ACCESS_TOKEN }
    );
    const previewRes = fetchJson_(previewUrl);
    if (previewRes.data && previewRes.data.length > 0 && previewRes.data[0].body) {
      const match = previewRes.data[0].body.match(/src="([^"]+)"/);
      if (match) adPreviewLink = match[1].replace(/&amp;/g, '&');
    }
  } catch (e) {
    if (isRateLimitError_(e.message)) throw e; // let the retry logic handle it
    Logger.log(`⚠️  Row ${rowNum}: preview fetch failed: ${e.message}`);
  }

  // ===== CREATIVE LINK + THUMBNAIL =====
  let creativeUrl = '';
  let thumbnailUrl = '';

  // (a) Video ad — use public watch link + thumbnail from creative (no /videos call)
  if (creative.video_id) {
    creativeUrl = `https://www.facebook.com/watch/?v=${creative.video_id}`;
    if (creative.object_story_spec
        && creative.object_story_spec.video_data
        && creative.object_story_spec.video_data.image_url) {
      thumbnailUrl = creative.object_story_spec.video_data.image_url;
    } else if (creative.thumbnail_url) {
      thumbnailUrl = creative.thumbnail_url;
    }
    Logger.log(`🎥 Row ${rowNum}: video watch link + thumbnail set`);
  }

  // (b) Image hash → /adimages
  if (!creativeUrl && creative.image_hash) {
    creativeUrl = resolveImageHash_(ACCOUNT_ID, API_VERSION, ACCESS_TOKEN, creative.image_hash);
  }

  // (c) Direct image_url
  if (!creativeUrl && creative.image_url) {
    creativeUrl = creative.image_url;
  }

  // (d) object_story_spec
  if (!creativeUrl && creative.object_story_spec) {
    const oss = creative.object_story_spec;
    if (oss.link_data && oss.link_data.picture) {
      creativeUrl = oss.link_data.picture;
    } else if (oss.photo_data && oss.photo_data.url) {
      creativeUrl = oss.photo_data.url;
    } else if (oss.video_data && oss.video_data.image_url) {
      creativeUrl = oss.video_data.image_url;
    }
  }

  // (e) asset_feed_spec (DCO)
  if (!creativeUrl && creative.asset_feed_spec) {
    const afs = creative.asset_feed_spec;
    if (afs.images && afs.images.length > 0 && afs.images[0].hash) {
      creativeUrl = resolveImageHash_(ACCOUNT_ID, API_VERSION, ACCESS_TOKEN, afs.images[0].hash);
    }
  }

  // (f) Page post fallback
  if (!creativeUrl && creative.effective_object_story_id) {
    const postUrl = buildUrl_(
      `https://graph.facebook.com/${API_VERSION}/${creative.effective_object_story_id}`,
      {
        fields: 'full_picture,permalink_url,attachments{media,media_type,url}',
        access_token: ACCESS_TOKEN
      }
    );
    const postRes = fetchJson_(postUrl);
    creativeUrl =
      postRes.full_picture ||
      (postRes.attachments && postRes.attachments.data && postRes.attachments.data[0]?.media?.image?.src) ||
      postRes.permalink_url || '';
  }

  // STEP 3 — Write results
  if (adPreviewLink) {
    sheet.getRange(rowNum, colV).setValue(adPreviewLink);
  }

  if (creativeUrl) {
    sheet.getRange(rowNum, colW).setValue(creativeUrl);
    const imageForDisplay = thumbnailUrl || creativeUrl;
    sheet.getRange(rowNum, colX).setFormula(
      `=IMAGE("${imageForDisplay}", 4, ${cfg.IMAGE_HEIGHT_PX}, ${cfg.IMAGE_WIDTH_PX})`
    );
    sheet.setRowHeightsForced(rowNum, 1, cfg.ROW_HEIGHT_PX);
    Logger.log(`✅ Row ${rowNum}: V=preview, W=creative, X=image`);
    return 'SUCCESS';
  } else {
    sheet.getRange(rowNum, colW).setValue('NO CREATIVE FOUND');
    Logger.log(`⚠️  Row ${rowNum}: ad found but no creative URL resolved.`);
    return 'FAILED';
  }
}

/****************************************************
 * HELPER — Find an ad by Campaign ID / Ad Set ID / Ad ID
 *
 * Priority:
 *   1. Ad ID      → fetch the ad directly (an Ad ID uniquely identifies the ad).
 *   2. Ad Set ID  → first ad on the ad set's /ads edge.
 *   3. Campaign ID→ first ad on the campaign's /ads edge.
 ****************************************************/
function findAdByIds_(apiVersion, accessToken, campaignId, adsetId, adId) {
  const adFields = 'id,name,campaign_id,adset_id,creative{id,image_hash,image_url,thumbnail_url,object_story_spec,asset_feed_spec,effective_object_story_id,video_id}';

  // (1) Ad ID — most reliable. Fetch the ad node directly.
  if (adId) {
    try {
      const url = buildUrl_(
        `https://graph.facebook.com/${apiVersion}/${adId}`,
        { fields: adFields, access_token: accessToken }
      );
      const ad = fetchJson_(url);
      if (ad && ad.id) {
        // Optional sanity checks against the Campaign / Ad Set IDs, if provided.
        if (campaignId && String(ad.campaign_id) !== String(campaignId)) {
          Logger.log(`⚠️  Ad ${adId} campaign_id ${ad.campaign_id} != expected ${campaignId}`);
        }
        if (adsetId && String(ad.adset_id) !== String(adsetId)) {
          Logger.log(`⚠️  Ad ${adId} adset_id ${ad.adset_id} != expected ${adsetId}`);
        }
        return ad;
      }
    } catch (e) {
      if (isRateLimitError_(e.message)) throw e;
      Logger.log(`⚠️  Direct Ad ID fetch failed for ${adId}: ${e.message}`);
    }
  }

  // (2) Ad Set ID — list ads under the ad set and take the first.
  if (adsetId) {
    const ad = findAdOnEdge_(adsetId, apiVersion, accessToken, adFields);
    if (ad) return ad;
  }

  // (3) Campaign ID — list ads under the campaign and take the first.
  if (campaignId) {
    const ad = findAdOnEdge_(campaignId, apiVersion, accessToken, adFields);
    if (ad) return ad;
  }

  return null;
}

/****************************************************
 * HELPER — Fetch the first ad on a parent's /ads edge
 *          (parent = an Ad Set ID or a Campaign ID)
 ****************************************************/
function findAdOnEdge_(parentId, apiVersion, accessToken, adFields) {
  try {
    const url = buildUrl_(
      `https://graph.facebook.com/${apiVersion}/${parentId}/ads`,
      { fields: adFields, limit: '10', access_token: accessToken }
    );
    const res = fetchJson_(url);
    if (res.data && res.data.length > 0) return res.data[0];
  } catch (e) {
    if (isRateLimitError_(e.message)) throw e;
    Logger.log(`⚠️  Edge lookup failed for ${parentId}: ${e.message}`);
  }
  return null;
}

/****************************************************
 * HELPER — Normalize an ID read from the sheet
 ****************************************************/
function cleanId_(raw) {
  let s = String(raw == null ? '' : raw)
    .replace(/ /g, ' ')  // non-breaking spaces
    .replace(/\s+/g, '')      // strip all whitespace
    .trim();
  // Drop a trailing ".0" that can appear if an ID was read as a number.
  s = s.replace(/\.0+$/, '');
  return s;
}

/****************************************************
 * HELPER — Resolve image hash to direct CDN URL
 ****************************************************/
function resolveImageHash_(accountId, apiVersion, accessToken, hash) {
  const url = buildUrl_(
    `https://graph.facebook.com/${apiVersion}/${accountId}/adimages`,
    {
      hashes: JSON.stringify([hash]),
      fields: 'hash,url,permalink_url,original_width,original_height',
      access_token: accessToken
    }
  );
  const res = fetchJson_(url);
  let imgObj = null;
  if (Array.isArray(res.data) && res.data.length > 0) {
    imgObj = res.data[0];
  } else if (res.data && typeof res.data === 'object') {
    const keys = Object.keys(res.data);
    if (keys.length > 0) imgObj = res.data[keys[0]];
  }
  return imgObj ? (imgObj.url || imgObj.permalink_url || '') : '';
}

/****************************************************
 * HELPER — Column letter (A, B, ..., AA) → 1-based number
 ****************************************************/
function colLetterToNum_(letter) {
  const s = String(letter).toUpperCase();
  let n = 0;
  for (let i = 0; i < s.length; i++) {
    n = n * 26 + (s.charCodeAt(i) - 64);
  }
  return n;
}

/****************************************************
 * HELPER — Build URL with encoded params
 ****************************************************/
function buildUrl_(base, params) {
  const qs = Object.keys(params)
    .map(k => `${encodeURIComponent(k)}=${encodeURIComponent(params[k])}`)
    .join('&');
  return `${base}?${qs}`;
}

/****************************************************
 * HELPER — Fetch JSON with error handling
 ****************************************************/
function fetchJson_(url) {
  const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  const code = res.getResponseCode();
  const body = res.getContentText();
  if (code !== 200) {
    Logger.log(`❌ HTTP ${code}: ${body.substring(0, 500)}`);
    throw new Error(`Meta API error ${code}: ${body.substring(0, 300)}`);
  }
  return JSON.parse(body);
}

/****************************************************
 * TRIGGER SETUP — Run once to install the daily trigger
 ****************************************************/
function installDailyTrigger() {
  const cfg = AD_CREATIVE_CONFIG;

  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'pullAdCreativeDailyAuto') {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger('pullAdCreativeDailyAuto')
    .timeBased()
    .everyDays(1)
    .atHour(cfg.TRIGGER_HOUR)
    .nearMinute(cfg.TRIGGER_MINUTE)
    .create();

  Logger.log(`✅ Daily trigger installed for ${cfg.TRIGGER_HOUR}:${String(cfg.TRIGGER_MINUTE).padStart(2, '0')}`);
}

/****************************************************
 * TRIGGER REMOVAL — Run to remove the daily trigger
 ****************************************************/
function removeDailyTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  let removed = 0;
  triggers.forEach(t => {
    if (t.getHandlerFunction() === 'pullAdCreativeDailyAuto') {
      ScriptApp.deleteTrigger(t);
      removed++;
    }
  });
  Logger.log(`✅ Removed ${removed} trigger(s).`);
}
// =============================================================================
//
//
//
// Testing Creative Pulling
//
//
//
// =============================================================================
