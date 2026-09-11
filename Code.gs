// =============================================================================
//
//   Creative Pull (Meta)  —  EXACT AD LOOKUP BY ID
//
//   Reads Campaign ID (K), Ad Set ID (L), Ad ID (M) from the "Lead Data" sheet,
//   finds the exact ad in the ad account, and writes back:
//       Preview Link          → Column N
//       Creative Preview Link → Column O
//       Ad Thumbnail          → Column P
//
//   Secrets (Ad Account ID, Access Token) and the daily trigger time are set
//   once via setSecrets().
//
// =============================================================================

/****************************************************
 * 1) SET SECRETS + INSTALL DAILY TRIGGER
 *    Run this ONCE (and again whenever you change the token or the time).
 ****************************************************/
function setSecrets() {
  const props = PropertiesService.getScriptProperties();

  // ----- Meta credentials -----
  props.setProperty('META_AD_ACCOUNT_ID', 'act_1229');   // e.g. act_1234567890
  props.setProperty('META_ACCESS_TOKEN',  'EAASWVrjY');  // Marketing API token
  props.setProperty('API_VERSION',        'v19.0');       // Graph API version

  // ----- Daily trigger time (24h clock, script timezone) -----
  // The daily pull runs once a day at this time. Change these and re-run
  // setSecrets() to move it.
  props.setProperty('TRIGGER_HOUR',   '15');  // 0–23  (15 = 3 PM)
  props.setProperty('TRIGGER_MINUTE', '0');   // 0–59

  // (Re)install the daily trigger using the time above.
  installDailyTrigger();

  Logger.log('✅ Secrets saved and daily trigger installed.');
}

/****************************************************
 * CONFIG — columns, sheet, sizing, pacing
 ****************************************************/
const CFG = {
  SHEET_NAME: 'Lead Data',

  // ----- SOURCE columns (the IDs used to find the exact ad) -----
  COL_CAMPAIGN_ID: 'K',
  COL_ADSET_ID:    'L',
  COL_AD_ID:       'M',
  // NOTE: format these ID columns as PLAIN TEXT — Meta IDs are 15–17 digits and
  // lose precision (and stop matching) if the cell is formatted as a Number.

  // ----- OUTPUT columns -----
  COL_PREVIEW_LINK:  'N',  // ad preview iframe link
  COL_CREATIVE_LINK: 'O',  // full-quality creative URL (image or video watch link)
  COL_THUMBNAIL:     'P',  // =IMAGE() formula for the HQ thumbnail

  // Only fill rows whose outputs are still empty (true = skip already-pulled rows).
  // Set false to re-pull every row that has all three IDs on each run.
  ONLY_FILL_EMPTY: true,

  // Image sizing (pixels)
  IMAGE_WIDTH_PX:  100,
  IMAGE_HEIGHT_PX: 100,
  ROW_HEIGHT_PX:   100,

  // Manual test range (used by pullCreativesManual)
  MANUAL_START_ROW: 2,
  MANUAL_END_ROW:   5,

  // ----- Pacing & rate-limit handling -----
  SLEEP_MS_BETWEEN_ROWS:    1500,
  JITTER_MS:                500,
  BATCH_PAUSE_EVERY_N_ROWS: 20,
  BATCH_PAUSE_MS:           30000,
  RATE_LIMIT_COOLDOWN_MS:   300000,  // 5 min pause when rate-limited
  MAX_RATE_LIMIT_RETRIES:   3,
  MAX_EXECUTION_MS:         330000,  // 5.5 min — bail before Apps Script's 6-min cap
};

/****************************************************
 * 2a) MANUAL RUN — process a fixed row range
 ****************************************************/
function pullCreativesManual() {
  Logger.log(`▶️  MANUAL run for rows ${CFG.MANUAL_START_ROW}–${CFG.MANUAL_END_ROW}`);
  const rows = [];
  for (let r = CFG.MANUAL_START_ROW; r <= CFG.MANUAL_END_ROW; r++) rows.push(r);
  processRows_(rows);
}

/****************************************************
 * 2b) DAILY AUTO RUN — installed by setSecrets()/installDailyTrigger()
 *     Processes every non-empty row that has all three IDs.
 ****************************************************/
function pullCreativesDailyAuto() {
  const sheet = getSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    Logger.log('ℹ️  No data rows.');
    return;
  }

  const cCamp  = colLetterToNum_(CFG.COL_CAMPAIGN_ID);
  const cAdset = colLetterToNum_(CFG.COL_ADSET_ID);
  const cAd    = colLetterToNum_(CFG.COL_AD_ID);
  const cN     = colLetterToNum_(CFG.COL_PREVIEW_LINK);
  const cO     = colLetterToNum_(CFG.COL_CREATIVE_LINK);
  const cP     = colLetterToNum_(CFG.COL_THUMBNAIL);

  const numRows = lastRow - 1;

  // Read the three ID columns (display values preserve large IDs as text).
  const campVals  = sheet.getRange(2, cCamp,  numRows, 1).getDisplayValues();
  const adsetVals = sheet.getRange(2, cAdset, numRows, 1).getDisplayValues();
  const adVals    = sheet.getRange(2, cAd,    numRows, 1).getDisplayValues();

  // Read the three output columns in one contiguous block.
  const minOut = Math.min(cN, cO, cP);
  const maxOut = Math.max(cN, cO, cP);
  const outVals = sheet.getRange(2, minOut, numRows, maxOut - minOut + 1).getValues();
  const nIdx = cN - minOut, oIdx = cO - minOut, pIdx = cP - minOut;

  const rows = [];
  for (let i = 0; i < numRows; i++) {
    const hasIds =
      cleanId_(campVals[i][0])  !== '' &&
      cleanId_(adsetVals[i][0]) !== '' &&
      cleanId_(adVals[i][0])    !== '';
    if (!hasIds) continue;   // only non-empty rows where all three IDs exist

    if (CFG.ONLY_FILL_EMPTY) {
      const alreadyFilled =
        String(outVals[i][nIdx] || '').trim() !== '' ||
        String(outVals[i][oIdx] || '').trim() !== '' ||
        String(outVals[i][pIdx] || '').trim() !== '';
      if (alreadyFilled) continue;
    }
    rows.push(i + 2);
  }

  if (rows.length === 0) {
    Logger.log('✅ Nothing to do — no rows with IDs need pulling.');
    return;
  }

  Logger.log(`▶️  AUTO run for ${rows.length} row(s): ${rows[0]}…${rows[rows.length - 1]}`);
  processRows_(rows);
}

/****************************************************
 * SHARED — process a list of row numbers with pacing + rate-limit retries
 ****************************************************/
function processRows_(rows) {
  const runStart = Date.now();
  let success = 0, failed = 0, skipped = 0, inBatch = 0;

  for (let idx = 0; idx < rows.length; idx++) {
    const row = rows[idx];

    if (Date.now() - runStart > CFG.MAX_EXECUTION_MS) {
      Logger.log(`⏰ Time limit approaching. Stopping before row ${row}. Next run resumes.`);
      break;
    }

    let attempts = 0, done = false;
    while (!done && attempts <= CFG.MAX_RATE_LIMIT_RETRIES) {
      try {
        const result = pullOne_(row);
        if (result === 'SUCCESS') success++;
        else if (result === 'SKIPPED') skipped++;
        else failed++;
        done = true;
      } catch (err) {
        if (isRateLimitError_(err.message) && attempts < CFG.MAX_RATE_LIMIT_RETRIES) {
          attempts++;
          Logger.log(`🛑 Rate limited on row ${row}. Cooldown ${CFG.RATE_LIMIT_COOLDOWN_MS / 1000}s (attempt ${attempts}/${CFG.MAX_RATE_LIMIT_RETRIES})…`);
          Utilities.sleep(CFG.RATE_LIMIT_COOLDOWN_MS);
        } else {
          Logger.log(`❌ Row ${row} threw: ${err.message}`);
          failed++;
          done = true;
        }
      }
    }

    inBatch++;
    Utilities.sleep(CFG.SLEEP_MS_BETWEEN_ROWS + Math.floor(Math.random() * CFG.JITTER_MS));

    if (inBatch >= CFG.BATCH_PAUSE_EVERY_N_ROWS && idx < rows.length - 1) {
      Logger.log(`⏸️  Batch pause ${CFG.BATCH_PAUSE_MS / 1000}s after ${inBatch} rows…`);
      Utilities.sleep(CFG.BATCH_PAUSE_MS);
      inBatch = 0;
    }
  }

  Logger.log(`✅ Done. Success: ${success} | Failed: ${failed} | Skipped: ${skipped}`);
}

/****************************************************
 * CORE — pull preview + creative + thumbnail for ONE row
 ****************************************************/
function pullOne_(rowNum) {
  if (rowNum < 2) {
    Logger.log(`⛔ Refusing to write to row ${rowNum} (header row).`);
    return 'FAILED';
  }

  const props = PropertiesService.getScriptProperties();
  const ACCOUNT_ID   = props.getProperty('META_AD_ACCOUNT_ID');
  const ACCESS_TOKEN = props.getProperty('META_ACCESS_TOKEN');
  const API_VERSION  = props.getProperty('API_VERSION') || 'v19.0';
  if (!ACCOUNT_ID || !ACCESS_TOKEN) {
    throw new Error('Missing META_AD_ACCOUNT_ID / META_ACCESS_TOKEN — run setSecrets() first.');
  }

  const sheet = getSheet_();
  const cCamp  = colLetterToNum_(CFG.COL_CAMPAIGN_ID);
  const cAdset = colLetterToNum_(CFG.COL_ADSET_ID);
  const cAd    = colLetterToNum_(CFG.COL_AD_ID);
  const cN     = colLetterToNum_(CFG.COL_PREVIEW_LINK);
  const cO     = colLetterToNum_(CFG.COL_CREATIVE_LINK);
  const cP     = colLetterToNum_(CFG.COL_THUMBNAIL);

  const campaignId = cleanId_(sheet.getRange(rowNum, cCamp).getDisplayValue());
  const adsetId    = cleanId_(sheet.getRange(rowNum, cAdset).getDisplayValue());
  const adId       = cleanId_(sheet.getRange(rowNum, cAd).getDisplayValue());

  if (!adId) {
    Logger.log(`⏭️  Row ${rowNum}: no Ad ID, skipping.`);
    return 'SKIPPED';
  }

  Logger.log(`🔎 Row ${rowNum}: campaign=${campaignId || '-'}, adset=${adsetId || '-'}, ad=${adId}`);

  const ad = findExactAd_(API_VERSION, ACCESS_TOKEN, campaignId, adsetId, adId);
  if (!ad) {
    sheet.getRange(rowNum, cO).setValue('NOT FOUND');
    Logger.log(`❌ Row ${rowNum}: ad ${adId} not found.`);
    return 'FAILED';
  }

  const creative = ad.creative || {};
  Logger.log(`✅ Row ${rowNum}: found ad ${ad.id}`);

  // ===== PREVIEW LINK (Column N) =====
  let previewLink = '';
  try {
    const previewUrl = buildUrl_(
      `https://graph.facebook.com/${API_VERSION}/${ad.id}/previews`,
      { ad_format: 'DESKTOP_FEED_STANDARD', access_token: ACCESS_TOKEN }
    );
    const previewRes = fetchJson_(previewUrl);
    if (previewRes.data && previewRes.data.length > 0 && previewRes.data[0].body) {
      const match = previewRes.data[0].body.match(/src="([^"]+)"/);
      if (match) previewLink = match[1].replace(/&amp;/g, '&');
    }
  } catch (e) {
    if (isRateLimitError_(e.message)) throw e;
    Logger.log(`⚠️  Row ${rowNum}: preview fetch failed: ${e.message}`);
  }

  // ===== CREATIVE LINK (Column O) + THUMBNAIL (Column P) =====
  let creativeUrl = '';
  let thumbnailUrl = '';

  // (a) Video ad
  if (creative.video_id) {
    creativeUrl = `https://www.facebook.com/watch/?v=${creative.video_id}`;
    if (creative.object_story_spec
        && creative.object_story_spec.video_data
        && creative.object_story_spec.video_data.image_url) {
      thumbnailUrl = creative.object_story_spec.video_data.image_url;
    } else if (creative.thumbnail_url) {
      thumbnailUrl = creative.thumbnail_url;
    }
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
      { fields: 'full_picture,permalink_url,attachments{media,media_type,url}', access_token: ACCESS_TOKEN }
    );
    const postRes = fetchJson_(postUrl);
    creativeUrl =
      postRes.full_picture ||
      (postRes.attachments && postRes.attachments.data && postRes.attachments.data[0]?.media?.image?.src) ||
      postRes.permalink_url || '';
  }

  // ===== WRITE =====
  if (previewLink) sheet.getRange(rowNum, cN).setValue(previewLink);

  if (creativeUrl) {
    sheet.getRange(rowNum, cO).setValue(creativeUrl);
    const imageForDisplay = thumbnailUrl || creativeUrl;
    sheet.getRange(rowNum, cP).setFormula(
      `=IMAGE("${imageForDisplay}", 4, ${CFG.IMAGE_HEIGHT_PX}, ${CFG.IMAGE_WIDTH_PX})`
    );
    sheet.setRowHeightsForced(rowNum, 1, CFG.ROW_HEIGHT_PX);
    Logger.log(`✅ Row ${rowNum}: N=preview, O=creative, P=image`);
    return 'SUCCESS';
  }

  sheet.getRange(rowNum, cO).setValue('NO CREATIVE FOUND');
  Logger.log(`⚠️  Row ${rowNum}: ad found but no creative URL resolved.`);
  return 'FAILED';
}

/****************************************************
 * HELPER — find the EXACT ad by Ad ID, verified against Campaign/Ad Set ID
 ****************************************************/
function findExactAd_(apiVersion, accessToken, campaignId, adsetId, adId) {
  const adFields = 'id,name,campaign_id,adset_id,creative{id,image_hash,image_url,thumbnail_url,object_story_spec,asset_feed_spec,effective_object_story_id,video_id}';

  const url = buildUrl_(
    `https://graph.facebook.com/${apiVersion}/${adId}`,
    { fields: adFields, access_token: accessToken }
  );

  let ad;
  try {
    ad = fetchJson_(url);
  } catch (e) {
    if (isRateLimitError_(e.message)) throw e;
    Logger.log(`⚠️  Ad ID ${adId} lookup failed: ${e.message}`);
    return null;
  }

  if (!ad || !ad.id) return null;

  // Verify this really is the exact ad the row points to.
  if (campaignId && String(ad.campaign_id) !== String(campaignId)) {
    Logger.log(`⚠️  Ad ${adId}: campaign_id ${ad.campaign_id} ≠ expected ${campaignId}`);
  }
  if (adsetId && String(ad.adset_id) !== String(adsetId)) {
    Logger.log(`⚠️  Ad ${adId}: adset_id ${ad.adset_id} ≠ expected ${adsetId}`);
  }

  return ad;
}

/****************************************************
 * HELPER — resolve an image hash to a direct CDN URL
 ****************************************************/
function resolveImageHash_(accountId, apiVersion, accessToken, hash) {
  const url = buildUrl_(
    `https://graph.facebook.com/${apiVersion}/${accountId}/adimages`,
    { hashes: JSON.stringify([hash]), fields: 'hash,url,permalink_url,original_width,original_height', access_token: accessToken }
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
 * HELPER — detect Meta rate-limit error messages
 ****************************************************/
function isRateLimitError_(message) {
  if (!message) return false;
  const m = message.toLowerCase();
  return (
    m.includes('user request limit reached') ||
    m.includes('too many api calls') ||
    m.includes('"code":17') ||
    m.includes('"code":4') ||   // app-level
    m.includes('"code":32') ||  // page-level
    m.includes('"code":613') || // custom limit
    m.includes('rate limit')
  );
}

/****************************************************
 * HELPER — normalize an ID read from the sheet
 ****************************************************/
function cleanId_(raw) {
  // \s also matches non-breaking spaces in JS, so this clears every kind of
  // whitespace (IDs never contain any). Also drop a trailing ".0" that appears
  // if an ID was accidentally read as a number.
  return String(raw == null ? '' : raw).replace(/\s+/g, '').replace(/\.0+$/, '').trim();
}

/****************************************************
 * HELPER — sheet handle
 ****************************************************/
function getSheet_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CFG.SHEET_NAME);
  if (!sheet) throw new Error(`Sheet "${CFG.SHEET_NAME}" not found.`);
  return sheet;
}

/****************************************************
 * HELPER — column letter (A, B, …, AA) → 1-based number
 ****************************************************/
function colLetterToNum_(letter) {
  const s = String(letter).toUpperCase();
  let n = 0;
  for (let i = 0; i < s.length; i++) n = n * 26 + (s.charCodeAt(i) - 64);
  return n;
}

/****************************************************
 * HELPER — build a URL with encoded params
 ****************************************************/
function buildUrl_(base, params) {
  const qs = Object.keys(params)
    .map(k => `${encodeURIComponent(k)}=${encodeURIComponent(params[k])}`)
    .join('&');
  return `${base}?${qs}`;
}

/****************************************************
 * HELPER — fetch JSON with error handling
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
 * TRIGGER — install the daily trigger from the saved time
 *    (called automatically by setSecrets)
 ****************************************************/
function installDailyTrigger() {
  const props = PropertiesService.getScriptProperties();
  const hour   = parseInt(props.getProperty('TRIGGER_HOUR')   || '15', 10);
  const minute = parseInt(props.getProperty('TRIGGER_MINUTE') || '0',  10);

  // Remove any existing trigger for the daily function first.
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'pullCreativesDailyAuto') ScriptApp.deleteTrigger(t);
  });

  ScriptApp.newTrigger('pullCreativesDailyAuto')
    .timeBased()
    .everyDays(1)
    .atHour(hour)
    .nearMinute(minute)
    .create();

  Logger.log(`✅ Daily trigger installed for ${hour}:${String(minute).padStart(2, '0')}`);
}

/****************************************************
 * TRIGGER — remove the daily trigger
 ****************************************************/
function removeDailyTrigger() {
  let removed = 0;
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'pullCreativesDailyAuto') {
      ScriptApp.deleteTrigger(t);
      removed++;
    }
  });
  Logger.log(`✅ Removed ${removed} trigger(s).`);
}
