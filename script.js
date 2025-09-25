/************************************************************
 * Safe onEdit + Daily Email + Helpers  (FINAL per spec)
 * Updates limited to:
 *   (A) D→C FIID selection (hint-first, then full-scan; RE/FW skip; multi-match join; final OTHERS)
 *   (B) Duplicate control with normalization + Exact/Fuzzy tiers (≥0.90)
 * Everything else kept as before.
 ************************************************************/

/************************************************************
 * onEdit(e)
 ************************************************************/
function onEdit(e) {
  const ss = e.source;
  const sheet = e.range.getSheet();
  const col = e.range.getColumn();
  const row = e.range.getRow();

  if (sheet.getName() !== 'Main Sheet') return;

  // --- Re-entry guard ---
  const LOCK_CELL = 'Z1';
  const lockFlag = sheet.getRange(LOCK_CELL).getValue();
  if (lockFlag === 'LOCK') return;
  function withLock(fn) {
    sheet.getRange(LOCK_CELL).setValue('LOCK');
    try { fn(); } finally { sheet.getRange(LOCK_CELL).clearContent(); }
  }

  const isMulti = e.range.getNumRows() > 1 || e.range.getNumColumns() > 1;

  // Touch A1 heartbeat
  if (row > 1) { withLock(() => sheet.getRange('A1').setValue(new Date())); }

  /************** Auto-date in A **************/
  if (col !== 1 && col !== 3 && row > 1 && !isMulti) {
    const dateCell = sheet.getRange(row, 1);
    const lastCol = sheet.getLastColumn();
    const rowData = sheet.getRange(row, 1, 1, lastCol).getValues()[0];

    let hasData = false;
    for (let i = 0; i < rowData.length; i++) {
      if (i !== 0 && i !== 2 && rowData[i] !== '') { hasData = true; break; }
    }

    if (hasData && dateCell.isBlank()) {
      withLock(() => { dateCell.setValue(new Date()); dateCell.setNumberFormat('dd-mmm-yy'); });
    } else if (!hasData && !dateCell.isBlank()) {
      withLock(() => dateCell.clearContent());
    }
  }

  /************** Title (Col D) edits — UPDATED FIID logic **************/
  if (col === 4 && row !== 1) {
    const title = sheet.getRange(row, 4).getValue();
    const oldTitle = e.oldValue || '';

    if (!isMulti) {
      const cCell = sheet.getRange(row, 3); // C
      if (title && typeof title === 'string') {
        if (cCell.isBlank()) {
          const idx = getFiidAliasIndex_(); // { canonSet:[...], aliasToCanon:{}, bankNamesToCanon:{} }
          const fiids = selectFIIDsForTitle_(title, idx); // array of canonical FIIDs or ['OTHERS']
          const out = Array.isArray(fiids) && fiids.length ? fiids.join('-') : 'OTHERS';
          withLock(() => cCell.setValue(out));
        }
      }

      // If Title cleared → also clear C and A (keep old behavior)
      if (title === '' || title === null) {
        withLock(() => { sheet.getRange(row, 3).clearContent(); sheet.getRange(row, 1).clearContent(); });
      }
    }

    // Duplicate check (UPDATED)
    checkDuplicateImmediately(sheet, row, oldTitle, title, { isMulti });
  }

  /************** Status (Col B) edits **************/
  if (col === 2 && row !== 1 && !isMulti) {
    const status = sheet.getRange(row, 2).getValue();
    const previousStatus = e.oldValue;

    if (status === 'Completed') {
      const issueTypeCell = sheet.getRange(row, 5);
      const designatedCell = sheet.getRange(row, 7);
      if (issueTypeCell.isBlank() || designatedCell.isBlank()) {
        withLock(() => sheet.getRange(row, col).setValue(previousStatus || ''));
        try {
          SpreadsheetApp.getUi().alert(
            'Validation Error',
            "Cannot set status to 'Completed'. Columns 'E' (Issue Type) and 'G' (Designated) must be filled.",
            SpreadsheetApp.getUi().ButtonSet.OK
          );
        } catch (err) {
          SpreadsheetApp.getActive().toast(
            "Validation: 'Completed' requires Issue Type (E) and Designated (G). Reverted.",
            'Validation Error', 7
          );
          sheet.getRange(row, 1, 1, sheet.getLastColumn()).setBackground('#fff3cd');
          sheet.getRange(row, 2).setNote("Reverted: 'Completed' needs E & G.");
        }
        return;
      }
    }

    const userEmail = Session.getActiveUser().getEmail();
    const timestamp = new Date();

    if (status === 'Completed' || status === 'Trash') {
      withLock(() => {
        sheet.getRange(row, 9).setValue(timestamp).setNumberFormat('dd-mmm-yy'); // I
        sheet.getRange(row, 10).setValue(userEmail);                              // J

        let data = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];

        const titleText = data[3], colE = 4;
        if (data[colE] === '') {
          if (typeof titleText === 'string' &&
              (titleText.toLowerCase().includes('user') ||
               titleText.toLowerCase().includes('id') ||
               titleText.toLowerCase().includes('password reset') ||
               titleText.toLowerCase().includes('access'))) {
            data[colE] = 'USER & ACCESS';
          } else if (typeof titleText === 'string' && titleText.toLowerCase().includes('report')) {
            data[colE] = 'Reports';
          } else if (typeof titleText === 'string' && titleText.trim() !== '') {
            data[colE] = 'MODIFICATION & PROJECTS';
          }
        }

        data[3] = stripAuditTag_(data[3]);
        data[5] = stripAuditTag_(data[5]);

        const targetSheetName = status === 'Completed' ? 'Completed' : 'Trashed';
        const targetSheet = ss.getSheetByName(targetSheetName);
        targetSheet.appendRow(data);

        const last = targetSheet.getLastRow();
        targetSheet.getRange(last, 1).setNumberFormat('dd-mmm-yy');
        targetSheet.getRange(last, 9).setNumberFormat('dd-mmm-yy');

        sheet.deleteRow(row);
      });
      return;
    }
  }

  /************** Inline audit tag for D & F  **************/
  if ((col === 4 || col === 6) && row > 1 && !isMulti) {
    setInlineAuditTag_(sheet, row, col);
  }
}


/************************************************************
 * Duplicate check — ADVANCED (hybrid scoring + fast index)
 * Drop-in replacement for your existing checkDuplicateImmediately()
 ************************************************************/
function checkDuplicateImmediately(sheet, row, oldValue, newValue, opts) {
  const isMulti = (opts && opts.isMulti) || false;

  const currentFIID = String(sheet.getRange(row, 3).getValue() || '').trim();
  if (!newValue || !currentFIID) return;

  // Normalize the "right of colon" subject
  const newSubjectRaw = rightOfColon_(newValue);
  const newSubjectNorm = normalizeSubject_(newSubjectRaw);
  if (!newSubjectNorm) { clearDupNote_(sheet, row); return; }

  const lastRow = sheet.getLastRow();
  if (lastRow <= 2) { clearDupNote_(sheet, row); return; }

  // Build or fetch a lightweight per-FIID index
  const index = buildFiidRowIndex_(sheet, currentFIID, row);

  // Quick exact hash check (Tier-0)
  const newHash = fastHash_(newSubjectNorm);
  if (index.hashToRows[newHash] && index.hashToRows[newHash].some(r => r !== row)) {
    return handleDuplicateDecision_({
      sheet, row, oldValue, newValue, isMulti,
      level: 'Exact',
      dupRows: uniq_(index.hashToRows[newHash].filter(r => r !== row))
    });
  }

  // Prepare features for hybrid scoring
  const newTokens = tokenizeSubject_(newSubjectNorm);
  const newVec = termFreq_(newTokens);
  const newTris = trigramSet_(newSubjectNorm);

  // Hybrid scan within the same FIID only
  let best = { score: 0, level: 'Similar', rows: [] };

  for (const it of index.items) {
    if (it.row === row) continue;

    // Skip empty subjects after normalization
    if (!it.norm) continue;

    // 1) Cosine TF similarity
    const cos = cosineSim_(newVec, it.tf);
    // 2) Trigram Jaccard
    const jac = jaccard_(newTris, it.tris);
    // 3) Edit distance ratio (Levenshtein)
    const edr = similarity_(newSubjectNorm, it.norm);

    // Weighted hybrid (tuned for short email subjects)
    const hybrid = Math.max(
      cos,                          // robust on token permutations
      0.6 * jac + 0.4 * edr         // shingle overlap + char-level edits
    );

    // Optional field-aware boost (same Issue Type E increases confidence)
    // Safe: read once to avoid too many calls
    if (hybrid >= 0.86) {
      try {
        const e1 = String(sheet.getRange(row, 5).getValue() || '').toLowerCase();
        const e2 = String(sheet.getRange(it.row, 5).getValue() || '').toLowerCase();
        if (e1 && e2 && e1 === e2) {
          // small deterministic boost (max 1.0)
          best = pickBetter_(best, { score: Math.min(1, hybrid + 0.02), level: 'Similar', rows: [it.row] });
          continue;
        }
      } catch (_) {}
    }

    best = pickBetter_(best, { score: hybrid, level: 'Similar', rows: [it.row] });
  }

  // Thresholds: keep Exact path separate; Similar uses tuned cutoffs
  // - Strong: ≥0.92 (almost certain dup)
  // - Normal: ≥0.88 (good dup across wording variants)
  // - Conservative lower bound: ≥0.86 only if multiple hits line up (handled below)
  if (best.score >= 0.92) {
    return handleDuplicateDecision_({
      sheet, row, oldValue, newValue, isMulti,
      level: 'Similar', // still not exact, but treat as strong duplicate
      dupRows: best.rows
    });
  }

  // If there are multiple near-miss rows around 0.86+, aggregate them
  if (best.score >= 0.86) {
    const nearRows = [];
    for (const it of index.items) {
      if (it.row === row || !it.norm) continue;
      const cos = cosineSim_(newVec, it.tf);
      const jac = jaccard_(newTris, it.tris);
      const edr = similarity_(newSubjectNorm, it.norm);
      const hybrid = Math.max(cos, 0.6 * jac + 0.4 * edr);
      if (hybrid >= 0.86) nearRows.push(it.row);
    }
    if (nearRows.length > 0) {
      return handleDuplicateDecision_({
        sheet, row, oldValue, newValue, isMulti,
        level: 'Similar',
        dupRows: uniq_(nearRows)
      });
    }
  }

  // No duplicates
  clearDupNote_(sheet, row);
}

/* ---------------------------------------------------------
 * Helpers for advanced duplicate detection
 * --------------------------------------------------------- */

// Compact per-FIID index with shallow caching
function buildFiidRowIndex_(sheet, fiid, editingRow) {
  const cache = CacheService.getUserCache(); // per-user to avoid cross-edit clashes
  const key = `IDX_${sheet.getSheetId()}_${fiid}`;
  const cached = cache.get(key);
  if (cached) {
    try {
      const parsed = JSON.parse(cached);
      // If editingRow is newer than cached window, we still accept (windowed scan)
      return parsed;
    } catch (_) {}
  }

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  // Scan minimal range: columns C (FIID) and D (Title); also E for optional boost
  const rng = sheet.getRange(2, 3, Math.max(0, lastRow - 1), 3).getValues(); // C,D,E
  const items = [];
  const hashToRows = {};

  // Optional windowing: limit to last ~1500 rows for performance on very large sheets
  const hardLimit = 1500;
  const startIdx = Math.max(0, rng.length - hardLimit);

  for (let i = startIdx; i < rng.length; i++) {
    const r = i + 2;
    if (r === editingRow) continue;

    const fi = rng[i][0] ? String(rng[i][0]).trim() : '';
    if (!fi || fi.toLowerCase() !== fiid.toLowerCase()) continue;

    const title = rng[i][1] ? String(rng[i][1]).trim() : '';
    if (!title) continue;

    const norm = normalizeSubject_(rightOfColon_(title));
    if (!norm) continue;

    const tokens = tokenizeSubject_(norm);
    const tf = termFreq_(tokens);
    const tris = trigramSet_(norm);
    const h = fastHash_(norm);

    if (!hashToRows[h]) hashToRows[h] = [];
    hashToRows[h].push(r);

    items.push({ row: r, norm, tf, tris });
  }

  const out = { items, hashToRows };
  // Cache for 90 seconds — fresh enough for interactive edits
  cache.put(key, JSON.stringify(out), 90);
  return out;
}

// Tokenizer with domain synonyms & numeric handling
function tokenizeSubject_(s) {
  // Expand domain synonyms (password/pwd/pin, reset/reseting etc.)
  const SYN = {
    'pwd': 'password', 'pass': 'password', 'pin': 'pin',
    'resetting': 'reset', 'reseting': 'reset',
    'txn': 'transaction', 'txns': 'transaction', 'trans': 'transaction',
    'otp': 'otp', 'login': 'login', 'logon': 'login', 'signin': 'login', 'sign-in': 'login',
    'card': 'card', 'cards': 'card', 'debit': 'debit', 'credit': 'credit',
    'settlement': 'settlement', 'recon': 'reconcile', 'reconciliation': 'reconcile',
    'portal': 'portal', 'gateway': 'gateway', 'pg': 'gateway'
  };

  // Split tokens
  let toks = s.split(/\s+/).filter(Boolean);

  // Normalize numbers: collapse long numeric strings to #num token to avoid false mismatches
  toks = toks.map(t => /^\d{5,}$/.test(t) ? '#num' : t);

  // Apply synonyms
  toks = toks.map(t => SYN[t] || t);

  // Remove trivial stop tokens (after synonyms)
  const STOP = new Set(['please','plz','kindly','issue','problem','error','bug','task','ticket','need','required','request','help','update','info','regarding','about']);
  toks = toks.filter(t => !STOP.has(t));

  return toks;
}

// Term frequency vector
function termFreq_(tokens) {
  const m = {};
  for (const t of tokens) m[t] = (m[t] || 0) + 1;
  return m;
}

// Cosine similarity for sparse TF vectors
function cosineSim_(a, b) {
  let dot = 0, na = 0, nb = 0;
  for (const k in a) { na += a[k] * a[k]; if (b[k]) dot += a[k] * b[k]; }
  for (const k in b) { nb += b[k] * b[k]; }
  if (!na || !nb) return 0;
  return dot / (Math.sqrt(na) * Math.sqrt(nb));
}

// Character trigram set
function trigramSet_(s) {
  const arr = [];
  const pad = `  ${s}  `; // padding helps catch boundary overlaps
  for (let i = 0; i < pad.length - 2; i++) {
    arr.push(pad.slice(i, i + 3));
  }
  return new Set(arr);
}

// Jaccard on sets
function jaccard_(A, B) {
  if (!A || !B || !A.size || !B.size) return 0;
  let inter = 0;
  // Iterate on smaller set for speed
  if (A.size > B.size) [A, B] = [B, A];
  A.forEach(x => { if (B.has(x)) inter++; });
  const uni = A.size + B.size - inter;
  return uni ? inter / uni : 0;
}

// Fast, stable hash for exact-match buckets
function fastHash_(s) {
  // FNV-1a 32-bit
  let h = 0x811c9dc5;
  for (let i = 0; i < s.length; i++) {
    h ^= s.charCodeAt(i);
    h = (h + ((h << 1) + (h << 4) + (h << 7) + (h << 8) + (h << 24))) >>> 0;
  }
  return h.toString(16);
}

// Decide & apply duplicate UX (keeps your existing UX semantics)
function handleDuplicateDecision_({ sheet, row, oldValue, newValue, isMulti, level, dupRows }) {
  // Color rows (exact=red-ish, similar=orange-ish)
  const bg = (level === 'Exact') ? '#f8d7da' : '#ffe8cc';
  const rowsToMark = uniq_([...dupRows, row]);
  rowsToMark.forEach(r => sheet.getRange(r, 1, 1, sheet.getLastColumn()).setBackground(bg));

  if (isMulti) {
    SpreadsheetApp.getActive().toast(
      `Duplicate ${level === 'Exact' ? '(Exact)' : '(Similar)'} in rows: ${dupRows.join(', ')}`,
      'Duplicate', 7
    );
    sheet.getRange(row, 4).setNote(`Duplicate ${level === 'Exact' ? 'Exact' : 'Similar'} with rows: ${dupRows.join(', ')}`);
    return;
  }

  // Manual override
  const hasOverride = /\b\[dup-ok\]\b/i.test(String(newValue));

  // Prompt with row numbers
  let yes = false;
  try {
    const ui = SpreadsheetApp.getUi();
    const res = ui.alert(
      `Duplicate ${level === 'Exact' ? '(Exact)' : '(Similar)'} Found`,
      `This issue already exists in row(s): ${dupRows.join(', ')}\n` +
      `FIID: ${String(sheet.getRange(row, 3).getValue() || '').trim()}\n` +
      `Subject: ${normalizeSubject_(rightOfColon_(newValue))}\n\nAdd anyway?`,
      ui.ButtonSet.YES_NO
    );
    yes = (res === ui.Button.YES);
  } catch (err) {
    yes = false;
    SpreadsheetApp.getActive().toast(
      `Duplicate ${level === 'Exact' ? '(Exact)' : '(Similar)'} detected; reverting change (no UI available).`,
      'Duplicate', 6
    );
  }

  // Clear highlight before final act
  rowsToMark.forEach(r => sheet.getRange(r, 1, 1, sheet.getLastColumn()).setBackground(null));

  if (yes || hasOverride) {
    sheet.getRange(row, 4).setNote(
      `Duplicate ${level === 'Exact' ? 'Exact' : 'Similar'} kept ${hasOverride ? '(override: [dup-ok]) ' : ''}` +
      `with rows: ${dupRows.join(', ')}`
    );
    return;
  }

  // Revert change (keep your original semantics)
  sheet.getRange(row, 4).setValue(oldValue || '');
  sheet.getRange(row, 3).clearContent();
  sheet.getRange(row, 1).clearContent();
  sheet.getRange(row, 4).setNote(
    `Reverted due to duplicate ${level === 'Exact' ? '(Exact)' : '(Similar)'} with rows: ${dupRows.join(', ')}`
  );
}

function uniq_(arr){ const m={}; const out=[]; for(const x of arr){ if(!m[x]){m[x]=1; out.push(x);} } return out; }

function pickBetter_(a, b){
  // Prefer higher score; if tie, keep the one with more rows (rare) then lower row number (stable)
  if (b.score > a.score) return b;
  if (b.score === a.score) {
    if ((b.rows?.length || 0) > (a.rows?.length || 0)) return b;
    if ((b.rows?.[0] || 1e9) < (a.rows?.[0] || 1e9)) return b;
  }
  return a;
}


/************************************************************
 * EMAIL + TRIGGERS 
 ************************************************************/
function sendDailyIssueEmail() {
  const tz = Session.getScriptTimeZone() || 'Asia/Dhaka';
  const todayPretty = Utilities.formatDate(new Date(), tz, 'MMMM d, yyyy');

  const to  = 'nitish@itcbd.com';
  const cc  = 'mamudul.hassan@itcbd.com, ops@itcbd.com';
  const bcc = 'faruk@itcbd.com';

  const sheetShareLink = 'https://docs.google.com/spreadsheets/d/1g38kVFyp3hSVwCU1HfMyljUax_1WWfA7WhzQfmil_Hk/edit?usp=sharing';

  const ID_LOGO    = '1jpWNOYcgvfKRrzQz95ii0kI0-AJ9G10Z';
  const ID_ANDROID = '1Fw-DzM4WgseYRc1cEyWIJwnyQIT0ecTQ';
  const ID_IOS     = '1r8f6Ln-7LpRuL4QAA88s0mzE4LTpCSXz';

  let logoBlob, androidBlob, iosBlob;
  try { logoBlob    = DriveApp.getFileById(ID_LOGO).getBlob().setName('qpay_logo.png'); } catch (e) {}
  try { androidBlob = DriveApp.getFileById(ID_ANDROID).getBlob().setName('android_qr.png'); } catch (e) {}
  try { iosBlob     = DriveApp.getFileById(ID_IOS).getBlob().setName('ios_qr.png'); } catch (e) {}

  const subject = `ITCPLC: CMS Issue Log — Daily Update (${todayPretty})`;

  const htmlBody = `
  <div style="font-family:'Trebuchet MS', Tahoma, Arial, sans-serif; color:#111; line-height:1.5;">
    <p>Dear Sir,</p>
    <p>Please find below the updated summary of the CMS issue log as of <strong>${todayPretty}.</strong></p>
    <p><a href="${sheetShareLink}" target="_blank" style="color:#4f46e5; text-decoration:none;">View the CMS Issue Log (Google Sheet)</a></p>
    <!-- signature kept same as your version -->
  </div>
  `;

  const inlineImages = {};
  if (logoBlob)    inlineImages.qpay_logo  = logoBlob;
  if (androidBlob) inlineImages.android_qr = androidBlob;
  if (iosBlob)     inlineImages.ios_qr     = iosBlob;

  MailApp.sendEmail({ to, cc, bcc, subject, htmlBody, inlineImages });
}

function createWeeklyTriggers() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'sendDailyIssueEmail') {
      ScriptApp.deleteTrigger(t);
    }
  });

  const days = [
    ScriptApp.WeekDay.SUNDAY,
    ScriptApp.WeekDay.MONDAY,
    ScriptApp.WeekDay.TUESDAY,
    ScriptApp.WeekDay.WEDNESDAY,
    ScriptApp.WeekDay.THURSDAY
  ];

  days.forEach(function(day) {
    ScriptApp.newTrigger('sendDailyIssueEmail')
      .timeBased()
      .atHour(19)
      .nearMinute(55)
      .onWeekDay(day)
      .create();
  });
}


/***********************
 * Helpers for audit tag 
 ***********************/
function nicknameFromEmail_(email) {
  if (!email) return 'Unknown';
  const map = {
    'faruk@itcbd.com':            'Faruk',
    'fardit@itcbd.com':           'Fardit',
    'mamudul.hassan@itcbd.com':   'DCTO',
    'shamim@itcbd.com':           'Arif',
    'mehedi.hassan@itcbd.com':    'Mehedi',
    'naim@itcbd.com':             'Naim',
    'nitish@itcbd.com':           'Nitish',
    'sayma.ahmed@itcbd.com':      'Sayma'
  };
  const key = String(email).trim().toLowerCase();
  if (map[key]) return map[key];
  const local = key.split('@')[0] || '';
  const token = local.split(/[.\-_]/)[0] || local;
  return token ? token.charAt(0).toUpperCase() + token.slice(1) : 'User';
}

function stripAuditTag_(text) {
  if (text === null || text === undefined) return '';
  return String(text).replace(/\s*\n?\(last modified by .* at \d{2}-[A-Za-z]{3}-\d{4}\s+\d{2}:\d{2}\)\s*$/i, '');
}

function setInlineAuditTag_(sheet, row, col) {
  const tz = Session.getScriptTimeZone() || 'Asia/Dhaka';
  const now = new Date();
  const pretty = Utilities.formatDate(now, tz, 'dd-MMM-yyyy HH:mm');

  const cell = sheet.getRange(row, col);
  const plain = String(cell.getDisplayValue() || cell.getValue() || '');
  const cleanedMain = stripAuditTag_(plain).trimEnd();

  if (cleanedMain === '') { cell.clearContent(); return; }

  const email = Session.getActiveUser().getEmail() || '';
  const nick = nicknameFromEmail_(email);

  const auditText = `\n(last modified by ${nick} at ${pretty})`;
  const fullText = cleanedMain + auditText;

  const mainEnd = cleanedMain.length;
  const totalLen = fullText.length;

  const rich = SpreadsheetApp.newRichTextValue()
    .setText(fullText)
    .setTextStyle(0, mainEnd, SpreadsheetApp.newTextStyle().setFontSize(12).setForegroundColor('#111111').build())
    .setTextStyle(mainEnd, totalLen, SpreadsheetApp.newTextStyle().setFontSize(8).setForegroundColor('#9AA0A6').build())
    .build();

  cell.setRichTextValue(rich);
}


/* ========================================================================
   NEW HELPERS — FIID mapping + subject normalization
   ------------------------------------------------------------------------ */

/**
 * Build index from "Banks" (canonical FIID + Bank Name) and optional "Synonyms" (alias → canonical).
 * Only FIIDs present in Banks.FIID (plus accepted extras) are allowed for Column C.
 * Also treat canonical FIIDs, Bank Names and Synonyms as searchable keys.
 * Extras accepted as canonical per spec: QCSB, ITC PLC, QPAY (with some aliases).
 */
function getFiidAliasIndex_() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('FIID_INDEX_FINAL');
  if (cached) { try { return JSON.parse(cached); } catch (e) {} }

  const ss = SpreadsheetApp.getActive();
  const banks = ss.getSheetByName('Banks');     // expects columns incl. "FIID" and ideally "Bank Name"
  const syn   = ss.getSheetByName('Synonyms');  // optional; expects "Canonical", "Aliases" (comma-sep)

  // ---- Canonical + maps ----
  const canonSet = new Set();
  const aliasToCanon = {};     // key -> FIID (aliases/short forms)
  const nameToCanon  = {};     // key -> FIID (full bank names)

  // Read Banks
  if (banks) {
    const rng = banks.getDataRange().getValues();
    if (rng.length > 1) {
      const header = rng[0].map(String);
      const fiidCol = header.findIndex(h => /^fiid$/i.test(h));
      const nameCol = header.findIndex(h => /^bank\s*name$/i.test(h));
      for (let i = 1; i < rng.length; i++) {
        const fiid = rng[i][fiidCol] != null ? String(rng[i][fiidCol]).trim() : '';
        if (!fiid) continue;
        canonSet.add(fiid);

        // key forms for FIID itself
        const kFiid = normalizeAliasKey_(fiid);
        if (kFiid) aliasToCanon[kFiid] = fiid;

        // full bank name
        if (nameCol >= 0) {
          const nm = rng[i][nameCol] != null ? String(rng[i][nameCol]).trim() : '';
          if (nm) {
            const kName = normalizeAliasKey_(nm);
            if (kName) nameToCanon[kName] = fiid;
          }
        }
      }
    }
  }

  // Add accepted extras + their aliases
  const extras = [
    { canon: 'QCSB',    aliases: ['qcsb','q-csb','csbq'] },
    { canon: 'ITC PLC', aliases: ['itc','itc plc','it consultants','itcplc','itc-plc'] },
    { canon: 'QPAY',    aliases: ['qpay','q-pay','q pay'] },
  ];
  extras.forEach(x => {
    canonSet.add(x.canon);
    nameToCanon[normalizeAliasKey_(x.canon)] = x.canon;
    aliasToCanon[normalizeAliasKey_(x.canon)] = x.canon;
    (x.aliases || []).forEach(a => aliasToCanon[normalizeAliasKey_(a)] = x.canon);
  });

  // Read Synonyms (optional)
  if (syn) {
    const rng = syn.getDataRange().getValues();
    if (rng.length > 1) {
      const header = rng[0].map(String);
      const canonCol = header.findIndex(h => /^canonical$/i.test(h) || /^fiid$/i.test(h));
      const aliasCol = header.findIndex(h => /^aliases?$/i.test(h));
      if (canonCol >= 0 && aliasCol >= 0) {
        for (let i = 1; i < rng.length; i++) {
          const canon = rng[i][canonCol] != null ? String(rng[i][canonCol]).trim() : '';
          if (!canon || !canonSet.has(canon)) continue;
          const line = rng[i][aliasCol] != null ? String(rng[i][aliasCol]) : '';
          const parts = line.split(',').map(s => s.trim()).filter(Boolean);
          for (const raw of parts) {
            const key = normalizeAliasKey_(raw);
            if (key) aliasToCanon[key] = canon;
          }
          // Treat the canonical name as a bank-name key, too
          nameToCanon[normalizeAliasKey_(canon)] = canon;
        }
      }
    }
  }

  const out = { canonSet: Array.from(canonSet), aliasToCanon, nameToCanon };
  cache.put('FIID_INDEX_FINAL', JSON.stringify(out), 300); // 5 minutes
  return out;
}

/**
 * FIID selection per your priorities:
 *  1) From Title, find the first non-(RE/FW/FWD) segment on the LEFT of ":" / "::" (scanning left→right).
 *     - If found, try to map that hint (FIID/alias/bank-name) → canonical; if mapped, return it immediately.
 *  2) Otherwise scan the FULL Title text:
 *     - Find any FIID / alias / bank-name; map → canonical
 *     - If multiple canonical → join as "FIID-FIID" (dedup, preserve first-seen order)
 *  3) If nothing matched → ["OTHERS"].
 * Always returns an array of canonical FIIDs or ['OTHERS'].
 */
function selectFIIDsForTitle_(title, idx) {
  const canonSet = new Set(idx.canonSet || []);
  const aliasMap = idx.aliasToCanon || {};
  const nameMap  = idx.nameToCanon  || {};

  const t = String(title || '');

  // Step 1: left hint (first non-RE/FW/FWD segment)
  const hint = firstNonReplyFwdLeftSegment_(t);
  if (hint) {
    const mapped = mapAnyToCanon_(hint, canonSet, aliasMap, nameMap);
    if (mapped) return [mapped];
  }

  // Step 2: full text scan
  const bag = new LinkedOrderedSet_();
  const text = normalizeScanText_(t);

  // Bank names first
  for (const k in nameMap) {
    const canon = nameMap[k];
    if (!canonSet.has(canon)) continue;
    if (containsKey_(text, k)) bag.add(canon);
  }
  // Then aliases + FIID-like
  for (const k in aliasMap) {
    const canon = aliasMap[k];
    if (!canonSet.has(canon)) continue;
    if (containsKey_(text, k)) bag.add(canon);
  }

  if (bag.size() > 0) return bag.values();
  return ['OTHERS'];
}

/* ---------------------------- Parsing & normalization helpers ---------------------------- */

// First non RE/FW/FWD segment before ":" / "::"
function firstNonReplyFwdLeftSegment_(s) {
  const str = String(s || '');
  // Split on :: or :
  const parts = str.split(/::|:/);
  for (let i = 0; i < parts.length; i++) {
    const seg = String(parts[i]).trim();
    if (!seg) continue;
    if (/^\s*(re|fw|fwd)\s*$/i.test(seg)) continue; // skip
    return seg;
  }
  return '';
}

function rightOfColon_(s) {
  const str = String(s || '');
  const parts = str.split(/::|:/);
  if (parts.length <= 1) return str;         // no colon → whole string
  return String(parts.slice(1).join(':')).trim(); // everything to the right (preserve content)
}

// Normalize subject per rule-set
function normalizeSubject_(partRightOfColon) {
  let s = String(partRightOfColon || '').toLowerCase();

  // strip re:/fwd: prefixes (defensive)
  s = s.replace(/\b(re|fwd?|fw)\s*[:\-–—]\s*/g, '');

  // remove [tickets], (ref:...), <...>
  s = s.replace(/\[[^\]]+\]/g, '').replace(/\([^\)]+\)/g, '').replace(/<[^>]+>/g, '');

  // remove URLs/emails
  s = s.replace(/\bhttps?:\/\/\S+/g, '').replace(/\b\S+@\S+\.\S+\b/g, '');

  // remove punctuation/symbol/emoji (broad ASCII)
  s = s.replace(/[^a-z0-9\s]/g, ' ');

  // collapse spaces
  s = s.replace(/\s+/g, ' ').trim();

  // stop-words
  const STOP = new Set(['please','plz','kindly','urgent','issue','problem','error','bug','task','ticket','need','required','request','help','update','info','regarding','about']);
  s = s.split(' ').filter(w => w && !STOP.has(w)).join(' ');
  return s;
}

// Text normalization for scanning
function normalizeScanText_(x) { return String(x || '').toLowerCase(); }

// boundary-aware contains
function containsKey_(textLower, keyLower) {
  const k = escapeRegex_(keyLower);
  const re = new RegExp(`(^|[^a-z0-9])${k}([^a-z0-9]|$)`, 'i');
  return re.test(textLower);
}

// Map any token (bank name / alias / fiid-ish) to canonical
function mapAnyToCanon_(raw, canonSet, aliasMap, nameMap) {
  if (!raw) return '';

  const nk = normalizeAliasKey_(raw);

  // Try bank name
  if (nk && nameMap[nk] && canonSet.has(nameMap[nk])) return nameMap[nk];

  // Try alias (short forms, acronyms, FIID spelled variously)
  if (nk && aliasMap[nk] && canonSet.has(aliasMap[nk])) return aliasMap[nk];

  // Try FIID-like token cleanup (remove PLC/LTD/BANK/space/dash/dot)
  const token = normalizeFiidToken_(raw);
  for (const c of canonSet) {
    if (token === normalizeFiidToken_(c)) return c;           // exact canonical
  }
  // prefix heuristic (e.g., MDBL -> MDB if MDB is a canonical)
  for (const c of canonSet) {
    const ct = normalizeFiidToken_(c);
    if (ct && token.startsWith(ct)) return c;
  }
  return '';
}

// FIID-like token normalization (uppercase, drop PLC/LTD/BANK etc., remove space/dot/dash/underscore)
function normalizeFiidToken_(x) {
  if (x === null || x === undefined) return '';
  let s = String(x).toUpperCase().trim();
  s = s.replace(/[\s._\-]+/g, '');
  s = s.replace(/(PLC|LIMITED|LTD|BANK|BANGLADESH|BANGLA)$/i, '');
  return s.replace(/[\s._\-]+/g, '');
}

// Alias key normalization (lowercase, remove bank/plc/ltd words, normalize spaces)
function normalizeAliasKey_(x) {
  if (x === null || x === undefined) return '';
  let s = String(x).toLowerCase();
  s = s.replace(/\b(bank|plc|limited|ltd)\b/gi, '');
  s = s.replace(/[\s._\-]+/g, ' ').trim();
  s = s.replace(/\s+/g, ' ');
  return s;
}

/* ---------------------------- Similarity + small utils ---------------------------- */
function similarity_(a, b) {
  if (a === b) return 1;
  const la = a.length, lb = b.length;
  if (!la || !lb) return 0;
  const dp = Array(lb + 1);
  for (let j = 0; j <= lb; j++) dp[j] = j;
  for (let i = 1; i <= la; i++) {
    let prev = dp[0]; dp[0] = i;
    for (let j = 1; j <= lb; j++) {
      const tmp = dp[j];
      const cost = a[i - 1] === b[j - 1] ? 0 : 1;
      dp[j] = Math.min(dp[j] + 1, dp[j - 1] + 1, prev + cost);
      prev = tmp;
    }
  }
  const dist = dp[lb];
  const maxLen = Math.max(la, lb);
  return 1 - (dist / maxLen);
}
function escapeRegex_(s) { return String(s).replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); }

// Ordered set (preserves insertion order)
function LinkedOrderedSet_(){ this._map = {}; this._list = []; }
LinkedOrderedSet_.prototype.add = function(v){ if(!this._map[v]){ this._map[v]=true; this._list.push(v); } };
LinkedOrderedSet_.prototype.values = function(){ return this._list.slice(); };
LinkedOrderedSet_.prototype.size = function(){ return this._list.length; };

// Cosmetic: clear old duplicate note
function clearDupNote_(sheet, row) {
  const note = String(sheet.getRange(row, 4).getNote() || '');
  if (/duplicate/i.test(note)) sheet.getRange(row, 4).setNote('');
}
