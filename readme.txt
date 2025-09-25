
# CMS Issue Log — README

Last updated: 2025-09-25
Owner: CMS Team, ITC PLC
Platform: Google Sheets + Google Apps Script

---

## 1) What this is

**CMS Issue Log** is a Google Sheet–based logging system for incidents/emails/events coming from your card management or operations mailbox. A Google Apps Script powers automation to:
- Normalize email subjects (drops "RE:", "FW:", trims refs/URLs, etc.)
- Map a raw bank name/abbreviation to a **Canonical Bank Name** using a **Synonyms** table
- Suggest a **FIID** from the **Banks** master list (hint-first, then full scan; multi-match join; fallback to "OTHERS")
- Detect **duplicates** using normalization + exact/fuzzy comparison (threshold ≥ 0.90 by default)
- Send a **daily summary email** with counts and breakdowns

> Triggered primarily via `onEdit(e)`: when a user edits a row in **Main Sheet** the rules run automatically.

---

## 2) Sheet structure

### 2.1 Main Sheet (the log)
Recommended columns (you may add/ edit more as needed):
- `Date`
- `Time`
- `Source/Email`
- `Subject (Raw)`
- `Subject (Clean)`  ← auto-generated
- `Bank Name (Raw)`
- `Bank (Canonical)` ← auto-selected via Synonyms
- `FIID (Suggested)` ← auto-suggested via Banks
- `Status` (Open / In Progress / Resolved)
- `Notes`

Special cell:
- `Z1` — **LOCK** cell used for re-entry guard. The script writes and clears this to prevent nested runs. Do not edit.

### 2.2 Banks (reference table)
Columns:
- `FIID`
- `Bank Name` (official)

Used by FIID suggestion logic.

### 2.3 Synonyms (reference table)
Columns:
- `Canonical` (official bank name)
- `Aliases` (comma-separated variants/abbreviations/spellings)

Example row:
```
Canonical: Shahjalal Islami Bank
Aliases: SJIBL,SJIBPLC,SJBL,SHAH JALAL ISLAMI
```

### 2.4 Config (optional but recommended)
Key–value settings such as:
- `FUZZY_THRESHOLD` → default `0.90`
- `DAILY_EMAIL_TO` → `ops@example.com,lead@example.com`

If omitted, the script uses hard-coded defaults.

---

## 3) First-time setup

1) Create the four tabs: **Main Sheet**, **Banks**, **Synonyms**, **Config** (optional).  
2) Paste the provided Apps Script into Extensions → **Apps Script** (replace or add to your project).  
3) Set triggers:
   - `onEdit` — simple trigger (works automatically).
   - Daily summary — **time-based** trigger (e.g., every day at 18:00).  
4) On first run, grant permissions when prompted.  
5) Populate:
   - **Banks** with official FIIDs and names.
   - **Synonyms** with canonical names + aliases.
   - **Config** with threshold and email recipients (optional).

---

## 4) How it works (runtime logic)

### 4.1 Subject & Bank normalization
- Cleans `Subject (Raw)` → writes to `Subject (Clean)` (drops RE/FW, trims trackers/URLs, lowercases for compare).
- Looks up `Bank Name (Raw)` → resolves to `Bank (Canonical)` using **Synonyms** (aliases → canonical).

### 4.2 FIID suggestion
Order of attempts:
1. **Hint-first**: prefer FIIDs whose bank names match the canonical/alias hints.
2. **Full scan** across `Banks`:
   - Normalize and attempt **exact** match.
   - Fall back to **fuzzy** match; keep matches ≥ `FUZZY_THRESHOLD` (default 0.90).
3. Multiple matches: join to show all candidates (e.g., `"FIID1 / FIID2"`).
4. No match: use `"OTHERS"` as a safe fallback.

### 4.3 Duplicate control
- Normalizes relevant fields (e.g., `Subject (Clean)`, `Source/Email`).
- Flags **exact** duplicates immediately.
- Flags **fuzzy** potential duplicates when similarity ≥ threshold.
- You can then review/merge/close duplicates via `Status` and `Notes`.

### 4.4 Daily summary email
- Time-based trigger compiles stats:
  - New vs Open counts
  - Potential duplicates
  - Bank-wise distribution
- Sends to recipients defined in **Config → DAILY_EMAIL_TO** (or script default).

---

## 5) Daily workflow (for operators)

1. Add/import new items to **Main Sheet**.  
2. As you type `Bank Name (Raw)`, the sheet fills `Bank (Canonical)` and suggests `FIID (Suggested)`.  
3. If `FIID (Suggested)` shows multiple values, pick the correct one manually.  
4. Update `Status`: `Open → In Progress → Resolved`.  
5. If a duplicate warning appears, check the suggested match and merge/close as needed.  
6. Review the **daily summary email** for oversight.

---

## 6) Managing reference data

### 6.1 Banks
- Add a row with the official `FIID` and `Bank Name`.
- Keep naming consistent with how you want it to appear in reports.

### 6.2 Synonyms
- For a bank with many variations, keep **one** row with its canonical name.
- Add all known variants to `Aliases`, comma-separated.
- Example additions for Mutual Trust Bank:
  `MDBL,MDB,MDBPLC,Mutual Trust Bank Ltd,Mutual Trust Bank`

### 6.3 Config
- `FUZZY_THRESHOLD`
  - Increase (e.g., `0.92`) → stricter matching (fewer false positives, more misses).
  - Decrease (e.g., `0.88`) → looser matching (fewer misses, more false positives).
- `DAILY_EMAIL_TO`: comma-separated list of recipients.

---

## 7) Triggers & permissions

- **onEdit**: runs automatically when editing **Main Sheet**.  
- **Time-driven**: set daily at a reasonable hour in your timezone (e.g., 18:00 Asia/Dhaka).  
- Permissions:
  - The script needs access to send email and read/write the spreadsheet.
  - If you copy the file to a new owner, reauthorize triggers under the new owner.

---

## 8) Security & governance

- Share the spreadsheet using least-privilege (Viewer for readers, Editor only for maintainers).
- Protect reference tabs (**Banks**, **Synonyms**, **Config**) with range protections where practical.
- Consider a monthly backup (File → Make a copy or Version history named checkpoint).
- Avoid placing secrets in the sheet; use script properties if you must store tokens.

---

## 9) Troubleshooting

**onEdit not firing**
- Verify the sheet name is exactly `"Main Sheet"` (or update the script accordingly).
- Ensure the `Z1` lock cell is empty (not stuck on `"LOCK"`).
- Simple triggers cannot run if the user lacks permissions—confirm user access.

**FIID suggestion blank or wrong**
- Confirm **Banks** contains a matching bank.
- Add missing variants to **Synonyms → Aliases**.
- Adjust `FUZZY_THRESHOLD` slightly and test.

**Too many duplicates flagged**
- Raise `FUZZY_THRESHOLD` (e.g., 0.90 → 0.92).

**Daily summary email missing**
- Ensure the time-based trigger is enabled.
- Confirm `DAILY_EMAIL_TO` has valid addresses.

**Script took too long / Exceeded execution time**
- Reduce batch edits.
- Use filtered views to limit mass recalculations.

---

## 10) Maintenance checklist

- [ ] Review and update **Synonyms** monthly.
- [ ] Append new banks to **Banks** as soon as they are onboarded.
- [ ] Verify triggers after ownership or permission changes.
- [ ] Export a CSV backup quarterly.
- [ ] Skim daily summary emails for anomalies.

---

## 11) Release notes (template)

**Date:** YYYY-MM-DD  
**Changes:**
- FIID selection improved (hint-first → full-scan fallback; multi-match join; OTHERS fallback)
- Duplicate control hardened (normalization + Exact/Fuzzy ≥ 0.90; RE/FW drop)
- Daily email enhanced (new/open count + bank-wise breakdown)
**Impact:** Higher match accuracy, fewer duplicates, clearer oversight  
**Action:** Update Synonyms; recalibrate threshold if needed

---

## 12) Sample test titles

- `FW: ATM Cash Withdrawal dispute – SJIBPLC Uttara`
- `RE: Card not working at POS – MDB`
- `Chargeback timeline request – SJBL`
- `PIN retry exceeded – MDBPLC`
- `Terminal down – Shahjalal Islami Bank`

---

## 13) Glossary

- **Canonical (Bank Name)**: The official, single source of truth name we want to use everywhere.
- **Alias**: Any alternate spelling/abbreviation users send in emails (e.g., SJIBL, SJBL).
- **FIID**: Financial Institution Identifier used for routing/identification in CMS.
- **Fuzzy match**: Text similarity scoring; matches above a threshold are treated as potential matches.

---

## 14) FAQ

**Q: Can I rename the tabs?**  
A: Yes, but update the script constants accordingly. `Main Sheet` name is referenced explicitly.

**Q: Can I store more metadata (ticket no., branch, etc.)?**  
A: Yes. Add columns to **Main Sheet**. The script ignores unknown columns.

**Q: We operate multiple business units—can we split logs?**  
A: Create separate sheets or add a `BU` column and filter your daily email/report by BU.

**Q: Where do I change recipients for the daily email?**  
A: In **Config → DAILY_EMAIL_TO** (comma-separated), or in the script default if not using Config.

---

## 15) Support

- Primary contact: (Name, Email)
- Escalation: (Name, Email)
- Script repository link (optional): (URL)

---

## 16) Change log (fill as you go)

- 2025-09-25: Initial README created for the CMS Issue Log.
